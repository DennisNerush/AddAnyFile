using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using VSTestGenerator.Helpers;

namespace VSTestGenerator
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
    [InstalledProductRegistration("#110", "#112", Vsix.Version, IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(PackageGuids.VSTestGeneratorPkgGuidString)]
    public sealed class TestGeneratorPackage : ExtensionPointPackage
    {
        public static DTE2 _dte;
        public const string VSTestGeneratorDirectory = "VS-Test-Generator";


        protected override void Initialize()
        {
            _dte = GetService(typeof(DTE)) as DTE2;

            Logger.Initialize(this, Vsix.Name);
            Telemetry.Initialize(_dte, Vsix.Version, "e146dff7-f7c5-49ab-a7d8-3557375f6624");

            base.Initialize();

            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                CommandID menuCommandID = new CommandID(PackageGuids.VSTestGeneratorCmdSetGuid, PackageIds.cmdidMyCommand);
                var menuItem = new OleMenuCommand(MenuItemCallback, menuCommandID);
                menuItem.BeforeQueryStatus += MenuItem_BeforeQueryStatus;
                mcs.AddCommand(menuItem);

                var settingsFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    VSTestGeneratorDirectory, "settings.json");

                var settingsDirectory = Path.GetDirectoryName(settingsFilePath);

                if (!Directory.Exists(settingsDirectory))
                    Directory.CreateDirectory(settingsDirectory);

                if (!File.Exists(settingsFilePath))
                {
                   TemplateMap.CopyDefaultSettingsFile(settingsFilePath);
                }
            }
        }

        private void MenuItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            var button = (OleMenuCommand)sender;
            button.Visible = button.Enabled = false;

            UIHierarchyItem item = GetSelectedItemInSolution();
            var project = item.Object as Project;

            if (project == null || !project.Kind.Equals(EnvDTE.Constants.vsProjectKindSolutionItems, StringComparison.OrdinalIgnoreCase))
                button.Visible = button.Enabled = true;
        }

        private async void MenuItemCallback(object sender, EventArgs e)
        {
            string currentFilePath = GetCurrentFilePath();

            if (string.IsNullOrEmpty(currentFilePath) || !File.Exists(currentFilePath))
                return;

            Project project = ProjectHelpers.GetActiveProject();
            if (project == null)
                return;

            var currentFileRelativePathFromCurrentProject = GetPathFromProjectFolder(project, currentFilePath);
            var testProjectData = FindMatchingTestProject(project);
            var testProject = GetTestProject(testProjectData);

            string file = Path.Combine(Path.GetDirectoryName(testProject.FullName) + testProjectData.Path) + currentFileRelativePathFromCurrentProject;
            string dir = Path.GetDirectoryName(file);

            PackageUtilities.EnsureOutputPath(dir);

            if (!File.Exists(file))
            {
                int position = await WriteFile(testProject, file);

                try
                {
                    testProject.AddFileToProject(file);
                    var window = (Window2)_dte.ItemOperations.OpenFile(file);

                    // Move cursor into position
                    if (position > 0)
                    {
                        var view = ProjectHelpers.GetCurentTextView();

                        if (view != null)
                            view.Caret.MoveTo(new SnapshotPoint(view.TextBuffer.CurrentSnapshot, position));
                    }

                    _dte.ExecuteCommand("SolutionExplorer.SyncWithActiveDocument");
                    _dte.ActiveDocument.Activate();
                    _dte.ExecuteCommand("File.SaveAll");
                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show($"Don't worry KAPARA, already got tests for this file ({0})");
            }
        }

        private Project GetTestProject(ProjectData projectData)
        {
            var allProjects = GetAllProjectsInSolution();

            return allProjects.Single(x => x.Name == projectData.Name);
        }

        private string GetPathFromProjectFolder(Project project, string currentFilePath)
        {
            var projectName = project.Name;
            var relativeFolder = currentFilePath.Split(new[] { projectName }, StringSplitOptions.RemoveEmptyEntries).Last();
            var testsFileName = relativeFolder.Replace(".cs","Tests.cs");

            return testsFileName;
        }

        private ProjectData FindMatchingTestProject(Project project)
        {
            var settingsFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                   VSTestGeneratorDirectory, "settings.json");
            using (StreamReader r = new StreamReader(settingsFilePath))
            {
                string json = r.ReadToEnd();
                var serializer = new JavaScriptSerializer();
                dynamic parsedJson = serializer.DeserializeObject(json);
                return new ProjectData
                {
                    Path = parsedJson[project.Name]["path"],
                    Name = parsedJson[project.Name]["name"],
                };
            }
        }

        private static async Task<int> WriteFile(Project project, string file)
        {
            Encoding encoding = new UTF8Encoding(false);
            string template = await TemplateMap.GetTemplateFilePath(project, file);

            if (!string.IsNullOrEmpty(template))
            {
                int index = template.IndexOf('$');
                template = template.Remove(index, 1);
                File.WriteAllText(file, template, encoding);
                return index;
            }

            File.WriteAllText(file, string.Empty, encoding);
            return 0;
        }

        private string GetCurrentFilePath()
        {
            Window2 window = _dte.ActiveWindow as Window2;

            if (window == null || window.Type != vsWindowType.vsWindowTypeDocument)
                return string.Empty;

            
            Document doc = _dte.ActiveDocument;
            if (doc != null && !string.IsNullOrEmpty(doc.FullName))
            {
                ProjectItem docItem = _dte.Solution.FindProjectItem(doc.FullName);

                if (docItem != null)
                    return docItem.Properties.Item("FullPath").Value.ToString();
            }

            return string.Empty;
        }

        private static UIHierarchyItem GetSelectedItemInSolution()
        {
            var items = (Array)_dte.ToolWindows.SolutionExplorer.SelectedItems;

            foreach (UIHierarchyItem selItem in items)
            {
                return selItem;
            }

            return null;
        }

        public static DTE2 GetActiveIDE()
        {
            // Get an instance of currently running Visual Studio IDE.
            DTE2 dte2 = GetGlobalService(typeof(DTE)) as DTE2;
            return dte2;
        }

        private static IList<Project> GetAllProjectsInSolution()
        {
            Projects projects = GetActiveIDE().Solution.Projects;
            List<Project> list = new List<Project>();
            var item = projects.GetEnumerator();
            while (item.MoveNext())
            {
                var project = item.Current as Project;
                if (project == null) 
                    continue;
                    
                if (project.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                    list.AddRange(GetSolutionFolderProjects(project));
                else
                    list.Add(project);
            }

            return list;
        }

        private static IEnumerable<Project> GetSolutionFolderProjects(Project solutionFolder)
        {
            List<Project> list = new List<Project>();
            for (var i = 1; i <= solutionFolder.ProjectItems.Count; i++)
            {
                var subProject = solutionFolder.ProjectItems.Item(i).SubProject;
                if (subProject == null)
                    continue;
                
                // If this is another solution folder, do a recursive call, otherwise add
                if (subProject.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                    list.AddRange(GetSolutionFolderProjects(subProject));
                else
                    list.Add(subProject);
            }

            return list;
        }

        private class ProjectData
        {
            public string Name { get; set; }
            public string Path { get; set; }
        }
    }
}
