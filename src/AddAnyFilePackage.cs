﻿using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Interop;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using System.Web.Script.Serialization;

namespace MadsKristensen.AddAnyFile
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
    [InstalledProductRegistration("#110", "#112", Vsix.Version, IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(PackageGuids.guidAddAnyFilePkgString)]
    public sealed class AddAnyFilePackage : ExtensionPointPackage
    {
        public static DTE2 _dte;

        protected override void Initialize()
        {
            _dte = GetService(typeof(DTE)) as DTE2;

            Logger.Initialize(this, Vsix.Name);
            Telemetry.Initialize(_dte, Vsix.Version, "e146dff7-f7c5-49ab-a7d8-3557375f6624");

            base.Initialize();

            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                CommandID menuCommandID = new CommandID(PackageGuids.guidAddAnyFileCmdSet, PackageIds.cmdidMyCommand);
                var menuItem = new OleMenuCommand(MenuItemCallback, menuCommandID);
                menuItem.BeforeQueryStatus += MenuItem_BeforeQueryStatus;
                mcs.AddCommand(menuItem);
            }
        }

        private void MenuItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            var button = (OleMenuCommand)sender;
            button.Visible = button.Enabled = false;

            UIHierarchyItem item = GetSelectedItem();
            var project = item.Object as Project;

            if (project == null || !project.Kind.Equals(EnvDTE.Constants.vsProjectKindSolutionItems, StringComparison.OrdinalIgnoreCase))
                button.Visible = button.Enabled = true;
        }

        private async void MenuItemCallback(object sender, EventArgs e)
        {
            UIHierarchyItem item = GetSelectedItem();

            if (item == null)
                return;

            string currentFilePath = "";
            string folder = FindFolder(item, out currentFilePath);

            if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder))
                return;

            Project project = ProjectHelpers.GetActiveProject();
            if (project == null)
                return;

            var testProjectData = FindMatchingTestProject(project);
            var allProjects = GetAllProjectsInSolution();
            var testProject = allProjects.Single(x => x.Name == testProjectData.Name);
            //string input = PromptForFileName(folder).TrimStart('/', '\\').Replace("/", "\\");

            //if (string.IsNullOrEmpty(input))
            //    return;


            var currentRelativePath = GetPathFromProjectFolder(project, currentFilePath);
            var input = testProjectData.Path +  currentRelativePath;

            string[] parsedInputs = GetParsedInput(input);

            foreach (string inputItem in parsedInputs)
            {
                input = inputItem;

                if (input.EndsWith("\\", StringComparison.Ordinal))
                {
                    input = input + "__dummy__";
                }

                folder = Path.GetDirectoryName(input);
                string file = Path.Combine(folder, input);
                string dir = Path.GetDirectoryName(file);

                PackageUtilities.EnsureOutputPath(dir);

                if (!File.Exists(file))
                {
                    int position = await WriteFile(testProject, file);

                    try
                    {
                        var projectItem = testProject.AddFileToProject(file);

                        if (file.EndsWith("__dummy__"))
                        {
                            Telemetry.TrackEvent("Folder added");
                            projectItem.Delete();
                            continue;
                        }

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
                    }
                    catch (Exception ex)
                    {
                        Logger.Log(ex);
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("The file '" + file + "' already exist.");
                }
            }
        }

        private string GetPathFromProjectFolder(Project project, string currentFilePath)
        {
            var projectName = project.Name;
            var relativeFolder = currentFilePath.Split(new[] { projectName }, StringSplitOptions.RemoveEmptyEntries).Last();
            var testsFileName = relativeFolder.Replace(".cs","Tests.cs");

            return testsFileName;
        }

        private TestProjectData FindMatchingTestProject(Project project)
        {
            var allProjectsInSolution = GetAllProjectsInSolution();

            using (StreamReader r = new StreamReader("settings.json"))
            {
                string json = r.ReadToEnd();
                var serializer = new JavaScriptSerializer();
                dynamic parsedJson = serializer.DeserializeObject(json);
                return new TestProjectData
                {
                    Path = parsedJson[project.Name]["path"],
                    Name = parsedJson[project.Name]["name"],
                };
            }
        }

        private class TestProjectData
        {
            public string Name { get; set; }
            public string Path { get; set; }
        }

        private static async Task<int> WriteFile(Project project, string file)
        {
            Encoding encoding = new UTF8Encoding(false);
            string extension = Path.GetExtension(file);
            string template = await TemplateMap.GetTemplateFilePath(project, file);

            var props = new Dictionary<string, string>() { { "extension", extension.ToLowerInvariant() } };
            Telemetry.TrackEvent("File added", props);

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

        static string[] GetParsedInput(string input)
        {
            // var tests = new string[] { "file1.txt", "file1.txt, file2.txt", ".ignore", ".ignore.(old,new)", "license", "folder/",
            //    "folder\\", "folder\\file.txt", "folder/.thing", "page.aspx.cs", "widget-1.(html,js)", "pages\\home.(aspx, aspx.cs)",
            //    "home.(html,js), about.(html,js,css)", "backup.2016.(old, new)", "file.(txt,txt,,)", "file_@#d+|%.3-2...3^&.txt" };
            Regex pattern = new Regex(@"[,]?([^(,]*)([\.\/\\]?)[(]?((?<=[^(])[^,]*|[^)]+)[)]?");
            List<string> results = new List<string>();
            Match match = pattern.Match(input);

            while (match.Success)
            {
                // Alwasy 4 matches w. Group[3] being the extension, extension list, folder terminator ("/" or "\"), or empty string
                string path = match.Groups[1].Value.Trim() + match.Groups[2].Value;
                string[] extensions = match.Groups[3].Value.Split(',');

                foreach (string ext in extensions)
                {
                    string value = path + ext.Trim();

                    // ensure "file.(txt,,txt)" or "file.txt,,file.txt,File.TXT" retuns as just ["file.txt"]
                    if (value != "" && !value.EndsWith(".", StringComparison.Ordinal) && !results.Contains(value, StringComparer.OrdinalIgnoreCase))
                    {
                        results.Add(value);
                    }
                }
                match = match.NextMatch();
            }
            return results.ToArray();
        }

        private string PromptForFileName(string folder)
        {
            DirectoryInfo dir = new DirectoryInfo(folder);
            var dialog = new FileNameDialog(dir.Name);

            var hwnd = new IntPtr(_dte.MainWindow.HWnd);
            var window = (System.Windows.Window)HwndSource.FromHwnd(hwnd).RootVisual;
            dialog.Owner = window;

            var result = dialog.ShowDialog();
            return (result.HasValue && result.Value) ? dialog.Input : string.Empty;
        }

        private static string FindFolder(UIHierarchyItem item, out string currentFilePath)
        {
            Window2 window = _dte.ActiveWindow as Window2;
            currentFilePath = "";

            if (window != null && window.Type == vsWindowType.vsWindowTypeDocument)
            {
                // if a document is active, use the document's containing directory
                Document doc = _dte.ActiveDocument;
                if (doc != null && !string.IsNullOrEmpty(doc.FullName))
                {
                    ProjectItem docItem = _dte.Solution.FindProjectItem(doc.FullName);

                    if (docItem != null)
                    {
                        currentFilePath = docItem.Properties.Item("FullPath").Value.ToString();
                        
                        if (File.Exists(currentFilePath))
                            return Path.GetDirectoryName(currentFilePath);
                    }
                }
            }

            string folder = null;

            ProjectItem projectItem = item.Object as ProjectItem;
            Project project = item.Object as Project;

            if (projectItem != null)
            {
                string fileName = projectItem.FileNames[0];

                if (File.Exists(fileName))
                {
                    folder = Path.GetDirectoryName(fileName);
                }
                else
                {
                    folder = fileName;
                }
            }
            else if (project != null)
            {
                folder = project.GetRootFolder();
            }

            return folder;
        }

        private static UIHierarchyItem GetSelectedItem()
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
            DTE2 dte2 = Package.GetGlobalService(typeof(DTE)) as DTE2;
            return dte2;
        }

        public static IList<Project> GetAllProjectsInSolution()
        {
            Projects projects = GetActiveIDE().Solution.Projects;
            List<Project> list = new List<Project>();
            var item = projects.GetEnumerator();
            while (item.MoveNext())
            {
                var project = item.Current as Project;
                if (project == null)
                {
                    continue;
                }

                if (project.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                {
                    list.AddRange(GetSolutionFolderProjects(project));
                }
                else
                {
                    list.Add(project);
                }
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
                {
                    continue;
                }

                // If this is another solution folder, do a recursive call, otherwise add
                if (subProject.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                {
                    list.AddRange(GetSolutionFolderProjects(subProject));
                }
                else
                {
                    list.Add(subProject);
                }
            }
            return list;
        }
    }
}
