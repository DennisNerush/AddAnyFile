using System;

namespace VSTestGenerator
{
    /// <summary>
    /// Helper class that exposes all GUIDs used across VS Package.
    /// </summary>
    internal sealed partial class PackageGuids
    {
        public const string VSTestGeneratorPkgGuidString = "27dd9dea-6dd2-403e-929d-3ff20d896c5e";
        public const string VSTestGeneratorCmdSetGuidString = "32af8a17-bbbc-4c56-877e-fc6c6575a8cf";
        public static Guid VSTestGeneratorPkgGuid = new Guid(VSTestGeneratorPkgGuidString);
        public static Guid VSTestGeneratorCmdSetGuid = new Guid(VSTestGeneratorCmdSetGuidString);
    }
    /// <summary>
    /// Helper class that encapsulates all CommandIDs uses across VS Package.
    /// </summary>
    internal sealed partial class PackageIds
    {
        public const int cmdidMyCommand = 666;
    }
}
