using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PsToolbox
{
    public static class ContextMenuManager
    {
        const string SessionKeyName = "PsToolbox.Session";
        static readonly string FileRoot = @"Software\Classes\*\shell\" + SessionKeyName;
        static readonly string FolderRoot = @"Software\Classes\Directory\shell\" + SessionKeyName;

        public static string Rebuild(string appDir, IList<ToolManifest> tools)
        {
            Cleanup(tools);

            var enabled = tools.Where(t => Config.ToolEnabled(t)).ToList();
            var fileTools = enabled.Where(t => t.SupportsFiles).ToList();
            var folderTools = enabled.Where(t => t.SupportsFolders).ToList();

            if (fileTools.Any())
            {
                EnsureParent(FileRoot);
                foreach (var tool in fileTools)
                    RegisterTool(FileRoot, tool, appDir, true);
            }

            if (folderTools.Any())
            {
                EnsureParent(FolderRoot);
                foreach (var tool in folderTools)
                    RegisterTool(FolderRoot, tool, appDir, false);
            }

            if (!enabled.Any())
                return "No active tools registered";

            return string.Format("Context menu active: {0}", string.Join(", ", enabled.Select(t => t.DisplayName)));
        }

        public static void Cleanup(IList<ToolManifest> tools)
        {
            TryDelete(FileRoot);
            TryDelete(FolderRoot);
        }

        static void EnsureParent(string parentKey)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(parentKey))
            {
                key.SetValue("MUIVerb", Config.MenuRootText);
                key.SetValue("Icon", "shell32.dll,44");
                key.SetValue("SubCommands", string.Empty);
            }
        }

        static void RegisterTool(string parentKey, ToolManifest tool, string appDir, bool fileContext)
        {
            var keyPath = parentKey + @"\shell\" + tool.Id;
            using (var key = Registry.CurrentUser.CreateSubKey(keyPath))
            {
                key.SetValue(string.Empty, tool.MenuText ?? tool.DisplayName ?? tool.Id);
                key.SetValue("Icon", "shell32.dll,16");
                key.SetValue("MultiSelectModel", "Player");

                var appliesTo = fileContext ? tool.BuildFileAppliesTo() : null;
                if (!string.IsNullOrWhiteSpace(appliesTo))
                    key.SetValue("AppliesTo", appliesTo);
            }

            using (var cmd = Registry.CurrentUser.CreateSubKey(keyPath + @"\command"))
            {
                cmd.SetValue(string.Empty, BuildCommand(appDir, tool.Id));
            }
        }

        static string BuildCommand(string appDir, string toolId)
        {
            var launchVbs = Path.Combine(appDir, "launch.vbs");
            return string.Format("wscript.exe \"{0}\" --invoke \"{1}\" \"%1\"", launchVbs, toolId);
        }

        static void TryDelete(string subKey)
        {
            try { Registry.CurrentUser.DeleteSubKeyTree(subKey, false); }
            catch { }
        }
    }
}


