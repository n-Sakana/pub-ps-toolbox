using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace PsToolbox
{
    public static class App
    {
        static App()
        {
            Tools = new List<ToolManifest>();
            LastStatus = "Ready";
        }

        public static string AppDir { get; private set; }
        public static List<ToolManifest> Tools { get; private set; }
        public static string LastStatus { get; set; }

        public static void Run(string appDir)
        {
            AppDir = appDir;
            Config.Load(Path.Combine(appDir, "config.json"));
            ReloadTools();
            LastStatus = ContextMenuManager.Rebuild(AppDir, Tools);

            var app = new Application();
            app.ShutdownMode = ShutdownMode.OnMainWindowClose;
            app.Exit += (s, e) =>
            {
                try { ContextMenuManager.Cleanup(Tools); }
                catch { }
            };
            app.Run(new HostWindow());
        }

        public static void ReloadTools()
        {
            Tools = ToolRegistry.Load(Path.Combine(AppDir, "tools"));
        }

        public static IEnumerable<ToolManifest> EnabledTools()
        {
            return Tools.Where(t => Config.ToolEnabled(t)).OrderBy(t => t.DisplayName);
        }

        public static string ApplyContextMenus()
        {
            ReloadTools();
            LastStatus = ContextMenuManager.Rebuild(AppDir, Tools);
            return LastStatus;
        }

        public static void ResetConfigFromDisk()
        {
            Config.Load(Path.Combine(AppDir, "config.json"));
            ReloadTools();
        }
    }
}
