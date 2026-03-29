using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;

namespace PsToolbox
{
    public class ToolSetting
    {
        public string Key { get; set; }
        public string Label { get; set; }
        public string Type { get; set; }
        public string Default { get; set; }
        public string Hint { get; set; }
        public string[] Options { get; set; }
    }

    public class ToolManifest
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string Description { get; set; }
        public string MenuText { get; set; }
        public string[] Targets { get; set; }
        public string[] Extensions { get; set; }
        public string Script { get; set; }
        public bool EnabledByDefault { get; set; }
        public List<ToolSetting> Settings { get; set; }
        public string ToolDir { get; set; }

        public bool SupportsFiles
        {
            get { return (Targets ?? new string[0]).Any(t => string.Equals(t, "file", StringComparison.OrdinalIgnoreCase)); }
        }

        public bool SupportsFolders
        {
            get { return (Targets ?? new string[0]).Any(t => string.Equals(t, "folder", StringComparison.OrdinalIgnoreCase)); }
        }

        public string BuildFileAppliesTo()
        {
            var exts = (Extensions ?? new string[0])
                .Where(e => !string.IsNullOrWhiteSpace(e) && e != "*")
                .Select(e => e.StartsWith(".") ? e.ToLowerInvariant() : "." + e.ToLowerInvariant())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();

            if (exts.Length == 0)
                return null;

            return string.Join(" OR ", exts.Select(e => string.Format("System.FileExtension:=\"{0}\"", e)));
        }

        public override string ToString()
        {
            return DisplayName ?? Id ?? base.ToString();
        }
    }

    public static class ToolRegistry
    {
        public static List<ToolManifest> Load(string toolsRoot)
        {
            var result = new List<ToolManifest>();
            if (!Directory.Exists(toolsRoot))
                return result;

            var serializer = new JavaScriptSerializer();
            foreach (var manifestPath in Directory.GetFiles(toolsRoot, "tool.json", SearchOption.AllDirectories))
            {
                var raw = File.ReadAllText(manifestPath);
                var tool = serializer.Deserialize<ToolManifest>(raw);
                if (tool == null || string.IsNullOrWhiteSpace(tool.Id))
                    continue;
                tool.ToolDir = Path.GetDirectoryName(manifestPath);
                tool.Settings = tool.Settings ?? new List<ToolSetting>();
                tool.Targets = tool.Targets ?? new string[0];
                tool.Extensions = tool.Extensions ?? new string[0];
                result.Add(tool);
            }

            return result.OrderBy(t => t.DisplayName).ToList();
        }
    }
}
