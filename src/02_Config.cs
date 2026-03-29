using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;

namespace PsToolbox
{
    public static class Config
    {
        static readonly Dictionary<string, string> Data = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        static string _path;

        public static void Load(string path)
        {
            _path = path;
            Data.Clear();
            if (!File.Exists(path)) return;

            var raw = File.ReadAllText(path, Encoding.UTF8);
            if (string.IsNullOrWhiteSpace(raw)) return;

            try
            {
                var ser = new JavaScriptSerializer();
                var loaded = ser.Deserialize<Dictionary<string, string>>(raw);
                if (loaded == null) return;
                foreach (var pair in loaded)
                    Data[pair.Key] = pair.Value ?? string.Empty;
            }
            catch
            {
            }
        }

        public static void Save()
        {
            if (string.IsNullOrWhiteSpace(_path)) return;
            var ser = new JavaScriptSerializer();
            var ordered = Data.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(k => k.Key, v => v.Value ?? string.Empty, StringComparer.OrdinalIgnoreCase);
            File.WriteAllText(_path, ser.Serialize(ordered), new UTF8Encoding(true));
        }

        public static string Get(string key, string def = "")
        {
            string value;
            return Data.TryGetValue(key, out value) ? value : def;
        }

        public static bool GetBool(string key, bool def = false)
        {
            var raw = Get(key, def ? "1" : "0");
            return raw == "1" || raw.Equals("true", StringComparison.OrdinalIgnoreCase);
        }

        public static int GetInt(string key, int def = 0)
        {
            int value;
            return int.TryParse(Get(key, def.ToString()), out value) ? value : def;
        }

        public static void Set(string key, string value)
        {
            Data[key] = value ?? string.Empty;
        }

        public static string MenuRootText
        {
            get { return Get("menu_root_text", "Toolbox"); }
            set { Set("menu_root_text", string.IsNullOrWhiteSpace(value) ? "Toolbox" : value.Trim()); }
        }

        public static string ToolSettingKey(string toolId, string key)
        {
            return string.Format("tool.{0}.{1}", toolId, key);
        }

        public static bool ToolEnabled(ToolManifest tool)
        {
            return GetBool(ToolSettingKey(tool.Id, "enabled"), tool.EnabledByDefault);
        }

        public static void SetToolEnabled(ToolManifest tool, bool enabled)
        {
            Set(ToolSettingKey(tool.Id, "enabled"), enabled ? "1" : "0");
        }
    }
}
