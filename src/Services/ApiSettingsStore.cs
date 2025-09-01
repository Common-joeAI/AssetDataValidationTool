using System;
using System.IO;
using System.Text.Json;
using AssetDataValidationTool.Models;

namespace AssetDataValidationTool.Services
{
    public static class ApiSettingsStore
    {
        private static readonly string SettingsPath =
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config", "api.settings.json");

        public static ApiSettings Load()
        {
            try
            {
                if (!File.Exists(SettingsPath))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
                    Save(new ApiSettings()); // write defaults
                }
                var json = File.ReadAllText(SettingsPath);
                var opts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                return JsonSerializer.Deserialize<ApiSettings>(json, opts) ?? new ApiSettings();
            }
            catch
            {
                return new ApiSettings();
            }
        }

        public static void Save(ApiSettings settings)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
                var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SettingsPath, json);
            }
            catch { /* ignore */ }
        }
    }
}
