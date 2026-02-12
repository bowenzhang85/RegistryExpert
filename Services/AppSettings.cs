using System;
using System.IO;
using System.Text.Json;

namespace RegistryExpert
{
    /// <summary>
    /// Application settings that persist between sessions
    /// </summary>
    public class AppSettings
    {
        /// <summary>
        /// Splitter distance for the detail pane in Analyze window (pixels from bottom)
        /// </summary>
        public int DetailPanelHeight { get; set; } = 180;

        /// <summary>
        /// Current theme preference
        /// </summary>
        public string Theme { get; set; } = "Dark";

        /// <summary>
        /// Main window width
        /// </summary>
        public int WindowWidth { get; set; } = 1400;

        /// <summary>
        /// Main window height
        /// </summary>
        public int WindowHeight { get; set; } = 900;

        /// <summary>
        /// Main window maximized state
        /// </summary>
        public bool WindowMaximized { get; set; } = false;

        private static readonly string SettingsDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "RegistryExpert");

        private static readonly string SettingsPath = Path.Combine(SettingsDirectory, "settings.json");

        /// <summary>
        /// Load settings from disk, or return defaults if not found
        /// </summary>
        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    var json = File.ReadAllText(SettingsPath);
                    var settings = JsonSerializer.Deserialize<AppSettings>(json);
                    if (settings != null)
                    {
                        return settings;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading settings: {ex.Message}");
            }

            return new AppSettings();
        }

        /// <summary>
        /// Save settings to disk
        /// </summary>
        public void Save()
        {
            try
            {
                // Ensure directory exists
                if (!Directory.Exists(SettingsDirectory))
                {
                    Directory.CreateDirectory(SettingsDirectory);
                }

                var options = new JsonSerializerOptions { WriteIndented = true };
                var json = JsonSerializer.Serialize(this, options);
                File.WriteAllText(SettingsPath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving settings: {ex.Message}");
            }
        }
    }
}
