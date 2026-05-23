using System;
using System.IO;
using System.Text.Json;

namespace RegistryExpert.Core
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
        /// When true, RegistryExpert silently downloads available updates in
        /// the background on startup and prompts to restart when ready.
        /// When false, the user must trigger the download from the
        /// "Update Available" dialog.
        /// </summary>
        public bool AutoDownloadUpdates { get; set; } = true;

        /// <summary>
        /// The version string that was running last time settings were saved.
        /// Used to detect when the user has just upgraded (via auto-update OR
        /// manual reinstall) so we can show a "Welcome to vX" banner.
        /// </summary>
        public string LastSeenVersion { get; set; } = "";

        /// <summary>
        /// Storage tab → Mounted Devices: show entries whose MountType is "Drive Letter".
        /// </summary>
        public bool MountedDevicesShowDriveLetters { get; set; } = true;

        /// <summary>
        /// Storage tab → Mounted Devices: show entries whose MountType is "Volume GUID".
        /// </summary>
        public bool MountedDevicesShowVolumeGuids { get; set; } = false;

        /// <summary>
        /// Storage tab → Mounted Devices: show entries whose MountType is "Other".
        /// </summary>
        public bool MountedDevicesShowOther { get; set; } = false;

        /// <summary>
        /// Storage tab → Mounted Devices: restrict displayed entries to only those
        /// flagged as stale (orphan volume registrations).
        /// </summary>
        public bool MountedDevicesStaleOnly { get; set; } = false;

        /// <summary>
        /// Storage tab → Disk Layout: whether the "Volume Details" Expander is
        /// expanded. Persisted so the user's preference survives app restart.
        /// </summary>
        public bool DiskLayoutDetailsExpanded { get; set; } = false;

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
