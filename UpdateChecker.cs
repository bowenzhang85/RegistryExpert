using System;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace RegistryExpert
{
    /// <summary>
    /// Contains information about an available update.
    /// </summary>
    public class UpdateInfo
    {
        public bool UpdateAvailable { get; init; }
        public string CurrentVersion { get; init; } = "";
        public string LatestVersion { get; init; } = "";
        public string ReleaseUrl { get; init; } = "";
        public string ReleaseNotes { get; init; } = "";
    }

    /// <summary>
    /// Helper class for checking for application updates via GitHub Releases API.
    /// </summary>
    public static class UpdateChecker
    {
        private const string GitHubApiUrl = "https://api.github.com/repos/bowenzhang85/RegistryExpert/releases/latest";
        private static readonly HttpClient _httpClient;

        static UpdateChecker()
        {
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "RegistryExpert-UpdateChecker");
            _httpClient.DefaultRequestHeaders.Add("Accept", "application/vnd.github.v3+json");
            _httpClient.Timeout = TimeSpan.FromSeconds(10);
        }

        /// <summary>
        /// Gets the current application version as a string (e.g., "1.0.1").
        /// </summary>
        public static string GetCurrentVersion()
        {
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            return $"{version?.Major ?? 1}.{version?.Minor ?? 0}.{version?.Build ?? 0}";
        }

        /// <summary>
        /// Checks for updates by querying the GitHub Releases API.
        /// </summary>
        /// <returns>UpdateInfo with update details, or null if the check failed.</returns>
        public static async Task<UpdateInfo?> CheckForUpdatesAsync()
        {
            try
            {
                var response = await _httpClient.GetAsync(GitHubApiUrl).ConfigureAwait(false);
                
                if (!response.IsSuccessStatusCode)
                {
                    System.Diagnostics.Debug.WriteLine($"GitHub API returned {response.StatusCode}");
                    return null;
                }

                var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                // Extract tag_name (e.g., "v1.0.2"), html_url, and body (release notes)
                var tagName = root.GetProperty("tag_name").GetString() ?? "";
                var htmlUrl = root.GetProperty("html_url").GetString() ?? "";
                var body = root.TryGetProperty("body", out var bodyElement) 
                    ? bodyElement.GetString() ?? "" 
                    : "";

                // Strip leading 'v' from tag if present
                var latestVersion = tagName.StartsWith("v", StringComparison.OrdinalIgnoreCase) 
                    ? tagName.Substring(1) 
                    : tagName;

                var currentVersion = GetCurrentVersion();

                // Compare versions
                var updateAvailable = false;
                if (Version.TryParse(currentVersion, out var current) && 
                    Version.TryParse(latestVersion, out var latest))
                {
                    updateAvailable = latest.CompareTo(current) > 0;
                }

                return new UpdateInfo
                {
                    UpdateAvailable = updateAvailable,
                    CurrentVersion = currentVersion,
                    LatestVersion = latestVersion,
                    ReleaseUrl = htmlUrl,
                    ReleaseNotes = body
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Update check failed: {ex.Message}");
                return null;
            }
        }
    }
}
