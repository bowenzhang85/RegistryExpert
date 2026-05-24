using System;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace RegistryExpert.Core
{
    /// <summary>Indicates the kind of payload referenced by an UpdateInfo's DownloadUrl.</summary>
    public enum DownloadKind
    {
        /// <summary>Portable single-file RegistryExpert.exe (legacy / fallback path).</summary>
        PortableExe = 0,
        /// <summary>Inno Setup installer RegistryExpert-Setup.exe (preferred when available).</summary>
        Installer = 1,
    }

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
        public string DownloadUrl { get; init; } = "";
        public long DownloadSize { get; init; }
        public DateTimeOffset? PublishedAt { get; init; }

        /// <summary>
        /// Indicates whether DownloadUrl points to the installer or the portable exe.
        /// AutoUpdater branches on this to choose between silent installer invocation
        /// and the legacy in-place batch-script swap.
        /// </summary>
        public DownloadKind DownloadKind { get; init; } = DownloadKind.PortableExe;
    }

    /// <summary>
    /// Helper class for checking for application updates via GitHub Releases API.
    /// </summary>
    public static class UpdateChecker
    {
        private const string DefaultGitHubApiUrl = "https://api.github.com/repos/bowenzhang85/RegistryExpert/releases/latest";

        /// <summary>
        /// Base of the GitHub releases API (everything before "/latest" or "/tags/...").
        /// Derived from GitHubApiUrl so the override env var also redirects tag lookups.
        /// </summary>
        private static string GitHubReleasesApiBase
        {
            get
            {
                var url = GitHubApiUrl;
                const string latestSuffix = "/latest";
                if (url.EndsWith(latestSuffix, StringComparison.OrdinalIgnoreCase))
                    return url.Substring(0, url.Length - latestSuffix.Length);
                // Fall back to trimming everything after "/releases"
                var idx = url.IndexOf("/releases", StringComparison.OrdinalIgnoreCase);
                return idx > 0 ? url.Substring(0, idx + "/releases".Length) : url;
            }
        }

        /// <summary>
        /// The endpoint queried for release info. Can be overridden via the
        /// REGEXPERT_UPDATE_URL environment variable so a private/test repo
        /// (e.g. RegistryExpert-test) can be targeted without rebuilding.
        /// </summary>
        public static string GitHubApiUrl { get; } =
            Environment.GetEnvironmentVariable("REGEXPERT_UPDATE_URL") is { Length: > 0 } overrideUrl
                ? overrideUrl
                : DefaultGitHubApiUrl;

        public static bool IsUsingOverrideUrl =>
            !string.Equals(GitHubApiUrl, DefaultGitHubApiUrl, StringComparison.OrdinalIgnoreCase);

        private static readonly HttpClient _httpClient;

        static UpdateChecker()
        {
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "RegistryExpert-UpdateChecker");
            _httpClient.DefaultRequestHeaders.Add("Accept", "application/vnd.github.v3+json");
            _httpClient.Timeout = TimeSpan.FromSeconds(15);
        }

        /// <summary>
        /// Gets the current application version as a string (e.g., "1.0.1").
        /// </summary>
        public static string GetCurrentVersion()
        {
            var version = (System.Reflection.Assembly.GetEntryAssembly() ?? System.Reflection.Assembly.GetExecutingAssembly()).GetName().Version;
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
                using var response = await _httpClient.GetAsync(GitHubApiUrl).ConfigureAwait(false);

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

                // Pick the preferred download asset (installer first, then portable).
                var (downloadUrl, downloadSize, downloadKind) = ParseAssetsFromRoot(root);

                return new UpdateInfo
                {
                    UpdateAvailable = updateAvailable,
                    CurrentVersion = currentVersion,
                    LatestVersion = latestVersion,
                    ReleaseUrl = htmlUrl,
                    ReleaseNotes = body,
                    DownloadUrl = downloadUrl,
                    DownloadSize = downloadSize,
                    PublishedAt = TryParsePublishedAt(root),
                    DownloadKind = downloadKind
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Update check failed: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Fetch release info for a specific tag (e.g. "v2.2.1"). Returns null on failure.
        /// Used by the in-app "What's New" / release notes window to look up the currently
        /// installed version regardless of whether it is the latest.
        /// </summary>
        public static async Task<UpdateInfo?> GetReleaseByTagAsync(string tag)
        {
            if (string.IsNullOrWhiteSpace(tag)) return null;

            try
            {
                var url = $"{GitHubReleasesApiBase}/tags/{Uri.EscapeDataString(tag)}";
                using var response = await _httpClient.GetAsync(url).ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                {
                    System.Diagnostics.Debug.WriteLine($"GitHub API (tags/{tag}) returned {response.StatusCode}");
                    return null;
                }

                var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                var tagName = root.TryGetProperty("tag_name", out var t) ? t.GetString() ?? tag : tag;
                var htmlUrl = root.TryGetProperty("html_url", out var h) ? h.GetString() ?? "" : "";
                var body = root.TryGetProperty("body", out var b) ? b.GetString() ?? "" : "";

                var version = tagName.StartsWith("v", StringComparison.OrdinalIgnoreCase)
                    ? tagName.Substring(1) : tagName;

                // Populate the download fields from the release's assets too —
                // the migration code path (TryMigrateToInstallerAsync) calls
                // GetReleaseByTagAsync to look up the installer for its current
                // version, so we must surface DownloadUrl/Size/Kind here.
                var (downloadUrl, downloadSize, downloadKind) = ParseAssetsFromRoot(root);

                return new UpdateInfo
                {
                    UpdateAvailable = false, // not a comparison; just a lookup
                    CurrentVersion = GetCurrentVersion(),
                    LatestVersion = version,
                    ReleaseUrl = htmlUrl,
                    ReleaseNotes = body,
                    DownloadUrl = downloadUrl,
                    DownloadSize = downloadSize,
                    PublishedAt = TryParsePublishedAt(root),
                    DownloadKind = downloadKind
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetReleaseByTagAsync({tag}) failed: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Scan the "assets" array of a GitHub release JSON for our preferred payload.
        /// Prefer the Inno Setup installer (filename matches "RegistryExpert-installer-v*.exe")
        /// when present; fall back to RegistryExpert.exe (portable) for older releases that
        /// don't yet ship the installer. Returns ("", 0, PortableExe) when neither matches.
        /// </summary>
        private static (string Url, long Size, DownloadKind Kind) ParseAssetsFromRoot(JsonElement root)
        {
            string downloadUrl = "";
            long downloadSize = 0;
            DownloadKind downloadKind = DownloadKind.PortableExe;

            if (root.TryGetProperty("assets", out var assets) && assets.ValueKind == JsonValueKind.Array)
            {
                string portableUrl = "";
                long portableSize = 0;

                foreach (var asset in assets.EnumerateArray())
                {
                    var name = asset.TryGetProperty("name", out var n) ? n.GetString() : null;
                    if (string.IsNullOrEmpty(name)) continue;

                    var url = asset.TryGetProperty("browser_download_url", out var u)
                        ? u.GetString() ?? "" : "";
                    var size = asset.TryGetProperty("size", out var s) && s.TryGetInt64(out var sz)
                        ? sz : 0;

                    if (IsInstallerAssetName(name))
                    {
                        downloadUrl = url;
                        downloadSize = size;
                        downloadKind = DownloadKind.Installer;
                        // Installer is preferred — stop scanning; we don't need the portable fallback.
                        break;
                    }
                    else if (string.Equals(name, "RegistryExpert.exe", StringComparison.OrdinalIgnoreCase))
                    {
                        portableUrl = url;
                        portableSize = size;
                    }
                }

                if (string.IsNullOrEmpty(downloadUrl) && !string.IsNullOrEmpty(portableUrl))
                {
                    downloadUrl = portableUrl;
                    downloadSize = portableSize;
                    downloadKind = DownloadKind.PortableExe;
                }
            }

            return (downloadUrl, downloadSize, downloadKind);
        }

        /// <summary>
        /// Returns true if a release-asset filename matches our installer naming convention.
        /// Accepts both the current versioned form (e.g. "RegistryExpert-installer-v2.3.0.exe")
        /// and the legacy unversioned form (e.g. "RegistryExpert-Setup.exe") for backward
        /// compatibility with any local builds or test releases that used the older name.
        /// </summary>
        private static bool IsInstallerAssetName(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            if (string.Equals(name, "RegistryExpert-Setup.exe", StringComparison.OrdinalIgnoreCase))
                return true;
            return name.StartsWith("RegistryExpert-installer-v", StringComparison.OrdinalIgnoreCase)
                && name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase);
        }

        private static DateTimeOffset? TryParsePublishedAt(JsonElement root)
        {
            if (root.TryGetProperty("published_at", out var pa)
                && pa.ValueKind == JsonValueKind.String
                && DateTimeOffset.TryParse(pa.GetString(), out var dto))
            {
                return dto;
            }
            return null;
        }

        /// <summary>
        /// HttpClient instance shared with AutoUpdater for download requests
        /// (so the same User-Agent header is reused).
        /// </summary>
        internal static HttpClient HttpClient => _httpClient;
    }
}
