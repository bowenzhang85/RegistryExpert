using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace RegistryExpert.Core
{
    /// <summary>
    /// Handles downloading new RegistryExpert.exe builds from a GitHub release
    /// and replacing the running executable via a small detached batch script.
    /// </summary>
    public static class AutoUpdater
    {
        private const string CacheFolderName = "RegistryExpert";
        private const string ProcessName = "RegistryExpert.exe";
        private const string PortableFileName = "RegistryExpert.exe";
        // Cache filename for installer downloads. Generic — the on-disk cache
        // doesn't need to match the GitHub asset name (which is versioned,
        // e.g. RegistryExpert-installer-v2.3.0.exe). The per-version cache
        // folder already disambiguates between releases.
        private const string InstallerFileName = "RegistryExpert-installer.exe";

        /// <summary>
        /// Returns the cache path used to store the downloaded payload for a
        /// specific version, e.g. %TEMP%\RegistryExpert\1.2.3\RegistryExpert-Setup.exe
        /// (installer) or %TEMP%\RegistryExpert\1.2.3\RegistryExpert.exe (portable).
        /// </summary>
        public static string GetDownloadCachePath(string version, DownloadKind kind = DownloadKind.PortableExe)
        {
            var dir = Path.Combine(Path.GetTempPath(), CacheFolderName, version);
            var fileName = kind == DownloadKind.Installer ? InstallerFileName : PortableFileName;
            return Path.Combine(dir, fileName);
        }

        /// <summary>
        /// Returns true when the cached file already exists and has the expected size.
        /// </summary>
        public static bool IsUpdateAlreadyDownloaded(UpdateInfo info)
        {
            if (string.IsNullOrEmpty(info.LatestVersion) || info.DownloadSize <= 0)
                return false;

            var path = GetDownloadCachePath(info.LatestVersion, info.DownloadKind);
            try
            {
                var fi = new FileInfo(path);
                return fi.Exists && fi.Length == info.DownloadSize && VerifyDownload(path, info.DownloadSize);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Downloads the new RegistryExpert.exe to the per-version temp cache.
        /// Returns the local path on success, or null on failure / cancellation.
        /// </summary>
        public static async Task<string?> DownloadUpdateAsync(
            UpdateInfo info,
            IProgress<double>? progress = null,
            CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrEmpty(info.DownloadUrl) || info.DownloadSize <= 0)
            {
                Debug.WriteLine("AutoUpdater: missing DownloadUrl or DownloadSize");
                return null;
            }

            var targetPath = GetDownloadCachePath(info.LatestVersion, info.DownloadKind);
            var targetDir = Path.GetDirectoryName(targetPath)!;
            try
            {
                Directory.CreateDirectory(targetDir);

                // Reuse partial download if already complete & valid
                if (IsUpdateAlreadyDownloaded(info))
                {
                    progress?.Report(1.0);
                    return targetPath;
                }

                // Always overwrite any prior partial file
                if (File.Exists(targetPath))
                    File.Delete(targetPath);

                using var response = await UpdateChecker.HttpClient
                    .GetAsync(info.DownloadUrl, HttpCompletionOption.ResponseHeadersRead, cancellationToken)
                    .ConfigureAwait(false);

                if (!response.IsSuccessStatusCode)
                {
                    Debug.WriteLine($"AutoUpdater: download HTTP {response.StatusCode}");
                    return null;
                }

                long total = response.Content.Headers.ContentLength ?? info.DownloadSize;
                long received = 0;

                using (var inStream = await response.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false))
                using (var outStream = new FileStream(targetPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, useAsync: true))
                {
                    var buffer = new byte[81920];
                    int read;
                    while ((read = await inStream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false)) > 0)
                    {
                        await outStream.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
                        received += read;
                        if (total > 0)
                            progress?.Report(Math.Min(1.0, (double)received / total));
                    }
                }

                if (!VerifyDownload(targetPath, info.DownloadSize))
                {
                    Debug.WriteLine("AutoUpdater: verification failed; deleting download");
                    SafeDelete(targetPath);
                    return null;
                }

                progress?.Report(1.0);
                return targetPath;
            }
            catch (OperationCanceledException)
            {
                SafeDelete(targetPath);
                return null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"AutoUpdater: download error: {ex.Message}");
                SafeDelete(targetPath);
                return null;
            }
        }

        /// <summary>
        /// Validates the downloaded file matches the expected size and is a
        /// well-formed Windows PE executable (MZ header + PE\0\0 signature).
        /// </summary>
        public static bool VerifyDownload(string path, long expectedSize)
        {
            try
            {
                var fi = new FileInfo(path);
                if (!fi.Exists || fi.Length != expectedSize) return false;

                using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                using var br = new BinaryReader(fs);

                // DOS header: 'M','Z'
                if (br.ReadByte() != (byte)'M' || br.ReadByte() != (byte)'Z')
                    return false;

                // e_lfanew is at offset 0x3C
                fs.Seek(0x3C, SeekOrigin.Begin);
                int peOffset = br.ReadInt32();
                if (peOffset <= 0 || peOffset > fs.Length - 4) return false;

                fs.Seek(peOffset, SeekOrigin.Begin);
                if (br.ReadByte() != (byte)'P' || br.ReadByte() != (byte)'E' ||
                    br.ReadByte() != 0 || br.ReadByte() != 0)
                    return false;

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Returns true if the current process can write to the directory that
        /// contains the running executable (i.e. no UAC elevation needed).
        /// </summary>
        public static bool CanWriteToInstallLocation(string exePath)
        {
            try
            {
                var dir = Path.GetDirectoryName(exePath);
                if (string.IsNullOrEmpty(dir)) return false;

                var probe = Path.Combine(dir, $".regexpert_write_probe_{Guid.NewGuid():N}");
                File.WriteAllBytes(probe, Array.Empty<byte>());
                File.Delete(probe);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Returns the full path of the currently running executable, even when
        /// PublishSingleFile is enabled (Assembly.Location can be empty in that case).
        /// </summary>
        public static string GetCurrentExecutablePath()
        {
            var p = Process.GetCurrentProcess().MainModule?.FileName;
            if (!string.IsNullOrEmpty(p)) return p!;
            return Environment.ProcessPath ?? AppContext.BaseDirectory;
        }

        /// <summary>
        /// Writes the helper batch script to %TEMP%\RegistryExpert\update.cmd and
        /// launches it (optionally elevated). Caller should immediately exit the
        /// application so the script can replace the .exe.
        /// </summary>
        /// <param name="newExePath">Local path to the downloaded new build.</param>
        /// <param name="currentExePath">Path of the running .exe to be replaced.</param>
        /// <param name="elevated">If true, request UAC elevation when launching the script.</param>
        /// <param name="currentVersionForArg">
        /// Version string of the currently running build (e.g. "2.1.0"). The
        /// script will pass this to the relaunched new build via
        /// "--just-updated &lt;version&gt;" so it can show the success banner.
        /// </param>
        /// <returns>true if the script was successfully launched.</returns>
        public static bool LaunchUpdaterAndExit(string newExePath, string currentExePath, bool elevated, string currentVersionForArg)
        {
            try
            {
                var scriptDir = Path.Combine(Path.GetTempPath(), CacheFolderName);
                Directory.CreateDirectory(scriptDir);
                var scriptPath = Path.Combine(scriptDir, "update.cmd");

                // %~1 = newExePath, %~2 = currentExePath, %~3 = currentVersionForArg
                var script =
                    "@echo off\r\n" +
                    "setlocal\r\n" +
                    "set SRC=%~1\r\n" +
                    "set DST=%~2\r\n" +
                    "set FROMVER=%~3\r\n" +
                    ":wait\r\n" +
                    $"tasklist /FI \"IMAGENAME eq {ProcessName}\" 2>NUL | find /I \"{ProcessName}\" >NUL\r\n" +
                    "if not errorlevel 1 (timeout /t 1 /nobreak >NUL & goto wait)\r\n" +
                    "move /Y \"%SRC%\" \"%DST%\"\r\n" +
                    "if errorlevel 1 (\r\n" +
                    "  echo Update failed: could not replace target file. >> \"%TEMP%\\RegistryExpert\\update.log\"\r\n" +
                    "  exit /b 1\r\n" +
                    ")\r\n" +
                    "start \"\" \"%DST%\" --just-updated %FROMVER%\r\n" +
                    "(goto) 2>nul & del \"%~f0\"\r\n";

                File.WriteAllText(scriptPath, script);

                // Sanitise version arg to avoid shell injection (only digits, dots, hyphens, letters)
                var safeVer = string.IsNullOrEmpty(currentVersionForArg)
                    ? "unknown"
                    : new string(currentVersionForArg.Where(c => char.IsLetterOrDigit(c) || c == '.' || c == '-').ToArray());
                if (string.IsNullOrEmpty(safeVer)) safeVer = "unknown";

                var psi = new ProcessStartInfo
                {
                    FileName = scriptPath,
                    Arguments = $"\"{newExePath}\" \"{currentExePath}\" {safeVer}",
                    UseShellExecute = true,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };
                if (elevated)
                    psi.Verb = "runas";

                Process.Start(psi);
                return true;
            }
            catch (Win32Exception wex)
            {
                // 1223 = user cancelled UAC
                Debug.WriteLine($"AutoUpdater: LaunchUpdater failed (Win32: {wex.NativeErrorCode}): {wex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"AutoUpdater: LaunchUpdater failed: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Launches the Inno Setup installer silently and exits the current process.
        /// The installer's /CLOSEAPPLICATIONS flag will wait for our process to
        /// terminate (Restart Manager), then write the new files and relaunch
        /// via the [Run] section in the .iss script with "--just-updated &lt;version&gt;".
        /// </summary>
        /// <param name="setupExePath">Local path to the downloaded RegistryExpert-Setup.exe.</param>
        /// <param name="currentVersionForArg">
        /// Version of the currently running build; passed to the installer's [Run]
        /// section so the relaunched new build can show the post-update banner.
        /// </param>
        /// <returns>true if the installer process was launched successfully.</returns>
        public static bool LaunchInstallerAndExit(string setupExePath, string currentVersionForArg)
        {
            try
            {
                if (string.IsNullOrEmpty(setupExePath) || !File.Exists(setupExePath))
                {
                    Debug.WriteLine($"AutoUpdater: installer file not found: {setupExePath}");
                    return false;
                }

                // Sanitize the from-version arg to avoid shell injection. Inno Setup's
                // {param:...} substitution permits letters, digits, dots, hyphens.
                var safeVer = string.IsNullOrEmpty(currentVersionForArg)
                    ? "unknown"
                    : new string(currentVersionForArg.Where(c => char.IsLetterOrDigit(c) || c == '.' || c == '-').ToArray());
                if (string.IsNullOrEmpty(safeVer)) safeVer = "unknown";

                var args =
                    "/VERYSILENT " +
                    "/SUPPRESSMSGBOXES " +
                    "/CLOSEAPPLICATIONS " +
                    "/NORESTART " +
                    $"/fromversion={safeVer}";

                var psi = new ProcessStartInfo
                {
                    FileName = setupExePath,
                    Arguments = args,
                    UseShellExecute = true,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                Process.Start(psi);
                return true;
            }
            catch (Win32Exception wex)
            {
                Debug.WriteLine($"AutoUpdater: LaunchInstaller failed (Win32: {wex.NativeErrorCode}): {wex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"AutoUpdater: LaunchInstaller failed: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Removes any cached download folders other than the one for currentVersion.
        /// Best-effort; failures are silently ignored.
        /// </summary>
        public static void CleanupOldCaches(string currentVersion)
        {
            try
            {
                var root = Path.Combine(Path.GetTempPath(), CacheFolderName);
                if (!Directory.Exists(root)) return;

                foreach (var dir in Directory.GetDirectories(root))
                {
                    var name = Path.GetFileName(dir);
                    if (string.Equals(name, currentVersion, StringComparison.OrdinalIgnoreCase))
                        continue;
                    // Only delete folders that look like version numbers (e.g. "1.2.3")
                    if (Version.TryParse(name, out _))
                    {
                        try { Directory.Delete(dir, recursive: true); } catch { /* ignore */ }
                    }
                }
            }
            catch { /* ignore */ }
        }

        private static void SafeDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { /* ignore */ }
        }
    }
}
