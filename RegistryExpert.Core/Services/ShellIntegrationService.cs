using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.Win32;

namespace RegistryExpert.Core
{
    /// <summary>
    /// Registers / unregisters the "Open with Registry Expert" right-click verb
    /// in Windows Explorer for offline registry hive files.
    ///
    /// Writes under HKCU (no admin required). Filters via the shell's AppliesTo
    /// query syntax to only show on known hive filenames — SYSTEM, SOFTWARE,
    /// NTUSER.DAT, etc. — instead of polluting every file's context menu.
    ///
    /// On Windows 11 the verb appears in the legacy menu accessed via
    /// "Show more options" or Shift+F10 (a limitation of non-MSIX,
    /// non-IExplorerCommand shell verbs).
    /// </summary>
    public static class ShellIntegrationService
    {
        /// <summary>Subkey name under HKCU\Software\Classes\*\shell.</summary>
        public const string VerbName = "OpenWithRegistryExpert";

        /// <summary>Display text shown on the context menu item.</summary>
        public const string VerbDisplayText = "Open with Registry Expert";

        private static readonly string RegPath = $@"Software\Classes\*\shell\{VerbName}";
        private static readonly string ShellRoot = @"Software\Classes\*\shell";

        /// <summary>Returns true when the registry verb is currently registered for the current user.</summary>
        public static bool IsRegistered()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RegPath);
                return key != null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ShellIntegrationService.IsRegistered failed: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Returns the exe path currently stored in the registered command value
        /// (parsed out of "&quot;C:\path\to\exe&quot; &quot;%1&quot;"), or null
        /// if the verb is not registered or the value cannot be parsed.
        /// </summary>
        public static string? GetRegisteredExePath()
        {
            try
            {
                using var cmdKey = Registry.CurrentUser.OpenSubKey($@"{RegPath}\command");
                var raw = cmdKey?.GetValue(null) as string;
                if (string.IsNullOrEmpty(raw)) return null;

                // Expected form: "<exe>" "%1"
                if (raw.Length > 0 && raw[0] == '"')
                {
                    var endQuote = raw.IndexOf('"', 1);
                    if (endQuote > 1) return raw.Substring(1, endQuote - 1);
                }

                // Fallback: take everything up to the first space
                var space = raw.IndexOf(' ');
                return space > 0 ? raw.Substring(0, space) : raw;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ShellIntegrationService.GetRegisteredExePath failed: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Write the shell verb under HKCU. Pass the canonical exact-match filenames
        /// (e.g. SYSTEM, SOFTWARE) and prefix matches (e.g. NTUSER, USRCLASS) so the
        /// AppliesTo filter limits the menu item to actual hive files.
        /// Idempotent — safe to call repeatedly.
        /// </summary>
        public static void Register(
            string exePath,
            IEnumerable<string> exactNames,
            IEnumerable<string> prefixes)
        {
            if (string.IsNullOrWhiteSpace(exePath))
                throw new ArgumentException("exePath is required", nameof(exePath));

            var appliesTo = BuildAppliesTo(exactNames, prefixes);

            using (var key = Registry.CurrentUser.CreateSubKey(RegPath, writable: true)
                ?? throw new InvalidOperationException($"Could not create HKCU\\{RegPath}"))
            {
                key.SetValue(null, VerbDisplayText, RegistryValueKind.String);
                key.SetValue("Icon", $"\"{exePath}\",0", RegistryValueKind.String);
                key.SetValue("AppliesTo", appliesTo, RegistryValueKind.String);
            }

            using (var cmdKey = Registry.CurrentUser.CreateSubKey($@"{RegPath}\command", writable: true)
                ?? throw new InvalidOperationException($"Could not create HKCU\\{RegPath}\\command"))
            {
                cmdKey.SetValue(null, $"\"{exePath}\" \"%1\"", RegistryValueKind.String);
            }
        }

        /// <summary>
        /// Remove the shell verb (and its command subkey) from HKCU.
        /// Safe to call even when not currently registered.
        /// </summary>
        public static void Unregister()
        {
            try
            {
                using var shellKey = Registry.CurrentUser.OpenSubKey(ShellRoot, writable: true);
                shellKey?.DeleteSubKeyTree(VerbName, throwOnMissingSubKey: false);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ShellIntegrationService.Unregister failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Build an AppliesTo query string that matches the supplied exact filenames
        /// (case-insensitive equality) and prefix patterns (starts-with). For example:
        /// <c>System.FileName:="SYSTEM" OR System.FileName:~&lt;"NTUSER"</c>.
        /// </summary>
        private static string BuildAppliesTo(IEnumerable<string> exactNames, IEnumerable<string> prefixes)
        {
            var clauses = new List<string>();

            foreach (var name in exactNames ?? Enumerable.Empty<string>())
            {
                if (!string.IsNullOrWhiteSpace(name))
                    clauses.Add($"System.FileName:=\"{name}\"");
            }

            foreach (var prefix in prefixes ?? Enumerable.Empty<string>())
            {
                if (!string.IsNullOrWhiteSpace(prefix))
                    clauses.Add($"System.FileName:~<\"{prefix}\"");
            }

            return string.Join(" OR ", clauses);
        }
    }
}
