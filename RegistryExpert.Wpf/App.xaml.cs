using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using RegistryExpert.Core;

namespace RegistryExpert.Wpf
{
    public partial class App : Application
    {
        /// <summary>
        /// When set, indicates the previous installed version that the auto-updater
        /// just upgraded from. Populated from the "--just-updated &lt;version&gt;"
        /// command-line argument written by AutoUpdater.LaunchUpdaterAndExit.
        /// MainWindow uses this to show the post-update success banner.
        /// </summary>
        public static string? UpgradedFromVersion { get; private set; }

        /// <summary>
        /// Positional file path arguments handed to the app at startup. Populated
        /// by Windows Explorer when the user opens a hive via the right-click
        /// "Open with Registry Expert" verb (which expands to "exe.exe" "%1").
        /// MainWindow processes this list after the window is ready.
        /// </summary>
        public static IReadOnlyList<string> StartupFilePaths { get; private set; } = Array.Empty<string>();

        /// <summary>
        /// Raised when a second instance forwards file paths via the named pipe.
        /// MainWindow subscribes to this to load the hives into the already-running window.
        /// </summary>
        public static event Action<IReadOnlyList<string>>? RemoteOpenRequested;

        /// <summary>Raised when a second instance asks us to come to the foreground.</summary>
        public static event Action? RemoteActivateRequested;

        // ── Single instance state ────────────────────────────────────────

        private static Mutex? _singleInstanceMutex;
        private static CancellationTokenSource? _pipeServerCts;
        private static Task? _pipeServerTask;

        // Per-user pipe name; second instance connects with the same name.
        private static readonly string PipeName = $"RegistryExpert.OpenHive.{GetCurrentUserSid()}";

        // PID file lets the second instance call AllowSetForegroundWindow on us.
        private static readonly string PidFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "RegistryExpert", "instance.pid");

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            ParseArgs(e.Args);

            // Single-instance: try to acquire a user-scoped mutex. If another
            // instance is already running, forward our file paths + activate
            // request to it via named pipe and exit.
            var mutexName = $"Local\\RegistryExpert.SingleInstance.{GetCurrentUserSid()}";
            _singleInstanceMutex = new Mutex(initiallyOwned: false, name: mutexName, out _);

            bool acquired = false;
            try { acquired = _singleInstanceMutex.WaitOne(0, exitContext: false); }
            catch (AbandonedMutexException) { acquired = true; } // previous owner crashed

            if (!acquired)
            {
                ForwardToExistingInstance();
                // Tell WPF not to create a window for this process.
                Shutdown(0);
                return;
            }

            // We are the primary instance. Continue normal startup.
            WritePidFile();
            StartPipeServer();

            // Stale shell-verb cleanup. The installer now OWNS the right-click
            // "Open with Registry Expert" verb (pointing at the install location).
            // Only remove the verb if it points at a *different / stale* exe path
            // (e.g. a leftover from an old in-app dev toggle or a moved portable) —
            // never the current install. Order is safe: the installer writes the
            // verb -> install path, then launches us from that same path, so
            // registered == current -> kept.
            try
            {
                if (ShellIntegrationService.IsRegistered())
                {
                    var registered = ShellIntegrationService.GetRegisteredExePath();
                    var currentExe = Environment.ProcessPath
                        ?? System.Reflection.Assembly.GetEntryAssembly()?.Location ?? "";
                    if (!string.IsNullOrEmpty(registered)
                        && !string.IsNullOrEmpty(currentExe)
                        && !string.Equals(Path.GetFullPath(registered), Path.GetFullPath(currentExe),
                                          StringComparison.OrdinalIgnoreCase))
                    {
                        ShellIntegrationService.Unregister(); // stale leftover -> remove
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Shell-verb stale cleanup failed: {ex.Message}");
            }

            // Register code page encoding support (required by Lib/Registry parser)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Catch unhandled exceptions so the app shows a message instead of silently crashing
            DispatcherUnhandledException += OnDispatcherUnhandledException;
        }

        protected override void OnExit(ExitEventArgs e)
        {
            try
            {
                _pipeServerCts?.Cancel();
                _pipeServerCts?.Dispose();
                _pipeServerCts = null;
            }
            catch { /* best-effort */ }

            try
            {
                if (_singleInstanceMutex != null)
                {
                    try { _singleInstanceMutex.ReleaseMutex(); } catch { /* not owner */ }
                    _singleInstanceMutex.Dispose();
                    _singleInstanceMutex = null;
                }
            }
            catch { /* best-effort */ }

            try { if (File.Exists(PidFilePath)) File.Delete(PidFilePath); }
            catch { /* best-effort */ }

            base.OnExit(e);
        }

        // ── Arg parsing ────────────────────────────────────────────────────

        private static void ParseArgs(string[] args)
        {
            var files = new List<string>();
            for (int i = 0; i < args.Length; i++)
            {
                var arg = args[i];

                // --just-updated <version> : consume both
                if (string.Equals(arg, "--just-updated", StringComparison.Ordinal))
                {
                    if (i + 1 < args.Length)
                    {
                        UpgradedFromVersion = args[i + 1];
                        i++; // skip the version value
                    }
                    continue;
                }

                // Skip any other flag-style argument and its value (defensive)
                if (arg.StartsWith("--", StringComparison.Ordinal)) continue;

                // Positional arg: treat as a file path if it points at an existing file
                try
                {
                    if (File.Exists(arg)) files.Add(Path.GetFullPath(arg));
                }
                catch { /* ignore malformed paths */ }
            }

            StartupFilePaths = files;
        }

        // ── Single-instance: pipe server (primary side) ───────────────────

        private static void StartPipeServer()
        {
            _pipeServerCts = new CancellationTokenSource();
            var ct = _pipeServerCts.Token;
            _pipeServerTask = Task.Run(() => RunPipeServerLoop(ct), ct);
        }

        private static async Task RunPipeServerLoop(CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                NamedPipeServerStream? server = null;
                try
                {
                    server = new NamedPipeServerStream(
                        PipeName,
                        PipeDirection.In,
                        maxNumberOfServerInstances: 1,
                        PipeTransmissionMode.Byte,
                        PipeOptions.Asynchronous);

                    await server.WaitForConnectionAsync(ct).ConfigureAwait(false);

                    using var reader = new StreamReader(server, Encoding.UTF8);
                    var files = new List<string>();
                    bool activate = false;

                    string? line;
                    while ((line = await reader.ReadLineAsync(ct).ConfigureAwait(false)) != null)
                    {
                        if (line.StartsWith("OPEN ", StringComparison.Ordinal))
                        {
                            var path = line.Substring(5).Trim();
                            if (!string.IsNullOrEmpty(path)) files.Add(path);
                        }
                        else if (string.Equals(line, "ACTIVATE", StringComparison.Ordinal))
                        {
                            activate = true;
                        }
                    }

                    // Marshal to UI thread to fire events
                    var capturedFiles = files;
                    var capturedActivate = activate;
                    if (Current?.Dispatcher != null)
                    {
                        await Current.Dispatcher.InvokeAsync(() =>
                        {
                            if (capturedFiles.Count > 0)
                                RemoteOpenRequested?.Invoke(capturedFiles);
                            if (capturedActivate)
                                RemoteActivateRequested?.Invoke();
                        }, DispatcherPriority.Normal, ct);
                    }
                }
                catch (OperationCanceledException) { break; }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Pipe server error: {ex.Message}");
                    // Small delay before retrying to avoid hot-loop on persistent failure
                    try { await Task.Delay(250, ct).ConfigureAwait(false); }
                    catch (OperationCanceledException) { break; }
                }
                finally
                {
                    try { server?.Dispose(); } catch { }
                }
            }
        }

        // ── Single-instance: client side (second instance forwards & exits) ──

        private static void ForwardToExistingInstance()
        {
            // Step 1: best-effort foreground permission grant so the primary can
            // bring itself to the front when we activate it.
            try
            {
                if (int.TryParse(SafeReadPidFile(), out int primaryPid))
                {
                    NativeMethods.AllowSetForegroundWindow(primaryPid);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"AllowSetForegroundWindow failed: {ex.Message}");
            }

            // Step 2: open the pipe and forward our payload. ~1.5s timeout.
            try
            {
                using var client = new NamedPipeClientStream(
                    serverName: ".",
                    pipeName: PipeName,
                    direction: PipeDirection.Out,
                    options: PipeOptions.Asynchronous);

                client.Connect(timeout: 1500);

                using var writer = new StreamWriter(client, new UTF8Encoding(false)) { AutoFlush = true };
                foreach (var path in StartupFilePaths)
                {
                    writer.WriteLine("OPEN " + path);
                }
                writer.WriteLine("ACTIVATE");
            }
            catch (TimeoutException)
            {
                // Existing instance unresponsive — silently shut down to preserve
                // the single-instance invariant. User can manually retry.
                Debug.WriteLine("Pipe connect timed out; existing instance unresponsive.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Forward to existing instance failed: {ex.Message}");
            }
        }

        // ── PID file ──────────────────────────────────────────────────────

        private static void WritePidFile()
        {
            try
            {
                var dir = Path.GetDirectoryName(PidFilePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                File.WriteAllText(PidFilePath, Environment.ProcessId.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"WritePidFile failed: {ex.Message}");
            }
        }

        private static string? SafeReadPidFile()
        {
            try
            {
                return File.Exists(PidFilePath) ? File.ReadAllText(PidFilePath).Trim() : null;
            }
            catch { return null; }
        }

        // ── Helpers ───────────────────────────────────────────────────────

        private static string GetCurrentUserSid()
        {
            try
            {
                return WindowsIdentity.GetCurrent().User?.Value ?? "default";
            }
            catch { return "default"; }
        }

        // ── Unhandled exception ──────────────────────────────────────────

        private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine($"Unhandled exception: {e.Exception}");

            // For non-recoverable errors, show the message but let the app shut down
            if (e.Exception is System.AccessViolationException
                or System.BadImageFormatException
                or System.TypeInitializationException
                or System.AppDomainUnloadedException)
            {
                MessageBox.Show(
                    $"A fatal error occurred and the application must close:\n\n{e.Exception.Message}",
                    "RegistryExpert - Fatal Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                // Don't set e.Handled — let the app terminate
                return;
            }

            MessageBox.Show(
                $"An unexpected error occurred:\n\n{e.Exception.Message}",
                "RegistryExpert - Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);

            e.Handled = true;
        }

        // ── P/Invoke ──────────────────────────────────────────────────────

        private static class NativeMethods
        {
            [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
            [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
            public static extern bool AllowSetForegroundWindow(int dwProcessId);
        }
    }
}
