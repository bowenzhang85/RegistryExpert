using System;
using System.Windows;
using System.Windows.Threading;

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

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Detect post-auto-update relaunch: "--just-updated <oldVersion>"
            var idx = Array.IndexOf(e.Args, "--just-updated");
            if (idx >= 0 && idx + 1 < e.Args.Length)
            {
                UpgradedFromVersion = e.Args[idx + 1];
            }

            // Register code page encoding support (required by Lib/Registry parser)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Catch unhandled exceptions so the app shows a message instead of silently crashing
            DispatcherUnhandledException += OnDispatcherUnhandledException;
        }

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
    }
}
