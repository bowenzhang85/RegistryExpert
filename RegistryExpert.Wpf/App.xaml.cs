using System.Windows;
using System.Windows.Threading;

namespace RegistryExpert.Wpf
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

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
