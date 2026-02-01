using System;
using System.Windows.Forms;

namespace RegistryExpert
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            
            using (var mainForm = new MainForm())
            {
                Application.Run(mainForm);
            }
            
            // Force exit after the message loop terminates.
            // This handles edge cases where SystemEvents or COM objects keep the process alive.
            Environment.Exit(0);
        }
    }
}
