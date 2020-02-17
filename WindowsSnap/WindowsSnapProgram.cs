using System;
using System.Windows.Forms;

namespace WindowsSnap {
    static class WindowsSnapProgram {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main() {
            try
            {
                Logger.Log("Windows Snap started.");
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new TrayIconForm());
            }
            catch (Exception ex)
            {
                Logger.Log("Fatal error! " + ex);
            }
        }
    }
}
