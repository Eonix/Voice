using System;
using System.Diagnostics;
using System.IO;
using System.IO.IsolatedStorage;
using System.Linq;
using System.Windows.Forms;
using Voice.Properties;

namespace Voice
{
    public static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main()
        {
            const string logFileName = "Trace.log";
            const string oldLogFileName = "Trace.old.log";

            using (var store = IsolatedStorageFile.GetStore(IsolatedStorageScope.User | IsolatedStorageScope.Domain | IsolatedStorageScope.Assembly, null, null))
            {
                if (store.FileExists(logFileName) && store.GetCreationTime(logFileName).AddDays(1) < DateTime.Now)
                {
                    if (store.FileExists(oldLogFileName))
                        store.DeleteFile(oldLogFileName);

                    store.MoveFile(logFileName, oldLogFileName);
                }
            }

            var fileStream = new IsolatedStorageFileStream(logFileName, FileMode.Append, FileAccess.Write, FileShare.Read);
            
            using (var traceListener = new TextWriterTraceListener(fileStream))
            {
                Trace.Listeners.Add(traceListener);
                Trace.AutoFlush = true;

                if (Settings.Default.UpgradeRequired)
                {
                    Settings.Default.Upgrade();
                    Settings.Default.UpgradeRequired = false;
                    Settings.Default.Save();
                }

                Trace.WriteLine(DateTime.Now + ": Application started.");

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainApplicationContext());
            }
        }
    }
}
