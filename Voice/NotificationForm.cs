using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Voice
{
    public class NotificationForm : Form
    {
        private static class NativeMethods
        {
            /// <summary>
            /// Places the given window in the system-maintained clipboard format listener list.
            /// </summary>
            [DllImport("User32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool AddClipboardFormatListener(IntPtr hwnd);

            /// <summary>
            /// Removes the given window from the system-maintained clipboard format listener list.
            /// </summary>
            [DllImport("User32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool RemoveClipboardFormatListener(IntPtr hwnd);
        }

        public Action<string> ClipboardUpdate;

        public NotificationForm()
        {
            NativeMethods.AddClipboardFormatListener(Handle);
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            const int WM_CLIPBOARDUPDATE = 0x031D; // Sent when the contents of the clipboard have changed.

            if (m.Msg == WM_CLIPBOARDUPDATE)
                ClipboardUpdate?.Invoke(GetClipboardText());
        }

        private static string GetClipboardText()
        {
            // It seems that faster computers may recieve the Clipboard event before it is available
            // in the clipboard api, so we will poll the OS a couple of times to ensure 
            // that there is in fact no text.
            const int retries = 10;
            for (int i = 0; i < retries; i++)
            {
                var text = Clipboard.GetText();
                if (!string.IsNullOrEmpty(text))
                    return text;

                Thread.Sleep(10);
            }

            return string.Empty; // It was probably not text that was added to the clipboard.
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                NativeMethods.RemoveClipboardFormatListener(Handle);

            base.Dispose(disposing);
        }
    }
}