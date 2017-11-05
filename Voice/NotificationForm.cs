using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Voice
{
    public class NotificationForm : Form
    {
        public Action<string> ClipboardUpdate;
        private IntPtr nextClipboardViewer;

        public NotificationForm()
        {
            nextClipboardViewer = (IntPtr)NativeMethods.SetClipboardViewer((int)Handle);
        }

        protected override void WndProc(ref Message m)
        {
            // defined in winuser.h
            const int drawClipboard = 0x308; // WM_DRAWCLIPBOARD
            const int changeClipboardChain = 0x030D; // WM_CHANGECBCHAIN

            switch (m.Msg)
            {
                case drawClipboard:
                    ClipboardUpdate?.Invoke(Clipboard.GetText());
                    Marshal.ThrowExceptionForHR(NativeMethods.SendMessage(nextClipboardViewer, m.Msg, m.WParam, m.LParam));
                    break;

                case changeClipboardChain:
                    if (m.WParam == nextClipboardViewer)
                        nextClipboardViewer = m.LParam;
                    else
                        Marshal.ThrowExceptionForHR(NativeMethods.SendMessage(nextClipboardViewer, m.Msg, m.WParam, m.LParam));
                    break;
                default:
                    base.WndProc(ref m);
                    break;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                NativeMethods.ChangeClipboardChain(this.Handle, nextClipboardViewer);

            base.Dispose(disposing);
        }
    }
}
