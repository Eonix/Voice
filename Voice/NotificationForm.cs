using System;
using System.Windows.Forms;

namespace Voice
{
    public class NotificationForm : Form
    {
        public event EventHandler<IDataObject> ClipboardUpdate;
        private IntPtr nextClipboardViewer;

        public NotificationForm()
        {
            nextClipboardViewer = (IntPtr)NativeMethods.SetClipboardViewer((int)Handle);
        }

        protected override void WndProc(ref Message m)
        {
            // defined in winuser.h
            const int WM_DRAWCLIPBOARD = 0x308;
            const int WM_CHANGECBCHAIN = 0x030D;

            switch (m.Msg)
            {
                case WM_DRAWCLIPBOARD:
                    ClipboardUpdate?.Invoke(null, Clipboard.GetDataObject());
                    NativeMethods.SendMessage(nextClipboardViewer, m.Msg, m.WParam, m.LParam);
                    break;

                case WM_CHANGECBCHAIN:
                    if (m.WParam == nextClipboardViewer)
                        nextClipboardViewer = m.LParam;
                    else
                        NativeMethods.SendMessage(nextClipboardViewer, m.Msg, m.WParam, m.LParam);
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
