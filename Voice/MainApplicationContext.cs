using System;
using System.ComponentModel;
using System.Drawing;
using System.Speech.Synthesis;
using System.Windows.Forms;

namespace Voice
{
    public class MainApplicationContext : ApplicationContext
    {
        private readonly Container _components;
        private readonly NotifyIcon _notifyIcon;
        private readonly SpeechSynthesizer _speechSynthesizer;

        private bool _listening = true;

        public MainApplicationContext()
        {
            _speechSynthesizer = new SpeechSynthesizer();
            _speechSynthesizer.SelectVoice("IVONA 2 Amy");
            _speechSynthesizer.Rate = 5;

            _components = new Container();
            _notifyIcon = new NotifyIcon(_components)
            {
                ContextMenuStrip = new ContextMenuStrip(),
                Icon = SystemIcons.WinLogo,
                Text = @"Voice",
                Visible = true
            };

            var menuItems = _notifyIcon.ContextMenuStrip.Items;

            var voicesMenuItem = new ToolStripMenuItem("Voices");
            
            foreach (var installedVoice in _speechSynthesizer.GetInstalledVoices())
            {
                voicesMenuItem.DropDownItems.Add(new ToolStripMenuItem(installedVoice.VoiceInfo.Name, null, OnVoiceItemClick)
                {
                    Checked = _speechSynthesizer.Voice.Equals(installedVoice.VoiceInfo)
                });
            }

            menuItems.Add(voicesMenuItem);
            menuItems.Add(new ToolStripMenuItem("Listening", null, OnListeningClick) {Checked = _listening});
            menuItems.Add(new ToolStripSeparator());
            menuItems.Add(new ToolStripMenuItem("Exit", null, OnExitClick));

            ClipboardNotification.ClipboardUpdate += ClipboardNotificationOnClipboardUpdate;
        }

        private void OnVoiceItemClick(object sender, EventArgs eventArgs)
        {
            var toolStripMenuItem = (ToolStripMenuItem)sender;
            _speechSynthesizer.SelectVoice(toolStripMenuItem.Text);

            foreach (ToolStripMenuItem item in toolStripMenuItem.Owner.Items)
            {
                item.Checked = _speechSynthesizer.Voice.Name == item.Text;
            }
        }

        private void OnListeningClick(object sender, EventArgs eventArgs)
        {
            var toolStripMenuItem = (ToolStripMenuItem) sender;
            _listening = !_listening;
            toolStripMenuItem.Checked = _listening;
        }

        private void ClipboardNotificationOnClipboardUpdate(object sender, EventArgs eventArgs)
        {
            if (!_listening)
                return;

            var text = Clipboard.GetText();
            if (string.IsNullOrWhiteSpace(text))
                return;

            _speechSynthesizer.SpeakAsyncCancelAll();
            _speechSynthesizer.SpeakAsync(text);
        }

        private void OnExitClick(object sender, EventArgs eventArgs)
        {
            ExitThread();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _components?.Dispose();
                _speechSynthesizer?.Dispose();
            }

            base.Dispose(disposing);
        }

        protected override void ExitThreadCore()
        {
            _notifyIcon.Visible = false;

            base.ExitThreadCore();
        }
    }
}
