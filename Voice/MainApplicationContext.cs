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
            voicesMenuItem.DropDownItems.AddRange(GetVoiceItems());

            var rateMenuItem = new ToolStripMenuItem("Rate");
            rateMenuItem.DropDownItems.AddRange(GetRateItems());
            
            menuItems.Add(voicesMenuItem);
            menuItems.Add(rateMenuItem);
            menuItems.Add(new ToolStripMenuItem("Listening", null, OnListeningClick) {Checked = _listening});
            menuItems.Add(new ToolStripSeparator());
            menuItems.Add(new ToolStripMenuItem("Exit", null, OnExitClick));

            ClipboardNotification.ClipboardUpdate += ClipboardNotificationOnClipboardUpdate;
        }

        private ToolStripItem[] GetRateItems()
        {
            var availableRates = new[] {10, 9, 8, 7, 6, 5, 4, 3, 2, 1, 0, -1, -2, -3, -4, -5, -6, -7, -8, -9, -10};
            var rates = new ToolStripItem[availableRates.Length];

            for (var index = 0; index < availableRates.Length; index++)
            {
                var rate = availableRates[index];
                rates[index] = new ToolStripMenuItem(Convert.ToString(rate), null, OnRateItemClick)
                {
                    Checked = _speechSynthesizer.Rate == rate
                };
            }

            return rates;
        }

        private void OnRateItemClick(object sender, EventArgs eventArgs)
        {
            var toolStripMenuItem = (ToolStripMenuItem)sender;
            _speechSynthesizer.Rate = Convert.ToInt32(toolStripMenuItem.Text);

            foreach (ToolStripMenuItem item in toolStripMenuItem.Owner.Items)
            {
                item.Checked = _speechSynthesizer.Rate == Convert.ToInt32(item.Text);
            }
        }

        private ToolStripItem[] GetVoiceItems()
        {
            var installedVoices = _speechSynthesizer.GetInstalledVoices();
            var voices = new ToolStripItem[installedVoices.Count];

            for (var index = 0; index < installedVoices.Count; index++)
            {
                var installedVoice = installedVoices[index];
                voices[index] = new ToolStripMenuItem(installedVoice.VoiceInfo.Name, null, OnVoiceItemClick)
                {
                    Checked = _speechSynthesizer.Voice.Equals(installedVoice.VoiceInfo)
                };
            }

            return voices;
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
