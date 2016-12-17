using System;
using System.ComponentModel;
using System.Drawing;
using System.Speech.Synthesis;
using System.Windows.Forms;
using Voice.Properties;

namespace Voice
{
    public class MainApplicationContext : ApplicationContext
    {
        private readonly Container components;
        private readonly NotifyIcon notifyIcon;
        private readonly SpeechSynthesizer speechSynthesizer;

        private bool listening;

        public MainApplicationContext()
        {
            speechSynthesizer = new SpeechSynthesizer();
            components = new Container();
            notifyIcon = new NotifyIcon(components)
            {
                ContextMenuStrip = new ContextMenuStrip(),
                Icon = SystemIcons.WinLogo,
                Text = @"Voice",
                Visible = true
            };

            speechSynthesizer.Rate = Settings.Default.Rate;
            if (!string.IsNullOrWhiteSpace(Settings.Default.Voice))
                speechSynthesizer.SelectVoice(Settings.Default.Voice);

            listening = Settings.Default.Listening;
            PopulateMenu();

            ClipboardNotification.ClipboardUpdate += ClipboardNotificationOnClipboardUpdate;
        }

        private void PopulateMenu()
        {
            var menuItems = notifyIcon.ContextMenuStrip.Items;

            var voicesMenuItem = new ToolStripMenuItem("Voices");
            voicesMenuItem.DropDownItems.AddRange(GetVoiceItems());

            var rateMenuItem = new ToolStripMenuItem("Rate");
            rateMenuItem.DropDownItems.AddRange(GetRateItems());

            menuItems.Add(voicesMenuItem);
            menuItems.Add(rateMenuItem);
            menuItems.Add(new ToolStripMenuItem("Listening", null, OnListeningClick) { Checked = listening });
            menuItems.Add(new ToolStripMenuItem("Stop Talking", null, OnStopTalkingClick));
            menuItems.Add(new ToolStripSeparator());
            menuItems.Add(new ToolStripMenuItem("Exit", null, OnExitClick));
        }

        private void OnStopTalkingClick(object sender, EventArgs eventArgs)
        {
            speechSynthesizer.SpeakAsyncCancelAll();
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
                    Checked = speechSynthesizer.Rate == rate
                };
            }

            return rates;
        }

        private void OnRateItemClick(object sender, EventArgs eventArgs)
        {
            var toolStripMenuItem = (ToolStripMenuItem)sender;
            speechSynthesizer.Rate = Convert.ToInt32(toolStripMenuItem.Text);

            Settings.Default.Rate = speechSynthesizer.Rate;
            Settings.Default.Save();

            foreach (ToolStripMenuItem item in toolStripMenuItem.Owner.Items)
            {
                item.Checked = speechSynthesizer.Rate == Convert.ToInt32(item.Text);
            }
        }

        private ToolStripItem[] GetVoiceItems()
        {
            var installedVoices = speechSynthesizer.GetInstalledVoices();
            var voices = new ToolStripItem[installedVoices.Count];

            for (var index = 0; index < installedVoices.Count; index++)
            {
                var installedVoice = installedVoices[index];
                voices[index] = new ToolStripMenuItem(installedVoice.VoiceInfo.Name, null, OnVoiceItemClick)
                {
                    Checked = speechSynthesizer.Voice.Equals(installedVoice.VoiceInfo)
                };
            }

            return voices;
        }

        private void OnVoiceItemClick(object sender, EventArgs eventArgs)
        {
            var toolStripMenuItem = (ToolStripMenuItem)sender;
            speechSynthesizer.SelectVoice(toolStripMenuItem.Text);

            Settings.Default.Voice = toolStripMenuItem.Text;
            Settings.Default.Save();

            foreach (ToolStripMenuItem item in toolStripMenuItem.Owner.Items)
            {
                item.Checked = speechSynthesizer.Voice.Name == item.Text;
            }
        }

        private void OnListeningClick(object sender, EventArgs eventArgs)
        {
            var toolStripMenuItem = (ToolStripMenuItem) sender;
            listening = !listening;
            toolStripMenuItem.Checked = listening;

            Settings.Default.Listening = listening;
            Settings.Default.Save();
        }

        private void ClipboardNotificationOnClipboardUpdate(object sender, EventArgs eventArgs)
        {
            if (!listening)
                return;

            var text = Clipboard.GetText();
            if (string.IsNullOrWhiteSpace(text))
                return;

            speechSynthesizer.SpeakAsyncCancelAll();
            speechSynthesizer.SpeakAsync(text);
        }

        private void OnExitClick(object sender, EventArgs eventArgs)
        {
            ExitThread();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                components?.Dispose();
                speechSynthesizer?.Dispose();
            }

            base.Dispose(disposing);
        }

        protected override void ExitThreadCore()
        {
            notifyIcon.Visible = false;

            base.ExitThreadCore();
        }
    }
}
