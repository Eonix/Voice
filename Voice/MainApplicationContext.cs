using System;
using System.Collections.Concurrent;
using System.ComponentModel;
using System.Linq;
using System.Speech.Synthesis;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Voice.Properties;

namespace Voice
{
    public class MainApplicationContext : ApplicationContext
    {
        private readonly Container components;
        private readonly NotifyIcon notifyIcon;

        private CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
        private bool listening;

        public MainApplicationContext()
        {
            components = new Container();
            notifyIcon = new NotifyIcon(components)
            {
                ContextMenuStrip = new ContextMenuStrip(),
                Icon = Resources.icon,
                Text = @"Voice",
                Visible = true
            };
            
            listening = Settings.Default.Listening;
            
            PopulateMenu();

            var notificationForm = new NotificationForm();
            components.Add(notificationForm);
            notificationForm.ClipboardUpdate += ClipboardNotificationOnClipboardUpdate;
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
            cancellationTokenSource.Cancel();
            cancellationTokenSource.Dispose();
            cancellationTokenSource = new CancellationTokenSource();
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
                    Checked = Settings.Default.Rate == rate
                };
            }

            return rates;
        }

        private void OnRateItemClick(object sender, EventArgs eventArgs)
        {
            var toolStripMenuItem = (ToolStripMenuItem)sender;

            Settings.Default.Rate = Convert.ToInt32(toolStripMenuItem.Text);
            Settings.Default.Save();

            foreach (ToolStripMenuItem item in toolStripMenuItem.Owner.Items)
            {
                item.Checked = Settings.Default.Rate == Convert.ToInt32(item.Text);
            }
        }

        private static ToolStripItem[] GetVoiceItems()
        {
            using (var speechSynthesizer = new SpeechSynthesizer())
            {
                var currentVoice = string.IsNullOrWhiteSpace(Settings.Default.Voice)
                    ? speechSynthesizer.Voice.Name // If no voice is set, use default from engine.
                    : Settings.Default.Voice;

                var installedVoices =
                    speechSynthesizer.GetInstalledVoices()
                        .Select(voice => new ToolStripMenuItem(voice.VoiceInfo.Name, null, OnVoiceItemClick)
                        {
                            Checked = currentVoice == voice.VoiceInfo.Name
                        });
                
                return installedVoices.Cast<ToolStripItem>().ToArray();
            }
        }

        private static void OnVoiceItemClick(object sender, EventArgs eventArgs)
        {
            var toolStripMenuItem = (ToolStripMenuItem)sender;

            Settings.Default.Voice = toolStripMenuItem.Text;
            Settings.Default.Save();

            foreach (ToolStripMenuItem item in toolStripMenuItem.Owner.Items)
            {
                item.Checked = Settings.Default.Voice == item.Text;
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

        private void ClipboardNotificationOnClipboardUpdate(object sender, IDataObject dataObject)
        {
            if (!listening)
                return;

            cancellationTokenSource.Cancel();
            cancellationTokenSource.Dispose();
            cancellationTokenSource = new CancellationTokenSource();

            Task.Run(async () =>
            {
                using (var synthesizer = new SpeechSynthesizer())
                {
                    synthesizer.Rate = Settings.Default.Rate;

                    if (!string.IsNullOrWhiteSpace(Settings.Default.Voice))
                        synthesizer.SelectVoice(Settings.Default.Voice);

                    await synthesizer.SpeakTextAsync(Convert.ToString(dataObject.GetData(DataFormats.Text)),
                        cancellationTokenSource.Token);
                }
            });
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
                cancellationTokenSource.Dispose();
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
