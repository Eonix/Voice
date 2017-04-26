using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Speech.Synthesis;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Voice.Properties;

namespace Voice
{
    public class MainApplicationContext : ApplicationContext
    {
        private readonly Stopwatch lastSpeechStopwatch = new Stopwatch();
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
                Icon = Resources.icon,
                Text = @"Voice",
                Visible = true
            };

            Settings.Default.Reload();
            
            listening = Settings.Default.Listening;
            speechSynthesizer.Rate = Settings.Default.Rate;
            speechSynthesizer.Volume = Settings.Default.Volume;

            if (!string.IsNullOrWhiteSpace(Settings.Default.Voice))
                speechSynthesizer.SelectVoice(Settings.Default.Voice);

            PopulateMenu();

            var notificationForm = new NotificationForm();
            components.Add(notificationForm);
            notificationForm.ClipboardUpdate += ClipboardNotificationOnClipboardUpdate;

            // Timer for restarting the application if the speech engine was last used 5 mins. ago.
            // This is sadly a necessary fix to prevent the memory leak by the SAPI engine.
            var restartTimer = new Timer(components);
            restartTimer.Tick += RestartTimerOnTick;
            restartTimer.Interval = (int) TimeSpan.FromSeconds(1).TotalMilliseconds;
            restartTimer.Start();
        }

        private void RestartTimerOnTick(object sender, EventArgs eventArgs)
        {
            if (speechSynthesizer.State != SynthesizerState.Ready)
                return;

            if (lastSpeechStopwatch.Elapsed < TimeSpan.FromMinutes(5))
                return;

            // Restart the application.
            Process.Start(Application.ExecutablePath);
            ExitThread();
        }

        private void PopulateMenu()
        {
            var menuItems = notifyIcon.ContextMenuStrip.Items;

            var voicesMenuItem = new ToolStripMenuItem("Voices");
            voicesMenuItem.DropDownItems.AddRange(GetVoiceItems());

            var rateMenuItem = new ToolStripMenuItem("Rate");
            rateMenuItem.DropDownItems.AddRange(GetRateItems());

            var volumeMenuItem = new ToolStripMenuItem("Volume");
            volumeMenuItem.DropDownItems.AddRange(GetVolumeItems());

            menuItems.Add(voicesMenuItem);
            menuItems.Add(rateMenuItem);
            menuItems.Add(volumeMenuItem);
            menuItems.Add(new ToolStripMenuItem("Listening", null, OnListeningClick) { Checked = listening });
            menuItems.Add(new ToolStripMenuItem("Stop Talking", null, (_, __) => speechSynthesizer.SpeakAsyncCancelAll()));
            menuItems.Add(new ToolStripSeparator());
            menuItems.Add(new ToolStripMenuItem("Exit", null, (_, __) => ExitThread()));
        }

        private ToolStripItem[] GetVolumeItems()
        {
            var availableVolumes = new[] { 100, 90, 80, 70, 60, 50, 40, 30, 20, 10, 0 };

            return
                availableVolumes.Select(volume => new ToolStripMenuItem(Convert.ToString(volume), null, OnVolumeItemClick)
                {
                    Checked = Settings.Default.Volume == volume
                }).Cast<ToolStripItem>().ToArray();
        }

        private void OnVolumeItemClick(object sender, EventArgs e)
        {
            var toolStripMenuItem = (ToolStripMenuItem)sender;

            Settings.Default.Volume = Convert.ToInt32(toolStripMenuItem.Text);
            Settings.Default.Save();

            speechSynthesizer.Volume = Settings.Default.Volume;

            foreach (ToolStripMenuItem item in toolStripMenuItem.Owner.Items)
            {
                item.Checked = Settings.Default.Volume == Convert.ToInt32(item.Text);
            }
        }

        private ToolStripItem[] GetRateItems()
        {
            var availableRates = new[] {10, 9, 8, 7, 6, 5, 4, 3, 2, 1, 0, -1, -2, -3, -4, -5, -6, -7, -8, -9, -10};
            return availableRates.Select(rate => new ToolStripMenuItem(Convert.ToString(rate), null, OnRateItemClick)
            {
                Checked = Settings.Default.Rate == rate
            }).Cast<ToolStripItem>().ToArray();
        }

        private void OnRateItemClick(object sender, EventArgs eventArgs)
        {
            var toolStripMenuItem = (ToolStripMenuItem)sender;

            Settings.Default.Rate = Convert.ToInt32(toolStripMenuItem.Text);
            Settings.Default.Save();

            speechSynthesizer.Rate = Settings.Default.Rate;

            foreach (ToolStripMenuItem item in toolStripMenuItem.Owner.Items)
            {
                item.Checked = Settings.Default.Rate == Convert.ToInt32(item.Text);
            }
        }

        private ToolStripItem[] GetVoiceItems()
        {
            var currentVoice = string.IsNullOrWhiteSpace(Settings.Default.Voice)
                ? speechSynthesizer.Voice.Name // If no voice is set, use default from engine.
                : Settings.Default.Voice;

            return speechSynthesizer.GetInstalledVoices()
                .Select(voice => new ToolStripMenuItem(voice.VoiceInfo.Name, null, OnVoiceItemClick)
                {
                    Checked = currentVoice == voice.VoiceInfo.Name
                }).Cast<ToolStripItem>().ToArray();
        }

        private void OnVoiceItemClick(object sender, EventArgs eventArgs)
        {
            var toolStripMenuItem = (ToolStripMenuItem)sender;

            Settings.Default.Voice = toolStripMenuItem.Text;
            Settings.Default.Save();

            speechSynthesizer.SelectVoice(Settings.Default.Voice);

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

            lastSpeechStopwatch.Restart();

            speechSynthesizer.SpeakAsyncCancelAll();
            speechSynthesizer.SpeakAsync(ShortenUrls(Convert.ToString(dataObject.GetData(DataFormats.Text))));
        }

        private static string ShortenUrls(string content)
        {
            return Regex.Replace(content,
                @"((http|ftp|https):\/\/[\w\-_]+(\.[\w\-_]+)+([\w\-\.,@?^=%&amp;:/~\+#]*[\w\-\@?^=%&amp;/~\+#])?)",
                ReplaceUrl);
        }

        private static string ReplaceUrl(Match match)
        {
            var matchValue = match.ToString();
            return Uri.TryCreate(matchValue, UriKind.RelativeOrAbsolute, out var result) ? result.Host : matchValue;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                components.Dispose();
                speechSynthesizer.Dispose();
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
