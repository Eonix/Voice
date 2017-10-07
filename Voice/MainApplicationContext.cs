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

            listening = Settings.Default.Listening;

            var currentVoice = Settings.Default.CurrentVoice;
            if (!string.IsNullOrWhiteSpace(currentVoice))
                speechSynthesizer.SelectVoice(currentVoice);

            if (Settings.Default.ProfileCollection == null)
                Settings.Default.ProfileCollection = new VoiceProfileCollection();

            PopulateMenu(GetOrAddProfile(speechSynthesizer, Settings.Default));

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

        private static VoiceProfile GetOrAddProfile(SpeechSynthesizer synthesizer, Settings settings)
        {
            var voiceName = synthesizer.Voice.Name;
            var voiceProfile = settings.ProfileCollection.Profiles.FirstOrDefault(x => x.Name == voiceName);

            if (voiceProfile != null)
            {
                synthesizer.Rate = voiceProfile.Rate;
                synthesizer.Volume = voiceProfile.Volume;
                return voiceProfile;
            }

            var newProfile = new VoiceProfile {Name = voiceName, Volume = synthesizer.Volume, Rate = synthesizer.Rate};
            settings.ProfileCollection.Profiles.Add(newProfile);

            return newProfile;
        }

        private void RestartTimerOnTick(object sender, EventArgs eventArgs)
        {
            if (speechSynthesizer.State != SynthesizerState.Ready)
                return;

            if (lastSpeechStopwatch.Elapsed < TimeSpan.FromMinutes(5))
                return;
            
            Application.Restart();
        }

        private void PopulateMenu(VoiceProfile profile)
        {
            var menuItems = notifyIcon.ContextMenuStrip.Items;

            var voicesMenuItem = new ToolStripMenuItem("Voices") {Name = "Voices"};
            voicesMenuItem.DropDownItems.AddRange(GetVoiceItems(profile));

            var rateMenuItem = new ToolStripMenuItem("Rate") {Name = "Rate"};
            rateMenuItem.DropDownItems.AddRange(GetRateItems(profile));

            var volumeMenuItem = new ToolStripMenuItem("Volume") {Name = "Volume"};
            volumeMenuItem.DropDownItems.AddRange(GetVolumeItems(profile));

            menuItems.Add(voicesMenuItem);
            menuItems.Add(rateMenuItem);
            menuItems.Add(volumeMenuItem);
            menuItems.Add(new ToolStripMenuItem("Listening", null, OnListeningClick) { Checked = listening });
            menuItems.Add(new ToolStripMenuItem("Stop Talking", null, (_, __) => speechSynthesizer.SpeakAsyncCancelAll()));
            menuItems.Add(new ToolStripSeparator());
            menuItems.Add(new ToolStripMenuItem("Exit", null, (_, __) => ExitThread()));
        }

        private ToolStripItem[] GetVolumeItems(VoiceProfile profile)
        {
            var availableVolumes = new[] { 100, 90, 80, 70, 60, 50, 40, 30, 20, 10, 0 };

            return
                availableVolumes.Select(volume => new ToolStripMenuItem(Convert.ToString(volume), null, OnVolumeItemClick)
                {
                    Checked = profile.Volume == volume
                }).Cast<ToolStripItem>().ToArray();
        }

        private void OnVolumeItemClick(object sender, EventArgs e)
        {
            var toolStripMenuItem = (ToolStripMenuItem)sender;

            var voiceProfile = GetOrAddProfile(speechSynthesizer, Settings.Default);
            voiceProfile.Volume = Convert.ToInt32(toolStripMenuItem.Text);
            speechSynthesizer.Volume = voiceProfile.Volume;
            
            ClearAndSetSelectedItem(toolStripMenuItem.Owner, voiceProfile.Volume);

            Settings.Default.Save();
        }

        private ToolStripItem[] GetRateItems(VoiceProfile profile)
        {
            var availableRates = new[] {10, 9, 8, 7, 6, 5, 4, 3, 2, 1, 0, -1, -2, -3, -4, -5, -6, -7, -8, -9, -10};
            return availableRates.Select(rate => new ToolStripMenuItem(Convert.ToString(rate), null, OnRateItemClick)
            {
                Checked = profile.Rate == rate
            }).Cast<ToolStripItem>().ToArray();
        }

        private void OnRateItemClick(object sender, EventArgs eventArgs)
        {
            var toolStripMenuItem = (ToolStripMenuItem)sender;

            var voiceProfile = GetOrAddProfile(speechSynthesizer, Settings.Default);
            voiceProfile.Rate = Convert.ToInt32(toolStripMenuItem.Text);
            speechSynthesizer.Rate = voiceProfile.Rate;
            
            ClearAndSetSelectedItem(toolStripMenuItem.Owner, voiceProfile.Rate);

            Settings.Default.Save();
        }

        private ToolStripItem[] GetVoiceItems(VoiceProfile profile)
        {
            return speechSynthesizer.GetInstalledVoices()
                .Select(voice => new ToolStripMenuItem(voice.VoiceInfo.Name, null, OnVoiceItemClick)
                {
                    Checked = profile.Name == voice.VoiceInfo.Name
                }).Cast<ToolStripItem>().ToArray();
        }

        private void OnVoiceItemClick(object sender, EventArgs eventArgs)
        {
            var toolStripMenuItem = (ToolStripMenuItem)sender;

            speechSynthesizer.SelectVoice(toolStripMenuItem.Text);
            var voiceProfile = GetOrAddProfile(speechSynthesizer, Settings.Default);
            Settings.Default.CurrentVoice = voiceProfile.Name;

            ClearAndSetSelectedItem(toolStripMenuItem.Owner, voiceProfile.Name);

            var rateToolStripItem = (ToolStripMenuItem) notifyIcon.ContextMenuStrip.Items["Rate"];
            ClearAndSetSelectedItem(rateToolStripItem.DropDownItems[0].Owner, voiceProfile.Rate);

            var volumeToolStripItem = (ToolStripMenuItem)notifyIcon.ContextMenuStrip.Items["Volume"];
            ClearAndSetSelectedItem(volumeToolStripItem.DropDownItems[0].Owner, voiceProfile.Volume);

            Settings.Default.Save();
        }

        private static void ClearAndSetSelectedItem(ToolStrip toolStrip, int value)
        {
            ClearAndSetSelectedItem(toolStrip, Convert.ToString(value));
        }

        private static void ClearAndSetSelectedItem(ToolStrip toolStrip, string value)
        {
            foreach (ToolStripMenuItem item in toolStrip.Items)
            {
                item.Checked = value == item.Text;
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
                match => Uri.TryCreate(match.ToString(), UriKind.RelativeOrAbsolute, out var result)
                    ? result.Host
                    : match.ToString());
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
