using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Voice.Properties;

namespace Voice
{
    public class MainApplicationContext : ApplicationContext
    {
        private readonly Stopwatch lastSpeechTimeout = new Stopwatch();
        private readonly Container components;
        private readonly NotifyIcon notifyIcon;
        private readonly Speaker speaker;

        public MainApplicationContext()
        {
            speaker = new Speaker((speaking) => { notifyIcon.Icon = speaking ? Resources.icon_inverted : Resources.icon; });
            components = new Container();
            notifyIcon = new NotifyIcon(components)
            {
                ContextMenuStrip = new ContextMenuStrip(),
                Icon = Resources.icon,
                Text = Resources.Voice,
                Visible = true
            };

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
            if (speaker.Speaking)
                return;

            if (lastSpeechTimeout.Elapsed < TimeSpan.FromMinutes(5))
                return;
            
            Application.Restart();
        }

        private void PopulateMenu()
        {
            var menuItems = notifyIcon.ContextMenuStrip.Items;

            var voicesMenuItem = new ToolStripMenuItem(Resources.Voices) {Name = "Voices"};
            voicesMenuItem.DropDownItems.AddRange(GetVoiceItems());

            var rateMenuItem = new ToolStripMenuItem(Resources.Rate) {Name = "Rate"};
            rateMenuItem.DropDownItems.AddRange(GetRateItems());

            var volumeMenuItem = new ToolStripMenuItem(Resources.Volume) {Name = "Volume"};
            volumeMenuItem.DropDownItems.AddRange(GetVolumeItems());

            menuItems.Add(voicesMenuItem);
            menuItems.Add(rateMenuItem);
            menuItems.Add(volumeMenuItem);
            menuItems.Add(new ToolStripMenuItem(Resources.Listening, null, OnListeningClick) { Checked = speaker.Listening });
            menuItems.Add(new ToolStripMenuItem(Resources.StopTalking, null, (_, __) => speaker.StopTalking()));
            menuItems.Add(new ToolStripSeparator());
            menuItems.Add(new ToolStripMenuItem(Resources.Exit, null, (_, __) => ExitThread()));
        }

        private ToolStripItem[] GetVolumeItems()
        {
            var availableVolumes = new[] { 100, 90, 80, 70, 60, 50, 40, 30, 20, 10, 0 };

            return
                availableVolumes.Select(volume => new ToolStripMenuItem(Convert.ToString(volume), null, OnVolumeItemClick)
                {
                    Checked = speaker.Volume == volume
                }).Cast<ToolStripItem>().ToArray();
        }

        private void OnVolumeItemClick(object sender, EventArgs e)
        {
            var toolStripMenuItem = (ToolStripMenuItem)sender;

            speaker.Volume = Convert.ToInt32(toolStripMenuItem.Text);
            ClearAndSetSelectedItem(toolStripMenuItem.Owner, speaker.Volume);
        }

        private ToolStripItem[] GetRateItems()
        {
            var availableRates = new[] {10, 9, 8, 7, 6, 5, 4, 3, 2, 1, 0, -1, -2, -3, -4, -5, -6, -7, -8, -9, -10};
            return availableRates.Select(rate => new ToolStripMenuItem(Convert.ToString(rate), null, OnRateItemClick)
            {
                Checked = speaker.Rate == rate
            }).Cast<ToolStripItem>().ToArray();
        }

        private void OnRateItemClick(object sender, EventArgs eventArgs)
        {
            var toolStripMenuItem = (ToolStripMenuItem)sender;

            speaker.Rate = Convert.ToInt32(toolStripMenuItem.Text);
            ClearAndSetSelectedItem(toolStripMenuItem.Owner, speaker.Rate);
        }

        private ToolStripItem[] GetVoiceItems()
        {
            var currentVoice = speaker.CurrentVoice;
            return speaker.Voices
                .Select(voice => new ToolStripMenuItem(voice, null, OnVoiceItemClick) {Checked = currentVoice == voice})
                .Cast<ToolStripItem>().ToArray();
        }

        private void OnVoiceItemClick(object sender, EventArgs eventArgs)
        {
            var toolStripMenuItem = (ToolStripMenuItem)sender;

            speaker.CurrentVoice = toolStripMenuItem.Text;

            ClearAndSetSelectedItem(toolStripMenuItem.Owner, speaker.CurrentVoice);

            var rateToolStripItem = (ToolStripMenuItem) notifyIcon.ContextMenuStrip.Items["Rate"];
            ClearAndSetSelectedItem(rateToolStripItem.DropDownItems[0].Owner, speaker.Rate);

            var volumeToolStripItem = (ToolStripMenuItem)notifyIcon.ContextMenuStrip.Items["Volume"];
            ClearAndSetSelectedItem(volumeToolStripItem.DropDownItems[0].Owner, speaker.Volume);
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
            speaker.Listening = !speaker.Listening;
            ((ToolStripMenuItem)sender).Checked = speaker.Listening;
        }

        private void ClipboardNotificationOnClipboardUpdate(string text)
        {
            if (!speaker.Listening)
                return;

            lastSpeechTimeout.Restart();
            speaker.Speak(text);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                components.Dispose();
                speaker.Dispose();
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
