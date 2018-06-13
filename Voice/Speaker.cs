using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Speech.Synthesis;
using System.Text.RegularExpressions;
using Voice.Properties;

namespace Voice
{
    public sealed class Speaker : IDisposable
    {
        private readonly SpeechSynthesizer speechSynthesizer;

        public Speaker(Action<bool> speechEventHandler)
        {
            speechSynthesizer = new SpeechSynthesizer();
            speechSynthesizer.StateChanged += (sender, args) => speechEventHandler(args.State == SynthesizerState.Speaking);

            if (string.IsNullOrWhiteSpace(CurrentVoice))
                CurrentVoice = speechSynthesizer.Voice.Name;
            else
                ChangeVoiceProfile(CurrentVoice);
        }

        public string CurrentVoice
        {
            get => Settings.Default.CurrentVoice;
            set
            {
                StopTalking();
                ChangeVoiceProfile(value);
                
                Settings.Default.CurrentVoice = value;
                Settings.Default.Save();
            }
        }

        public bool Listening
        {
            get => Settings.Default.Listening;
            set
            {
                StopTalking();
                Settings.Default.Listening = value;
                Settings.Default.Save();
            }
        }

        public int Volume
        {
            get => GetVoiceProfile(CurrentVoice).Volume;
            set
            {
                SetProfileProperty(x => x.Volume = value);
                speechSynthesizer.Volume = value;
            }
        }

        public int Rate
        {
            get => GetVoiceProfile(CurrentVoice).Rate;
            set
            {
                SetProfileProperty(x => x.Rate = value);
                speechSynthesizer.Rate = value;
            }
        }

        public bool Speaking => speechSynthesizer.State != SynthesizerState.Ready;

        public IEnumerable<string> Voices => speechSynthesizer.GetInstalledVoices().Select(x => x.VoiceInfo.Name);

        public void Speak(string text)
        {
            // Only one prompt should exist at any time, so get the current one and cancel it if it's not finished.
            var currentlySpokenPrompt = speechSynthesizer.GetCurrentlySpokenPrompt();
            if (speechSynthesizer.State != SynthesizerState.Ready && currentlySpokenPrompt != null)
                speechSynthesizer.SpeakAsyncCancel(currentlySpokenPrompt);

            if (!string.IsNullOrWhiteSpace(text))
                speechSynthesizer.SpeakAsync(ShortenUrls(text));
        }
        
        public void StopTalking()
        {
            speechSynthesizer.SpeakAsyncCancelAll();
        }

        private void ChangeVoiceProfile(string voiceName)
        {
            var profile = GetVoiceProfile(voiceName);
            speechSynthesizer.SelectVoice(voiceName);
            speechSynthesizer.Volume = profile.Volume;
            speechSynthesizer.Rate = profile.Rate;
        }

        private void SetProfileProperty(Action<VoiceProfile> action)
        {
            action(GetVoiceProfile(CurrentVoice));
            Settings.Default.Save();
        }
        
        private static string ShortenUrls(string content)
        {
            return Regex.Replace(content,
                @"((http|ftp|https):\/\/[\w\-_]+(\.[\w\-_]+)+([\w\-\.,@?^=%&amp;:/~\+#]*[\w\-\@?^=%&amp;/~\+#])?)",
                match => Uri.TryCreate(match.ToString(), UriKind.RelativeOrAbsolute, out var result)
                    ? result.Host
                    : match.ToString());
        }

        private VoiceProfile GetVoiceProfile(string voiceName)
        {
            if (Settings.Default.ProfileCollection == null)
                Settings.Default.ProfileCollection = new VoiceProfileCollection();

            var profile = Settings.Default.ProfileCollection.Profiles.SingleOrDefault(x => x.Name == voiceName);
            if (profile != null)
                return profile;

            var newProfile = new VoiceProfile { Name = voiceName, Volume = speechSynthesizer.Volume, Rate = speechSynthesizer.Rate };
            Settings.Default.ProfileCollection.Profiles.Add(newProfile);
            Settings.Default.Save();

            return newProfile;
        }
        
        public void Dispose()
        {
            speechSynthesizer.Dispose();
        }
    }
}
