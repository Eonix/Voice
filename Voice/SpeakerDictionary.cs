using System.Collections.Generic;
using System.Deployment.Application;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Voice
{
    public class SpeakerDictionary
    {
        private readonly string rootDirectory;
        private readonly Dictionary<string, string> entries;

        public SpeakerDictionary()
        {
            rootDirectory = ApplicationDeployment.IsNetworkDeployed
                ? ApplicationDeployment.CurrentDeployment.DataDirectory
                : Application.UserAppDataPath;

            entries = new Dictionary<string, string>();
        }

        public void OpenDictionary(string voiceName)
        {
            var path = Path.Combine(rootDirectory, GetFileName(voiceName));
            if (!File.Exists(path))
            {
                using (var stream = File.CreateText(path))
                {
                    stream.WriteLine("# Syntax: word=replacement");
                }
            }

            using (Process.Start(path)) { }
        }

        public void ReloadDictionary(string voiceName)
        {
            entries.Clear();
            var path = Path.Combine(rootDirectory, GetFileName(voiceName));
            if (!File.Exists(path))
                return;

            foreach (var pair in File.ReadAllLines(path).Where(x => !x.StartsWith("#")).Select(x => x.Split('=')))
            {
                entries.Add(pair[0]?.Trim().ToLower() ?? string.Empty, pair[1]?.Trim() ?? string.Empty);
            }
        }

        public string Transform(string text)
        {
            var words = text.Split(' ');
            for (int i = 0; i < words.Length; i++)
            {
                var word = words[i].ToLower();
                if (entries.TryGetValue(word, out var replacement))
                    words[i] = word.Replace(word.TrimEnd('.', '!', '?', ',', ';'), replacement);
            }
            
            return string.Join(" ", words);
        }

        private static string GetFileName(string voiceName) => $"{voiceName.Replace(" ", "-")}-dictionary.txt";
    }
}
