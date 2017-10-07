using System;

namespace Voice
{
    [Serializable]
    public class VoiceProfile
    {
        public string Name { get; set; } = string.Empty;
        public int Rate { get; set; } = 0;
        public int Volume { get; set; } = 100;
    }
}