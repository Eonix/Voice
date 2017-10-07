using System;
using System.Collections.Generic;

namespace Voice
{
    [Serializable]
    public class VoiceProfileCollection
    {
        public List<VoiceProfile> Profiles { get; } = new List<VoiceProfile>();
    }
}
