using System.Threading;
using System.Threading.Tasks;

// ReSharper disable once CheckNamespace
namespace System.Speech.Synthesis
{
    public static class SpeechSynthesizerExtensions
    {
        public static Task<Prompt> SpeakTextAsync(this SpeechSynthesizer synthesizer, string textToSpeak, CancellationToken token)
        {
            var source = new TaskCompletionSource<Prompt>();
            synthesizer.SpeakCompleted += (_, eventArgs) =>
            {
                if (eventArgs.Error != null)
                    source.SetException(eventArgs.Error);
                else if (eventArgs.Cancelled)
                    source.SetCanceled();
                else if (eventArgs.Prompt != null)
                    source.SetResult(eventArgs.Prompt);
            };

            var prompt = synthesizer.SpeakAsync(textToSpeak);
            token.Register(() =>
            {
                if (!prompt.IsCompleted)
                    synthesizer.SpeakAsyncCancel(prompt);
            });

            return source.Task;
        }
    }
}
