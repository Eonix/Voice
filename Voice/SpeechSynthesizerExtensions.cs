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
            synthesizer.SpeakCompleted += OnSpeakCompleted(source);

            var prompt = synthesizer.SpeakAsync(textToSpeak);
            token.Register(() =>
            {
                if (!prompt.IsCompleted)
                    synthesizer.SpeakAsyncCancel(prompt);
            });

            return source.Task;
        }

        private static EventHandler<SpeakCompletedEventArgs> OnSpeakCompleted(TaskCompletionSource<Prompt> source)
        {
            return (_, eventArgs) =>
            {
                if (eventArgs.Cancelled)
                    source.SetCanceled();
                else if (eventArgs.Error != null)
                    source.SetException(eventArgs.Error);
                else
                    source.SetResult(eventArgs.Prompt);
            };
        }
    }
}
