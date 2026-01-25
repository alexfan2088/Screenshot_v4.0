using System;
using System.Threading;
using System.Threading.Tasks;

namespace Screenshot.Core
{
    public enum OutputMode
    {
        None,
        AudioOnly,
        VideoOnly,
        AudioAndVideo
    }

    public sealed record RecordingSessionOptions(
        string OutputDirectory,
        string BaseFileName,
        OutputMode OutputMode,
        RecordingConfig Config);

    public sealed record RecordingSessionResult(
        string? VideoPath,
        string? AudioPath,
        string? PptPath,
        TimeSpan Duration);

    public interface IRecordingBackend : IAsyncDisposable
    {
        bool IsRecording { get; }
        Task StartAsync(RecordingSessionOptions options, CancellationToken cancellationToken);
        Task<RecordingSessionResult> StopAsync(CancellationToken cancellationToken);
    }
}
