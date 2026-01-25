using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Screenshot.Core;

namespace Screenshot.Platform.Mac
{
    public sealed class MacRecordingBackend : IRecordingBackend
    {
        private readonly string _helperPath;
        private Process? _process;
        private DateTime _startTime;
        private RecordingSessionResult? _lastResult;
        private string? _videoPath;
        private string? _audioPath;
        private string? _lastStdOut;
        private string? _lastStdErr;

        public MacRecordingBackend(string helperPath)
        {
            _helperPath = helperPath;
        }

        public bool IsRecording => _process != null && !_process.HasExited;

        public Task StartAsync(RecordingSessionOptions options, CancellationToken cancellationToken)
        {
            if (IsRecording)
            {
                throw new InvalidOperationException("Recording already started");
            }

            if (!File.Exists(_helperPath))
            {
                throw new FileNotFoundException($"RecorderHelper not found: {_helperPath}");
            }

            Directory.CreateDirectory(options.OutputDirectory);
            _videoPath = Path.Combine(options.OutputDirectory, $"{options.BaseFileName}.mp4");
            _audioPath = Path.Combine(options.OutputDirectory, $"{options.BaseFileName}.wav");

            var args = $"--output \"{_videoPath}\" --wav \"{_audioPath}\" --fps {options.Config.VideoFrameRate} --audio-mode {(options.Config.AudioCaptureMode == AudioCaptureMode.NativeSystemAudio ? "native" : "virtual")}";
            if (options.Config.RegionWidth > 0 && options.Config.RegionHeight > 0)
            {
                args += $" --width {options.Config.RegionWidth} --height {options.Config.RegionHeight} --left {options.Config.RegionLeft} --top {options.Config.RegionTop}";
            }
            if (options.OutputMode == OutputMode.AudioOnly)
            {
                args += " --no-video";
            }
            if (options.OutputMode == OutputMode.VideoOnly)
            {
                args += " --no-audio";
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = _helperPath,
                Arguments = args,
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            _process = new Process { StartInfo = startInfo };
            _process.Start();
            _process.OutputDataReceived += (_, e) =>
            {
                if (string.IsNullOrWhiteSpace(e.Data)) return;
                _lastStdOut = e.Data;
                Logger.WriteInfo($"RecorderHelper: {e.Data}");
            };
            _process.ErrorDataReceived += (_, e) =>
            {
                if (string.IsNullOrWhiteSpace(e.Data)) return;
                _lastStdErr = e.Data;
                Logger.WriteError($"RecorderHelper: {e.Data}");
            };
            _process.BeginOutputReadLine();
            _process.BeginErrorReadLine();
            _startTime = DateTime.UtcNow;

            if (_process.WaitForExit(800))
            {
                var detail = _lastStdErr ?? _lastStdOut ?? "RecorderHelper exited immediately";
                throw new InvalidOperationException($"RecorderHelper start failed (code {_process.ExitCode}). {detail}");
            }

            return Task.CompletedTask;
        }

        public async Task<RecordingSessionResult> StopAsync(CancellationToken cancellationToken)
        {
            if (_process == null)
            {
                throw new InvalidOperationException("Recording not started");
            }

            try
            {
                await _process.StandardInput.WriteLineAsync("stop");
                await _process.StandardInput.FlushAsync();
            }
            catch
            {
                // Ignore stdin failure
            }

            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutCts.CancelAfter(TimeSpan.FromSeconds(10));

            try
            {
                await _process.WaitForExitAsync(timeoutCts.Token);
            }
            catch
            {
                try { _process.Kill(); } catch { }
            }

            Logger.WriteInfo($"RecorderHelper exited with code {_process.ExitCode}");

            var duration = DateTime.UtcNow - _startTime;
            _lastResult = new RecordingSessionResult(
                _videoPath,
                _audioPath,
                null,
                null,
                duration);

            return _lastResult;
        }

        public ValueTask DisposeAsync()
        {
            if (_process != null)
            {
                try { if (!_process.HasExited) _process.Kill(); } catch { }
                _process.Dispose();
                _process = null;
            }

            return ValueTask.CompletedTask;
        }
    }
}
