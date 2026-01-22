using System;
using System.Diagnostics;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using NAudio.Wave;
using Screenshot.Core;

namespace Screenshot.Platform.Windows
{
    public sealed class WindowsRecordingBackend : IRecordingBackend
    {
        private readonly object _writerLock = new();
        private CancellationTokenSource? _cts;
        private Task? _videoTask;
        private WmfSinkWriter? _writer;
        private WasapiLoopbackCapture? _audioCapture;
        private WaveFileWriter? _audioWriter;
        private RecordingSessionOptions? _options;
        private long _audioTime;
        private long _videoFrameIndex;
        private int _videoWidth;
        private int _videoHeight;
        private int _captureLeft;
        private int _captureTop;
        private int _captureWidth;
        private int _captureHeight;
        private int _fps;
        private int _audioSampleRate;
        private int _audioChannels;
        private int _audioBitsPerSample;
        private bool _audioIsFloat;
        private DateTime _startTime;

        public bool IsRecording => _cts != null && !_cts.IsCancellationRequested;

        public Task StartAsync(RecordingSessionOptions options, CancellationToken cancellationToken)
        {
            if (_cts != null) throw new InvalidOperationException("Recording already started");

            _options = options;
            Directory.CreateDirectory(options.OutputDirectory);
            var includeVideo = options.OutputMode != OutputMode.AudioOnly;
            var includeAudio = options.OutputMode != OutputMode.VideoOnly;
            var videoPath = includeVideo || includeAudio ? Path.Combine(options.OutputDirectory, $"{options.BaseFileName}.mp4") : null;
            var audioPath = includeAudio ? Path.Combine(options.OutputDirectory, $"{options.BaseFileName}.wav") : null;
            if (!includeVideo && !includeAudio)
            {
                throw new InvalidOperationException("No output mode selected.");
            }

            var config = options.Config;
            _fps = Math.Max(1, config.VideoFrameRate);
            ComputeCaptureBounds(config);
            _videoWidth = _captureWidth;
            _videoHeight = _captureHeight;

            if (includeAudio)
            {
                _audioCapture = new WasapiLoopbackCapture();
                _audioSampleRate = _audioCapture.WaveFormat.SampleRate;
                _audioChannels = _audioCapture.WaveFormat.Channels;
                _audioBitsPerSample = _audioCapture.WaveFormat.BitsPerSample;
                _audioIsFloat = _audioCapture.WaveFormat.Encoding == WaveFormatEncoding.IeeeFloat;

                // Force PCM 16-bit for sink writer input
                _audioBitsPerSample = 16;
                if (!string.IsNullOrWhiteSpace(audioPath))
                {
                    _audioWriter = new WaveFileWriter(audioPath, new WaveFormat(_audioSampleRate, _audioBitsPerSample, _audioChannels));
                }
            }

            var writer = new WmfSinkWriter();
            var videoBitrate = config.GetVideoBitrateMbps(_videoWidth, _videoHeight) * 1_000_000;
            writer.Initialize(videoPath ?? string.Empty, _videoWidth, _videoHeight, _fps, videoBitrate, _audioSampleRate, _audioChannels, _audioBitsPerSample, includeVideo, includeAudio);
            _writer = writer;

            if (includeAudio && _audioCapture != null)
            {
                _audioCapture.DataAvailable += OnAudioDataAvailable;
                _audioCapture.StartRecording();
            }

            _cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            _startTime = DateTime.UtcNow;
            _videoTask = includeVideo ? Task.Run(() => VideoLoop(_cts.Token)) : Task.CompletedTask;

            return Task.CompletedTask;
        }

        public async Task<RecordingSessionResult> StopAsync(CancellationToken cancellationToken)
        {
            if (_cts == null || _options == null)
            {
                throw new InvalidOperationException("Recording not started");
            }

            _cts.Cancel();
            if (_videoTask != null)
            {
                await _videoTask;
            }

            if (_audioCapture != null)
            {
                _audioCapture.DataAvailable -= OnAudioDataAvailable;
                _audioCapture.StopRecording();
                _audioCapture.Dispose();
                _audioCapture = null;
            }
            _audioWriter?.Dispose();
            _audioWriter = null;

            lock (_writerLock)
            {
                _writer?.Dispose();
                _writer = null;
            }

            var result = new RecordingSessionResult(
                videoPath,
                audioPath,
                null,
                null,
                DateTime.UtcNow - _startTime);

            _cts.Dispose();
            _cts = null;
            return result;
        }

        private void VideoLoop(CancellationToken token)
        {
            var frameDuration = 10_000_000L / _fps;
            using var bitmap = new Bitmap(_videoWidth, _videoHeight, PixelFormat.Format24bppRgb);
            using var graphics = Graphics.FromImage(bitmap);

            while (!token.IsCancellationRequested)
            {
                var hdcDest = graphics.GetHdc();
                var hdcSrc = GetDC(IntPtr.Zero);
                try
                {
                    BitBlt(hdcDest, 0, 0, _videoWidth, _videoHeight, hdcSrc, _captureLeft, _captureTop, SRCCOPY | CAPTUREBLT);
                }
                finally
                {
                    graphics.ReleaseHdc(hdcDest);
                    ReleaseDC(IntPtr.Zero, hdcSrc);
                }

                var rect = new Rectangle(0, 0, _videoWidth, _videoHeight);
                var bmpData = bitmap.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
                var bytes = bmpData.Stride * bmpData.Height;
                var buffer = new byte[bytes];
                Marshal.Copy(bmpData.Scan0, buffer, 0, bytes);
                bitmap.UnlockBits(bmpData);

                var sampleTime = _videoFrameIndex * frameDuration;
                _videoFrameIndex++;

                lock (_writerLock)
                {
                    _writer?.WriteVideoSample(buffer, sampleTime, frameDuration);
                }

                Thread.Sleep(1000 / _fps);
            }
        }

        private void OnAudioDataAvailable(object? sender, WaveInEventArgs e)
        {
            if (_writer == null) return;

            byte[] pcmBuffer;
            int bytes;
            if (_audioIsFloat)
            {
                // Convert float32 to int16
                var samples = e.BytesRecorded / 4;
                pcmBuffer = new byte[samples * 2];
                for (int i = 0; i < samples; i++)
                {
                    var sample = BitConverter.ToSingle(e.Buffer, i * 4);
                    var val = (short)Math.Clamp(sample * short.MaxValue, short.MinValue, short.MaxValue);
                    pcmBuffer[i * 2] = (byte)(val & 0xFF);
                    pcmBuffer[i * 2 + 1] = (byte)((val >> 8) & 0xFF);
                }
                bytes = pcmBuffer.Length;
            }
            else
            {
                pcmBuffer = new byte[e.BytesRecorded];
                Buffer.BlockCopy(e.Buffer, 0, pcmBuffer, 0, e.BytesRecorded);
                bytes = e.BytesRecorded;
            }

            var blockAlign = _audioChannels * (_audioBitsPerSample / 8);
            var samples = bytes / blockAlign;
            var duration = samples * 10_000_000L / _audioSampleRate;
            var sampleTime = _audioTime;
            _audioTime += duration;

            lock (_writerLock)
            {
                _writer?.WriteAudioSample(pcmBuffer, bytes, sampleTime, duration);
            }

            _audioWriter?.Write(pcmBuffer, 0, bytes);
        }

        private void ComputeCaptureBounds(RecordingConfig config)
        {
            var screenLeft = GetSystemMetrics(SM_XVIRTUALSCREEN);
            var screenTop = GetSystemMetrics(SM_YVIRTUALSCREEN);
            var screenWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
            var screenHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN);

            if (config.UseCustomRegion && config.RegionWidth > 0 && config.RegionHeight > 0)
            {
                _captureLeft = Math.Max(config.RegionLeft, screenLeft);
                _captureTop = Math.Max(config.RegionTop, screenTop);
                _captureWidth = Math.Min(config.RegionWidth, screenWidth - (_captureLeft - screenLeft));
                _captureHeight = Math.Min(config.RegionHeight, screenHeight - (_captureTop - screenTop));
            }
            else
            {
                _captureLeft = screenLeft;
                _captureTop = screenTop;
                _captureWidth = screenWidth;
                _captureHeight = screenHeight;
            }

            if (_captureWidth <= 0) _captureWidth = screenWidth;
            if (_captureHeight <= 0) _captureHeight = screenHeight;
        }

        public ValueTask DisposeAsync()
        {
            if (_cts != null)
            {
                _cts.Cancel();
                _cts.Dispose();
                _cts = null;
            }

            return ValueTask.CompletedTask;
        }

        private const int SRCCOPY = 0x00CC0020;
        private const int CAPTUREBLT = 0x40000000;
        private const int SM_XVIRTUALSCREEN = 76;
        private const int SM_YVIRTUALSCREEN = 77;
        private const int SM_CXVIRTUALSCREEN = 78;
        private const int SM_CYVIRTUALSCREEN = 79;

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("gdi32.dll")]
        private static extern bool BitBlt(
            IntPtr hdcDest,
            int nXDest,
            int nYDest,
            int nWidth,
            int nHeight,
            IntPtr hdcSrc,
            int nXSrc,
            int nYSrc,
            int dwRop);
    }
}
