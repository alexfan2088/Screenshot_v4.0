using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using ImageSharpImage = SixLabors.ImageSharp.Image;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Screenshot.Core;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Screenshot.App.Services
{
    internal sealed class DocumentCapturePipeline : IAsyncDisposable
    {
        public readonly record struct CaptureStatus(int Count, double ChangeRate, DateTime NextIntervalAtUtc);

        public event Action<CaptureStatus>? CaptureCompleted;

        private readonly RecordingConfig _config;
        private readonly string _outputDir;
        private readonly string _baseName;
        private readonly SemaphoreSlim _captureLock = new(1, 1);
        private CancellationTokenSource? _cts;
        private Task? _loopTask;
        private int _captureIndex;
        private PPTGenerator? _ppt;
        private bool _initialized;
        private Image<Rgba32>? _lastIntervalImage;
        private DateTime _nextIntervalAtUtc;

        public string? PptPath { get; private set; }

        public DocumentCapturePipeline(RecordingConfig config, string outputDir, string baseName)
        {
            _config = config;
            _outputDir = outputDir;
            _baseName = baseName;
        }

        public async Task StartAsync(CancellationToken cancellationToken)
        {
            if (_loopTask != null)
            {
                throw new InvalidOperationException("Document capture already started");
            }

            Directory.CreateDirectory(_outputDir);
            _cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

            _nextIntervalAtUtc = DateTime.UtcNow.AddSeconds(Math.Max(1, _config.ScreenshotInterval));

            _loopTask = Task.Run(async () =>
            {
                try
                {
                    while (true)
                    {
                        await CaptureOnceAsync(_cts.Token);
                        await Task.Delay(TimeSpan.FromSeconds(1), _cts.Token);
                    }
                }
                catch (OperationCanceledException)
                {
                    // Expected during stop.
                }
            }, _cts.Token);
        }

        public async Task StopAsync(CancellationToken cancellationToken)
        {
            if (_cts == null) return;

            _cts.Cancel();
            if (_loopTask != null)
            {
                await _loopTask;
            }

            await _captureLock.WaitAsync(cancellationToken);
            try
            {
                if (_ppt != null)
                {
                    _ppt.Finish();
                    _ppt.Dispose();
                    _ppt = null;
                }

                _lastIntervalImage?.Dispose();
                _lastIntervalImage = null;
            }
            finally
            {
                _captureLock.Release();
            }
        }

        private async Task CaptureOnceAsync(CancellationToken cancellationToken)
        {
            if (!await _captureLock.WaitAsync(0, cancellationToken))
            {
                return;
            }

            try
            {
                var nowUtc = DateTime.UtcNow;
                var tempPath = Path.Combine(Path.GetTempPath(), $"{_baseName}_{Guid.NewGuid():N}.jpg");
                await CaptureScreenAsync(tempPath, cancellationToken);
                if (!File.Exists(tempPath))
                {
                    Logger.WriteWarning($"Screenshot not captured: {tempPath}");
                    CaptureCompleted?.Invoke(new CaptureStatus(_captureIndex, 0, _nextIntervalAtUtc));
                    return;
                }

                using var currentImage = ImageSharpImage.Load<Rgba32>(tempPath);
                var changeRate = CalculateChangeRate(currentImage);
                var intervalReached = nowUtc >= _nextIntervalAtUtc;
                var shouldCapture = intervalReached && (_captureIndex == 0 || changeRate >= _config.ScreenChangeRate);

                if (intervalReached)
                {
                    UpdateIntervalBaseline(currentImage);
                    _nextIntervalAtUtc = nowUtc.AddSeconds(Math.Max(1, _config.ScreenshotInterval));
                }

                if (shouldCapture)
                {
                    var fileName = $"{_baseName}_{_captureIndex:D5}.jpg";
                    var imagePath = Path.Combine(_outputDir, fileName);
                    _captureIndex++;

                    if (File.Exists(imagePath))
                    {
                        File.Delete(imagePath);
                    }

                    File.Move(tempPath, imagePath);

                    try
                    {
                        if (!_initialized)
                        {
                            InitializeDocuments(imagePath);
                        }

                        if (_ppt != null)
                        {
                            _ppt.AddImage(imagePath);
                        }

                }
                catch (Exception ex)
                {
                    Logger.WriteError("Append image to documents failed", ex);
                }

                    if (!_config.KeepJpgFiles)
                    {
                        TryDelete(imagePath);
                    }
                }
                else
                {
                    TryDelete(tempPath);
                }

                CaptureCompleted?.Invoke(new CaptureStatus(_captureIndex, changeRate, _nextIntervalAtUtc));
            }
            catch (Exception ex)
            {
                Logger.WriteError("Capture pipeline error", ex);
            }
            finally
            {
                _captureLock.Release();
            }
        }

        private double CalculateChangeRate(Image<Rgba32> current)
        {
            try
            {
                if (_lastIntervalImage == null)
                {
                    _lastIntervalImage = current.Clone();
                    return 100.0;
                }

                if (current.Width != _lastIntervalImage.Width || current.Height != _lastIntervalImage.Height)
                {
                    _lastIntervalImage.Dispose();
                    _lastIntervalImage = current.Clone();
                    return 100.0;
                }

                int changedPixels = 0;
                int sampledPixels = 0;
                int sampleRate = 5;
                int width = current.Width;
                int height = current.Height;

                for (int y = 0; y < height; y += sampleRate)
                {
                    for (int x = 0; x < width; x += sampleRate)
                    {
                        var currentColor = current[x, y];
                        var lastColor = _lastIntervalImage[x, y];
                        int diffR = Math.Abs(currentColor.R - lastColor.R);
                        int diffG = Math.Abs(currentColor.G - lastColor.G);
                        int diffB = Math.Abs(currentColor.B - lastColor.B);
                        if (diffR > 20 || diffG > 20 || diffB > 20)
                        {
                            changedPixels++;
                        }
                        sampledPixels++;
                    }
                }

                var rate = sampledPixels == 0 ? 0.0 : (changedPixels / (double)sampledPixels) * 100.0;
                return rate;
            }
            catch (Exception ex)
            {
                Logger.WriteError("Compute change rate failed", ex);
                return 0.0;
            }
        }

        private void UpdateIntervalBaseline(Image<Rgba32> current)
        {
            _lastIntervalImage?.Dispose();
            _lastIntervalImage = current.Clone();
        }

        private async Task<bool> CaptureScreenAsync(string imagePath, CancellationToken cancellationToken)
        {
            if (OperatingSystem.IsMacOS())
            {
                var windowArg = "";
                if (_config.CaptureMode == CaptureMode.Window && _config.WindowId > 0)
                {
                    windowArg = $"-l {_config.WindowId} ";
                }
                var regionArgs = "";
                if (string.IsNullOrWhiteSpace(windowArg) && _config.UseCustomRegion && _config.RegionWidth > 0 && _config.RegionHeight > 0)
                {
                    regionArgs = $"-R {_config.RegionLeft},{_config.RegionTop},{_config.RegionWidth},{_config.RegionHeight} ";
                }

                // Try most specific first, then fall back to more permissive captures.
                var displayArg = _config.DisplayId > 0 ? $"-D {_config.DisplayId} " : "";
                if (await TryScreenCapture($"{displayArg}{windowArg}{regionArgs}-x -t jpg \"{imagePath}\"", cancellationToken) && File.Exists(imagePath))
                {
                    return true;
                }

                if (!string.IsNullOrWhiteSpace(windowArg))
                {
                    return await TryScreenCapture($"{windowArg}-x -t jpg \"{imagePath}\"", cancellationToken) && File.Exists(imagePath);
                }

                if (!string.IsNullOrWhiteSpace(regionArgs))
                {
                    if (await TryScreenCapture($"{regionArgs}-x -t jpg \"{imagePath}\"", cancellationToken) && File.Exists(imagePath))
                    {
                        return true;
                    }
                }

                return await TryScreenCapture($"-x -t jpg \"{imagePath}\"", cancellationToken) && File.Exists(imagePath);
            }

            if (OperatingSystem.IsWindows())
            {
                CaptureWindowsScreenshot(imagePath, _config);
                return true;
            }

            throw new PlatformNotSupportedException("Screenshot capture not implemented for this OS yet.");
        }

        private static async Task<bool> TryScreenCapture(string arguments, CancellationToken cancellationToken)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "screencapture",
                Arguments = arguments,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardError = true,
                RedirectStandardOutput = true
            };

            using var process = Process.Start(startInfo);
            if (process == null)
            {
                Logger.WriteWarning("screencapture failed to start.");
                return false;
            }

            var stdOutTask = process.StandardOutput.ReadToEndAsync();
            var stdErrTask = process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync(cancellationToken);
            var stdOut = await stdOutTask;
            var stdErr = await stdErrTask;

            if (process.ExitCode != 0)
            {
                Logger.WriteWarning($"screencapture exit code {process.ExitCode}. args: {arguments}. stdout: {stdOut}. stderr: {stdErr}");
                return false;
            }

            if (!string.IsNullOrWhiteSpace(stdErr))
            {
                Logger.WriteWarning($"screencapture stderr: {stdErr}");
            }

            return true;
        }

        [SupportedOSPlatform("windows")]
        private static void CaptureWindowsScreenshot(string imagePath, RecordingConfig config)
        {
            var screenLeft = GetSystemMetrics(SM_XVIRTUALSCREEN);
            var screenTop = GetSystemMetrics(SM_YVIRTUALSCREEN);
            var screenWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
            var screenHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN);

            var left = screenLeft;
            var top = screenTop;
            var width = screenWidth;
            var height = screenHeight;

            if (config.UseCustomRegion && config.RegionWidth > 0 && config.RegionHeight > 0)
            {
                left = config.RegionLeft;
                top = config.RegionTop;
                width = config.RegionWidth;
                height = config.RegionHeight;

                if (left < screenLeft) left = screenLeft;
                if (top < screenTop) top = screenTop;
                if (left + width > screenLeft + screenWidth)
                {
                    width = Math.Max(0, screenLeft + screenWidth - left);
                }
                if (top + height > screenTop + screenHeight)
                {
                    height = Math.Max(0, screenTop + screenHeight - top);
                }
            }

            if (width <= 0 || height <= 0)
            {
                return;
            }

            using var bitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            using var graphics = Graphics.FromImage(bitmap);
            var hdcDest = graphics.GetHdc();
            var hdcSrc = GetDC(IntPtr.Zero);

            try
            {
                BitBlt(hdcDest, 0, 0, width, height, hdcSrc, left, top, SRCCOPY | CAPTUREBLT);
                bitmap.Save(imagePath, ImageFormat.Jpeg);
            }
            finally
            {
                graphics.ReleaseHdc(hdcDest);
                ReleaseDC(IntPtr.Zero, hdcSrc);
            }
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

        private void InitializeDocuments(string imagePath)
        {
            using var image = ImageSharpImage.Load(imagePath);
            var width = image.Width;
            var height = image.Height;

            if (_config.GeneratePPT)
            {
                PptPath = Path.Combine(_outputDir, $"{_baseName}.pptx");
                _ppt = new PPTGenerator(PptPath, width, height);
                _ppt.Initialize();
            }

            _initialized = true;
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
                // ignore
            }
        }

        public async ValueTask DisposeAsync()
        {
            if (_cts != null)
            {
                await StopAsync(CancellationToken.None);
                _cts.Dispose();
                _cts = null;
            }

            _lastIntervalImage?.Dispose();
            _lastIntervalImage = null;
            _captureLock.Dispose();
        }

        public void UpdateCaptureSettings(int intervalSeconds, double screenChangeRate)
        {
            _config.ScreenshotInterval = Math.Max(1, intervalSeconds);
            _config.ScreenChangeRate = Math.Max(0, screenChangeRate);
            _nextIntervalAtUtc = DateTime.UtcNow.AddSeconds(_config.ScreenshotInterval);
        }
    }
}
