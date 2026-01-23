using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using Screenshot.Core;
using Screenshot.App.Services;
using Screenshot.Platform.Mac;
#if NET10_0_WINDOWS
using Screenshot.Platform.Windows;
#endif

namespace Screenshot.App.ViewModels
{
    public sealed class MainViewModel : INotifyPropertyChanged
    {
        private OutputMode _selectedOutputMode = OutputMode.AudioAndVideo;
        private AudioCaptureMode _selectedAudioCaptureMode = AudioCaptureMode.NativeSystemAudio;
        private string _outputDirectory;
        private string _statusMessage = "待机";
        private string _lastOutputSummary = "暂无录制";
        private string _macHelperPath;
        private string _logDirectory;
        private bool _isRecording;
        private IRecordingBackend? _backend;
        private DocumentCapturePipeline? _docPipeline;
        private string? _lastPptPath;
        private string? _lastPdfPath;
        private DateTime _recordingStart;
        private string? _sessionBaseName;
        private string _sessionName;
        private string _sessionDirectoryPreview = "";
        private bool _freezeSessionPreview;
        private bool _generatePpt = true;
        private bool _generatePdf = true;
        private bool _keepJpgFiles = true;
        private string _screenshotIntervalText = "10";
        private bool _logEnabled = true;
        private bool _logAppendMode = false;
        private bool _useCustomRegion = false;
        private string _regionLeftText = "0";
        private string _regionTopText = "0";
        private string _regionWidthText = "0";
        private string _regionHeightText = "0";
        private string _recordingDurationMinutesText = "60";
        private readonly string _settingsPath;
        private string _settingsPathDisplay;
        private CancellationTokenSource? _autoStopCts;
        private bool _stopInProgress;

        public event PropertyChangedEventHandler? PropertyChanged;

        public ObservableCollection<OutputMode> OutputModes { get; } =
            new(Enum.GetValues<OutputMode>());

        public ObservableCollection<AudioCaptureMode> AudioCaptureModes { get; } =
            new(Enum.GetValues<AudioCaptureMode>());

        public OutputMode SelectedOutputMode
        {
            get => _selectedOutputMode;
            set => SetField(ref _selectedOutputMode, value);
        }

        public AudioCaptureMode SelectedAudioCaptureMode
        {
            get => _selectedAudioCaptureMode;
            set => SetField(ref _selectedAudioCaptureMode, value);
        }

        public string OutputDirectory
        {
            get => _outputDirectory;
            set => SetField(ref _outputDirectory, value);
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set => SetField(ref _statusMessage, value);
        }

        public string LastOutputSummary
        {
            get => _lastOutputSummary;
            set => SetField(ref _lastOutputSummary, value);
        }

        public string MacHelperPath
        {
            get => _macHelperPath;
            set => SetField(ref _macHelperPath, value);
        }

        public string LogDirectory
        {
            get => _logDirectory;
            private set => SetField(ref _logDirectory, value);
        }

        public string SessionName
        {
            get => _sessionName;
            set => SetField(ref _sessionName, value);
        }

        public bool IsEditingLocked => _isRecording;
        public bool IsEditingUnlocked => !_isRecording;

        public string SessionDirectoryPreview
        {
            get => _sessionDirectoryPreview;
            private set => SetField(ref _sessionDirectoryPreview, value);
        }

        private string _sessionDirectoryStatus = "";

        public string SessionDirectoryStatus
        {
            get => _sessionDirectoryStatus;
            private set => SetField(ref _sessionDirectoryStatus, value);
        }

        public bool HasSessionDirectory => !string.IsNullOrWhiteSpace(SessionDirectoryStatus);

        public bool GeneratePpt
        {
            get => _generatePpt;
            set => SetField(ref _generatePpt, value);
        }

        public bool GeneratePdf
        {
            get => _generatePdf;
            set => SetField(ref _generatePdf, value);
        }

        public bool KeepJpgFiles
        {
            get => _keepJpgFiles;
            set => SetField(ref _keepJpgFiles, value);
        }

        public string ScreenshotInterval
        {
            get => _screenshotIntervalText;
            set => SetField(ref _screenshotIntervalText, value);
        }

        public string RecordingDurationMinutes
        {
            get => _recordingDurationMinutesText;
            set => SetField(ref _recordingDurationMinutesText, value);
        }

        public bool UseCustomRegion
        {
            get => _useCustomRegion;
            set => SetField(ref _useCustomRegion, value);
        }

        public string RegionLeft
        {
            get => _regionLeftText;
            set => SetField(ref _regionLeftText, value);
        }

        public string RegionTop
        {
            get => _regionTopText;
            set => SetField(ref _regionTopText, value);
        }

        public string RegionWidth
        {
            get => _regionWidthText;
            set => SetField(ref _regionWidthText, value);
        }

        public string RegionHeight
        {
            get => _regionHeightText;
            set => SetField(ref _regionHeightText, value);
        }

        public bool LogEnabled
        {
            get => _logEnabled;
            set => SetField(ref _logEnabled, value);
        }

        public bool LogAppendMode
        {
            get => _logAppendMode;
            set => SetField(ref _logAppendMode, value);
        }

        public string SettingsPathDisplay
        {
            get => _settingsPathDisplay;
            private set => SetField(ref _settingsPathDisplay, value);
        }

        public ICommand StartCommand { get; }
        public ICommand StopCommand { get; }
        public ICommand CopySessionPathCommand { get; }
        public ICommand OpenSessionDirectoryCommand { get; }
        public ICommand OpenOutputDirectoryCommand { get; }
        public ICommand OpenLogDirectoryCommand { get; }
        public ICommand ResetSettingsCommand { get; }
        public ICommand OpenSettingsDirectoryCommand { get; }

        public MainViewModel()
        {
            _outputDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "ScreenshotV4.0");

            _logDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "ScreenshotV4.0-Logs");

            _settingsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "ScreenshotV4.0",
                "settings.json");
            _settingsPathDisplay = _settingsPath;

            Logger.SetLogDirectory(_logDirectory);
            Logger.Enabled = _logEnabled;
            Logger.SetLogFileMode(_logAppendMode ? 1 : 0);

            _macHelperPath = Environment.GetEnvironmentVariable("SCREENSHOT_MAC_HELPER")
                ?? Path.Combine(AppContext.BaseDirectory, "RecorderHelper");

            if (OperatingSystem.IsMacOS() && !string.IsNullOrWhiteSpace(_macHelperPath))
            {
                var baseDir = AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                if (_macHelperPath.Contains("@executable_path", StringComparison.Ordinal))
                {
                    _macHelperPath = _macHelperPath.Replace("@executable_path", baseDir, StringComparison.Ordinal);
                }
                _macHelperPath = Path.GetFullPath(_macHelperPath, AppContext.BaseDirectory);
            }

            _sessionName = $"Screenshot_{DateTime.Now:yyyyMMdd}";
            LoadSettings();
            RefreshSessionDirectoryPreview();

            StartCommand = new AsyncDelegateCommand(StartRecordingAsync, () => !_isRecording);
            StopCommand = new AsyncDelegateCommand(StopRecordingAsync, () => _isRecording);
            CopySessionPathCommand = new DelegateCommand(CopySessionPath, () => !string.IsNullOrWhiteSpace(SessionDirectoryStatus));
            OpenSessionDirectoryCommand = new DelegateCommand(OpenSessionDirectory, () => !string.IsNullOrWhiteSpace(SessionDirectoryStatus));
            OpenOutputDirectoryCommand = new DelegateCommand(OpenOutputDirectory, () => Directory.Exists(OutputDirectory));
            OpenLogDirectoryCommand = new DelegateCommand(OpenLogDirectory, () => Directory.Exists(LogDirectory));
            ResetSettingsCommand = new DelegateCommand(ResetSettings, () => !_isRecording);
            OpenSettingsDirectoryCommand = new DelegateCommand(OpenSettingsDirectory, () => Directory.Exists(Path.GetDirectoryName(_settingsPath) ?? string.Empty));
        }

        private async System.Threading.Tasks.Task StartRecordingAsync()
        {
            Directory.CreateDirectory(OutputDirectory);
            Directory.CreateDirectory(LogDirectory);
            Logger.SetLogDirectory(LogDirectory);
            Logger.Enabled = LogEnabled;
            Logger.SetLogFileMode(LogAppendMode ? 1 : 0);
            _isRecording = true;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsEditingLocked)));
            StatusMessage = "录制中";
            _freezeSessionPreview = true;
            RaiseCommandStates();

            var config = new RecordingConfig
            {
                AudioCaptureMode = SelectedAudioCaptureMode,
                GeneratePPT = GeneratePpt,
                GeneratePDF = GeneratePdf,
                KeepJpgFiles = KeepJpgFiles,
                ScreenshotInterval = ParseIntervalSeconds(),
                UseCustomRegion = UseCustomRegion,
                RegionLeft = ParseInt(RegionLeft),
                RegionTop = ParseInt(RegionTop),
                RegionWidth = ParseInt(RegionWidth),
                RegionHeight = ParseInt(RegionHeight)
            };

            var sanitizedName = SanitizeSessionName(SessionName);
            _sessionBaseName = $"{DateTime.Now:yyyyMMdd_HHmmss}_{sanitizedName}";
            var sessionDirName = _sessionBaseName;
            var sessionDir = Path.Combine(OutputDirectory, sessionDirName);
            SessionDirectoryStatus = sessionDir;
            var options = new RecordingSessionOptions(
                sessionDir,
                _sessionBaseName,
                SelectedOutputMode,
                config);

            try
            {
                _docPipeline = new DocumentCapturePipeline(config, sessionDir, _sessionBaseName);
                await _docPipeline.StartAsync(System.Threading.CancellationToken.None);

                _backend = CreateBackend();
                _recordingStart = DateTime.UtcNow;
                await _backend.StartAsync(options, System.Threading.CancellationToken.None);
                StartAutoStopTimer();
            }
            catch (Exception ex)
            {
                _isRecording = false;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsEditingLocked)));
                _freezeSessionPreview = false;
                StatusMessage = $"启动失败: {ex.Message}";
                Logger.WriteError("Start recording failed", ex);
                await DisposeBackendAsync();
                await DisposeDocPipelineAsync();
                _autoStopCts?.Cancel();
            }
            finally
            {
                RaiseCommandStates();
            }
        }

        private async System.Threading.Tasks.Task StopRecordingAsync()
        {
            if (_stopInProgress) return;
            _stopInProgress = true;
            _autoStopCts?.Cancel();
            _isRecording = false;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsEditingLocked)));
            StatusMessage = "停止中...";
            RaiseCommandStates();

            if (_backend != null)
            {
                var result = await _backend.StopAsync(System.Threading.CancellationToken.None);
                await DisposeBackendAsync();
                var duration = DateTime.UtcNow - _recordingStart;
                StatusMessage = $"已停止（{duration:hh\\:mm\\:ss}）";
                var ppt = _lastPptPath ?? "-";
                var pdf = _lastPdfPath ?? "-";
                LastOutputSummary = $"mp4: {result.VideoPath ?? "-"} | wav: {result.AudioPath ?? "-"} | ppt: {ppt} | pdf: {pdf}";
            }
            else
            {
                StatusMessage = "已停止（无后端）";
            }
            _freezeSessionPreview = false;
            if (_docPipeline != null)
            {
                _lastPptPath = _docPipeline.PptPath;
                _lastPdfPath = _docPipeline.PdfPath;
                await _docPipeline.StopAsync(System.Threading.CancellationToken.None);
                await DisposeDocPipelineAsync();
            }
            SessionDirectoryStatus = "";
            RaiseCommandStates();
            _stopInProgress = false;
        }

        private void RaiseCommandStates()
        {
            if (StartCommand is DelegateCommand start) start.RaiseCanExecuteChanged();
            if (StopCommand is DelegateCommand stop) stop.RaiseCanExecuteChanged();
            if (StartCommand is AsyncDelegateCommand asyncStart) asyncStart.RaiseCanExecuteChanged();
            if (StopCommand is AsyncDelegateCommand asyncStop) asyncStop.RaiseCanExecuteChanged();
            if (CopySessionPathCommand is DelegateCommand copy) copy.RaiseCanExecuteChanged();
            if (OpenSessionDirectoryCommand is DelegateCommand open) open.RaiseCanExecuteChanged();
            if (OpenOutputDirectoryCommand is DelegateCommand openOutput) openOutput.RaiseCanExecuteChanged();
            if (OpenLogDirectoryCommand is DelegateCommand openLog) openLog.RaiseCanExecuteChanged();
            if (ResetSettingsCommand is DelegateCommand reset) reset.RaiseCanExecuteChanged();
            if (OpenSettingsDirectoryCommand is DelegateCommand openSettings) openSettings.RaiseCanExecuteChanged();
        }

        private async System.Threading.Tasks.Task DisposeBackendAsync()
        {
            if (_backend == null) return;
            await _backend.DisposeAsync();
            _backend = null;
        }

        private async System.Threading.Tasks.Task DisposeDocPipelineAsync()
        {
            if (_docPipeline == null) return;
            await _docPipeline.DisposeAsync();
            _docPipeline = null;
        }

        private int ParseIntervalSeconds()
        {
            if (int.TryParse(ScreenshotInterval, out var seconds))
            {
                return Math.Max(1, seconds);
            }

            return 10;
        }

        private static int ParseInt(string value)
        {
            if (int.TryParse(value, out var result))
            {
                return result;
            }
            return 0;
        }

        private void StartAutoStopTimer()
        {
            _autoStopCts?.Cancel();
            _autoStopCts = new CancellationTokenSource();
            var minutes = ParseInt(RecordingDurationMinutes);
            if (minutes <= 0) return;

            var token = _autoStopCts.Token;
            _ = System.Threading.Tasks.Task.Run(async () =>
            {
                try
                {
                    await System.Threading.Tasks.Task.Delay(TimeSpan.FromMinutes(minutes), token);
                    if (!token.IsCancellationRequested && _isRecording)
                    {
                        await StopRecordingAsync();
                    }
                }
                catch
                {
                    // ignore
                }
            }, token);
        }

        private static string SanitizeSessionName(string? name)
        {
            var value = string.IsNullOrWhiteSpace(name) ? "Session" : name.Trim();
            foreach (var ch in Path.GetInvalidFileNameChars())
            {
                value = value.Replace(ch, '_');
            }
            return string.IsNullOrWhiteSpace(value) ? "Session" : value;
        }

        private void RefreshSessionDirectoryPreview()
        {
            if (_freezeSessionPreview) return;
            var sanitizedName = SanitizeSessionName(SessionName);
            var previewName = $"{DateTime.Now:yyyyMMdd_HHmmss}_{sanitizedName}";
            SessionDirectoryPreview = Path.Combine(OutputDirectory, previewName);
        }

        private void SetField<T>(ref T field, T value, [CallerMemberName] string? name = null)
        {
            if (Equals(field, value)) return;
            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

            if (name == nameof(SessionName) || name == nameof(OutputDirectory))
            {
                RefreshSessionDirectoryPreview();
            }
            if (name == nameof(_isRecording) || name == nameof(IsEditingLocked))
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsEditingLocked)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsEditingUnlocked)));
            }
            if (name == nameof(SessionDirectoryStatus))
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasSessionDirectory)));
            }
            if (name == nameof(OutputDirectory) || name == nameof(LogDirectory))
            {
                if (OpenOutputDirectoryCommand is DelegateCommand openOutput) openOutput.RaiseCanExecuteChanged();
                if (OpenLogDirectoryCommand is DelegateCommand openLog) openLog.RaiseCanExecuteChanged();
            }
            if (name == nameof(SettingsPathDisplay))
            {
                if (OpenSettingsDirectoryCommand is DelegateCommand openSettings) openSettings.RaiseCanExecuteChanged();
            }
            if (name == nameof(LogEnabled))
            {
                Logger.Enabled = LogEnabled;
            }
            if (name == nameof(LogAppendMode))
            {
                Logger.SetLogFileMode(LogAppendMode ? 1 : 0);
            }
            if (ShouldPersistSetting(name))
            {
                SaveSettings();
            }

        }

        private void ResetSettings()
        {
            if (_isRecording) return;
            var defaults = new AppSettings();
            _outputDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "ScreenshotV4.0");
            _logDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "ScreenshotV4.0-Logs");
            _sessionName = $"Screenshot_{DateTime.Now:yyyyMMdd}";

            _generatePpt = defaults.GeneratePpt;
            _generatePdf = defaults.GeneratePdf;
            _keepJpgFiles = defaults.KeepJpgFiles;
            _screenshotIntervalText = defaults.ScreenshotInterval;
            _logEnabled = defaults.LogEnabled;
            _logAppendMode = defaults.LogAppendMode;
            _selectedOutputMode = defaults.SelectedOutputMode;
            _selectedAudioCaptureMode = defaults.SelectedAudioCaptureMode;
            _useCustomRegion = defaults.UseCustomRegion;
            _regionLeftText = defaults.RegionLeft;
            _regionTopText = defaults.RegionTop;
            _regionWidthText = defaults.RegionWidth;
            _regionHeightText = defaults.RegionHeight;
            _recordingDurationMinutesText = defaults.RecordingDurationMinutes;

            Logger.SetLogDirectory(_logDirectory);
            Logger.Enabled = _logEnabled;
            Logger.SetLogFileMode(_logAppendMode ? 1 : 0);

            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(OutputDirectory)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(LogDirectory)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SessionName)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(GeneratePpt)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(GeneratePdf)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(KeepJpgFiles)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ScreenshotInterval)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(LogEnabled)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(LogAppendMode)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedOutputMode)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedAudioCaptureMode)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(UseCustomRegion)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(RegionLeft)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(RegionTop)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(RegionWidth)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(RegionHeight)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(RecordingDurationMinutes)));

            RefreshSessionDirectoryPreview();
            SaveSettings();
        }

        public void ExportSettingsTo(string path)
        {
            if (_isRecording) return;
            try
            {
                var settings = new AppSettings
                {
                    OutputDirectory = OutputDirectory,
                    SessionName = SessionName,
                    GeneratePpt = GeneratePpt,
                    GeneratePdf = GeneratePdf,
                KeepJpgFiles = KeepJpgFiles,
                ScreenshotInterval = ScreenshotInterval,
                LogEnabled = LogEnabled,
                LogAppendMode = LogAppendMode,
                SelectedOutputMode = SelectedOutputMode,
                SelectedAudioCaptureMode = SelectedAudioCaptureMode,
                RecordingDurationMinutes = RecordingDurationMinutes
            };
                settings.Save(path);
                StatusMessage = $"已导出设置: {path}";
                SettingsPathDisplay = _settingsPath;
            }
            catch
            {
                StatusMessage = "导出设置失败";
            }
        }

        public void ImportSettingsFrom(string path)
        {
            if (_isRecording) return;
            try
            {
                var settings = AppSettings.Load(path);
                ApplySettings(settings);
                StatusMessage = $"已导入设置: {path}";
                SettingsPathDisplay = _settingsPath;
            }
            catch
            {
                StatusMessage = "导入设置失败";
            }
        }

        private void ApplySettings(AppSettings settings)
        {
            if (!string.IsNullOrWhiteSpace(settings.OutputDirectory))
            {
                _outputDirectory = settings.OutputDirectory;
            }
            if (!string.IsNullOrWhiteSpace(settings.SessionName))
            {
                _sessionName = settings.SessionName;
            }

            _generatePpt = settings.GeneratePpt;
            _generatePdf = settings.GeneratePdf;
            _keepJpgFiles = settings.KeepJpgFiles;
            _screenshotIntervalText = settings.ScreenshotInterval;
            _logEnabled = settings.LogEnabled;
            _logAppendMode = settings.LogAppendMode;
            _selectedOutputMode = settings.SelectedOutputMode;
            _selectedAudioCaptureMode = settings.SelectedAudioCaptureMode;
            _useCustomRegion = settings.UseCustomRegion;
            _regionLeftText = settings.RegionLeft;
            _regionTopText = settings.RegionTop;
            _regionWidthText = settings.RegionWidth;
            _regionHeightText = settings.RegionHeight;
            _recordingDurationMinutesText = settings.RecordingDurationMinutes;
            _useCustomRegion = settings.UseCustomRegion;
            _regionLeftText = settings.RegionLeft;
            _regionTopText = settings.RegionTop;
            _regionWidthText = settings.RegionWidth;
            _regionHeightText = settings.RegionHeight;
            _recordingDurationMinutesText = settings.RecordingDurationMinutes;

            Logger.Enabled = _logEnabled;
            Logger.SetLogFileMode(_logAppendMode ? 1 : 0);

            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(OutputDirectory)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SessionName)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(GeneratePpt)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(GeneratePdf)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(KeepJpgFiles)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ScreenshotInterval)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(LogEnabled)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(LogAppendMode)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedOutputMode)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedAudioCaptureMode)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(UseCustomRegion)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(RegionLeft)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(RegionTop)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(RegionWidth)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(RegionHeight)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(RecordingDurationMinutes)));

            RefreshSessionDirectoryPreview();
            SaveSettings();
        }

        private void LoadSettings()
        {
            var settings = AppSettings.Load(_settingsPath);
            if (!string.IsNullOrWhiteSpace(settings.OutputDirectory))
            {
                _outputDirectory = settings.OutputDirectory;
            }
            if (!string.IsNullOrWhiteSpace(settings.SessionName))
            {
                _sessionName = settings.SessionName;
            }

            _generatePpt = settings.GeneratePpt;
            _generatePdf = settings.GeneratePdf;
            _keepJpgFiles = settings.KeepJpgFiles;
            _screenshotIntervalText = settings.ScreenshotInterval;
            _logEnabled = settings.LogEnabled;
            _logAppendMode = settings.LogAppendMode;
            _selectedOutputMode = settings.SelectedOutputMode;
            _selectedAudioCaptureMode = settings.SelectedAudioCaptureMode;

            Logger.SetLogDirectory(_logDirectory);
            Logger.Enabled = _logEnabled;
            Logger.SetLogFileMode(_logAppendMode ? 1 : 0);
        }

        private void SaveSettings()
        {
            var settings = new AppSettings
            {
                OutputDirectory = OutputDirectory,
                SessionName = SessionName,
                GeneratePpt = GeneratePpt,
                GeneratePdf = GeneratePdf,
                KeepJpgFiles = KeepJpgFiles,
                ScreenshotInterval = ScreenshotInterval,
                LogEnabled = LogEnabled,
                LogAppendMode = LogAppendMode,
                SelectedOutputMode = SelectedOutputMode,
                SelectedAudioCaptureMode = SelectedAudioCaptureMode,
                UseCustomRegion = UseCustomRegion,
                RegionLeft = RegionLeft,
                RegionTop = RegionTop,
                RegionWidth = RegionWidth,
                RegionHeight = RegionHeight,
                RecordingDurationMinutes = RecordingDurationMinutes
            };
            settings.Save(_settingsPath);
        }

        private static bool ShouldPersistSetting(string? name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            return name == nameof(OutputDirectory)
                || name == nameof(SessionName)
                || name == nameof(GeneratePpt)
                || name == nameof(GeneratePdf)
                || name == nameof(KeepJpgFiles)
                || name == nameof(ScreenshotInterval)
                || name == nameof(LogEnabled)
                || name == nameof(LogAppendMode)
                || name == nameof(SelectedOutputMode)
                || name == nameof(SelectedAudioCaptureMode)
                || name == nameof(UseCustomRegion)
                || name == nameof(RegionLeft)
                || name == nameof(RegionTop)
                || name == nameof(RegionWidth)
                || name == nameof(RegionHeight)
                || name == nameof(RecordingDurationMinutes);
        }

        private IRecordingBackend CreateBackend()
        {
            if (OperatingSystem.IsMacOS())
            {
                return new MacRecordingBackend(MacHelperPath);
            }
#if NET10_0_WINDOWS
            if (OperatingSystem.IsWindows())
            {
                return new WindowsRecordingBackend();
            }
#endif

            throw new PlatformNotSupportedException("Windows backend not wired yet.");
        }

        private void CopySessionPath()
        {
            if (string.IsNullOrWhiteSpace(SessionDirectoryStatus)) return;
            if (!OperatingSystem.IsMacOS()) return;
            try
            {
                var startInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "pbcopy",
                    RedirectStandardInput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using var process = System.Diagnostics.Process.Start(startInfo);
                if (process?.StandardInput != null)
                {
                    process.StandardInput.Write(SessionDirectoryStatus);
                    process.StandardInput.Close();
                }
            }
            catch
            {
                // Clipboard integration is platform-specific; ignore failures for now.
            }
        }

        private void OpenSessionDirectory()
        {
            if (string.IsNullOrWhiteSpace(SessionDirectoryStatus)) return;
            if (!OperatingSystem.IsMacOS()) return;
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "open",
                    Arguments = $"\"{SessionDirectoryStatus}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                });
            }
            catch
            {
                // ignore
            }
        }

        private void OpenOutputDirectory()
        {
            if (!OperatingSystem.IsMacOS()) return;
            if (!Directory.Exists(OutputDirectory)) return;
            TryOpenDirectory(OutputDirectory);
        }

        private void OpenLogDirectory()
        {
            if (!OperatingSystem.IsMacOS()) return;
            if (!Directory.Exists(LogDirectory)) return;
            TryOpenDirectory(LogDirectory);
        }

        private void OpenSettingsDirectory()
        {
            if (!OperatingSystem.IsMacOS()) return;
            var dir = Path.GetDirectoryName(_settingsPath);
            if (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir)) return;
            TryOpenDirectory(dir);
        }

        private static void TryOpenDirectory(string path)
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "open",
                    Arguments = $"\"{path}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                });
            }
            catch
            {
                // ignore
            }
        }
    }
}
