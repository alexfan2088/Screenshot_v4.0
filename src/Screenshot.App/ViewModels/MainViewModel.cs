using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using Avalonia.Threading;
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
        private string? _lastVideoPath;
        private string? _lastAudioPath;
        private string? _lastLogPath;
        private string? _lastSessionDirectory;
        private DateTime _recordingStart;
        private string? _sessionBaseName;
        private string _sessionName;
        private string _sessionDirectoryPreview = "";
        private bool _freezeSessionPreview;
        private bool _generatePpt = true;
        private bool _keepJpgFiles = true;
        private string _screenshotIntervalText = "10";
        private string _screenChangeRateText = "11.12";
        private bool _logEnabled = true;
        private bool _logAppendMode = false;
        private bool _useCustomRegion = false;
        private string _regionLeftText = "0";
        private string _regionTopText = "0";
        private string _regionWidthText = "0";
        private string _regionHeightText = "0";
        private bool _hasLastRegion;
        private int _lastRegionLeft;
        private int _lastRegionTop;
        private int _lastRegionWidth;
        private int _lastRegionHeight;
        private string _recordingDurationMinutesText = "60";
        private int _videoMergeMode = 1;
        private double _currentChangeRate;
        private int _captureCount;
        private int _nextCheckSeconds;
        private DateTime _nextCheckDueAtUtc;
        private string _remainingTimeText = "--";
        private DispatcherTimer? _statusTimer;
        private DispatcherTimer? _inputApplyTimer;
        private int _lastDurationMinutesApplied;
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
        public bool IsRecording => _isRecording;

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

        public string ScreenChangeRate
        {
            get => _screenChangeRateText;
            set => SetField(ref _screenChangeRateText, value);
        }

        public string RecordingDurationMinutes
        {
            get => _recordingDurationMinutesText;
            set => SetField(ref _recordingDurationMinutesText, value);
        }

        public int VideoMergeMode
        {
            get => _videoMergeMode;
            set => SetField(ref _videoMergeMode, value);
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

        public bool LogDisabled
        {
            get => !LogEnabled;
            set
            {
                if (value)
                {
                    LogEnabled = false;
                }
            }
        }

        public bool LogOverwriteMode
        {
            get => !LogAppendMode;
            set
            {
                if (value)
                {
                    LogAppendMode = false;
                }
            }
        }

        public bool IsOutputModeNone
        {
            get => SelectedOutputMode == OutputMode.None;
            set
            {
                if (value)
                {
                    SelectedOutputMode = OutputMode.None;
                }
            }
        }

        public bool IsOutputModeAudioOnly
        {
            get => SelectedOutputMode == OutputMode.AudioOnly;
            set
            {
                if (value)
                {
                    SelectedOutputMode = OutputMode.AudioOnly;
                }
            }
        }

        public bool IsOutputModeVideoOnly
        {
            get => SelectedOutputMode == OutputMode.VideoOnly;
            set
            {
                if (value)
                {
                    SelectedOutputMode = OutputMode.VideoOnly;
                }
            }
        }

        public bool IsOutputModeAudioAndVideo
        {
            get => SelectedOutputMode == OutputMode.AudioAndVideo;
            set
            {
                if (value)
                {
                    SelectedOutputMode = OutputMode.AudioAndVideo;
                }
            }
        }

        public bool IsVideoMergeLive
        {
            get => VideoMergeMode == 1;
            set
            {
                if (value)
                {
                    VideoMergeMode = 1;
                }
            }
        }

        public bool IsVideoMergePost
        {
            get => VideoMergeMode == 0;
            set
            {
                if (value)
                {
                    VideoMergeMode = 0;
                }
            }
        }

        public string OutputModeNoneLabel => SelectedOutputMode == OutputMode.None ? "✓ 不生成音视频文件" : "不生成音视频文件";
        public string OutputModeAudioOnlyLabel => SelectedOutputMode == OutputMode.AudioOnly ? "✓ 只生成音频文件" : "只生成音频文件";
        public string OutputModeVideoOnlyLabel => SelectedOutputMode == OutputMode.VideoOnly ? "✓ 只生成视频文件" : "只生成视频文件";
        public string OutputModeAudioAndVideoLabel => SelectedOutputMode == OutputMode.AudioAndVideo ? "✓ 生成音频+视频文件" : "生成音频+视频文件";
        public string VideoMergeLiveLabel => VideoMergeMode == 1 ? "✓ 边录边合（推荐，秒级完成）" : "边录边合（推荐，秒级完成）";
        public string VideoMergePostLabel => VideoMergeMode == 0 ? "✓ 后期合成（长视频需等待）" : "后期合成（长视频需等待）";
        public string KeepJpgLabel => KeepJpgFiles ? "✓ 保留JPG文件" : "保留JPG文件";
        public string GeneratePptLabel => GeneratePpt ? "✓ 生成PPT" : "生成PPT";
        public string LogEnabledLabel => LogEnabled ? "✓ 输出日志" : "输出日志";
        public string LogDisabledLabel => LogEnabled ? "不输出日志" : "✓ 不输出日志";
        public string LogOverwriteLabel => LogAppendMode ? "日志文件覆盖" : "✓ 日志文件覆盖";
        public string LogAppendLabel => LogAppendMode ? "✓ 日志文件追加" : "日志文件追加";

        public string SettingsPathDisplay
        {
            get => _settingsPathDisplay;
            private set => SetField(ref _settingsPathDisplay, value);
        }

        public string CurrentChangeRateText => $"{Clamp(_currentChangeRate, 0, 99.99):0.00}%";
        public string CaptureCountText => $"{Clamp(_captureCount, 0, 9999)}";
        public string RemainingTimeText => _remainingTimeText;
        public string NextCheckSecondsText
            => _nextCheckSeconds <= 0 ? "0秒后" : $"{Clamp(_nextCheckSeconds, 1, 9999)}秒后";

        public string? LastVideoPath
        {
            get => _lastVideoPath;
            private set => SetField(ref _lastVideoPath, value);
        }

        public string? LastAudioPath
        {
            get => _lastAudioPath;
            private set => SetField(ref _lastAudioPath, value);
        }

        public string? LastPptPath
        {
            get => _lastPptPath;
            private set => SetField(ref _lastPptPath, value);
        }


        public string? LastLogPath
        {
            get => _lastLogPath;
            private set => SetField(ref _lastLogPath, value);
        }

        public string? LastSessionDirectory
        {
            get => _lastSessionDirectory;
            private set => SetField(ref _lastSessionDirectory, value);
        }

        public bool HasLastVideo => HasFile(LastVideoPath);
        public bool HasLastAudio => HasFile(LastAudioPath);
        public bool HasLastPpt => HasFile(LastPptPath);
        public bool HasLastLog => HasFile(LastLogPath);
        public bool HasLastSessionDirectory => !string.IsNullOrWhiteSpace(LastSessionDirectory) && Directory.Exists(LastSessionDirectory);

        public ICommand StartCommand { get; }
        public ICommand StopCommand { get; }
        public ICommand CopySessionPathCommand { get; }
        public ICommand OpenSessionDirectoryCommand { get; }
        public ICommand OpenOutputDirectoryCommand { get; }
        public ICommand OpenLogDirectoryCommand { get; }
        public ICommand ResetSettingsCommand { get; }
        public ICommand OpenSettingsDirectoryCommand { get; }
        public ICommand OpenLastVideoCommand { get; }
        public ICommand OpenLastAudioCommand { get; }
        public ICommand OpenLastPptCommand { get; }
        public ICommand OpenLastLogCommand { get; }
        public ICommand OpenLastSessionDirectoryCommand { get; }
        public ICommand SetOutputModeNoneCommand { get; }
        public ICommand SetOutputModeAudioOnlyCommand { get; }
        public ICommand SetOutputModeVideoOnlyCommand { get; }
        public ICommand SetOutputModeAudioAndVideoCommand { get; }
        public ICommand SetVideoMergeLiveCommand { get; }
        public ICommand SetVideoMergePostCommand { get; }
        public ICommand ToggleKeepJpgCommand { get; }
        public ICommand ToggleGeneratePptCommand { get; }
        public ICommand SetLogEnabledCommand { get; }
        public ICommand SetLogDisabledCommand { get; }
        public ICommand SetLogOverwriteCommand { get; }
        public ICommand SetLogAppendCommand { get; }

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
            ResetSettingsCommand = new DelegateCommand(RestoreLastRegion, () => !_isRecording);
            OpenSettingsDirectoryCommand = new DelegateCommand(OpenSettingsDirectory, () => Directory.Exists(Path.GetDirectoryName(_settingsPath) ?? string.Empty));
            OpenLastVideoCommand = new DelegateCommand(OpenLastVideo, () => HasLastVideo);
            OpenLastAudioCommand = new DelegateCommand(OpenLastAudio, () => HasLastAudio);
            OpenLastPptCommand = new DelegateCommand(OpenLastPpt, () => HasLastPpt);
            OpenLastLogCommand = new DelegateCommand(OpenLastLogFile, () => HasLastLog);
            OpenLastSessionDirectoryCommand = new DelegateCommand(OpenLastSessionDirectory, () => HasLastSessionDirectory);
            SetOutputModeNoneCommand = new DelegateCommand(() => SelectedOutputMode = OutputMode.None);
            SetOutputModeAudioOnlyCommand = new DelegateCommand(() => SelectedOutputMode = OutputMode.AudioOnly);
            SetOutputModeVideoOnlyCommand = new DelegateCommand(() => SelectedOutputMode = OutputMode.VideoOnly);
            SetOutputModeAudioAndVideoCommand = new DelegateCommand(() => SelectedOutputMode = OutputMode.AudioAndVideo);
            SetVideoMergeLiveCommand = new DelegateCommand(() => VideoMergeMode = 1);
            SetVideoMergePostCommand = new DelegateCommand(() => VideoMergeMode = 0);
            ToggleKeepJpgCommand = new DelegateCommand(() => KeepJpgFiles = !KeepJpgFiles);
            ToggleGeneratePptCommand = new DelegateCommand(() => GeneratePpt = !GeneratePpt);
            SetLogEnabledCommand = new DelegateCommand(() => LogEnabled = true);
            SetLogDisabledCommand = new DelegateCommand(() => LogEnabled = false);
            SetLogOverwriteCommand = new DelegateCommand(() => LogAppendMode = false);
            SetLogAppendCommand = new DelegateCommand(() => LogAppendMode = true);
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
                KeepJpgFiles = KeepJpgFiles,
                ScreenshotInterval = ParseIntervalSeconds(),
                ScreenChangeRate = ParseScreenChangeRate(),
                VideoMergeMode = VideoMergeMode,
                LogEnabled = LogEnabled ? 1 : 0,
                LogFileMode = LogAppendMode ? 1 : 0,
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
                _recordingStart = DateTime.UtcNow;
                _docPipeline = new DocumentCapturePipeline(config, sessionDir, _sessionBaseName);
                _docPipeline.CaptureCompleted += OnCaptureCompleted;
                InitializeStatusInfo();
                StartStatusTimer();
                await _docPipeline.StartAsync(System.Threading.CancellationToken.None);

                _backend = CreateBackend();
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
                StopStatusTimer();
                if (_docPipeline != null)
                {
                    _docPipeline.CaptureCompleted -= OnCaptureCompleted;
                }
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
            var currentSessionDir = SessionDirectoryStatus;

            RecordingSessionResult? result = null;
            if (_backend != null)
            {
                result = await _backend.StopAsync(System.Threading.CancellationToken.None);
                await DisposeBackendAsync();
                var duration = DateTime.UtcNow - _recordingStart;
                StatusMessage = $"已停止（{duration:hh\\:mm\\:ss}）";
                LastVideoPath = result.VideoPath;
                LastAudioPath = result.AudioPath;
            }
            else
            {
                StatusMessage = "已停止（无后端）";
            }
            _freezeSessionPreview = false;
            if (_docPipeline != null)
            {
                _docPipeline.CaptureCompleted -= OnCaptureCompleted;
                LastPptPath = _docPipeline.PptPath;
                await _docPipeline.StopAsync(System.Threading.CancellationToken.None);
                await DisposeDocPipelineAsync();
            }
            StopStatusTimer();
            _nextCheckSeconds = 0;
            _remainingTimeText = "--";
            UpdateStatusInfo();
            LastLogPath = Logger.LogFilePath;
            LastSessionDirectory = currentSessionDir;
            var ppt = LastPptPath ?? "-";
            var mp4 = LastVideoPath ?? "-";
            var wav = LastAudioPath ?? "-";
            LastOutputSummary = $"mp4: {mp4} | wav: {wav} | ppt: {ppt}";
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
            if (OpenLastVideoCommand is DelegateCommand lastVideo) lastVideo.RaiseCanExecuteChanged();
            if (OpenLastAudioCommand is DelegateCommand lastAudio) lastAudio.RaiseCanExecuteChanged();
            if (OpenLastPptCommand is DelegateCommand lastPpt) lastPpt.RaiseCanExecuteChanged();
            if (OpenLastLogCommand is DelegateCommand lastLog) lastLog.RaiseCanExecuteChanged();
            if (OpenLastSessionDirectoryCommand is DelegateCommand lastSession) lastSession.RaiseCanExecuteChanged();
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

        private void InitializeStatusInfo()
        {
            _currentChangeRate = ParseScreenChangeRate();
            _captureCount = 0;
            _nextCheckSeconds = ParseIntervalSeconds();
            _nextCheckDueAtUtc = DateTime.UtcNow.AddSeconds(_nextCheckSeconds);
            _lastDurationMinutesApplied = ParseInt(RecordingDurationMinutes);
            UpdateRemainingTime();
            UpdateStatusInfo();
        }

        private void StartStatusTimer()
        {
            Dispatcher.UIThread.Post(() =>
            {
                _statusTimer?.Stop();
                _statusTimer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(1)
                };
                _statusTimer.Tick += OnStatusTimerTick;
                _statusTimer.Start();
            });
        }

        private void StopStatusTimer()
        {
            if (_statusTimer == null) return;
            _statusTimer.Stop();
            _statusTimer.Tick -= OnStatusTimerTick;
            _statusTimer = null;
        }

        private void OnStatusTimerTick(object? sender, EventArgs e)
        {
            if (!_isRecording) return;
            var remainingSeconds = (int)Math.Ceiling((_nextCheckDueAtUtc - DateTime.UtcNow).TotalSeconds);
            _nextCheckSeconds = Math.Max(0, remainingSeconds);
            UpdateRemainingTime();
            UpdateStatusInfo();
        }

        private void ScheduleInputApply()
        {
            if (_inputApplyTimer == null)
            {
                _inputApplyTimer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(2)
                };
                _inputApplyTimer.Tick += (_, _) =>
                {
                    _inputApplyTimer?.Stop();
                    ApplyInputChanges();
                };
            }

            _inputApplyTimer.Stop();
            _inputApplyTimer.Start();
        }

        private void ApplyInputChanges()
        {
            _currentChangeRate = ParseScreenChangeRate();
            _nextCheckSeconds = ParseIntervalSeconds();
            _nextCheckDueAtUtc = DateTime.UtcNow.AddSeconds(_nextCheckSeconds);
            var durationMinutes = ParseInt(RecordingDurationMinutes);
            if (_isRecording && durationMinutes != _lastDurationMinutesApplied)
            {
                _recordingStart = DateTime.UtcNow;
            }
            _lastDurationMinutesApplied = durationMinutes;
            if (_isRecording && _docPipeline != null)
            {
                _docPipeline.UpdateCaptureSettings(_nextCheckSeconds, _currentChangeRate);
            }
            UpdateRemainingTime();
            UpdateStatusInfo();
        }

        private void OnCaptureCompleted(DocumentCapturePipeline.CaptureStatus status)
        {
            Dispatcher.UIThread.Post(() =>
            {
                _captureCount = status.Count;
                _nextCheckDueAtUtc = status.NextIntervalAtUtc;
                _nextCheckSeconds = Math.Max(0, (int)Math.Ceiling((_nextCheckDueAtUtc - DateTime.UtcNow).TotalSeconds));
                _currentChangeRate = status.ChangeRate;
                UpdateRemainingTime();
                UpdateStatusInfo();
            });
        }

        private void UpdateRemainingTime()
        {
            var totalMinutes = ParseInt(RecordingDurationMinutes);
            if (totalMinutes <= 0)
            {
                _remainingTimeText = "--";
                return;
            }

            var elapsed = DateTime.UtcNow - _recordingStart;
            var totalSeconds = Math.Max(0, (int)TimeSpan.FromMinutes(totalMinutes).TotalSeconds - (int)elapsed.TotalSeconds);
            var minutes = totalSeconds / 60;
            var seconds = totalSeconds % 60;
            _remainingTimeText = $"{Clamp(minutes, 0, 999)}分{Clamp(seconds, 0, 59)}秒";
        }

        private void UpdateStatusInfo()
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CurrentChangeRateText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CaptureCountText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(RemainingTimeText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(NextCheckSecondsText)));
        }

        private int ParseIntervalSeconds()
        {
            if (int.TryParse(ScreenshotInterval, out var seconds))
            {
                return Clamp(seconds, 1, 9999);
            }

            return 1;
        }

        private double ParseScreenChangeRate()
        {
            if (double.TryParse(ScreenChangeRate, out var rate))
            {
                if (rate < 0) return 0;
                if (rate > 99.99) return 99.99;
                return rate;
            }
            return 0;
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
            if (name == nameof(LogDirectory))
            {
                Logger.SetLogDirectory(LogDirectory);
            }
            if (name == nameof(SettingsPathDisplay))
            {
                if (OpenSettingsDirectoryCommand is DelegateCommand openSettings) openSettings.RaiseCanExecuteChanged();
            }
            if (name == nameof(ScreenChangeRate))
            {
                ScheduleInputApply();
            }
            if (name == nameof(ScreenshotInterval))
            {
                ScheduleInputApply();
            }
            if (name == nameof(RecordingDurationMinutes))
            {
                ScheduleInputApply();
            }
            if (name == nameof(SelectedOutputMode))
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsOutputModeNone)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsOutputModeAudioOnly)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsOutputModeVideoOnly)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsOutputModeAudioAndVideo)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(OutputModeNoneLabel)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(OutputModeAudioOnlyLabel)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(OutputModeVideoOnlyLabel)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(OutputModeAudioAndVideoLabel)));
            }
            if (name == nameof(VideoMergeMode))
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsVideoMergeLive)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsVideoMergePost)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(VideoMergeLiveLabel)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(VideoMergePostLabel)));
            }
            if (name == nameof(GeneratePpt))
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(GeneratePptLabel)));
            }
            if (name == nameof(KeepJpgFiles))
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(KeepJpgLabel)));
            }
            if (name == nameof(LogEnabled))
            {
                Logger.Enabled = LogEnabled;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(LogDisabled)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(LogEnabledLabel)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(LogDisabledLabel)));
            }
            if (name == nameof(LogAppendMode))
            {
                Logger.SetLogFileMode(LogAppendMode ? 1 : 0);
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(LogOverwriteMode)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(LogOverwriteLabel)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(LogAppendLabel)));
            }
            if (name == nameof(LastVideoPath))
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasLastVideo)));
                if (OpenLastVideoCommand is DelegateCommand lastVideo) lastVideo.RaiseCanExecuteChanged();
            }
            if (name == nameof(LastAudioPath))
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasLastAudio)));
                if (OpenLastAudioCommand is DelegateCommand lastAudio) lastAudio.RaiseCanExecuteChanged();
            }
            if (name == nameof(LastPptPath))
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasLastPpt)));
                if (OpenLastPptCommand is DelegateCommand lastPpt) lastPpt.RaiseCanExecuteChanged();
            }
            if (name == nameof(LastLogPath))
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasLastLog)));
                if (OpenLastLogCommand is DelegateCommand lastLog) lastLog.RaiseCanExecuteChanged();
            }
            if (name == nameof(LastSessionDirectory))
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasLastSessionDirectory)));
                if (OpenLastSessionDirectoryCommand is DelegateCommand lastSession) lastSession.RaiseCanExecuteChanged();
            }
            if (ShouldPersistSetting(name))
            {
                SaveSettings();
            }

        }

        public void ApplyCustomRegion(int left, int top, int width, int height, bool remember)
        {
            UseCustomRegion = true;
            RegionLeft = left.ToString();
            RegionTop = top.ToString();
            RegionWidth = width.ToString();
            RegionHeight = height.ToString();
            if (remember)
            {
                _hasLastRegion = true;
                _lastRegionLeft = left;
                _lastRegionTop = top;
                _lastRegionWidth = width;
                _lastRegionHeight = height;
            }
        }

        private void RestoreLastRegion()
        {
            if (_isRecording) return;
            if (!_hasLastRegion) return;
            UseCustomRegion = true;
            RegionLeft = _lastRegionLeft.ToString();
            RegionTop = _lastRegionTop.ToString();
            RegionWidth = _lastRegionWidth.ToString();
            RegionHeight = _lastRegionHeight.ToString();
        }

        public void ExportSettingsTo(string path)
        {
            if (_isRecording) return;
            try
            {
                var settings = new AppSettings
                {
                    OutputDirectory = OutputDirectory,
                    LogDirectory = LogDirectory,
                    SessionName = SessionName,
                    GeneratePpt = GeneratePpt,
                    KeepJpgFiles = KeepJpgFiles,
                    ScreenshotInterval = ScreenshotInterval,
                    ScreenChangeRate = ScreenChangeRate,
                    LogEnabled = LogEnabled,
                    LogAppendMode = LogAppendMode,
                    SelectedOutputMode = SelectedOutputMode,
                    SelectedAudioCaptureMode = SelectedAudioCaptureMode,
                    RecordingDurationMinutes = RecordingDurationMinutes,
                    VideoMergeMode = VideoMergeMode
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

        public void UpdateLogDirectory(string path)
        {
            if (_isRecording) return;
            if (string.IsNullOrWhiteSpace(path)) return;
            SetField(ref _logDirectory, path, nameof(LogDirectory));
            StatusMessage = $"已更新日志目录: {path}";
        }

        private void ApplySettings(AppSettings settings)
        {
            if (!string.IsNullOrWhiteSpace(settings.OutputDirectory))
            {
                _outputDirectory = settings.OutputDirectory;
            }
            if (!string.IsNullOrWhiteSpace(settings.LogDirectory))
            {
                _logDirectory = settings.LogDirectory;
            }
            if (!string.IsNullOrWhiteSpace(settings.SessionName))
            {
                _sessionName = settings.SessionName;
            }

            _generatePpt = settings.GeneratePpt;
            _keepJpgFiles = settings.KeepJpgFiles;
            _screenshotIntervalText = settings.ScreenshotInterval;
            _screenChangeRateText = settings.ScreenChangeRate;
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
            _videoMergeMode = settings.VideoMergeMode;

            Logger.Enabled = _logEnabled;
            Logger.SetLogFileMode(_logAppendMode ? 1 : 0);

            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(OutputDirectory)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(LogDirectory)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SessionName)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(GeneratePpt)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(KeepJpgFiles)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ScreenshotInterval)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ScreenChangeRate)));
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
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(VideoMergeMode)));

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
            if (!string.IsNullOrWhiteSpace(settings.LogDirectory))
            {
                _logDirectory = settings.LogDirectory;
            }
            if (!string.IsNullOrWhiteSpace(settings.SessionName))
            {
                _sessionName = settings.SessionName;
            }

            _generatePpt = settings.GeneratePpt;
            _keepJpgFiles = settings.KeepJpgFiles;
            _screenshotIntervalText = settings.ScreenshotInterval;
            _screenChangeRateText = settings.ScreenChangeRate;
            _logEnabled = settings.LogEnabled;
            _logAppendMode = settings.LogAppendMode;
            _selectedOutputMode = settings.SelectedOutputMode;
            _selectedAudioCaptureMode = settings.SelectedAudioCaptureMode;
            _videoMergeMode = settings.VideoMergeMode;

            Logger.SetLogDirectory(_logDirectory);
            Logger.Enabled = _logEnabled;
            Logger.SetLogFileMode(_logAppendMode ? 1 : 0);
        }

        private void SaveSettings()
        {
            var settings = new AppSettings
            {
                OutputDirectory = OutputDirectory,
                LogDirectory = LogDirectory,
                SessionName = SessionName,
                GeneratePpt = GeneratePpt,
                KeepJpgFiles = KeepJpgFiles,
                ScreenshotInterval = ScreenshotInterval,
                ScreenChangeRate = ScreenChangeRate,
                LogEnabled = LogEnabled,
                LogAppendMode = LogAppendMode,
                SelectedOutputMode = SelectedOutputMode,
                SelectedAudioCaptureMode = SelectedAudioCaptureMode,
                UseCustomRegion = UseCustomRegion,
                RegionLeft = RegionLeft,
                RegionTop = RegionTop,
                RegionWidth = RegionWidth,
                RegionHeight = RegionHeight,
                RecordingDurationMinutes = RecordingDurationMinutes,
                VideoMergeMode = VideoMergeMode
            };
            settings.Save(_settingsPath);
        }

        private static bool ShouldPersistSetting(string? name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            return name == nameof(OutputDirectory)
                || name == nameof(LogDirectory)
                || name == nameof(SessionName)
                || name == nameof(GeneratePpt)
                || name == nameof(KeepJpgFiles)
                || name == nameof(ScreenshotInterval)
                || name == nameof(ScreenChangeRate)
                || name == nameof(LogEnabled)
                || name == nameof(LogAppendMode)
                || name == nameof(SelectedOutputMode)
                || name == nameof(SelectedAudioCaptureMode)
                || name == nameof(UseCustomRegion)
                || name == nameof(RegionLeft)
                || name == nameof(RegionTop)
                || name == nameof(RegionWidth)
                || name == nameof(RegionHeight)
                || name == nameof(RecordingDurationMinutes)
                || name == nameof(VideoMergeMode);
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

        private void OpenLastVideo()
        {
            TryOpenFile(LastVideoPath, "MP4");
        }

        private void OpenLastAudio()
        {
            TryOpenFile(LastAudioPath, "WAV");
        }

        private void OpenLastPpt()
        {
            TryOpenFile(LastPptPath, "PPT");
        }


        private void OpenLastLogFile()
        {
            TryOpenFile(LastLogPath, "日志文件");
        }

        private void OpenLastSessionDirectory()
        {
            if (!OperatingSystem.IsMacOS()) return;
            if (!HasLastSessionDirectory) return;
            TryOpenDirectory(LastSessionDirectory!);
        }

        private void TryOpenFile(string? path, string label)
        {
            if (!OperatingSystem.IsMacOS()) return;
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                StatusMessage = $"{label}不存在";
                return;
            }
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
                StatusMessage = $"{label}打开失败";
            }
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

        private static bool HasFile(string? path)
        {
            return !string.IsNullOrWhiteSpace(path) && File.Exists(path);
        }

        private static int Clamp(int value, int min, int max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }

        private static double Clamp(double value, double min, double max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }
    }
}
