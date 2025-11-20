using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Runtime.InteropServices;
using System.Drawing;
using static Screenshot_v3_0.Logger;

namespace Screenshot_v3_0
{
    public partial class MainWindow : Window
    {
        private readonly AudioRecorder _audioRecorder = new();
        private VideoEncoder? _videoEncoder;
        private RecordingConfig _config;
        private readonly string _workDir;
        private readonly string _configPath;
        private bool _isVideoRecording;
        private RegionHighlightWindow? _regionHighlightWindow;

        public MainWindow()
        {
            InitializeComponent();
            
            // 设置窗口关闭事件
            this.Closing += MainWindow_Closing;
            
            // 获取 .exe 所在目录
            string exeDirectory = AppDomain.CurrentDomain.BaseDirectory;
            
            // 配置文件路径：放在 .exe 所在目录
            _configPath = Path.Combine(exeDirectory, "config.json");

            // 设置工作目录：优先使用D盘，如果D盘不存在则使用C盘（用于保存录制的视频文件）
            string driveLetter = Directory.Exists("D:\\") ? "D:" : "C:";
            _workDir = Path.Combine(driveLetter, "ScreenshotV3.0");
            Directory.CreateDirectory(_workDir);

            // 加载配置
            _config = RecordingConfig.Load(_configPath);
            
            // 初始化日志系统
            Logger.Enabled = _config.LogEnabled == 1;
            Logger.SetLogFileMode(_config.LogFileMode);
            // 设置日志文件目录为工作目录（与视频、音频文件同一目录）
            // SetLogDirectory 会根据模式决定是否清空日志文件
            Logger.SetLogDirectory(_workDir);
            
            // 初始化菜单项状态
            UpdateLogMenuItems();
            
            Logger.WriteLine("程序启动");
            Logger.WriteLine($"日志状态: {(Logger.Enabled ? "启用" : "禁用")}");
            Logger.WriteLine($"日志文件模式: {(_config.LogFileMode == 0 ? "覆盖" : "叠加")}");
            
            UpdateConfigDisplay();
            
            // 延迟初始化区域高亮，直到窗口加载完成
            this.Loaded += (s, e) =>
            {
                UpdateRegionOverlay();
            };
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            // 立即允许关闭，不阻塞
            e.Cancel = false;

            Dispatcher.Invoke(() =>
            {
                _regionHighlightWindow?.Close();
                _regionHighlightWindow = null;
            });

            // 在后台线程执行清理，不阻塞窗口关闭
            ThreadPool.QueueUserWorkItem(_ =>
            {
                try
                {
                    // 如果正在录制，快速停止（不等待完成）
                    if (_isVideoRecording)
                    {
                        try
                        {
                            // 快速停止，不等待完成（FFmpeg 会自动停止音频录制）
                            // 使用快速退出模式完成编码，立即终止 FFmpeg
                            if (_videoEncoder != null)
                            {
                                try
                                {
                                    _videoEncoder.Finish(quickExit: true);
                                }
                                catch (Exception ex)
                                {
                                    WriteError($"完成编码失败", ex);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteError($"停止录制时出错", ex);
                        }
                    }
                    else
                    {
                        // 如果只是音频录制，也停止
                        try
                        {
                            _audioRecorder.Stop();
                        }
                        catch (Exception ex)
                        {
                            WriteError($"停止音频录制时出错", ex);
                        }
                    }

                    // 延迟释放资源，给 Finish 一些时间
                    Thread.Sleep(1000);
                    
                    // 释放资源
                    try
                    {
                        _videoEncoder?.Dispose();
                        _audioRecorder?.Dispose();
                    }
                    catch (Exception ex)
                    {
                        WriteError($"释放资源时出错", ex);
                    }
                }
                catch (Exception ex)
                {
                    WriteError($"窗口关闭处理时出错", ex);
                }
            });
        }

        private void MenuSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var settingsWindow = new SettingsWindow(_config);
                if (settingsWindow.ShowDialog() == true)
                {
                    _config = settingsWindow.Config;
                    _config.Save(_configPath);
                    UpdateConfigDisplay();
                    if (StatusBarInfo != null)
                    {
                        StatusBarInfo.Text = "设置已保存";
                    }
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"打开设置失败：{ex.Message}";
                if (ex.InnerException != null)
                {
                    errorMsg += $"\n详细：{ex.InnerException.Message}";
                }
                if (StatusBarInfo != null)
                {
                    StatusBarInfo.Text = errorMsg;
                }
                WriteError($"打开设置窗口异常", ex);
            }
        }

        private void BtnStartVideo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_isVideoRecording) return;

                var videoPath = Path.Combine(_workDir, $"video{DateTime.Now:yyMMddHHmmss}.mp4");

                int videoWidth;
                int videoHeight;
                int offsetX = 0;
                int offsetY = 0;

                var (screenWidthPixels, screenHeightPixels) = GetPrimaryScreenPixelSize();

                if (HasValidCustomRegion())
                {
                    int regionRight = _config.RegionLeft + _config.RegionWidth;
                    int regionBottom = _config.RegionTop + _config.RegionHeight;
                    
                    if (_config.RegionLeft < 0 || _config.RegionTop < 0 ||
                        regionRight > screenWidthPixels || regionBottom > screenHeightPixels)
                    {
                        string errorMsg = $"选择的录制区域超出主屏幕范围！\n" +
                                         $"区域: ({_config.RegionLeft}, {_config.RegionTop}) 到 ({regionRight}, {regionBottom})\n" +
                                         $"主屏幕: (0, 0) 到 ({screenWidthPixels}, {screenHeightPixels})\n" +
                                         $"请重新选择区域或清除区域设置使用全屏录制。";
                        
                        WriteError(errorMsg);
                        if (StatusBarInfo != null)
                        {
                            StatusBarInfo.Text = "录制区域超出屏幕范围，请重新选择";
                        }
                        
                        MessageBox.Show(errorMsg, "区域超出范围", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    
                    videoWidth = _config.RegionWidth;
                    videoHeight = _config.RegionHeight;
                    offsetX = _config.RegionLeft;
                    offsetY = _config.RegionTop;
                    WriteLine($"使用自定义区域录制: {videoWidth}x{videoHeight} @ 起点 ({offsetX}, {offsetY})");
                }
                else
                {
                    double resolutionScale = _config.VideoResolutionScale / 100.0;
                    videoWidth = (int)(screenWidthPixels * resolutionScale);
                    videoHeight = (int)(screenHeightPixels * resolutionScale);
                }

                // 创建视频编码器（不再需要 VideoRecorder，FFmpeg 直接录制）
                _videoEncoder = new VideoEncoder(videoPath, _config);

                // 初始化编码器（支持自定义区域偏移）
                _videoEncoder.Initialize(videoWidth, videoHeight, _config.VideoFrameRate, 
                    _config.AudioSampleRate, 2, offsetX, offsetY);

                // 开始音频录制（使用 NAudio WasapiLoopbackCapture，无需虚拟声卡）
                // 注意：无论是否有声音输出，都要录制音频（包括静音）
                // AudioRecorder.Start() 会创建 temp_{filename}.wav 文件
                var timestamp = DateTime.Now.ToString("yyMMddHHmmss");
                var audioOutputPath = Path.Combine(_workDir, $"audio_{timestamp}.m4a"); // 最终输出路径（不会被使用）
                var expectedAudioPath = Path.Combine(_workDir, $"temp_audio_{timestamp}.wav"); // AudioRecorder 实际创建的文件
                WriteLine($"========== 开始录制 ==========");
                WriteLine($"视频文件: {videoPath}");
                WriteLine($"开始音频录制（NAudio WasapiLoopbackCapture，即使静音也会录制），输出路径: {audioOutputPath}");
                WriteLine($"预期音频文件: {expectedAudioPath}");
                
                // 连接音频数据事件，实现边录边合成
                _audioRecorder.AudioSampleAvailable += OnAudioSampleAvailable;
                
                _audioRecorder.Start(audioOutputPath);
                
                // 等待音频录制启动
                System.Threading.Thread.Sleep(300);
                
                // 启动 FFmpeg 录制视频
                // VideoEncoder.Start() 会尝试直接录制有声视频，如果失败则使用音频管道实时合成
                WriteLine($"启动 FFmpeg 录制视频...");
                _videoEncoder.Start();

                _isVideoRecording = true;
                BtnStartVideo.IsEnabled = false;
                BtnStopVideo.IsEnabled = true;

                if (StatusBarInfo != null)
                {
                    if (HasValidCustomRegion())
                    {
                        StatusBarInfo.Text = $"开始录制自定义区域：{videoWidth}x{videoHeight} @ ({offsetX},{offsetY})";
                    }
                    else
                    {
                        StatusBarInfo.Text = $"开始录制视频 → {videoPath} | 分辨率: {videoWidth}x{videoHeight} @ {_config.VideoFrameRate}fps";
                    }
                }
            }
            catch (Exception ex)
            {
                if (StatusBarInfo != null)
                {
                    StatusBarInfo.Text = $"开始录制视频失败：{ex.Message}";
                }
            }
        }

        private void BtnStopVideo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!_isVideoRecording) return;

                // 先禁用按钮，防止重复点击
                BtnStopVideo.IsEnabled = false;
                if (StatusBarInfo != null)
                {
                    StatusBarInfo.Text = "正在停止录制...";
                }

                // 在后台线程执行停止操作，避免阻塞 UI
                ThreadPool.QueueUserWorkItem(_ =>
                {
                    try
                    {
                        // 同时停止视频和音频录制，确保时长一致
                        WriteLine($"========== 停止录制 ==========");
                        WriteLine($"同时停止视频和音频录制...");
                        
                        // 先发送停止信号给 FFmpeg（不等待完成）
                        _videoEncoder?.RequestStop();
                        
                        // 断开音频事件连接（避免继续写入数据）
                        _audioRecorder.AudioSampleAvailable -= OnAudioSampleAvailable;
                        
                        // 立即停止音频录制（与视频同时停止）
                        WriteLine($"停止音频录制");
                        _audioRecorder.Stop();
                        
                        // 等待音频管道数据刷新（实时合成模式）或音频文件写入完成（文件合并模式）
                        System.Threading.Thread.Sleep(1000);
                        
                        // 检查是否是实时合成模式（_hasAudioInVideo = true 且没有找到音频设备）
                        // 如果是实时合成模式，不需要查找和合并音频文件
                        // 如果是文件合并模式，需要查找音频文件并合并
                        
                        // 注意：VideoEncoder.Finish() 会根据 _hasAudioInVideo 判断是否需要合并
                        // 如果是实时合成模式，_hasAudioInVideo = true，Finish() 会跳过合并
                        // 如果是文件合并模式，_hasAudioInVideo = false，Finish() 会查找 _tempAudioPath 并合并
                        
                        // 只有在文件合并模式下，才需要查找和设置音频文件路径
                        // 实时合成模式下，音频已经通过管道写入 FFmpeg，无需合并
                        
                        // 尝试查找临时音频文件（仅在文件合并模式下需要）
                        var tempAudioFiles = Directory.GetFiles(_workDir, "temp_*.wav");
                        if (tempAudioFiles.Length == 0)
                        {
                            tempAudioFiles = Directory.GetFiles(_workDir, "temp_audio_*.wav");
                        }
                        
                        if (tempAudioFiles.Length > 0)
                        {
                            var latestAudioFile = tempAudioFiles.OrderByDescending(f => File.GetCreationTime(f)).First();
                            var audioFileInfo = new FileInfo(latestAudioFile);
                            WriteLine($"找到临时音频文件: {latestAudioFile}, 大小: {audioFileInfo.Length} 字节");
                            
                            // 设置音频文件路径（VideoEncoder.Finish() 会根据 _hasAudioInVideo 决定是否使用）
                            if (audioFileInfo.Length > 0)
                            {
                                _videoEncoder?.SetAudioFile(latestAudioFile);
                            }
                        }
                        else
                        {
                            WriteLine($"未找到临时音频文件（可能是实时合成模式，音频已通过管道写入）");
                        }
                        
                        // 现在停止视频录制并完成编码
                        // 注意：如果是实时合成模式，Finish() 会跳过合并；如果是文件合并模式，Finish() 会合并音频
                        WriteLine($"准备停止视频录制（FFmpeg）并完成编码");
                        _videoEncoder?.Finish(); // 这会发送 'q' 给 FFmpeg，等待退出，然后根据模式决定是否合并音频
                        WriteLine($"视频编码完成");

                // 释放资源
                _videoEncoder?.Dispose();

                        // 在 UI 线程更新界面
                        Dispatcher.Invoke(() =>
                        {
                            _videoEncoder = null;
                            _isVideoRecording = false;
                            BtnStartVideo.IsEnabled = true;
                            
                            if (StatusBarInfo != null)
                            {
                                StatusBarInfo.Text = "视频录制已停止（编码完成需几秒）";
                            }
                        });
                    }
                    catch (Exception ex)
                    {
                        WriteError($"停止录制失败", ex);
                        Dispatcher.Invoke(() =>
                        {
                            if (StatusBarInfo != null)
                            {
                                StatusBarInfo.Text = $"停止录制失败：{ex.Message}";
                            }
                            BtnStartVideo.IsEnabled = true;
                            BtnStopVideo.IsEnabled = true;
                        });
                    }
                });
            }
            catch (Exception ex)
            {
                if (StatusBarInfo != null)
                {
                    StatusBarInfo.Text = $"停止录制视频失败：{ex.Message}";
                }
                BtnStartVideo.IsEnabled = true;
                BtnStopVideo.IsEnabled = true;
            }
        }

        private void OnAudioSampleAvailable(byte[]? audioData, int bytesRecorded)
        {
            // 音频样本可用事件：将音频数据实时传递给 FFmpeg（边录边合成）
            if (audioData != null && bytesRecorded > 0 && _videoEncoder != null)
            {
                _videoEncoder.WriteAudioData(audioData, bytesRecorded);
            }
        }


        private void UpdateConfigDisplay()
        {
            try
            {
                if (StatusBarConfig != null)
                {
                    string regionInfo = HasValidCustomRegion()
                        ? $"区域：{_config.RegionWidth}x{_config.RegionHeight} @ ({_config.RegionLeft},{_config.RegionTop})"
                        : "区域：全屏";

                    StatusBarConfig.Text = $"视频：{_config.VideoResolutionScale}%分辨率，{_config.VideoFrameRate}fps，H.264 | " +
                                           $"音频：{_config.AudioSampleRate / 1000}kHz，{_config.AudioBitrate}kbps | {regionInfo}";
                }
            }
            catch (Exception ex)
            {
                WriteError($"更新配置显示失败", ex);
            }
        }

        private bool HasValidCustomRegion()
        {
            return _config.UseCustomRegion &&
                   _config.RegionWidth > 0 &&
                   _config.RegionHeight > 0;
        }

        private void UpdateRegionOverlay()
        {
            try
            {
                if (HasValidCustomRegion())
                {
                    if (_regionHighlightWindow == null)
                    {
                        _regionHighlightWindow = new RegionHighlightWindow();
                        // 只在主窗口已显示时设置 Owner
                        if (IsLoaded && IsVisible)
                        {
                            _regionHighlightWindow.Owner = this;
                        }
                    }
                    else if (IsLoaded && IsVisible && _regionHighlightWindow.Owner == null)
                    {
                        // 如果之前没有设置 Owner，现在设置
                        _regionHighlightWindow.Owner = this;
                    }

                    var rect = new Int32Rect(
                        _config.RegionLeft,
                        _config.RegionTop,
                        _config.RegionWidth,
                        _config.RegionHeight);

                    _regionHighlightWindow.ShowRegion(rect);
                }
                else
                {
                    _regionHighlightWindow?.HideRegion();
                }
            }
            catch (Exception ex)
            {
                WriteError($"更新区域高亮失败", ex);
            }
        }

        private void BtnSelectRegion_Click(object sender, RoutedEventArgs e)
        {
            RegionSelectionWindow? selector = null;
            try
            {
                if (_isVideoRecording)
                {
                    if (StatusBarInfo != null)
                    {
                        StatusBarInfo.Text = "请先停止录制，再调整录制区域";
                    }
                    return;
                }

                selector = new RegionSelectionWindow
                {
                    Owner = this
                };

                bool? dialogResult = selector.ShowDialog();
                if (dialogResult == true)
                {
                    var rect = selector.SelectedRect;
                    if (rect.Width > 0 && rect.Height > 0)
                    {
                        // 验证区域是否在主屏幕范围内
                        var (screenWidth, screenHeight) = GetPrimaryScreenPixelSize();
                        int regionRight = rect.X + rect.Width;
                        int regionBottom = rect.Y + rect.Height;
                        
                        if (rect.X < 0 || rect.Y < 0 || regionRight > screenWidth || regionBottom > screenHeight)
                        {
                            string warningMsg = $"选择的区域部分超出主屏幕范围！\n" +
                                               $"区域: ({rect.X}, {rect.Y}) 到 ({regionRight}, {regionBottom})\n" +
                                               $"主屏幕: (0, 0) 到 ({screenWidth}, {screenHeight})\n" +
                                               $"录制时只会录制主屏幕范围内的部分。\n\n" +
                                               $"是否继续使用此区域？";
                            
                            WriteWarning(warningMsg);
                            
                            var result = MessageBox.Show(warningMsg, "区域超出范围", 
                                MessageBoxButton.YesNo, MessageBoxImage.Warning);
                            
                            if (result == MessageBoxResult.No)
                            {
                                if (StatusBarInfo != null)
                                {
                                    StatusBarInfo.Text = "已取消选择区域";
                                }
                                return;
                            }
                        }
                        
                        _config.UseCustomRegion = true;
                        _config.RegionLeft = rect.X;
                        _config.RegionTop = rect.Y;
                        _config.RegionWidth = rect.Width;
                        _config.RegionHeight = rect.Height;
                        _config.Save(_configPath);

                        UpdateRegionOverlay();
                        UpdateConfigDisplay();

                        WriteLine($"选择录制区域: 左上=({rect.X},{rect.Y}), 大小={rect.Width}x{rect.Height}");
                        if (StatusBarInfo != null)
                        {
                            StatusBarInfo.Text = $"已选择区域：{rect.Width}x{rect.Height} @ ({rect.X},{rect.Y})";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteError($"选择录制区域失败", ex);
                if (StatusBarInfo != null)
                {
                    StatusBarInfo.Text = $"选择录制区域失败：{ex.Message}";
                }
            }
            finally
            {
                // 确保窗口正确关闭和清理
                try
                {
                    if (selector != null && selector.IsVisible)
                    {
                        selector.Close();
                    }
                }
                catch
                {
                    // 忽略关闭错误
                }
            }
        }

        private void BtnClearRegion_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_isVideoRecording)
                {
                    if (StatusBarInfo != null)
                    {
                        StatusBarInfo.Text = "录制中无法清除区域，请先停止录制";
                    }
                    return;
                }

                _config.UseCustomRegion = false;
                _config.RegionLeft = 0;
                _config.RegionTop = 0;
                _config.RegionWidth = 0;
                _config.RegionHeight = 0;
                _config.Save(_configPath);

                UpdateRegionOverlay();
                UpdateConfigDisplay();

                WriteLine("已清除自定义录制区域设置，恢复全屏录制");
                if (StatusBarInfo != null)
                {
                    StatusBarInfo.Text = "已清除自定义录制区域，恢复全屏";
                }
            }
            catch (Exception ex)
            {
                WriteError($"清除录制区域失败", ex);
                if (StatusBarInfo != null)
                {
                    StatusBarInfo.Text = $"清除录制区域失败：{ex.Message}";
                }
            }
        }

        private (int width, int height) GetPrimaryScreenPixelSize()
        {
            int width = GetSystemMetrics(SM_CXSCREEN);
            int height = GetSystemMetrics(SM_CYSCREEN);

            if (width <= 0 || height <= 0)
            {
                // 回退到 WPF 的 SystemParameters（单位为 DIP，需要乘 DPI 比例，但通常足够）
                width = (int)SystemParameters.PrimaryScreenWidth;
                height = (int)SystemParameters.PrimaryScreenHeight;
            }

            return (width, height);
        }

        private const int SM_CXSCREEN = 0;
        private const int SM_CYSCREEN = 1;

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        /// <summary>
        /// 更新日志菜单项状态
        /// </summary>
        private void UpdateLogMenuItems()
        {
            try
            {
                // 更新日志开关状态
                if (MenuLogEnabled != null)
                {
                    MenuLogEnabled.IsChecked = _config.LogEnabled == 1;
                }
                if (MenuLogDisabled != null)
                {
                    MenuLogDisabled.IsChecked = _config.LogEnabled == 0;
                }

                // 更新日志文件模式状态
                if (MenuLogFileOverwrite != null)
                {
                    MenuLogFileOverwrite.IsChecked = _config.LogFileMode == 0;
                }
                if (MenuLogFileAppend != null)
                {
                    MenuLogFileAppend.IsChecked = _config.LogFileMode == 1;
                }
            }
            catch (Exception ex)
            {
                WriteError($"更新日志菜单项状态失败", ex);
            }
        }

        /// <summary>
        /// 保存日志配置
        /// </summary>
        private void SaveLogConfig()
        {
            try
            {
                _config.Save(_configPath);
                Logger.SetLogFileMode(_config.LogFileMode);
                // 如果切换到覆盖模式，需要重新设置日志目录以清空文件
                if (_config.LogFileMode == 0)
                {
                    Logger.SetLogDirectory(_workDir, applyMode: true);
                }
                // 只有在日志启用时才写入日志（避免在禁用时写入）
                if (_config.LogEnabled == 1)
                {
                    Logger.WriteLine($"日志配置已更新: 状态={(_config.LogEnabled == 1 ? "启用" : "禁用")}, 模式={(_config.LogFileMode == 0 ? "覆盖" : "叠加")}");
                }
            }
            catch (Exception ex)
            {
                WriteError($"保存日志配置失败", ex);
            }
        }

        private void MenuLogEnabled_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _config.LogEnabled = 1;
                Logger.Enabled = true;
                UpdateLogMenuItems();
                SaveLogConfig();
                if (StatusBarInfo != null)
                {
                    StatusBarInfo.Text = "日志已启用";
                }
            }
            catch (Exception ex)
            {
                WriteError($"启用日志失败", ex);
            }
        }

        private void MenuLogDisabled_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _config.LogEnabled = 0;
                Logger.Enabled = false;
                UpdateLogMenuItems();
                SaveLogConfig();
                if (StatusBarInfo != null)
                {
                    StatusBarInfo.Text = "日志已禁用";
                }
            }
            catch (Exception ex)
            {
                WriteError($"禁用日志失败", ex);
            }
        }

        private void MenuLogFileOverwrite_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _config.LogFileMode = 0;
                UpdateLogMenuItems();
                SaveLogConfig();
                if (StatusBarInfo != null)
                {
                    StatusBarInfo.Text = "日志文件模式：覆盖";
                }
            }
            catch (Exception ex)
            {
                WriteError($"设置日志文件覆盖模式失败", ex);
            }
        }

        private void MenuLogFileAppend_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _config.LogFileMode = 1;
                UpdateLogMenuItems();
                SaveLogConfig();
                if (StatusBarInfo != null)
                {
                    StatusBarInfo.Text = "日志文件模式：叠加";
                }
            }
            catch (Exception ex)
            {
                WriteError($"设置日志文件叠加模式失败", ex);
            }
        }
    }
}
