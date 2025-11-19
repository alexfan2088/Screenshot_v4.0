using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
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
            // 设置日志文件目录为工作目录（与视频、音频文件同一目录）
            // SetLogDirectory 会自动清空日志文件（覆盖模式）
            Logger.SetLogDirectory(_workDir);
            Logger.WriteLine("程序启动");
            Logger.WriteLine($"日志状态: {(Logger.Enabled ? "启用" : "禁用")}");
            
            UpdateConfigDisplay();
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            // 立即允许关闭，不阻塞
            e.Cancel = false;

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
                
                // 计算分辨率
                double resolutionScale = _config.VideoResolutionScale / 100.0;
                int screenWidth = (int)SystemParameters.PrimaryScreenWidth;
                int screenHeight = (int)SystemParameters.PrimaryScreenHeight;
                int videoWidth = (int)(screenWidth * resolutionScale);
                int videoHeight = (int)(screenHeight * resolutionScale);

                // 创建视频编码器（不再需要 VideoRecorder，FFmpeg 直接录制）
                _videoEncoder = new VideoEncoder(videoPath, _config);

                // 初始化编码器（offsetX 和 offsetY 都是 0，从屏幕左上角开始）
                _videoEncoder.Initialize(videoWidth, videoHeight, _config.VideoFrameRate, 
                    _config.AudioSampleRate, 2, offsetX: 0, offsetY: 0);

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
                    StatusBarInfo.Text = $"开始录制视频 → {videoPath} | 分辨率: {videoWidth}x{videoHeight} @ {_config.VideoFrameRate}fps";
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
                    StatusBarConfig.Text = $"视频：{_config.VideoResolutionScale}%分辨率，{_config.VideoFrameRate}fps，H.264 | " +
                                         $"音频：{_config.AudioSampleRate / 1000}kHz，{_config.AudioBitrate}kbps";
                }
            }
            catch (Exception ex)
            {
                WriteError($"更新配置显示失败", ex);
            }
        }
    }
}
