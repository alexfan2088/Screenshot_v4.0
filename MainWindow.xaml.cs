using System;
using System.IO;
using System.Windows;

namespace Screenshot_v3_0
{
    public partial class MainWindow : Window
    {
        private readonly Screenshot_v3_0.AudioRecorder _rec = new();
        private readonly string _workDir;

        public MainWindow()
        {
            InitializeComponent();
            
            // 设置工作目录：优先使用D盘，如果D盘不存在则使用C盘
            string driveLetter = Directory.Exists("D:\\") ? "D:" : "C:";
            _workDir = Path.Combine(driveLetter, "ScreenshotV3.0");
            Directory.CreateDirectory(_workDir);
        }

        private void BtnStartAudio_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var aacPath = Path.Combine(_workDir, $"audio{DateTime.Now:yyMMddHHmmss}.m4a");
                _rec.Start(aacPath);
                TxtLog.Text = $"开始录音 → {aacPath}\n（停止后自动转AAC并删除临时WAV）";
            }
            catch (Exception ex)
            {
                TxtLog.Text = "开始录音失败：" + ex.Message;
            }
        }

        private void BtnStopAudio_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _rec.Stop();
                TxtLog.Text = "已请求停止录音（编码完成需几秒，视音频长度而定）";
            }
            catch (Exception ex)
            {
                TxtLog.Text = "停止录音失败：" + ex.Message;
            }
        }
    }
}
