using System;
using System.IO;
using NAudio.Wave;
using NAudio.Lame;
using NAudio.CoreAudioApi;

namespace Screenshot_v3_0
{
    /// <summary>
    /// 音频录制器（使用NAudio录制系统音频 - 声卡输出）
    /// </summary>
    public sealed class AudioRecorder : IDisposable
    {
        private bool _isRecording;
        private string? _outputPath;
        private WasapiLoopbackCapture? _loopbackCapture;
        private WaveFileWriter? _waveFileWriter;
        private string? _tempWavPath;

        public bool IsRecording => _isRecording;

        /// <param name="outAacPath">最终输出的文件路径</param>
        public void Start(string outAacPath)
        {
            if (_isRecording) return;

            try
            {
                _outputPath = outAacPath;
                
                // 创建临时WAV文件路径
                var tempDir = Path.GetDirectoryName(outAacPath);
                var tempFileName = $"temp_{Path.GetFileNameWithoutExtension(outAacPath)}.wav";
                _tempWavPath = Path.Combine(tempDir!, tempFileName);

                // 使用WASAPI Loopback捕获系统音频（声卡输出）
                _loopbackCapture = new WasapiLoopbackCapture();

                // 设置数据可用事件
                _loopbackCapture.DataAvailable += LoopbackCapture_DataAvailable;
                _loopbackCapture.RecordingStopped += LoopbackCapture_RecordingStopped;

                // 创建WAV文件写入器
                _waveFileWriter = new WaveFileWriter(_tempWavPath, _loopbackCapture.WaveFormat);

                // 开始录制系统音频
                _loopbackCapture.StartRecording();
                _isRecording = true;

                System.Diagnostics.Debug.WriteLine($"开始录制系统音频到: {_tempWavPath}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"开始录制系统音频失败: {ex.Message}");
                throw;
            }
        }

        public void Stop()
        {
            if (!_isRecording) return;
            
            try
            {
                _isRecording = false;
                _loopbackCapture?.StopRecording();
                System.Diagnostics.Debug.WriteLine($"停止录制系统音频，准备转换文件");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"停止录制系统音频失败: {ex.Message}");
                throw;
            }
        }

        private void LoopbackCapture_DataAvailable(object? sender, WaveInEventArgs e)
        {
            if (_waveFileWriter != null)
            {
                _waveFileWriter.Write(e.Buffer, 0, e.BytesRecorded);
            }
        }

        private void LoopbackCapture_RecordingStopped(object? sender, StoppedEventArgs e)
        {
            try
            {
                // 关闭WAV文件
                _waveFileWriter?.Dispose();
                _waveFileWriter = null;

                // 转换WAV到M4A
                if (!string.IsNullOrEmpty(_tempWavPath) && !string.IsNullOrEmpty(_outputPath) && File.Exists(_tempWavPath))
                {
                    ConvertWavToM4A(_tempWavPath, _outputPath);
                    
                    // 删除临时WAV文件
                    File.Delete(_tempWavPath);
                }

                System.Diagnostics.Debug.WriteLine($"音频文件已保存到: {_outputPath}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"处理录制文件时出错: {ex.Message}");
            }
        }

        private void ConvertWavToM4A(string wavPath, string m4aPath)
        {
            try
            {
                using (var reader = new AudioFileReader(wavPath))
                using (var writer = new LameMP3FileWriter(m4aPath, reader.WaveFormat, 128))
                {
                    reader.CopyTo(writer);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"转换音频格式失败: {ex.Message}");
                // 如果转换失败，至少保留WAV文件
                File.Copy(wavPath, m4aPath, true);
            }
        }

        public void Dispose()
        {
            if (_isRecording) Stop();
            
            _loopbackCapture?.Dispose();
            _waveFileWriter?.Dispose();
        }
    }
}

