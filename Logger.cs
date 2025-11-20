using System;
using System.IO;
using System.Threading;

namespace Screenshot_v3_0
{
    /// <summary>
    /// 日志记录器
    /// </summary>
    public static class Logger
    {
        private static bool _enabled = true; // 默认启用日志
        private static int _logFileMode = 0; // 日志文件模式 0=覆盖，1=叠加
        private static string? _logFilePath;
        private static readonly object _lockObject = new object();

        /// <summary>
        /// 初始化日志系统
        /// </summary>
        static Logger()
        {
            try
            {
                // 默认使用程序目录，可以通过 SetLogDirectory 修改
                string exeDir = AppDomain.CurrentDomain.BaseDirectory;
                string logFileName = $"log_{DateTime.Now:yyyyMMdd}.txt";
                _logFilePath = Path.Combine(exeDir, logFileName);
            }
            catch
            {
                _logFilePath = null;
            }
        }

        /// <summary>
        /// 设置日志文件模式（0=覆盖，1=叠加）
        /// </summary>
        public static void SetLogFileMode(int mode)
        {
            _logFileMode = mode;
        }

        /// <summary>
        /// 获取日志文件模式（0=覆盖，1=叠加）
        /// </summary>
        public static int LogFileMode => _logFileMode;

        /// <summary>
        /// 设置日志文件目录
        /// </summary>
        public static void SetLogDirectory(string directory, bool applyMode = true)
        {
            try
            {
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
                string logFileName = $"log_{DateTime.Now:yyyyMMdd}.txt";
                _logFilePath = Path.Combine(directory, logFileName);
                // 根据模式决定是否清空日志文件
                if (applyMode && _logFileMode == 0)
                {
                    ClearLog();
                }
            }
            catch
            {
                // 如果设置失败，保持原有路径
            }
        }

        /// <summary>
        /// 清空日志文件（覆盖模式）
        /// </summary>
        public static void ClearLog()
        {
            if (string.IsNullOrEmpty(_logFilePath)) return;

            try
            {
                lock (_lockObject)
                {
                    // 如果文件存在，删除它（下次写入时会创建新文件）
                    if (File.Exists(_logFilePath))
                    {
                        File.Delete(_logFilePath);
                    }
                }
            }
            catch
            {
                // 忽略清空日志错误
            }
        }

        /// <summary>
        /// 启用或禁用日志（1=启用，0=禁用）
        /// </summary>
        public static bool Enabled
        {
            get => _enabled;
            set => _enabled = value;
        }

        /// <summary>
        /// 写入日志
        /// </summary>
        public static void WriteLine(string message)
        {
            if (!_enabled || string.IsNullOrEmpty(_logFilePath)) return;

            try
            {
                string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}";
                
                lock (_lockObject)
                {
                    // 如果文件不存在，创建新文件；如果存在，根据模式决定追加或覆盖
                    if (_logFileMode == 0 && !File.Exists(_logFilePath))
                    {
                        // 覆盖模式：如果文件不存在，创建新文件
                        File.WriteAllText(_logFilePath, logMessage + Environment.NewLine);
                    }
                    else
                    {
                        // 叠加模式：总是追加；覆盖模式：文件已存在时也追加（因为 SetLogDirectory 已清空）
                        File.AppendAllText(_logFilePath, logMessage + Environment.NewLine);
                    }
                }

                // 同时输出到调试窗口
                System.Diagnostics.Debug.WriteLine(logMessage);
            }
            catch
            {
                // 忽略日志写入错误，避免影响主程序
            }
        }

        /// <summary>
        /// 写入错误日志
        /// </summary>
        public static void WriteError(string message, Exception? ex = null)
        {
            if (!_enabled) return;

            string errorMessage = $"错误: {message}";
            if (ex != null)
            {
                errorMessage += $"\n异常: {ex.Message}";
                if (ex.StackTrace != null)
                {
                    errorMessage += $"\n堆栈: {ex.StackTrace}";
                }
            }

            WriteLine(errorMessage);
        }

        /// <summary>
        /// 写入警告日志
        /// </summary>
        public static void WriteWarning(string message)
        {
            if (!_enabled) return;
            WriteLine($"警告: {message}");
        }

        /// <summary>
        /// 写入信息日志
        /// </summary>
        public static void WriteInfo(string message)
        {
            if (!_enabled) return;
            WriteLine($"信息: {message}");
        }
    }
}

