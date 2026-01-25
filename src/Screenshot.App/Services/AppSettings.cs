using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using Screenshot.Core;

namespace Screenshot.App.Services
{
    internal sealed class AppSettings
    {
        public string? OutputDirectory { get; set; }
        public string? LogDirectory { get; set; }
        public string? SessionName { get; set; }
        public bool GeneratePpt { get; set; } = true;
        public bool KeepJpgFiles { get; set; } = true;
        public string ScreenshotInterval { get; set; } = "10";
        public string ScreenChangeRate { get; set; } = "11.12";
        public bool LogEnabled { get; set; } = true;
        public bool LogAppendMode { get; set; } = false;
        public OutputMode SelectedOutputMode { get; set; } = OutputMode.AudioAndVideo;
        public AudioCaptureMode SelectedAudioCaptureMode { get; set; } = AudioCaptureMode.NativeSystemAudio;
        public bool UseCustomRegion { get; set; } = false;
        public string RegionLeft { get; set; } = "0";
        public string RegionTop { get; set; } = "0";
        public string RegionWidth { get; set; } = "0";
        public string RegionHeight { get; set; } = "0";
        public string RecordingDurationMinutes { get; set; } = "60";
        public int VideoMergeMode { get; set; } = 1;

        public static AppSettings Load(string path)
        {
            try
            {
                if (!File.Exists(path)) return new AppSettings();
                var json = File.ReadAllText(path);
                var options = CreateOptions();
                return JsonSerializer.Deserialize<AppSettings>(json, options) ?? new AppSettings();
            }
            catch
            {
                return new AppSettings();
            }
        }

        public void Save(string path)
        {
            try
            {
                var directory = Path.GetDirectoryName(path);
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var options = CreateOptions();
                var json = JsonSerializer.Serialize(this, options);
                File.WriteAllText(path, json);
            }
            catch
            {
                // ignore settings save failures
            }
        }

        private static JsonSerializerOptions CreateOptions()
        {
            return new JsonSerializerOptions
            {
                WriteIndented = true,
                Converters = { new JsonStringEnumConverter() }
            };
        }
    }
}
