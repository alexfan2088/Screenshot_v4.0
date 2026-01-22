using System;
using System.Runtime.InteropServices;

namespace Screenshot.Platform.Windows
{
    internal sealed class WmfSinkWriter : IDisposable
    {
        private WmfInterop.IMFSinkWriter? _writer;
        private int _videoStreamIndex;
        private int _audioStreamIndex;
        private bool _hasAudio;

        public void Initialize(string outputPath, int width, int height, int fps, int videoBitrate, int audioSampleRate, int audioChannels, int audioBitsPerSample, bool includeVideo, bool includeAudio)
        {
            var hr = WmfInterop.MFStartup(WmfInterop.MF_VERSION, WmfInterop.MFSTARTUP_FULL);
            WmfInterop.ThrowIfFailed(hr, "MFStartup failed");

            hr = WmfInterop.MFCreateSinkWriterFromURL(outputPath, IntPtr.Zero, IntPtr.Zero, out var writer);
            WmfInterop.ThrowIfFailed(hr, "MFCreateSinkWriterFromURL failed");
            _writer = writer;

            if (includeVideo)
            {
                ConfigureVideoStream(width, height, fps, videoBitrate);
            }
            if (includeAudio)
            {
                ConfigureAudioStream(audioSampleRate, audioChannels, audioBitsPerSample);
            }

            hr = _writer.BeginWriting();
            WmfInterop.ThrowIfFailed(hr, "BeginWriting failed");
        }

        public void WriteVideoSample(byte[] buffer, long sampleTime, long duration)
        {
            if (_writer == null) return;
            var sample = CreateSample(buffer, buffer.Length, sampleTime, duration);
            var hr = _writer.WriteSample(_videoStreamIndex, sample);
            Marshal.ReleaseComObject(sample);
            WmfInterop.ThrowIfFailed(hr, "WriteSample(video) failed");
        }

        public void WriteAudioSample(byte[] buffer, int bytes, long sampleTime, long duration)
        {
            if (_writer == null || !_hasAudio) return;
            var sample = CreateSample(buffer, bytes, sampleTime, duration);
            var hr = _writer.WriteSample(_audioStreamIndex, sample);
            Marshal.ReleaseComObject(sample);
            WmfInterop.ThrowIfFailed(hr, "WriteSample(audio) failed");
        }

        private void ConfigureVideoStream(int width, int height, int fps, int bitrate)
        {
            var hr = WmfInterop.MFCreateMediaType(out var mediaTypeOut);
            WmfInterop.ThrowIfFailed(hr, "MFCreateMediaType video out failed");

            WmfHelpers.SetGuid(mediaTypeOut, WmfInterop.MF_MT_MAJOR_TYPE, WmfInterop.MFMediaType_Video);
            WmfHelpers.SetGuid(mediaTypeOut, WmfInterop.MF_MT_SUBTYPE, WmfInterop.MFVideoFormat_H264);
            WmfHelpers.SetUInt32(mediaTypeOut, WmfInterop.MF_MT_AVG_BITRATE, bitrate);
            WmfHelpers.SetUInt32(mediaTypeOut, WmfInterop.MF_MT_INTERLACE_MODE, WmfInterop.MFVideoInterlace_Progressive);
            WmfHelpers.SetUInt32(mediaTypeOut, WmfInterop.MF_MT_ALL_SAMPLES_INDEPENDENT, 1);
            WmfInterop.ThrowIfFailed(WmfHelpers.MFSetAttributeSize(mediaTypeOut, WmfInterop.MF_MT_FRAME_SIZE, width, height), "Set frame size failed");
            WmfInterop.ThrowIfFailed(WmfHelpers.MFSetAttributeRatio(mediaTypeOut, WmfInterop.MF_MT_FRAME_RATE, fps, 1), "Set frame rate failed");
            WmfInterop.ThrowIfFailed(WmfHelpers.MFSetAttributeRatio(mediaTypeOut, WmfInterop.MF_MT_PIXEL_ASPECT_RATIO, 1, 1), "Set pixel aspect ratio failed");

            hr = _writer!.AddStream(mediaTypeOut, out _videoStreamIndex);
            WmfInterop.ThrowIfFailed(hr, "AddStream video failed");

            hr = WmfInterop.MFCreateMediaType(out var mediaTypeIn);
            WmfInterop.ThrowIfFailed(hr, "MFCreateMediaType video in failed");
            WmfHelpers.SetGuid(mediaTypeIn, WmfInterop.MF_MT_MAJOR_TYPE, WmfInterop.MFMediaType_Video);
            WmfHelpers.SetGuid(mediaTypeIn, WmfInterop.MF_MT_SUBTYPE, new Guid("00000016-0000-0010-8000-00AA00389B71")); // MFVideoFormat_RGB24
            WmfHelpers.SetUInt32(mediaTypeIn, WmfInterop.MF_MT_INTERLACE_MODE, WmfInterop.MFVideoInterlace_Progressive);
            WmfInterop.ThrowIfFailed(WmfHelpers.MFSetAttributeSize(mediaTypeIn, WmfInterop.MF_MT_FRAME_SIZE, width, height), "Set frame size in failed");
            WmfInterop.ThrowIfFailed(WmfHelpers.MFSetAttributeRatio(mediaTypeIn, WmfInterop.MF_MT_FRAME_RATE, fps, 1), "Set frame rate in failed");
            WmfInterop.ThrowIfFailed(WmfHelpers.MFSetAttributeRatio(mediaTypeIn, WmfInterop.MF_MT_PIXEL_ASPECT_RATIO, 1, 1), "Set pixel aspect in failed");
            WmfHelpers.SetUInt32(mediaTypeIn, WmfInterop.MF_MT_DEFAULT_STRIDE, width * 3);

            hr = _writer.SetInputMediaType(_videoStreamIndex, mediaTypeIn, IntPtr.Zero);
            WmfInterop.ThrowIfFailed(hr, "SetInputMediaType video failed");

            Marshal.ReleaseComObject(mediaTypeOut);
            Marshal.ReleaseComObject(mediaTypeIn);
        }

        private void ConfigureAudioStream(int sampleRate, int channels, int bitsPerSample)
        {
            if (channels <= 0 || sampleRate <= 0) return;
            _hasAudio = true;

            var hr = WmfInterop.MFCreateMediaType(out var mediaTypeOut);
            WmfInterop.ThrowIfFailed(hr, "MFCreateMediaType audio out failed");

            WmfHelpers.SetGuid(mediaTypeOut, WmfInterop.MF_MT_MAJOR_TYPE, WmfInterop.MFMediaType_Audio);
            WmfHelpers.SetGuid(mediaTypeOut, WmfInterop.MF_MT_SUBTYPE, WmfInterop.MFAudioFormat_AAC);
            WmfHelpers.SetUInt32(mediaTypeOut, WmfInterop.MF_MT_AUDIO_NUM_CHANNELS, channels);
            WmfHelpers.SetUInt32(mediaTypeOut, WmfInterop.MF_MT_AUDIO_SAMPLES_PER_SECOND, sampleRate);
            WmfHelpers.SetUInt32(mediaTypeOut, WmfInterop.MF_MT_AUDIO_BITS_PER_SAMPLE, bitsPerSample);
            WmfHelpers.SetUInt32(mediaTypeOut, WmfInterop.MF_MT_AUDIO_AAC_PAYLOAD_TYPE, 0);
            WmfHelpers.SetUInt32(mediaTypeOut, WmfInterop.MF_MT_AAC_AUDIO_PROFILE_LEVEL_INDICATION, 0x29);

            hr = _writer!.AddStream(mediaTypeOut, out _audioStreamIndex);
            WmfInterop.ThrowIfFailed(hr, "AddStream audio failed");

            hr = WmfInterop.MFCreateMediaType(out var mediaTypeIn);
            WmfInterop.ThrowIfFailed(hr, "MFCreateMediaType audio in failed");
            WmfHelpers.SetGuid(mediaTypeIn, WmfInterop.MF_MT_MAJOR_TYPE, WmfInterop.MFMediaType_Audio);
            WmfHelpers.SetGuid(mediaTypeIn, WmfInterop.MF_MT_SUBTYPE, WmfInterop.MFAudioFormat_PCM);
            WmfHelpers.SetUInt32(mediaTypeIn, WmfInterop.MF_MT_AUDIO_NUM_CHANNELS, channels);
            WmfHelpers.SetUInt32(mediaTypeIn, WmfInterop.MF_MT_AUDIO_SAMPLES_PER_SECOND, sampleRate);
            WmfHelpers.SetUInt32(mediaTypeIn, WmfInterop.MF_MT_AUDIO_BITS_PER_SAMPLE, bitsPerSample);
            var blockAlign = channels * (bitsPerSample / 8);
            var bytesPerSecond = blockAlign * sampleRate;
            WmfHelpers.SetUInt32(mediaTypeIn, WmfInterop.MF_MT_AUDIO_BLOCK_ALIGNMENT, blockAlign);
            WmfHelpers.SetUInt32(mediaTypeIn, WmfInterop.MF_MT_AUDIO_AVG_BYTES_PER_SECOND, bytesPerSecond);

            hr = _writer.SetInputMediaType(_audioStreamIndex, mediaTypeIn, IntPtr.Zero);
            WmfInterop.ThrowIfFailed(hr, "SetInputMediaType audio failed");

            Marshal.ReleaseComObject(mediaTypeOut);
            Marshal.ReleaseComObject(mediaTypeIn);
        }

        private static WmfInterop.IMFSample CreateSample(byte[] data, int length, long sampleTime, long duration)
        {
            var hr = WmfInterop.MFCreateSample(out var sample);
            WmfInterop.ThrowIfFailed(hr, "MFCreateSample failed");

            hr = WmfInterop.MFCreateMemoryBuffer(length, out var buffer);
            WmfInterop.ThrowIfFailed(hr, "MFCreateMemoryBuffer failed");

            hr = buffer.Lock(out var ptr, out _, out _);
            WmfInterop.ThrowIfFailed(hr, "Lock buffer failed");
            Marshal.Copy(data, 0, ptr, length);
            buffer.Unlock();
            buffer.SetCurrentLength(length);

            sample.AddBuffer(buffer);
            sample.SetSampleTime(sampleTime);
            sample.SetSampleDuration(duration);

            Marshal.ReleaseComObject(buffer);
            return sample;
        }

        public void Dispose()
        {
            if (_writer != null)
            {
                _writer.Finalize_();
                Marshal.ReleaseComObject(_writer);
                _writer = null;
            }
            WmfInterop.MFShutdown();
        }
    }
}
