using System;
using System.Runtime.InteropServices;

namespace Screenshot.Platform.Windows
{
    internal static class WmfInterop
    {
        public const int MF_VERSION = 0x0002;
        public const int MFSTARTUP_FULL = 0;

        public static readonly Guid MFMediaType_Video = new("73646976-0000-0010-8000-00AA00389B71");
        public static readonly Guid MFMediaType_Audio = new("73647561-0000-0010-8000-00AA00389B71");
        public static readonly Guid MFVideoFormat_H264 = new("34363248-0000-0010-8000-00AA00389B71");
        public static readonly Guid MFAudioFormat_AAC = new("00001610-0000-0010-8000-00AA00389B71");
        public static readonly Guid MFAudioFormat_PCM = new("00000001-0000-0010-8000-00AA00389B71");

        public static readonly Guid MF_MT_MAJOR_TYPE = new("48eba18e-f8c9-4687-bf11-0a74c9f96a8f");
        public static readonly Guid MF_MT_SUBTYPE = new("f7e34c9a-42e8-4714-b74b-cb29d72c35e5");
        public static readonly Guid MF_MT_AVG_BITRATE = new("20332624-fb0d-4d9e-bd0d-cbf6786c102e");
        public static readonly Guid MF_MT_INTERLACE_MODE = new("e2724bb8-e676-4806-b4b2-a8d6efb44ccd");
        public static readonly Guid MF_MT_FRAME_SIZE = new("1652c33d-d6b2-4012-b834-72030849a37d");
        public static readonly Guid MF_MT_FRAME_RATE = new("c459a2e8-3d2c-4e44-b132-fee5156c7bb0");
        public static readonly Guid MF_MT_PIXEL_ASPECT_RATIO = new("c6376a1e-8d0a-4027-be45-6d9a0ad39bb6");
        public static readonly Guid MF_MT_DEFAULT_STRIDE = new("644b4e48-1e02-4516-b0eb-c01ca9d49ac6");
        public static readonly Guid MF_MT_AUDIO_NUM_CHANNELS = new("37e48bf5-645e-4c5b-89de-ada9e29b696a");
        public static readonly Guid MF_MT_AUDIO_SAMPLES_PER_SECOND = new("5faeeae7-0290-4c31-8d1a-cf3962b0b6cf");
        public static readonly Guid MF_MT_AUDIO_BITS_PER_SAMPLE = new("f2deb57f-40fa-4764-aa33-ed4f2d1ff669");
        public static readonly Guid MF_MT_AUDIO_BLOCK_ALIGNMENT = new("322de230-9eeb-43bd-ab7a-ff412251541d");
        public static readonly Guid MF_MT_AUDIO_AVG_BYTES_PER_SECOND = new("1aab75c8-cfef-451c-ab95-ac034b8e1731");
        public static readonly Guid MF_MT_AUDIO_AAC_PAYLOAD_TYPE = new("b8ebefaf-b718-4e04-bf58-5f7f3b9d4c8f");
        public static readonly Guid MF_MT_AAC_AUDIO_PROFILE_LEVEL_INDICATION = new("7632f0e6-9532-4d02-86ed-4d8bf9a9ab6e");
        public static readonly Guid MF_MT_ALL_SAMPLES_INDEPENDENT = new("c9173739-5e56-461c-b713-46fb995cb95f");

        public const int MFVideoInterlace_Progressive = 2;

        [DllImport("mfplat.dll", ExactSpelling = true)]
        public static extern int MFStartup(int version, int dwFlags);

        [DllImport("mfplat.dll", ExactSpelling = true)]
        public static extern int MFShutdown();

        [DllImport("mfplat.dll", ExactSpelling = true)]
        public static extern int MFCreateMediaType(out IMFMediaType mediaType);

        [DllImport("mfplat.dll", ExactSpelling = true)]
        public static extern int MFCreateSample(out IMFSample sample);

        [DllImport("mfplat.dll", ExactSpelling = true)]
        public static extern int MFCreateMemoryBuffer(int cbMaxLength, out IMFMediaBuffer buffer);

        [DllImport("mfreadwrite.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
        public static extern int MFCreateSinkWriterFromURL(string pwszOutputURL, IntPtr pByteStream, IntPtr pAttributes, out IMFSinkWriter sinkWriter);

        public static void ThrowIfFailed(int hr, string message)
        {
            if (hr < 0)
            {
                Marshal.ThrowExceptionForHR(hr, new IntPtr(-1));
                throw new InvalidOperationException(message + $" (HRESULT: 0x{hr:X8})");
            }
        }
    }

    [ComImport, Guid("44ae0fa8-ea31-4109-8d2e-4cae4997c555"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMFAttributes
    {
        int GetItem([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, [In, Out] ref PropVariant pValue);
        int SetItem([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, [In] ref PropVariant pValue);
        int CompareItem([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, [In] ref PropVariant pValue, out bool pbResult);
        int Compare([MarshalAs(UnmanagedType.Interface)] IMFAttributes pTheirs, int MatchType, out bool pbResult);
        int GetUINT32([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, out int punValue);
        int GetUINT64([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, out long punValue);
        int GetDouble([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, out double pfValue);
        int GetGUID([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, out Guid pguidValue);
        int GetStringLength([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, out int pcchLength);
        int GetString([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, [Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder pwszValue, int cchBufSize, out int pcchLength);
        int GetAllocatedString([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, [MarshalAs(UnmanagedType.LPWStr)] out string ppwszValue, out int pcchLength);
        int GetBlobSize([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, out int pcbBlobSize);
        int GetBlob([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, [Out] byte[] pBuf, int cbBufSize, out int pcbBlobSize);
        int GetAllocatedBlob([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, out IntPtr ip, out int pcbSize);
        int GetUnknown([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, [In] ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv);
        int SetUINT32([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, int unValue);
        int SetUINT64([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, long unValue);
        int SetDouble([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, double fValue);
        int SetGUID([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, [In, MarshalAs(UnmanagedType.LPStruct)] Guid guidValue);
        int SetString([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, [In, MarshalAs(UnmanagedType.LPWStr)] string wszValue);
        int SetBlob([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, [In] byte[] pBuf, int cbBufSize);
        int SetUnknown([In, MarshalAs(UnmanagedType.LPStruct)] Guid guidKey, [MarshalAs(UnmanagedType.IUnknown)] object pUnknown);
        int LockStore();
        int UnlockStore();
        int GetCount(out int pcItems);
        int GetItemByIndex(int unIndex, out Guid pguidKey, [In, Out] ref PropVariant pValue);
        int CopyAllItems([MarshalAs(UnmanagedType.Interface)] IMFAttributes pDest);
    }

    [ComImport, Guid("2cd2d921-c447-44a7-a13c-4adabfc247e3"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMFMediaType : IMFAttributes
    {
    }

    [ComImport, Guid("df598932-f10c-4e39-bba2-c308f101daa3"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMFSample
    {
        int GetSampleFlags(out int pdwSampleFlags);
        int SetSampleFlags(int dwSampleFlags);
        int GetSampleTime(out long phnsSampleTime);
        int SetSampleTime(long hnsSampleTime);
        int GetSampleDuration(out long phnsSampleDuration);
        int SetSampleDuration(long hnsSampleDuration);
        int GetBufferCount(out int pdwBufferCount);
        int GetBufferByIndex(int dwIndex, out IMFMediaBuffer ppBuffer);
        int ConvertToContiguousBuffer(out IMFMediaBuffer ppBuffer);
        int AddBuffer(IMFMediaBuffer pBuffer);
        int RemoveBufferByIndex(int dwIndex);
        int RemoveAllBuffers();
        int GetTotalLength(out int pcbTotalLength);
        int CopyToBuffer(IMFMediaBuffer pBuffer);
    }

    [ComImport, Guid("045FA593-8799-42b8-BC8D-8968C6453507"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMFMediaBuffer
    {
        int Lock(out IntPtr ppbBuffer, out int pcbMaxLength, out int pcbCurrentLength);
        int Unlock();
        int GetCurrentLength(out int pcbCurrentLength);
        int SetCurrentLength(int cbCurrentLength);
        int GetMaxLength(out int pcbMaxLength);
    }

    [ComImport, Guid("3137f1cd-fe5e-4805-a5d8-fb477448cb3d"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMFSinkWriter
    {
        int AddStream(IMFMediaType pTargetMediaType, out int pdwStreamIndex);
        int SetInputMediaType(int dwStreamIndex, IMFMediaType pInputMediaType, IntPtr pEncodingParameters);
        int BeginWriting();
        int WriteSample(int dwStreamIndex, IMFSample pSample);
        int SendStreamTick(int dwStreamIndex, long llTimestamp);
        int PlaceMarker(int dwStreamIndex, IntPtr pContext);
        int NotifyEndOfSegment(int dwStreamIndex);
        int Flush(int dwStreamIndex);
        int Finalize_();
        int GetServiceForStream(int dwStreamIndex, ref Guid guidService, ref Guid riid, out IntPtr ppvObject);
        int GetStatistics(int dwStreamIndex, out IntPtr pStats);
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct PropVariant
    {
        public short vt;
        public short wReserved1;
        public short wReserved2;
        public short wReserved3;
        public IntPtr p;
        public int p2;

        public static PropVariant FromLong(long value)
        {
            return new PropVariant { vt = 20, p = new IntPtr(value) };
        }

        public static PropVariant FromInt(int value)
        {
            return new PropVariant { vt = 3, p2 = value };
        }

        public static PropVariant FromGuid(Guid value)
        {
            var pv = new PropVariant { vt = 72 };
            pv.p = Marshal.AllocHGlobal(Marshal.SizeOf<Guid>());
            Marshal.StructureToPtr(value, pv.p, false);
            return pv;
        }
    }
}
