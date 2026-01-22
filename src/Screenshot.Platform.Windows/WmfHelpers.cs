using System;
using System.Runtime.InteropServices;

namespace Screenshot.Platform.Windows
{
    internal static class WmfHelpers
    {
        [DllImport("mfplat.dll", ExactSpelling = true)]
        public static extern int MFSetAttributeSize(IMFAttributes pAttributes, Guid guidKey, int unWidth, int unHeight);

        [DllImport("mfplat.dll", ExactSpelling = true)]
        public static extern int MFSetAttributeRatio(IMFAttributes pAttributes, Guid guidKey, int unNumerator, int unDenominator);

        public static void SetUInt32(IMFAttributes attrs, Guid key, int value)
        {
            var hr = attrs.SetUINT32(key, value);
            WmfInterop.ThrowIfFailed(hr, "SetUINT32 failed");
        }

        public static void SetUInt64(IMFAttributes attrs, Guid key, long value)
        {
            var hr = attrs.SetUINT64(key, value);
            WmfInterop.ThrowIfFailed(hr, "SetUINT64 failed");
        }

        public static void SetGuid(IMFAttributes attrs, Guid key, Guid value)
        {
            var hr = attrs.SetGUID(key, value);
            WmfInterop.ThrowIfFailed(hr, "SetGUID failed");
        }
    }
}
