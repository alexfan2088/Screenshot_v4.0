using System;
using System.Runtime.InteropServices;

namespace Screenshot.App.Services
{
    internal sealed record MacWindowInfo(int WindowId, int X, int Y, int Width, int Height);

    internal static class MacWindowPicker
    {
        private const string CoreGraphics = "/System/Library/Frameworks/CoreGraphics.framework/CoreGraphics";
        private const string CoreFoundation = "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation";
        private const uint kCGWindowListOptionOnScreenOnly = 1 << 0;
        private const uint kCGWindowListExcludeDesktopElements = 1 << 4;
        private const uint kCGNullWindowID = 0;
        private const uint kCFStringEncodingUTF8 = 0x08000100;
        private const int kCFNumberSInt64Type = 11;
        private const int kCFNumberDoubleType = 13;

        private static readonly IntPtr KeyWindowNumber = CreateKey("kCGWindowNumber");
        private static readonly IntPtr KeyWindowLayer = CreateKey("kCGWindowLayer");
        private static readonly IntPtr KeyWindowBounds = CreateKey("kCGWindowBounds");
        private static readonly IntPtr KeyWindowOwnerPid = CreateKey("kCGWindowOwnerPID");
        private static readonly IntPtr KeyBoundsX = CreateKey("X");
        private static readonly IntPtr KeyBoundsY = CreateKey("Y");
        private static readonly IntPtr KeyBoundsWidth = CreateKey("Width");
        private static readonly IntPtr KeyBoundsHeight = CreateKey("Height");

        public static bool TryGetWindowAtPoint(int screenX, int screenY, out MacWindowInfo? info)
        {
            info = null;
            if (TryGetWindowAtPointInternal(screenX, screenY, out var found))
            {
                info = found;
                return true;
            }
            return false;
        }

        public static bool TryGetWindowBounds(int windowId, out MacWindowInfo? info)
        {
            info = null;
            if (!OperatingSystem.IsMacOS()) return false;
            if (windowId <= 0) return false;

            var list = CGWindowListCopyWindowInfo(kCGWindowListOptionOnScreenOnly | kCGWindowListExcludeDesktopElements, kCGNullWindowID);
            if (list == IntPtr.Zero) return false;

            try
            {
                var count = CFArrayGetCount(list);
                for (nint i = 0; i < count; i++)
                {
                    var dict = CFArrayGetValueAtIndex(list, i);
                    if (dict == IntPtr.Zero) continue;

                    if (!TryGetInt64(dict, KeyWindowNumber, out var foundId)) continue;
                    if (foundId != windowId) continue;
                    if (!TryGetBounds(dict, out var x, out var y, out var width, out var height)) continue;
                    if (width <= 0 || height <= 0) continue;
                    info = new MacWindowInfo((int)foundId, x, y, width, height);
                    return true;
            }
            }
            finally
            {
                CFRelease(list);
            }

            return false;
        }

        private static bool TryGetWindowAtPointInternal(int screenX, int screenY, out MacWindowInfo? info)
        {
            info = null;
            if (!OperatingSystem.IsMacOS()) return false;

            var list = CGWindowListCopyWindowInfo(kCGWindowListOptionOnScreenOnly | kCGWindowListExcludeDesktopElements, kCGNullWindowID);
            if (list == IntPtr.Zero) return false;

            try
            {
                if (TryFindWindow(list, screenX, screenY, out info)) return true;
                var flippedY = FlipPrimaryDisplayY(screenY);
                if (flippedY != screenY)
                {
                    if (TryFindWindow(list, screenX, flippedY, out info)) return true;
                }
            }
            finally
            {
                CFRelease(list);
            }

            return false;
        }

        private static bool TryFindWindow(IntPtr list, int screenX, int screenY, out MacWindowInfo? info)
        {
            info = null;
            var count = CFArrayGetCount(list);
            var currentPid = Environment.ProcessId;
            for (nint i = 0; i < count; i++)
            {
                var dict = CFArrayGetValueAtIndex(list, i);
                if (dict == IntPtr.Zero) continue;

                if (!TryGetInt64(dict, KeyWindowOwnerPid, out var ownerPid)) continue;
                if (ownerPid == currentPid) continue;

                if (!TryGetInt64(dict, KeyWindowLayer, out var layer) || layer != 0) continue;
                if (!TryGetInt64(dict, KeyWindowNumber, out var windowId)) continue;
                if (!TryGetBounds(dict, out var x, out var y, out var width, out var height)) continue;

                if (width <= 0 || height <= 0) continue;
                if (screenX < x || screenY < y || screenX > x + width || screenY > y + height) continue;

                info = new MacWindowInfo((int)windowId, x, y, width, height);
                return true;
            }
            return false;
        }

        private static int FlipPrimaryDisplayY(int y)
        {
            var display = CGMainDisplayID();
            var bounds = CGDisplayBounds(display);
            var height = (int)Math.Round(bounds.Size.Height);
            if (height <= 0) return y;
            return height - y;
        }

        private static bool TryGetBounds(IntPtr dict, out int x, out int y, out int width, out int height)
        {
            x = y = width = height = 0;
            var boundsDict = CFDictionaryGetValue(dict, KeyWindowBounds);
            if (boundsDict == IntPtr.Zero) return false;

            if (!TryGetDouble(boundsDict, KeyBoundsX, out var dx)) return false;
            if (!TryGetDouble(boundsDict, KeyBoundsY, out var dy)) return false;
            if (!TryGetDouble(boundsDict, KeyBoundsWidth, out var dw)) return false;
            if (!TryGetDouble(boundsDict, KeyBoundsHeight, out var dh)) return false;

            x = (int)Math.Round(dx);
            y = (int)Math.Round(dy);
            width = (int)Math.Round(dw);
            height = (int)Math.Round(dh);
            return true;
        }

        private static bool TryGetInt64(IntPtr dict, IntPtr key, out long value)
        {
            value = 0;
            var number = CFDictionaryGetValue(dict, key);
            if (number == IntPtr.Zero) return false;
            return CFNumberGetValue(number, kCFNumberSInt64Type, out value);
        }

        private static bool TryGetDouble(IntPtr dict, IntPtr key, out double value)
        {
            value = 0;
            var number = CFDictionaryGetValue(dict, key);
            if (number == IntPtr.Zero) return false;
            if (CFNumberGetValueDouble(number, kCFNumberDoubleType, out value)) return true;
            return CFNumberGetValue(number, kCFNumberSInt64Type, out var fallback) && (value = fallback) >= 0;
        }

        private static IntPtr CreateKey(string value)
        {
            return CFStringCreateWithCString(IntPtr.Zero, value, kCFStringEncodingUTF8);
        }

        [DllImport(CoreGraphics)]
        private static extern IntPtr CGWindowListCopyWindowInfo(uint option, uint relativeToWindow);

        [DllImport(CoreFoundation)]
        private static extern nint CFArrayGetCount(IntPtr array);

        [DllImport(CoreFoundation)]
        private static extern IntPtr CFArrayGetValueAtIndex(IntPtr array, nint index);

        [DllImport(CoreFoundation)]
        private static extern IntPtr CFDictionaryGetValue(IntPtr dict, IntPtr key);

        [DllImport(CoreFoundation)]
        private static extern IntPtr CFStringCreateWithCString(IntPtr alloc, string str, uint encoding);

        [DllImport(CoreFoundation)]
        private static extern bool CFNumberGetValue(IntPtr number, int theType, out long value);

        [DllImport(CoreFoundation, EntryPoint = "CFNumberGetValue")]
        private static extern bool CFNumberGetValueDouble(IntPtr number, int theType, out double value);

        [DllImport(CoreFoundation)]
        private static extern void CFRelease(IntPtr cf);

        [DllImport(CoreGraphics)]
        private static extern uint CGMainDisplayID();

        [DllImport(CoreGraphics)]
        private static extern CGRect CGDisplayBounds(uint display);

        [StructLayout(LayoutKind.Sequential)]
        private struct CGPoint
        {
            public double X;
            public double Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct CGSize
        {
            public double Width;
            public double Height;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct CGRect
        {
            public CGPoint Origin;
            public CGSize Size;
        }
    }
}
