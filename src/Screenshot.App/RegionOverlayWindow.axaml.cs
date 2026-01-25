using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Platform;
using System.Runtime.InteropServices;

namespace Screenshot.App
{
    public partial class RegionOverlayWindow : Window
    {
        public RegionOverlayWindow()
        {
            InitializeComponent();
            IsHitTestVisible = false;
            Focusable = false;
            Opened += (_, _) => EnableMousePassthrough();
        }

        private void InitializeComponent()
        {
            AvaloniaXamlLoader.Load(this);
        }

        public void SetScreenBounds(PixelRect bounds)
        {
            Position = bounds.Position;
            Width = bounds.Width;
            Height = bounds.Height;
        }

        public void SetRegion(PixelRect region)
        {
            var border = this.FindControl<Border>("HighlightBorder");
            if (border == null) return;

            Canvas.SetLeft(border, region.X);
            Canvas.SetTop(border, region.Y);
            border.Width = region.Width;
            border.Height = region.Height;
        }

        private void EnableMousePassthrough()
        {
            if (!OperatingSystem.IsMacOS()) return;
            var handle = this.TryGetPlatformHandle();
            if (handle == null || handle.Handle == IntPtr.Zero) return;
            var sel = sel_registerName("setIgnoresMouseEvents:");
            objc_msgSend(handle.Handle, sel, true);
        }

        [DllImport("/usr/lib/libobjc.A.dylib")]
        private static extern IntPtr sel_registerName(string name);

        [DllImport("/usr/lib/libobjc.A.dylib")]
        private static extern void objc_msgSend(IntPtr receiver, IntPtr selector, bool value);
    }
}
