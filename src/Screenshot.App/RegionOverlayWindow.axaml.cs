using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace Screenshot.App
{
    public partial class RegionOverlayWindow : Window
    {
        public RegionOverlayWindow()
        {
            InitializeComponent();
            IsHitTestVisible = false;
            Focusable = false;
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
    }
}
