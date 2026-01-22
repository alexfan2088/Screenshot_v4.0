using Avalonia.Controls;
using Screenshot.App.ViewModels;
using Avalonia;

namespace Screenshot.App
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void OnOpenSettingsClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            if (DataContext is not MainViewModel vm) return;
            var window = new SettingsWindow
            {
                DataContext = vm
            };
            window.Show();
        }

        private async void OnSelectRegionClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            if (DataContext is not MainViewModel vm) return;
            if (!OperatingSystem.IsWindows())
            {
                vm.StatusMessage = "区域选择仅支持 Windows";
                return;
            }

            var window = new RegionSelectionWindow();
            var virtualBounds = GetVirtualScreenBounds();
            if (virtualBounds.HasValue)
            {
                var bounds = virtualBounds.Value;
                window.Position = bounds.Position;
                window.Width = bounds.Width;
                window.Height = bounds.Height;
            }
            window.Show();

            var rect = await window.PickAsync();
            if (rect.HasValue)
            {
                var value = rect.Value;
                vm.UseCustomRegion = true;
                vm.RegionLeft = value.X.ToString();
                vm.RegionTop = value.Y.ToString();
                vm.RegionWidth = value.Width.ToString();
                vm.RegionHeight = value.Height.ToString();
            }
        }

        private PixelRect? GetVirtualScreenBounds()
        {
            var screens = Screens;
            if (screens == null || screens.All.Count == 0)
            {
                return null;
            }

            var left = int.MaxValue;
            var top = int.MaxValue;
            var right = int.MinValue;
            var bottom = int.MinValue;

            foreach (var screen in screens.All)
            {
                var bounds = screen.Bounds;
                if (bounds.X < left) left = bounds.X;
                if (bounds.Y < top) top = bounds.Y;
                if (bounds.Right > right) right = bounds.Right;
                if (bounds.Bottom > bottom) bottom = bounds.Bottom;
            }

            if (left == int.MaxValue) return null;
            return new PixelRect(left, top, Math.Max(0, right - left), Math.Max(0, bottom - top));
        }
    }
}
