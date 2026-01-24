using Avalonia.Controls;
using Avalonia.Platform.Storage;
using Screenshot.App.ViewModels;
using Avalonia;
using Avalonia.Input;
using Avalonia.Threading;
using System.Linq;

namespace Screenshot.App
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Opened += (_, _) => PositionTopCenter();
            HookNumericInputBehavior();
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

        private async void OnPickOutputDirectoryClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            if (DataContext is not MainViewModel vm) return;
            if (vm.IsEditingLocked)
            {
                vm.StatusMessage = "录制中无法修改目录";
                return;
            }

            var storage = StorageProvider;
            if (storage is null) return;

            var folders = await storage.OpenFolderPickerAsync(new FolderPickerOpenOptions
            {
                Title = "选择输出目录",
                AllowMultiple = false
            });

            var path = folders.FirstOrDefault()?.TryGetLocalPath();
            if (!string.IsNullOrWhiteSpace(path))
            {
                vm.OutputDirectory = path;
                vm.StatusMessage = $"已更新输出目录: {path}";
            }
        }

        private async void OnPickLogDirectoryClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            if (DataContext is not MainViewModel vm) return;
            if (vm.IsEditingLocked)
            {
                vm.StatusMessage = "录制中无法修改目录";
                return;
            }

            var storage = StorageProvider;
            if (storage is null) return;

            var folders = await storage.OpenFolderPickerAsync(new FolderPickerOpenOptions
            {
                Title = "选择日志目录",
                AllowMultiple = false
            });

            var path = folders.FirstOrDefault()?.TryGetLocalPath();
            if (!string.IsNullOrWhiteSpace(path))
            {
                vm.UpdateLogDirectory(path);
            }
        }

        private void OnOutputMenuClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            if (sender is not Button button) return;
            button.ContextMenu?.Open(button);
        }

        private void HookNumericInputBehavior()
        {
            var boxes = new[]
            {
                this.FindControl<TextBox>("ScreenChangeRateBox"),
                this.FindControl<TextBox>("ScreenshotIntervalBox"),
                this.FindControl<TextBox>("RecordingDurationBox")
            };

            foreach (var textBox in boxes)
            {
                if (textBox == null) continue;
                textBox.PointerEntered += OnNumericTextBoxPointerEntered;
                textBox.PointerExited += OnNumericTextBoxPointerExited;
            }
        }

        private static readonly TimeSpan HoverSelectDelay = TimeSpan.FromSeconds(1);

        private void OnNumericTextBoxPointerEntered(object? sender, PointerEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                var timer = new DispatcherTimer { Interval = HoverSelectDelay };
                timer.Tick += (_, _) =>
                {
                    timer.Stop();
                    if (textBox.IsFocused)
                    {
                        textBox.SelectAll();
                    }
                    else
                    {
                        textBox.Focus();
                        textBox.SelectAll();
                    }
                };
                textBox.Tag = timer;
                timer.Start();
            }
        }

        private void OnNumericTextBoxPointerExited(object? sender, PointerEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                if (textBox.Tag is DispatcherTimer timer)
                {
                    timer.Stop();
                }
                textBox.Tag = null;
            }
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

        private void PositionTopCenter()
        {
            var screen = Screens?.Primary;
            if (screen == null) return;
            var area = screen.WorkingArea;
            var width = (int)Math.Max(0, Bounds.Width);
            var x = area.X + Math.Max(0, (area.Width - width) / 2);
            var y = area.Y;
            Position = new PixelPoint(x, y);
        }
    }
}
