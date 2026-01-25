using Avalonia.Controls;
using Avalonia.Platform.Storage;
using Screenshot.App.ViewModels;
using Avalonia;
using Avalonia.Input;
using Avalonia.Threading;
using System.ComponentModel;
using System.Linq;

namespace Screenshot.App
{
    public partial class MainWindow : Window
    {
        private MainViewModel? _vm;
        private bool _isOpened;

        public MainWindow()
        {
            InitializeComponent();
            Opened += (_, _) =>
            {
                _isOpened = true;
                PositionTopCenter();
                UpdateRegionOverlay();
            };
            HookNumericInputBehavior();
            DataContextChanged += OnDataContextChanged;
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
                vm.ApplyCustomRegion(value.X, value.Y, value.Width, value.Height, remember: true);
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

        private void OnDataContextChanged(object? sender, EventArgs e)
        {
            if (_vm != null)
            {
                _vm.PropertyChanged -= OnVmPropertyChanged;
            }

            _vm = DataContext as MainViewModel;
            if (_vm != null)
            {
                _vm.PropertyChanged += OnVmPropertyChanged;
                UpdateRegionOverlay();
            }
        }

        private void OnVmPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(MainViewModel.UseCustomRegion) ||
                e.PropertyName == nameof(MainViewModel.RegionLeft) ||
                e.PropertyName == nameof(MainViewModel.RegionTop) ||
                e.PropertyName == nameof(MainViewModel.RegionWidth) ||
                e.PropertyName == nameof(MainViewModel.RegionHeight))
            {
                if (_vm?.IsRecording == true)
                {
                    UpdateRegionOverlay();
                }
            }
        }

        private void UpdateRegionOverlay()
        {
            if (!_isOpened) return;
            if (_vm == null) return;
            var bounds = GetVirtualScreenBounds();
            if (!bounds.HasValue) return;

            var useCustom = _vm.UseCustomRegion;
            var left = ParseInt(_vm.RegionLeft);
            var top = ParseInt(_vm.RegionTop);
            var width = ParseInt(_vm.RegionWidth);
            var height = ParseInt(_vm.RegionHeight);
            var rect = useCustom && width > 0 && height > 0
                ? new PixelRect(left, top, width, height)
                : new PixelRect(0, 0, bounds.Value.Width, bounds.Value.Height);

            var canvas = this.FindControl<Canvas>("RegionOverlayCanvas");
            var border = this.FindControl<Border>("RegionOverlayBorder");
            if (canvas == null || border == null) return;

            canvas.Width = bounds.Value.Width;
            canvas.Height = bounds.Value.Height;
            Canvas.SetLeft(canvas, bounds.Value.X);
            Canvas.SetTop(canvas, bounds.Value.Y);

            Canvas.SetLeft(border, rect.X);
            Canvas.SetTop(border, rect.Y);
            border.Width = rect.Width;
            border.Height = rect.Height;
            border.IsVisible = true;
        }

        private static int ParseInt(string value)
        {
            return int.TryParse(value, out var result) ? result : 0;
        }
    }
}
