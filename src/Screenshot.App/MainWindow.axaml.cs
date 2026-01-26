using Avalonia.Controls;
using Avalonia.Platform.Storage;
using Screenshot.App.ViewModels;
using Avalonia;
using Avalonia.Input;
using Avalonia.Threading;
using Avalonia.Media.Imaging;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Linq;

namespace Screenshot.App
{
    public partial class MainWindow : Window
    {
        private RegionOverlayWindow? _regionOverlay;
        private MainViewModel? _vm;
        private bool _isOpened;

        public MainWindow()
        {
            InitializeComponent();
            Opened += (_, _) =>
            {
                _isOpened = true;
                PositionTopCenter();
                UpdateCurrentDisplayId();
                UpdateRegionOverlay();
            };
            HookNumericInputBehavior();
            DataContextChanged += OnDataContextChanged;
            Activated += (_, _) => UpdateCurrentDisplayId();
            PositionChanged += (_, _) => UpdateCurrentDisplayId();
            Closed += (_, _) => _regionOverlay?.Close();
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
            UpdateCurrentDisplayId();
            var window = new RegionSelectionWindow();
            _regionOverlay?.Hide();
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
            _regionOverlay?.Show();
            Activate();
        }

        private async void OnSingleScreenshotClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            if (DataContext is not MainViewModel vm) return;
            if (!OperatingSystem.IsMacOS())
            {
                vm.StatusMessage = "截图仅支持 macOS";
                return;
            }

            try
            {
                UpdateCurrentDisplayId();
                var targetDir = !string.IsNullOrWhiteSpace(vm.SessionDirectoryStatus)
                    ? vm.SessionDirectoryStatus
                    : vm.SessionDirectoryPreview;
                Directory.CreateDirectory(targetDir);
                var fileName = $"截图{DateTime.Now:yyMMddHHmmssfff}.jpg";
                var imagePath = Path.Combine(targetDir, fileName);

                var regionArgs = "";
                if (vm.UseCustomRegion)
                {
                    var left = ParseInt(vm.RegionLeft);
                    var top = ParseInt(vm.RegionTop);
                    var width = ParseInt(vm.RegionWidth);
                    var height = ParseInt(vm.RegionHeight);
                    if (width > 0 && height > 0)
                    {
                        regionArgs = $"-R {left},{top},{width},{height} ";
                    }
                }

                var displayArg = vm.CurrentDisplayId > 0 ? $"-D {vm.CurrentDisplayId} " : "";
                var startInfo = new ProcessStartInfo
                {
                    FileName = "screencapture",
                    Arguments = $"{displayArg}{regionArgs}-x -t jpg \"{imagePath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var process = Process.Start(startInfo);
                if (process != null)
                {
                    await process.WaitForExitAsync();
                }

                if (!File.Exists(imagePath))
                {
                    vm.StatusMessage = "截图失败";
                    return;
                }

                var topLevel = TopLevel.GetTopLevel(this);
                if (OperatingSystem.IsMacOS())
                {
                    var escapedPath = imagePath.Replace("\"", "\\\"");
                    var script = $"set the clipboard to (read (POSIX file \\\"{escapedPath}\\\") as JPEG picture)";
                    var clipInfo = new ProcessStartInfo
                    {
                        FileName = "osascript",
                        Arguments = $"-e \"{script}\"",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };
                    Process.Start(clipInfo);
                }

                vm.StatusMessage = $"已截图: {imagePath}";
            }
            catch (Exception ex)
            {
                vm.StatusMessage = $"截图失败: {ex.Message}";
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
                UpdateRegionOverlay();
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
            if (!useCustom || width <= 0 || height <= 0)
            {
                _vm.ApplyCustomRegion(0, 0, bounds.Value.Width, bounds.Value.Height, remember: true);
                return;
            }

            var rect = new PixelRect(left, top, width, height);

            EnsureRegionOverlay();
            if (_regionOverlay == null) return;
            _regionOverlay.SetScreenBounds(bounds.Value);
            _regionOverlay.SetRegion(rect);
            _regionOverlay.Show();
        }

        private void EnsureRegionOverlay()
        {
            if (_regionOverlay != null) return;
            if (!_isOpened) return;
            _regionOverlay = new RegionOverlayWindow();
            _regionOverlay.Show(this);
        }

        private static int ParseInt(string value)
        {
            return int.TryParse(value, out var result) ? result : 0;
        }

        private void UpdateCurrentDisplayId()
        {
            if (_vm == null) return;
            if (!OperatingSystem.IsMacOS()) return;
            var id = GetDisplayIdForWindow();
            if (id > 0)
            {
                _vm.UpdateCurrentDisplayId(id);
            }
        }

        private int GetDisplayIdForWindow()
        {
            var handle = this.TryGetPlatformHandle();
            if (handle == null || handle.Handle == IntPtr.Zero) return 0;
            var nsWindow = handle.Handle;
            var selScreen = sel_registerName("screen");
            var screen = objc_msgSend(nsWindow, selScreen);
            if (screen == IntPtr.Zero) return 0;
            var selDeviceDesc = sel_registerName("deviceDescription");
            var deviceDesc = objc_msgSend(screen, selDeviceDesc);
            if (deviceDesc == IntPtr.Zero) return 0;

            var nsString = objc_getClass("NSString");
            var selStringWithUTF8 = sel_registerName("stringWithUTF8String:");
            var key = objc_msgSend_str(nsString, selStringWithUTF8, "NSScreenNumber");
            var selObjectForKey = sel_registerName("objectForKey:");
            var number = objc_msgSend_ptr(deviceDesc, selObjectForKey, key);
            if (number == IntPtr.Zero) return 0;
            var selUnsignedIntValue = sel_registerName("unsignedIntValue");
            return (int)objc_msgSend_uint(number, selUnsignedIntValue);
        }

        [DllImport("/usr/lib/libobjc.A.dylib")]
        private static extern IntPtr objc_getClass(string name);

        [DllImport("/usr/lib/libobjc.A.dylib")]
        private static extern IntPtr sel_registerName(string name);

        [DllImport("/usr/lib/libobjc.A.dylib")]
        private static extern IntPtr objc_msgSend(IntPtr receiver, IntPtr selector);

        [DllImport("/usr/lib/libobjc.A.dylib", EntryPoint = "objc_msgSend")]
        private static extern IntPtr objc_msgSend_str(IntPtr receiver, IntPtr selector, string arg1);

        [DllImport("/usr/lib/libobjc.A.dylib", EntryPoint = "objc_msgSend")]
        private static extern IntPtr objc_msgSend_ptr(IntPtr receiver, IntPtr selector, IntPtr arg1);

        [DllImport("/usr/lib/libobjc.A.dylib", EntryPoint = "objc_msgSend")]
        private static extern uint objc_msgSend_uint(IntPtr receiver, IntPtr selector);
    }
}
