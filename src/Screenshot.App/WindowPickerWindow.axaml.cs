using System;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Markup.Xaml;
using Screenshot.App.Services;

namespace Screenshot.App
{
    public partial class WindowPickerWindow : Window
    {
        private readonly TaskCompletionSource<WindowSelection?> _tcs = new();
        private int _currentWindowId;
        private PixelRect _currentRect;

        public WindowPickerWindow()
        {
            InitializeComponent();
            this.PointerMoved += OnPointerMoved;
            this.PointerPressed += OnPointerPressed;
            this.KeyDown += OnKeyDown;
            this.Closed += OnClosed;
            Cursor = new Cursor(StandardCursorType.Cross);
        }

        private void InitializeComponent()
        {
            AvaloniaXamlLoader.Load(this);
        }

        public Task<WindowSelection?> PickAsync()
        {
            return _tcs.Task;
        }

        private void OnPointerMoved(object? sender, PointerEventArgs e)
        {
            var screenPoint = this.PointToScreen(e.GetPosition(this));
            if (!MacWindowPicker.TryGetWindowAtPoint(screenPoint.X, screenPoint.Y, out var window) || window == null)
            {
                HideSelection();
                return;
            }

            if (window.WindowId != _currentWindowId || !windowBoundsEquals(window, _currentRect))
            {
                _currentWindowId = window.WindowId;
                _currentRect = new PixelRect(window.X, window.Y, window.Width, window.Height);
                UpdateSelection(_currentRect);
            }
        }

        private void OnPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (_currentWindowId <= 0) return;
            _tcs.TrySetResult(new WindowSelection(_currentWindowId, _currentRect));
            Close();
        }

        private void OnKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                _tcs.TrySetResult(null);
                Close();
            }
        }

        private void OnClosed(object? sender, EventArgs e)
        {
            _tcs.TrySetResult(null);
        }

        private void UpdateSelection(PixelRect rect)
        {
            var border = this.FindControl<Border>("SelectionBorder");
            if (border == null) return;
            var topLeft = this.PointToClient(new PixelPoint(rect.X, rect.Y));
            Canvas.SetLeft(border, topLeft.X);
            Canvas.SetTop(border, topLeft.Y);
            border.Width = rect.Width;
            border.Height = rect.Height;
            border.IsVisible = rect.Width > 0 && rect.Height > 0;
        }

        private void HideSelection()
        {
            var border = this.FindControl<Border>("SelectionBorder");
            if (border == null) return;
            border.IsVisible = false;
            _currentWindowId = 0;
            _currentRect = default;
        }

        private static bool windowBoundsEquals(MacWindowInfo window, PixelRect rect)
        {
            return window.X == rect.X && window.Y == rect.Y && window.Width == rect.Width && window.Height == rect.Height;
        }
    }

    public readonly record struct WindowSelection(int WindowId, PixelRect Bounds);
}
