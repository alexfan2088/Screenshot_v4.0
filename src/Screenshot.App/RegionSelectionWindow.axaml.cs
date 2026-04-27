using System;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Markup.Xaml;

namespace Screenshot.App
{
    public partial class RegionSelectionWindow : Window
    {
        private Point? _start;
        private PixelRect _currentRect;
        private bool _hasRect;
        private readonly TaskCompletionSource<PixelRect?> _tcs = new();

        public RegionSelectionWindow()
        {
            InitializeComponent();
            this.PointerPressed += OnPointerPressed;
            this.PointerMoved += OnPointerMoved;
            this.PointerReleased += OnPointerReleased;
            this.KeyDown += OnKeyDown;
            this.Closed += OnClosed;
            Cursor = new Cursor(StandardCursorType.Cross);
        }

        private void InitializeComponent()
        {
            AvaloniaXamlLoader.Load(this);
        }

        public Task<PixelRect?> PickAsync()
        {
            return _tcs.Task;
        }

        private void OnPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            var point = e.GetPosition(this);
            _start = point;
            UpdateSelection(_start.Value, _start.Value);
        }

        private void OnPointerMoved(object? sender, PointerEventArgs e)
        {
            if (_start == null) return;
            var current = e.GetPosition(this);
            UpdateSelection(_start.Value, current);
        }

        private void OnPointerReleased(object? sender, PointerReleasedEventArgs e)
        {
            if (_start == null) return;

            var end = e.GetPosition(this);
            var rect = ClampRect(BuildRect(_start.Value, end));

            _currentRect = rect;
            _hasRect = rect.Width > 0 && rect.Height > 0;
            _start = null;

            _tcs.TrySetResult(_hasRect ? rect : null);
            Close();
        }

        private void OnKeyDown(object? sender, KeyEventArgs e)
        {
            // No Enter confirm. Esc cancels selection.
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

        private void UpdateSelection(Point a, Point b)
        {
            var rect = BuildRect(a, b);
            _currentRect = ClampRect(rect);
            _hasRect = _currentRect.Width > 0 && _currentRect.Height > 0;
            var border = this.FindControl<Border>("SelectionBorder");
            if (border == null) return;

            Canvas.SetLeft(border, _currentRect.X);
            Canvas.SetTop(border, _currentRect.Y);
            border.Width = _currentRect.Width;
            border.Height = _currentRect.Height;
            border.IsVisible = true;
        }

        private PixelRect BuildRect(Point a, Point b)
        {
            var x1 = (int)Math.Round(Math.Min(a.X, b.X));
            var y1 = (int)Math.Round(Math.Min(a.Y, b.Y));
            var x2 = (int)Math.Round(Math.Max(a.X, b.X));
            var y2 = (int)Math.Round(Math.Max(a.Y, b.Y));
            return new PixelRect(x1, y1, Math.Max(0, x2 - x1), Math.Max(0, y2 - y1));
        }

        private PixelRect ClampRect(PixelRect rect)
        {
            var maxX = (int)Math.Max(0, Bounds.Width - 1);
            var maxY = (int)Math.Max(0, Bounds.Height - 1);
            var x = Math.Clamp(rect.X, 0, maxX);
            var y = Math.Clamp(rect.Y, 0, maxY);
            var width = Math.Clamp(rect.Width, 1, maxX - x + 1);
            var height = Math.Clamp(rect.Height, 1, maxY - y + 1);
            return new PixelRect(x, y, width, height);
        }
    }
}
