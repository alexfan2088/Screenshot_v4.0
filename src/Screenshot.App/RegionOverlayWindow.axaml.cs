using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Markup.Xaml;
using Avalonia.Platform;
using System.Runtime.InteropServices;

namespace Screenshot.App
{
    public partial class RegionOverlayWindow : Window
    {
        private const int MinRegionSize = 20;
        private const int GripSize = 8;
        private const int CornerGripSize = 14;
        private PixelRect _currentRegion;
        private bool _editable;
        private bool _isDragging;
        private Point _dragStartClient;
        private PixelRect _dragStartRegion;
        private PixelPoint _dragStartWindowPos;
        private int _dragStartWidthDip;
        private int _dragStartHeightDip;
        private DragMode _dragMode = DragMode.None;

        public event Action<PixelRect>? RegionChanged;

        public RegionOverlayWindow()
        {
            InitializeComponent();
            Focusable = false;
            Opened += (_, _) =>
            {
                HookPointerEvents();
                ApplyPointerMode();
            };
            SizeChanged += (_, _) => RefreshLayout();
        }

        private void InitializeComponent()
        {
            AvaloniaXamlLoader.Load(this);
        }

        public void SetScreenBounds(PixelRect bounds, double scale)
        {
            Position = bounds.Position;
            var safeScale = scale <= 0 ? 1.0 : scale;
            Width = bounds.Width / safeScale;
            Height = bounds.Height / safeScale;
            SetRegion(new PixelRect(0, 0, Math.Max(1, (int)Math.Round(Width)), Math.Max(1, (int)Math.Round(Height))));
        }

        public void SetRegion(PixelRect region)
        {
            var border = this.FindControl<Border>("HighlightBorder");
            if (border == null) return;

            _currentRegion = region;
            Canvas.SetLeft(border, region.X);
            Canvas.SetTop(border, region.Y);
            border.Width = region.Width;
            border.Height = region.Height;
            RefreshLayout();
        }

        public void SetEditable(bool editable)
        {
            if (_editable == editable) return;
            _editable = editable;
            ApplyPointerMode();
        }

        private void HookPointerEvents()
        {
            HookGrip("TopGrip", DragMode.ResizeTop);
            HookGrip("BottomGrip", DragMode.ResizeBottom);
            HookGrip("LeftGrip", DragMode.ResizeLeft);
            HookGrip("RightGrip", DragMode.ResizeRight);
            HookGrip("BottomRightGrip", DragMode.ResizeBottomRight);

            PointerMoved += OnOverlayPointerMoved;
            PointerReleased += OnOverlayPointerReleased;
        }

        private void HookGrip(string name, DragMode mode)
        {
            var grip = this.FindControl<Border>(name);
            if (grip == null) return;
            grip.PointerPressed += (sender, e) =>
            {
                if (!_editable) return;
                if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed) return;
                if (sender is not Control control) return;
                StartDrag(control, e, mode);
            };
        }

        private void StartDrag(Control control, PointerEventArgs e, DragMode mode)
        {
            _dragMode = mode;
            _isDragging = true;
            _dragStartClient = e.GetPosition(this);
            _dragStartRegion = _currentRegion;
            _dragStartWindowPos = Position;
            _dragStartWidthDip = Math.Max(1, (int)Math.Round(Bounds.Width));
            _dragStartHeightDip = Math.Max(1, (int)Math.Round(Bounds.Height));
            e.Pointer.Capture(control);
            e.Handled = true;
        }

        private void OnOverlayPointerMoved(object? sender, PointerEventArgs e)
        {
            if (!_editable || !_isDragging) return;

            var current = e.GetPosition(this);
            var dx = (int)Math.Round(current.X - _dragStartClient.X);
            var dy = (int)Math.Round(current.Y - _dragStartClient.Y);
            var newX = _dragStartWindowPos.X;
            var newY = _dragStartWindowPos.Y;
            var newWidth = _dragStartWidthDip;
            var newHeight = _dragStartHeightDip;
            switch (_dragMode)
            {
                case DragMode.ResizeTop:
                    newY = _dragStartWindowPos.Y + dy;
                    newHeight = _dragStartHeightDip - dy;
                    break;
                case DragMode.ResizeBottom:
                    newHeight = _dragStartHeightDip + dy;
                    break;
                case DragMode.ResizeLeft:
                    newX = _dragStartWindowPos.X + dx;
                    newWidth = _dragStartWidthDip - dx;
                    break;
                case DragMode.ResizeRight:
                    newWidth = _dragStartWidthDip + dx;
                    break;
                case DragMode.ResizeBottomRight:
                    newWidth = _dragStartWidthDip + dx;
                    newHeight = _dragStartHeightDip + dy;
                    break;
            }

            if (newWidth < MinRegionSize)
            {
                if (_dragMode == DragMode.ResizeLeft)
                {
                    newX -= (MinRegionSize - newWidth);
                }
                newWidth = MinRegionSize;
            }
            if (newHeight < MinRegionSize)
            {
                if (_dragMode == DragMode.ResizeTop)
                {
                    newY -= (MinRegionSize - newHeight);
                }
                newHeight = MinRegionSize;
            }

            Position = new PixelPoint(newX, newY);
            Width = newWidth;
            Height = newHeight;
            SetRegion(new PixelRect(0, 0, Math.Max(1, newWidth), Math.Max(1, newHeight)));
            RegionChanged?.Invoke(_currentRegion);
            e.Handled = true;
        }

        private void OnOverlayPointerReleased(object? sender, PointerReleasedEventArgs e)
        {
            if (!_isDragging) return;
            _isDragging = false;
            _dragMode = DragMode.None;
            e.Pointer.Capture(null);
            RegionChanged?.Invoke(_currentRegion);
            e.Handled = true;
        }

        private void RefreshLayout()
        {
            var w = Math.Max(1, (int)Math.Round(Bounds.Width));
            var h = Math.Max(1, (int)Math.Round(Bounds.Height));

            PositionGrip("TopGrip", 0, 0, w, GripSize);
            PositionGrip("BottomGrip", 0, Math.Max(0, h - GripSize), w, GripSize);
            PositionGrip("LeftGrip", 0, 0, GripSize, h);
            PositionGrip("RightGrip", Math.Max(0, w - GripSize), 0, GripSize, h);
            PositionGrip("BottomRightGrip", Math.Max(0, w - CornerGripSize), Math.Max(0, h - CornerGripSize), CornerGripSize, CornerGripSize);
        }

        private void PositionGrip(string name, double x, double y, double width, double height)
        {
            var grip = this.FindControl<Border>(name);
            if (grip == null) return;
            Canvas.SetLeft(grip, x);
            Canvas.SetTop(grip, y);
            grip.Width = width;
            grip.Height = height;
        }

        private void ApplyPointerMode()
        {
            IsHitTestVisible = _editable;
            UpdateMousePassthrough(!_editable);
        }

        private void UpdateMousePassthrough(bool passThrough)
        {
            if (!OperatingSystem.IsMacOS()) return;
            var handle = this.TryGetPlatformHandle();
            if (handle == null || handle.Handle == IntPtr.Zero) return;
            var sel = sel_registerName("setIgnoresMouseEvents:");
            objc_msgSend(handle.Handle, sel, passThrough);
        }

        private enum DragMode
        {
            None = 0,
            ResizeTop = 1,
            ResizeBottom = 2,
            ResizeLeft = 3,
            ResizeRight = 4,
            ResizeBottomRight = 5
        }

        [DllImport("/usr/lib/libobjc.A.dylib")]
        private static extern IntPtr sel_registerName(string name);

        [DllImport("/usr/lib/libobjc.A.dylib")]
        private static extern void objc_msgSend(IntPtr receiver, IntPtr selector, bool value);
    }
}
