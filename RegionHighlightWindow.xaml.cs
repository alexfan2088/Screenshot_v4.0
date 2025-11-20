using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;

namespace Screenshot_v3_0
{
    /// <summary>
    /// 显示红色边框的高亮窗口
    /// </summary>
    public partial class RegionHighlightWindow : Window
    {
        private const int GwlExstyle = -20;
        private const int WsExTransparent = 0x00000020;
        private const int WsExToolwindow = 0x00000080;

        public RegionHighlightWindow()
        {
            InitializeComponent();
            ShowInTaskbar = false;
            IsHitTestVisible = false;
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);

            var handle = new WindowInteropHelper(this).Handle;
            int style = GetWindowLong(handle, GwlExstyle);
            style |= WsExTransparent | WsExToolwindow;
            SetWindowLong(handle, GwlExstyle, style);
        }

        public void ShowRegion(Int32Rect rect)
        {
            if (rect.Width <= 0 || rect.Height <= 0)
            {
                Hide();
                return;
            }

            // rect 是物理像素坐标（屏幕坐标）
            // 使用 Win32 API 直接设置窗口位置，避免 DPI 转换问题
            if (!IsVisible)
            {
                Show();
            }

            // 确保窗口句柄已创建
            var helper = new WindowInteropHelper(this);
            if (helper.Handle == IntPtr.Zero)
            {
                // 如果句柄还未创建，先显示窗口
                Show();
                helper.EnsureHandle();
            }

            // 使用 SetWindowPos 直接设置窗口位置和大小（物理像素）
            SetWindowPos(
                helper.Handle,
                IntPtr.Zero, // HWND_TOP
                rect.X,
                rect.Y,
                rect.Width,
                rect.Height,
                0x0040); // SWP_NOZORDER | SWP_NOACTIVATE
        }

        public void HideRegion()
        {
            try
            {
                if (IsVisible)
                {
                    Hide();
                }
            }
            catch
            {
                // 忽略隐藏错误
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            try
            {
                Hide();
            }
            catch
            {
                // 忽略关闭错误
            }
            base.OnClosed(e);
        }

        [DllImport("user32.dll")]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    }
}

