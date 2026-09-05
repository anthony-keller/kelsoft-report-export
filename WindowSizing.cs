using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using Microsoft.Win32;

namespace KelsoftReportExport;

/// <summary>
/// Keeps the window inside the work area of whichever monitor it is on.
///
/// Two problems share that one answer. With the standard frame removed, Windows maximises a
/// window to the whole monitor — overhanging the edges and covering the taskbar. And the
/// design size is larger than a scaled display has room for: at 150% on a 1080p screen the
/// desktop is only about 690 device-independent pixels tall, so a window asking for 880
/// opens with its lower half, including the Export button, off the bottom of the screen.
///
/// Everything here works in physical pixels, which is the one measure that stays true
/// whatever the display scaling is; WPF's own units do not.
/// </summary>
internal sealed class WindowSizing
{
    private const int WmGetMinMaxInfo = 0x0024;
    private const int MonitorDefaultToNearest = 0x0002;
    private const uint SwpNoZOrder = 0x0004;
    private const uint SwpNoActivate = 0x0010;

    /// <summary>Desktop left showing on each side when the window has to be cut down.</summary>
    private const int Inset = 16;

    private readonly Window _window;

    private WindowSizing(Window window) => _window = window;

    public static void Apply(Window window) => new WindowSizing(window).Attach();

    private void Attach()
    {
        _window.SourceInitialized += (_, _) =>
        {
            HwndSource.FromHwnd(Handle)?.AddHook(Hook);
            FitToWorkArea();
        };

        // Dragged to a monitor with different scaling, the window is resized to match it —
        // and may no longer fit.
        _window.DpiChanged += (_, _) => FitToWorkArea();

        // So might a resolution change, a monitor being unplugged, or a remote session
        // reconnecting at a different size.
        void OnDisplaySettingsChanged(object? sender, EventArgs e) =>
            _window.Dispatcher.BeginInvoke(FitToWorkArea);

        SystemEvents.DisplaySettingsChanged += OnDisplaySettingsChanged;
        _window.Closed += (_, _) => SystemEvents.DisplaySettingsChanged -= OnDisplaySettingsChanged;
    }

    private IntPtr Handle => new WindowInteropHelper(_window).Handle;

    /// <summary>
    /// Shrinks and re-centres the window if it is wider or taller than the monitor can show.
    /// A window that already fits is left exactly where the user put it.
    /// </summary>
    private void FitToWorkArea()
    {
        if (_window.WindowState != WindowState.Normal) return;

        var handle = Handle;
        if (handle == IntPtr.Zero) return;
        if (!TryGetWorkArea(handle, out var work)) return;
        if (!GetWindowRect(handle, out var bounds)) return;

        var available = new Size(
            work.Right - work.Left - (2 * Inset),
            work.Bottom - work.Top - (2 * Inset));

        var width = bounds.Right - bounds.Left;
        var height = bounds.Bottom - bounds.Top;
        if (width <= available.Width && height <= available.Height) return;

        width = Math.Min(width, available.Width);
        height = Math.Min(height, available.Height);

        SetWindowPos(handle, IntPtr.Zero,
            work.Left + ((work.Right - work.Left - width) / 2),
            work.Top + ((work.Bottom - work.Top - height) / 2),
            width, height, SwpNoZOrder | SwpNoActivate);
    }

    /// <summary>
    /// Answers WM_GETMINMAXINFO with the work area, so maximising stops at the taskbar. The
    /// window's own minimum is passed on in the same breath, because handling the message at
    /// all takes it out of WPF's hands — and it is held to the work area too, so the minimum
    /// can never be the reason a window will not fit on the screen.
    /// </summary>
    private IntPtr Hook(IntPtr hwnd, int message, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (message != WmGetMinMaxInfo) return IntPtr.Zero;
        if (!TryGetWorkArea(hwnd, out var work, out var monitor)) return IntPtr.Zero;

        var bounds = Marshal.PtrToStructure<MinMaxInfo>(lParam);

        // Positions are relative to the monitor, not the desktop.
        bounds.MaxPosition = new Point(work.Left - monitor.Left, work.Top - monitor.Top);
        bounds.MaxSize = new Point(work.Right - work.Left, work.Bottom - work.Top);
        bounds.MaxTrackSize = bounds.MaxSize;

        var dpi = GetDpiForWindow(hwnd);
        var scale = dpi > 0 ? dpi / 96.0 : 1.0;
        bounds.MinTrackSize = new Point(
            (int)Math.Min(_window.MinWidth * scale, bounds.MaxSize.X),
            (int)Math.Min(_window.MinHeight * scale, bounds.MaxSize.Y));

        Marshal.StructureToPtr(bounds, lParam, true);
        handled = true;
        return IntPtr.Zero;
    }

    private static bool TryGetWorkArea(IntPtr hwnd, out Rect work) =>
        TryGetWorkArea(hwnd, out work, out _);

    private static bool TryGetWorkArea(IntPtr hwnd, out Rect work, out Rect monitor)
    {
        work = default;
        monitor = default;

        var handle = MonitorFromWindow(hwnd, MonitorDefaultToNearest);
        if (handle == IntPtr.Zero) return false;

        var info = new MonitorInfo { Size = Marshal.SizeOf<MonitorInfo>() };
        if (!GetMonitorInfo(handle, ref info)) return false;

        work = info.Work;
        monitor = info.Monitor;
        return true;
    }

    [DllImport("user32.dll")]
    private static extern IntPtr MonitorFromWindow(IntPtr handle, int flags);

    [DllImport("user32.dll")]
    private static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo info);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr handle, out Rect bounds);

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(IntPtr handle, IntPtr insertAfter,
        int x, int y, int width, int height, uint flags);

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr handle);

    private readonly record struct Size(int Width, int Height);

    [StructLayout(LayoutKind.Sequential)]
    private struct Point(int x, int y)
    {
        public int X = x;
        public int Y = y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Rect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MonitorInfo
    {
        public int Size;
        public Rect Monitor;
        public Rect Work;
        public int Flags;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MinMaxInfo
    {
        public Point Reserved;
        public Point MaxSize;
        public Point MaxPosition;
        public Point MinTrackSize;
        public Point MaxTrackSize;
    }
}
