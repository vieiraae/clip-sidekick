using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Text;

namespace ClipSidekick;

/// <summary>
/// A small dialog that captures a keyboard shortcut from the user.
/// Shows "Press your desired shortcut..." and captures modifier+key combinations.
/// </summary>
internal sealed class HotkeyDialog
{
    private const int WIDTH = 340;
    private const int HEIGHT = 150;
    private const int WM_KEYDOWN = 0x0100;
    private const int WM_SYSKEYDOWN = 0x0104;

    private IntPtr _hWnd;
    private NativeMethods.WndProc _wndProc = null!;
    private string? _result;
    private string _currentDisplay = "";
    private string _currentHotkey;
    private bool _done;
    private bool _captured;

    // Colors
    private static readonly Color BgColor = Color.FromArgb(32, 32, 32);
    private static readonly Color TextColor = Color.FromArgb(230, 230, 230);
    private static readonly Color DimText = Color.FromArgb(140, 140, 140);
    private static readonly Color AccentColor = Color.FromArgb(96, 165, 250);
    private static readonly Color BorderColor = Color.FromArgb(60, 60, 60);
    private static readonly Color KeyBg = Color.FromArgb(55, 55, 55);

    [DllImport("user32.dll")]
    private static extern bool SetLayeredWindowAttributes(IntPtr hwnd, uint crKey, byte bAlpha, uint dwFlags);

    public HotkeyDialog()
    {
        _currentHotkey = "";
    }

    public string? Show(IntPtr parent, string currentHotkey)
    {
        _currentHotkey = currentHotkey;
        _currentDisplay = currentHotkey;
        _result = null;
        _done = false;
        _captured = false;

        var hInstance = NativeMethods.GetModuleHandleW(IntPtr.Zero);
        _wndProc = WndProc;

        var className = "ClipSidekickHotkeyDialog";
        var classNamePtr = Marshal.StringToHGlobalUni(className);

        var wc = new NativeMethods.WNDCLASSEX
        {
            cbSize = Marshal.SizeOf<NativeMethods.WNDCLASSEX>(),
            style = 0,
            lpfnWndProc = Marshal.GetFunctionPointerForDelegate(_wndProc),
            hInstance = hInstance,
            hCursor = NativeMethods.LoadCursorW(IntPtr.Zero, NativeMethods.IDC_ARROW),
            hbrBackground = IntPtr.Zero,
            lpszClassName = classNamePtr
        };

        NativeMethods.RegisterClassExW(ref wc);

        // Center on screen near cursor
        NativeMethods.GetCursorPos(out var cursor);
        var monitor = NativeMethods.MonitorFromPoint(cursor, NativeMethods.MONITOR_DEFAULTTONEAREST);
        var mi = new NativeMethods.MONITORINFO { cbSize = Marshal.SizeOf<NativeMethods.MONITORINFO>() };
        NativeMethods.GetMonitorInfoW(monitor, ref mi);

        int x = cursor.X - WIDTH / 2;
        int y = cursor.Y - HEIGHT / 2;
        x = Math.Clamp(x, mi.rcWork.Left, mi.rcWork.Right - WIDTH);
        y = Math.Clamp(y, mi.rcWork.Top, mi.rcWork.Bottom - HEIGHT);

        _hWnd = NativeMethods.CreateWindowExW(
            NativeMethods.WS_EX_TOOLWINDOW | NativeMethods.WS_EX_TOPMOST | NativeMethods.WS_EX_LAYERED,
            className,
            "Set Hotkey",
            NativeMethods.WS_POPUP,
            x, y, WIDTH, HEIGHT,
            IntPtr.Zero, IntPtr.Zero, hInstance, IntPtr.Zero);

        int pref = NativeMethods.DWMWCP_ROUND;
        NativeMethods.DwmSetWindowAttribute(_hWnd, NativeMethods.DWMWA_WINDOW_CORNER_PREFERENCE, ref pref, sizeof(int));
        SetLayeredWindowAttributes(_hWnd, 0, 240, 0x02);

        NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_SHOW);
        NativeMethods.SetForegroundWindow(_hWnd);
        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);

        // Modal message loop
        while (!_done && NativeMethods.GetMessageW(out var msg, IntPtr.Zero, 0, 0))
        {
            NativeMethods.TranslateMessage(ref msg);
            NativeMethods.DispatchMessageW(ref msg);
        }

        NativeMethods.DestroyWindow(_hWnd);
        Marshal.FreeHGlobal(classNamePtr);

        return _result;
    }

    private IntPtr WndProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam)
    {
        switch (msg)
        {
            case NativeMethods.WM_PAINT:
                OnPaint(hWnd);
                return IntPtr.Zero;

            case NativeMethods.WM_ERASEBKGND:
                return new IntPtr(1);

            case WM_KEYDOWN:
            case WM_SYSKEYDOWN:
                OnKeyDown(wParam.ToInt32());
                return IntPtr.Zero;

            case NativeMethods.WM_ACTIVATE:
                if (NativeMethods.LOWORD(wParam) == NativeMethods.WA_INACTIVE)
                {
                    // Cancel on deactivation
                    _done = true;
                }
                return IntPtr.Zero;

            case NativeMethods.WM_TIMER:
                _done = true;
                return IntPtr.Zero;

            case NativeMethods.WM_DESTROY:
                _done = true;
                return IntPtr.Zero;
        }

        return NativeMethods.DefWindowProcW(hWnd, msg, wParam, lParam);
    }

    private void OnKeyDown(int vk)
    {
        // Escape cancels
        if (vk == NativeMethods.VK_ESCAPE)
        {
            _done = true;
            return;
        }

        // Ignore standalone modifier presses
        if (vk is 0x10 or 0x11 or 0x12 or 0x5B or 0x5C or 0xA0 or 0xA1 or 0xA2 or 0xA3 or 0xA4 or 0xA5)
        {
            // Update display with current modifiers
            _currentDisplay = BuildModifierString() + "...";
            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            return;
        }

        // Build the hotkey string
        var sb = new StringBuilder();
        var mods = BuildModifierString();

        if (string.IsNullOrEmpty(mods))
        {
            // Must have at least one modifier
            _currentDisplay = "Need at least one modifier key";
            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            return;
        }

        sb.Append(mods);
        sb.Append(GetKeyName(vk));

        _result = sb.ToString();
        _currentDisplay = _result;
        _captured = true;
        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);

        // Auto-close after a brief moment to show what was captured
        NativeMethods.SetTimer(_hWnd, 1, 400, IntPtr.Zero);
    }

    private static string BuildModifierString()
    {
        var sb = new StringBuilder();

        if ((NativeMethods.GetAsyncKeyState(0x5B) & 0x8000) != 0 ||
            (NativeMethods.GetAsyncKeyState(0x5C) & 0x8000) != 0)
            sb.Append("Win+");

        if ((NativeMethods.GetAsyncKeyState(0x11) & 0x8000) != 0) // Ctrl
            sb.Append("Ctrl+");

        if ((NativeMethods.GetAsyncKeyState(0x12) & 0x8000) != 0) // Alt
            sb.Append("Alt+");

        if ((NativeMethods.GetAsyncKeyState(0x10) & 0x8000) != 0) // Shift
            sb.Append("Shift+");

        return sb.ToString();
    }

    private static string GetKeyName(int vk)
    {
        if (vk is >= 0x41 and <= 0x5A)
            return ((char)vk).ToString();
        if (vk is >= 0x30 and <= 0x39)
            return ((char)vk).ToString();
        if (vk is >= 0x70 and <= 0x7B)
            return $"F{vk - 0x70 + 1}";

        return vk switch
        {
            0x20 => "Space",
            0x09 => "Tab",
            0x0D => "Enter",
            0x2D => "Insert",
            0x2E => "Delete",
            0x24 => "Home",
            0x23 => "End",
            0x21 => "PageUp",
            0x22 => "PageDown",
            0x25 => "Left",
            0x26 => "Up",
            0x27 => "Right",
            0x28 => "Down",
            0xC0 => "`",
            0xBD => "-",
            0xBB => "=",
            0xDB => "[",
            0xDD => "]",
            0xDC => "\\",
            0xBA => ";",
            0xDE => "'",
            0xBC => ",",
            0xBE => ".",
            0xBF => "/",
            _ => $"Key({vk:X2})"
        };
    }

    private void OnPaint(IntPtr hWnd)
    {
        NativeMethods.BeginPaint(hWnd, out var ps);
        NativeMethods.GetClientRect(hWnd, out var cr);
        int w = cr.Right - cr.Left;
        int h = cr.Bottom - cr.Top;

        if (w <= 0 || h <= 0)
        {
            NativeMethods.EndPaint(hWnd, ref ps);
            return;
        }

        using var bmp = new Bitmap(w, h);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

        // Background
        using (var bgBrush = new SolidBrush(BgColor))
            g.FillRectangle(bgBrush, 0, 0, w, h);
        using (var borderPen = new Pen(BorderColor))
            g.DrawRectangle(borderPen, 0, 0, w - 1, h - 1);

        // Title
        using var titleFont = new Font("Segoe UI", 11f, FontStyle.Regular);
        using var titleBrush = new SolidBrush(TextColor);
        g.DrawString("Set Hotkey", titleFont, titleBrush, 16, 12);

        // Instruction
        using var instrFont = new Font("Segoe UI", 9f);
        using var dimBrush = new SolidBrush(DimText);
        g.DrawString("Press your desired shortcut combination", instrFont, dimBrush, 16, 38);

        // Current hotkey display (centered, in a rounded box)
        var keyText = _currentDisplay;
        using var keyFont = new Font("Segoe UI Semibold", 13f);
        var keySize = g.MeasureString(keyText, keyFont);
        int boxW = Math.Max((int)keySize.Width + 32, 120);
        int boxH = 38;
        int boxX = (w - boxW) / 2;
        int boxY = 68;

        var boxRect = new Rectangle(boxX, boxY, boxW, boxH);
        using var keyBgBrush = new SolidBrush(KeyBg);
        using var keyPath = CreateRoundedRect(boxRect, 6);
        g.FillPath(keyBgBrush, keyPath);

        if (_captured)
        {
            using var accentPen = new Pen(AccentColor, 1.5f);
            g.DrawPath(accentPen, keyPath);
        }
        else
        {
            using var boxBorderPen = new Pen(BorderColor);
            g.DrawPath(boxBorderPen, keyPath);
        }

        using var keyBrush = new SolidBrush(_captured ? AccentColor : TextColor);
        float textX = boxX + (boxW - keySize.Width) / 2;
        float textY = boxY + (boxH - keySize.Height) / 2;
        g.DrawString(keyText, keyFont, keyBrush, textX, textY);

        // Hint at bottom
        using var hintFont = new Font("Segoe UI", 8f);
        g.DrawString("Esc to cancel", hintFont, dimBrush, 16, h - 22);

        // Blit
        using var screenG = Graphics.FromHdc(ps.hdc);
        screenG.DrawImageUnscaled(bmp, 0, 0);

        NativeMethods.EndPaint(hWnd, ref ps);
    }

    private static GraphicsPath CreateRoundedRect(Rectangle bounds, int radius)
    {
        var path = new GraphicsPath();
        int d = radius * 2;
        path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
        path.AddArc(bounds.Right - d - 1, bounds.Y, d, d, 270, 90);
        path.AddArc(bounds.Right - d - 1, bounds.Bottom - d - 1, d, d, 0, 90);
        path.AddArc(bounds.X, bounds.Bottom - d - 1, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }
}
