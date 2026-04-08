using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Runtime.InteropServices;

namespace ClipSidekick;

internal sealed class SettingsDialog
{
    private const int WIDTH = 380;
    private const int HEIGHT = 224;
    private const int ROW_HEIGHT = 44;
    private const int LABEL_X = 20;
    private const int CONTROL_RIGHT = 360;
    private const int HEADER_HEIGHT = 48;

    private IntPtr _hWnd;
    private NativeMethods.WndProc _wndProc = null!;
    private bool _done;
    private AppSettings _settings = null!;
    private bool _inHotkeyDialog;
    private IntPtr _mouseHook;
    private NativeMethods.LowLevelMouseProc _mouseHookProc = null!;

    // Editable values
    private int _notificationDurationMs;
    private string _hotkey = "";
    private int _maxHistoryItems;

    // UI state
    private int _hoverRow = -1; // 0=duration, 1=hotkey, 2=max items
    private int _hoverButton = -1; // for duration: 0=minus, 1=plus; for hotkey: 0=change
    private bool _hoverClose;

    private static readonly Color BgColor = Color.FromArgb(32, 32, 32);
    private static readonly Color TextColor = Color.FromArgb(230, 230, 230);
    private static readonly Color DimText = Color.FromArgb(140, 140, 140);
    private static readonly Color AccentColor = Color.FromArgb(96, 165, 250);
    private static readonly Color BorderColor = Color.FromArgb(60, 60, 60);
    private static readonly Color RowHoverBg = Color.FromArgb(42, 42, 42);
    private static readonly Color ButtonBg = Color.FromArgb(55, 55, 55);
    private static readonly Color ButtonHoverBg = Color.FromArgb(70, 70, 70);
    private static readonly Color SeparatorColor = Color.FromArgb(50, 50, 50);

    [DllImport("user32.dll")]
    private static extern bool SetLayeredWindowAttributes(IntPtr hwnd, uint crKey, byte bAlpha, uint dwFlags);

    public bool Show(IntPtr parent, AppSettings settings)
    {
        _settings = settings;
        _notificationDurationMs = settings.NotificationDurationMs;
        _hotkey = settings.Hotkey;
        _maxHistoryItems = settings.MaxHistoryItems;
        _done = false;
        _inHotkeyDialog = false;

        var hInstance = NativeMethods.GetModuleHandleW(IntPtr.Zero);
        _wndProc = WndProc;

        var className = "ClipSidekickSettings";
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

        NativeMethods.GetCursorPos(out var cursor);
        var monitor = NativeMethods.MonitorFromPoint(cursor, NativeMethods.MONITOR_DEFAULTTONEAREST);
        var mi = new NativeMethods.MONITORINFO { cbSize = Marshal.SizeOf<NativeMethods.MONITORINFO>() };
        NativeMethods.GetMonitorInfoW(monitor, ref mi);

        int x = Math.Clamp(cursor.X - WIDTH / 2, mi.rcWork.Left, mi.rcWork.Right - WIDTH);
        int y = Math.Clamp(cursor.Y - HEIGHT / 2, mi.rcWork.Top, mi.rcWork.Bottom - HEIGHT);

        _hWnd = NativeMethods.CreateWindowExW(
            NativeMethods.WS_EX_TOOLWINDOW | NativeMethods.WS_EX_TOPMOST | NativeMethods.WS_EX_LAYERED,
            className, "Settings",
            NativeMethods.WS_POPUP,
            x, y, WIDTH, HEIGHT,
            IntPtr.Zero, IntPtr.Zero, hInstance, IntPtr.Zero);

        int pref = NativeMethods.DWMWCP_ROUND;
        NativeMethods.DwmSetWindowAttribute(_hWnd, NativeMethods.DWMWA_WINDOW_CORNER_PREFERENCE, ref pref, sizeof(int));
        SetLayeredWindowAttributes(_hWnd, 0, 245, 0x02);

        NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_SHOW);
        NativeMethods.SetForegroundWindow(_hWnd);
        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        InstallMouseHook();

        while (!_done && NativeMethods.GetMessageW(out var msg, IntPtr.Zero, 0, 0))
        {
            NativeMethods.TranslateMessage(ref msg);
            NativeMethods.DispatchMessageW(ref msg);
        }

        RemoveMouseHook();
        NativeMethods.DestroyWindow(_hWnd);
        Marshal.FreeHGlobal(classNamePtr);

        return true;
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

            case NativeMethods.WM_MOUSEMOVE:
                OnMouseMove(lParam);
                return IntPtr.Zero;

            case NativeMethods.WM_LBUTTONDOWN:
                OnMouseDown(lParam);
                return IntPtr.Zero;

            case NativeMethods.WM_APP_DISMISS:
                _done = true;
                return IntPtr.Zero;

            case NativeMethods.WM_KEYDOWN:
                if (wParam.ToInt32() == NativeMethods.VK_ESCAPE)
                    _done = true;
                return IntPtr.Zero;

            case NativeMethods.WM_ACTIVATE:
                if (NativeMethods.LOWORD(wParam) == NativeMethods.WA_INACTIVE && !_inHotkeyDialog)
                    _done = true;
                return IntPtr.Zero;

            case NativeMethods.WM_DESTROY:
                _done = true;
                return IntPtr.Zero;
        }

        return NativeMethods.DefWindowProcW(hWnd, msg, wParam, lParam);
    }

    private void OnMouseMove(IntPtr lParam)
    {
        int mx = NativeMethods.GET_X_LPARAM(lParam);
        int my = NativeMethods.GET_Y_LPARAM(lParam);

        int newRow = -1;
        int newButton = -1;
        bool newHoverClose = false;

        // Close button
        if (mx >= WIDTH - 40 && mx < WIDTH - 8 && my >= 8 && my < 40)
            newHoverClose = true;

        int rowY = HEADER_HEIGHT;
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            newRow = 0; // Duration row
            // Minus button: right-aligned before value
            int minusBtnX = CONTROL_RIGHT - 130;
            int plusBtnX = CONTROL_RIGHT - 30;
            if (mx >= minusBtnX && mx < minusBtnX + 28 && my >= rowY + 8 && my < rowY + 36)
                newButton = 0;
            else if (mx >= plusBtnX && mx < plusBtnX + 28 && my >= rowY + 8 && my < rowY + 36)
                newButton = 1;
        }

        rowY += ROW_HEIGHT;
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            newRow = 1; // Hotkey row
            int changeBtnX = CONTROL_RIGHT - 80;
            if (mx >= changeBtnX && mx < changeBtnX + 70 && my >= rowY + 8 && my < rowY + 36)
                newButton = 0;
        }

        rowY += ROW_HEIGHT;
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            newRow = 2; // Max history items row
            int minusBtnX = CONTROL_RIGHT - 130;
            int plusBtnX = CONTROL_RIGHT - 30;
            if (mx >= minusBtnX && mx < minusBtnX + 28 && my >= rowY + 8 && my < rowY + 36)
                newButton = 0;
            else if (mx >= plusBtnX && mx < plusBtnX + 28 && my >= rowY + 8 && my < rowY + 36)
                newButton = 1;
        }

        if (newRow != _hoverRow || newButton != _hoverButton || newHoverClose != _hoverClose)
        {
            _hoverRow = newRow;
            _hoverButton = newButton;
            _hoverClose = newHoverClose;
            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        }
    }

    private void OnMouseDown(IntPtr lParam)
    {
        int mx = NativeMethods.GET_X_LPARAM(lParam);
        int my = NativeMethods.GET_Y_LPARAM(lParam);

        // Close button
        if (mx >= WIDTH - 40 && mx < WIDTH - 8 && my >= 8 && my < 40)
        {
            _done = true;
            return;
        }

        // Duration row
        int rowY = HEADER_HEIGHT;
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            int minusBtnX = CONTROL_RIGHT - 130;
            int plusBtnX = CONTROL_RIGHT - 30;
            if (mx >= minusBtnX && mx < minusBtnX + 28 && my >= rowY + 8 && my < rowY + 36)
            {
                _notificationDurationMs = Math.Max(500, _notificationDurationMs - 500);
                SaveSettings();
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            }
            else if (mx >= plusBtnX && mx < plusBtnX + 28 && my >= rowY + 8 && my < rowY + 36)
            {
                _notificationDurationMs = Math.Min(10000, _notificationDurationMs + 500);
                SaveSettings();
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            }
        }

        // Hotkey row
        rowY += ROW_HEIGHT;
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            int changeBtnX = CONTROL_RIGHT - 80;
            if (mx >= changeBtnX && mx < changeBtnX + 70 && my >= rowY + 8 && my < rowY + 36)
            {
                _inHotkeyDialog = true;
                RemoveMouseHook();
                NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_HIDE);
                var dialog = new HotkeyDialog();
                var result = dialog.Show(_hWnd, _hotkey);
                if (result != null)
                {
                    _hotkey = result;
                    SaveSettings();
                }
                NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_SHOW);
                NativeMethods.SetForegroundWindow(_hWnd);
                _inHotkeyDialog = false;
                InstallMouseHook();
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            }
        }

        // Max history items row
        rowY += ROW_HEIGHT;
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            int minusBtnX = CONTROL_RIGHT - 130;
            int plusBtnX = CONTROL_RIGHT - 30;
            if (mx >= minusBtnX && mx < minusBtnX + 28 && my >= rowY + 8 && my < rowY + 36)
            {
                _maxHistoryItems = Math.Max(10, _maxHistoryItems - 10);
                SaveSettings();
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            }
            else if (mx >= plusBtnX && mx < plusBtnX + 28 && my >= rowY + 8 && my < rowY + 36)
            {
                _maxHistoryItems = Math.Min(200, _maxHistoryItems + 10);
                SaveSettings();
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            }
        }

    }

    private void SaveSettings()
    {
        _settings.NotificationDurationMs = _notificationDurationMs;
        _settings.Hotkey = _hotkey;
        _settings.MaxHistoryItems = _maxHistoryItems;
        _settings.Save();
    }

    private void InstallMouseHook()
    {
        RemoveMouseHook();
        _mouseHookProc = MouseHookCallback;
        _mouseHook = NativeMethods.SetWindowsHookExW(
            NativeMethods.WH_MOUSE_LL, _mouseHookProc,
            NativeMethods.GetModuleHandleW(IntPtr.Zero), 0);
    }

    private void RemoveMouseHook()
    {
        if (_mouseHook != IntPtr.Zero)
        {
            NativeMethods.UnhookWindowsHookEx(_mouseHook);
            _mouseHook = IntPtr.Zero;
        }
    }

    private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && wParam.ToInt32() == NativeMethods.WM_LBUTTONDOWN && !_inHotkeyDialog)
        {
            var info = Marshal.PtrToStructure<NativeMethods.MSLLHOOKSTRUCT>(lParam);
            NativeMethods.GetWindowRect(_hWnd, out var rect);
            if (info.pt.X < rect.Left || info.pt.X > rect.Right ||
                info.pt.Y < rect.Top || info.pt.Y > rect.Bottom)
            {
                NativeMethods.PostMessageW(_hWnd, NativeMethods.WM_APP_DISMISS, IntPtr.Zero, IntPtr.Zero);
            }
        }
        return NativeMethods.CallNextHookEx(_mouseHook, nCode, wParam, lParam);
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

        // Header
        using var titleFont = new Font("Segoe UI", 12f, FontStyle.Regular);
        using var titleBrush = new SolidBrush(TextColor);
        g.DrawString("Settings", titleFont, titleBrush, 20, 14);

        // Close button
        int closeBtnX = w - 36;
        int closeBtnY = 10;
        if (_hoverClose)
        {
            using var hoverBrush = new SolidBrush(ButtonHoverBg);
            var closeRect = new Rectangle(closeBtnX - 2, closeBtnY - 2, 28, 28);
            using var closePath = CreateRoundedRect(closeRect, 4);
            g.FillPath(hoverBrush, closePath);
        }
        using var closeFont = new Font("Segoe UI", 10f);
        using var closeBrush = new SolidBrush(DimText);
        g.DrawString("\u00D7", closeFont, closeBrush, closeBtnX, closeBtnY);

        using var sepPen = new Pen(SeparatorColor);
        g.DrawLine(sepPen, 0, HEADER_HEIGHT - 1, w, HEADER_HEIGHT - 1);

        using var labelFont = new Font("Segoe UI", 9.5f);
        using var labelBrush = new SolidBrush(TextColor);
        using var dimBrush = new SolidBrush(DimText);
        using var valueFont = new Font("Segoe UI Semibold", 10f);

        // --- Row 0: Notification Duration ---
        int rowY = HEADER_HEIGHT;
        if (_hoverRow == 0)
        {
            using var rowBrush = new SolidBrush(RowHoverBg);
            g.FillRectangle(rowBrush, 1, rowY, w - 2, ROW_HEIGHT);
        }

        g.DrawString("Notification duration", labelFont, labelBrush, LABEL_X, rowY + 8);
        g.DrawString("How long the popup stays visible", new Font("Segoe UI", 7.5f), dimBrush, LABEL_X, rowY + 26);

        // Minus button
        int minusBtnX = CONTROL_RIGHT - 130;
        DrawSmallButton(g, minusBtnX, rowY + 8, 28, 28, "−", _hoverRow == 0 && _hoverButton == 0);

        // Value
        var durationText = $"{_notificationDurationMs / 1000.0:0.#}s";
        var durationSize = g.MeasureString(durationText, valueFont);
        float valueX = minusBtnX + 28 + (72 - durationSize.Width) / 2;
        g.DrawString(durationText, valueFont, new SolidBrush(AccentColor), valueX, rowY + 12);

        // Plus button
        int plusBtnX = CONTROL_RIGHT - 30;
        DrawSmallButton(g, plusBtnX, rowY + 8, 28, 28, "+", _hoverRow == 0 && _hoverButton == 1);

        g.DrawLine(sepPen, 20, rowY + ROW_HEIGHT - 1, w - 20, rowY + ROW_HEIGHT - 1);

        // --- Row 1: Hotkey ---
        rowY += ROW_HEIGHT;
        if (_hoverRow == 1)
        {
            using var rowBrush = new SolidBrush(RowHoverBg);
            g.FillRectangle(rowBrush, 1, rowY, w - 2, ROW_HEIGHT);
        }

        g.DrawString("Hotkey", labelFont, labelBrush, LABEL_X, rowY + 8);
        g.DrawString(_hotkey, new Font("Segoe UI", 7.5f), dimBrush, LABEL_X, rowY + 26);

        // Change button
        int changeBtnX = CONTROL_RIGHT - 80;
        DrawPillButton(g, changeBtnX, rowY + 10, 70, 24, "Change", _hoverRow == 1 && _hoverButton == 0);

        g.DrawLine(sepPen, 20, rowY + ROW_HEIGHT - 1, w - 20, rowY + ROW_HEIGHT - 1);

        // --- Row 2: Max History Items ---
        rowY += ROW_HEIGHT;
        if (_hoverRow == 2)
        {
            using var rowBrush = new SolidBrush(RowHoverBg);
            g.FillRectangle(rowBrush, 1, rowY, w - 2, ROW_HEIGHT);
        }

        g.DrawString("Max history items", labelFont, labelBrush, LABEL_X, rowY + 8);
        g.DrawString("Number of items to keep", new Font("Segoe UI", 7.5f), dimBrush, LABEL_X, rowY + 26);

        // Minus button
        int itemsMinusBtnX = CONTROL_RIGHT - 130;
        DrawSmallButton(g, itemsMinusBtnX, rowY + 8, 28, 28, "\u2212", _hoverRow == 2 && _hoverButton == 0);

        // Value
        var itemsText = _maxHistoryItems.ToString();
        var itemsSize = g.MeasureString(itemsText, valueFont);
        float itemsValueX = itemsMinusBtnX + 28 + (72 - itemsSize.Width) / 2;
        g.DrawString(itemsText, valueFont, new SolidBrush(AccentColor), itemsValueX, rowY + 12);

        // Plus button
        int itemsPlusBtnX = CONTROL_RIGHT - 30;
        DrawSmallButton(g, itemsPlusBtnX, rowY + 8, 28, 28, "+", _hoverRow == 2 && _hoverButton == 1);

        // Blit
        using var screenG = Graphics.FromHdc(ps.hdc);
        screenG.DrawImageUnscaled(bmp, 0, 0);

        NativeMethods.EndPaint(hWnd, ref ps);
    }

    private void DrawSmallButton(Graphics g, int x, int y, int w, int h, string text, bool hover)
    {
        var rect = new Rectangle(x, y, w, h);
        using var path = CreateRoundedRect(rect, 4);
        using (var brush = new SolidBrush(hover ? ButtonHoverBg : ButtonBg))
            g.FillPath(brush, path);

        using var font = new Font("Segoe UI", 11f, FontStyle.Bold);
        var size = g.MeasureString(text, font);
        using var textBrush = new SolidBrush(TextColor);
        g.DrawString(text, font, textBrush,
            x + (w - size.Width) / 2,
            y + (h - size.Height) / 2);
    }

    private void DrawPillButton(Graphics g, int x, int y, int w, int h, string text, bool hover)
    {
        var rect = new Rectangle(x, y, w, h);
        using var path = CreateRoundedRect(rect, h / 2);
        using (var brush = new SolidBrush(hover ? ButtonHoverBg : ButtonBg))
            g.FillPath(brush, path);

        using var font = new Font("Segoe UI", 8.5f);
        var size = g.MeasureString(text, font);
        using var textBrush = new SolidBrush(TextColor);
        g.DrawString(text, font, textBrush,
            x + (w - size.Width) / 2,
            y + (h - size.Height) / 2);
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
