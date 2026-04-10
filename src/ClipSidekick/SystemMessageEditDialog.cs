using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Runtime.InteropServices;

namespace ClipSidekick;

internal sealed class SystemMessageEditDialog
{
    private const int WIDTH = 420;
    private const int HEIGHT = 340;
    private const int FIELD_LEFT = 20;
    private const int FIELD_RIGHT = 20;
    private const int LABEL_Y = 56;
    private const int FIELD_Y = 82;
    private const int FIELD_H = 180;

    private IntPtr _hWnd;
    private IntPtr _hEdit;
    private IntPtr _hFont;
    private IntPtr _bgBrush;
    private NativeMethods.WndProc _wndProc = null!;
    private bool _done;
    private string? _result;
    private int _hoverButton = -1; // 0=save, 1=cancel, 2=reset

    private static readonly Color BgColor = Color.FromArgb(32, 32, 32);
    private static readonly Color TextColor = Color.FromArgb(230, 230, 230);
    private static readonly Color AccentColor = Color.FromArgb(96, 165, 250);
    private static readonly Color DimText = Color.FromArgb(140, 140, 140);
    private static readonly Color FieldBg = Color.FromArgb(45, 45, 45);
    private static readonly Color ButtonBg = Color.FromArgb(55, 55, 55);
    private static readonly Color ButtonHoverBg = Color.FromArgb(70, 70, 70);

    private const int WM_CTLCOLOREDIT = 0x0133;
    private const int WM_SETFONT = 0x0030;
    private const int ES_MULTILINE = 0x0004;
    private const int ES_AUTOVSCROLL = 0x0040;
    private const int ES_WANTRETURN = 0x1000;
    private const int WS_CHILD = 0x40000000;
    private const int WS_VISIBLE = 0x10000000;
    private const int WS_VSCROLL = 0x00200000;
    private const int WS_TABSTOP = 0x00010000;
    private const int WS_EX_CLIENTEDGE = 0x00000200;
    private const int WM_GETTEXT = 0x000D;
    private const int WM_GETTEXTLENGTH = 0x000E;

    [DllImport("user32.dll")]
    private static extern bool SetLayeredWindowAttributes(IntPtr hwnd, uint crKey, byte bAlpha, uint dwFlags);

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateSolidBrush(int crColor);

    [DllImport("gdi32.dll")]
    private static extern int SetTextColor(IntPtr hdc, int crColor);

    [DllImport("gdi32.dll")]
    private static extern int SetBkColor(IntPtr hdc, int crColor);

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateFontW(int cHeight, int cWidth, int cEscapement, int cOrientation,
        int cWeight, uint bItalic, uint bUnderline, uint bStrikeOut, uint iCharSet,
        uint iOutPrecision, uint iClipPrecision, uint iQuality, uint iPitchAndFamily,
        [MarshalAs(UnmanagedType.LPWStr)] string pszFaceName);

    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DeleteObject(IntPtr hObject);

    private static int ToCOLORREF(Color c) => c.R | (c.G << 8) | (c.B << 16);

    public string? Show(IntPtr parent, string message)
    {
        _done = false;
        _result = null;

        var hInstance = NativeMethods.GetModuleHandleW(IntPtr.Zero);
        _wndProc = WndProc;

        _hFont = CreateFontW(-18, 0, 0, 0, 400, 0, 0, 0, 0, 0, 0, 5, 0, "Segoe UI");
        _bgBrush = CreateSolidBrush(ToCOLORREF(FieldBg));

        var className = "ClipSidekickSysMsgEdit";
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
            className, "System Message",
            NativeMethods.WS_POPUP | NativeMethods.WS_CLIPCHILDREN,
            x, y, WIDTH, HEIGHT,
            IntPtr.Zero, IntPtr.Zero, hInstance, IntPtr.Zero);

        int pref = NativeMethods.DWMWCP_ROUND;
        NativeMethods.DwmSetWindowAttribute(_hWnd, NativeMethods.DWMWA_WINDOW_CORNER_PREFERENCE, ref pref, sizeof(int));
        SetLayeredWindowAttributes(_hWnd, 0, 245, 0x02);

        int fieldW = WIDTH - FIELD_LEFT - FIELD_RIGHT;

        _hEdit = NativeMethods.CreateWindowExW(
            WS_EX_CLIENTEDGE,
            "EDIT", message,
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_WANTRETURN,
            FIELD_LEFT, FIELD_Y, fieldW, FIELD_H,
            _hWnd, IntPtr.Zero, hInstance, IntPtr.Zero);

        NativeMethods.SendMessageW(_hEdit, WM_SETFONT, _hFont, new IntPtr(1));
        NativeMethods.SetFocus(_hEdit);

        NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_SHOW);
        NativeMethods.SetForegroundWindow(_hWnd);
        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);

        while (!_done && NativeMethods.GetMessageW(out var msg, IntPtr.Zero, 0, 0))
        {
            NativeMethods.TranslateMessage(ref msg);
            NativeMethods.DispatchMessageW(ref msg);
        }

        if (!_done)
            NativeMethods.PostQuitMessage(0);

        NativeMethods.DestroyWindow(_hWnd);
        NativeMethods.UnregisterClassW(classNamePtr, hInstance);
        Marshal.FreeHGlobal(classNamePtr);
        DeleteObject(_hFont);
        DeleteObject(_bgBrush);

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

            case NativeMethods.WM_LBUTTONDOWN:
                OnMouseDown(lParam);
                return IntPtr.Zero;

            case NativeMethods.WM_MOUSEMOVE:
                OnMouseMove(lParam);
                return IntPtr.Zero;

            case WM_CTLCOLOREDIT:
                SetTextColor(wParam, ToCOLORREF(TextColor));
                SetBkColor(wParam, ToCOLORREF(FieldBg));
                return _bgBrush;

            case NativeMethods.WM_KEYDOWN:
                if (wParam.ToInt32() == NativeMethods.VK_ESCAPE)
                    _done = true;
                return IntPtr.Zero;

            case NativeMethods.WM_DESTROY:
                _done = true;
                return IntPtr.Zero;
        }

        return NativeMethods.DefWindowProcW(hWnd, msg, wParam, lParam);
    }

    private string GetEditText(IntPtr hEdit)
    {
        int len = (int)NativeMethods.SendMessageW(hEdit, WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero);
        if (len <= 0) return "";
        var buf = new char[len + 1];
        var ptr = Marshal.AllocHGlobal((len + 1) * 2);
        NativeMethods.SendMessageW(hEdit, WM_GETTEXT, new IntPtr(len + 1), ptr);
        Marshal.Copy(ptr, buf, 0, len + 1);
        Marshal.FreeHGlobal(ptr);
        return new string(buf, 0, len);
    }

    private void TrySave()
    {
        _result = GetEditText(_hEdit).Trim();
        _done = true;
    }

    private void ResetToDefault()
    {
        _result = new AppSettings().SystemMessage;
        _done = true;
    }

    private void OnMouseDown(IntPtr lParam)
    {
        int mx = NativeMethods.GET_X_LPARAM(lParam);
        int my = NativeMethods.GET_Y_LPARAM(lParam);

        if (my >= HEIGHT - 40 && my < HEIGHT - 14)
        {
            if (mx >= WIDTH - 160 && mx < WIDTH - 90) TrySave();
            else if (mx >= WIDTH - 80 && mx < WIDTH - 20) _done = true;
        }
        if (mx >= FIELD_LEFT && mx < FIELD_LEFT + 80 && my >= HEIGHT - 40 && my < HEIGHT - 14)
            ResetToDefault();

        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
    }

    private void OnMouseMove(IntPtr lParam)
    {
        int mx = NativeMethods.GET_X_LPARAM(lParam);
        int my = NativeMethods.GET_Y_LPARAM(lParam);

        int newHover = -1;
        if (my >= HEIGHT - 40 && my < HEIGHT - 14)
        {
            if (mx >= WIDTH - 160 && mx < WIDTH - 90) newHover = 0;
            else if (mx >= WIDTH - 80 && mx < WIDTH - 20) newHover = 1;
            else if (mx >= FIELD_LEFT && mx < FIELD_LEFT + 80) newHover = 2;
        }

        if (newHover != _hoverButton)
        {
            _hoverButton = newHover;
            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        }
    }

    private void OnPaint(IntPtr hWnd)
    {
        NativeMethods.BeginPaint(hWnd, out var ps);
        NativeMethods.GetClientRect(hWnd, out var clientRect);
        int w = clientRect.Right;
        int h = clientRect.Bottom;

        if (w <= 0 || h <= 0) { NativeMethods.EndPaint(hWnd, ref ps); return; }

        using var bmp = new Bitmap(w, h);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

        using (var bgBrush = new SolidBrush(BgColor))
            g.FillRectangle(bgBrush, 0, 0, w, h);
        using (var borderPen = new Pen(Color.FromArgb(60, 60, 60)))
            g.DrawRectangle(borderPen, 0, 0, w - 1, h - 1);

        using var titleFont = new Font("Segoe UI", 11f);
        using var labelFont = new Font("Segoe UI", 9.5f);
        using var labelBrush = new SolidBrush(TextColor);

        g.DrawString("System Message", titleFont, labelBrush, 20, 16);
        g.DrawString("Instructions for the AI assistant", labelFont, new SolidBrush(DimText), 20, LABEL_Y);

        // Buttons
        DrawButton(g, WIDTH - 160, h - 40, 64, 26, "Save", _hoverButton == 0, true);
        DrawButton(g, WIDTH - 80, h - 40, 54, 26, "Cancel", _hoverButton == 1, false);
        DrawButton(g, FIELD_LEFT, h - 40, 74, 26, "Reset", _hoverButton == 2, false);

        using var screenG = Graphics.FromHdc(ps.hdc);
        screenG.DrawImageUnscaled(bmp, 0, 0);
        NativeMethods.EndPaint(hWnd, ref ps);
    }

    private void DrawButton(Graphics g, int x, int y, int w, int h, string text, bool hover, bool accent)
    {
        var rect = new Rectangle(x, y, w, h);
        using var path = CreateRoundedRect(rect, 4);
        var bg = accent && !hover ? AccentColor : (hover ? ButtonHoverBg : ButtonBg);
        using (var bgBrush = new SolidBrush(bg))
            g.FillPath(bgBrush, path);
        using var font = new Font("Segoe UI", 8.5f);
        var textSize = g.MeasureString(text, font);
        using var brush = new SolidBrush(accent ? Color.White : TextColor);
        g.DrawString(text, font, brush, x + (w - textSize.Width) / 2, y + (h - textSize.Height) / 2);
    }

    private static GraphicsPath CreateRoundedRect(Rectangle rect, int radius)
    {
        var path = new GraphicsPath();
        int d = radius * 2;
        path.AddArc(rect.X, rect.Y, d, d, 180, 90);
        path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
        path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
        path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }
}
