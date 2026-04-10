using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Text;

namespace ClipSidekick;

/// <summary>
/// A small, modern toast popup that appears near the cursor when a clipboard item is captured.
/// Clicking it opens the main history window.
/// </summary>
internal sealed class NotificationPopup
{
    private const int POPUP_WIDTH = 280;
    private const int POPUP_HEIGHT = 36;
    private const int CORNER_RADIUS = 10;
    private const int BORDER_WIDTH = 2;
    private const nuint TIMER_ID = 100;
    private const nuint ANIM_TIMER_ID = 101;

    private float _dpiScale = 1f;

    private int ScaledWidth => (int)(POPUP_WIDTH * _dpiScale);
    private int ScaledHeight => (int)(POPUP_HEIGHT * _dpiScale);

    private IntPtr _hWnd;
    private IntPtr _hInstance;
    private NativeMethods.WndProc _wndProc = null!;
    private string _previewText = string.Empty;
    private bool _isImage;
    private bool _hovered;
    private bool _aiHovered;
    private bool _arrowHovered;
    private bool _dropdownOpen;
    private int _dropdownHoverIndex = -1;
    private bool _tracking;
    private bool _processing;
    private float _animPhase;

    // Freeform input mode
    private bool _inputMode;
    private string _inputText = "";
    private IntPtr _keyboardHook;
    private NativeMethods.LowLevelKeyboardProc _keyboardHookProc = null!;
    private IntPtr _savedKeyboardLayout;

    public int DurationMs { get; set; } = 1000;

    public event Action? Clicked;
    public event Action? AIEditClicked;
    public event Action<int>? QuickTaskClicked;
    public event Action? CancelRequested;
    public event Action<string>? FreeformSubmitted;

    public string[] TaskHotkeys { get; set; } = new string[10];
    public List<CustomTask> CustomTasks { get; set; } = [];
    public string FreeformHotkey { get; set; } = "";

    private int TotalTaskCount => TaskLabels.Length + CustomTasks.Count + 1; // +1 for freeform

    // Task labels (must match MainWindow.EditTasks, skip index 0)
    private static readonly string[] TaskLabels = ["Proofread", "Rewrite", "Use synonyms", "Minor revise", "Major revise", "Describe", "Answer", "Explain", "Expand", "Summarize"];
    private static readonly string[] TaskIcons = ["\ud83d\udc41", "\u270f\ufe0f", "\ud83d\udcda", "\ud83d\udd27", "\ud83d\udd28", "\ud83d\udcdd", "\ud83d\udcac", "\ud83d\udca1", "\ud83c\udf1f", "\ud83d\udccb"];

    // Colors
    private static readonly Color BgColor = Color.FromArgb(30, 30, 30);
    private static readonly Color BgHoverColor = Color.FromArgb(40, 40, 40);
    private static readonly Color TextColor = Color.FromArgb(220, 220, 220);

    [DllImport("user32.dll")]
    private static extern bool SetLayeredWindowAttributes(IntPtr hwnd, uint crKey, byte bAlpha, uint dwFlags);

    public void Create(IntPtr hInstance)
    {
        _hInstance = hInstance;
        _wndProc = WndProc;

        var className = "ClipSidekickNotification";
        var classNamePtr = Marshal.StringToHGlobalUni(className);

        var wc = new NativeMethods.WNDCLASSEX
        {
            cbSize = Marshal.SizeOf<NativeMethods.WNDCLASSEX>(),
            style = 0,
            lpfnWndProc = Marshal.GetFunctionPointerForDelegate(_wndProc),
            hInstance = _hInstance,
            hCursor = NativeMethods.LoadCursorW(IntPtr.Zero, NativeMethods.IDC_HAND),
            hbrBackground = IntPtr.Zero,
            lpszClassName = classNamePtr
        };

        NativeMethods.RegisterClassExW(ref wc);

        _hWnd = NativeMethods.CreateWindowExW(
            NativeMethods.WS_EX_TOOLWINDOW | NativeMethods.WS_EX_TOPMOST |
            NativeMethods.WS_EX_NOACTIVATE | NativeMethods.WS_EX_LAYERED,
            className,
            "",
            NativeMethods.WS_POPUP,
            0, 0, POPUP_WIDTH, POPUP_HEIGHT,
            IntPtr.Zero, IntPtr.Zero, _hInstance, IntPtr.Zero);

        // Rounded corners on Win11
        int pref = NativeMethods.DWMWCP_ROUND;
        NativeMethods.DwmSetWindowAttribute(_hWnd, NativeMethods.DWMWA_WINDOW_CORNER_PREFERENCE, ref pref, sizeof(int));

        // Fully opaque layered window (same as main window)
        SetLayeredWindowAttributes(_hWnd, 0, 255, 0x02);

        Marshal.FreeHGlobal(classNamePtr);
    }

    public void Show(ClipboardItem item)
    {
        _isImage = item.Type == ClipboardItemType.Image;
        _previewText = _isImage
            ? $"Image ({item.Image?.Width}\u00d7{item.Image?.Height})"
            : TruncateText(item.Text, 45);
        _hovered = false;

        // Position near cursor
        NativeMethods.GetCursorPos(out var cursor);
        var monitor = NativeMethods.MonitorFromPoint(cursor, NativeMethods.MONITOR_DEFAULTTONEAREST);
        var mi = new NativeMethods.MONITORINFO { cbSize = Marshal.SizeOf<NativeMethods.MONITORINFO>() };
        NativeMethods.GetMonitorInfoW(monitor, ref mi);

        // Detect monitor DPI for scaling
        if (NativeMethods.GetDpiForMonitor(monitor, NativeMethods.MDT_EFFECTIVE_DPI, out uint dpiX, out _) == 0)
            _dpiScale = dpiX / 96f;
        else
            _dpiScale = 1f;

        int pw = ScaledWidth;
        int ph = ScaledHeight;

        var workArea = mi.rcWork;
        int x = cursor.X - pw / 2;
        int y = cursor.Y - ph - 12; // Above cursor

        if (y < workArea.Top)
            y = cursor.Y + 20;

        x = Math.Clamp(x, workArea.Left + 4, workArea.Right - pw - 4);
        y = Math.Clamp(y, workArea.Top + 4, workArea.Bottom - ph - 4);

        NativeMethods.SetWindowPos(_hWnd, NativeMethods.HWND_TOPMOST, x, y, pw, ph,
            NativeMethods.SWP_NOACTIVATE);

        NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_SHOWNOACTIVATE);
        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);

        // Auto-dismiss timer
        NativeMethods.SetTimer(_hWnd, TIMER_ID, (uint)DurationMs, IntPtr.Zero);
    }

    public void Dismiss()
    {
        NativeMethods.KillTimer(_hWnd, TIMER_ID);
        NativeMethods.KillTimer(_hWnd, ANIM_TIMER_ID);
        _dropdownOpen = false;
        _dropdownHoverIndex = -1;
        _processing = false;
        if (_inputMode)
            ExitInputMode();
        NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_HIDE);
        ResizePopup(ScaledWidth, ScaledHeight);
    }

    public void UpdatePreviewText(string text)
    {
        _previewText = TruncateText(text, 45);
        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
    }

    public void StartProcessing()
    {
        _processing = true;
        _animPhase = 0;
        _dropdownOpen = false;
        _dropdownHoverIndex = -1;
        NativeMethods.KillTimer(_hWnd, TIMER_ID);
        ResizePopup(ScaledWidth, ScaledHeight);
        NativeMethods.SetTimer(_hWnd, ANIM_TIMER_ID, 30, IntPtr.Zero);
        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        InstallKeyboardHook();
    }

    public void StopProcessing()
    {
        _processing = false;
        NativeMethods.KillTimer(_hWnd, ANIM_TIMER_ID);
        RemoveKeyboardHook();
        Dismiss();
    }

    /// <summary>
    /// Shows the notification in freeform input mode (used by the global hotkey path).
    /// </summary>
    public void ShowInputMode()
    {
        // Position near cursor (same as Show)
        NativeMethods.GetCursorPos(out var cursor);
        var monitor = NativeMethods.MonitorFromPoint(cursor, NativeMethods.MONITOR_DEFAULTTONEAREST);
        var mi = new NativeMethods.MONITORINFO { cbSize = Marshal.SizeOf<NativeMethods.MONITORINFO>() };
        NativeMethods.GetMonitorInfoW(monitor, ref mi);

        if (NativeMethods.GetDpiForMonitor(monitor, NativeMethods.MDT_EFFECTIVE_DPI, out uint dpiX, out _) == 0)
            _dpiScale = dpiX / 96f;
        else
            _dpiScale = 1f;

        int pw = ScaledWidth;
        int ph = ScaledHeight;

        var workArea = mi.rcWork;
        int x = cursor.X - pw / 2;
        int y = cursor.Y - ph - 12;

        if (y < workArea.Top)
            y = cursor.Y + 20;

        x = Math.Clamp(x, workArea.Left + 4, workArea.Right - pw - 4);
        y = Math.Clamp(y, workArea.Top + 4, workArea.Bottom - ph - 4);

        NativeMethods.SetWindowPos(_hWnd, NativeMethods.HWND_TOPMOST, x, y, pw, ph,
            NativeMethods.SWP_NOACTIVATE);

        CaptureKeyboardLayout();
        _inputMode = true;
        _inputText = "";
        _dropdownOpen = false;
        _dropdownHoverIndex = -1;
        _processing = false;
        _hovered = false;

        NativeMethods.KillTimer(_hWnd, TIMER_ID);
        NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_SHOWNOACTIVATE);
        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        InstallKeyboardHook();
    }

    public void EnterInputMode()
    {
        CaptureKeyboardLayout();
        _inputMode = true;
        _inputText = "";
        _dropdownOpen = false;
        _dropdownHoverIndex = -1;
        NativeMethods.KillTimer(_hWnd, TIMER_ID);
        ResizePopup(ScaledWidth, ScaledHeight);
        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        InstallKeyboardHook();
    }

    private void CaptureKeyboardLayout()
    {
        var fgWnd = NativeMethods.GetForegroundWindow();
        uint tid = NativeMethods.GetWindowThreadProcessId(fgWnd, out _);
        _savedKeyboardLayout = NativeMethods.GetKeyboardLayout(tid);
    }

    private void ExitInputMode()
    {
        _inputMode = false;
        _inputText = "";
        RemoveKeyboardHook();
    }

    private void InstallKeyboardHook()
    {
        RemoveKeyboardHook();
        _keyboardHookProc = KeyboardHookCallback;
        _keyboardHook = NativeMethods.SetWindowsHookExW(
            NativeMethods.WH_KEYBOARD_LL, _keyboardHookProc, _hInstance, 0);
    }

    private void RemoveKeyboardHook()
    {
        if (_keyboardHook != IntPtr.Zero)
        {
            NativeMethods.UnhookWindowsHookEx(_keyboardHook);
            _keyboardHook = IntPtr.Zero;
        }
    }

    private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && _processing && !_inputMode)
        {
            int wmMsg = wParam.ToInt32();
            if (wmMsg == NativeMethods.WM_KEYDOWN || wmMsg == 0x0104)
            {
                var hookStruct = Marshal.PtrToStructure<NativeMethods.KBDLLHOOKSTRUCT>(lParam);
                if (hookStruct.vkCode == 0x1B) // VK_ESCAPE
                {
                    CancelRequested?.Invoke();
                    return new IntPtr(1);
                }
            }
            return NativeMethods.CallNextHookEx(_keyboardHook, nCode, wParam, lParam);
        }

        if (nCode >= 0 && _inputMode)
        {
            int wmMsg = wParam.ToInt32();
            var hookStruct = Marshal.PtrToStructure<NativeMethods.KBDLLHOOKSTRUCT>(lParam);
            uint vk = hookStruct.vkCode;

            // Let modifier keys pass through so the OS tracks key state correctly
            if (vk is 0xA0 or 0xA1 or 0xA2 or 0xA3 or 0xA4 or 0xA5 // L/R Shift, Ctrl, Alt
                     or 0x10 or 0x11 or 0x12  // VK_SHIFT, VK_CONTROL, VK_MENU
                     or 0x14 or 0x90 or 0x91  // CapsLock, NumLock, ScrollLock
                     or 0x5B or 0x5C)         // Win keys
                return NativeMethods.CallNextHookEx(_keyboardHook, nCode, wParam, lParam);

            // Only handle key-down events for the rest
            if (wmMsg != NativeMethods.WM_KEYDOWN && wmMsg != 0x0104) // WM_KEYDOWN or WM_SYSKEYDOWN
                return new IntPtr(1);

            if (vk == 0x1B) // VK_ESCAPE
            {
                ExitInputMode();
                Dismiss();
                return new IntPtr(1);
            }

            if (vk == 0x0D) // VK_RETURN
            {
                if (_inputText.Length > 0)
                {
                    var instruction = _inputText;
                    ExitInputMode();
                    StartProcessing();
                    FreeformSubmitted?.Invoke(instruction);
                }
                else
                {
                    ExitInputMode();
                    Dismiss();
                }
                return new IntPtr(1);
            }

            if (vk == 0x08) // VK_BACK
            {
                if (_inputText.Length > 0)
                {
                    _inputText = _inputText[..^1];
                    NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                }
                return new IntPtr(1);
            }

            // Ctrl+V — paste from clipboard
            if (vk == 0x56 && NativeMethods.KBDLLHOOKSTRUCT_IsCtrl(hookStruct))
            {
                PasteFromClipboard();
                return new IntPtr(1);
            }

            // Ctrl+A — swallow
            if (vk == 0x41 && NativeMethods.KBDLLHOOKSTRUCT_IsCtrl(hookStruct))
                return new IntPtr(1);

            // Convert keypress to character
            // Build key state from physical keyboard (GetKeyboardState is per-thread and stale here)
            byte[] keyState = new byte[256];
            foreach (int mk in new[] { 0x10, 0xA0, 0xA1,   // VK_SHIFT, VK_LSHIFT, VK_RSHIFT
                                        0x11, 0xA2, 0xA3,   // VK_CONTROL, VK_LCONTROL, VK_RCONTROL
                                        0x12, 0xA4, 0xA5 }) // VK_MENU, VK_LMENU, VK_RMENU
            {
                if ((NativeMethods.GetAsyncKeyState(mk) & 0x8000) != 0)
                    keyState[mk] = 0x80;
            }
            // Toggle keys: low bit of GetAsyncKeyState = toggled on
            if ((NativeMethods.GetAsyncKeyState(0x14) & 1) != 0) keyState[0x14] = 1; // CapsLock
            if ((NativeMethods.GetAsyncKeyState(0x90) & 1) != 0) keyState[0x90] = 1; // NumLock

            var sb = new StringBuilder(4);
            var fgWnd = NativeMethods.GetForegroundWindow();
            uint tid = NativeMethods.GetWindowThreadProcessId(fgWnd, out _);
            IntPtr layout = NativeMethods.GetKeyboardLayout(tid);
            if (layout == IntPtr.Zero) layout = _savedKeyboardLayout;
            int result = NativeMethods.ToUnicodeEx(vk, hookStruct.scanCode, keyState, sb, sb.Capacity, 0, layout);
            if (result > 0 && _inputText.Length < 200)
            {
                string chars = sb.ToString(0, result);
                if (chars.Length > 0 && chars[0] >= 32)
                {
                    _inputText += chars;
                    NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                }
            }

            return new IntPtr(1); // swallow non-modifier keystrokes
        }

        return NativeMethods.CallNextHookEx(_keyboardHook, nCode, wParam, lParam);
    }

    private void PasteFromClipboard()
    {
        try
        {
            if (NativeMethods.OpenClipboard(IntPtr.Zero))
            {
                try
                {
                    var hData = NativeMethods.GetClipboardData(13); // CF_UNICODETEXT
                    if (hData != IntPtr.Zero)
                    {
                        var ptr = NativeMethods.GlobalLock(hData);
                        if (ptr != IntPtr.Zero)
                        {
                            var clipText = Marshal.PtrToStringUni(ptr) ?? "";
                            NativeMethods.GlobalUnlock(hData);
                            if (_inputText.Length + clipText.Length <= 200)
                                _inputText += clipText;
                            else
                                _inputText += clipText[..(200 - _inputText.Length)];
                            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                        }
                    }
                }
                finally { NativeMethods.CloseClipboard(); }
            }
        }
        catch { }
    }

    private void ResizePopup(int w, int h)
    {
        NativeMethods.GetWindowRect(_hWnd, out var wr);
        NativeMethods.SetWindowPos(_hWnd, NativeMethods.HWND_TOPMOST, wr.Left, wr.Top, w, h,
            NativeMethods.SWP_NOACTIVATE | NativeMethods.SWP_NOMOVE);
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

            case NativeMethods.WM_TIMER:
                if ((nuint)wParam.ToInt64() == TIMER_ID)
                    Dismiss();
                else if ((nuint)wParam.ToInt64() == ANIM_TIMER_ID)
                {
                    _animPhase = (_animPhase + 0.04f) % 1f;
                    NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                }
                return IntPtr.Zero;

            case NativeMethods.WM_LBUTTONDOWN:
                int clickX = NativeMethods.GET_X_LPARAM(lParam);
                int clickY = NativeMethods.GET_Y_LPARAM(lParam);

                // In input mode, ignore clicks
                if (_inputMode)
                    return IntPtr.Zero;

                if (_dropdownOpen)
                {
                    // Check if clicking on a dropdown item
                    if (_dropdownHoverIndex >= 0)
                    {
                        if (_dropdownHoverIndex == 0)
                        {
                            // Ask AI... (freeform input)
                            EnterInputMode();
                        }
                        else
                        {
                            // Index 1-10 → built-in tasks (EditTasks 1-10)
                            // Index 11+ → custom tasks (EditTasks 11+)
                            int di = _dropdownHoverIndex - 1; // offset past freeform
                            int taskIdx = di < TaskLabels.Length
                                ? di + 1
                                : 11 + (di - TaskLabels.Length);
                            StartProcessing();
                            QuickTaskClicked?.Invoke(taskIdx);
                        }
                    }
                    else
                    {
                        // Click outside dropdown — close it
                        _dropdownOpen = false;
                        ResizePopup(ScaledWidth, ScaledHeight);
                        NativeMethods.SetTimer(_hWnd, TIMER_ID, (uint)DurationMs, IntPtr.Zero);
                        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                    }
                    return IntPtr.Zero;
                }

                NativeMethods.KillTimer(_hWnd, TIMER_ID);

                if (_processing)
                {
                    // During processing, the stars area becomes a cancel button
                    if (clickX >= ScaledWidth - 56)
                    {
                        CancelRequested?.Invoke();
                        return IntPtr.Zero;
                    }
                    return IntPtr.Zero;
                }

                if (clickX >= ScaledWidth - 20) // Arrow area
                {
                    // Open dropdown
                    _dropdownOpen = true;
                    int dropH = TotalTaskCount * 26 + 16;
                    ResizePopup(ScaledWidth, ScaledHeight + dropH);
                    NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                    return IntPtr.Zero;
                }

                NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_HIDE);
                if (clickX >= ScaledWidth - 56)
                    AIEditClicked?.Invoke();
                else
                    Clicked?.Invoke();
                return IntPtr.Zero;

            case NativeMethods.WM_MOUSEMOVE:
                if (!_tracking)
                {
                    var tme = new NativeMethods.TRACKMOUSEEVENT
                    {
                        cbSize = Marshal.SizeOf<NativeMethods.TRACKMOUSEEVENT>(),
                        dwFlags = NativeMethods.TME_LEAVE,
                        hwndTrack = _hWnd
                    };
                    NativeMethods.TrackMouseEvent(ref tme);
                    _tracking = true;
                }
                if (!_hovered)
                {
                    _hovered = true;
                    // Pause auto-dismiss while hovering
                    NativeMethods.KillTimer(_hWnd, TIMER_ID);
                    NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                }
                int moveX = NativeMethods.GET_X_LPARAM(lParam);
                int moveY = NativeMethods.GET_Y_LPARAM(lParam);

                if (_dropdownOpen)
                {
                    // Dropdown hover detection
                    int dropY = ScaledHeight;
                    int dropItemH = 26;
                    int newDropIdx = -1;
                    if (moveY >= dropY + 8 && moveX >= 4 && moveX < ScaledWidth - 4)
                    {
                        int idx = (moveY - dropY - 8) / dropItemH;
                        if (idx >= 0 && idx < TotalTaskCount)
                            newDropIdx = idx;
                    }
                    if (newDropIdx != _dropdownHoverIndex)
                    {
                        _dropdownHoverIndex = newDropIdx;
                        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                    }
                    return IntPtr.Zero;
                }

                bool newAiHover = moveX >= ScaledWidth - 56 && moveX < ScaledWidth - 20;
                bool newArrowHover = moveX >= ScaledWidth - 20;
                if (newAiHover != _aiHovered || newArrowHover != _arrowHovered)
                {
                    _aiHovered = newAiHover;
                    _arrowHovered = newArrowHover;
                    NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                }
                return IntPtr.Zero;

            case NativeMethods.WM_MOUSELEAVE:
                _tracking = false;
                _hovered = false;
                _aiHovered = false;
                _arrowHovered = false;
                if (_dropdownOpen)
                {
                    _dropdownOpen = false;
                    _dropdownHoverIndex = -1;
                    ResizePopup(ScaledWidth, ScaledHeight);
                }
                // Resume auto-dismiss (unless processing or in input mode)
                if (!_processing && !_inputMode)
                    NativeMethods.SetTimer(_hWnd, TIMER_ID, (uint)DurationMs, IntPtr.Zero);
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                return IntPtr.Zero;
        }

        return NativeMethods.DefWindowProcW(hWnd, msg, wParam, lParam);
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

        var rect = new Rectangle(0, 0, w, h);

        // Background with rounded rect
        using var bgPath = CreateRoundedRect(rect, CORNER_RADIUS);
        using (var bgBrush = new SolidBrush(_hovered ? BgHoverColor : BgColor))
            g.FillPath(bgBrush, bgPath);

        // Colorful gradient border (animated when processing)
        {
            int totalPerimeter = 2 * (w + h);
            int segCount = 40;
            float segLen = totalPerimeter / (float)segCount;
            for (int si = 0; si < segCount; si++)
            {
                Color segColor;
                if (_processing)
                {
                    float t = (si / (float)segCount + _animPhase) % 1f;
                    segColor = Color.FromArgb(220, ColorFromHSL(t * 360f, 0.8f, 0.6f));
                }
                else
                {
                    Color[] borderColors = [
                        Color.FromArgb(168, 85, 247),
                        Color.FromArgb(59, 130, 246),
                        Color.FromArgb(16, 185, 129),
                        Color.FromArgb(245, 158, 11),
                        Color.FromArgb(239, 68, 68),
                    ];
                    float t = si / (float)segCount;
                    int ci = (int)(t * borderColors.Length) % borderColors.Length;
                    float blend = (t * borderColors.Length) % 1f;
                    var c1 = borderColors[ci];
                    var c2 = borderColors[(ci + 1) % borderColors.Length];
                    segColor = Color.FromArgb(
                        (int)(c1.R + (c2.R - c1.R) * blend),
                        (int)(c1.G + (c2.G - c1.G) * blend),
                        (int)(c1.B + (c2.B - c1.B) * blend));
                }
                using var pen = new Pen(segColor, BORDER_WIDTH);
                float pos = si * segLen;
                float nextPos = pos + segLen;
                float[] edges = [w, h, w, h];
                float offset = 0;
                for (int e = 0; e < 4; e++)
                {
                    float edgeEnd = offset + edges[e];
                    if (pos >= edgeEnd) { offset = edgeEnd; continue; }
                    if (nextPos <= offset) break;
                    float s = Math.Max(pos - offset, 0);
                    float en = Math.Min(nextPos - offset, edges[e]);
                    switch (e)
                    {
                        case 0: g.DrawLine(pen, s, 1, en, 1); break;
                        case 1: g.DrawLine(pen, w - 2, s, w - 2, en); break;
                        case 2: g.DrawLine(pen, w - s - 1, h - 2, w - en - 1, h - 2); break;
                        case 3: g.DrawLine(pen, 1, h - s - 1, 1, h - en - 1); break;
                    }
                    offset = edgeEnd;
                }
            }
        }

        // Preview text - centered (leave space for AI button + arrow)
        using var textFont = new Font("Segoe UI", 9f);
        using var textBrush = new SolidBrush(TextColor);

        if (_inputMode)
        {
            // Input mode: show ✨ icon + typed text or placeholder
            using var inputIconFont = new Font("Segoe UI Emoji", 8f);
            using var inputIconBrush = new SolidBrush(Color.FromArgb(220, 200, 100));
            var iconSize = g.MeasureString("\u2728", inputIconFont);
            float textH = g.MeasureString("Wg", textFont).Height;
            float centerY = (ScaledHeight - textH) / 2f;
            g.DrawString("\u2728", inputIconFont, inputIconBrush, 8, (ScaledHeight - iconSize.Height) / 2f);

            var inputRect = new RectangleF(36, centerY, w - 44, textH + 2);
            if (_inputText.Length == 0)
            {
                using var phBrush = new SolidBrush(Color.FromArgb(120, 120, 120));
                g.DrawString("Type an instruction...", textFont, phBrush, inputRect);
            }
            else
            {
                // Show text with cursor
                var displayText = _inputText + "\u2502"; // thin pipe cursor
                // Scroll: measure full text, if wider than rect, shift left
                var fullSize = g.MeasureString(displayText, textFont);
                float offsetX = 0;
                if (fullSize.Width > inputRect.Width)
                    offsetX = inputRect.Width - fullSize.Width;
                g.SetClip(inputRect);
                g.DrawString(displayText, textFont, textBrush, inputRect.X + offsetX, inputRect.Y);
                g.ResetClip();
            }
        }
        else
        {
        var textSize = g.MeasureString(_previewText, textFont);
        var textRect = new RectangleF(12, (ScaledHeight - textSize.Height) / 2f, w - 64, textSize.Height + 2);
        using var fmt = new StringFormat
        {
            Trimming = StringTrimming.EllipsisCharacter,
            FormatFlags = StringFormatFlags.NoWrap,
            Alignment = StringAlignment.Center
        };
        g.DrawString(_previewText, textFont, textBrush, textRect, fmt);
        }

        // AI edit button (stars) or Cancel button during processing (hidden in input mode)
        if (!_inputMode)
        {
        int aiBtnX = w - 52;
        int aiBtnY = (ScaledHeight - 24) / 2;
        if (_processing)
        {
            // Show compact cancel button (■)
            var cancelBg = _aiHovered || _arrowHovered ? Color.FromArgb(220, 60, 60) : Color.FromArgb(190, 40, 40);
            using var cancelBgBrush = new SolidBrush(cancelBg);
            int cancelSize = 22;
            var cancelRect = new Rectangle(aiBtnX + (24 - cancelSize) / 2, aiBtnY + (24 - cancelSize) / 2, cancelSize, cancelSize);
            using var cancelPath = CreateRoundedRect(cancelRect, cancelSize / 2);
            g.FillPath(cancelBgBrush, cancelPath);
            using var cancelFont = new Font("Segoe UI", 6f);
            using var cancelBrush = new SolidBrush(Color.White);
            using var cancelFmt = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString("\u25A0", cancelFont, cancelBrush, cancelRect, cancelFmt);
        }
        else
        {
            if (_aiHovered)
            {
                using var aiBgBrush = new SolidBrush(Color.FromArgb(60, 60, 60));
                var aiRect = new Rectangle(aiBtnX - 2, aiBtnY - 2, 28, 28);
                using var aiPath = CreateRoundedRect(aiRect, 6);
                g.FillPath(aiBgBrush, aiPath);
            }
            using var aiFont = new Font("Segoe UI Emoji", 10f);
            using var aiBrush = new SolidBrush(Color.FromArgb(220, 200, 100));
            var aiSize = g.MeasureString("\u2728", aiFont);
            g.DrawString("\u2728", aiFont, aiBrush,
                aiBtnX + (24 - aiSize.Width) / 2,
                aiBtnY + (24 - aiSize.Height) / 2);

            // Dropdown arrow button
            int arrowX = w - 18;
            int arrowY = (ScaledHeight - 24) / 2;
            if (_arrowHovered || _dropdownOpen)
            {
                using var arrBg = new SolidBrush(Color.FromArgb(60, 60, 60));
                var arrRect = new Rectangle(arrowX - 2, arrowY, 18, 24);
                using var arrPath = CreateRoundedRect(arrRect, 4);
                g.FillPath(arrBg, arrPath);
            }
            using var arrowFont = new Font("Segoe UI", 7f);
            using var arrowBrush = new SolidBrush(Color.FromArgb(180, 180, 180));
            string arrowChar = _dropdownOpen ? "▲" : "▼";
            var arrowSize = g.MeasureString(arrowChar, arrowFont);
            g.DrawString(arrowChar, arrowFont, arrowBrush,
                arrowX + (14 - arrowSize.Width) / 2,
                arrowY + (24 - arrowSize.Height) / 2);
        }
        } // end !_inputMode

        // Dropdown list
        if (_dropdownOpen)
        {
            int dropY = ScaledHeight;
            int dropItemH = 26;

            using var dropBg = new SolidBrush(Color.FromArgb(38, 38, 38));
            var dropRect = new Rectangle(0, dropY, w, h - dropY);
            g.FillRectangle(dropBg, dropRect);

            using var sepPen = new Pen(Color.FromArgb(55, 55, 55));
            g.DrawLine(sepPen, 4, dropY, w - 4, dropY);

            using var taskFont = new Font("Segoe UI", 8.5f);
            using var emojiFont = new Font("Segoe UI Emoji", 8.5f);
            using var dimBrush = new SolidBrush(Color.FromArgb(160, 160, 160));

            for (int i = 0; i < TotalTaskCount; i++)
            {
                int iy = dropY + 8 + i * dropItemH;

                // Separator after freeform entry
                if (i == 1)
                    g.DrawLine(sepPen, 12, iy, w - 12, iy);

                if (_dropdownHoverIndex == i)
                {
                    using var hlBrush = new SolidBrush(Color.FromArgb(55, 55, 55));
                    var hlRect = new Rectangle(4, iy, w - 8, dropItemH);
                    using var hlPath = CreateRoundedRect(hlRect, 4);
                    g.FillPath(hlBrush, hlPath);
                }

                string icon, label, hk;
                if (i == 0)
                {
                    icon = "\u2728";
                    label = "Ask AI...";
                    hk = FreeformHotkey;
                }
                else if (i <= TaskLabels.Length)
                {
                    int ti = i - 1;
                    icon = TaskIcons[ti];
                    label = TaskLabels[ti];
                    hk = (TaskHotkeys != null && ti < TaskHotkeys.Length) ? TaskHotkeys[ti] : "";
                }
                else
                {
                    int ci = i - 1 - TaskLabels.Length;
                    icon = "\u26a1";
                    label = CustomTasks[ci].Name;
                    hk = CustomTasks[ci].Hotkey;
                }

                var iconSz = g.MeasureString(icon, emojiFont);
                g.DrawString(icon, emojiFont, textBrush, 12, iy + (dropItemH - iconSz.Height) / 2);

                var displayText = string.IsNullOrEmpty(hk) ? label : $"{label}  ({hk})";
                g.DrawString(displayText, taskFont, _dropdownHoverIndex == i ? textBrush : dimBrush, 12 + iconSz.Width + 2, iy + 5);
            }
        }

        // Blit to screen using GDI BitBlt (DPI-immune raw pixel copy)
        IntPtr hdcMem = NativeMethods.CreateCompatibleDC(ps.hdc);
        IntPtr hBitmap = bmp.GetHbitmap();
        IntPtr hOldBmp = NativeMethods.SelectObject(hdcMem, hBitmap);
        BitBlt(ps.hdc, 0, 0, w, h, hdcMem, 0, 0, SRCCOPY);
        NativeMethods.SelectObject(hdcMem, hOldBmp);
        NativeMethods.DeleteObject(hBitmap);
        NativeMethods.DeleteDC(hdcMem);

        NativeMethods.EndPaint(hWnd, ref ps);
    }

    [DllImport("gdi32.dll")]
    private static extern bool BitBlt(IntPtr hdcDest, int xDest, int yDest, int wDest, int hDest,
        IntPtr hdcSrc, int xSrc, int ySrc, uint rop);
    private const uint SRCCOPY = 0x00CC0020;

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

    private static string TruncateText(string text, int maxLength)
    {
        var single = text.ReplaceLineEndings(" ");
        return single.Length > maxLength ? single[..maxLength] + "…" : single;
    }

    private static Color ColorFromHSL(float h, float s, float l)
    {
        float c = (1 - Math.Abs(2 * l - 1)) * s;
        float x = c * (1 - Math.Abs(h / 60 % 2 - 1));
        float m = l - c / 2;
        float r, g, b;
        if (h < 60) { r = c; g = x; b = 0; }
        else if (h < 120) { r = x; g = c; b = 0; }
        else if (h < 180) { r = 0; g = c; b = x; }
        else if (h < 240) { r = 0; g = x; b = c; }
        else if (h < 300) { r = x; g = 0; b = c; }
        else { r = c; g = 0; b = x; }
        return Color.FromArgb((int)((r + m) * 255), (int)((g + m) * 255), (int)((b + m) * 255));
    }
}
