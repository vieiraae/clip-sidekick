using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace ClipSidekick;

internal sealed class SettingsDialog
{
    private const int WIDTH = 440;
    private const int ADD_BUTTON_ROW_HEIGHT = 36;
    private const int ROW_HEIGHT = 44;
    private const int TASK_ROW_HEIGHT = 36;
    private const int LABEL_X = 20;
    private const int CONTROL_RIGHT = 420;
    private const int HEADER_HEIGHT = 48;
    private const int TAB_BAR_HEIGHT = 36;
    private const int CONTENT_TOP = HEADER_HEIGHT + TAB_BAR_HEIGHT;

    private static readonly string[] TabLabels = ["General", "Quick Tasks", "Custom Tasks", "Advanced"];
    private static readonly string[] TaskNames = ["Proofread", "Rewrite", "Use synonyms", "Minor revise", "Major revise", "Describe", "Answer", "Explain", "Expand", "Summarize"];

    private IntPtr _hWnd;
    private NativeMethods.WndProc _wndProc = null!;
    private bool _done;
    private AppSettings _settings = null!;
    private bool _inHotkeyDialog;
    private bool _inEditDialog;
    private IntPtr _mouseHook;
    private NativeMethods.LowLevelMouseProc _mouseHookProc = null!;

    // Editable values
    private int _notificationDurationMs;
    private string _hotkey = "";
    private int _maxHistoryItems;
    private bool _showNotification;
    private string[] _taskHotkeys = new string[10];
    private List<CustomTask> _customTasks = [];
    private string _freeformHotkey = "";
    private string _bubbleHotkey = "";
    private bool _mcpEnabled;
    private string[] _mcpServerNames = [];
    private HashSet<string> _disabledMcpServers = [];
    private bool _skillsEnabled;

    // Model dropdown
    private string[] _models = [""];
    private int _modelIndex;
    private bool _modelDropdownOpen;
    private int _modelScrollOffset; // scroll offset in pixels for dropdown
    private bool _modelScrollbarDragging;
    private int _modelScrollbarDragStartY;
    private int _modelScrollbarDragStartOffset;
    private string _systemMessage = "";

    // UI state — Tab 0=General, 1=Quick Tasks, 2=Custom Tasks
    // Row indices: 0-4=general, 5=ask AI, 6-15=built-in task hotkeys, 16=add custom task, 17+=custom tasks
    private int _activeSettingsTab;
    private int _hoverTab = -1;
    private int _hoverRow = -1;
    private int _hoverButton = -1;
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
    private static readonly Color TabBg = Color.FromArgb(28, 28, 28);
    private static readonly Color TabHoverBg = Color.FromArgb(40, 40, 40);

    [DllImport("user32.dll")]
    private static extern bool SetLayeredWindowAttributes(IntPtr hwnd, uint crKey, byte bAlpha, uint dwFlags);

    public bool Show(IntPtr parent, AppSettings settings, string[] models, int modelIndex)
    {
        _settings = settings;
        _models = models.Length > 0 ? models : [""];
        _modelIndex = Math.Clamp(modelIndex, 0, _models.Length - 1);
        _modelDropdownOpen = false;
        _modelScrollOffset = 0;
        _notificationDurationMs = settings.NotificationDurationMs;
        _hotkey = settings.Hotkey;
        _maxHistoryItems = settings.MaxHistoryItems;
        _showNotification = settings.ShowNotification;
        _taskHotkeys = (string[])settings.QuickTaskHotkeys.Clone();
        if (_taskHotkeys.Length < 10)
            _taskHotkeys = [.. _taskHotkeys, .. new string[10 - _taskHotkeys.Length]];
        _customTasks = settings.CustomTasks.Select(t => new CustomTask { Name = t.Name, Prompt = t.Prompt, Hotkey = t.Hotkey }).ToList();
        _freeformHotkey = settings.FreeformHotkey ?? "";
        _bubbleHotkey = settings.BubbleHotkey ?? "";
        _mcpEnabled = settings.McpEnabled;
        _disabledMcpServers = new HashSet<string>(settings.DisabledMcpServers);
        _skillsEnabled = settings.SkillsEnabled;
        _systemMessage = settings.SystemMessage ?? "";
        _mcpServerNames = LoadMcpServerNames();
        _activeSettingsTab = 0;
        _done = false;
        _inHotkeyDialog = false;
        _inEditDialog = false;

        int windowHeight = ComputeWindowHeight();

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
        int y = Math.Clamp(cursor.Y - windowHeight / 2, mi.rcWork.Top, mi.rcWork.Bottom - windowHeight);

        _hWnd = NativeMethods.CreateWindowExW(
            NativeMethods.WS_EX_TOOLWINDOW | NativeMethods.WS_EX_TOPMOST | NativeMethods.WS_EX_LAYERED,
            className, "Settings",
            NativeMethods.WS_POPUP,
            x, y, WIDTH, windowHeight,
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

        if (!_done)
            NativeMethods.PostQuitMessage(0);

        RemoveMouseHook();
        NativeMethods.DestroyWindow(_hWnd);
        NativeMethods.UnregisterClassW(classNamePtr, hInstance);
        Marshal.FreeHGlobal(classNamePtr);

        return true;
    }

    private int ComputeWindowHeight() => _activeSettingsTab switch
    {
        0 => CONTENT_TOP + ROW_HEIGHT * 7 + 8,
        1 => CONTENT_TOP + ROW_HEIGHT + 28 + TASK_ROW_HEIGHT * 10 + 8,
        2 => CONTENT_TOP + ADD_BUTTON_ROW_HEIGHT + Math.Max(1, _customTasks.Count) * TASK_ROW_HEIGHT + 8,
        3 => CONTENT_TOP + ROW_HEIGHT * 4 + (_mcpServerNames.Length > 0 ? 24 + _mcpServerNames.Length * TASK_ROW_HEIGHT : 0) + 8,
        _ => CONTENT_TOP + ROW_HEIGHT * 5 + 8
    };

    private void ResizeWindow()
    {
        int newHeight = ComputeWindowHeight();
        NativeMethods.GetWindowRect(_hWnd, out var wr);
        int x = wr.Left;
        int y = wr.Top;

        // Ensure window stays on screen after resize
        var pt = new NativeMethods.POINT { X = x + WIDTH / 2, Y = y + newHeight / 2 };
        var monitor = NativeMethods.MonitorFromPoint(pt, NativeMethods.MONITOR_DEFAULTTONEAREST);
        var mi = new NativeMethods.MONITORINFO { cbSize = Marshal.SizeOf<NativeMethods.MONITORINFO>() };
        NativeMethods.GetMonitorInfoW(monitor, ref mi);
        y = Math.Clamp(y, mi.rcWork.Top, Math.Max(mi.rcWork.Top, mi.rcWork.Bottom - newHeight));

        NativeMethods.SetWindowPos(_hWnd, IntPtr.Zero, x, y, WIDTH, newHeight, NativeMethods.SWP_NOZORDER);
        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
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

            case NativeMethods.WM_LBUTTONUP:
                if (_modelScrollbarDragging)
                {
                    _modelScrollbarDragging = false;
                    NativeMethods.ReleaseCapture();
                }
                return IntPtr.Zero;

            case NativeMethods.WM_APP_DISMISS:
                _done = true;
                return IntPtr.Zero;

            case NativeMethods.WM_KEYDOWN:
                if (wParam.ToInt32() == NativeMethods.VK_ESCAPE)
                {
                    if (_modelDropdownOpen)
                    { _modelDropdownOpen = false; NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false); }
                    else
                        _done = true;
                }
                return IntPtr.Zero;

            case 0x020A: // WM_MOUSEWHEEL
                if (_modelDropdownOpen)
                {
                    int delta = (short)(wParam.ToInt64() >> 16);
                    const int DROP_ITEM_H = 28;
                    int maxVisible = 8;
                    int totalH = _models.Length * DROP_ITEM_H + 4;
                    int visibleH = maxVisible * DROP_ITEM_H + 4;
                    if (totalH > visibleH)
                    {
                        _modelScrollOffset = Math.Clamp(_modelScrollOffset - delta / 4, 0, totalH - visibleH);
                        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                    }
                    return IntPtr.Zero;
                }
                break;

            case NativeMethods.WM_ACTIVATE:
                if (NativeMethods.LOWORD(wParam) == NativeMethods.WA_INACTIVE && !_inHotkeyDialog && !_inEditDialog)
                    _done = true;
                return IntPtr.Zero;

            case NativeMethods.WM_DESTROY:
                _done = true;
                return IntPtr.Zero;
        }

        return NativeMethods.DefWindowProcW(hWnd, msg, wParam, lParam);
    }

    private const int TAB_COUNT = 4;

    // Tab hit test: returns 0..TAB_COUNT-1 if in tab bar, else -1
    private int HitTestTab(int mx, int my)
    {
        if (my < HEADER_HEIGHT || my >= HEADER_HEIGHT + TAB_BAR_HEIGHT) return -1;
        int tabW = WIDTH / TAB_COUNT;
        for (int i = 0; i < TAB_COUNT; i++)
        {
            int tabX = i * tabW;
            int tabRight = (i == TAB_COUNT - 1) ? WIDTH : tabX + tabW;
            if (mx >= tabX && mx < tabRight) return i;
        }
        return -1;
    }

    private void OnMouseMove(IntPtr lParam)
    {
        int mx = NativeMethods.GET_X_LPARAM(lParam);
        int my = NativeMethods.GET_Y_LPARAM(lParam);

        if (_modelScrollbarDragging)
        {
            const int DROP_ITEM_H = 28;
            int maxVisible = Math.Min(_models.Length, 8);
            int dropH = maxVisible * DROP_ITEM_H + 4;
            int totalH = _models.Length * DROP_ITEM_H + 4;
            int maxScroll = totalH - dropH;
            int sbTrackH = dropH - 8;
            int sbThumbH = Math.Max(20, sbTrackH * dropH / totalH);
            int trackSpace = sbTrackH - sbThumbH;
            if (trackSpace > 0)
            {
                int deltaY = my - _modelScrollbarDragStartY;
                int newOffset = _modelScrollbarDragStartOffset + (int)((float)deltaY / trackSpace * maxScroll);
                _modelScrollOffset = Math.Clamp(newOffset, 0, maxScroll);
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            }
            return;
        }

        int newRow = -1, newButton = -1;
        bool newHoverClose = false;
        int newHoverTab = HitTestTab(mx, my);

        if (mx >= WIDTH - 40 && mx < WIDTH - 8 && my >= 8 && my < 40)
            newHoverClose = true;

        if (my >= CONTENT_TOP)
        {
            switch (_activeSettingsTab)
            {
                case 0: HoverGeneral(mx, my, ref newRow, ref newButton); break;
                case 1: HoverQuickTasks(mx, my, ref newRow, ref newButton); break;
                case 2: HoverCustomTasks(mx, my, ref newRow, ref newButton); break;
                case 3: HoverAdvanced(mx, my, ref newRow, ref newButton); break;
            }
        }

        if (newRow != _hoverRow || newButton != _hoverButton || newHoverClose != _hoverClose || newHoverTab != _hoverTab)
        {
            _hoverRow = newRow;
            _hoverButton = newButton;
            _hoverClose = newHoverClose;
            _hoverTab = newHoverTab;
            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        }
    }

    private void HoverGeneral(int mx, int my, ref int newRow, ref int newButton)
    {
        // Model dropdown overlay takes priority
        if (_modelDropdownOpen)
        {
            const int DROP_ITEM_H = 28;
            int maxVisible = Math.Min(_models.Length, 8);
            int dropH = maxVisible * DROP_ITEM_H + 4;
            int dropY = CONTENT_TOP + ROW_HEIGHT + 2;
            int dropW = WIDTH - 40;
            int dropX = LABEL_X;
            if (mx >= dropX && mx < dropX + dropW && my >= dropY && my < dropY + dropH)
            {
                int idx = (my - dropY - 2 + _modelScrollOffset) / DROP_ITEM_H;
                if (idx >= 0 && idx < _models.Length)
                    newRow = 51 + idx;
            }
            return;
        }

        int rowY = CONTENT_TOP;

        // Row 50: Model selector
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            newRow = 50;
            if (mx >= CONTROL_RIGHT - 200 && mx < CONTROL_RIGHT - 6 && my >= rowY + 8 && my < rowY + 36) newButton = 0;
        }
        rowY += ROW_HEIGHT;

        // Row 55: System message
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            newRow = 55;
            if (mx >= CONTROL_RIGHT - 60 && mx < CONTROL_RIGHT - 10 && my >= rowY + 8 && my < rowY + 36) newButton = 0;
        }
        rowY += ROW_HEIGHT;

        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            newRow = 0;
            if (mx >= CONTROL_RIGHT - 130 && mx < CONTROL_RIGHT - 102 && my >= rowY + 8 && my < rowY + 36) newButton = 0;
            else if (mx >= CONTROL_RIGHT - 30 && mx < CONTROL_RIGHT - 2 && my >= rowY + 8 && my < rowY + 36) newButton = 1;
        }
        rowY += ROW_HEIGHT;
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            newRow = 1;
            if (mx >= CONTROL_RIGHT - 50 && mx < CONTROL_RIGHT - 6 && my >= rowY + 10 && my < rowY + 34) newButton = 0;
        }
        rowY += ROW_HEIGHT;
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            newRow = 2;
            if (mx >= CONTROL_RIGHT - 130 && mx < CONTROL_RIGHT - 102 && my >= rowY + 8 && my < rowY + 36) newButton = 0;
            else if (mx >= CONTROL_RIGHT - 30 && mx < CONTROL_RIGHT - 2 && my >= rowY + 8 && my < rowY + 36) newButton = 1;
        }
        rowY += ROW_HEIGHT;
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            newRow = 3;
            if (mx >= CONTROL_RIGHT - 60 && mx < CONTROL_RIGHT - 10 && my >= rowY + 8 && my < rowY + 36) newButton = 0;
            else if (!string.IsNullOrEmpty(_bubbleHotkey) && mx >= CONTROL_RIGHT - 84 && mx < CONTROL_RIGHT - 60 && my >= rowY + 10 && my < rowY + 34) newButton = 1;
        }
        rowY += ROW_HEIGHT;
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            newRow = 4;
            if (mx >= CONTROL_RIGHT - 60 && mx < CONTROL_RIGHT - 10 && my >= rowY + 8 && my < rowY + 36) newButton = 0;
            else if (mx >= CONTROL_RIGHT - 84 && mx < CONTROL_RIGHT - 60 && my >= rowY + 10 && my < rowY + 34) newButton = 1;
        }
    }

    private void HoverQuickTasks(int mx, int my, ref int newRow, ref int newButton)
    {
        int rowY = CONTENT_TOP;
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            newRow = 5;
            if (mx >= CONTROL_RIGHT - 60 && mx < CONTROL_RIGHT - 10 && my >= rowY + 8 && my < rowY + 36) newButton = 0;
            else if (!string.IsNullOrEmpty(_freeformHotkey) && mx >= CONTROL_RIGHT - 84 && mx < CONTROL_RIGHT - 60 && my >= rowY + 10 && my < rowY + 34) newButton = 1;
        }
        int taskStartY = CONTENT_TOP + ROW_HEIGHT + 28;
        for (int i = 0; i < 10; i++)
        {
            int ty = taskStartY + i * TASK_ROW_HEIGHT;
            if (my >= ty && my < ty + TASK_ROW_HEIGHT)
            {
                newRow = 6 + i;
                if (mx >= CONTROL_RIGHT - 60 && mx < CONTROL_RIGHT - 10 && my >= ty + 6 && my < ty + 30) newButton = 0;
                else if (!string.IsNullOrEmpty(_taskHotkeys[i]) && mx >= CONTROL_RIGHT - 84 && mx < CONTROL_RIGHT - 60 && my >= ty + 6 && my < ty + 30) newButton = 1;
                break;
            }
        }
    }

    private void HoverCustomTasks(int mx, int my, ref int newRow, ref int newButton)
    {
        if (my >= CONTENT_TOP && my < CONTENT_TOP + ADD_BUTTON_ROW_HEIGHT)
        {
            newRow = 16;
            if (mx >= LABEL_X && mx < LABEL_X + 110 && my >= CONTENT_TOP + 6 && my < CONTENT_TOP + 30) newButton = 0;
        }
        int customStartY = CONTENT_TOP + ADD_BUTTON_ROW_HEIGHT;
        for (int ci = 0; ci < _customTasks.Count; ci++)
        {
            int cy = customStartY + ci * TASK_ROW_HEIGHT;
            if (my >= cy && my < cy + TASK_ROW_HEIGHT)
            {
                newRow = 17 + ci;
                if (mx >= LABEL_X && mx < LABEL_X + 120 && my >= cy + 6 && my < cy + 30) newButton = 2;
                int delX = LABEL_X + 124;
                if (mx >= delX && mx < delX + 24 && my >= cy + 6 && my < cy + 30) newButton = 3;
                if (mx >= CONTROL_RIGHT - 60 && mx < CONTROL_RIGHT - 10 && my >= cy + 6 && my < cy + 30) newButton = 0;
                else if (!string.IsNullOrEmpty(_customTasks[ci].Hotkey) && mx >= CONTROL_RIGHT - 84 && mx < CONTROL_RIGHT - 60 && my >= cy + 6 && my < cy + 30) newButton = 1;
                break;
            }
        }
    }

    private void OnMouseDown(IntPtr lParam)
    {
        int mx = NativeMethods.GET_X_LPARAM(lParam);
        int my = NativeMethods.GET_Y_LPARAM(lParam);

        if (mx >= WIDTH - 40 && mx < WIDTH - 8 && my >= 8 && my < 40)
        {
            _done = true;
            return;
        }

        int clickedTab = HitTestTab(mx, my);
        if (clickedTab >= 0 && clickedTab != _activeSettingsTab)
        {
            _activeSettingsTab = clickedTab;
            _hoverRow = -1;
            _hoverButton = -1;
            _modelDropdownOpen = false;
            ResizeWindow();
            return;
        }

        if (my < CONTENT_TOP) return;

        switch (_activeSettingsTab)
        {
            case 0: ClickGeneral(mx, my); break;
            case 1: ClickQuickTasks(mx, my); break;
            case 2: ClickCustomTasks(mx, my); break;
            case 3: ClickAdvanced(mx, my); break;
        }
    }

    private void ClickGeneral(int mx, int my)
    {
        // Handle model dropdown selection
        if (_modelDropdownOpen)
        {
            const int DROP_ITEM_H = 28;
            int maxVisible = Math.Min(_models.Length, 8);
            int dropH = maxVisible * DROP_ITEM_H + 4;
            int totalH = _models.Length * DROP_ITEM_H + 4;
            int dropY = CONTENT_TOP + ROW_HEIGHT + 2;
            int dropW = WIDTH - 40;
            int dropX = LABEL_X;

            // Check scrollbar click first
            if (totalH > dropH && mx >= dropX + dropW - 10 && mx < dropX + dropW && my >= dropY && my < dropY + dropH)
            {
                _modelScrollbarDragging = true;
                _modelScrollbarDragStartY = my;
                _modelScrollbarDragStartOffset = _modelScrollOffset;
                NativeMethods.SetCapture(_hWnd);
                return;
            }

            if (mx >= dropX && mx < dropX + dropW && my >= dropY && my < dropY + dropH)
            {
                int idx = (my - dropY - 2 + _modelScrollOffset) / DROP_ITEM_H;
                if (idx >= 0 && idx < _models.Length && idx != _modelIndex)
                {
                    _modelIndex = idx;
                    SaveSettings();
                }
            }
            _modelDropdownOpen = false;
            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            return;
        }

        int rowY = CONTENT_TOP;

        // Row 50: Model selector - toggle dropdown
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            if (mx >= CONTROL_RIGHT - 200 && mx < CONTROL_RIGHT - 6 && my >= rowY + 8 && my < rowY + 36)
            {
                _modelDropdownOpen = !_modelDropdownOpen;
                _modelScrollOffset = 0;
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            }
            return;
        }
        rowY += ROW_HEIGHT;

        // Row 55: System message - edit
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            if (mx >= CONTROL_RIGHT - 60 && mx < CONTROL_RIGHT - 10 && my >= rowY + 8 && my < rowY + 36)
                OpenSystemMessageDialog();
            return;
        }
        rowY += ROW_HEIGHT;

        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            if (mx >= CONTROL_RIGHT - 130 && mx < CONTROL_RIGHT - 102 && my >= rowY + 8 && my < rowY + 36)
            { _maxHistoryItems = Math.Max(10, _maxHistoryItems - 10); SaveSettings(); NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false); }
            else if (mx >= CONTROL_RIGHT - 30 && mx < CONTROL_RIGHT - 2 && my >= rowY + 8 && my < rowY + 36)
            { _maxHistoryItems = Math.Min(200, _maxHistoryItems + 10); SaveSettings(); NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false); }
            return;
        }
        rowY += ROW_HEIGHT;
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            if (mx >= CONTROL_RIGHT - 50 && mx < CONTROL_RIGHT - 6 && my >= rowY + 10 && my < rowY + 34)
            { _showNotification = !_showNotification; SaveSettings(); NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false); }
            return;
        }
        rowY += ROW_HEIGHT;
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            if (mx >= CONTROL_RIGHT - 130 && mx < CONTROL_RIGHT - 102 && my >= rowY + 8 && my < rowY + 36)
            { _notificationDurationMs = Math.Max(500, _notificationDurationMs - 500); SaveSettings(); NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false); }
            else if (mx >= CONTROL_RIGHT - 30 && mx < CONTROL_RIGHT - 2 && my >= rowY + 8 && my < rowY + 36)
            { _notificationDurationMs = Math.Min(10000, _notificationDurationMs + 500); SaveSettings(); NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false); }
            return;
        }
        rowY += ROW_HEIGHT;
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            if (mx >= CONTROL_RIGHT - 60 && mx < CONTROL_RIGHT - 10 && my >= rowY + 8 && my < rowY + 36)
                OpenHotkeyDialog(ref _bubbleHotkey);
            else if (!string.IsNullOrEmpty(_bubbleHotkey) && mx >= CONTROL_RIGHT - 84 && mx < CONTROL_RIGHT - 60 && my >= rowY + 10 && my < rowY + 34)
            { _bubbleHotkey = ""; SaveSettings(); NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false); }
            return;
        }
        rowY += ROW_HEIGHT;
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            if (mx >= CONTROL_RIGHT - 60 && mx < CONTROL_RIGHT - 10 && my >= rowY + 8 && my < rowY + 36)
                OpenHotkeyDialog(ref _hotkey);
            else if (mx >= CONTROL_RIGHT - 84 && mx < CONTROL_RIGHT - 60 && my >= rowY + 10 && my < rowY + 34)
            { _hotkey = "Win+Shift+V"; SaveSettings(); NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false); }
        }
    }

    private void ClickQuickTasks(int mx, int my)
    {
        int rowY = CONTENT_TOP;
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            if (mx >= CONTROL_RIGHT - 60 && mx < CONTROL_RIGHT - 10 && my >= rowY + 8 && my < rowY + 36)
                OpenHotkeyDialog(ref _freeformHotkey);
            else if (!string.IsNullOrEmpty(_freeformHotkey) && mx >= CONTROL_RIGHT - 84 && mx < CONTROL_RIGHT - 60 && my >= rowY + 10 && my < rowY + 34)
            { _freeformHotkey = ""; SaveSettings(); NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false); }
            return;
        }
        int taskStartY = CONTENT_TOP + ROW_HEIGHT + 28;
        for (int i = 0; i < 10; i++)
        {
            int ty = taskStartY + i * TASK_ROW_HEIGHT;
            if (my >= ty && my < ty + TASK_ROW_HEIGHT)
            {
                if (mx >= CONTROL_RIGHT - 60 && mx < CONTROL_RIGHT - 10 && my >= ty + 6 && my < ty + 30)
                    OpenTaskHotkeyDialog(i);
                else if (!string.IsNullOrEmpty(_taskHotkeys[i]) && mx >= CONTROL_RIGHT - 84 && mx < CONTROL_RIGHT - 60 && my >= ty + 6 && my < ty + 30)
                { _taskHotkeys[i] = ""; SaveSettings(); NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false); }
                return;
            }
        }
    }

    private void ClickCustomTasks(int mx, int my)
    {
        if (my >= CONTENT_TOP && my < CONTENT_TOP + ADD_BUTTON_ROW_HEIGHT)
        {
            if (mx >= LABEL_X && mx < LABEL_X + 110 && my >= CONTENT_TOP + 6 && my < CONTENT_TOP + 30)
            {
                _inEditDialog = true;
                RemoveMouseHook();
                NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_HIDE);
                var editDialog = new CustomTaskEditDialog();
                var result = editDialog.Show(_hWnd, "", "");
                if (result != null)
                {
                    _customTasks.Add(new CustomTask { Name = result.Value.name, Prompt = result.Value.prompt });
                    SaveSettings();
                    ResizeWindow();
                }
                NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_SHOW);
                NativeMethods.SetForegroundWindow(_hWnd);
                _inEditDialog = false;
                InstallMouseHook();
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            }
            return;
        }
        int customStartY = CONTENT_TOP + ADD_BUTTON_ROW_HEIGHT;
        for (int ci = 0; ci < _customTasks.Count; ci++)
        {
            int cy = customStartY + ci * TASK_ROW_HEIGHT;
            if (my >= cy && my < cy + TASK_ROW_HEIGHT)
            {
                if (mx >= LABEL_X && mx < LABEL_X + 120 && my >= cy + 6 && my < cy + 30)
                {
                    _inEditDialog = true;
                    RemoveMouseHook();
                    NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_HIDE);
                    var editDialog = new CustomTaskEditDialog();
                    var result = editDialog.Show(_hWnd, _customTasks[ci].Name, _customTasks[ci].Prompt);
                    if (result != null)
                    {
                        _customTasks[ci].Name = result.Value.name;
                        _customTasks[ci].Prompt = result.Value.prompt;
                        SaveSettings();
                    }
                    NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_SHOW);
                    NativeMethods.SetForegroundWindow(_hWnd);
                    _inEditDialog = false;
                    InstallMouseHook();
                    NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                }
                int delX = LABEL_X + 124;
                if (mx >= delX && mx < delX + 24 && my >= cy + 6 && my < cy + 30)
                {
                    _customTasks.RemoveAt(ci);
                    SaveSettings();
                    ResizeWindow();
                    return;
                }
                if (mx >= CONTROL_RIGHT - 60 && mx < CONTROL_RIGHT - 10 && my >= cy + 6 && my < cy + 30)
                {
                    _inHotkeyDialog = true;
                    RemoveMouseHook();
                    NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_HIDE);
                    var dialog = new HotkeyDialog();
                    var result = dialog.Show(_hWnd, _customTasks[ci].Hotkey);
                    if (result != null) { _customTasks[ci].Hotkey = result; SaveSettings(); }
                    NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_SHOW);
                    NativeMethods.SetForegroundWindow(_hWnd);
                    _inHotkeyDialog = false;
                    InstallMouseHook();
                    NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                }
                else if (!string.IsNullOrEmpty(_customTasks[ci].Hotkey) && mx >= CONTROL_RIGHT - 84 && mx < CONTROL_RIGHT - 60 && my >= cy + 6 && my < cy + 30)
                {
                    _customTasks[ci].Hotkey = "";
                    SaveSettings();
                    NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                }
                return;
            }
        }
    }

    private void OpenHotkeyDialog(ref string field)
    {
        _inHotkeyDialog = true;
        RemoveMouseHook();
        NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_HIDE);
        var dialog = new HotkeyDialog();
        var result = dialog.Show(_hWnd, field);
        if (result != null) { field = result; SaveSettings(); }
        NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_SHOW);
        NativeMethods.SetForegroundWindow(_hWnd);
        _inHotkeyDialog = false;
        InstallMouseHook();
        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
    }

    private void OpenSystemMessageDialog()
    {
        _inEditDialog = true;
        RemoveMouseHook();
        NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_HIDE);
        var editDialog = new SystemMessageEditDialog();
        var result = editDialog.Show(_hWnd, _systemMessage);
        if (result != null)
        {
            _systemMessage = result;
            SaveSettings();
        }
        NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_SHOW);
        NativeMethods.SetForegroundWindow(_hWnd);
        _inEditDialog = false;
        InstallMouseHook();
        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
    }

    private void OpenTaskHotkeyDialog(int i)
    {
        _inHotkeyDialog = true;
        RemoveMouseHook();
        NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_HIDE);
        var dialog = new HotkeyDialog();
        var result = dialog.Show(_hWnd, _taskHotkeys[i]);
        if (result != null) { _taskHotkeys[i] = result; SaveSettings(); }
        NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_SHOW);
        NativeMethods.SetForegroundWindow(_hWnd);
        _inHotkeyDialog = false;
        InstallMouseHook();
        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
    }

    private void SaveSettings()
    {
        _settings.Model = _modelIndex < _models.Length ? _models[_modelIndex] : "";
        _settings.NotificationDurationMs = _notificationDurationMs;
        _settings.Hotkey = _hotkey;
        _settings.MaxHistoryItems = _maxHistoryItems;
        _settings.ShowNotification = _showNotification;
        _settings.QuickTaskHotkeys = (string[])_taskHotkeys.Clone();
        _settings.CustomTasks = _customTasks.Select(t => new CustomTask { Name = t.Name, Prompt = t.Prompt, Hotkey = t.Hotkey }).ToList();
        _settings.FreeformHotkey = _freeformHotkey;
        _settings.BubbleHotkey = _bubbleHotkey;
        _settings.McpEnabled = _mcpEnabled;
        _settings.DisabledMcpServers = [.. _disabledMcpServers];
        _settings.SkillsEnabled = _skillsEnabled;
        _settings.SystemMessage = _systemMessage;
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

        if (w <= 0 || h <= 0) { NativeMethods.EndPaint(hWnd, ref ps); return; }

        using var bmp = new Bitmap(w, h);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

        // Background
        using (var bgBrush = new SolidBrush(BgColor))
            g.FillRectangle(bgBrush, 0, 0, w, h);
        using (var borderPen = new Pen(BorderColor))
            g.DrawRectangle(borderPen, 0, 0, w - 1, h - 1);

        // Header title
        using var titleFont = new Font("Segoe UI", 12f, FontStyle.Regular);
        using var titleBrush = new SolidBrush(TextColor);
        g.DrawString("Settings", titleFont, titleBrush, 20, 14);

        // Close button
        int closeBtnX = w - 36;
        if (_hoverClose)
        {
            using var hoverBrush = new SolidBrush(ButtonHoverBg);
            var closeRect = new Rectangle(closeBtnX - 2, 8, 28, 28);
            using var closePath = CreateRoundedRect(closeRect, 4);
            g.FillPath(hoverBrush, closePath);
        }
        using var closeFont = new Font("Segoe UI", 10f);
        using var closeBrush = new SolidBrush(DimText);
        g.DrawString("\u00D7", closeFont, closeBrush, closeBtnX, 10);

        using var sepPen = new Pen(SeparatorColor);
        g.DrawLine(sepPen, 0, HEADER_HEIGHT - 1, w, HEADER_HEIGHT - 1);

        // Draw tab bar
        int tabW = w / TAB_COUNT;
        for (int i = 0; i < TAB_COUNT; i++)
        {
            int tabX = i * tabW;
            int tabRight = (i == TAB_COUNT - 1) ? w : tabX + tabW;
            int thisTabW = tabRight - tabX;
            bool isActive = i == _activeSettingsTab;
            bool isHover = i == _hoverTab && !isActive;

            Color tabBg = isActive ? BgColor : (isHover ? TabHoverBg : TabBg);
            using (var tabBrush = new SolidBrush(tabBg))
                g.FillRectangle(tabBrush, tabX, HEADER_HEIGHT, thisTabW, TAB_BAR_HEIGHT);

            if (isActive)
            {
                using var accentPen = new Pen(AccentColor, 2);
                g.DrawLine(accentPen, tabX + 2, HEADER_HEIGHT + TAB_BAR_HEIGHT - 2, tabRight - 2, HEADER_HEIGHT + TAB_BAR_HEIGHT - 2);
            }

            if (i < TAB_COUNT - 1)
            {
                using var tabSepPen = new Pen(SeparatorColor);
                g.DrawLine(tabSepPen, tabRight - 1, HEADER_HEIGHT + 6, tabRight - 1, HEADER_HEIGHT + TAB_BAR_HEIGHT - 6);
            }

            using var tabFont = new Font("Segoe UI", 8.5f);
            using var tabLabelBrush = new SolidBrush(isActive ? AccentColor : (isHover ? TextColor : DimText));
            var labelSize = g.MeasureString(TabLabels[i], tabFont);
            g.DrawString(TabLabels[i], tabFont, tabLabelBrush,
                tabX + (thisTabW - labelSize.Width) / 2,
                HEADER_HEIGHT + (TAB_BAR_HEIGHT - labelSize.Height) / 2);
        }

        g.DrawLine(sepPen, 0, CONTENT_TOP - 1, w, CONTENT_TOP - 1);

        using var labelFont = new Font("Segoe UI", 9.5f);
        using var labelBrush = new SolidBrush(TextColor);
        using var dimBrush = new SolidBrush(DimText);
        using var valueFont = new Font("Segoe UI Semibold", 10f);
        using var hkFont = new Font("Segoe UI", 8f);
        using var accentBrush = new SolidBrush(AccentColor);

        switch (_activeSettingsTab)
        {
            case 0: DrawGeneralTab(g, w, sepPen, labelFont, labelBrush, dimBrush, valueFont, hkFont, accentBrush); break;
            case 1: DrawQuickTasksTab(g, w, sepPen, labelFont, labelBrush, dimBrush, hkFont, accentBrush); break;
            case 2: DrawCustomTasksTab(g, w, sepPen, labelBrush, dimBrush, hkFont, accentBrush); break;
            case 3: DrawAdvancedTab(g, w, sepPen, labelFont, labelBrush, dimBrush); break;
        }

        using var screenG = Graphics.FromHdc(ps.hdc);
        screenG.DrawImageUnscaled(bmp, 0, 0);
        NativeMethods.EndPaint(hWnd, ref ps);
    }

    private void DrawGeneralTab(Graphics g, int w, Pen sepPen, Font labelFont, SolidBrush labelBrush, SolidBrush dimBrush, Font valueFont, Font hkFont, SolidBrush accentBrush)
    {
        int rowY = CONTENT_TOP;

        // Row 50: AI Model
        if (_hoverRow == 50) { using var rb = new SolidBrush(RowHoverBg); g.FillRectangle(rb, 1, rowY, w - 2, ROW_HEIGHT); }
        g.DrawString("AI model", labelFont, labelBrush, LABEL_X, rowY + 8);
        g.DrawString("Model used for AI tasks", new Font("Segoe UI", 7.5f), dimBrush, LABEL_X, rowY + 26);
        {
            string modelName = _modelIndex < _models.Length ? _models[_modelIndex] : "";
            if (string.IsNullOrEmpty(modelName)) modelName = "(default)";
            string arrow = _modelDropdownOpen ? " \u25B2" : " \u25BC";
            using var pillFont = new Font("Segoe UI", 8.5f);
            var nameSize = g.MeasureString(modelName, pillFont);
            var arrowSize = g.MeasureString(arrow, pillFont);
            int pillW = Math.Min((int)(nameSize.Width + arrowSize.Width + 16), 200);
            int pillX = CONTROL_RIGHT - pillW - 6;
            var pillRect = new Rectangle(pillX, rowY + 10, pillW, 24);
            using var pillPath = CreateRoundedRect(pillRect, 12);
            bool pillHover = _hoverRow == 50 && _hoverButton == 0;
            using (var pillBg = new SolidBrush(pillHover ? ButtonHoverBg : ButtonBg)) g.FillPath(pillBg, pillPath);
            using var pillBrush = new SolidBrush(AccentColor);
            g.DrawString(modelName, pillFont, pillBrush, pillX + 8, rowY + 13);
            using var arrowBrush = new SolidBrush(DimText);
            g.DrawString(arrow, pillFont, arrowBrush, pillX + 8 + nameSize.Width, rowY + 13);
        }
        g.DrawLine(sepPen, 20, rowY + ROW_HEIGHT - 1, w - 20, rowY + ROW_HEIGHT - 1);
        rowY += ROW_HEIGHT;

        // Row 55: System message
        if (_hoverRow == 55) { using var rb = new SolidBrush(RowHoverBg); g.FillRectangle(rb, 1, rowY, w - 2, ROW_HEIGHT); }
        g.DrawString("System message", labelFont, labelBrush, LABEL_X, rowY + 8);
        {
            string preview = string.IsNullOrEmpty(_systemMessage) ? "(default)" : _systemMessage;
            if (preview.Length > 30) preview = preview[..30] + "\u2026";
            using var previewFont = new Font("Segoe UI", 7.5f);
            g.DrawString(preview, previewFont, dimBrush, LABEL_X, rowY + 26);
        }
        DrawPillButton(g, CONTROL_RIGHT - 60, rowY + 10, 50, 24, "Edit", _hoverRow == 55 && _hoverButton == 0);
        g.DrawLine(sepPen, 20, rowY + ROW_HEIGHT - 1, w - 20, rowY + ROW_HEIGHT - 1);
        rowY += ROW_HEIGHT;

        // Row 0: Max history items
        if (_hoverRow == 0) { using var rb = new SolidBrush(RowHoverBg); g.FillRectangle(rb, 1, rowY, w - 2, ROW_HEIGHT); }
        g.DrawString("Max history items", labelFont, labelBrush, LABEL_X, rowY + 8);
        g.DrawString("Number of items to keep", new Font("Segoe UI", 7.5f), dimBrush, LABEL_X, rowY + 26);
        DrawSmallButton(g, CONTROL_RIGHT - 130, rowY + 8, 28, 28, "\u2212", _hoverRow == 0 && _hoverButton == 0);
        var itText = _maxHistoryItems.ToString();
        var itSize = g.MeasureString(itText, valueFont);
        g.DrawString(itText, valueFont, accentBrush, CONTROL_RIGHT - 130 + 28 + (72 - itSize.Width) / 2, rowY + 12);
        DrawSmallButton(g, CONTROL_RIGHT - 30, rowY + 8, 28, 28, "+", _hoverRow == 0 && _hoverButton == 1);
        g.DrawLine(sepPen, 20, rowY + ROW_HEIGHT - 1, w - 20, rowY + ROW_HEIGHT - 1);
        rowY += ROW_HEIGHT;

        // Row 1: Show notification
        if (_hoverRow == 1) { using var rb = new SolidBrush(RowHoverBg); g.FillRectangle(rb, 1, rowY, w - 2, ROW_HEIGHT); }
        g.DrawString("Copy notification bubble", labelFont, labelBrush, LABEL_X, rowY + 8);
        g.DrawString("Show bubble on Ctrl+C", new Font("Segoe UI", 7.5f), dimBrush, LABEL_X, rowY + 26);
        int toggleX = CONTROL_RIGHT - 50, toggleY = rowY + 12;
        const int toggleW = 44, toggleH = 22;
        var trackColor = _showNotification ? AccentColor : (_hoverRow == 1 && _hoverButton == 0 ? Color.FromArgb(80, 80, 80) : Color.FromArgb(60, 60, 60));
        using (var trackBrush = new SolidBrush(trackColor))
        {
            using var trackPath = CreateRoundedRect(new Rectangle(toggleX, toggleY, toggleW, toggleH), toggleH / 2);
            g.FillPath(trackBrush, trackPath);
        }
        int knobX = _showNotification ? toggleX + toggleW - toggleH + 2 : toggleX + 2;
        using (var kn = new SolidBrush(Color.White)) g.FillEllipse(kn, knobX, toggleY + 2, toggleH - 4, toggleH - 4);
        g.DrawLine(sepPen, 20, rowY + ROW_HEIGHT - 1, w - 20, rowY + ROW_HEIGHT - 1);
        rowY += ROW_HEIGHT;

        // Row 2: Notification duration
        if (_hoverRow == 2) { using var rb = new SolidBrush(RowHoverBg); g.FillRectangle(rb, 1, rowY, w - 2, ROW_HEIGHT); }
        g.DrawString("Notification duration", labelFont, labelBrush, LABEL_X, rowY + 8);
        g.DrawString("How long the popup stays visible", new Font("Segoe UI", 7.5f), dimBrush, LABEL_X, rowY + 26);
        DrawSmallButton(g, CONTROL_RIGHT - 130, rowY + 8, 28, 28, "\u2212", _hoverRow == 2 && _hoverButton == 0);
        var durText = $"{_notificationDurationMs / 1000.0:0.#}s";
        var durSize = g.MeasureString(durText, valueFont);
        g.DrawString(durText, valueFont, accentBrush, CONTROL_RIGHT - 130 + 28 + (72 - durSize.Width) / 2, rowY + 12);
        DrawSmallButton(g, CONTROL_RIGHT - 30, rowY + 8, 28, 28, "+", _hoverRow == 2 && _hoverButton == 1);
        g.DrawLine(sepPen, 20, rowY + ROW_HEIGHT - 1, w - 20, rowY + ROW_HEIGHT - 1);
        rowY += ROW_HEIGHT;

        // Row 3: Bubble hotkey
        if (_hoverRow == 3) { using var rb = new SolidBrush(RowHoverBg); g.FillRectangle(rb, 1, rowY, w - 2, ROW_HEIGHT); }
        g.DrawString("Bubble hotkey", labelFont, labelBrush, LABEL_X, rowY + 8);
        g.DrawString("Show notification bubble", new Font("Segoe UI", 7.5f), dimBrush, LABEL_X, rowY + 26);
        DrawHotkeyControls(g, rowY, _bubbleHotkey, 3, hkFont, accentBrush);
        g.DrawLine(sepPen, 20, rowY + ROW_HEIGHT - 1, w - 20, rowY + ROW_HEIGHT - 1);
        rowY += ROW_HEIGHT;

        // Row 4: Clipboard hotkey
        if (_hoverRow == 4) { using var rb = new SolidBrush(RowHoverBg); g.FillRectangle(rb, 1, rowY, w - 2, ROW_HEIGHT); }
        g.DrawString("Clipboard hotkey", labelFont, labelBrush, LABEL_X, rowY + 8);
        g.DrawString("Show clipboard history", new Font("Segoe UI", 7.5f), dimBrush, LABEL_X, rowY + 26);
        DrawHotkeyControls(g, rowY, _hotkey, 4, hkFont, accentBrush);

        // Model dropdown overlay (drawn last so it appears on top)
        if (_modelDropdownOpen)
        {
            const int DROP_ITEM_H = 28;
            int maxVisible = Math.Min(_models.Length, 8);
            int dropH = maxVisible * DROP_ITEM_H + 4;
            int totalH = _models.Length * DROP_ITEM_H + 4;
            int dropW = w - 40;
            int dropX = LABEL_X;
            int dropY = CONTENT_TOP + ROW_HEIGHT + 2;

            // Shadow + background
            using (var shadowBrush = new SolidBrush(Color.FromArgb(120, 0, 0, 0)))
                g.FillRectangle(shadowBrush, dropX + 2, dropY + 2, dropW, dropH);
            var dropRect = new Rectangle(dropX, dropY, dropW, dropH);
            using (var dropBg = new SolidBrush(Color.FromArgb(40, 40, 40)))
            using (var dropPath = CreateRoundedRect(dropRect, 6))
                g.FillPath(dropBg, dropPath);
            using (var dropBorder = new Pen(BorderColor))
            using (var dropPath2 = CreateRoundedRect(dropRect, 6))
                g.DrawPath(dropBorder, dropPath2);

            // Clip to dropdown area
            var oldClip = g.Clip;
            g.SetClip(dropRect);

            using var itemFont = new Font("Segoe UI", 9f);
            for (int i = 0; i < _models.Length; i++)
            {
                int iy = dropY + 2 + i * DROP_ITEM_H - _modelScrollOffset;
                if (iy + DROP_ITEM_H < dropY || iy > dropY + dropH) continue;
                bool isHover = _hoverRow == 51 + i;
                bool isSelected = i == _modelIndex;
                if (isHover || isSelected)
                {
                    var hlColor = isHover ? Color.FromArgb(60, 60, 60) : Color.FromArgb(50, 50, 50);
                    using var hlBrush = new SolidBrush(hlColor);
                    g.FillRectangle(hlBrush, dropX + 2, iy, dropW - 4, DROP_ITEM_H);
                }
                using var itemBrush = new SolidBrush(isSelected ? AccentColor : TextColor);
                g.DrawString(_models[i], itemFont, itemBrush, dropX + 10, iy + 6);
            }

            g.Clip = oldClip;

            // Scrollbar
            if (totalH > dropH)
            {
                int sbX = dropX + dropW - 6;
                int sbTrackH = dropH - 8;
                int sbThumbH = Math.Max(20, sbTrackH * dropH / totalH);
                int sbThumbY = dropY + 4 + (int)((float)_modelScrollOffset / (totalH - dropH) * (sbTrackH - sbThumbH));
                using var sbBrush = new SolidBrush(Color.FromArgb(80, 80, 80));
                g.FillRectangle(sbBrush, sbX, sbThumbY, 4, sbThumbH);
            }
        }
    }

    private void DrawQuickTasksTab(Graphics g, int w, Pen sepPen, Font labelFont, SolidBrush labelBrush, SolidBrush dimBrush, Font hkFont, SolidBrush accentBrush)
    {
        int rowY = CONTENT_TOP;

        // Ask AI hotkey
        if (_hoverRow == 5) { using var rb = new SolidBrush(RowHoverBg); g.FillRectangle(rb, 1, rowY, w - 2, ROW_HEIGHT); }
        g.DrawString("Ask AI hotkey", labelFont, labelBrush, LABEL_X, rowY + 8);
        g.DrawString("Type a freeform AI instruction", new Font("Segoe UI", 7.5f), dimBrush, LABEL_X, rowY + 26);
        DrawHotkeyControls(g, rowY, _freeformHotkey, 5, hkFont, accentBrush);
        g.DrawLine(sepPen, 0, rowY + ROW_HEIGHT, w, rowY + ROW_HEIGHT);

        // Section header
        using var sectionFont = new Font("Segoe UI", 9f);
        g.DrawString("Built-in AI Tasks", sectionFont, dimBrush, LABEL_X, CONTENT_TOP + ROW_HEIGHT + 6);

        // Task hotkey rows
        int taskStartY = CONTENT_TOP + ROW_HEIGHT + 28;
        using var taskFont = new Font("Segoe UI", 9f);
        using var hotkeyFont = new Font("Segoe UI", 8f);
        for (int i = 0; i < 10; i++)
        {
            int ty = taskStartY + i * TASK_ROW_HEIGHT;
            if (_hoverRow == 6 + i) { using var rb = new SolidBrush(RowHoverBg); g.FillRectangle(rb, 1, ty, w - 2, TASK_ROW_HEIGHT); }
            g.DrawString(TaskNames[i], taskFont, labelBrush, LABEL_X, ty + 9);

            var hk = _taskHotkeys[i];
            if (!string.IsNullOrEmpty(hk))
            {
                var hkSize = g.MeasureString(hk, hotkeyFont);
                int clearX = CONTROL_RIGHT - 84;
                g.DrawString(hk, hotkeyFont, accentBrush, clearX - hkSize.Width - 4, ty + 10);
                bool clearHover = _hoverRow == 6 + i && _hoverButton == 1;
                var clearRect = new Rectangle(clearX, ty + 6, 24, 24);
                using var clearPath = CreateRoundedRect(clearRect, 4);
                using (var clearBg = new SolidBrush(clearHover ? ButtonHoverBg : ButtonBg)) g.FillPath(clearBg, clearPath);
                using var clearTxt = new SolidBrush(clearHover ? TextColor : DimText);
                var xSz = g.MeasureString("\u00D7", hotkeyFont);
                g.DrawString("\u00D7", hotkeyFont, clearTxt, clearX + (24 - xSz.Width) / 2, ty + 6 + (24 - xSz.Height) / 2);
            }
            DrawPillButton(g, CONTROL_RIGHT - 60, ty + 6, 50, 24,
                string.IsNullOrEmpty(hk) ? "Set" : "Edit",
                _hoverRow == 6 + i && _hoverButton == 0);

            if (i < 9) g.DrawLine(sepPen, LABEL_X, ty + TASK_ROW_HEIGHT - 1, w - 20, ty + TASK_ROW_HEIGHT - 1);
        }
    }

    private void DrawCustomTasksTab(Graphics g, int w, Pen sepPen, SolidBrush labelBrush, SolidBrush dimBrush, Font hkFont, SolidBrush accentBrush)
    {
        // Add button
        DrawPillButton(g, LABEL_X, CONTENT_TOP + 6, 110, 24, "+ Add task", _hoverRow == 16 && _hoverButton == 0);

        if (_customTasks.Count == 0)
        {
            using var emptyFont = new Font("Segoe UI", 9f);
            g.DrawString("No custom tasks yet. Click \"+ Add task\" to create one.", emptyFont, dimBrush, LABEL_X, CONTENT_TOP + ADD_BUTTON_ROW_HEIGHT + 12);
            return;
        }

        using var taskFont = new Font("Segoe UI", 9f);
        using var hotkeyFont = new Font("Segoe UI", 8f);
        int customStartY = CONTENT_TOP + ADD_BUTTON_ROW_HEIGHT;

        for (int ci = 0; ci < _customTasks.Count; ci++)
        {
            int cy = customStartY + ci * TASK_ROW_HEIGHT;
            if (_hoverRow == 17 + ci) { using var rb = new SolidBrush(RowHoverBg); g.FillRectangle(rb, 1, cy, w - 2, TASK_ROW_HEIGHT); }

            var nameText = string.IsNullOrEmpty(_customTasks[ci].Name) ? "(unnamed)" : _customTasks[ci].Name;
            bool nameHover = _hoverRow == 17 + ci && _hoverButton == 2;
            using (var nameBrush = new SolidBrush(nameHover ? AccentColor : TextColor))
                g.DrawString(nameText, taskFont, nameBrush, LABEL_X, cy + 9);

            // Delete button
            int delX = LABEL_X + 124;
            bool delHover = _hoverRow == 17 + ci && _hoverButton == 3;
            var delRect = new Rectangle(delX, cy + 6, 24, 24);
            using var delPath = CreateRoundedRect(delRect, 4);
            using (var delBg = new SolidBrush(delHover ? Color.FromArgb(180, 60, 60) : ButtonBg)) g.FillPath(delBg, delPath);
            using (var delTxt = new SolidBrush(delHover ? TextColor : DimText))
            {
                var dSz = g.MeasureString("\u00D7", hotkeyFont);
                g.DrawString("\u00D7", hotkeyFont, delTxt, delX + (24 - dSz.Width) / 2, cy + 6 + (24 - dSz.Height) / 2);
            }

            var chk = _customTasks[ci].Hotkey;
            if (!string.IsNullOrEmpty(chk))
            {
                var chkSize = g.MeasureString(chk, hotkeyFont);
                int chkClearX = CONTROL_RIGHT - 84;
                g.DrawString(chk, hotkeyFont, accentBrush, chkClearX - chkSize.Width - 4, cy + 10);
                bool chkClearHover = _hoverRow == 17 + ci && _hoverButton == 1;
                var chkRect = new Rectangle(chkClearX, cy + 6, 24, 24);
                using var chkPath = CreateRoundedRect(chkRect, 4);
                using (var chkBg = new SolidBrush(chkClearHover ? ButtonHoverBg : ButtonBg)) g.FillPath(chkBg, chkPath);
                using var chkTxt = new SolidBrush(chkClearHover ? TextColor : DimText);
                var xSz = g.MeasureString("\u00D7", hotkeyFont);
                g.DrawString("\u00D7", hotkeyFont, chkTxt, chkClearX + (24 - xSz.Width) / 2, cy + 6 + (24 - xSz.Height) / 2);
            }
            DrawPillButton(g, CONTROL_RIGHT - 60, cy + 6, 50, 24,
                string.IsNullOrEmpty(chk) ? "Set" : "Edit",
                _hoverRow == 17 + ci && _hoverButton == 0);

            if (ci < _customTasks.Count - 1)
                g.DrawLine(sepPen, LABEL_X, cy + TASK_ROW_HEIGHT - 1, w - 20, cy + TASK_ROW_HEIGHT - 1);
        }
    }

    private void HoverAdvanced(int mx, int my, ref int newRow, ref int newButton)
    {
        int rowY = CONTENT_TOP;

        // Row 25: Skills toggle
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            newRow = 25;
            if (mx >= CONTROL_RIGHT - 50 && mx < CONTROL_RIGHT - 6 && my >= rowY + 10 && my < rowY + 34) newButton = 0;
            return;
        }
        rowY += ROW_HEIGHT;

        // Row 26: Skills folder link
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            newRow = 26;
            if (mx >= LABEL_X && mx < LABEL_X + 200 && my >= rowY + 8 && my < rowY + 28) newButton = 0;
            return;
        }
        rowY += ROW_HEIGHT;

        // Row 20: MCP toggle
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            newRow = 20;
            if (mx >= CONTROL_RIGHT - 50 && mx < CONTROL_RIGHT - 6 && my >= rowY + 10 && my < rowY + 34) newButton = 0;
            return;
        }
        rowY += ROW_HEIGHT;

        // Row 21: mcp.json link
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            newRow = 21;
            if (mx >= LABEL_X && mx < LABEL_X + 200 && my >= rowY + 8 && my < rowY + 28) newButton = 0;
            return;
        }
        rowY += ROW_HEIGHT;

        // Server toggle rows: row 30+i
        if (_mcpServerNames.Length > 0)
        {
            int serverStartY = rowY + 24;
            for (int i = 0; i < _mcpServerNames.Length; i++)
            {
                int sy = serverStartY + i * TASK_ROW_HEIGHT;
                if (my >= sy && my < sy + TASK_ROW_HEIGHT)
                {
                    newRow = 30 + i;
                    if (mx >= CONTROL_RIGHT - 50 && mx < CONTROL_RIGHT - 6 && my >= sy + 7 && my < sy + 29) newButton = 0;
                    break;
                }
            }
        }
    }

    private void ClickAdvanced(int mx, int my)
    {
        int rowY = CONTENT_TOP;

        // Skills toggle
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            if (mx >= CONTROL_RIGHT - 50 && mx < CONTROL_RIGHT - 6 && my >= rowY + 10 && my < rowY + 34)
            { _skillsEnabled = !_skillsEnabled; SaveSettings(); NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false); }
            return;
        }
        rowY += ROW_HEIGHT;

        // Skills folder link
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            if (mx >= LABEL_X && mx < LABEL_X + 200 && my >= rowY + 8 && my < rowY + 28)
            {
                try { Process.Start(new ProcessStartInfo(AppSettings.SkillsDir) { UseShellExecute = true }); }
                catch { }
            }
            return;
        }
        rowY += ROW_HEIGHT;

        // MCP toggle
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            if (mx >= CONTROL_RIGHT - 50 && mx < CONTROL_RIGHT - 6 && my >= rowY + 10 && my < rowY + 34)
            { _mcpEnabled = !_mcpEnabled; SaveSettings(); NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false); }
            return;
        }
        rowY += ROW_HEIGHT;

        // mcp.json link
        if (my >= rowY && my < rowY + ROW_HEIGHT)
        {
            if (mx >= LABEL_X && mx < LABEL_X + 200 && my >= rowY + 8 && my < rowY + 28)
            {
                try { Process.Start(new ProcessStartInfo(AppSettings.McpJsonFile) { UseShellExecute = true }); }
                catch { }
            }
            return;
        }
        rowY += ROW_HEIGHT;

        // Server toggles
        if (_mcpServerNames.Length > 0)
        {
            int serverStartY = rowY + 24;
            for (int i = 0; i < _mcpServerNames.Length; i++)
            {
                int sy = serverStartY + i * TASK_ROW_HEIGHT;
                if (my >= sy && my < sy + TASK_ROW_HEIGHT)
                {
                    if (mx >= CONTROL_RIGHT - 50 && mx < CONTROL_RIGHT - 6 && my >= sy + 7 && my < sy + 29)
                    {
                        var name = _mcpServerNames[i];
                        if (!_disabledMcpServers.Remove(name))
                            _disabledMcpServers.Add(name);
                        SaveSettings();
                        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                    }
                    return;
                }
            }
        }
    }

    private void DrawAdvancedTab(Graphics g, int w, Pen sepPen, Font labelFont, SolidBrush labelBrush, SolidBrush dimBrush)
    {
        int rowY = CONTENT_TOP;

        // Row 25: Skills toggle
        if (_hoverRow == 25) { using var rb = new SolidBrush(RowHoverBg); g.FillRectangle(rb, 1, rowY, w - 2, ROW_HEIGHT); }
        g.DrawString("Skills", labelFont, labelBrush, LABEL_X, rowY + 8);
        g.DrawString("Load skills from skills folder", new Font("Segoe UI", 7.5f), dimBrush, LABEL_X, rowY + 26);
        DrawToggle(g, CONTROL_RIGHT - 50, rowY + 12, _skillsEnabled, _hoverRow == 25 && _hoverButton == 0);
        g.DrawLine(sepPen, 20, rowY + ROW_HEIGHT - 1, w - 20, rowY + ROW_HEIGHT - 1);
        rowY += ROW_HEIGHT;

        // Row 26: Skills folder link
        if (_hoverRow == 26) { using var rb = new SolidBrush(RowHoverBg); g.FillRectangle(rb, 1, rowY, w - 2, ROW_HEIGHT); }
        g.DrawString("Skills folder", labelFont, labelBrush, LABEL_X, rowY + 8);
        {
            bool skillsLinkHover = _hoverRow == 26 && _hoverButton == 0;
            using var skillsLinkFont = new Font("Segoe UI", 7.5f, skillsLinkHover ? FontStyle.Underline : FontStyle.Regular);
            using var skillsLinkBrush = new SolidBrush(AccentColor);
            g.DrawString("Open skills folder", skillsLinkFont, skillsLinkBrush, LABEL_X, rowY + 26);
        }
        g.DrawLine(sepPen, 20, rowY + ROW_HEIGHT - 1, w - 20, rowY + ROW_HEIGHT - 1);
        rowY += ROW_HEIGHT;

        // Row 20: MCP enabled toggle
        if (_hoverRow == 20) { using var rb = new SolidBrush(RowHoverBg); g.FillRectangle(rb, 1, rowY, w - 2, ROW_HEIGHT); }
        g.DrawString("MCP servers", labelFont, labelBrush, LABEL_X, rowY + 8);
        g.DrawString("Enable Model Context Protocol", new Font("Segoe UI", 7.5f), dimBrush, LABEL_X, rowY + 26);
        DrawToggle(g, CONTROL_RIGHT - 50, rowY + 12, _mcpEnabled, _hoverRow == 20 && _hoverButton == 0);
        g.DrawLine(sepPen, 20, rowY + ROW_HEIGHT - 1, w - 20, rowY + ROW_HEIGHT - 1);
        rowY += ROW_HEIGHT;

        // Row 21: mcp.json link
        if (_hoverRow == 21) { using var rb = new SolidBrush(RowHoverBg); g.FillRectangle(rb, 1, rowY, w - 2, ROW_HEIGHT); }
        g.DrawString("MCP configuration", labelFont, labelBrush, LABEL_X, rowY + 8);
        bool linkHover = _hoverRow == 21 && _hoverButton == 0;
        using var linkFont = new Font("Segoe UI", 7.5f, linkHover ? FontStyle.Underline : FontStyle.Regular);
        using var linkBrush = new SolidBrush(AccentColor);
        g.DrawString("Open mcp.json", linkFont, linkBrush, LABEL_X, rowY + 26);
        g.DrawLine(sepPen, 20, rowY + ROW_HEIGHT - 1, w - 20, rowY + ROW_HEIGHT - 1);
        rowY += ROW_HEIGHT;

        // Server toggles
        if (_mcpServerNames.Length > 0)
        {
            using var sectionFont = new Font("Segoe UI", 9f);
            g.DrawString("Servers", sectionFont, dimBrush, LABEL_X, rowY + 4);
            int serverStartY = rowY + 24;

            using var serverFont = new Font("Segoe UI", 9f);
            for (int i = 0; i < _mcpServerNames.Length; i++)
            {
                int sy = serverStartY + i * TASK_ROW_HEIGHT;
                if (_hoverRow == 30 + i) { using var rb = new SolidBrush(RowHoverBg); g.FillRectangle(rb, 1, sy, w - 2, TASK_ROW_HEIGHT); }
                bool enabled = !_disabledMcpServers.Contains(_mcpServerNames[i]);
                using var nameBrush = new SolidBrush(enabled ? TextColor : DimText);
                g.DrawString(_mcpServerNames[i], serverFont, nameBrush, LABEL_X, sy + 9);
                DrawToggle(g, CONTROL_RIGHT - 50, sy + 7, enabled, _hoverRow == 30 + i && _hoverButton == 0);
                if (i < _mcpServerNames.Length - 1)
                    g.DrawLine(sepPen, LABEL_X, sy + TASK_ROW_HEIGHT - 1, w - 20, sy + TASK_ROW_HEIGHT - 1);
            }
        }
    }

    private void DrawToggle(Graphics g, int x, int y, bool on, bool hover)
    {
        const int toggleW = 44, toggleH = 22;
        var trackColor = on ? AccentColor : (hover ? Color.FromArgb(80, 80, 80) : Color.FromArgb(60, 60, 60));
        using (var trackBrush = new SolidBrush(trackColor))
        {
            using var trackPath = CreateRoundedRect(new Rectangle(x, y, toggleW, toggleH), toggleH / 2);
            g.FillPath(trackBrush, trackPath);
        }
        int knobX = on ? x + toggleW - toggleH + 2 : x + 2;
        using (var kn = new SolidBrush(Color.White)) g.FillEllipse(kn, knobX, y + 2, toggleH - 4, toggleH - 4);
    }

    private static string[] LoadMcpServerNames()
    {
        try
        {
            if (!File.Exists(AppSettings.McpJsonFile)) return [];
            var json = File.ReadAllText(AppSettings.McpJsonFile);
            var doc = JsonDocument.Parse(json);
            if (!doc.RootElement.TryGetProperty("servers", out var servers)) return [];
            return servers.EnumerateObject().Select(p => p.Name).ToArray();
        }
        catch { return []; }
    }

    // Draw hotkey value text + clear button + Set/Edit pill button
    private void DrawHotkeyControls(Graphics g, int rowY, string hotkey, int rowId, Font hkFont, SolidBrush accentBrush)
    {
        if (!string.IsNullOrEmpty(hotkey))
        {
            var hkSize = g.MeasureString(hotkey, hkFont);
            int clearX = CONTROL_RIGHT - 84;
            g.DrawString(hotkey, hkFont, accentBrush, clearX - hkSize.Width - 4, rowY + 14);

            bool clearHover = _hoverRow == rowId && _hoverButton == 1;
            var clearRect = new Rectangle(clearX, rowY + 10, 24, 24);
            using var clearPath = CreateRoundedRect(clearRect, 4);
            using (var clearBg = new SolidBrush(clearHover ? ButtonHoverBg : ButtonBg)) g.FillPath(clearBg, clearPath);
            using var clearTxt = new SolidBrush(clearHover ? TextColor : DimText);
            var xSz = g.MeasureString("\u00D7", hkFont);
            g.DrawString("\u00D7", hkFont, clearTxt, clearX + (24 - xSz.Width) / 2, rowY + 10 + (24 - xSz.Height) / 2);
        }
        DrawPillButton(g, CONTROL_RIGHT - 60, rowY + 10, 50, 24,
            string.IsNullOrEmpty(hotkey) ? "Set" : "Edit",
            _hoverRow == rowId && _hoverButton == 0);
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
        g.DrawString(text, font, textBrush, x + (w - size.Width) / 2, y + (h - size.Height) / 2);
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
        g.DrawString(text, font, textBrush, x + (w - size.Width) / 2, y + (h - size.Height) / 2);
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
