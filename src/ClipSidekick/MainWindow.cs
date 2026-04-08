using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using GitHub.Copilot.SDK;

namespace ClipSidekick;

internal sealed class MainWindow
{
    private const int WINDOW_WIDTH_FALLBACK = 420;
    private const int WINDOW_HEIGHT = 520;
    private const int ITEM_HEIGHT = 56;
    private const int IMAGE_ITEM_HEIGHT = 80;
    private const int TAB_HEIGHT = 40;
    private const int HEADER_HEIGHT = 32;
    private const int TOP_OFFSET = HEADER_HEIGHT + TAB_HEIGHT; // total header + tabs
    private const int PIN_BUTTON_SIZE = 24;
    private const int BUTTON_SIZE = 28;
    private const int BUTTON_MARGIN = 4;
    private const int BORDER_RADIUS = 4;
    private const int RESIZE_BORDER = 6;
    private const int SCROLLBAR_WIDTH = 6;
    private const int SCROLLBAR_MIN_THUMB = 20;
    private const int HOTKEY_ID = 1;
    private const nuint PASTE_TIMER_ID = 100;

    private IntPtr _hWnd;
    private IntPtr _hInstance;
    private NativeMethods.WndProc _wndProc = null!;
    private readonly ClipboardMonitor _monitor;
    private bool _alwaysOnTop;
    private int _scrollOffset;
    private bool _scrollbarHover;
    private bool _scrollbarDragging;
    private int _scrollbarDragStartY;
    private int _scrollbarDragStartOffset;
    private int _hoverIndex = -1;
    private int _hoverButton = -1; // normal: 0=paste, 1=AI, 2=more; expanded: 0=bookmark, 1=delete
    private bool _tracking;
    private int _expandedIndex = -1; // item showing bookmark/delete buttons
    private bool _visible;
    private int _windowWidth = WINDOW_WIDTH_FALLBACK;
    private int _windowHeight = WINDOW_HEIGHT;
    private IntPtr _previousForegroundWindow;
    private string? _extractedText;
    private IntPtr _extractPasteTarget;
    private NotificationPopup _notification = null!;
    private AppSettings _settings;
    private Icon? _appIconSmall;
    private Icon? _appIconLarge;
    private IntPtr _mouseHook;
    private NativeMethods.LowLevelMouseProc _mouseHookProc = null!;
    private int _activeTab; // 0=history, 1=bookmarks, 2=edit
    private int _hoverTab = -1;
    private bool _pinHover;
    private bool _closeHover;
    private string? _tooltipText;
    private Point _tooltipPos;

    // Edit tab state
    private ClipboardItem? _editItem;
    private int _editTaskIndex;    // 0=<unset>, then Rewrite..Summarize
    private int _editToneIndex;    // 0=<unset>, then Professional..Funny
    private int _editFormatIndex;  // 0=<unset>, then Single paragraph..JSON
    private int _editLengthIndex;  // 0=<unset>, then Headline..Verbose
    private int _editChoices = 2;   // 1..5 choices slider
    private int _editModelIndex;   // index into _editModels
    private bool _editModelDropdownOpen;
    private Rectangle _editModelDropdownRect; // stored during drawing
    private bool _editChoicesDragging;
    private Rectangle _editChoicesSliderRect; // stored during drawing
    private int _editHoverElement = -1; // which edit UI element is hovered
    private int _editTextScrollOffset; // scroll offset for edit text box
    private bool _editScrollbarDragging;
    private int _editScrollbarDragStartY;
    private int _editScrollbarDragStartOffset;
    private bool _aiResultScrollbarDragging;
    private int _aiResultScrollbarDragIndex;
    private int _aiResultScrollbarDragStartY;
    private int _aiResultScrollbarDragStartOffset;
    private Rectangle _editTextBoxRect; // stored during drawing for scroll calculations
    private bool _editTaskDropdownOpen;
    private bool _editToneDropdownOpen;
    private bool _editFormatDropdownOpen;
    // Pill rects for hover detection (set during drawing)
    private Rectangle _pillTaskRect, _pillToneRect, _pillFormatRect, _pillLengthRect;
    private bool _editLengthDropdownOpen;

    private static readonly string[] EditTasks = ["<unset>", "Proofread", "Rewrite", "Use synonyms", "Minor revise", "Major revise", "Describe", "Answer", "Explain", "Expand", "Summarize"];
    private static readonly string[] EditTaskIcons = ["", "\u270f\ufe0f", "\ud83d\udc41", "\ud83d\udcda", "\ud83d\udd27", "\ud83d\udd28", "\ud83d\udcdd", "\ud83d\udcac", "\ud83d\udca1", "\ud83c\udf1f", "\ud83d\udccb"];
    private static readonly string[] EditTones = ["<unset>", "Professional", "Casual", "Enthusiastic", "Informational", "Confident", "Technical", "Funny"];
    private static readonly string[] EditToneIcons = ["", "\ud83d\udcbc", "\ud83d\ude0a", "\ud83e\udd29", "\u2139\ufe0f", "\ud83e\udd19", "\u2699\ufe0f", "\ud83e\udd20"];
    private static readonly string[] EditFormats = ["<unset>", "Single paragraph", "Paragraphs with line breaks", "List", "Ordered list", "Table", "Task list", "Headings", "Blockquotes", "Code blocks", "Emojis", "HTML", "JSON"];
    private static readonly string[] EditFormatIcons = ["", "\u00b6", "\u21b5", "\u2022", "#", "\u2261", "\u2611\ufe0f", "H", "\u201c", "</>", "\ud83d\ude00", "\ud83c\udf10", "{}"];
    private static readonly string[] EditLengths = ["<unset>", "Headline", "Minimal", "Tight", "Normal", "Verbose"];
    private static readonly string[] EditLengthIcons = ["", "\ud83d\udccc", "\u2702\ufe0f", "\ud83d\udccf", "\ud83d\udcdd", "\ud83d\udcd6"];
    private string[] _editModels = [""];

    // Copilot SDK
    private CopilotClient? _copilotClient;
    private CopilotSession? _copilotSession;

    // AI processing state
    private bool _aiProcessing;
    private CancellationTokenSource? _aiCancellation;
    private List<string> _aiResults = [];
    private int[] _aiResultScrollOffsets = [];
    private Rectangle _editSendButtonRect;
    private Rectangle[] _aiResultTextBoxRects = [];
    private Rectangle[] _aiResultButtonRects = []; // flattened: [result0_btn0, result0_btn1, result0_btn2, result1_btn0, ...]
    private float _aiAnimPhase; // 0..1 animation phase for edge glow
    private const nuint ANIM_TIMER_ID = 200;
    private const nuint TRAY_ANIM_TIMER_ID = 300;
    private Icon?[] _trayAnimFrames = [];
    private int _trayAnimFrame;

    // Inline AI processing state (from bubble task dropdown)
    private int _inlineAiItemIndex = -1; // index of item being AI-processed inline
    private bool _inlineAiProcessing;
    private CancellationTokenSource? _inlineAiCancellation;
    private bool _bubbleTaskDropdownOpen;
    private int _bubbleTaskDropdownItemIndex = -1;
    private int _bubbleTaskHoverIndex = -1; // hover index within task dropdown
    private bool _inlineAiFromNotification; // true when triggered from notification popup

    // Colors - modern dark theme
    private static readonly Color BgColor = Color.FromArgb(32, 32, 32);
    private static readonly Color HeaderBg = Color.FromArgb(40, 40, 40);
    private static readonly Color ItemBg = Color.FromArgb(45, 45, 45);
    private static readonly Color ItemHoverBg = Color.FromArgb(55, 55, 55);
    private static readonly Color TextColor = Color.FromArgb(230, 230, 230);
    private static readonly Color SecondaryText = Color.FromArgb(150, 150, 150);
    private static readonly Color AccentColor = Color.FromArgb(96, 165, 250);
    private static readonly Color ButtonHover = Color.FromArgb(70, 70, 70);
    private static readonly Color PinColor = Color.FromArgb(250, 204, 21);
    private static readonly Color BorderColor = Color.FromArgb(60, 60, 60);
    private static readonly Color SeparatorColor = Color.FromArgb(55, 55, 55);
    private static readonly Color ScrollbarColor = Color.FromArgb(80, 80, 80);
    private static readonly Color ScrollbarHoverColor = Color.FromArgb(110, 110, 110);

    // Tray icon
    private NativeMethods.NOTIFYICONDATA _trayIconData;
    private const int TRAY_ID = 1;
    private const int TRAY_MENU_SHOW = 1001;
    private const int TRAY_MENU_ONTOP = 1002;
    private const int TRAY_MENU_CLEAR = 1003;
    private const int TRAY_MENU_SETTINGS = 1005;
    private const int TRAY_MENU_EXIT = 1004;

    public MainWindow(ClipboardMonitor monitor)
    {
        _monitor = monitor;
        _monitor.ClipboardChanged += OnClipboardChanged;
        _settings = AppSettings.Load();
        _monitor.MaxItems = _settings.MaxHistoryItems;

        // Restore edit dropdown settings
        _editTaskIndex = Math.Clamp(_settings.EditTaskIndex, 0, EditTasks.Length - 1);
        _editToneIndex = Math.Clamp(_settings.EditToneIndex, 0, EditTones.Length - 1);
        _editFormatIndex = Math.Clamp(_settings.EditFormatIndex, 0, EditFormats.Length - 1);
        _editLengthIndex = Math.Clamp(_settings.EditLengthIndex, 0, EditLengths.Length - 1);
        _editChoices = Math.Clamp(_settings.EditChoices, 1, 5);
        _editModelIndex = Math.Max(0, Array.IndexOf(_editModels, _settings.Model));
    }

    private IReadOnlyList<ClipboardItem> ActiveItems =>
        _activeTab == 0 ? _monitor.Items : _monitor.BookmarkedItems;

    public void Create()
    {
        _hInstance = NativeMethods.GetModuleHandleW(IntPtr.Zero);
        _wndProc = WndProc;

        var className = "ClipSidekickWindow";
        var classNamePtr = Marshal.StringToHGlobalUni(className);

        var wc = new NativeMethods.WNDCLASSEX
        {
            cbSize = Marshal.SizeOf<NativeMethods.WNDCLASSEX>(),
            style = 0,
            lpfnWndProc = Marshal.GetFunctionPointerForDelegate(_wndProc),
            hInstance = _hInstance,
            hCursor = NativeMethods.LoadCursorW(IntPtr.Zero, NativeMethods.IDC_ARROW),
            hbrBackground = IntPtr.Zero,
            lpszClassName = classNamePtr
        };

        NativeMethods.RegisterClassExW(ref wc);

        _hWnd = NativeMethods.CreateWindowExW(
            NativeMethods.WS_EX_TOOLWINDOW | NativeMethods.WS_EX_LAYERED | NativeMethods.WS_EX_NOACTIVATE,
            className,
            "Clip Sidekick",
            NativeMethods.WS_POPUP | NativeMethods.WS_THICKFRAME,
            0, 0, WINDOW_WIDTH_FALLBACK, WINDOW_HEIGHT,
            IntPtr.Zero, IntPtr.Zero, _hInstance, IntPtr.Zero);

        // Set rounded corners on Win11
        int pref = NativeMethods.DWMWCP_ROUND;
        NativeMethods.DwmSetWindowAttribute(_hWnd, NativeMethods.DWMWA_WINDOW_CORNER_PREFERENCE, ref pref, sizeof(int));

        // Load embedded app icon
        LoadAppIcon();

        // Make layered window fully opaque (needed for WS_EX_LAYERED to prevent flicker during creation)
        SetLayeredWindowAttributes(_hWnd, 0, 255, 0x02);

        // Register for clipboard updates
        NativeMethods.AddClipboardFormatListener(_hWnd);

        // Register configurable hotkey
        RegisterConfiguredHotkey();

        // Add tray icon
        AddTrayIcon();

        // Create notification popup
        _notification = new NotificationPopup();
        _notification.DurationMs = _settings.NotificationDurationMs;
        _notification.Create(_hInstance);
        _notification.Clicked += () =>
            NativeMethods.PostMessageW(_hWnd, NativeMethods.WM_APP_SHOW, IntPtr.Zero, IntPtr.Zero);
        _notification.AIEditClicked += () =>
            NativeMethods.PostMessageW(_hWnd, NativeMethods.WM_APP_AIEDIT, IntPtr.Zero, IntPtr.Zero);
        _notification.QuickTaskClicked += (taskIdx) =>
            NativeMethods.PostMessageW(_hWnd, NativeMethods.WM_APP_QUICKTASK, new IntPtr(taskIdx), IntPtr.Zero);
        _notification.CancelRequested += () =>
        {
            if (_inlineAiFromNotification)
                CancelInlineAiRequest();
        };

        Marshal.FreeHGlobal(classNamePtr);

        // Measure pill row to set fixed window width
        MeasureWindowWidth();

        // Initialize Copilot SDK client and session
        InitCopilotAsync();

        // Auto-show on startup
        Show();
    }

    [DllImport("user32.dll")]
    private static extern bool SetLayeredWindowAttributes(IntPtr hwnd, uint crKey, byte bAlpha, uint dwFlags);

    public void Run()
    {
        while (NativeMethods.GetMessageW(out var msg, IntPtr.Zero, 0, 0))
        {
            NativeMethods.TranslateMessage(ref msg);
            NativeMethods.DispatchMessageW(ref msg);
        }
    }

    public void Cleanup()
    {
        RemoveMouseHook();
        NativeMethods.RemoveClipboardFormatListener(_hWnd);
        NativeMethods.UnregisterHotKey(_hWnd, HOTKEY_ID);
        RemoveTrayIcon();
        // Dispose Copilot SDK
        _aiCancellation?.Cancel();
        NativeMethods.KillTimer(_hWnd, ANIM_TIMER_ID);
        try { _copilotSession?.DisposeAsync().AsTask().Wait(TimeSpan.FromSeconds(3)); } catch { }
        try { _copilotClient?.DisposeAsync().AsTask().Wait(TimeSpan.FromSeconds(3)); } catch { }
    }

    private async void InitCopilotAsync()
    {
        try
        {
            _copilotClient = new CopilotClient();

            _copilotSession = await _copilotClient.CreateSessionAsync(new SessionConfig
            {
                OnPermissionRequest = PermissionHandler.ApproveAll,
                Model = _editModels[_editModelIndex],
            });

            // Load available models from the SDK (after session is ready)
            var models = await _copilotClient.ListModelsAsync();
            if (models is { Count: > 0 })
            {
                _editModels = models.Select(m => m.Id).ToArray();
                _editModelIndex = Math.Max(0, Array.IndexOf(_editModels, _settings.Model));
            }

            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        }
        catch
        {
            // Copilot CLI not available — features will be disabled
        }
    }

    private async void RecreateCopilotSession()
    {
        try
        {
            if (_copilotSession != null)
                await _copilotSession.DisposeAsync();
            if (_copilotClient != null)
            {
                _copilotSession = await _copilotClient.CreateSessionAsync(new SessionConfig
                {
                    OnPermissionRequest = PermissionHandler.ApproveAll,
                    Model = _editModels[_editModelIndex],
                });
            }
        }
        catch { }
    }

    private async void SendAiRequest()
    {
        if (_copilotSession == null || _editItem == null || _aiProcessing)
            return;

        _aiProcessing = true;
        _aiResults.Clear();
        _aiAnimPhase = 0;
        NativeMethods.SetTimer(_hWnd, ANIM_TIMER_ID, 30, IntPtr.Zero);
        StartTrayAnimation();
        _aiCancellation = new CancellationTokenSource();
        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);

        try
        {
            var ct = _aiCancellation.Token;
            string inputText = _editItem.Text;

            for (int i = 0; i < _editChoices && !ct.IsCancellationRequested; i++)
            {
                string prompt = ComposePrompt(inputText, i + 1);
                var response = await _copilotSession.SendAndWaitAsync(
                    new MessageOptions { Prompt = prompt }, null, ct);
                if (response?.Data?.Content is { } content && content.Length > 0)
                {
                    _aiResults.Add(content);
                    _aiResultScrollOffsets = new int[_aiResults.Count];
                    AdjustWindowHeightForResults();
                    NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                }
            }
        }
        catch (OperationCanceledException) { }
        catch { }
        finally
        {
            _aiProcessing = false;
            _aiCancellation?.Dispose();
            _aiCancellation = null;
            NativeMethods.KillTimer(_hWnd, ANIM_TIMER_ID);
            StopTrayAnimation();
            _aiResultScrollOffsets = new int[_aiResults.Count];
            AdjustWindowHeightForResults();
            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        }
    }

    private void CancelAiRequest()
    {
        _aiCancellation?.Cancel();
        try { _copilotSession?.AbortAsync().GetAwaiter().GetResult(); } catch { }
    }

    private async void StartInlineAiRequest(int itemIndex, int taskIndex)
    {
        if (_copilotSession == null || _inlineAiProcessing)
            return;

        // Always use history items (not bookmarks) — notification popup triggers use index 0 of history
        var items = _inlineAiFromNotification ? _monitor.Items : ActiveItems;
        if (itemIndex < 0 || itemIndex >= items.Count)
            return;

        var item = items[itemIndex];
        string inputText = item.Type == ClipboardItemType.Image ? item.PreviewText : item.Text;
        if (string.IsNullOrWhiteSpace(inputText))
            return;

        _inlineAiProcessing = true;
        _inlineAiItemIndex = itemIndex;
        _aiAnimPhase = 0;
        NativeMethods.SetTimer(_hWnd, ANIM_TIMER_ID, 30, IntPtr.Zero);
        StartTrayAnimation();
        _inlineAiCancellation = new CancellationTokenSource();
        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);

        try
        {
            var ct = _inlineAiCancellation.Token;
            string prompt = $"Task: {EditTasks[taskIndex]}.\n\nText to edit:\n\n{inputText}\n\nRespond ONLY with the edited text. No explanations, no markdown fences, no preamble.";
            var response = await _copilotSession.SendAndWaitAsync(
                new MessageOptions { Prompt = prompt }, null, ct);
            if (response?.Data?.Content is { } content && content.Length > 0 && !ct.IsCancellationRequested)
            {
                _monitor.SetIgnoreNext();
                SetClipboardText(content);
                SchedulePaste();
            }
        }
        catch (OperationCanceledException) { }
        catch { }
        finally
        {
            _inlineAiProcessing = false;
            _inlineAiItemIndex = -1;
            _inlineAiCancellation?.Dispose();
            _inlineAiCancellation = null;
            NativeMethods.KillTimer(_hWnd, ANIM_TIMER_ID);
            StopTrayAnimation();
            if (_inlineAiFromNotification)
            {
                _inlineAiFromNotification = false;
                _notification.StopProcessing();
            }
            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        }
    }

    private void CancelInlineAiRequest()
    {
        _inlineAiCancellation?.Cancel();
        try { _copilotSession?.AbortAsync().GetAwaiter().GetResult(); } catch { }
    }

    private string ComposePrompt(string inputText, int choiceNumber)
    {
        var parts = new List<string>();

        if (_editTaskIndex > 0)
            parts.Add($"Task: {EditTasks[_editTaskIndex]}.");
        if (_editToneIndex > 0)
            parts.Add($"Tone: {EditTones[_editToneIndex]}.");
        if (_editFormatIndex > 0)
            parts.Add($"Format: {EditFormats[_editFormatIndex]}.");
        if (_editLengthIndex > 0)
            parts.Add($"Length: {EditLengths[_editLengthIndex]}.");

        if (_editChoices > 1)
            parts.Add($"This is variation {choiceNumber} of {_editChoices}. Make each variation meaningfully different.");

        parts.Add($"Text to edit:\n\n{inputText}");
        parts.Add("Respond ONLY with the edited text. No explanations, no markdown fences, no preamble.");

        return string.Join("\n", parts);
    }

    private void AdjustWindowHeightForResults()
    {
        if (_aiResults.Count > 0)
        {
            // Measure actual content height needed
            int contentH = TOP_OFFSET + 1; // top offset
            int pad = 16;
            contentH += pad; // top padding

            // Pills row
            contentH += 28 + 12; // pillHeight + gap

            // Textbox: measure actual text height
            int cw = _windowWidth - pad * 2 - 16; // approx client width minus padding and scrollbar
            float lineH;
            using (var measG = Graphics.FromHwnd(_hWnd))
            using (var labelFont = new Font("Segoe UI", 9f))
            using (var sf = new StringFormat(StringFormat.GenericTypographic))
            {
                sf.FormatFlags = 0;
                lineH = measG.MeasureString("Ag", labelFont).Height;
                string fullText = _editItem?.Text ?? "";
                int txtW = cw - 20 - 10; // textPadX*2 + scrollbar
                var textSize = measG.MeasureString(fullText, labelFont, new SizeF(txtW, 999999), sf);
                int totalTextH = Math.Max((int)Math.Ceiling(lineH), (int)Math.Ceiling(textSize.Height));
                int innerH = Math.Clamp(totalTextH, (int)Math.Ceiling(lineH), (int)Math.Ceiling(lineH * 4));
                contentH += innerH + 16 + 12; // textPadY*2 + gap

                // Model + slider + send button row
                contentH += 28 + 8; // rowH + gap

                // Results
                foreach (var result in _aiResults)
                {
                    var rs = measG.MeasureString(result, labelFont, new SizeF(txtW, 999999), sf);
                    int rMinH = (int)Math.Ceiling(lineH * 1.5);
                    int rTotalH = Math.Max(rMinH, (int)Math.Ceiling(rs.Height));
                    int rInnerH = Math.Clamp(rTotalH, rMinH, (int)Math.Ceiling(lineH * 4));
                    int boxH = rInnerH + 16; // textPadY*2
                    contentH += boxH + 10; // box + gap (buttons are beside textbox now)
                }
            }

            contentH += pad; // bottom padding
            int targetH = Math.Clamp(contentH, WINDOW_HEIGHT, 900);
            NativeMethods.GetWindowRect(_hWnd, out var wr);
            NativeMethods.SetWindowPos(_hWnd, IntPtr.Zero, wr.Left, wr.Top, _windowWidth, targetH,
                NativeMethods.SWP_NOACTIVATE | NativeMethods.SWP_NOZORDER);
        }
        else if (_windowHeight > WINDOW_HEIGHT)
        {
            NativeMethods.GetWindowRect(_hWnd, out var wr);
            NativeMethods.SetWindowPos(_hWnd, IntPtr.Zero, wr.Left, wr.Top, _windowWidth, WINDOW_HEIGHT,
                NativeMethods.SWP_NOACTIVATE | NativeMethods.SWP_NOZORDER);
        }
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
        if (nCode >= 0 && _visible)
        {
            int wmsg = wParam.ToInt32();
            if (wmsg == NativeMethods.WM_LBUTTONDOWN)
            {
                var hookStruct = Marshal.PtrToStructure<NativeMethods.MSLLHOOKSTRUCT>(lParam);
                NativeMethods.GetWindowRect(_hWnd, out var rc);
                if (hookStruct.pt.X < rc.Left || hookStruct.pt.X > rc.Right ||
                    hookStruct.pt.Y < rc.Top || hookStruct.pt.Y > rc.Bottom)
                {
                    NativeMethods.PostMessageW(_hWnd, NativeMethods.WM_APP_DISMISS, IntPtr.Zero, IntPtr.Zero);
                }
            }
            else if (wmsg == NativeMethods.WM_MOUSEWHEEL)
            {
                var hookStruct = Marshal.PtrToStructure<NativeMethods.MSLLHOOKSTRUCT>(lParam);
                NativeMethods.GetWindowRect(_hWnd, out var rc);
                if (hookStruct.pt.X >= rc.Left && hookStruct.pt.X <= rc.Right &&
                    hookStruct.pt.Y >= rc.Top && hookStruct.pt.Y <= rc.Bottom)
                {
                    // Extract wheel delta from mouseData high word
                    short delta = (short)(hookStruct.mouseData >> 16);
                    if (_activeTab == 2)
                    {
                        int step = Math.Max(8, Math.Abs(delta) / 3);
                        // Check if mouse is over a result textbox
                        int localX = hookStruct.pt.X - rc.Left;
                        int localY = hookStruct.pt.Y - rc.Top;
                        bool scrolledResult = false;
                        for (int ri = 0; ri < _aiResultTextBoxRects.Length; ri++)
                        {
                            if (_aiResultTextBoxRects[ri].Contains(localX, localY) && ri < _aiResultScrollOffsets.Length)
                            {
                                _aiResultScrollOffsets[ri] += delta < 0 ? step : -step;
                                if (_aiResultScrollOffsets[ri] < 0) _aiResultScrollOffsets[ri] = 0;
                                scrolledResult = true;
                                break;
                            }
                        }
                        if (!scrolledResult)
                        {
                            _editTextScrollOffset += delta < 0 ? step : -step;
                            if (_editTextScrollOffset < 0) _editTextScrollOffset = 0;
                        }
                    }
                    else
                    {
                        NativeMethods.GetClientRect(_hWnd, out var cr);
                        int clientHeight = cr.Bottom - cr.Top;
                        int maxScroll = Math.Max(0, GetTotalContentHeight() - (clientHeight - TOP_OFFSET - 1));
                        _scrollOffset = Math.Clamp(_scrollOffset - delta, 0, maxScroll);
                    }
                    NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                }
            }
        }
        return NativeMethods.CallNextHookEx(_mouseHook, nCode, wParam, lParam);
    }

    private void ToggleWindow()
    {
        if (_visible)
        {
            Hide();
        }
        else
        {
            Show();
        }
    }

    private void Show()
    {
        // _previousForegroundWindow may already be set by the hotkey handler
        if (_previousForegroundWindow == IntPtr.Zero || _previousForegroundWindow == _hWnd)
            _previousForegroundWindow = NativeMethods.GetForegroundWindow();

        PositionWindowAtCursor();
        _scrollOffset = 0;
        _hoverIndex = -1;
        _hoverButton = -1;

        var insertAfter = _alwaysOnTop ? NativeMethods.HWND_TOPMOST : NativeMethods.HWND_TOPMOST;
        NativeMethods.SetWindowPos(_hWnd, insertAfter, 0, 0, 0, 0,
            NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_SHOWWINDOW | NativeMethods.SWP_NOACTIVATE);
        NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_SHOWNOACTIVATE);
        _visible = true;
        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);

        InstallMouseHook();
    }

    private void Hide()
    {
        RemoveMouseHook();
        NativeMethods.ShowWindow(_hWnd, NativeMethods.SW_HIDE);
        _visible = false;
        _bubbleTaskDropdownOpen = false;
        _bubbleTaskDropdownItemIndex = -1;
        _bubbleTaskHoverIndex = -1;
    }

    private void PositionWindowAtCursor()
    {
        var anchor = new NativeMethods.POINT();
        bool usedCaret = false;

        // Strategy 1: UI Automation (works for VS Code, Chrome, WinUI, modern apps)
        if (CaretLocator.TryGetCaretPosition(out int cx, out int cy))
        {
            anchor.X = cx;
            anchor.Y = cy;
            usedCaret = true;
        }

        // Strategy 2: Win32 GetGUIThreadInfo (classic Win32 apps like old Notepad)
        if (!usedCaret && _previousForegroundWindow != IntPtr.Zero && _previousForegroundWindow != _hWnd)
        {
            uint targetThread = NativeMethods.GetWindowThreadProcessId(_previousForegroundWindow, out _);
            if (targetThread != 0)
            {
                var gti = new NativeMethods.GUITHREADINFO { cbSize = Marshal.SizeOf<NativeMethods.GUITHREADINFO>() };
                if (NativeMethods.GetGUIThreadInfo(targetThread, ref gti) && gti.hwndCaret != IntPtr.Zero)
                {
                    var pt = new NativeMethods.POINT { X = gti.rcCaret.Left, Y = gti.rcCaret.Bottom };
                    NativeMethods.ClientToScreen(gti.hwndCaret, ref pt);
                    anchor = pt;
                    usedCaret = true;
                }
            }
        }

        // Fallback: mouse cursor
        if (!usedCaret)
            NativeMethods.GetCursorPos(out anchor);

        var monitor = NativeMethods.MonitorFromPoint(anchor, NativeMethods.MONITOR_DEFAULTTONEAREST);
        var mi = new NativeMethods.MONITORINFO { cbSize = Marshal.SizeOf<NativeMethods.MONITORINFO>() };
        NativeMethods.GetMonitorInfoW(monitor, ref mi);

        var workArea = mi.rcWork;
        int x = anchor.X - _windowWidth / 2;
        int y = anchor.Y - _windowHeight - 10;

        if (y < workArea.Top)
            y = anchor.Y + 20;

        x = Math.Clamp(x, workArea.Left, workArea.Right - _windowWidth);
        y = Math.Clamp(y, workArea.Top, workArea.Bottom - _windowHeight);

        NativeMethods.SetWindowPos(_hWnd, IntPtr.Zero, x, y, _windowWidth, _windowHeight,
            NativeMethods.SWP_NOACTIVATE | NativeMethods.SWP_NOZORDER);
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

            case NativeMethods.WM_CLIPBOARDUPDATE:
                _monitor.OnClipboardUpdate();
                return IntPtr.Zero;

            case NativeMethods.WM_HOTKEY:
                if (wParam.ToInt32() == HOTKEY_ID)
                {
                    if (!_visible)
                        _previousForegroundWindow = NativeMethods.GetForegroundWindow();
                    ToggleWindow();
                }
                return IntPtr.Zero;

            case NativeMethods.WM_MOUSEWHEEL:
                // Handled by the low-level mouse hook (MouseHookCallback)
                return IntPtr.Zero;

            case NativeMethods.WM_MOUSEMOVE:
                OnMouseMove(lParam);
                return IntPtr.Zero;

            case NativeMethods.WM_LBUTTONDOWN:
                OnMouseDown(lParam);
                return IntPtr.Zero;

            case 0x0202: // WM_LBUTTONUP
                if (_scrollbarDragging || _editScrollbarDragging || _editChoicesDragging || _aiResultScrollbarDragging)
                {
                    if (_editChoicesDragging)
                    {
                        _settings.EditChoices = _editChoices;
                        _settings.Save();
                    }
                    _scrollbarDragging = false;
                    _editScrollbarDragging = false;
                    _aiResultScrollbarDragging = false;
                    _editChoicesDragging = false;
                    NativeMethods.ReleaseCapture();
                    NativeMethods.InvalidateRect(hWnd, IntPtr.Zero, false);
                }
                return IntPtr.Zero;

            case NativeMethods.WM_MOUSELEAVE:
                _tracking = false;
                _scrollbarHover = false;
                _tooltipText = null;
                if (_hoverIndex != -1 || _hoverButton != -1 || _hoverTab != -1)
                {
                    _hoverIndex = -1;
                    _hoverButton = -1;
                    _hoverTab = -1;
                    NativeMethods.InvalidateRect(hWnd, IntPtr.Zero, false);
                }
                return IntPtr.Zero;

            case NativeMethods.WM_ACTIVATE:
                return IntPtr.Zero;

            case NativeMethods.WM_MOUSEACTIVATE:
                return new IntPtr(3); // MA_NOACTIVATE

            case NativeMethods.WM_APP_DISMISS:
                if (_visible && !_alwaysOnTop)
                    Hide();
                return IntPtr.Zero;

            case NativeMethods.WM_APP_SHOW:
                _activeTab = 0;
                Show();
                return IntPtr.Zero;

            case NativeMethods.WM_APP_AIEDIT:
                var latest = _monitor.Items.FirstOrDefault();
                if (latest != null)
                {
                    _editItem = latest;
                    _editTextScrollOffset = 0;
                    _editTaskDropdownOpen = false;
                    _editToneDropdownOpen = false;
                    _editFormatDropdownOpen = false;
                    _editLengthDropdownOpen = false;
                    _editModelDropdownOpen = false;
                    _aiResults.Clear();
                    _aiResultScrollOffsets = [];
                    _activeTab = 2;
                    AdjustWindowHeightForResults();
                    Show();
                }
                return IntPtr.Zero;

            case NativeMethods.WM_APP_QUICKTASK:
                int quickTaskIdx = wParam.ToInt32();
                if (quickTaskIdx > 0 && quickTaskIdx < EditTasks.Length && _monitor.Items.Count > 0)
                {
                    _inlineAiFromNotification = true;
                    StartInlineAiRequest(0, quickTaskIdx);
                }
                return IntPtr.Zero;

            case NativeMethods.WM_APP_EXTRACTPASTE:
                if (_extractedText != null)
                {
                    _previousForegroundWindow = _extractPasteTarget;
                    _monitor.SetIgnoreNext();
                    SetClipboardText(_extractedText);
                    _extractedText = null;
                    ForceForegroundWindow(_extractPasteTarget);
                    SchedulePaste();
                }
                return IntPtr.Zero;

            case NativeMethods.WM_KEYDOWN:
                if (wParam.ToInt32() == NativeMethods.VK_ESCAPE && _visible)
                {
                    if (_aiProcessing)
                    {
                        CancelAiRequest();
                    }
                    else if (_activeTab == 2)
                    {
                        _activeTab = 0;
                        _editItem = null;
                        _aiResults.Clear();
                        _aiResultScrollOffsets = [];
                        AdjustWindowHeightForResults();
                        NativeMethods.InvalidateRect(hWnd, IntPtr.Zero, false);
                    }
                    else
                        Hide();
                }
                return IntPtr.Zero;

            case NativeMethods.WM_TIMER:
                if ((nuint)wParam.ToInt64() == PASTE_TIMER_ID)
                {
                    NativeMethods.KillTimer(hWnd, PASTE_TIMER_ID);
                    ReleaseModifierKeys();
                    SimulatePaste();
                }
                else if ((nuint)wParam.ToInt64() == ANIM_TIMER_ID)
                {
                    _aiAnimPhase = (_aiAnimPhase + 0.04f) % 1f;
                    NativeMethods.InvalidateRect(hWnd, IntPtr.Zero, false);
                }
                else if ((nuint)wParam.ToInt64() == TRAY_ANIM_TIMER_ID)
                {
                    if (_trayAnimFrames.Length > 0)
                    {
                        _trayAnimFrame = (_trayAnimFrame + 1) % _trayAnimFrames.Length;
                        var frame = _trayAnimFrames[_trayAnimFrame];
                        if (frame != null)
                        {
                            _trayIconData.hIcon = frame.Handle;
                            _trayIconData.uFlags = NativeMethods.NIF_ICON;
                            NativeMethods.Shell_NotifyIconW(NativeMethods.NIM_MODIFY, ref _trayIconData);
                        }
                    }
                }
                return IntPtr.Zero;

            case NativeMethods.WM_NCHITTEST:
                return OnNcHitTest(hWnd, lParam);

            case 0x00A1: // WM_NCLBUTTONDOWN
                if (wParam.ToInt32() == NativeMethods.HTCAPTION)
                {
                    // Check if click is on a tab — handle tab switch instead of drag
                    int sx = NativeMethods.GET_X_LPARAM(lParam);
                    int sy = NativeMethods.GET_Y_LPARAM(lParam);
                    NativeMethods.GetWindowRect(hWnd, out var wr);
                    int relMx = sx - wr.Left;
                    int relMy = sy - wr.Top;
                    if (relMy >= HEADER_HEIGHT && relMy < TOP_OFFSET)
                    {
                        NativeMethods.GetClientRect(hWnd, out var ccr);
                        int cw = ccr.Right - ccr.Left;
                        int tabCount = 3;
                        int tabW = cw / tabCount;
                        int newTab = Math.Clamp(Math.Max(relMx - 1, 0) / tabW, 0, tabCount - 1);
                        if (newTab != _activeTab)
                        {
                            _activeTab = newTab;
                            _scrollOffset = 0;
                            _hoverIndex = -1;
                            _hoverButton = -1;
                            NativeMethods.InvalidateRect(hWnd, IntPtr.Zero, false);
                            return IntPtr.Zero;
                        }
                    }
                }
                break;

            case 0x02E0: // WM_DPICHANGED
                {
                    // Remeasure pill row width for new DPI, then reposition
                    MeasureWindowWidth();
                    var rect = Marshal.PtrToStructure<NativeMethods.RECT>(lParam);
                    NativeMethods.SetWindowPos(hWnd, IntPtr.Zero,
                        rect.Left, rect.Top, _windowWidth, _windowHeight,
                        NativeMethods.SWP_NOZORDER | NativeMethods.SWP_NOACTIVATE);
                }
                return IntPtr.Zero;

            case NativeMethods.WM_GETMINMAXINFO:
                OnGetMinMaxInfo(lParam);
                return IntPtr.Zero;

            case NativeMethods.WM_SIZE:
                OnSizeChanged(hWnd);
                return IntPtr.Zero;

            case NativeMethods.WM_TRAY_ICON:
                OnTrayIcon(lParam);
                return IntPtr.Zero;

            case NativeMethods.WM_COMMAND:
                OnCommand(wParam.ToInt32());
                return IntPtr.Zero;

            case NativeMethods.WM_DESTROY:
                Cleanup();
                NativeMethods.PostQuitMessage(0);
                return IntPtr.Zero;
        }

        return NativeMethods.DefWindowProcW(hWnd, msg, wParam, lParam);
    }

    private void OnSizeChanged(IntPtr hWnd)
    {
        NativeMethods.GetWindowRect(hWnd, out var wr);
        _windowWidth = wr.Right - wr.Left;
        _windowHeight = wr.Bottom - wr.Top;
        NativeMethods.InvalidateRect(hWnd, IntPtr.Zero, false);
    }

    private IntPtr OnNcHitTest(IntPtr hWnd, IntPtr lParam)
    {
        int x = NativeMethods.GET_X_LPARAM(lParam);
        int y = NativeMethods.GET_Y_LPARAM(lParam);

        NativeMethods.GetWindowRect(hWnd, out var rc);
        int relX = x - rc.Left;
        int relY = y - rc.Top;
        int w = rc.Right - rc.Left;
        int h = rc.Bottom - rc.Top;

        // Resize borders (vertical only — width is fixed)
        bool top = relY < RESIZE_BORDER;
        bool bottom = relY >= h - RESIZE_BORDER;

        if (top) return NativeMethods.HTTOP;
        if (bottom) return NativeMethods.HTBOTTOM;

        // Tab bar: draggable
        // Header bar: pin/close button area is HTCLIENT, rest is draggable
        if (relY < HEADER_HEIGHT)
        {
            // Pin/close buttons on the right
            if (relX >= w - (PIN_BUTTON_SIZE + 4) * 2 - 4)
                return NativeMethods.HTCLIENT;
            return NativeMethods.HTCAPTION;
        }

        // Tab bar: draggable
        if (relY < TOP_OFFSET)
            return NativeMethods.HTCAPTION;

        return NativeMethods.HTCLIENT;
    }

    private void OnGetMinMaxInfo(IntPtr lParam)
    {
        var mmi = Marshal.PtrToStructure<NativeMethods.MINMAXINFO>(lParam);
        mmi.ptMinTrackSize.X = _windowWidth;
        mmi.ptMaxTrackSize.X = _windowWidth;
        mmi.ptMinTrackSize.Y = 200;
        Marshal.StructureToPtr(mmi, lParam, false);
    }

    private void MeasureWindowWidth()
    {
        using var g = Graphics.FromHwnd(_hWnd);
        using var emojiFont = new Font("Segoe UI Emoji", 8.5f);
        using var textFont = new Font("Segoe UI", 8.5f);
        // Measure worst-case pill labels (unset state with dropdown arrow)
        string[] labels = ["Task \u25bc", "Tone \u25bc", "Format \u25bc", "Length \u25bc"];
        int pillRow = 0;
        foreach (var label in labels)
        {
            int pw = (int)g.MeasureString(label, textFont).Width + 20; // pill padding
            pillRow += pw;
        }
        pillRow += 8 * (labels.Length - 1); // gaps between pills
        pillRow += 16 * 2; // left + right content padding
        // Account for window frame (WS_THICKFRAME borders)
        NativeMethods.GetClientRect(_hWnd, out var cr);
        NativeMethods.GetWindowRect(_hWnd, out var wr);
        int frameW = (wr.Right - wr.Left) - (cr.Right - cr.Left);
        _windowWidth = Math.Max(pillRow + frameW, 320);
        NativeMethods.SetWindowPos(_hWnd, IntPtr.Zero, 0, 0, _windowWidth, _windowHeight,
            NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOZORDER | NativeMethods.SWP_NOACTIVATE);
    }

    private void OnPaint(IntPtr hWnd)
    {
        NativeMethods.BeginPaint(hWnd, out var ps);
        NativeMethods.GetClientRect(hWnd, out var clientRect);
        int w = clientRect.Right - clientRect.Left;
        int h = clientRect.Bottom - clientRect.Top;

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

        // Border
        using (var borderPen = new Pen(BorderColor, 1))
            g.DrawRectangle(borderPen, 0, 0, w - 1, h - 1);

        DrawTabs(g, w);
        if (_activeTab == 2)
            DrawEditTab(g, w, h);
        else
        {
            DrawItems(g, w, h);
            DrawScrollbar(g, w, h);
        }
        DrawTooltip(g, w, h);

        // Animated edge glow during AI processing
        if (_aiProcessing)
            DrawEdgeAnimation(g, w, h);

        // Blit to screen
        using var screenG = Graphics.FromHdc(ps.hdc);
        screenG.DrawImageUnscaled(bmp, 0, 0);

        NativeMethods.EndPaint(hWnd, ref ps);
    }

    private void DrawTabs(Graphics g, int w)
    {
        // Header bar background (above tabs)
        using (var headerBrush = new SolidBrush(Color.FromArgb(35, 35, 35)))
            g.FillRectangle(headerBrush, 1, 1, w - 2, HEADER_HEIGHT);

        // App icon in header (top-left)
        if (_appIconSmall != null)
        {
            int iconSize = 16;
            int iconX = 8;
            int iconY = (HEADER_HEIGHT - iconSize) / 2;
            g.DrawIcon(_appIconSmall, new Rectangle(iconX, iconY, iconSize, iconSize));
        }

        // Pin button in header bar (right side, before close button)
        int closeX = w - PIN_BUTTON_SIZE - 6;
        int pinX = closeX - PIN_BUTTON_SIZE - 4;
        int pinY = (HEADER_HEIGHT - PIN_BUTTON_SIZE) / 2;
        var pinRect = new Rectangle(pinX, pinY, PIN_BUTTON_SIZE, PIN_BUTTON_SIZE);

        if (_pinHover)
        {
            using var hoverBrush = new SolidBrush(Color.FromArgb(60, 60, 60));
            using var hoverPath = CreateRoundedRect(pinRect, 4);
            g.FillPath(hoverBrush, hoverPath);
        }

        using var pinFont = new Font("Segoe UI Emoji", 9f);
        using var pinBrush = new SolidBrush(_alwaysOnTop ? AccentColor : SecondaryText);
        var pinSize = g.MeasureString("\ud83d\udccc", pinFont);
        g.DrawString("\ud83d\udccc", pinFont, pinBrush,
            pinX + (PIN_BUTTON_SIZE - pinSize.Width) / 2,
            pinY + (PIN_BUTTON_SIZE - pinSize.Height) / 2);

        // Close button (right-most)
        int closeY = pinY;
        var closeRect = new Rectangle(closeX, closeY, PIN_BUTTON_SIZE, PIN_BUTTON_SIZE);

        if (_closeHover)
        {
            using var hoverBrush = new SolidBrush(Color.FromArgb(180, 200, 50, 50));
            using var hoverPath = CreateRoundedRect(closeRect, 4);
            g.FillPath(hoverBrush, hoverPath);
        }

        using var closeFont = new Font("Segoe UI", 10f, FontStyle.Bold);
        using var closeBrush = new SolidBrush(SecondaryText);
        var closeSize = g.MeasureString("\u00d7", closeFont);
        g.DrawString("\u00d7", closeFont, closeBrush,
            closeX + (PIN_BUTTON_SIZE - closeSize.Width) / 2,
            closeY + (PIN_BUTTON_SIZE - closeSize.Height) / 2);

        // Tab bar background
        int tabTop = HEADER_HEIGHT;
        using (var tabBgBrush = new SolidBrush(HeaderBg))
            g.FillRectangle(tabBgBrush, 1, tabTop, w - 2, TAB_HEIGHT);

        // Separator below tabs
        using (var sepPen = new Pen(SeparatorColor))
            g.DrawLine(sepPen, 1, tabTop + TAB_HEIGHT, w - 2, tabTop + TAB_HEIGHT);

        int tabCount = 3;
        int tabW = w / tabCount;
        using var tabFont = new Font("Segoe UI", 9.5f);
        using var emojiFont = new Font("Segoe UI Emoji", 9.5f);
        string[] tabIcons = ["\ud83d\udccb", "\ud83d\udd16", "\u2728"];
        string[] tabTexts = [" History", " Bookmarks", " Edit"];

        for (int i = 0; i < tabCount; i++)
        {
            int tabX = i * tabW + 1;
            int tw = i == tabCount - 1 ? w - i * tabW - 2 : tabW - 1;

            // Hover highlight
            if (_hoverTab == i && _activeTab != i)
            {
                using var hoverBrush = new SolidBrush(Color.FromArgb(50, 50, 50));
                g.FillRectangle(hoverBrush, tabX, tabTop, tw, TAB_HEIGHT);
            }

            // Tab label: emoji + text drawn separately
            var iconSize = g.MeasureString(tabIcons[i], emojiFont);
            var textSize = g.MeasureString(tabTexts[i], tabFont);
            float totalWidth = iconSize.Width + textSize.Width - 6;
            float lx = tabX + (tw - totalWidth) / 2;
            float ly = tabTop + (TAB_HEIGHT - Math.Max(iconSize.Height, textSize.Height)) / 2;
            using var textBrush = new SolidBrush(_activeTab == i ? TextColor : SecondaryText);
            g.DrawString(tabIcons[i], emojiFont, textBrush, lx, ly);
            g.DrawString(tabTexts[i], tabFont, textBrush, lx + iconSize.Width - 6, ly);

            // Active indicator
            if (_activeTab == i)
            {
                using var indicatorBrush = new SolidBrush(AccentColor);
                g.FillRectangle(indicatorBrush, tabX + 4, tabTop + TAB_HEIGHT - 3, tw - 8, 2);
            }
        }
    }

    private void DrawHeaderButton(Graphics g, Rectangle rect, string icon, Color color)
    {
        using var font = new Font("Segoe UI Emoji", 10f);
        using var brush = new SolidBrush(color);
        var textSize = g.MeasureString(icon, font);
        float x = rect.X + (rect.Width - textSize.Width) / 2;
        float y = rect.Y + (rect.Height - textSize.Height) / 2;
        g.DrawString(icon, font, brush, x, y);
    }

    private int GetItemHeight(ClipboardItem item)
    {
        return item.Type == ClipboardItemType.Image ? IMAGE_ITEM_HEIGHT : ITEM_HEIGHT;
    }

    private int GetTotalContentHeight()
    {
        int total = 0;
        foreach (var item in ActiveItems)
            total += GetItemHeight(item);
        return total;
    }

    private void DrawItems(Graphics g, int w, int h)
    {
        var items = ActiveItems;
        if (items.Count == 0)
        {
            DrawEmptyState(g, w, h);
            return;
        }

        int contentY = TOP_OFFSET + 1;
        int visibleHeight = h - contentY;
        int y = contentY - _scrollOffset;

        using var itemFont = new Font("Segoe UI", 9.5f);
        using var smallFont = new Font("Segoe UI", 8f);
        using var textBrush = new SolidBrush(TextColor);
        using var secondaryBrush = new SolidBrush(SecondaryText);

        // Clip to content area
        g.SetClip(new Rectangle(0, contentY, w, visibleHeight));

        for (int i = 0; i < items.Count; i++)
        {
            var item = items[i];
            int itemH = GetItemHeight(item);

            if (y + itemH < contentY)
            {
                y += itemH;
                continue;
            }
            if (y > h) break;

            var itemRect = new Rectangle(1, y, w - 2, itemH);

            // Item background
            using (var itemBgBrush = new SolidBrush(i == _hoverIndex ? ItemHoverBg : ItemBg))
                g.FillRectangle(itemBgBrush, itemRect);

            // Draw edge animation if this item is being inline-AI-processed
            if (_inlineAiProcessing && i == _inlineAiItemIndex)
            {
                int thickness = 2;
                int iw = itemRect.Width, ih = itemRect.Height;
                int perimeter = 2 * (iw + ih);
                int segments = 40;
                float segLen = perimeter / (float)segments;
                for (int si = 0; si < segments; si++)
                {
                    float t = (si / (float)segments + _aiAnimPhase) % 1f;
                    float hue = t * 360f;
                    var color = ColorFromHSL(hue, 0.8f, 0.6f);
                    using var pen = new Pen(Color.FromArgb(200, color), thickness);
                    float pos = si * segLen;
                    float nextPos = pos + segLen;
                    // Draw segment around item rect
                    float topLen = iw, rightLen2 = ih, bottomLen = iw, leftLen = ih;
                    float[] edges = [topLen, rightLen2, bottomLen, leftLen];
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
                            case 0: g.DrawLine(pen, itemRect.X + s, itemRect.Y, itemRect.X + en, itemRect.Y); break;
                            case 1: g.DrawLine(pen, itemRect.Right - 1, itemRect.Y + s, itemRect.Right - 1, itemRect.Y + en); break;
                            case 2: g.DrawLine(pen, itemRect.Right - s, itemRect.Bottom - 1, itemRect.Right - en, itemRect.Bottom - 1); break;
                            case 3: g.DrawLine(pen, itemRect.X, itemRect.Bottom - s, itemRect.X, itemRect.Bottom - en); break;
                        }
                        offset = edgeEnd;
                    }
                }
            }

            // Text/image area (left side, leave space for buttons)
            int textLeft = 12;
            int buttonsWidth = (BUTTON_SIZE + BUTTON_MARGIN) * 3 + 14 + 8; // extra 14 for arrow
            int textRight = w - buttonsWidth - 8;

            if (item.Type == ClipboardItemType.Image && item.Image != null)
            {
                // Draw image thumbnail
                int thumbHeight = itemH - 16;
                int thumbWidth = (int)(thumbHeight * ((float)item.Image.Width / item.Image.Height));
                thumbWidth = Math.Min(thumbWidth, textRight - textLeft - 8);
                var thumbRect = new Rectangle(textLeft, y + 4, thumbWidth, thumbHeight);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(item.Image, thumbRect);

                // Image info text next to thumbnail
                var infoText = $"{item.Image.Width}×{item.Image.Height}";
                g.DrawString(infoText, smallFont, secondaryBrush, textLeft + thumbWidth + 8, y + 4);

                // Timestamp
                var timeText = FormatTimestamp(item.Timestamp);
                g.DrawString(timeText, smallFont, secondaryBrush, textLeft + thumbWidth + 8, y + itemH - 18);
            }
            else
            {
                // Draw text
                var textRect = new Rectangle(textLeft, y + 6, textRight - textLeft, itemH - 20);
                using var textFormat = new StringFormat
                {
                    Trimming = StringTrimming.EllipsisCharacter,
                    FormatFlags = StringFormatFlags.NoWrap | StringFormatFlags.LineLimit
                };
                g.DrawString(item.PreviewText, itemFont, textBrush, textRect, textFormat);

                // Timestamp
                var timeText = FormatTimestamp(item.Timestamp);
                g.DrawString(timeText, smallFont, secondaryBrush, textLeft, y + itemH - 18);
            }

            // Action buttons (right side) - only show on hover
            if (i == _hoverIndex)
            {
                int btnY = y + (itemH - BUTTON_SIZE) / 2;

                if (i == _expandedIndex)
                {
                    // Expanded: show bookmark + delete (wider spacing)
                    int expandedGap = 20;
                    int btnX = w - BUTTON_SIZE * 2 - expandedGap - 12;
                    string bookmarkIcon = _activeTab == 0 ? "\uD83D\uDD16" : "\uD83D\uDD16";
                    var bookmarkColor = _activeTab == 0 ? TextColor : Color.FromArgb(220, 160, 60);
                    DrawActionButton(g, new Rectangle(btnX, btnY, BUTTON_SIZE, BUTTON_SIZE),
                        bookmarkIcon, _hoverButton == 0 ? ButtonHover : Color.Transparent, bookmarkColor);
                    btnX += BUTTON_SIZE + expandedGap;

                    DrawActionButton(g, new Rectangle(btnX, btnY, BUTTON_SIZE, BUTTON_SIZE),
                        "🗑️", _hoverButton == 1 ? ButtonHover : Color.Transparent, Color.FromArgb(220, 80, 80));
                }
                else if (_inlineAiProcessing && i == _inlineAiItemIndex)
                {
                    // Show cancel button while inline AI is processing
                    int btnX = w - BUTTON_SIZE - 12;
                    DrawActionButton(g, new Rectangle(btnX, btnY, BUTTON_SIZE, BUTTON_SIZE),
                        "■", _hoverButton == 1 ? Color.FromArgb(220, 60, 60) : Color.FromArgb(190, 40, 40), Color.White);
                }
                else
                {
                    // Normal: paste, AI (split: stars + arrow), more
                    int btnX = w - buttonsWidth;

                    DrawActionButton(g, new Rectangle(btnX, btnY, BUTTON_SIZE, BUTTON_SIZE),
                        "📋", _hoverButton == 0 ? ButtonHover : Color.Transparent, TextColor);
                    btnX += BUTTON_SIZE + BUTTON_MARGIN;

                    // AI split button: stars (left) + arrow (right)
                    DrawActionButton(g, new Rectangle(btnX, btnY, BUTTON_SIZE, BUTTON_SIZE),
                        "✨", _hoverButton == 1 ? ButtonHover : Color.Transparent, TextColor);
                    btnX += BUTTON_SIZE;
                    // Small arrow button
                    int arrowW = 14;
                    var arrowRect = new Rectangle(btnX, btnY, arrowW, BUTTON_SIZE);
                    if (_hoverButton == 3)
                    {
                        using var abg = new SolidBrush(ButtonHover);
                        using var apath = CreateRoundedRect(arrowRect, 3);
                        g.FillPath(abg, apath);
                    }
                    using (var arrowFont = new Font("Segoe UI", 7f))
                    using (var arrowBrush = new SolidBrush(TextColor))
                    {
                        var arrowSize = g.MeasureString("▼", arrowFont);
                        g.DrawString("▼", arrowFont, arrowBrush, btnX + (arrowW - arrowSize.Width) / 2, btnY + (BUTTON_SIZE - arrowSize.Height) / 2);
                    }
                    btnX += arrowW + BUTTON_MARGIN;

                    DrawActionButton(g, new Rectangle(btnX, btnY, BUTTON_SIZE, BUTTON_SIZE),
                        "⋯", _hoverButton == 2 ? ButtonHover : Color.Transparent, TextColor);
                }
            }

            // Separator between items
            using (var sepPen = new Pen(SeparatorColor))
                g.DrawLine(sepPen, 1, y + itemH - 1, w - 2, y + itemH - 1);

            y += itemH;
        }

        g.ResetClip();

        // Draw bubble task dropdown if open
        if (_bubbleTaskDropdownOpen && _bubbleTaskDropdownItemIndex >= 0)
        {
            var items2 = ActiveItems;
            if (_bubbleTaskDropdownItemIndex < items2.Count)
            {
                // Calculate dropdown position (below the arrow button of the item)
                int itemY = contentY - _scrollOffset;
                for (int di = 0; di < _bubbleTaskDropdownItemIndex && di < items2.Count; di++)
                    itemY += GetItemHeight(items2[di]);
                int itemH2 = GetItemHeight(items2[_bubbleTaskDropdownItemIndex]);

                int dropItemH = 28;
                // Skip index 0 ("<unset>") — show tasks starting from index 1
                int taskCount = EditTasks.Length - 1;
                int dropH = taskCount * dropItemH + 4;
                int dropW = 160;
                int dropX = w - 180;
                int dropY = itemY + itemH2;

                // Clamp to window bounds
                if (dropY + dropH > h) dropY = itemY - dropH;
                if (dropX < 4) dropX = 4;

                var dropRect = new Rectangle(dropX, dropY, dropW, dropH);
                using (var dropBg = new SolidBrush(Color.FromArgb(42, 42, 42)))
                using (var dropBorder = new Pen(Color.FromArgb(70, 70, 70)))
                using (var dropPath = CreateRoundedRect(dropRect, 6))
                {
                    g.FillPath(dropBg, dropPath);
                    g.DrawPath(dropBorder, dropPath);
                }

                using var taskFont = new Font("Segoe UI", 8.5f);
                using var taskEmoji = new Font("Segoe UI Emoji", 8.5f);
                using var taskBrush = new SolidBrush(TextColor);
                using var taskDimBrush = new SolidBrush(SecondaryText);

                for (int ti = 0; ti < taskCount; ti++)
                {
                    int taskIdx = ti + 1; // skip "<unset>"
                    int iy = dropY + 2 + ti * dropItemH;
                    bool isHover = _bubbleTaskHoverIndex == taskIdx;

                    if (isHover)
                    {
                        var hlRect = new Rectangle(dropX + 2, iy, dropW - 4, dropItemH);
                        using var hlBrush = new SolidBrush(Color.FromArgb(60, 60, 60));
                        using var hlPath = CreateRoundedRect(hlRect, 4);
                        g.FillPath(hlBrush, hlPath);
                    }

                    string icon = EditTaskIcons[taskIdx];
                    var iconSize2 = g.MeasureString(icon, taskEmoji);
                    g.DrawString(icon, taskEmoji, taskBrush, dropX + 10, iy + (dropItemH - iconSize2.Height) / 2);
                    g.DrawString(EditTasks[taskIdx], taskFont, isHover ? taskBrush : taskDimBrush, dropX + 10 + iconSize2.Width + 2, iy + 6);
                }
            }
        }
    }

    private void DrawScrollbar(Graphics g, int w, int h)
    {
        int contentY = TOP_OFFSET + 1;
        int visibleHeight = h - contentY - 1;
        int totalHeight = GetTotalContentHeight();
        if (totalHeight <= visibleHeight)
            return;

        float thumbRatio = (float)visibleHeight / totalHeight;
        int thumbHeight = Math.Max(SCROLLBAR_MIN_THUMB, (int)(visibleHeight * thumbRatio));
        int trackSpace = visibleHeight - thumbHeight;
        int maxScroll = totalHeight - visibleHeight;
        int thumbY = contentY + (maxScroll > 0 ? (int)((float)_scrollOffset / maxScroll * trackSpace) : 0);

        int scrollX = w - SCROLLBAR_WIDTH - 2;
        var thumbRect = new Rectangle(scrollX, thumbY, SCROLLBAR_WIDTH, thumbHeight);
        var color = _scrollbarHover || _scrollbarDragging ? ScrollbarHoverColor : ScrollbarColor;
        using var brush = new SolidBrush(color);
        using var path = CreateRoundedRect(thumbRect, SCROLLBAR_WIDTH / 2);
        g.FillPath(brush, path);
    }

    private void DrawEditTab(Graphics g, int w, int h)
    {
        int contentY = TOP_OFFSET + 1;
        int pad = 16;
        int cx = pad;  // content x
        int cw = w - pad * 2; // content width
        int cy = contentY + pad;

        using var labelFont = new Font("Segoe UI", 9f);
        using var smallFont = new Font("Segoe UI", 8.5f);
        using var emojiFont = new Font("Segoe UI Emoji", 8.5f);
        using var textBrush = new SolidBrush(TextColor);
        using var dimBrush = new SolidBrush(SecondaryText);
        var fieldBg = Color.FromArgb(38, 38, 38);
        var fieldBorder = Color.FromArgb(65, 65, 65);
        var pillBg = Color.FromArgb(55, 55, 55);
        var pillActiveBg = Color.FromArgb(70, 70, 70);
        var checkColor = AccentColor;

        // --- Pill buttons for options (above text box) ---
        int pillHeight = 28;
        int pillY = cy;
        int pillX = cx;

        // Task dropdown button
        {
            string? taskIcon = _editTaskIndex == 0 ? null : EditTaskIcons[_editTaskIndex];
            string taskLabel = _editTaskIndex == 0 ? "Task" : EditTaskIcons[_editTaskIndex];
            var taskBrush2 = _editTaskIndex == 0 ? dimBrush : textBrush;
            var taskBg = _editTaskIndex == 0 ? pillBg : pillActiveBg;
            string dropArrow = _editTaskDropdownOpen ? " \u25b2" : " \u25bc";
            int taskStartX = pillX;
            pillX = DrawEditPill(g, pillX, pillY, taskLabel + dropArrow, taskBg, _editHoverElement == 20, emojiFont, smallFont, taskBrush2, taskIcon);
            _pillTaskRect = new Rectangle(taskStartX, pillY, pillX - taskStartX, pillHeight);
            pillX += 8;
        }
        // Tone dropdown button
        {
            string? toneIcon = _editToneIndex == 0 ? null : EditToneIcons[_editToneIndex];
            string toneLabel = _editToneIndex == 0 ? "Tone" : EditToneIcons[_editToneIndex];
            var toneBrush2 = _editToneIndex == 0 ? dimBrush : textBrush;
            var toneBg = _editToneIndex == 0 ? pillBg : pillActiveBg;
            string toneArrow = _editToneDropdownOpen ? " \u25b2" : " \u25bc";
            int toneStartX = pillX;
            pillX = DrawEditPill(g, pillX, pillY, toneLabel + toneArrow, toneBg, _editHoverElement == 21, emojiFont, smallFont, toneBrush2, toneIcon);
            _pillToneRect = new Rectangle(toneStartX, pillY, pillX - toneStartX, pillHeight);
            pillX += 8;
        }
        // Format dropdown button
        {
            string? fmtIcon = _editFormatIndex == 0 ? null : EditFormatIcons[_editFormatIndex];
            string fmtLabel = _editFormatIndex == 0 ? "Format" : EditFormatIcons[_editFormatIndex];
            var fmtBrush2 = _editFormatIndex == 0 ? dimBrush : textBrush;
            var fmtBg = _editFormatIndex == 0 ? pillBg : pillActiveBg;
            string fmtArrow = _editFormatDropdownOpen ? " \u25b2" : " \u25bc";
            int fmtStartX = pillX;
            pillX = DrawEditPill(g, pillX, pillY, fmtLabel + fmtArrow, fmtBg, _editHoverElement == 22, emojiFont, smallFont, fmtBrush2, fmtIcon);
            _pillFormatRect = new Rectangle(fmtStartX, pillY, pillX - fmtStartX, pillHeight);
            pillX += 8;
        }
        // Length dropdown button
        {
            string? lenIcon = _editLengthIndex == 0 ? null : EditLengthIcons[_editLengthIndex];
            string lenLabel = _editLengthIndex == 0 ? "Length" : EditLengthIcons[_editLengthIndex];
            var lenBrush2 = _editLengthIndex == 0 ? dimBrush : textBrush;
            var lenBg = _editLengthIndex == 0 ? pillBg : pillActiveBg;
            string lenArrow = _editLengthDropdownOpen ? " \u25b2" : " \u25bc";
            int lenStartX = pillX;
            pillX = DrawEditPill(g, pillX, pillY, lenLabel + lenArrow, lenBg, _editHoverElement == 23, emojiFont, smallFont, lenBrush2, lenIcon);
            _pillLengthRect = new Rectangle(lenStartX, pillY, pillX - lenStartX, pillHeight);
        }

        cy += pillHeight + 12;

        // --- Content text box (1 line default, grows to 4 lines, then scrolls) ---
        string fullText = _editItem?.Text ?? "";
        int textPadX = 10, textPadY = 8;
        int scrollbarW = 6;
        int textAreaW = cw - textPadX * 2 - scrollbarW - 4;
        float lineHeight = g.MeasureString("Ag", labelFont).Height;
        using var sf = new StringFormat(StringFormat.GenericTypographic);
        sf.FormatFlags = 0; // enable word wrap
        var textSize2 = g.MeasureString(fullText, labelFont, new SizeF(textAreaW, 999999), sf);
        int totalTextHeight = Math.Max((int)Math.Ceiling(lineHeight), (int)Math.Ceiling(textSize2.Height));
        int minH = (int)Math.Ceiling(lineHeight);
        int maxH = (int)Math.Ceiling(lineHeight * 4);
        int innerH = Math.Clamp(totalTextHeight, minH, maxH);
        int previewHeight = innerH + textPadY * 2;
        bool needsScroll = totalTextHeight > innerH;

        // Store rect for mouse interaction
        _editTextBoxRect = new Rectangle(cx, cy, cw, previewHeight);

        var previewRect = new Rectangle(cx, cy, cw, previewHeight);
        using (var bgBrush = new SolidBrush(fieldBg))
        using (var borderPen = new Pen(fieldBorder))
        using (var path = CreateRoundedRect(previewRect, 6))
        {
            g.FillPath(bgBrush, path);
            g.DrawPath(borderPen, path);
        }

        // Draw text with vertical scrolling using TranslateTransform
        var textArea = new Rectangle(cx + textPadX, cy + textPadY, textAreaW, innerH);
        int maxTextScroll = Math.Max(0, totalTextHeight - innerH);
        _editTextScrollOffset = Math.Clamp(_editTextScrollOffset, 0, maxTextScroll);

        g.SetClip(textArea);
        var state = g.Save();
        g.TranslateTransform(0, -_editTextScrollOffset);
        var textDrawRect = new RectangleF(textArea.X, textArea.Y, textArea.Width, totalTextHeight + 20);
        g.DrawString(fullText, labelFont, dimBrush, textDrawRect, sf);
        g.Restore(state);
        g.ResetClip();

        // Draw scrollbar if needed
        if (needsScroll)
        {
            int sbX = cx + cw - scrollbarW - 6;
            int sbTrackY = cy + textPadY;
            int sbTrackH = innerH;
            float thumbRatio = (float)innerH / totalTextHeight;
            int sbThumbH = Math.Max(20, (int)(sbTrackH * thumbRatio));
            int sbTrackSpace = sbTrackH - sbThumbH;
            int sbThumbY = sbTrackY + (maxTextScroll > 0 ? (int)((float)_editTextScrollOffset / maxTextScroll * sbTrackSpace) : 0);
            var sbThumbRect = new Rectangle(sbX, sbThumbY, scrollbarW, sbThumbH);
            var sbColor = _editScrollbarDragging ? ScrollbarHoverColor : ScrollbarColor;
            using var sbBrush = new SolidBrush(sbColor);
            using var sbPath = CreateRoundedRect(sbThumbRect, scrollbarW / 2);
            g.FillPath(sbBrush, sbPath);
        }

        cy += previewHeight + 12;

        // --- Model dropdown + Choices slider (same row) ---
        {
            int rowH = 28;

            // Model dropdown pill (icon-only when closed)
            string modelIcon = "\ud83e\udde0";
            string modelArrow = _editModelDropdownOpen ? " \u25b2" : " \u25bc";
            var iconSize = g.MeasureString(modelIcon, emojiFont);
            var arrowSize = g.MeasureString(modelArrow, smallFont);
            int modelPillW = 10 + (int)iconSize.Width + (int)arrowSize.Width + 6;
            var modelPillRect = new Rectangle(cx, cy, modelPillW, rowH);
            var modelBg = _editHoverElement == 33 ? Color.FromArgb(70, 70, 70) : Color.FromArgb(55, 55, 55);
            using (var mBgBrush = new SolidBrush(modelBg))
            using (var mPath = CreateRoundedRect(modelPillRect, rowH / 2))
                g.FillPath(mBgBrush, mPath);

            // Draw emoji icon with emojiFont, arrow with smallFont
            float mTextX = cx + 10;
            g.DrawString(modelIcon, emojiFont, textBrush, mTextX, cy + (rowH - iconSize.Height) / 2);
            mTextX += iconSize.Width;
            g.DrawString(modelArrow, smallFont, textBrush, mTextX, cy + (rowH - arrowSize.Height) / 2);
            _editModelDropdownRect = modelPillRect;

            // Choices slider (middle)
            int sendBtnW = rowH; // square button at end
            int sliderLeft = cx + modelPillW + 16;
            string choicesLabel = $"Choices: {_editChoices}";
            var labelSize = g.MeasureString(choicesLabel, smallFont);
            int labelW = (int)labelSize.Width + 4;
            g.DrawString(choicesLabel, smallFont, dimBrush, sliderLeft, cy + (rowH - labelSize.Height) / 2);

            int trackX = sliderLeft + labelW + 8;
            int trackW = cx + cw - trackX - sendBtnW - 12; // leave room for send button
            int trackY = cy + rowH / 2;
            int thumbR = 8;

            // Store slider rect for interaction
            _editChoicesSliderRect = new Rectangle(trackX, cy, trackW, rowH);

            // Draw track
            using (var trackPen = new Pen(Color.FromArgb(70, 70, 70), 2))
                g.DrawLine(trackPen, trackX, trackY, trackX + trackW, trackY);

            // Draw filled portion
            float ratio = (_editChoices - 1) / 4f; // 1..5 mapped to 0..1
            int filledW = (int)(trackW * ratio);
            using (var filledPen = new Pen(AccentColor, 2))
                g.DrawLine(filledPen, trackX, trackY, trackX + filledW, trackY);

            // Draw thumb
            int thumbX = trackX + filledW;
            var thumbRect = new Rectangle(thumbX - thumbR, trackY - thumbR, thumbR * 2, thumbR * 2);
            bool thumbHover = _editHoverElement == 32 || _editChoicesDragging;
            using var thumbBrush = new SolidBrush(thumbHover ? Color.White : Color.FromArgb(220, 220, 220));
            g.FillEllipse(thumbBrush, thumbRect);

            // Send / Cancel button (end of row, icon only)
            int sendX = cx + cw - sendBtnW;
            var sendRect = new Rectangle(sendX, cy, sendBtnW, rowH);
            _editSendButtonRect = sendRect;

            using var btnIconFont = new Font("Segoe UI", 11f);
            if (_aiProcessing)
            {
                var cancelBg = _editHoverElement == 34 ? Color.FromArgb(220, 60, 60) : Color.FromArgb(190, 40, 40);
                using (var bg = new SolidBrush(cancelBg))
                using (var path = CreateRoundedRect(sendRect, rowH / 2))
                    g.FillPath(bg, path);
                using var stopBrush = new SolidBrush(Color.White);
                using var btnFmt = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                g.DrawString("■", btnIconFont, stopBrush, sendRect, btnFmt);
            }
            else
            {
                var sendBg = _editHoverElement == 34 ? Color.FromArgb(120, 185, 255) : AccentColor;
                using (var bg = new SolidBrush(sendBg))
                using (var path = CreateRoundedRect(sendRect, rowH / 2))
                    g.FillPath(bg, path);
                using var playBrush = new SolidBrush(Color.FromArgb(20, 20, 20));
                using var btnFmt = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                g.DrawString("▶", btnIconFont, playBrush, sendRect, btnFmt);
            }

            cy += rowH + 8;
        }

        // --- Results ---
        if (_aiResults.Count > 0)
        {
            int resultCount = _aiResults.Count;
            _aiResultTextBoxRects = new Rectangle[resultCount];
            var btnRects = new List<Rectangle>();

            for (int ri = 0; ri < resultCount; ri++)
            {
                if (ri >= _aiResults.Count) break;
                string resultText = _aiResults[ri];
                int textPadX2 = 10, textPadY2 = 8;
                int sbW = 6;
                int resultCw = (int)(cw * 0.95); // 95% width to leave room for action buttons
                int txtW = resultCw - textPadX2 * 2 - sbW - 4;
                var resultSize = g.MeasureString(resultText, labelFont, new SizeF(txtW, 999999), sf);
                int minResultH = (int)Math.Ceiling(lineHeight * 1.5);
                int totalH = Math.Max(minResultH, (int)Math.Ceiling(resultSize.Height));
                int innerRH = Math.Clamp(totalH, minResultH, (int)Math.Ceiling(lineHeight * 4));
                int boxH = innerRH + textPadY2 * 2;
                bool resultNeedsScroll = totalH > innerRH;

                var boxRect = new Rectangle(cx, cy, resultCw, boxH);
                _aiResultTextBoxRects[ri] = boxRect;

                // Draw result textbox
                using (var bgBrush2 = new SolidBrush(fieldBg))
                using (var borderPen2 = new Pen(fieldBorder))
                using (var path2 = CreateRoundedRect(boxRect, 6))
                {
                    g.FillPath(bgBrush2, path2);
                    g.DrawPath(borderPen2, path2);
                }

                // Draw result text with scrolling
                var resultArea = new Rectangle(cx + textPadX2, cy + textPadY2, txtW, innerRH);
                if (ri < _aiResultScrollOffsets.Length)
                {
                    int maxRS = Math.Max(0, totalH - innerRH);
                    _aiResultScrollOffsets[ri] = Math.Clamp(_aiResultScrollOffsets[ri], 0, maxRS);
                }
                int rScroll = ri < _aiResultScrollOffsets.Length ? _aiResultScrollOffsets[ri] : 0;

                g.SetClip(resultArea);
                var rState = g.Save();
                g.TranslateTransform(0, -rScroll);
                var rDrawRect = new RectangleF(resultArea.X, resultArea.Y, resultArea.Width, totalH + 20);
                g.DrawString(resultText, labelFont, textBrush, rDrawRect, sf);
                g.Restore(rState);
                g.ResetClip();

                // Draw scrollbar if needed
                if (resultNeedsScroll)
                {
                    int rsbX = cx + resultCw - sbW - 6;
                    int rsbTrackY = cy + textPadY2;
                    int rsbTrackH = innerRH;
                    float rThumbRatio = (float)innerRH / totalH;
                    int rThumbH = Math.Max(20, (int)(rsbTrackH * rThumbRatio));
                    int rTrackSpace = rsbTrackH - rThumbH;
                    int maxRScroll = Math.Max(0, totalH - innerRH);
                    int rThumbY = rsbTrackY + (maxRScroll > 0 ? (int)((float)rScroll / maxRScroll * rTrackSpace) : 0);
                    var rThumbRect = new Rectangle(rsbX, rThumbY, sbW, rThumbH);
                    using var rsbBrush = new SolidBrush(ScrollbarColor);
                    using var rsbPath = CreateRoundedRect(rThumbRect, sbW / 2);
                    g.FillPath(rsbBrush, rsbPath);
                }

                cy += boxH + 10;

                // Action buttons: refine, copy, paste (vertically stacked to the right of textbox)
                int abtnSize = 20;
                int abtnGap = 0;
                int abtnColH = abtnSize * 3 + abtnGap * 2;
                int abtnX = cx + resultCw + 4;
                int abtnY = cy - boxH - 4 + (boxH - abtnColH) / 2;
                string[] abtnIcons = ["✏️", "➕", "📋"];

                for (int bi = 0; bi < 3; bi++)
                {
                    var abtnRect = new Rectangle(abtnX, abtnY, abtnSize, abtnSize);
                    btnRects.Add(abtnRect);
                    int hoverKey = 600 + ri * 10 + bi;
                    bool isHover = _editHoverElement == hoverKey;
                    var abtnBg = isHover ? Color.FromArgb(70, 70, 70) : Color.Transparent;
                    if (abtnBg != Color.Transparent)
                    {
                        using var abBg = new SolidBrush(abtnBg);
                        using var abPath = CreateRoundedRect(abtnRect, 4);
                        g.FillPath(abBg, abPath);
                    }
                    using var abFont = new Font("Segoe UI Emoji", 7f);
                    var abIconSize = g.MeasureString(abtnIcons[bi], abFont);
                    g.DrawString(abtnIcons[bi], abFont, dimBrush, abtnX + (abtnSize - abIconSize.Width) / 2, abtnY + (abtnSize - abIconSize.Height) / 2);
                    abtnY += abtnSize + abtnGap;
                }
            }
            _aiResultButtonRects = btnRects.ToArray();
        }
        else
        {
            _aiResultTextBoxRects = [];
            _aiResultButtonRects = [];
        }

        // --- Dropdown lists (drawn over other content when open) ---
        if (_editLengthDropdownOpen)
        {
            int dropX = cx;
            int dropItemH = 28;
            int dropW = cw;
            int dropY = pillY + pillHeight + 2;
            int dropH = EditLengths.Length * dropItemH + 4;

            var dropRect = new Rectangle(dropX, dropY, dropW, dropH);
            using (var dropBg4 = new SolidBrush(Color.FromArgb(42, 42, 42)))
            using (var dropBorder4 = new Pen(Color.FromArgb(70, 70, 70)))
            using (var dropPath4 = CreateRoundedRect(dropRect, 6))
            {
                g.FillPath(dropBg4, dropPath4);
                g.DrawPath(dropBorder4, dropPath4);
            }

            for (int i = 0; i < EditLengths.Length; i++)
            {
                int iy = dropY + 2 + i * dropItemH;
                bool isHover = _editHoverElement == 400 + i;
                bool isSelected = i == _editLengthIndex;

                if (isHover || isSelected)
                {
                    var hlRect = new Rectangle(dropX + 2, iy, dropW - 4, dropItemH);
                    var hlColor = isHover ? Color.FromArgb(60, 60, 60) : Color.FromArgb(50, 50, 50);
                    using var hlBrush = new SolidBrush(hlColor);
                    using var hlPath = CreateRoundedRect(hlRect, 4);
                    g.FillPath(hlBrush, hlPath);
                }

                if (i == 0)
                {
                    g.DrawString(EditLengths[i], smallFont, dimBrush, dropX + 10, iy + 6);
                }
                else
                {
                    string icon = EditLengthIcons[i];
                    var iconSize = g.MeasureString(icon, emojiFont);
                    g.DrawString(icon, emojiFont, textBrush, dropX + 10, iy + 6);
                    g.DrawString(EditLengths[i], smallFont, textBrush, dropX + 10 + iconSize.Width + 2, iy + 6);
                }
            }
        }

        if (_editFormatDropdownOpen)
        {
            int dropX = cx;
            int dropItemH = 28;
            int dropW = cw;
            int dropY = pillY + pillHeight + 2;
            int dropH = EditFormats.Length * dropItemH + 4;

            var dropRect = new Rectangle(dropX, dropY, dropW, dropH);
            using (var dropBg3 = new SolidBrush(Color.FromArgb(42, 42, 42)))
            using (var dropBorder3 = new Pen(Color.FromArgb(70, 70, 70)))
            using (var dropPath3 = CreateRoundedRect(dropRect, 6))
            {
                g.FillPath(dropBg3, dropPath3);
                g.DrawPath(dropBorder3, dropPath3);
            }

            for (int i = 0; i < EditFormats.Length; i++)
            {
                int iy = dropY + 2 + i * dropItemH;
                bool isHover = _editHoverElement == 300 + i;
                bool isSelected = i == _editFormatIndex;

                if (isHover || isSelected)
                {
                    var hlRect = new Rectangle(dropX + 2, iy, dropW - 4, dropItemH);
                    var hlColor = isHover ? Color.FromArgb(60, 60, 60) : Color.FromArgb(50, 50, 50);
                    using var hlBrush = new SolidBrush(hlColor);
                    using var hlPath = CreateRoundedRect(hlRect, 4);
                    g.FillPath(hlBrush, hlPath);
                }

                if (i == 0)
                {
                    g.DrawString(EditFormats[i], smallFont, dimBrush, dropX + 10, iy + 6);
                }
                else
                {
                    string icon = EditFormatIcons[i];
                    var iconSize = g.MeasureString(icon, emojiFont);
                    g.DrawString(icon, emojiFont, textBrush, dropX + 10, iy + 6);
                    g.DrawString(EditFormats[i], smallFont, textBrush, dropX + 10 + iconSize.Width + 2, iy + 6);
                }
            }
        }

        if (_editToneDropdownOpen)
        {
            int dropX = cx;
            int dropItemH = 28;
            int dropW = cw;
            int dropY = pillY + pillHeight + 2;
            int dropH = EditTones.Length * dropItemH + 4;

            var dropRect = new Rectangle(dropX, dropY, dropW, dropH);
            using (var dropBg2 = new SolidBrush(Color.FromArgb(42, 42, 42)))
            using (var dropBorder2 = new Pen(Color.FromArgb(70, 70, 70)))
            using (var dropPath2 = CreateRoundedRect(dropRect, 6))
            {
                g.FillPath(dropBg2, dropPath2);
                g.DrawPath(dropBorder2, dropPath2);
            }

            for (int i = 0; i < EditTones.Length; i++)
            {
                int iy = dropY + 2 + i * dropItemH;
                bool isHover = _editHoverElement == 200 + i;
                bool isSelected = i == _editToneIndex;

                if (isHover || isSelected)
                {
                    var hlRect = new Rectangle(dropX + 2, iy, dropW - 4, dropItemH);
                    var hlColor = isHover ? Color.FromArgb(60, 60, 60) : Color.FromArgb(50, 50, 50);
                    using var hlBrush = new SolidBrush(hlColor);
                    using var hlPath = CreateRoundedRect(hlRect, 4);
                    g.FillPath(hlBrush, hlPath);
                }

                if (i == 0)
                {
                    g.DrawString(EditTones[i], smallFont, dimBrush, dropX + 10, iy + 6);
                }
                else
                {
                    string icon = EditToneIcons[i];
                    var iconSize = g.MeasureString(icon, emojiFont);
                    g.DrawString(icon, emojiFont, textBrush, dropX + 10, iy + 6);
                    g.DrawString(EditTones[i], smallFont, textBrush, dropX + 10 + iconSize.Width + 2, iy + 6);
                }
            }
        }

        if (_editTaskDropdownOpen)
        {
            int dropX = cx;
            int dropItemH = 28;
            int dropW = cw;
            int dropY = pillY + pillHeight + 2;
            int dropH = EditTasks.Length * dropItemH + 4;

            // Background + border
            var dropRect = new Rectangle(dropX, dropY, dropW, dropH);
            using (var dropBg = new SolidBrush(Color.FromArgb(42, 42, 42)))
            using (var dropBorder = new Pen(Color.FromArgb(70, 70, 70)))
            using (var dropPath = CreateRoundedRect(dropRect, 6))
            {
                g.FillPath(dropBg, dropPath);
                g.DrawPath(dropBorder, dropPath);
            }

            for (int i = 0; i < EditTasks.Length; i++)
            {
                int iy = dropY + 2 + i * dropItemH;
                bool isHover = _editHoverElement == 100 + i;
                bool isSelected = i == _editTaskIndex;

                if (isHover || isSelected)
                {
                    var hlRect = new Rectangle(dropX + 2, iy, dropW - 4, dropItemH);
                    var hlColor = isHover ? Color.FromArgb(60, 60, 60) : Color.FromArgb(50, 50, 50);
                    using var hlBrush = new SolidBrush(hlColor);
                    using var hlPath = CreateRoundedRect(hlRect, 4);
                    g.FillPath(hlBrush, hlPath);
                }

                string itemText = i == 0 ? EditTasks[i] : EditTaskIcons[i] + "  " + EditTasks[i];
                var itemBrush = i == 0 ? dimBrush : textBrush;
                if (i == 0)
                {
                    g.DrawString(itemText, smallFont, itemBrush, dropX + 10, iy + 6);
                }
                else
                {
                    // Draw emoji icon with emoji font, then text with regular font
                    string icon = EditTaskIcons[i];
                    var iconSize = g.MeasureString(icon, emojiFont);
                    g.DrawString(icon, emojiFont, itemBrush, dropX + 10, iy + 6);
                    g.DrawString(EditTasks[i], smallFont, itemBrush, dropX + 10 + iconSize.Width + 2, iy + 6);
                }
            }
        }

        if (_editModelDropdownOpen)
        {
            int dropItemH = 28;
            // Measure widest model name to size the dropdown
            int maxTextW = 0;
            foreach (var m in _editModels)
            {
                int tw = (int)g.MeasureString(m, smallFont).Width;
                if (tw > maxTextW) maxTextW = tw;
            }
            int dropW = Math.Max(_editModelDropdownRect.Width + 40, maxTextW + 24);
            int dropX = _editModelDropdownRect.X;
            int dropY = _editModelDropdownRect.Y + _editModelDropdownRect.Height + 2;
            int dropH = _editModels.Length * dropItemH + 4;

            var dropRect = new Rectangle(dropX, dropY, dropW, dropH);
            using (var dropBgM = new SolidBrush(Color.FromArgb(42, 42, 42)))
            using (var dropBorderM = new Pen(Color.FromArgb(70, 70, 70)))
            using (var dropPathM = CreateRoundedRect(dropRect, 6))
            {
                g.FillPath(dropBgM, dropPathM);
                g.DrawPath(dropBorderM, dropPathM);
            }

            for (int i = 0; i < _editModels.Length; i++)
            {
                int iy = dropY + 2 + i * dropItemH;
                bool isHover = _editHoverElement == 500 + i;
                bool isSelected = i == _editModelIndex;

                if (isHover || isSelected)
                {
                    var hlRect = new Rectangle(dropX + 2, iy, dropW - 4, dropItemH);
                    var hlColor = isHover ? Color.FromArgb(60, 60, 60) : Color.FromArgb(50, 50, 50);
                    using var hlBrush = new SolidBrush(hlColor);
                    using var hlPath = CreateRoundedRect(hlRect, 4);
                    g.FillPath(hlBrush, hlPath);
                }

                g.DrawString(_editModels[i], smallFont, i == _editModelIndex ? textBrush : dimBrush, dropX + 10, iy + 6);
            }
        }
    }

    private void DrawEditCheckbox(Graphics g, int x, int y, string label, bool isChecked, bool isHover)
    {
        int boxSize = 16;
        var boxRect = new Rectangle(x, y, boxSize, boxSize);
        if (isChecked)
        {
            using var fillBrush = new SolidBrush(AccentColor);
            using var path = CreateRoundedRect(boxRect, 3);
            g.FillPath(fillBrush, path);
            // Checkmark
            using var checkFont = new Font("Segoe UI", 8f, FontStyle.Bold);
            using var checkBrush = new SolidBrush(Color.White);
            g.DrawString("\u2713", checkFont, checkBrush, x + 1, y);
        }
        else
        {
            using var borderPen = new Pen(Color.FromArgb(80, 80, 80));
            using var path = CreateRoundedRect(boxRect, 3);
            g.DrawPath(borderPen, path);
        }

        using var font = new Font("Segoe UI", 9f);
        using var brush = new SolidBrush(isHover ? TextColor : SecondaryText);
        g.DrawString(label, font, brush, x + boxSize + 4, y);
    }

    private int DrawEditPill(Graphics g, int x, int y, string text, Color bg, bool isHover, Font emojiFont, Font textFont, Brush textBrush, string? icon = null)
    {
        // Measure: icon (emoji font) + text (text font)
        float iconW = 0;
        if (icon != null)
            iconW = g.MeasureString(icon, emojiFont).Width;
        string displayText = icon != null ? text.Substring(icon.Length) : text;
        var textSize = g.MeasureString(displayText, textFont);
        int pw = (int)(iconW + textSize.Width) + 20;
        int ph = 28;
        var pillRect = new Rectangle(x, y, pw, ph);
        var bgColor = isHover ? Color.FromArgb(bg.R + 15, bg.G + 15, bg.B + 15) : bg;
        using var bgBrush = new SolidBrush(bgColor);
        using var path = CreateRoundedRect(pillRect, ph / 2);
        g.FillPath(bgBrush, path);
        float dx = x + 10;
        if (icon != null)
        {
            g.DrawString(icon, emojiFont, textBrush, dx, y + 5);
            dx += iconW;
        }
        g.DrawString(displayText, textFont, textBrush, dx, y + 5);
        return x + pw;
    }

    private void DrawTooltip(Graphics g, int w, int h)
    {
        if (_tooltipText == null)
            return;

        using var font = new Font("Segoe UI", 8.5f);
        var size = g.MeasureString(_tooltipText, font);
        int pw = (int)size.Width + 12;
        int ph = (int)size.Height + 6;

        // Position below cursor, clamped to window bounds
        int tx = Math.Clamp(_tooltipPos.X - pw / 2, 2, w - pw - 2);
        int ty = _tooltipPos.Y + 22;
        if (ty + ph > h - 2)
            ty = _tooltipPos.Y - ph - 4; // flip above if no room below

        var tooltipRect = new Rectangle(tx, ty, pw, ph);
        using var bgBrush = new SolidBrush(Color.FromArgb(240, 50, 50, 50));
        using var borderPen = new Pen(Color.FromArgb(80, 80, 80));
        using var path = CreateRoundedRect(tooltipRect, 4);
        g.FillPath(bgBrush, path);
        g.DrawPath(borderPen, path);

        using var textBrush = new SolidBrush(Color.FromArgb(220, 220, 220));
        g.DrawString(_tooltipText, font, textBrush, tx + 6, ty + 3);
    }

    private void DrawEdgeAnimation(Graphics g, int w, int h)
    {
        // Colorful gradient border that rotates around the window edges
        int thickness = 2;
        int perimeter = 2 * (w + h);
        int segments = 60;
        float segLen = perimeter / (float)segments;

        for (int i = 0; i < segments; i++)
        {
            float t = (i / (float)segments + _aiAnimPhase) % 1f;
            // Rainbow hue
            float hue = t * 360f;
            var color = ColorFromHSL(hue, 0.8f, 0.6f);

            float pos = i * segLen;
            float nextPos = pos + segLen;

            // Map perimeter position to edge coordinates
            using var pen = new Pen(Color.FromArgb(200, color), thickness);
            DrawEdgeSegment(g, w, h, pos, nextPos, pen);
        }
    }

    private static void DrawEdgeSegment(Graphics g, int w, int h, float start, float end, Pen pen)
    {
        // Perimeter: top (0..w) → right (w..w+h) → bottom (w+h..2w+h) → left (2w+h..2w+2h)
        float topLen = w, rightLen = h, bottomLen = w, leftLen = h;
        float[] edges = [topLen, rightLen, bottomLen, leftLen];
        float offset = 0;

        for (int e = 0; e < 4; e++)
        {
            float edgeEnd = offset + edges[e];
            if (start >= edgeEnd) { offset = edgeEnd; continue; }
            if (end <= offset) break;

            float s = Math.Max(start - offset, 0);
            float en = Math.Min(end - offset, edges[e]);

            switch (e)
            {
                case 0: g.DrawLine(pen, s, 0, en, 0); break;           // top
                case 1: g.DrawLine(pen, w - 1, s, w - 1, en); break;   // right
                case 2: g.DrawLine(pen, w - s, h - 1, w - en, h - 1); break; // bottom (reversed)
                case 3: g.DrawLine(pen, 0, h - s, 0, h - en); break;   // left (reversed)
            }
            offset = edgeEnd;
        }
    }

    private static Color ColorFromHSL(float hue, float sat, float lit)
    {
        float c = (1 - Math.Abs(2 * lit - 1)) * sat;
        float x = c * (1 - Math.Abs(hue / 60 % 2 - 1));
        float m = lit - c / 2;
        float r, g2, b;
        if (hue < 60) { r = c; g2 = x; b = 0; }
        else if (hue < 120) { r = x; g2 = c; b = 0; }
        else if (hue < 180) { r = 0; g2 = c; b = x; }
        else if (hue < 240) { r = 0; g2 = x; b = c; }
        else if (hue < 300) { r = x; g2 = 0; b = c; }
        else { r = c; g2 = 0; b = x; }
        return Color.FromArgb((int)((r + m) * 255), (int)((g2 + m) * 255), (int)((b + m) * 255));
    }

    private static void DrawActionButton(Graphics g, Rectangle rect, string icon, Color bgColor, Color fgColor)
    {
        if (bgColor != Color.Transparent)
        {
            using var bgBrush = new SolidBrush(bgColor);
            using var path = CreateRoundedRect(rect, BORDER_RADIUS);
            g.FillPath(bgBrush, path);
        }

        using var font = new Font("Segoe UI Emoji", 10f);
        using var brush = new SolidBrush(fgColor);
        var size = g.MeasureString(icon, font);
        float x = rect.X + (rect.Width - size.Width) / 2;
        float y = rect.Y + (rect.Height - size.Height) / 2;
        g.DrawString(icon, font, brush, x, y);
    }

    private void DrawEmptyState(Graphics g, int w, int h)
    {
        using var font = new Font("Segoe UI", 11f);
        using var brush = new SolidBrush(SecondaryText);
        var text = _activeTab == 0 ? "No clipboard history yet" : "No bookmarks";
        var size = g.MeasureString(text, font);
        g.DrawString(text, font, brush, (w - size.Width) / 2, h / 2 - size.Height);

        using var smallFont = new Font("Segoe UI", 9f);
        var hint = _activeTab == 0 ? "Copy something to get started" : "Bookmark items from the history tab";
        var hintSize = g.MeasureString(hint, smallFont);
        g.DrawString(hint, smallFont, brush, (w - hintSize.Width) / 2, h / 2 + 4);
    }

    private static GraphicsPath CreateRoundedRect(Rectangle bounds, int radius)
    {
        var path = new GraphicsPath();
        int d = radius * 2;
        path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
        path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
        path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
        path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }

    private void OnMouseWheel(IntPtr wParam)
    {
        int delta = NativeMethods.GET_WHEEL_DELTA_WPARAM(wParam);

        if (_activeTab == 2)
        {
            // Scroll the edit text box (smaller step for the compact text area)
            int step = Math.Max(1, Math.Abs(delta) / 4);
            _editTextScrollOffset += delta < 0 ? step : -step;
            _editTextScrollOffset = Math.Max(0, _editTextScrollOffset);
            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            return;
        }

        NativeMethods.GetClientRect(_hWnd, out var cr);
        int clientHeight = cr.Bottom - cr.Top;
        int maxScroll = Math.Max(0, GetTotalContentHeight() - (clientHeight - TOP_OFFSET - 1));
        _scrollOffset = Math.Clamp(_scrollOffset - delta, 0, maxScroll);
        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
    }

    // Compute the Y offset of item at given index (from content top)
    private int GetItemYOffset(int index)
    {
        int offset = 0;
        var items = ActiveItems;
        for (int i = 0; i < index && i < items.Count; i++)
            offset += GetItemHeight(items[i]);
        return offset;
    }

    private void OnMouseMove(IntPtr lParam)
    {
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

        int mx = NativeMethods.GET_X_LPARAM(lParam);
        int my = NativeMethods.GET_Y_LPARAM(lParam);

        NativeMethods.GetClientRect(_hWnd, out var cr);
        int clientWidth = cr.Right - cr.Left;
        int clientHeight = cr.Bottom - cr.Top;
        int contentY = TOP_OFFSET + 1;

        // Handle scrollbar dragging (history/bookmarks)
        if (_scrollbarDragging)
        {
            int visibleHeight = clientHeight - contentY - 1;
            int totalHeight = GetTotalContentHeight();
            int maxScroll = Math.Max(0, totalHeight - visibleHeight);
            float thumbRatio = (float)visibleHeight / totalHeight;
            int thumbHeight = Math.Max(SCROLLBAR_MIN_THUMB, (int)(visibleHeight * thumbRatio));
            int trackSpace = visibleHeight - thumbHeight;
            if (trackSpace > 0)
            {
                int deltaY = my - _scrollbarDragStartY;
                int newOffset = _scrollbarDragStartOffset + (int)((float)deltaY / trackSpace * maxScroll);
                _scrollOffset = Math.Clamp(newOffset, 0, maxScroll);
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            }
            return;
        }
        // Handle edit textbox scrollbar dragging
        if (_editScrollbarDragging)
        {
            var r = _editTextBoxRect;
            int textPadY = 8, scrollbarW = 6;
            int textAreaW = r.Width - 20 - scrollbarW - 4; // match drawing calc
            using var tmpG = Graphics.FromHwnd(_hWnd);
            using var tmpFont = new Font("Segoe UI", 9f);
            using var tmpSf = new StringFormat(StringFormat.GenericTypographic);
            tmpSf.FormatFlags = 0;
            float lineH = tmpG.MeasureString("Ag", tmpFont).Height;
            string text = _editItem?.Text ?? "";
            var ts = tmpG.MeasureString(text, tmpFont, new SizeF(textAreaW, 999999), tmpSf);
            int totalTH = Math.Max((int)Math.Ceiling(lineH), (int)Math.Ceiling(ts.Height));
            int innerH = r.Height - textPadY * 2;
            int maxScroll = Math.Max(0, totalTH - innerH);
            float thumbRatio = (float)innerH / totalTH;
            int thumbH = Math.Max(20, (int)(innerH * thumbRatio));
            int trackSpace = innerH - thumbH;
            if (trackSpace > 0)
            {
                int deltaY = my - _editScrollbarDragStartY;
                int newOffset = _editScrollbarDragStartOffset + (int)((float)deltaY / trackSpace * maxScroll);
                _editTextScrollOffset = Math.Clamp(newOffset, 0, maxScroll);
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            }
            return;
        }
        // Handle result textbox scrollbar dragging
        if (_aiResultScrollbarDragging)
        {
            int ri = _aiResultScrollbarDragIndex;
            if (ri >= 0 && ri < _aiResultTextBoxRects.Length && ri < _aiResultScrollOffsets.Length && ri < _aiResults.Count)
            {
                var rr = _aiResultTextBoxRects[ri];
                int textPadY = 8;
                int textAreaW = rr.Width - 20 - 6 - 4;
                using var tmpG = Graphics.FromHwnd(_hWnd);
                using var tmpFont = new Font("Segoe UI", 9f);
                using var tmpSf = new StringFormat(StringFormat.GenericTypographic);
                tmpSf.FormatFlags = 0;
                float lineH = tmpG.MeasureString("Ag", tmpFont).Height;
                var rs = tmpG.MeasureString(_aiResults[ri], tmpFont, new SizeF(textAreaW, 999999), tmpSf);
                int totalH = Math.Max((int)Math.Ceiling(lineH), (int)Math.Ceiling(rs.Height));
                int innerH = rr.Height - textPadY * 2;
                int maxScroll = Math.Max(0, totalH - innerH);
                float thumbRatio = (float)innerH / totalH;
                int thumbH = Math.Max(20, (int)(innerH * thumbRatio));
                int trackSpace = innerH - thumbH;
                if (trackSpace > 0)
                {
                    int deltaY = my - _aiResultScrollbarDragStartY;
                    int newOffset = _aiResultScrollbarDragStartOffset + (int)((float)deltaY / trackSpace * maxScroll);
                    _aiResultScrollOffsets[ri] = Math.Clamp(newOffset, 0, maxScroll);
                    NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                }
            }
            return;
        }
        // Handle choices slider dragging
        if (_editChoicesDragging)
        {
            var sr = _editChoicesSliderRect;
            if (sr.Width > 0)
            {
                float ratio = Math.Clamp((float)(mx - sr.X) / sr.Width, 0f, 1f);
                _editChoices = Math.Clamp((int)Math.Round(ratio * 4) + 1, 1, 5);
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            }
            return;
        }
        int newHoverIndex = -1;
        int newHoverButton = -1;
        int newHoverTab = -1;
        bool newPinHover = false;
        bool newCloseHover = false;

        // Pin/close button hover detection (header bar)
        if (my < HEADER_HEIGHT)
        {
            int closeX = clientWidth - PIN_BUTTON_SIZE - 6;
            int btnPinX = closeX - PIN_BUTTON_SIZE - 4;
            int btnY = (HEADER_HEIGHT - PIN_BUTTON_SIZE) / 2;
            if (mx >= closeX && mx <= closeX + PIN_BUTTON_SIZE && my >= btnY && my <= btnY + PIN_BUTTON_SIZE)
                newCloseHover = true;
            else if (mx >= btnPinX && mx <= btnPinX + PIN_BUTTON_SIZE && my >= btnY && my <= btnY + PIN_BUTTON_SIZE)
                newPinHover = true;
        }

        // Tab hover detection
        if (my >= HEADER_HEIGHT && my < TOP_OFFSET)
        {
            int tabCount = 3;
            int tabW = clientWidth / tabCount;
            int adjustedX = Math.Max(mx - 1, 0);
            int tab = Math.Clamp(adjustedX / tabW, 0, tabCount - 1);
            newHoverTab = tab;
        }

        var items = ActiveItems;
        int newEditHover = -1;

        if (_activeTab == 2 && my > contentY)
        {
            // Edit tab hover detection — pills are above text box
            int pad = 16;
            int cx = pad;
            int cw = clientWidth - pad * 2;

            // Pills are at contentY + pad
            int pillY = contentY + pad;
            int pillHeight = 28;

            // Dropdown lists (priority over pills when open)
            if (_editTaskDropdownOpen)
            {
                int dropX = cx;
                int dropItemH = 28;
                int dropW = cw;
                int dropY = pillY + pillHeight + 2;
                int dropH = EditTasks.Length * dropItemH + 4;
                if (mx >= dropX && mx < dropX + dropW && my >= dropY && my < dropY + dropH)
                {
                    int idx = (my - dropY - 2) / dropItemH;
                    if (idx >= 0 && idx < EditTasks.Length)
                        newEditHover = 100 + idx;
                }
            }
            if (_editToneDropdownOpen)
            {
                int dropX = cx;
                int dropItemH = 28;
                int dropW = cw;
                int dropY = pillY + pillHeight + 2;
                int dropH = EditTones.Length * dropItemH + 4;
                if (mx >= dropX && mx < dropX + dropW && my >= dropY && my < dropY + dropH)
                {
                    int idx = (my - dropY - 2) / dropItemH;
                    if (idx >= 0 && idx < EditTones.Length)
                        newEditHover = 200 + idx;
                }
            }

            if (_editFormatDropdownOpen)
            {
                int dropX = cx;
                int dropItemH = 28;
                int dropW = cw;
                int dropY = pillY + pillHeight + 2;
                int dropH = EditFormats.Length * dropItemH + 4;
                if (mx >= dropX && mx < dropX + dropW && my >= dropY && my < dropY + dropH)
                {
                    int idx = (my - dropY - 2) / dropItemH;
                    if (idx >= 0 && idx < EditFormats.Length)
                        newEditHover = 300 + idx;
                }
            }

            if (_editLengthDropdownOpen)
            {
                int dropX = cx;
                int dropItemH = 28;
                int dropW = cw;
                int dropY = pillY + pillHeight + 2;
                int dropH = EditLengths.Length * dropItemH + 4;
                if (mx >= dropX && mx < dropX + dropW && my >= dropY && my < dropY + dropH)
                {
                    int idx = (my - dropY - 2) / dropItemH;
                    if (idx >= 0 && idx < EditLengths.Length)
                        newEditHover = 400 + idx;
                }
            }

            // Pill buttons: Task(20), Tone(21), Format(22), Length(23)
            if (newEditHover == -1 && my >= pillY && my < pillY + pillHeight)
            {
                if (_pillTaskRect.Contains(mx, my)) newEditHover = 20;
                else if (_pillToneRect.Contains(mx, my)) newEditHover = 21;
                else if (_pillFormatRect.Contains(mx, my)) newEditHover = 22;
                else if (_pillLengthRect.Contains(mx, my)) newEditHover = 23;
            }

            // Choices slider hover (element 32)
            if (newEditHover == -1 && _editChoicesSliderRect.Width > 0)
            {
                var sr = _editChoicesSliderRect;
                if (mx >= sr.X && mx < sr.X + sr.Width && my >= sr.Y && my < sr.Y + sr.Height)
                    newEditHover = 32;
            }

            // Model dropdown pill hover (element 33)
            if (newEditHover == -1 && _editModelDropdownRect.Width > 0 && _editModelDropdownRect.Contains(mx, my))
                newEditHover = 33;

            // Model dropdown list hover (500+)
            if (_editModelDropdownOpen && newEditHover == -1)
            {
                int dropItemH = 28;
                // Measure widest model name to match drawing width
                int maxTextW2 = 0;
                using (var tempG = Graphics.FromHwnd(_hWnd))
                using (var sf = new Font("Segoe UI", 9f))
                {
                    foreach (var m in _editModels)
                    {
                        int tw = (int)tempG.MeasureString(m, sf).Width;
                        if (tw > maxTextW2) maxTextW2 = tw;
                    }
                }
                int dropW = Math.Max(_editModelDropdownRect.Width + 40, maxTextW2 + 24);
                int dropX = _editModelDropdownRect.X;
                int dropY = _editModelDropdownRect.Y + _editModelDropdownRect.Height + 2;
                int dropH = _editModels.Length * dropItemH + 4;
                if (mx >= dropX && mx < dropX + dropW && my >= dropY && my < dropY + dropH)
                {
                    int idx = (my - dropY - 2) / dropItemH;
                    if (idx >= 0 && idx < _editModels.Length)
                        newEditHover = 500 + idx;
                }
            }

            // Send/Cancel button hover (element 34)
            if (newEditHover == -1 && _editSendButtonRect.Width > 0 && _editSendButtonRect.Contains(mx, my))
                newEditHover = 34;

            // Result action buttons hover (600 + resultIndex*10 + buttonIndex)
            if (newEditHover == -1)
            {
                for (int bi = 0; bi < _aiResultButtonRects.Length; bi++)
                {
                    if (_aiResultButtonRects[bi].Contains(mx, my))
                    {
                        newEditHover = 600 + bi / 3 * 10 + bi % 3;
                        break;
                    }
                }
            }

        }
        else if (my > contentY)
        {
            // Find which item the mouse is over using variable heights
            int relY = my - contentY + _scrollOffset;
            int accum = 0;
            for (int i = 0; i < items.Count; i++)
            {
                int itemH = GetItemHeight(items[i]);
                if (relY >= accum && relY < accum + itemH)
                {
                    newHoverIndex = i;
                    break;
                }
                accum += itemH;
            }

            if (newHoverIndex >= 0)
            {
                int itemYStart = contentY + GetItemYOffset(newHoverIndex) - _scrollOffset;
                int itemH = GetItemHeight(items[newHoverIndex]);
                int btnY = itemYStart + (itemH - BUTTON_SIZE) / 2;

                if (newHoverIndex == _expandedIndex)
                {
                    // Expanded: 2 buttons (bookmark + delete) with wider spacing
                    int expandedGap = 20;
                    int btnStartX = clientWidth - BUTTON_SIZE * 2 - expandedGap - 12;
                    if (mx >= btnStartX && my >= btnY && my <= btnY + BUTTON_SIZE)
                    {
                        int relX = mx - btnStartX;
                        int btnStep = BUTTON_SIZE + expandedGap;
                        if (relX < BUTTON_SIZE)
                            newHoverButton = 0;
                        else if (relX >= btnStep && relX < btnStep + BUTTON_SIZE)
                            newHoverButton = 1;
                    }
                }
                else if (_inlineAiProcessing && newHoverIndex == _inlineAiItemIndex)
                {
                    // Inline AI processing: only cancel button (index 1)
                    int cancelBtnX = clientWidth - BUTTON_SIZE - 12;
                    if (mx >= cancelBtnX && mx <= cancelBtnX + BUTTON_SIZE && my >= btnY && my <= btnY + BUTTON_SIZE)
                        newHoverButton = 1;
                }
                else
                {
                    // Normal: paste(0), AI stars(1), more(2), arrow(3)
                    int buttonsWidth = (BUTTON_SIZE + BUTTON_MARGIN) * 3 + 14 + 8;
                    int btnStartX = clientWidth - buttonsWidth;
                    if (mx >= btnStartX && my >= btnY && my <= btnY + BUTTON_SIZE)
                    {
                        int relX = mx - btnStartX;
                        // Button 0: paste
                        if (relX < BUTTON_SIZE)
                            newHoverButton = 0;
                        else
                        {
                            relX -= BUTTON_SIZE + BUTTON_MARGIN;
                            // Button 1: AI stars
                            if (relX >= 0 && relX < BUTTON_SIZE)
                                newHoverButton = 1;
                            else
                            {
                                relX -= BUTTON_SIZE;
                                // Button 3: arrow
                                if (relX >= 0 && relX < 14)
                                    newHoverButton = 3;
                                else
                                {
                                    relX -= 14 + BUTTON_MARGIN;
                                    // Button 2: more
                                    if (relX >= 0 && relX < BUTTON_SIZE)
                                        newHoverButton = 2;
                                }
                            }
                        }
                    }
                }

                // Bubble task dropdown hover
                if (_bubbleTaskDropdownOpen && _bubbleTaskDropdownItemIndex == newHoverIndex)
                {
                    int itemYStart2 = contentY + GetItemYOffset(newHoverIndex) - _scrollOffset;
                    int itemH2 = GetItemHeight(items[newHoverIndex]);
                    int dropItemH = 28;
                    int taskCount = EditTasks.Length - 1;
                    int dropH = taskCount * dropItemH + 4;
                    int dropW = 160;
                    int dropX = clientWidth - 180;
                    int dropY = itemYStart2 + itemH2;
                    if (dropY + dropH > clientHeight) dropY = itemYStart2 - dropH;
                    if (mx >= dropX && mx < dropX + dropW && my >= dropY && my < dropY + dropH)
                    {
                        int idx = (my - dropY - 2) / dropItemH;
                        if (idx >= 0 && idx < taskCount)
                            _bubbleTaskHoverIndex = idx + 1; // +1 to skip "<unset>"
                        else
                            _bubbleTaskHoverIndex = -1;
                    }
                    else
                    {
                        _bubbleTaskHoverIndex = -1;
                    }
                }
                else
                {
                    _bubbleTaskHoverIndex = -1;
                }
            }
        }

        // Scrollbar hover detection
        bool newScrollbarHover = false;
        if (my > contentY)
        {
            int visibleHeight = clientHeight - contentY - 1;
            int totalHeight = GetTotalContentHeight();
            if (totalHeight > visibleHeight)
            {
                int scrollX = clientWidth - SCROLLBAR_WIDTH - 2;
                if (mx >= scrollX - 4 && mx <= scrollX + SCROLLBAR_WIDTH + 4)
                    newScrollbarHover = true;
            }
        }

        // Bubble task dropdown hover (even when mouse is outside the item area)
        if (_bubbleTaskDropdownOpen && _bubbleTaskDropdownItemIndex >= 0 && _bubbleTaskDropdownItemIndex < items.Count)
        {
            int itemYStart3 = contentY + GetItemYOffset(_bubbleTaskDropdownItemIndex) - _scrollOffset;
            int itemH3 = GetItemHeight(items[_bubbleTaskDropdownItemIndex]);
            int dropItemH3 = 28;
            int taskCount3 = EditTasks.Length - 1;
            int dropH3 = taskCount3 * dropItemH3 + 4;
            int dropW3 = 160;
            int dropX3 = clientWidth - 180;
            int dropY3 = itemYStart3 + itemH3;
            if (dropY3 + dropH3 > clientHeight) dropY3 = itemYStart3 - dropH3;
            if (mx >= dropX3 && mx < dropX3 + dropW3 && my >= dropY3 && my < dropY3 + dropH3)
            {
                int idx = (my - dropY3 - 2) / dropItemH3;
                if (idx >= 0 && idx < taskCount3)
                    _bubbleTaskHoverIndex = idx + 1;
                else
                    _bubbleTaskHoverIndex = -1;
                // Keep hover on the dropdown item so it stays visible
                newHoverIndex = _bubbleTaskDropdownItemIndex;
            }
            else
            {
                _bubbleTaskHoverIndex = -1;
            }
        }

        if (newHoverIndex != _hoverIndex || newHoverButton != _hoverButton || newHoverTab != _hoverTab || newPinHover != _pinHover || newCloseHover != _closeHover || newScrollbarHover != _scrollbarHover || newEditHover != _editHoverElement)
        {
            if (newHoverIndex != _hoverIndex)
            {
                _expandedIndex = -1;
                if (!_bubbleTaskDropdownOpen || _bubbleTaskDropdownItemIndex != newHoverIndex)
                {
                    _bubbleTaskDropdownOpen = false;
                    _bubbleTaskDropdownItemIndex = -1;
                    _bubbleTaskHoverIndex = -1;
                }
            }
            _hoverIndex = newHoverIndex;
            _hoverButton = newHoverButton;
            _hoverTab = newHoverTab;
            _pinHover = newPinHover;
            _closeHover = newCloseHover;
            _scrollbarHover = newScrollbarHover;
            _editHoverElement = newEditHover;

            // Compute tooltip text
            string? newTooltip = null;
            if (newPinHover)
                newTooltip = _alwaysOnTop ? "Unpin window" : "Pin window on top";
            else if (newCloseHover)
                newTooltip = "Close";
            else if (newEditHover == 20)
                newTooltip = _editTaskIndex == 0 ? "Task" : "Task: " + EditTasks[_editTaskIndex];
            else if (newEditHover == 21)
                newTooltip = _editToneIndex == 0 ? "Tone" : "Tone: " + EditTones[_editToneIndex];
            else if (newEditHover == 22)
                newTooltip = _editFormatIndex == 0 ? "Format" : "Format: " + EditFormats[_editFormatIndex];
            else if (newEditHover == 23)
                newTooltip = _editLengthIndex == 0 ? "Length" : "Length: " + EditLengths[_editLengthIndex];
            else if (newEditHover == 33)
                newTooltip = "Model: " + _editModels[_editModelIndex];
            else if (newEditHover == 34)
                newTooltip = _aiProcessing ? "Cancel" : "Send";
            else if (newEditHover >= 600)
            {
                int btn = (newEditHover - 600) % 10;
                newTooltip = btn switch { 0 => "Refine", 1 => "Copy", 2 => "Paste", _ => null };
            }
            else if (newHoverIndex >= 0 && newHoverButton >= 0)
            {
                if (newHoverIndex == _expandedIndex)
                    newTooltip = newHoverButton == 0 ? (_activeTab == 0 ? "Bookmark" : "Unbookmark") : "Delete";
                else
                    newTooltip = newHoverButton switch { 0 => "Paste as plain text", 1 => _inlineAiProcessing && newHoverIndex == _inlineAiItemIndex ? "Cancel" : "AI edit", 2 => "More actions", 3 => "Quick AI tasks", _ => null };
            }
            _tooltipText = newTooltip;
            _tooltipPos = new Point(mx, my);

            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        }
    }

    private void OnMouseDown(IntPtr lParam)
    {
        _tooltipText = null;
        int mx = NativeMethods.GET_X_LPARAM(lParam);
        int my = NativeMethods.GET_Y_LPARAM(lParam);

        NativeMethods.GetClientRect(_hWnd, out var cr);
        int clientWidth = cr.Right - cr.Left;

        // Check pin/close button click (header bar)
        if (my < HEADER_HEIGHT)
        {
            int closeX = clientWidth - PIN_BUTTON_SIZE - 6;
            int btnPinX = closeX - PIN_BUTTON_SIZE - 4;
            int btnY = (HEADER_HEIGHT - PIN_BUTTON_SIZE) / 2;
            if (mx >= closeX && mx <= closeX + PIN_BUTTON_SIZE && my >= btnY && my <= btnY + PIN_BUTTON_SIZE)
            {
                Hide();
            }
            else if (mx >= btnPinX && mx <= btnPinX + PIN_BUTTON_SIZE && my >= btnY && my <= btnY + PIN_BUTTON_SIZE)
            {
                _alwaysOnTop = !_alwaysOnTop;
                var insertAfter = _alwaysOnTop ? NativeMethods.HWND_TOPMOST : NativeMethods.HWND_NOTOPMOST;
                NativeMethods.SetWindowPos(_hWnd, insertAfter, 0, 0, 0, 0,
                    NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOACTIVATE);
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            }
            return;
        }

        // Check tab clicks
        if (my >= HEADER_HEIGHT && my < TOP_OFFSET)
        {
            int tabCount = 3;
            int tabW = clientWidth / tabCount;
            int adjustedX = Math.Max(mx - 1, 0); // account for 1px border offset in drawing
            int newTab = Math.Clamp(adjustedX / tabW, 0, tabCount - 1);
            if (newTab != _activeTab)
            {
                _activeTab = newTab;
                _scrollOffset = 0;
                _hoverIndex = -1;
                _hoverButton = -1;
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            }
            return;
        }

        // Check scrollbar click
        int clientHeight = cr.Bottom - cr.Top;
        int contentY = TOP_OFFSET + 1;

        // Edit tab: check edit textbox scrollbar click
        if (_activeTab == 2 && my > contentY)
        {
            var r = _editTextBoxRect;
            int scrollbarW = 6;
            int sbX = r.X + r.Width - scrollbarW - 6;
            if (mx >= sbX - 4 && mx <= sbX + scrollbarW + 4 && my >= r.Y && my < r.Y + r.Height)
            {
                _editScrollbarDragging = true;
                _editScrollbarDragStartY = my;
                _editScrollbarDragStartOffset = _editTextScrollOffset;
                NativeMethods.SetCapture(_hWnd);
                return;
            }

            // Check result textbox scrollbar clicks
            for (int ri = 0; ri < _aiResultTextBoxRects.Length; ri++)
            {
                var rr = _aiResultTextBoxRects[ri];
                int rsbX = rr.X + rr.Width - 6 - 6;
                if (mx >= rsbX - 4 && mx <= rsbX + 6 + 4 && my >= rr.Y && my < rr.Y + rr.Height)
                {
                    _aiResultScrollbarDragging = true;
                    _aiResultScrollbarDragIndex = ri;
                    _aiResultScrollbarDragStartY = my;
                    _aiResultScrollbarDragStartOffset = ri < _aiResultScrollOffsets.Length ? _aiResultScrollOffsets[ri] : 0;
                    NativeMethods.SetCapture(_hWnd);
                    return;
                }
            }

            // Check choices slider click
            var sr = _editChoicesSliderRect;
            if (sr.Width > 0 && mx >= sr.X && mx < sr.X + sr.Width && my >= sr.Y && my < sr.Y + sr.Height)
            {
                float ratio = Math.Clamp((float)(mx - sr.X) / sr.Width, 0f, 1f);
                _editChoices = Math.Clamp((int)Math.Round(ratio * 4) + 1, 1, 5);
                _editChoicesDragging = true;
                NativeMethods.SetCapture(_hWnd);
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                return;
            }
        }

        if (_activeTab != 2 && my > contentY)
        {
            int visibleHeight = clientHeight - contentY - 1;
            int totalHeight = GetTotalContentHeight();
            if (totalHeight > visibleHeight)
            {
                int scrollX = clientWidth - SCROLLBAR_WIDTH - 2;
                if (mx >= scrollX - 4)
                {
                    _scrollbarDragging = true;
                    _scrollbarDragStartY = my;
                    _scrollbarDragStartOffset = _scrollOffset;
                    NativeMethods.SetCapture(_hWnd);
                    return;
                }
            }
        }

        // Edit tab click handling
        if (_activeTab == 2)
        {
            bool needsRepaint = false;
            // Handle dropdown item selection first
            if (_editHoverElement >= 100 && _editHoverElement < 100 + EditTasks.Length)
            {
                _editTaskIndex = _editHoverElement - 100;
                _editTaskDropdownOpen = false;
                _settings.EditTaskIndex = _editTaskIndex; _settings.Save();
                needsRepaint = true;
            }
            else if (_editHoverElement >= 200 && _editHoverElement < 200 + EditTones.Length)
            {
                _editToneIndex = _editHoverElement - 200;
                _editToneDropdownOpen = false;
                _settings.EditToneIndex = _editToneIndex; _settings.Save();
                needsRepaint = true;
            }
            else if (_editHoverElement >= 300 && _editHoverElement < 300 + EditFormats.Length)
            {
                _editFormatIndex = _editHoverElement - 300;
                _editFormatDropdownOpen = false;
                _settings.EditFormatIndex = _editFormatIndex; _settings.Save();
                needsRepaint = true;
            }
            else if (_editHoverElement >= 400 && _editHoverElement < 400 + EditLengths.Length)
            {
                _editLengthIndex = _editHoverElement - 400;
                _editLengthDropdownOpen = false;
                _settings.EditLengthIndex = _editLengthIndex; _settings.Save();
                needsRepaint = true;
            }
            else if (_editHoverElement >= 500 && _editHoverElement < 500 + _editModels.Length)
            {
                int newModel = _editHoverElement - 500;
                if (newModel != _editModelIndex)
                {
                    _editModelIndex = newModel;
                    _settings.Model = _editModels[_editModelIndex];
                    _settings.Save();
                    RecreateCopilotSession();
                }
                _editModelDropdownOpen = false;
                needsRepaint = true;
            }
            else if (_editHoverElement == 34)
            {
                // Send or Cancel button
                if (_aiProcessing)
                    CancelAiRequest();
                else
                    SendAiRequest();
                needsRepaint = true;
            }
            else if (_editHoverElement >= 600)
            {
                int ri = (_editHoverElement - 600) / 10;
                int btn = (_editHoverElement - 600) % 10;
                if (ri >= 0 && ri < _aiResults.Count)
                {
                    switch (btn)
                    {
                        case 0: // Refine — copy result to input textbox
                            _editItem = new ClipboardItem { Type = ClipboardItemType.Text, Text = _aiResults[ri] };
                            _editTextScrollOffset = 0;
                            _aiResults.Clear();
                            _aiResultScrollOffsets = [];
                            AdjustWindowHeightForResults();
                            break;
                        case 1: // Copy to clipboard
                            SetClipboardText(_aiResults[ri]);
                            break;
                        case 2: // Paste
                            _monitor.SetIgnoreNext();
                            SetClipboardText(_aiResults[ri]);
                            SchedulePaste();
                            break;
                    }
                }
                needsRepaint = true;
            }
            else
            {
                switch (_editHoverElement)
                {
                    case 20: _editTaskDropdownOpen = !_editTaskDropdownOpen; _editToneDropdownOpen = false; _editFormatDropdownOpen = false; _editLengthDropdownOpen = false; _editModelDropdownOpen = false; needsRepaint = true; break;
                    case 21: _editToneDropdownOpen = !_editToneDropdownOpen; _editTaskDropdownOpen = false; _editFormatDropdownOpen = false; _editLengthDropdownOpen = false; _editModelDropdownOpen = false; needsRepaint = true; break;
                    case 22: _editFormatDropdownOpen = !_editFormatDropdownOpen; _editTaskDropdownOpen = false; _editToneDropdownOpen = false; _editLengthDropdownOpen = false; _editModelDropdownOpen = false; needsRepaint = true; break;
                    case 23: _editLengthDropdownOpen = !_editLengthDropdownOpen; _editTaskDropdownOpen = false; _editToneDropdownOpen = false; _editFormatDropdownOpen = false; _editModelDropdownOpen = false; needsRepaint = true; break;
                    case 33: _editModelDropdownOpen = !_editModelDropdownOpen; _editTaskDropdownOpen = false; _editToneDropdownOpen = false; _editFormatDropdownOpen = false; _editLengthDropdownOpen = false; needsRepaint = true; break;

                    default:
                        if (_editTaskDropdownOpen || _editToneDropdownOpen || _editFormatDropdownOpen || _editLengthDropdownOpen || _editModelDropdownOpen) { _editTaskDropdownOpen = false; _editToneDropdownOpen = false; _editFormatDropdownOpen = false; _editLengthDropdownOpen = false; _editModelDropdownOpen = false; needsRepaint = true; }
                        break;
                }
            }
            if (needsRepaint)
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            return;
        }

        var items = ActiveItems;
        if (_hoverIndex < 0 || _hoverIndex >= items.Count)
        {
            // Check if clicking in the bubble task dropdown area
            if (_bubbleTaskDropdownOpen && _bubbleTaskHoverIndex > 0)
            {
                int taskIdx = _bubbleTaskHoverIndex;
                int itemIdx = _bubbleTaskDropdownItemIndex;
                _bubbleTaskDropdownOpen = false;
                _bubbleTaskDropdownItemIndex = -1;
                _bubbleTaskHoverIndex = -1;
                var items4 = ActiveItems;
                if (itemIdx >= 0 && itemIdx < items4.Count)
                    StartInlineAiRequest(itemIdx, taskIdx);
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            }
            return;
        }

        // Handle bubble task dropdown click first
        if (_bubbleTaskDropdownOpen && _bubbleTaskHoverIndex > 0)
        {
            int taskIdx = _bubbleTaskHoverIndex;
            int itemIdx = _bubbleTaskDropdownItemIndex;
            _bubbleTaskDropdownOpen = false;
            _bubbleTaskDropdownItemIndex = -1;
            _bubbleTaskHoverIndex = -1;
            if (itemIdx >= 0 && itemIdx < items.Count)
                StartInlineAiRequest(itemIdx, taskIdx);
            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            return;
        }

        // Close bubble dropdown if clicking elsewhere
        if (_bubbleTaskDropdownOpen)
        {
            _bubbleTaskDropdownOpen = false;
            _bubbleTaskDropdownItemIndex = -1;
            _bubbleTaskHoverIndex = -1;
            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        }

        if (_hoverIndex == _expandedIndex)
        {
            // Expanded mode: bookmark/unbookmark (0), delete (1)
            if (_hoverButton == 0)
            {
                if (_activeTab == 0)
                    _monitor.BookmarkItem(_hoverIndex);
                else
                    _monitor.UnbookmarkItem(_hoverIndex);
                _expandedIndex = -1;
            }
            else if (_hoverButton == 1)
            {
                if (_activeTab == 0)
                    _monitor.DeleteItem(_hoverIndex);
                else
                    _monitor.DeleteBookmarkedItem(_hoverIndex);
                _expandedIndex = -1;
                _hoverIndex = -1;
                _hoverButton = -1;
            }
            else
            {
                // Click on item content while expanded = collapse
                _expandedIndex = -1;
            }
            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        }
        else if (_hoverButton == 0) // Paste as text
        {
            PasteAsText(_hoverIndex);
        }
        else if (_hoverButton == 1 && _inlineAiProcessing && _hoverIndex == _inlineAiItemIndex) // Cancel inline AI
        {
            CancelInlineAiRequest();
            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        }
        else if (_hoverButton == 1) // AI edit — open edit tab
        {
            var items2 = ActiveItems;
            if (_hoverIndex >= 0 && _hoverIndex < items2.Count)
            {
                _editItem = items2[_hoverIndex];
                _editTextScrollOffset = 0;
                _editTaskDropdownOpen = false;
                _editToneDropdownOpen = false;
                _editFormatDropdownOpen = false;
                _editLengthDropdownOpen = false;
                _activeTab = 2;
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
            }
        }
        else if (_hoverButton == 2) // More options → expand
        {
            _expandedIndex = _hoverIndex;
            _hoverButton = -1;
            _bubbleTaskDropdownOpen = false;
            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        }
        else if (_hoverButton == 3) // Arrow → toggle bubble task dropdown
        {
            if (_bubbleTaskDropdownOpen && _bubbleTaskDropdownItemIndex == _hoverIndex)
            {
                _bubbleTaskDropdownOpen = false;
                _bubbleTaskDropdownItemIndex = -1;
            }
            else
            {
                _bubbleTaskDropdownOpen = true;
                _bubbleTaskDropdownItemIndex = _hoverIndex;
            }
            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        }
        else if (_hoverButton == -1 && _bubbleTaskDropdownOpen && _bubbleTaskHoverIndex > 0)
        {
            // Clicked on a task in the bubble dropdown
            int taskIdx = _bubbleTaskHoverIndex;
            int itemIdx = _bubbleTaskDropdownItemIndex;
            _bubbleTaskDropdownOpen = false;
            _bubbleTaskDropdownItemIndex = -1;
            _bubbleTaskHoverIndex = -1;
            var items3 = ActiveItems;
            if (itemIdx >= 0 && itemIdx < items3.Count)
                StartInlineAiRequest(itemIdx, taskIdx);
            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        }
        else
        {
            // Click on item content = copy to clipboard and paste
            PasteItem(_hoverIndex);
        }
    }

    private void PasteItem(int index)
    {
        var items = ActiveItems;
        if (index < 0 || index >= items.Count)
            return;

        var item = items[index];
        _monitor.SetIgnoreNext();

        if (item.Type == ClipboardItemType.Image && item.Image != null)
            SetClipboardImage(item.Image);
        else
            SetClipboardText(item.Text);

        SchedulePaste();
    }

    private void PasteAsText(int index)
    {
        var items = ActiveItems;
        if (index < 0 || index >= items.Count)
            return;

        var item = items[index];

        if (item.Type == ClipboardItemType.Image && item.Image != null)
        {
            ExtractTextFromImage(index);
            return;
        }

        _monitor.SetIgnoreNext();
        SetClipboardText(item.Text);
        SchedulePaste();
    }

    private async void ExtractTextFromImage(int itemIndex)
    {
        if (_copilotSession == null || _inlineAiProcessing)
            return;

        var items = ActiveItems;
        if (itemIndex < 0 || itemIndex >= items.Count)
            return;

        var item = items[itemIndex];
        if (item.Type != ClipboardItemType.Image || item.Image == null)
            return;

        // Capture target window now — after the async wait it will be stale
        var targetWindow = NativeMethods.GetForegroundWindow();
        if (targetWindow == _hWnd)
            targetWindow = _previousForegroundWindow;

        _inlineAiProcessing = true;
        _inlineAiItemIndex = itemIndex;
        _aiAnimPhase = 0;
        NativeMethods.SetTimer(_hWnd, ANIM_TIMER_ID, 30, IntPtr.Zero);
        StartTrayAnimation();
        _inlineAiCancellation = new CancellationTokenSource();
        NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);

        string? extractedText = null;
        try
        {
            var ct = _inlineAiCancellation.Token;

            string base64;
            using (var ms = new MemoryStream())
            {
                item.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                if (ms.Length > 10 * 1024 * 1024) // 10 MB limit
                    return;
                base64 = Convert.ToBase64String(ms.ToArray());
            }

            var response = await _copilotSession.SendAndWaitAsync(
                new MessageOptions
                {
                    Prompt = "Extract all text from this image. Respond ONLY with the extracted text exactly as it appears. No explanations, no markdown fences, no preamble.",
                    Attachments = [new UserMessageDataAttachmentsItemBlob
                    {
                        Type = "blob",
                        Data = base64,
                        MimeType = "image/png",
                        DisplayName = "clipboard-image.png"
                    }]
                }, null, ct);

            if (response?.Data?.Content is { } content && content.Length > 0 && !ct.IsCancellationRequested)
                extractedText = content;
        }
        catch (OperationCanceledException) { }
        catch { }
        finally
        {
            _inlineAiProcessing = false;
            _inlineAiItemIndex = -1;
            _inlineAiCancellation?.Dispose();
            _inlineAiCancellation = null;
            NativeMethods.KillTimer(_hWnd, ANIM_TIMER_ID);
            StopTrayAnimation();
            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        }

        // Marshal paste back to the UI thread — after await we're on a thread pool thread
        if (extractedText != null)
        {
            _extractedText = extractedText;
            _extractPasteTarget = targetWindow;
            NativeMethods.PostMessageW(_hWnd, NativeMethods.WM_APP_EXTRACTPASTE, IntPtr.Zero, IntPtr.Zero);
        }
    }

    private void SchedulePaste()
    {
        // With WS_EX_NOACTIVATE, the target window still has focus.
        // Hide first (unless always-on-top), then use a timer to send Ctrl+V
        // after the mouse click has fully processed.
        if (!_alwaysOnTop)
            Hide();

        NativeMethods.SetTimer(_hWnd, PASTE_TIMER_ID, 50, IntPtr.Zero);
    }

    private static void ReleaseModifierKeys()
    {
        // Release Win, Shift, Alt, Ctrl if they appear stuck in the input state
        ushort[] mods = [0x5B, 0x5C, 0x10, 0x11, 0x12]; // LWin, RWin, Shift, Ctrl, Alt
        foreach (var vk in mods)
        {
            if ((NativeMethods.GetAsyncKeyState(vk) & 0x8000) != 0)
            {
                var input = new NativeMethods.INPUT
                {
                    type = NativeMethods.INPUT_KEYBOARD,
                    u = { ki = { wVk = vk, dwFlags = NativeMethods.KEYEVENTF_KEYUP } }
                };
                NativeMethods.SendInput(1, [input], Marshal.SizeOf<NativeMethods.INPUT>());
            }
        }
    }

    private void SetClipboardText(string text)
    {
        if (!NativeMethods.OpenClipboard(_hWnd))
            return;

        try
        {
            NativeMethods.EmptyClipboard();
            var bytes = (text.Length + 1) * 2;
            var hMem = NativeMethods.GlobalAlloc(NativeMethods.GMEM_MOVEABLE, (nuint)bytes);
            if (hMem == IntPtr.Zero) return;

            var ptr = NativeMethods.GlobalLock(hMem);
            if (ptr == IntPtr.Zero) { NativeMethods.GlobalFree(hMem); return; }

            try
            {
                Marshal.Copy(text.ToCharArray(), 0, ptr, text.Length);
                Marshal.WriteInt16(ptr, text.Length * 2, 0); // null terminator
            }
            finally
            {
                NativeMethods.GlobalUnlock(hMem);
            }

            NativeMethods.SetClipboardData(NativeMethods.CF_UNICODETEXT, hMem);
        }
        finally
        {
            NativeMethods.CloseClipboard();
        }
    }

    private void SetClipboardImage(Bitmap image)
    {
        if (!NativeMethods.OpenClipboard(_hWnd))
            return;

        try
        {
            NativeMethods.EmptyClipboard();

            // Convert to 32bpp ARGB for maximum compatibility
            using var bmp32 = new Bitmap(image.Width, image.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp32))
                g.DrawImage(image, 0, 0, image.Width, image.Height);

            var bmpData = bmp32.LockBits(
                new Rectangle(0, 0, bmp32.Width, bmp32.Height),
                System.Drawing.Imaging.ImageLockMode.ReadOnly,
                System.Drawing.Imaging.PixelFormat.Format32bppArgb);

            int stride = Math.Abs(bmpData.Stride);
            int imageSize = stride * bmp32.Height;
            int headerSize = 40; // BITMAPINFOHEADER size
            int totalSize = headerSize + imageSize;

            var hMem = NativeMethods.GlobalAlloc(NativeMethods.GMEM_MOVEABLE, (nuint)totalSize);
            if (hMem != IntPtr.Zero)
            {
                var ptr = NativeMethods.GlobalLock(hMem);
                if (ptr == IntPtr.Zero) { NativeMethods.GlobalFree(hMem); }
                else
                {
                    // Write BITMAPINFOHEADER
                    Marshal.WriteInt32(ptr, 0, headerSize);       // biSize
                    Marshal.WriteInt32(ptr, 4, bmp32.Width);      // biWidth
                    Marshal.WriteInt32(ptr, 8, bmp32.Height);     // biHeight (positive = bottom-up)
                    Marshal.WriteInt16(ptr, 12, 1);               // biPlanes
                    Marshal.WriteInt16(ptr, 14, 32);              // biBitCount
                    Marshal.WriteInt32(ptr, 16, 0);               // biCompression (BI_RGB)
                    Marshal.WriteInt32(ptr, 20, imageSize);       // biSizeImage
                    Marshal.WriteInt32(ptr, 24, 0);               // biXPelsPerMeter
                    Marshal.WriteInt32(ptr, 28, 0);               // biYPelsPerMeter
                    Marshal.WriteInt32(ptr, 32, 0);               // biClrUsed
                    Marshal.WriteInt32(ptr, 36, 0);               // biClrImportant

                    // Copy pixel data (flip vertically — DIB is bottom-up)
                    for (int y = 0; y < bmp32.Height; y++)
                    {
                        var srcRow = bmpData.Scan0 + y * bmpData.Stride;
                        var dstRow = ptr + headerSize + (bmp32.Height - 1 - y) * stride;
                        unsafe
                        {
                            Buffer.MemoryCopy(srcRow.ToPointer(), dstRow.ToPointer(), stride, stride);
                        }
                    }

                    NativeMethods.GlobalUnlock(hMem);
                    NativeMethods.SetClipboardData(NativeMethods.CF_DIB, hMem);
                }
            }

            bmp32.UnlockBits(bmpData);
        }
        catch
        {
            // Silently fail
        }
        finally
        {
            NativeMethods.CloseClipboard();
        }
    }

    private static void SimulatePaste()
    {
        var inputs = new NativeMethods.INPUT[4];

        // Ctrl down
        inputs[0].type = NativeMethods.INPUT_KEYBOARD;
        inputs[0].u.ki.wVk = NativeMethods.VK_CONTROL;

        // V down
        inputs[1].type = NativeMethods.INPUT_KEYBOARD;
        inputs[1].u.ki.wVk = NativeMethods.VK_V;

        // V up
        inputs[2].type = NativeMethods.INPUT_KEYBOARD;
        inputs[2].u.ki.wVk = NativeMethods.VK_V;
        inputs[2].u.ki.dwFlags = NativeMethods.KEYEVENTF_KEYUP;

        // Ctrl up
        inputs[3].type = NativeMethods.INPUT_KEYBOARD;
        inputs[3].u.ki.wVk = NativeMethods.VK_CONTROL;
        inputs[3].u.ki.dwFlags = NativeMethods.KEYEVENTF_KEYUP;

        NativeMethods.SendInput(4, inputs, Marshal.SizeOf<NativeMethods.INPUT>());
    }

    private static void ForceForegroundWindow(IntPtr target)
    {
        if (target == IntPtr.Zero)
            return;

        var foreground = NativeMethods.GetForegroundWindow();
        if (foreground == target)
            return;

        uint foreThread = NativeMethods.GetWindowThreadProcessId(foreground, out _);
        uint appThread = NativeMethods.GetCurrentThreadId();

        if (foreThread != appThread)
            NativeMethods.AttachThreadInput(foreThread, appThread, true);

        NativeMethods.SetForegroundWindow(target);

        if (foreThread != appThread)
            NativeMethods.AttachThreadInput(foreThread, appThread, false);
    }

    private void OnClipboardChanged()
    {
        if (_visible)
        {
            if (_activeTab != 2)
                _activeTab = 0;
            NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
        }
        else if (_monitor.Items.Count > 0)
        {
            _notification.Show(_monitor.Items[0]);
        }
    }

    private static string FormatTimestamp(DateTime ts)
    {
        var diff = DateTime.Now - ts;
        if (diff.TotalSeconds < 60) return "just now";
        if (diff.TotalMinutes < 60) return $"{(int)diff.TotalMinutes}m ago";
        if (diff.TotalHours < 24) return $"{(int)diff.TotalHours}h ago";
        return ts.ToString("MMM d, HH:mm");
    }

    // Icon loading
    private void LoadAppIcon()
    {
        try
        {
            var stream = typeof(MainWindow).Assembly.GetManifestResourceStream("ClipSidekick.app.ico");
            if (stream != null)
            {
                var icon = new Icon(stream);
                stream.Dispose();

                _appIconSmall = new Icon(icon, NativeMethods.GetSystemMetrics(NativeMethods.SM_CXSMICON),
                                               NativeMethods.GetSystemMetrics(NativeMethods.SM_CYSMICON));
                _appIconLarge = new Icon(icon, 32, 32);
                icon.Dispose();

                NativeMethods.SendMessageW(_hWnd, NativeMethods.WM_SETICON, (IntPtr)NativeMethods.ICON_SMALL, _appIconSmall.Handle);
                NativeMethods.SendMessageW(_hWnd, NativeMethods.WM_SETICON, (IntPtr)NativeMethods.ICON_BIG, _appIconLarge.Handle);
            }
        }
        catch
        {
            // Fall back to default icon
        }
    }

    // Tray icon management
    private void AddTrayIcon()
    {
        var iconHandle = _appIconSmall?.Handle ?? IntPtr.Zero;

        // Fallback to system default if embedded icon not loaded
        if (iconHandle == IntPtr.Zero)
        {
            iconHandle = NativeMethods.LoadImageW(
                IntPtr.Zero, (IntPtr)NativeMethods.IDI_APPLICATION,
                NativeMethods.IMAGE_ICON,
                NativeMethods.GetSystemMetrics(NativeMethods.SM_CXSMICON),
                NativeMethods.GetSystemMetrics(NativeMethods.SM_CYSMICON),
                NativeMethods.LR_SHARED);
        }

        _trayIconData = new NativeMethods.NOTIFYICONDATA
        {
            cbSize = Marshal.SizeOf<NativeMethods.NOTIFYICONDATA>(),
            hWnd = _hWnd,
            uID = TRAY_ID,
            uFlags = NativeMethods.NIF_MESSAGE | NativeMethods.NIF_ICON | NativeMethods.NIF_TIP,
            uCallbackMessage = NativeMethods.WM_TRAY_ICON,
            hIcon = iconHandle,
            szTip = $"Clip Sidekick ({_settings.HotkeyDisplay})"
        };

        NativeMethods.Shell_NotifyIconW(NativeMethods.NIM_ADD, ref _trayIconData);
    }

    private void RemoveTrayIcon()
    {
        NativeMethods.Shell_NotifyIconW(NativeMethods.NIM_DELETE, ref _trayIconData);
    }

    private void BuildTrayAnimationFrames()
    {
        int sz = NativeMethods.GetSystemMetrics(NativeMethods.SM_CXSMICON);
        int frameCount = 12;
        _trayAnimFrames = new Icon?[frameCount];

        // Get the base icon bitmap
        using var baseBmp = new Bitmap(sz, sz);
        if (_appIconSmall != null)
        {
            using (var g = Graphics.FromImage(baseBmp))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawIcon(_appIconSmall, new Rectangle(0, 0, sz, sz));
            }
        }

        // Read base pixels once using LockBits
        var baseLock = baseBmp.LockBits(
            new Rectangle(0, 0, sz, sz),
            System.Drawing.Imaging.ImageLockMode.ReadOnly,
            System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        var basePixels = new byte[sz * sz * 4];
        Marshal.Copy(baseLock.Scan0, basePixels, 0, basePixels.Length);
        baseBmp.UnlockBits(baseLock);

        for (int i = 0; i < frameCount; i++)
        {
            float hueShift = i * (360f / frameCount);
            var tinted = new Bitmap(sz, sz, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            var tLock = tinted.LockBits(
                new Rectangle(0, 0, sz, sz),
                System.Drawing.Imaging.ImageLockMode.WriteOnly,
                System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            var tPixels = new byte[basePixels.Length];

            for (int p = 0; p < tPixels.Length; p += 4)
            {
                byte bv = basePixels[p], gv = basePixels[p + 1], rv = basePixels[p + 2], av = basePixels[p + 3];
                if (av == 0) continue;
                RgbToHsl(rv, gv, bv, out float h, out float s, out float l);
                h = (h + hueShift) % 360f;
                s = Math.Min(1f, s + 0.3f);
                var (r, g, b) = HslToRgb(h, s, l);
                tPixels[p] = (byte)b; tPixels[p + 1] = (byte)g; tPixels[p + 2] = (byte)r; tPixels[p + 3] = av;
            }

            Marshal.Copy(tPixels, 0, tLock.Scan0, tPixels.Length);
            tinted.UnlockBits(tLock);
            IntPtr hIcon = tinted.GetHicon();
            _trayAnimFrames[i] = Icon.FromHandle(hIcon);
            tinted.Dispose();
        }
    }

    private static void RgbToHsl(int r, int g, int b, out float h, out float s, out float l)
    {
        float rf = r / 255f, gf = g / 255f, bf = b / 255f;
        float max = Math.Max(rf, Math.Max(gf, bf));
        float min = Math.Min(rf, Math.Min(gf, bf));
        l = (max + min) / 2f;
        if (max == min) { h = 0; s = 0; return; }
        float d = max - min;
        s = l > 0.5f ? d / (2f - max - min) : d / (max + min);
        if (max == rf) h = ((gf - bf) / d + (gf < bf ? 6 : 0)) * 60f;
        else if (max == gf) h = ((bf - rf) / d + 2) * 60f;
        else h = ((rf - gf) / d + 4) * 60f;
    }

    private static (int r, int g, int b) HslToRgb(float h, float s, float l)
    {
        if (s == 0) { int v = (int)(l * 255); return (v, v, v); }
        float c = (1f - Math.Abs(2f * l - 1f)) * s;
        float x = c * (1f - Math.Abs((h / 60f) % 2f - 1f));
        float m = l - c / 2f;
        float rf, gf, bf;
        if (h < 60) { rf = c; gf = x; bf = 0; }
        else if (h < 120) { rf = x; gf = c; bf = 0; }
        else if (h < 180) { rf = 0; gf = c; bf = x; }
        else if (h < 240) { rf = 0; gf = x; bf = c; }
        else if (h < 300) { rf = x; gf = 0; bf = c; }
        else { rf = c; gf = 0; bf = x; }
        return ((int)((rf + m) * 255), (int)((gf + m) * 255), (int)((bf + m) * 255));
    }

    private void StartTrayAnimation()
    {
        if (_trayAnimFrames.Length == 0)
            BuildTrayAnimationFrames();
        _trayAnimFrame = 0;
        NativeMethods.SetTimer(_hWnd, TRAY_ANIM_TIMER_ID, 100, IntPtr.Zero);
    }

    private void StopTrayAnimation()
    {
        NativeMethods.KillTimer(_hWnd, TRAY_ANIM_TIMER_ID);
        // Restore original icon
        _trayIconData.hIcon = _appIconSmall?.Handle ?? IntPtr.Zero;
        _trayIconData.uFlags = NativeMethods.NIF_ICON;
        NativeMethods.Shell_NotifyIconW(NativeMethods.NIM_MODIFY, ref _trayIconData);

        // Destroy animation icon handles to prevent GDI handle leak
        foreach (var frame in _trayAnimFrames)
        {
            if (frame != null)
                NativeMethods.DestroyIcon(frame.Handle);
        }
        _trayAnimFrames = [];
    }

    private void OnTrayIcon(IntPtr lParam)
    {
        int msg = lParam.ToInt32();
        switch (msg)
        {
            case NativeMethods.WM_LBUTTONDBLCLK:
                ToggleWindow();
                break;
            case NativeMethods.WM_RBUTTONUP:
                ShowTrayMenu();
                break;
        }
    }

    private void ShowTrayMenu()
    {
        var hMenu = NativeMethods.CreatePopupMenu();
        NativeMethods.AppendMenuW(hMenu, NativeMethods.MF_STRING, TRAY_MENU_SHOW, "Show/Hide");
        NativeMethods.AppendMenuW(hMenu, NativeMethods.MF_STRING, TRAY_MENU_ONTOP, _alwaysOnTop ? "✓ Always on Top" : "  Always on Top");
        NativeMethods.AppendMenuW(hMenu, NativeMethods.MF_SEPARATOR, 0, null);
        NativeMethods.AppendMenuW(hMenu, NativeMethods.MF_STRING, TRAY_MENU_CLEAR, "Clear History");
        NativeMethods.AppendMenuW(hMenu, NativeMethods.MF_SEPARATOR, 0, null);
        NativeMethods.AppendMenuW(hMenu, NativeMethods.MF_STRING, TRAY_MENU_SETTINGS, "Settings...");
        NativeMethods.AppendMenuW(hMenu, NativeMethods.MF_SEPARATOR, 0, null);
        NativeMethods.AppendMenuW(hMenu, NativeMethods.MF_STRING, TRAY_MENU_EXIT, "Exit");

        NativeMethods.GetCursorPos(out var pt);
        NativeMethods.SetForegroundWindow(_hWnd);

        int cmd = NativeMethods.TrackPopupMenu(hMenu,
            NativeMethods.TPM_RIGHTBUTTON | NativeMethods.TPM_RETURNCMD,
            pt.X, pt.Y, 0, _hWnd, IntPtr.Zero);

        NativeMethods.DestroyMenu(hMenu);
        OnCommand(cmd);
    }

    private void OnCommand(int cmdId)
    {
        switch (cmdId)
        {
            case TRAY_MENU_SHOW:
                ToggleWindow();
                break;
            case TRAY_MENU_ONTOP:
                _alwaysOnTop = !_alwaysOnTop;
                if (_visible)
                {
                    var insertAfter = _alwaysOnTop ? NativeMethods.HWND_TOPMOST : NativeMethods.HWND_NOTOPMOST;
                    NativeMethods.SetWindowPos(_hWnd, insertAfter, 0, 0, 0, 0,
                        NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOACTIVATE);
                }
                NativeMethods.InvalidateRect(_hWnd, IntPtr.Zero, false);
                break;
            case TRAY_MENU_CLEAR:
                _monitor.ClearUnbookmarked();
                break;
            case TRAY_MENU_SETTINGS:
                ShowSettingsDialog();
                break;
            case TRAY_MENU_EXIT:
                NativeMethods.DestroyWindow(_hWnd);
                break;
        }
    }

    private void RegisterConfiguredHotkey()
    {
        var (mods, vk) = _settings.ParseHotkey();
        NativeMethods.RegisterHotKey(_hWnd, HOTKEY_ID, mods, vk);
    }

    private void ShowHotkeyDialog()
    {
        // Create a small modal-like window to capture a new hotkey
        var dialog = new HotkeyDialog();
        var result = dialog.Show(_hWnd, _settings.Hotkey);
        if (result != null && result != _settings.Hotkey)
        {
            // Unregister old, save new, register new
            NativeMethods.UnregisterHotKey(_hWnd, HOTKEY_ID);
            _settings.Hotkey = result;
            _settings.Save();
            RegisterConfiguredHotkey();
        }
    }

    private void ShowSettingsDialog()
    {
        var oldHotkey = _settings.Hotkey;
        var dialog = new SettingsDialog();
        if (dialog.Show(_hWnd, _settings))
        {
            // Apply notification duration
            _notification.DurationMs = _settings.NotificationDurationMs;

            // Apply max history items
            _monitor.MaxItems = _settings.MaxHistoryItems;

            // Re-register hotkey if changed
            if (_settings.Hotkey != oldHotkey)
            {
                NativeMethods.UnregisterHotKey(_hWnd, HOTKEY_ID);
                RegisterConfiguredHotkey();
            }
        }
    }
}
