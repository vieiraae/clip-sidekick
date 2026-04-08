# Clip Sidekick

A lightweight clipboard history manager for Windows with built-in AI editing powered by [GitHub Copilot SDK](https://github.com/github/copilot-sdk).

Copy text or screenshots as you work. Press **Win+Shift+V** to browse your clipboard history, pin items as bookmarks, and use AI to proofread, rewrite, summarize, or transform any clipboard content — all without leaving your current app.

## What it does

**Clipboard history** — Keeps your last 50 clipboard entries (text and images), accessible via a global hotkey. Click any item to paste it, or bookmark items you use often.

**AI editing** — Select a clipboard item and apply AI tasks like *Proofread*, *Rewrite*, *Summarize*, *Expand*, or *Use synonyms*. Fine-tune with tone (Professional, Casual, Technical…), format (Paragraphs, List, Table, JSON…), and length preferences. Generate up to 5 variations at once.

**Image-to-text** — Paste a screenshot as text. The AI extracts readable text from clipboard images automatically.

**Non-intrusive** — The popup appears near your cursor without stealing focus from your active window, just like the built-in Windows clipboard history. The cursor keeps blinking in Notepad.

## Features

- **History tab** — Browse recent clipboard items (text + images) with timestamps
- **Bookmarks tab** — Pin frequently used items; persisted across sessions
- **Edit tab** — Full AI editing interface with task, tone, format, length, and variation controls
- **Notification bubble** — Toast popup on every capture with quick AI task access
- **Paste automation** — AI results are automatically pasted into your active app
- **Tray icon** — Animated during AI processing; right-click for settings
- **Dark theme** — Modern dark UI with rounded corners and smooth animations
- **DPI-aware** — Scales properly across monitors with different DPI settings
- **Configurable hotkey** — Change from Win+Shift+V to any key combination

### AI tasks

| Task | Description |
|------|-------------|
| Proofread | Fix grammar and spelling |
| Rewrite | Rephrase the text |
| Use synonyms | Replace words with alternatives |
| Minor revise | Light editing pass |
| Major revise | Substantial rewrite |
| Describe | Add description or context |
| Answer | Transform into Q&A format |
| Explain | Clarify the content |
| Expand | Add more detail |
| Summarize | Create a concise version |

## Getting started

### Prerequisites

Install the .NET 10 Desktop Runtime and GitHub CLI:

```
winget install Microsoft.DotNet.DesktopRuntime.10
winget install GitHub.cli
```

Authenticate with GitHub and install the Copilot extension:

```
gh auth login
gh extension install github/gh-copilot
```

### Install

1. Download the latest release from [Releases](https://github.com/vieiraae/clip-sidekick/releases)
2. Extract the zip to a folder of your choice
3. Run `ClipSidekick.exe`

### First steps

1. Copy some text — a notification bubble appears near your cursor
2. Press **Win+Shift+V** to open the history window
3. Click an item to paste it, or click **✨** to edit it with AI
4. Right-click the tray icon to access settings

## Settings

Right-click the tray icon → **Settings** to configure:

| Setting | Default | Description |
|---------|---------|-------------|
| Hotkey | Win+Shift+V | Global keyboard shortcut |
| Notification duration | 1 second | How long the toast popup stays visible |
| Max history items | 50 | Number of clipboard entries to keep (10–200) |

Settings are stored in `%APPDATA%\ClipSidekick\settings.json`.

## How it works

Clip Sidekick is a native Win32 application built with C# and .NET 10. It uses P/Invoke for window management, clipboard monitoring, and input simulation — no WPF or WinForms. The UI is rendered with GDI+ double-buffering for smooth, flicker-free drawing.

AI capabilities are powered by the **GitHub Copilot SDK**, which connects to the Copilot CLI running on your machine. When you trigger an AI task, the app sends your clipboard content (or image) to Copilot and streams the result back. The Copilot CLI must be installed and authenticated separately.

### Architecture

```
src/ClipSidekick/
├── MainWindow.cs          Main window, rendering, and event handling
├── NotificationPopup.cs   Toast notification UI
├── ClipboardMonitor.cs    Clipboard capture and history management
├── ClipboardItem.cs       Data model for clipboard entries
├── CaretLocator.cs        Caret detection via UI Automation
├── AppSettings.cs         Settings persistence
├── SettingsDialog.cs      Settings UI
├── HotkeyDialog.cs        Hotkey capture dialog
├── NativeMethods.cs       Win32 P/Invoke declarations
└── Program.cs             Entry point
```

## Make it your own

The entire app is a single Visual Studio project with no complex dependencies beyond the Copilot SDK. You can use GitHub Copilot's coding abilities to customize it:

- **Add new AI tasks** — Edit the `EditTasks` array in `MainWindow.cs` and the matching `TaskLabels`/`TaskIcons` in `NotificationPopup.cs`
- **Change the theme** — Modify the color constants at the top of `MainWindow.cs` (`BgColor`, `AccentColor`, etc.)
- **Adjust the UI** — All rendering is in `OnPaint` methods using GDI+ — ask Copilot to add new buttons, change layouts, or add new tabs
- **Add new clipboard formats** — Extend `ClipboardMonitor.CaptureClipboard()` to handle rich text, files, or other formats
- **Change the AI prompt** — Modify `ComposePrompt()` in `MainWindow.cs` to change how instructions are sent to Copilot

To build from source:

```
cd src/ClipSidekick
dotnet run
```

## License

[MIT](LICENSE)

---

**Disclaimer:** This software is provided "as is", without warranty of any kind, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose, and noninfringement. In no event shall the authors be liable for any claim, damages, or other liability arising from the use of this software.
