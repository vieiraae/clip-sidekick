# Clip Sidekick

<p align="center">
  <img src="src/ClipSidekick/app.png" alt="Clip Sidekick" width="128">
</p>

**Your clipboard, supercharged with AI.** A lightweight Windows clipboard manager powered by [GitHub Copilot SDK](https://github.com/github/copilot-sdk) — copy, browse, bookmark, and let AI rewrite, proofread, or transform anything you've copied. All without leaving your current app.

## Highlights

**Copy → AI → Paste.** Copy text or a screenshot, pick an AI task from the notification bubble, and the result is pasted back instantly. Proofread an email, rewrite a paragraph, summarize a doc — one hotkey away.

**Ask AI anything.** Type a freeform instruction in the bubble — with or without context text. Generate content, answer questions, transform data — the result lands right where you're working.

**Hotkeys for everything.** Open the bubble, trigger clipboard history, or fire any AI task — all from global hotkeys you define. Select text in any app, press your bubble hotkey, and you're editing with AI.

**Your tasks, your rules.** Create custom AI tasks with your own name, instruction, and hotkey. They show up everywhere — the edit tab, the notification dropdown, and as global shortcuts.

**Extensible with MCP and Skills.** Connect to [MCP servers](https://modelcontextprotocol.io/) for tools like web search, APIs, and databases. Load agent skills from a local folder to teach the AI new behaviors — no code changes needed.

## Features

- **Clipboard history** — Last entries (text + images), searchable, with timestamps
- **Bookmarks** — Pin favorites; persisted across sessions
- **AI editing** — Proofread, Rewrite, Summarize, Expand, Explain, and 5 more built-in tasks
- **Custom tasks** — Unlimited custom AI tasks with name, instruction, and hotkey
- **Image-to-text** — Paste a screenshot; AI extracts readable text automatically
- **Ask AI** — Freeform AI instructions from the notification bubble
- **Notification bubble** — Non-intrusive toast with quick AI task dropdown on every copy
- **Bubble hotkey** — Open the bubble anywhere; copies selection or enters Ask AI mode
- **Tone / Format / Length** — Fine-tune AI output (Professional, Casual, Paragraphs, JSON…)
- **Multiple variations** — Generate up to 5 alternatives and pick the best
- **Model selection** — Choose from available Copilot models
- **System message** — Customize the AI system prompt
- **MCP servers** — Extend AI capabilities with Model Context Protocol servers
- **Skills** — Load custom skills from a local folder to teach AI new behaviors
- **Auto-paste** — AI results paste directly into your active app
- **Dark theme** — Modern UI with rounded corners, smooth animations, DPI-aware

## Getting started

### Prerequisites
Open a Windows terminal and run:
```
winget install Microsoft.DotNet.DesktopRuntime.10
winget install GitHub.Copilot
copilot /login
```

### Install & run

1. Download from [Releases](https://github.com/vieiraae/clip-sidekick/releases), extract, and run `ClipSidekick.exe`
2. Copy some text — a notification bubble appears near your cursor
3. Press **Win+Shift+V** to open clipboard history
4. Right-click the tray icon → **Settings** to configure hotkeys, tasks, and preferences

## Settings

### General

| Setting | Default | Description |
|---------|---------|-------------|
| Model | Auto | AI model to use |
| System message | Default prompt | Customize the AI system prompt |
| Max history items | 50 | Clipboard entries to keep (10–200) |
| Show notification | On | Toast popup on clipboard capture |
| Notification duration | 1s | How long the toast stays visible |
| Bubble hotkey | — | Show the notification bubble |
| Clipboard hotkey | Win+Shift+V | Open clipboard history |

### Quick Tasks

| Setting | Default | Description |
|---------|---------|-------------|
| Ask AI hotkey | — | Type a freeform AI instruction |
| Task hotkeys | — | Global hotkeys for built-in AI tasks |

### Custom Tasks

| Setting | Default | Description |
|---------|---------|-------------|
| Custom tasks | — | Your own AI tasks with name, instruction, and hotkey |

### Advanced

| Setting | Default | Description |
|---------|---------|-------------|
| Skills | Off | Load skills from `%APPDATA%\ClipSidekick\skills` |
| MCP servers | Off | Enable Model Context Protocol servers |

Settings are stored in `%APPDATA%\ClipSidekick\settings.json`.

## MCP servers

[Model Context Protocol](https://modelcontextprotocol.io/) servers extend what the AI can do — access APIs, query databases, interact with external services, and more.

1. Open **Settings → Advanced** and enable **MCP servers**
2. Click **Open mcp.json** to edit the configuration
3. Add servers in stdio or HTTP format:

```json
{
  "servers": {
    "weather": {
      "type": "stdio",
      "command": "npx",
      "args": ["-y", "@mcp-sidekick/weather"],
      "tools": ["*"]
    }
  }
}
```

Each server can be individually enabled or disabled from the Advanced tab.

## Skills

Skills let you teach the AI new behaviors by placing instruction files in a local folder.

1. Open **Settings → Advanced** and enable **Skills**
2. Click **Open skills folder** to open `%APPDATA%\ClipSidekick\skills`
3. Add skill files following the [Copilot SDK skills format](https://github.com/github/copilot-sdk)

Skills are loaded when the AI session starts and apply to all interactions.

## Build from source

```
cd src/ClipSidekick
dotnet run
```

Native Win32 app — C#, .NET 10, P/Invoke, GDI+ rendering. No WPF or WinForms. AI powered by the GitHub Copilot SDK via the Copilot CLI.

The entire app is a single-project codebase — easy to read, modify, and extend. Open it in VS Code with [GitHub Copilot](https://github.com/features/copilot) and you can add new AI tasks, tweak the UI, wire up new MCP servers, or build entirely new features with Copilot as your coding partner.

## License

[MIT](LICENSE)
