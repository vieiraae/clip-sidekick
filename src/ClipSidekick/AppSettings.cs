using System.Text.Json;

namespace ClipSidekick;

internal sealed class AppSettings
{
    internal static readonly string SettingsDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ClipSidekick");
    private static readonly string SettingsFile = Path.Combine(SettingsDir, "settings.json");
    internal static readonly string McpJsonFile = Path.Combine(SettingsDir, "mcp.json");

    public string Hotkey { get; set; } = "Win+Shift+V";
    public int NotificationDurationMs { get; set; } = 1000;
    public int MaxHistoryItems { get; set; } = 50;
    public bool ShowNotification { get; set; } = true;
    public int EditTaskIndex { get; set; }
    public int EditToneIndex { get; set; }
    public int EditFormatIndex { get; set; }
    public int EditLengthIndex { get; set; }
    public int EditChoices { get; set; } = 1;
    public string Model { get; set; } = "";
    public string[] QuickTaskHotkeys { get; set; } = ["", "", "", "", "", "", "", "", "", ""];
    public List<CustomTask> CustomTasks { get; set; } = [];
    public string FreeformHotkey { get; set; } = "";
    public string BubbleHotkey { get; set; } = "";
    public bool McpEnabled { get; set; }
    public List<string> DisabledMcpServers { get; set; } = [];
    public bool SkillsEnabled { get; set; }

    internal static readonly string SkillsDir = Path.Combine(SettingsDir, "skills");
    public string SystemMessage { get; set; } = "You are a helpful assistant called Sidekick that provides succinct and accurate responses. Respond ONLY with the generated text. No explanations, no markdown fences, no preamble.";

    public string HotkeyDisplay => Hotkey;

    public (uint mods, uint vk) ParseHotkey()
    {
        return ParseHotkeyString(Hotkey);
    }

    public static (uint mods, uint vk) ParseHotkeyString(string hotkey)
    {
        uint mods = NativeMethods.MOD_NOREPEAT;
        uint vk = 0;

        var parts = hotkey.Split('+');
        foreach (var part in parts)
        {
            switch (part)
            {
                case "Win": mods |= NativeMethods.MOD_WIN; break;
                case "Ctrl": mods |= NativeMethods.MOD_CONTROL; break;
                case "Alt": mods |= NativeMethods.MOD_ALT; break;
                case "Shift": mods |= NativeMethods.MOD_SHIFT; break;
                default: vk = ParseVk(part); break;
            }
        }

        return (mods, vk);
    }

    private static uint ParseVk(string key)
    {
        if (key.Length == 1)
        {
            char c = char.ToUpperInvariant(key[0]);
            if (c is >= 'A' and <= 'Z') return c;
            if (c is >= '0' and <= '9') return c;
        }

        if (key.StartsWith("F") && int.TryParse(key.AsSpan(1), out int fn) && fn is >= 1 and <= 12)
            return (uint)(0x70 + fn - 1);

        return key switch
        {
            "Space" => 0x20,
            "Tab" => 0x09,
            "Enter" => 0x0D,
            "Insert" => 0x2D,
            "Delete" => 0x2E,
            "Home" => 0x24,
            "End" => 0x23,
            "PageUp" => 0x21,
            "PageDown" => 0x22,
            "Left" => 0x25,
            "Up" => 0x26,
            "Right" => 0x27,
            "Down" => 0x28,
            "`" => 0xC0,
            "-" => 0xBD,
            "=" => 0xBB,
            "[" => 0xDB,
            "]" => 0xDD,
            "\\" => 0xDC,
            ";" => 0xBA,
            "'" => 0xDE,
            "," => 0xBC,
            "." => 0xBE,
            "/" => 0xBF,
            _ => 0x56 // fallback to 'V'
        };
    }

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(SettingsFile))
            {
                var json = File.ReadAllText(SettingsFile);
                return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
            }
        }
        catch { }
        return new AppSettings();
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(SettingsDir);
            var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(SettingsFile, json);
        }
        catch { }
    }
}

internal sealed class CustomTask
{
    public string Name { get; set; } = "";
    public string Prompt { get; set; } = "";
    public string Hotkey { get; set; } = "";
}
