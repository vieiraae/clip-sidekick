using System.Drawing;

namespace ClipSidekick;

internal enum ClipboardItemType
{
    Text,
    Image
}

internal sealed class ClipboardItem
{
    public ClipboardItemType Type { get; set; } = ClipboardItemType.Text;
    public string Text { get; set; } = string.Empty;
    public Bitmap? Image { get; set; }
    public DateTime Timestamp { get; set; } = DateTime.Now;
    public bool IsBookmarked { get; set; }

    public string PreviewText
    {
        get
        {
            if (Type == ClipboardItemType.Image)
                return Image != null ? $"Image ({Image.Width}×{Image.Height})" : "Image";
            var t = Text.Length > 200 ? Text[..200] + "…" : Text;
            return t.ReplaceLineEndings(" ");
        }
    }
}
