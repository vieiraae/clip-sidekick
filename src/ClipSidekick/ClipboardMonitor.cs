using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace ClipSidekick;

internal sealed class ClipboardMonitor
{
    private readonly List<ClipboardItem> _items = new();
    private readonly List<ClipboardItem> _bookmarkedItems = new();
    private int _maxItems = 50;
    private bool _ignoreNext;
    private DateTime _lastCaptureTime;
    private const int DebounceMs = 100;

    private static readonly string DataDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ClipSidekick");
    private static readonly string BookmarksPath = Path.Combine(DataDir, "bookmarks.json");
    private static readonly string BookmarkImagesDir = Path.Combine(DataDir, "bookmark_images");

    public IReadOnlyList<ClipboardItem> Items => _items;
    public IReadOnlyList<ClipboardItem> BookmarkedItems => _bookmarkedItems;

    public event Action? ClipboardChanged;

    public int MaxItems
    {
        get => _maxItems;
        set
        {
            _maxItems = Math.Max(10, value);
            while (_items.Count > _maxItems)
                _items.RemoveAt(_items.Count - 1);
        }
    }

    public ClipboardMonitor()
    {
        LoadBookmarkedItems();
    }

    public void SetIgnoreNext() => _ignoreNext = true;

    public void OnClipboardUpdate()
    {
        if (_ignoreNext)
        {
            _ignoreNext = false;
            return;
        }

        // Debounce: many apps fire multiple WM_CLIPBOARDUPDATE for a single copy
        var now = DateTime.UtcNow;
        if ((now - _lastCaptureTime).TotalMilliseconds < DebounceMs)
            return;

        var item = CaptureClipboard();
        if (item == null)
            return;

        _lastCaptureTime = now;

        // Don't add duplicates at the top, but still notify so the bubble shows
        if (_items.Count > 0)
        {
            var top = _items[0];
            if (item.Type == ClipboardItemType.Text && top.Type == ClipboardItemType.Text && top.Text == item.Text)
            {
                ClipboardChanged?.Invoke();
                return;
            }
            if (item.Type == ClipboardItemType.Image && top.Type == ClipboardItemType.Image)
            {
                ClipboardChanged?.Invoke();
                return;
            }
        }

        // Remove any existing non-bookmarked text duplicate
        if (item.Type == ClipboardItemType.Text)
        {
            for (int i = _items.Count - 1; i >= 0; i--)
            {
                if (_items[i].Type == ClipboardItemType.Text && _items[i].Text == item.Text && !_items[i].IsBookmarked)
                {
                    _items.RemoveAt(i);
                    break;
                }
            }
        }

        _items.Insert(0, item);

        // Trim items from the end
        while (_items.Count > _maxItems)
            _items.RemoveAt(_items.Count - 1);

        ClipboardChanged?.Invoke();
    }

    public void BookmarkItem(int index)
    {
        if (index >= 0 && index < _items.Count)
        {
            var item = _items[index];
            _items.RemoveAt(index);
            item.IsBookmarked = true;
            _bookmarkedItems.Insert(0, item);
            SaveBookmarkedItems();
            ClipboardChanged?.Invoke();
        }
    }

    public void UnbookmarkItem(int index)
    {
        if (index >= 0 && index < _bookmarkedItems.Count)
        {
            var item = _bookmarkedItems[index];
            _bookmarkedItems.RemoveAt(index);
            item.IsBookmarked = false;
            _items.Insert(0, item);
            SaveBookmarkedItems();
            ClipboardChanged?.Invoke();
        }
    }

    public void DeleteItem(int index)
    {
        if (index >= 0 && index < _items.Count)
        {
            _items.RemoveAt(index);
            ClipboardChanged?.Invoke();
        }
    }

    public void ClearHistory()
    {
        _items.Clear();
        ClipboardChanged?.Invoke();
    }

    public void DeleteBookmarkedItem(int index)
    {
        if (index >= 0 && index < _bookmarkedItems.Count)
        {
            _bookmarkedItems.RemoveAt(index);
            SaveBookmarkedItems();
            ClipboardChanged?.Invoke();
        }
    }

    public void ClearUnbookmarked()
    {
        _items.Clear();
        ClipboardChanged?.Invoke();
    }

    private static ClipboardItem? CaptureClipboard()
    {
        if (!NativeMethods.OpenClipboard(IntPtr.Zero))
            return null;

        try
        {
            // Try text first
            if (NativeMethods.IsClipboardFormatAvailable(NativeMethods.CF_UNICODETEXT))
            {
                var text = GetClipboardText();
                if (!string.IsNullOrEmpty(text))
                    return new ClipboardItem { Type = ClipboardItemType.Text, Text = text };
            }

            // Try bitmap
            if (NativeMethods.IsClipboardFormatAvailable(NativeMethods.CF_BITMAP))
            {
                var image = GetClipboardBitmap();
                if (image != null)
                    return new ClipboardItem { Type = ClipboardItemType.Image, Image = image };
            }

            return null;
        }
        finally
        {
            NativeMethods.CloseClipboard();
        }
    }

    private static string? GetClipboardText()
    {
        var hData = NativeMethods.GetClipboardData(NativeMethods.CF_UNICODETEXT);
        if (hData == IntPtr.Zero)
            return null;

        var pData = NativeMethods.GlobalLock(hData);
        if (pData == IntPtr.Zero)
            return null;

        try
        {
            return Marshal.PtrToStringUni(pData);
        }
        finally
        {
            NativeMethods.GlobalUnlock(hData);
        }
    }

    private static Bitmap? GetClipboardBitmap()
    {
        var hBitmap = NativeMethods.GetClipboardData(NativeMethods.CF_BITMAP);
        if (hBitmap == IntPtr.Zero)
            return null;

        try
        {
            using var src = Image.FromHbitmap(hBitmap);
            // Make a copy so we don't hold the clipboard handle
            var copy = new Bitmap(src.Width, src.Height, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(copy))
                g.DrawImage(src, 0, 0, src.Width, src.Height);
            return copy;
        }
        catch
        {
            return null;
        }
    }

    // --- Bookmarked items persistence ---

    private record BookmarkItemData(string Type, string? Text, string? ImageFile, DateTime Timestamp);

    private void SaveBookmarkedItems()
    {
        try
        {
            Directory.CreateDirectory(DataDir);
            Directory.CreateDirectory(BookmarkImagesDir);

            var entries = new List<BookmarkItemData>();
            for (int i = 0; i < _bookmarkedItems.Count; i++)
            {
                var item = _bookmarkedItems[i];
                string? imageFile = null;
                if (item.Type == ClipboardItemType.Image && item.Image != null)
                {
                    imageFile = $"bookmark_{i}.png";
                    item.Image.Save(Path.Combine(BookmarkImagesDir, imageFile), ImageFormat.Png);
                }
                entries.Add(new BookmarkItemData(
                    item.Type.ToString(),
                    item.Type == ClipboardItemType.Text ? item.Text : null,
                    imageFile,
                    item.Timestamp));
            }

            var json = JsonSerializer.Serialize(entries);
            File.WriteAllText(BookmarksPath, json);
        }
        catch
        {
            // Silently fail
        }
    }

    private void LoadBookmarkedItems()
    {
        try
        {
            // Migrate from old pinned.json if bookmarks.json doesn't exist
            var oldPinnedPath = Path.Combine(DataDir, "pinned.json");
            if (!File.Exists(BookmarksPath) && File.Exists(oldPinnedPath))
            {
                try { File.Move(oldPinnedPath, BookmarksPath); } catch { }
            }
            if (!File.Exists(BookmarksPath)) return;

            var json = File.ReadAllText(BookmarksPath);
            var entries = JsonSerializer.Deserialize<List<BookmarkItemData>>(json);
            if (entries == null) return;

            foreach (var entry in entries)
            {
                var item = new ClipboardItem
                {
                    IsBookmarked = true,
                    Timestamp = entry.Timestamp
                };

                if (entry.Type == nameof(ClipboardItemType.Image) && entry.ImageFile != null)
                {
                    // Try new path first, fall back to old pinned_images path
                    var imgPath = Path.Combine(BookmarkImagesDir, entry.ImageFile);
                    if (!File.Exists(imgPath))
                        imgPath = Path.Combine(DataDir, "pinned_images", entry.ImageFile);
                    if (File.Exists(imgPath))
                    {
                        using var stream = File.OpenRead(imgPath);
                        item.Type = ClipboardItemType.Image;
                        item.Image = new Bitmap(stream);
                    }
                    else continue;
                }
                else
                {
                    item.Type = ClipboardItemType.Text;
                    item.Text = entry.Text ?? string.Empty;
                }

                _bookmarkedItems.Add(item);
            }
        }
        catch
        {
            // Corrupted file — skip
        }
    }
}
