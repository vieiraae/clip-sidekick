using System.Runtime.InteropServices;

namespace ClipSidekick;

/// <summary>
/// Detects the text caret (blinking cursor) position using UI Automation.
/// Works across modern apps (VS Code, Chrome, WinUI) that don't use the classic Win32 caret.
/// </summary>
internal static class CaretLocator
{
    private static readonly Guid CLSID_CUIAutomation8 = new("E22AD333-B25F-460C-83D0-0581107395C9");
    private static readonly Guid IID_IUIAutomationTextPattern2 = new("506A921A-FCC9-409F-B23B-37EB74106872");
    private const int UIA_TextPattern2Id = 10024;

    public static bool TryGetCaretPosition(out int x, out int y)
    {
        x = y = 0;
        object? automation = null;
        IUIAutomationElement? element = null;
        object? patternObj = null;
        IUIAutomationTextRange? range = null;

        try
        {
            var type = Type.GetTypeFromCLSID(CLSID_CUIAutomation8);
            if (type == null) return false;

            automation = Activator.CreateInstance(type);
            if (automation == null) return false;

            var uia = (IUIAutomation)automation;
            if (uia.GetFocusedElement(out element) != 0 || element == null)
                return false;

            var iid = IID_IUIAutomationTextPattern2;
            if (element.GetCurrentPatternAs(UIA_TextPattern2Id, ref iid, out patternObj) != 0 || patternObj == null)
                return false;

            var textPattern2 = (IUIAutomationTextPattern2)patternObj;
            if (textPattern2.GetCaretRange(out _, out range) != 0 || range == null)
                return false;

            if (range.GetBoundingRectangles(out var rects) != 0 || rects == null || rects.Length < 4)
                return false;

            // rects is groups of 4 doubles: x, y, width, height
            x = (int)rects[0];
            y = (int)(rects[1] + rects[3]); // bottom of caret

            return x > 0 || y > 0;
        }
        catch
        {
            return false;
        }
        finally
        {
            if (range != null) try { Marshal.ReleaseComObject(range); } catch { }
            if (patternObj != null) try { Marshal.ReleaseComObject(patternObj); } catch { }
            if (element != null) try { Marshal.ReleaseComObject(element); } catch { }
            if (automation != null) try { Marshal.ReleaseComObject(automation); } catch { }
        }
    }

    // --- Minimal COM interface stubs for UI Automation ---
    // Only the methods we actually call have real signatures.
    // Reserved slots map to vtable entries we never invoke.

    [ComImport, Guid("30cbe57d-d9d0-452a-ab13-7ac5ac4825ee")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IUIAutomation
    {
        void _reserved0(); // CompareElements
        void _reserved1(); // CompareRuntimeIds
        void _reserved2(); // GetRootElement
        void _reserved3(); // ElementFromHandle
        void _reserved4(); // ElementFromPoint

        [PreserveSig]
        int GetFocusedElement([MarshalAs(UnmanagedType.Interface)] out IUIAutomationElement element);
    }

    [ComImport, Guid("d22108aa-8ac5-49a5-837b-37bbb3d7591e")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IUIAutomationElement
    {
        void _r0();  // SetFocus
        void _r1();  // GetRuntimeId
        void _r2();  // FindFirst
        void _r3();  // FindAll
        void _r4();  // FindFirstBuildCache
        void _r5();  // FindAllBuildCache
        void _r6();  // BuildUpdatedCache
        void _r7();  // GetCurrentPropertyValue
        void _r8();  // GetCurrentPropertyValueEx
        void _r9();  // GetCachedPropertyValue
        void _r10(); // GetCachedPropertyValueEx

        [PreserveSig]
        int GetCurrentPatternAs(int patternId, ref Guid riid,
            [MarshalAs(UnmanagedType.IUnknown)] out object patternObject);
    }

    [ComImport, Guid("506a921a-fcc9-409f-b23b-37eb74106872")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IUIAutomationTextPattern2
    {
        // IUIAutomationTextPattern methods
        void _r0(); // RangeFromPoint
        void _r1(); // RangeFromChild
        void _r2(); // GetSelection
        void _r3(); // GetVisibleRanges
        void _r4(); // get_DocumentRange
        void _r5(); // get_SupportedTextSelection

        // IUIAutomationTextPattern2 methods
        void _r6(); // RangeFromAnnotation

        [PreserveSig]
        int GetCaretRange(out int isActive,
            [MarshalAs(UnmanagedType.Interface)] out IUIAutomationTextRange range);
    }

    [ComImport, Guid("a543cc6a-f4ae-494b-8239-c814481187a8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IUIAutomationTextRange
    {
        void _r0(); // Clone
        void _r1(); // Compare
        void _r2(); // CompareEndpoints
        void _r3(); // ExpandToEnclosingUnit
        void _r4(); // FindAttribute
        void _r5(); // FindText
        void _r6(); // GetAttributeValue

        [PreserveSig]
        int GetBoundingRectangles(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_R8)] out double[] boundingRects);
    }
}
