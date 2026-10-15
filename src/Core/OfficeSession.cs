using System.Runtime.InteropServices;

namespace MetaCleaner.Core;

/// <summary>
/// Пересохранение файлов через установленный Microsoft Office (COM-автоматизация).
/// Office сам перестраивает структуру файла: RSID, settings.xml, тема, миниатюра.
/// Использовать только из STA-потока.
/// </summary>
public sealed class OfficeSession : IDisposable
{
    dynamic? word, ppt, excel;

    static string ProgId(DocKind k) => k switch
    {
        DocKind.Word => "Word.Application",
        DocKind.PowerPoint => "PowerPoint.Application",
        _ => "Excel.Application",
    };

    public static bool IsAvailable(DocKind k) =>
        OperatingSystem.IsWindows() && Type.GetTypeFromProgID(ProgId(k)) != null;

    public static bool AnyAvailable() =>
        IsAvailable(DocKind.Word) || IsAvailable(DocKind.PowerPoint) || IsAvailable(DocKind.Excel);

    static dynamic Create(DocKind k)
    {
        var t = Type.GetTypeFromProgID(ProgId(k))
                ?? throw new InvalidOperationException($"{Fields.KindName(k)} не установлен");
        return Activator.CreateInstance(t)!;
    }

    public void Resave(string path, DocKind k)
    {
        if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException("пересохранение через Office работает только в Windows");
        var full = Path.GetFullPath(path);
        switch (k)
        {
            case DocKind.Word:
            {
                word ??= Create(k);
                word.Visible = false;
                word.DisplayAlerts = 0; // wdAlertsNone
                // Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
                dynamic doc = word.Documents.Open(full, false, false, false);
                try { doc.Saved = false; doc.Save(); }
                finally { doc.Close(0); Release(doc); }
                break;
            }
            case DocKind.PowerPoint:
            {
                ppt ??= Create(k);
                ppt.DisplayAlerts = 1; // ppAlertsNone
                // Open(FileName, ReadOnly, Untitled, WithWindow) — msoFalse = 0
                dynamic pres = ppt.Presentations.Open(full, 0, 0, 0);
                try { pres.Saved = 0; pres.Save(); }
                finally { pres.Close(); Release(pres); }
                break;
            }
            default:
            {
                excel ??= Create(k);
                excel.Visible = false;
                excel.DisplayAlerts = false;
                dynamic wb = excel.Workbooks.Open(full);
                try { wb.Saved = false; wb.Save(); }
                finally { wb.Close(false); Release(wb); }
                break;
            }
        }
    }

    static void Release(object? o)
    {
        try
        {
            if (o != null && OperatingSystem.IsWindows() && Marshal.IsComObject(o)) Marshal.FinalReleaseComObject(o);
        }
        catch { /* ignore */ }
    }

    public void Dispose()
    {
        try { word?.Quit(0); } catch { }
        try { ppt?.Quit(); } catch { }
        try { excel?.Quit(); } catch { }
        Release(word); Release(ppt); Release(excel);
        word = ppt = excel = null;
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
}
