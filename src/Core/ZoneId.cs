using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace MetaCleaner.Core;

/// <summary>
/// Отметка Windows «файл получен из интернета» — скрытый поток NTFS Zone.Identifier.
/// В нём может храниться адрес, откуда скачан файл (HostUrl, ReferrerUrl).
/// </summary>
public static class ZoneId
{
    const uint GENERIC_READ = 0x80000000, FILE_SHARE_READ = 1, OPEN_EXISTING = 3;

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern SafeFileHandle CreateFileW(string name, uint access, uint share, IntPtr sec, uint disp, uint flags, IntPtr tmpl);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    static extern bool DeleteFileW(string name);

    static string Stream(string path) => Path.GetFullPath(path) + ":Zone.Identifier";

    /// <summary>Содержимое отметки или null, если её нет (или не Windows).</summary>
    public static string? Read(string path)
    {
        if (!OperatingSystem.IsWindows()) return null;
        try
        {
            using var h = CreateFileW(Stream(path), GENERIC_READ, FILE_SHARE_READ, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
            if (h.IsInvalid) return null;
            using var fs = new FileStream(h, FileAccess.Read);
            using var r = new StreamReader(fs, Encoding.UTF8);
            return r.ReadToEnd();
        }
        catch { return null; }
    }

    public static bool Remove(string path)
    {
        if (!OperatingSystem.IsWindows()) return false;
        try { return DeleteFileW(Stream(path)); }
        catch { return false; }
    }

    /// <summary>Адреса из отметки (HostUrl, ReferrerUrl).</summary>
    public static List<string> Urls(string content) =>
        content.Split('\n')
            .Select(l => l.Trim())
            .Where(l => l.StartsWith("HostUrl=", StringComparison.OrdinalIgnoreCase) ||
                        l.StartsWith("ReferrerUrl=", StringComparison.OrdinalIgnoreCase))
            .Select(l => l[(l.IndexOf('=') + 1)..])
            .ToList();
}
