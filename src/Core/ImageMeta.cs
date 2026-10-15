using System.Buffers.Binary;
using System.Text;

namespace MetaCleaner.Core;

/// <summary>
/// Чтение и удаление метаданных PNG/JPEG без перекодирования картинки:
/// PNG — текстовые блоки, время, EXIF, C2PA; JPEG — EXIF/XMP, IPTC, комментарии, C2PA.
/// Цветовой профиль (ICC) сохраняется.
/// </summary>
public static class ImageMeta
{
    static readonly byte[] PngSig = { 137, 80, 78, 71, 13, 10, 26, 10 };
    static readonly HashSet<string> PngDrop = new() { "tEXt", "zTXt", "iTXt", "tIME", "eXIf", "caBX" };

    public static bool IsImage(string name)
    {
        var e = Path.GetExtension(name).ToLowerInvariant();
        return e is ".png" or ".jpg" or ".jpeg";
    }

    static bool IsPng(byte[] d) => d.Length > 8 && d.AsSpan(0, 8).SequenceEqual(PngSig);
    static bool IsJpeg(byte[] d) => d.Length > 4 && d[0] == 0xFF && d[1] == 0xD8;

    /// <summary>Описание найденных метаданных; флаг ai — признак ИИ-происхождения.</summary>
    public static List<(string Text, bool Ai)> Describe(byte[] d)
    {
        var list = new List<(string, bool)>();
        if (IsPng(d))
        {
            WalkPng(d, (type, off, len) =>
            {
                switch (type)
                {
                    case "tEXt":
                        var raw = Encoding.Latin1.GetString(d, off, Math.Min(len, 200));
                        var txt = raw.Replace('\0', ':');
                        list.Add(($"PNG-текст «{Short(txt)}»", Signatures.Ai.IsMatch(raw)));
                        break;
                    case "zTXt":
                    case "iTXt":
                        var kw = Encoding.Latin1.GetString(d, off, Math.Min(len, 80)).Split('\0')[0];
                        list.Add(($"PNG-текст «{kw}»", Signatures.Ai.IsMatch(kw)));
                        break;
                    case "eXIf": list.Add(("EXIF", false)); break;
                    case "tIME": list.Add(("время создания картинки", false)); break;
                    case "caBX": list.Add(("подпись C2PA (происхождение ИИ-изображения)", true)); break;
                }
            });
        }
        else if (IsJpeg(d))
        {
            WalkJpeg(d, (m, off, len) =>
            {
                var seg = d.AsSpan(off, len);
                string ascii = Encoding.Latin1.GetString(seg[..Math.Min(len, 4096)]);
                bool ai = Signatures.Ai.IsMatch(ascii);
                if (m == 0xE1 && ascii.StartsWith("Exif")) list.Add(("EXIF", ai));
                else if (m == 0xE1) list.Add(("XMP", ai));
                else if (m == 0xED) list.Add(("IPTC/Photoshop", ai));
                else if (m == 0xFE) list.Add(($"комментарий «{Short(ascii)}»", ai));
                else if (m == 0xEB) list.Add(("подпись C2PA (происхождение ИИ-изображения)", true));
            });
        }
        return list;
    }

    /// <summary>Возвращает картинку без метаданных (или исходные байты, если чистить нечего).</summary>
    public static byte[] Strip(byte[] d, out bool changed)
    {
        changed = false;
        try
        {
            if (IsPng(d)) return StripPng(d, out changed);
            if (IsJpeg(d)) return StripJpeg(d, out changed);
        }
        catch { changed = false; }
        return d;
    }

    static string Short(string s)
    {
        s = new string(s.Where(c => !char.IsControl(c)).ToArray()).Trim();
        return s.Length > 60 ? s[..60] + "…" : s;
    }

    // ---------- PNG ----------

    static bool WalkPng(byte[] d, Action<string, int, int> onChunk)
    {
        int pos = 8;
        while (pos + 12 <= d.Length)
        {
            int len = BinaryPrimitives.ReadInt32BigEndian(d.AsSpan(pos));
            if (len < 0 || pos + 12 + len > d.Length) return false;
            string type = Encoding.ASCII.GetString(d, pos + 4, 4);
            onChunk(type, pos + 8, len);
            pos += 12 + len;
            if (type == "IEND") return true;
        }
        return false;
    }

    static byte[] StripPng(byte[] d, out bool changed)
    {
        changed = false;
        using var o = new MemoryStream(d.Length);
        o.Write(PngSig);
        bool drop = false;
        bool ok = WalkPng(d, (type, off, len) =>
        {
            if (PngDrop.Contains(type)) { drop = true; return; }
            o.Write(d, off - 8, len + 12);
        });
        if (!ok || !drop) return d;
        changed = true;
        return o.ToArray();
    }

    // ---------- JPEG ----------

    /// <summary>Обходит сегменты JPEG до начала данных изображения (SOS).</summary>
    static int WalkJpeg(byte[] d, Action<byte, int, int>? onSeg, Action<int, int>? copy = null)
    {
        int pos = 2;
        copy?.Invoke(0, 2);
        while (pos + 4 <= d.Length)
        {
            if (d[pos] != 0xFF) return -1;
            byte m = d[pos + 1];
            if (m == 0xFF) { pos++; continue; }
            if (m == 0xDA || m == 0xD9) { copy?.Invoke(pos, d.Length - pos); return d.Length; }
            if (m is >= 0xD0 and <= 0xD7 or 0x01) { copy?.Invoke(pos, 2); pos += 2; continue; }
            int len = (d[pos + 2] << 8) | d[pos + 3];
            if (len < 2 || pos + 2 + len > d.Length) return -1;
            onSeg?.Invoke(m, pos + 4, len - 2);
            pos += 2 + len;
        }
        return -1;
    }

    static byte[] StripJpeg(byte[] d, out bool changed)
    {
        changed = false;
        using var o = new MemoryStream(d.Length);
        bool drop = false;
        int pos = 2;
        o.Write(d, 0, 2);
        while (pos + 4 <= d.Length)
        {
            if (d[pos] != 0xFF) return d;
            byte m = d[pos + 1];
            if (m == 0xFF) { pos++; continue; }
            if (m == 0xDA || m == 0xD9) { o.Write(d, pos, d.Length - pos); pos = d.Length; break; }
            if (m is >= 0xD0 and <= 0xD7 or 0x01) { o.Write(d, pos, 2); pos += 2; continue; }
            int len = (d[pos + 2] << 8) | d[pos + 3];
            if (len < 2 || pos + 2 + len > d.Length) return d;
            bool remove = m is 0xE1 or 0xED or 0xFE or 0xEB;
            if (remove) drop = true;
            else o.Write(d, pos, 2 + len);
            pos += 2 + len;
        }
        if (pos != d.Length || !drop) return d;
        changed = true;
        return o.ToArray();
    }
}
