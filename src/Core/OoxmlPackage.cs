using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace MetaCleaner.Core;

public sealed class PackEntry
{
    public string Name { get; set; } = "";
    public byte[] Data { get; set; } = Array.Empty<byte>();
    public DateTimeOffset Time { get; set; }
}

/// <summary>Документ Office (OOXML) как список частей ZIP-архива в памяти.</summary>
public sealed class OoxmlPackage
{
    public const string NsRel = "http://schemas.openxmlformats.org/package/2006/relationships";
    public const string NsCt = "http://schemas.openxmlformats.org/package/2006/content-types";

    /// <summary>Office пишет в архив нулевую DOS-дату: 1980-01-01 00:00.</summary>
    public static readonly DateTimeOffset OfficeZipTime =
        new(new DateTime(1980, 1, 1, 0, 0, 0, DateTimeKind.Unspecified), TimeSpan.Zero);

    public List<PackEntry> Entries { get; } = new();

    public static OoxmlPackage Load(string path) => Load(File.ReadAllBytes(path));

    public static OoxmlPackage Load(byte[] bytes)
    {
        var p = new OoxmlPackage();
        using var ms = new MemoryStream(bytes, writable: false);
        using var za = new ZipArchive(ms, ZipArchiveMode.Read);
        foreach (var e in za.Entries)
        {
            if (e.FullName.EndsWith('/')) continue;
            using var s = e.Open();
            using var o = new MemoryStream();
            s.CopyTo(o);
            p.Entries.Add(new PackEntry { Name = e.FullName, Data = o.ToArray(), Time = e.LastWriteTime });
        }
        if (p.Find("[Content_Types].xml") == null)
            throw new InvalidDataException("в архиве нет [Content_Types].xml — это не документ Office");
        return p;
    }

    public PackEntry? Find(string name)
    {
        name = name.TrimStart('/');
        return Entries.FirstOrDefault(e => string.Equals(e.Name, name, StringComparison.OrdinalIgnoreCase));
    }

    public bool Has(string name) => Find(name) != null;
    public byte[]? Get(string name) => Find(name)?.Data;

    public void Set(string name, byte[] data)
    {
        var e = Find(name);
        if (e == null) Entries.Add(new PackEntry { Name = name.TrimStart('/'), Data = data, Time = OfficeZipTime });
        else e.Data = data;
    }

    public void Remove(string name)
    {
        name = name.TrimStart('/');
        Entries.RemoveAll(e => string.Equals(e.Name, name, StringComparison.OrdinalIgnoreCase));
    }

    public IEnumerable<PackEntry> Matching(Func<string, bool> pred) => Entries.Where(e => pred(e.Name));

    public XDocument? Xml(string name)
    {
        var d = Get(name);
        if (d == null) return null;
        using var ms = new MemoryStream(d, writable: false);
        return XDocument.Load(ms, LoadOptions.PreserveWhitespace);
    }

    public void SetXml(string name, XDocument doc) => Set(name, XmlBytes(doc));

    /// <summary>Сериализует XML так же, как Office: заголовок с UTF-8 в верхнем регистре и перевод строки CRLF.</summary>
    public static byte[] XmlBytes(XDocument doc)
    {
        var settings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false), Indent = false, OmitXmlDeclaration = true,
        };
        using var ms = new MemoryStream();
        var head = Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
        ms.Write(head);
        using (var w = XmlWriter.Create(ms, settings)) doc.Root!.WriteTo(w);
        // .NET пишет пустые элементы как «<a />», Office — как «<a/>».
        // Последовательность « />» встречается только в концах тегов: «>» в тексте и атрибутах экранируется.
        return RemoveSpaceBeforeSlash(ms.ToArray());
    }

    static byte[] RemoveSpaceBeforeSlash(byte[] b)
    {
        var o = new MemoryStream(b.Length);
        for (int i = 0; i < b.Length; i++)
        {
            if (b[i] == (byte)' ' && i + 2 < b.Length && b[i + 1] == (byte)'/' && b[i + 2] == (byte)'>') continue;
            o.WriteByte(b[i]);
        }
        return o.ToArray();
    }

    public static bool ContainsAscii(byte[] data, string text) =>
        data.AsSpan().IndexOf(Encoding.ASCII.GetBytes(text)) >= 0;

    // ---------- связи и типы содержимого ----------

    public string? PartByRelType(string relType)
    {
        var x = Xml("_rels/.rels");
        var r = x?.Root?.Elements(XName.Get("Relationship", NsRel))
            .FirstOrDefault(e => (string?)e.Attribute("Type") == relType);
        return ((string?)r?.Attribute("Target"))?.TrimStart('/');
    }

    public void AddRel(string relType, string target)
    {
        var x = Xml("_rels/.rels") ?? new XDocument(new XElement(XName.Get("Relationships", NsRel)));
        var root = x.Root!;
        var ids = root.Elements(XName.Get("Relationship", NsRel)).Select(e => (string?)e.Attribute("Id")).ToHashSet();
        int n = 100;
        while (ids.Contains("rId" + n)) n++;
        root.Add(new XElement(XName.Get("Relationship", NsRel),
            new XAttribute("Id", "rId" + n), new XAttribute("Type", relType), new XAttribute("Target", target)));
        SetXml("_rels/.rels", x);
    }

    public void RemoveRelByType(string relType)
    {
        var x = Xml("_rels/.rels");
        if (x?.Root == null) return;
        var rels = x.Root.Elements(XName.Get("Relationship", NsRel))
            .Where(e => (string?)e.Attribute("Type") == relType).ToList();
        if (rels.Count == 0) return;
        rels.ForEach(r => r.Remove());
        SetXml("_rels/.rels", x);
    }

    public void SetOverride(string partName, string contentType)
    {
        var x = Xml("[Content_Types].xml");
        if (x?.Root == null) return;
        var ns = XNamespace.Get(NsCt);
        var o = x.Root.Elements(ns + "Override").FirstOrDefault(e =>
            string.Equals((string?)e.Attribute("PartName"), partName, StringComparison.OrdinalIgnoreCase));
        if (o != null && (string?)o.Attribute("ContentType") == contentType) return;
        if (o == null)
            x.Root.Add(new XElement(ns + "Override", new XAttribute("PartName", partName),
                new XAttribute("ContentType", contentType)));
        else o.SetAttributeValue("ContentType", contentType);
        SetXml("[Content_Types].xml", x);
    }

    public void RemoveOverride(string partName)
    {
        var x = Xml("[Content_Types].xml");
        if (x?.Root == null) return;
        var ns = XNamespace.Get(NsCt);
        var list = x.Root.Elements(ns + "Override").Where(e =>
            string.Equals((string?)e.Attribute("PartName"), partName, StringComparison.OrdinalIgnoreCase)).ToList();
        if (list.Count == 0) return;
        list.ForEach(e => e.Remove());
        SetXml("[Content_Types].xml", x);
    }

    // ---------- сохранение ----------

    /// <summary>
    /// Собирает архив заново: [Content_Types].xml первым, затем _rels/.rels, затем остальные части.
    /// При officeLikeZip все даты внутри архива = 1980-01-01, как у Microsoft Office.
    /// </summary>
    public byte[] ToBytes(bool officeLikeZip)
    {
        var ct = Find("[Content_Types].xml");
        var rr = Find("_rels/.rels");
        var ordered = new List<PackEntry>();
        if (ct != null) ordered.Add(ct);
        if (rr != null) ordered.Add(rr);
        ordered.AddRange(Entries.Where(e => e != ct && e != rr));

        using var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var e in ordered)
            {
                var ze = za.CreateEntry(e.Name, CompressionLevel.Optimal);
                ze.LastWriteTime = officeLikeZip || e.Time.Year < 1980 || e.Time.Year > 2107 ? OfficeZipTime : e.Time;
                using var s = ze.Open();
                s.Write(e.Data, 0, e.Data.Length);
            }
        }
        return ms.ToArray();
    }

    /// <summary>Проверяет, что архив открывается и все XML-части корректны.</summary>
    public static void Validate(byte[] bytes)
    {
        var p = Load(bytes);
        foreach (var e in p.Entries)
        {
            if (!e.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                !e.Name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;
            try { p.Xml(e.Name); }
            catch (Exception ex) { throw new InvalidDataException($"повреждён XML в части {e.Name}: {ex.Message}"); }
        }
        if (!p.Has("_rels/.rels")) throw new InvalidDataException("нет _rels/.rels");
    }
}
