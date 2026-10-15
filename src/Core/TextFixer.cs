using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace MetaCleaner.Core;

public sealed class TextFixOptions
{
    /// <summary>Заменять длинное тире «—».</summary>
    public bool FixDashes { get; set; }
    /// <summary>На что заменять: «–» (короткое тире, как автозамена Word) или «-».</summary>
    public string DashReplacement { get; set; } = "–";
    /// <summary>Заменять длинное тире и в имени файла.</summary>
    public bool RenameFiles { get; set; }
    public bool RemoveMarkdown { get; set; }
    public bool RemoveChatPhrases { get; set; }
    public bool FillPlaceholders { get; set; }
    public bool CleanAltText { get; set; }

    public bool AnyText => FixDashes || RemoveMarkdown || RemoveChatPhrases || FillPlaceholders;
}

public sealed class TextFixReport
{
    public int Dashes, Markdown, Placeholders, AltTexts;
    public List<string> RemovedParagraphs { get; } = new();
    public List<string> Unfilled { get; } = new();
}

/// <summary>
/// Правка текста внутри документа без потери форматирования: изменения применяются
/// к исходным текстовым узлам (прогонам) абзаца, а не к абзацу целиком.
/// </summary>
public static class TextFixer
{
    static readonly XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static readonly XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
    static readonly XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    static readonly XName XmlSpace = XNamespace.Xml + "space";

    const RegexOptions O = RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled;
    static readonly Regex MdHeading = new(@"^\s*#{1,6}\s+", O);
    static readonly Regex MdBold = new(@"(\*\*|__)(?=\S)(.+?)(?<=\S)\1", O);
    static readonly Regex MdCode = new(@"(?<!`)`(?=\S)([^`\n]+?)`(?!`)", O);
    static readonly Regex MdLink = new(@"\[([^\]\n]+)\]\((https?://[^)\s]+)\)", O);
    static readonly Regex MdBullet = new(@"^\s*[*+]\s+", O);
    static readonly Regex MdOnly = new(@"^\s*(```[\w-]*|-{3,}|\*{3,}|_{3,})\s*$", O);

    enum EditKind { Dash, Markdown, Placeholder, Chat }
    readonly record struct Edit(int Start, int Length, string Replacement, EditKind Kind);

    /// <summary>Части документа, в которых лежит текст.</summary>
    public static bool IsTextPart(string name, DocKind kind)
    {
        var n = name.ToLowerInvariant();
        if (!n.EndsWith(".xml")) return false;
        return kind switch
        {
            DocKind.Word => n.StartsWith("word/") &&
                            (n.Contains("document") || n.Contains("header") || n.Contains("footer") ||
                             n.Contains("footnotes") || n.Contains("endnotes") || n.EndsWith("comments.xml")),
            DocKind.PowerPoint => n.StartsWith("ppt/slides/slide") || n.StartsWith("ppt/notesslides/"),
            _ => n == "xl/sharedstrings.xml" || n.StartsWith("xl/comments"),
        };
    }

    public static string FixFileName(string fileName, TextFixOptions o) =>
        o.FixDashes && o.RenameFiles && o.DashReplacement != "—" ? fileName.Replace("—", o.DashReplacement) : fileName;

    public static void Apply(OoxmlPackage p, DocKind kind, DocProps props, TextFixOptions o, TextFixReport rep)
    {
        if (o.CleanAltText) rep.AltTexts += CleanAlt(p);
        if (!o.AnyText) return;

        XName pName = kind switch { DocKind.Word => w + "p", DocKind.PowerPoint => a + "p", _ => x + "si" };
        XName tName = kind switch { DocKind.Word => w + "t", DocKind.PowerPoint => a + "t", _ => x + "t" };

        foreach (var e in p.Entries.ToList())
        {
            if (!IsTextPart(e.Name, kind)) continue;
            XDocument? doc;
            try { doc = p.Xml(e.Name); } catch { continue; }
            if (doc?.Root == null) continue;

            var paras = kind == DocKind.Excel
                ? doc.Root.Descendants().Where(el => el.Name == x + "si" || (el.Name == x + "text" && el.Parent?.Name == x + "comment")).ToList()
                : doc.Root.Descendants(pName).ToList();

            bool changed = false;
            foreach (var para in paras)
            {
                if (para.Document == null) continue; // уже удалён вместе с родителем
                var nodes = kind == DocKind.Excel
                    ? para.Descendants(tName).ToList()
                    : para.Descendants(tName).Where(t => t.Ancestors(pName).First() == para).ToList();
                if (nodes.Count == 0) continue;
                var texts = nodes.Select(n => n.Value).ToArray();
                var full = string.Concat(texts);
                var trimmed = full.Trim();
                if (trimmed.Length == 0) continue;

                // 1. абзац целиком — служебная фраза чата
                if (o.RemoveChatPhrases && trimmed.Length < 400 && Signatures.ChatPreamble.IsMatch(trimmed))
                {
                    rep.RemovedParagraphs.Add(trimmed);
                    RemoveParagraph(para, nodes, kind, pName);
                    changed = true;
                    continue;
                }
                // 2. абзац целиком — разделитель Markdown (```, ---, ***)
                if (o.RemoveMarkdown && MdOnly.IsMatch(full))
                {
                    rep.Markdown++;
                    RemoveParagraph(para, nodes, kind, pName);
                    changed = true;
                    continue;
                }

                // 3. точечные правки
                var edits = new List<Edit>();
                if (o.RemoveChatPhrases)
                {
                    var tail = Signatures.TrailingOffer.Match(full);
                    if (tail.Success)
                    {
                        rep.RemovedParagraphs.Add(tail.Groups[1].Value.Trim());
                        TryAdd(edits, new Edit(tail.Index, full.Length - tail.Index, "", EditKind.Chat));
                    }
                }
                if (o.RemoveMarkdown) CollectMarkdown(full, edits);
                if (o.FillPlaceholders) CollectPlaceholders(full, props, edits, rep);
                if (o.FixDashes && o.DashReplacement != "—")
                    for (int i = 0; i < full.Length; i++)
                        if (full[i] == '—') TryAdd(edits, new Edit(i, 1, o.DashReplacement, EditKind.Dash));
                if (edits.Count == 0) continue;

                var result = ApplyEdits(texts, edits);
                for (int i = 0; i < nodes.Count; i++)
                {
                    if (nodes[i].Value == result[i]) continue;
                    nodes[i].Value = result[i];
                    if (kind != DocKind.PowerPoint && result[i].Length > 0 &&
                        (char.IsWhiteSpace(result[i][0]) || char.IsWhiteSpace(result[i][^1])))
                        nodes[i].SetAttributeValue(XmlSpace, "preserve");
                }
                rep.Dashes += edits.Count(ed => ed.Kind == EditKind.Dash);
                rep.Markdown += edits.Count(ed => ed.Kind == EditKind.Markdown);
                rep.Placeholders += edits.Count(ed => ed.Kind == EditKind.Placeholder);
                changed = true;
            }
            if (changed) p.SetXml(e.Name, doc);
        }
    }

    // ---------------- Markdown ----------------

    static void CollectMarkdown(string s, List<Edit> edits)
    {
        var h = MdHeading.Match(s);
        if (h.Success) TryAdd(edits, new Edit(h.Index, h.Length, "", EditKind.Markdown));
        var b = MdBullet.Match(s);
        if (b.Success && !h.Success) TryAdd(edits, new Edit(b.Index, b.Length, "– ", EditKind.Markdown));

        foreach (Match m in MdLink.Matches(s))
        {
            // [текст](адрес) → текст
            TryAdd(edits, new Edit(m.Index, 1, "", EditKind.Markdown));
            int tail = m.Groups[1].Index + m.Groups[1].Length;
            TryAdd(edits, new Edit(tail, m.Index + m.Length - tail, "", EditKind.Markdown));
        }
        foreach (Match m in MdBold.Matches(s))
        {
            int mk = m.Groups[1].Length;
            if (Overlaps(edits, m.Index, mk) || Overlaps(edits, m.Index + m.Length - mk, mk)) continue;
            TryAdd(edits, new Edit(m.Index, mk, "", EditKind.Markdown));
            TryAdd(edits, new Edit(m.Index + m.Length - mk, mk, "", EditKind.Markdown));
        }
        foreach (Match m in MdCode.Matches(s))
        {
            if (Overlaps(edits, m.Index, 1) || Overlaps(edits, m.Index + m.Length - 1, 1)) continue;
            TryAdd(edits, new Edit(m.Index, 1, "", EditKind.Markdown));
            TryAdd(edits, new Edit(m.Index + m.Length - 1, 1, "", EditKind.Markdown));
        }
    }

    // ---------------- шаблоны [Ваше имя] ----------------

    static readonly Regex PhName = new(@"имя|фио|автор|преподавател|студент|name|author", O);
    static readonly Regex PhOrg = new(@"организац|компани|вуз|университет|кафедр|учрежден|company|organi[sz]ation|university", O);
    static readonly Regex PhDate = new(@"дат|число|date", O);

    static void CollectPlaceholders(string s, DocProps props, List<Edit> edits, TextFixReport rep)
    {
        foreach (Match m in Signatures.Placeholder.Matches(s))
        {
            var inner = m.Value;
            string? value = null;
            if (PhDate.IsMatch(inner)) value = DateFor(props);
            else if (PhOrg.IsMatch(inner)) value = props.Get(Fields.Company).Trim();
            else if (PhName.IsMatch(inner)) value = props.Get(Fields.Creator).Trim();
            if (string.IsNullOrEmpty(value)) { rep.Unfilled.Add(m.Value); continue; }
            TryAdd(edits, new Edit(m.Index, m.Length, value, EditKind.Placeholder));
        }
    }

    static string DateFor(DocProps props)
    {
        if (Fields.TryParseLocal(props.Get(Fields.Modified), out var utc))
            return utc.ToLocalTime().ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
        return DateTime.Now.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
    }

    // ---------------- применение правок ----------------

    static bool Overlaps(List<Edit> edits, int start, int len) =>
        edits.Any(e => start < e.Start + e.Length && e.Start < start + len);

    static void TryAdd(List<Edit> edits, Edit e)
    {
        if (e.Length <= 0 || Overlaps(edits, e.Start, e.Length)) return;
        edits.Add(e);
    }

    /// <summary>
    /// Применяет правки к строке, склеенной из текстов узлов, и раскладывает результат обратно
    /// по тем же узлам: текст замены попадает в узел, где начинается правка.
    /// </summary>
    static string[] ApplyEdits(string[] texts, List<Edit> edits)
    {
        var full = string.Concat(texts);
        var owner = new int[full.Length];
        for (int i = 0, pos = 0; i < texts.Length; i++)
            for (int j = 0; j < texts[i].Length; j++) owner[pos++] = i;

        var outp = texts.Select(_ => new StringBuilder()).ToArray();
        var sorted = edits.OrderBy(e => e.Start).ToList();
        int k = 0;
        for (int pos = 0; pos < full.Length;)
        {
            if (k < sorted.Count && sorted[k].Start == pos)
            {
                var e = sorted[k++];
                outp[owner[pos]].Append(e.Replacement);
                pos += e.Length;
                continue;
            }
            outp[owner[pos]].Append(full[pos]);
            pos++;
        }
        return outp.Select(sb => sb.ToString()).ToArray();
    }

    static readonly HashSet<string> KeepMarkers = new()
    {
        "sectPr", "commentRangeStart", "commentRangeEnd", "commentReference", "bookmarkStart", "bookmarkEnd",
        "ins", "del", "footnoteReference", "endnoteReference", "fldChar", "fldSimple", "drawing", "pict", "object",
    };

    /// <summary>Удаляет абзац, а если это небезопасно (единственный в ячейке, есть закладки и т.п.) — только его текст.</summary>
    static void RemoveParagraph(XElement para, List<XElement> nodes, DocKind kind, XName pName)
    {
        bool canRemove = kind != DocKind.Excel
                         && para.Parent != null
                         && para.Parent.Elements(pName).Count() > 1
                         && !para.Descendants().Any(d => KeepMarkers.Contains(d.Name.LocalName));
        if (canRemove) para.Remove();
        else foreach (var n in nodes) n.Value = "";
    }

    // ---------------- alt-текст и имена картинок ----------------

    public static bool IsSuspiciousName(string v) =>
        Signatures.PathLike.IsMatch(v) || Signatures.ImageName.IsMatch(v) || Signatures.Ai.IsMatch(v);

    public static bool IsAltPart(string name)
    {
        var n = name.ToLowerInvariant();
        return n.EndsWith(".xml") && (n.StartsWith("word/") || n.StartsWith("ppt/slides/") ||
               n.StartsWith("ppt/slidelayouts/") || n.StartsWith("ppt/slidemasters/") ||
               n.StartsWith("ppt/notesslides/") || n.StartsWith("xl/drawings/"));
    }

    static int CleanAlt(OoxmlPackage p)
    {
        int count = 0;
        foreach (var e in p.Entries.Where(en => IsAltPart(en.Name)).ToList())
        {
            if (!OoxmlPackage.ContainsAscii(e.Data, "descr=") && !OoxmlPackage.ContainsAscii(e.Data, "name=")) continue;
            var doc = p.Xml(e.Name);
            if (doc?.Root == null) continue;
            bool changed = false;
            foreach (var el in doc.Root.Descendants().Where(d => d.Name.LocalName is "docPr" or "cNvPr"))
            {
                foreach (var attr in new[] { "descr", "title" })
                {
                    var v = (string?)el.Attribute(attr);
                    if (string.IsNullOrEmpty(v) || !IsSuspiciousName(v)) continue;
                    el.Attribute(attr)!.Remove();
                    changed = true;
                    count++;
                }
                var name = (string?)el.Attribute("name");
                if (!string.IsNullOrEmpty(name) && IsSuspiciousName(name))
                {
                    el.SetAttributeValue("name", "Рисунок " + ((string?)el.Attribute("id") ?? "1"));
                    changed = true;
                    count++;
                }
            }
            if (changed) p.SetXml(e.Name, doc);
        }
        return count;
    }
}
