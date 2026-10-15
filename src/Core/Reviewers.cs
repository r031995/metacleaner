using System.Xml.Linq;

namespace MetaCleaner.Core;

/// <summary>Имена авторов примечаний и исправлений внутри документа.</summary>
public static class Reviewers
{
    static readonly XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static readonly XNamespace w15 = "http://schemas.microsoft.com/office/word/2012/wordml";

    public static string Initials(string name) =>
        string.Concat(name.Split(new[] { ' ', '.', '-' }, StringSplitOptions.RemoveEmptyEntries)
            .Take(3).Select(s => char.ToUpperInvariant(s[0])));

    /// <summary>Все имена авторов примечаний/исправлений.</summary>
    public static List<string> Collect(OoxmlPackage p, DocKind kind)
    {
        var names = new List<string>();
        Visit(p, kind, n => { names.Add(n); return null; });
        return names.Where(n => n.Trim() != "").Distinct().ToList();
    }

    /// <summary>Заменяет всех авторов на указанное имя. Возвращает число замен.</summary>
    public static int Replace(OoxmlPackage p, DocKind kind, string name)
    {
        int count = 0;
        Visit(p, kind, old => { if (old != name) count++; return name; });
        return count;
    }

    /// <summary>map: старое имя → новое (null — не менять).</summary>
    static void Visit(OoxmlPackage p, DocKind kind, Func<string, string?> map)
    {
        foreach (var e in p.Entries.ToList())
        {
            var n = e.Name.ToLowerInvariant();
            bool word = kind == DocKind.Word && n.StartsWith("word/") && n.EndsWith(".xml");
            bool ppt = kind == DocKind.PowerPoint && (n == "ppt/commentauthors.xml" || n == "ppt/authors.xml");
            bool xlComments = kind == DocKind.Excel && n.StartsWith("xl/comments") && n.EndsWith(".xml");
            bool xlPersons = kind == DocKind.Excel && n.StartsWith("xl/persons/") && n.EndsWith(".xml");
            if (!(word || ppt || xlComments || xlPersons)) continue;
            if (!OoxmlPackage.ContainsAscii(e.Data, word ? "author" : ppt ? "name" : xlPersons ? "displayName" : "author"))
                continue;

            var x = p.Xml(e.Name);
            if (x?.Root == null) continue;
            bool changed = false;

            void SetAttr(XElement el, XName attr, string? value)
            {
                var a = el.Attribute(attr);
                if (a == null || value == null || a.Value == value) return;
                a.Value = value;
                changed = true;
            }

            if (word)
            {
                foreach (var el in x.Root.DescendantsAndSelf())
                {
                    var a = el.Attribute(w + "author");
                    if (a != null)
                    {
                        var nv = map(a.Value);
                        SetAttr(el, w + "author", nv);
                        if (nv != null) SetAttr(el, w + "initials", Initials(nv));
                    }
                    var pa = el.Attribute(w15 + "author");
                    if (pa != null && el.Name == w15 + "person")
                    {
                        var nv = map(pa.Value);
                        SetAttr(el, w15 + "author", nv);
                        var pi = el.Element(w15 + "presenceInfo");
                        if (pi != null && nv != null)
                        {
                            SetAttr(pi, w15 + "providerId", "None");
                            SetAttr(pi, w15 + "userId", nv);
                        }
                    }
                }
            }
            else if (ppt)
            {
                foreach (var el in x.Root.Descendants().Where(d => d.Name.LocalName is "cmAuthor" or "author"))
                {
                    var a = el.Attribute("name");
                    if (a == null) continue;
                    var nv = map(a.Value);
                    if (nv == null) continue;
                    SetAttr(el, "name", nv);
                    SetAttr(el, "initials", Initials(nv));
                    SetAttr(el, "userId", nv);
                    SetAttr(el, "providerId", "None");
                }
            }
            else if (xlComments)
            {
                foreach (var el in x.Root.Descendants().Where(d => d.Name.LocalName == "author"))
                {
                    var nv = map(el.Value);
                    if (nv != null && el.Value != nv) { el.Value = nv; changed = true; }
                }
            }
            else
            {
                foreach (var el in x.Root.Descendants().Where(d => d.Name.LocalName == "person"))
                {
                    var a = el.Attribute("displayName");
                    if (a == null) continue;
                    var nv = map(a.Value);
                    if (nv == null) continue;
                    SetAttr(el, "displayName", nv);
                    SetAttr(el, "userId", nv);
                    SetAttr(el, "providerId", "None");
                }
            }

            if (changed) p.SetXml(e.Name, x);
        }
    }
}
