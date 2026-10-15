using System.Globalization;
using System.Xml.Linq;

namespace MetaCleaner.Core;

/// <summary>Чтение и запись свойств: docProps/core.xml, app.xml, custom.xml.</summary>
public static class PropsIO
{
    static readonly XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    static readonly XNamespace dc = "http://purl.org/dc/elements/1.1/";
    static readonly XNamespace dcterms = "http://purl.org/dc/terms/";
    static readonly XNamespace dcmitype = "http://purl.org/dc/dcmitype/";
    static readonly XNamespace xsi = "http://www.w3.org/2001/XMLSchema-instance";
    static readonly XNamespace ep = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    static readonly XNamespace vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
    static readonly XNamespace cu = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

    public const string RelCore = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    public const string RelApp = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    public const string RelCustom = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
    public const string RelThumb = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";
    const string CtCore = "application/vnd.openxmlformats-package.core-properties+xml";
    const string CtApp = "application/vnd.openxmlformats-officedocument.extended-properties+xml";

    // Порядок элементов такой же, как пишет Microsoft Office.
    static readonly (string Key, XName Name)[] CoreMap =
    {
        (Fields.Title, dc + "title"), (Fields.Subject, dc + "subject"), (Fields.Creator, dc + "creator"),
        (Fields.Keywords, cp + "keywords"), (Fields.Description, dc + "description"),
        (Fields.LastModifiedBy, cp + "lastModifiedBy"), (Fields.Revision, cp + "revision"),
        (Fields.Created, dcterms + "created"), (Fields.Modified, dcterms + "modified"),
        (Fields.Category, cp + "category"), (Fields.ContentStatus, cp + "contentStatus"),
    };

    static readonly (string Key, string El)[] AppMap =
    {
        (Fields.Template, "Template"), (Fields.Manager, "Manager"), (Fields.Company, "Company"),
        (Fields.TotalTime, "TotalTime"), (Fields.Application, "Application"), (Fields.AppVersion, "AppVersion"),
    };

    public static string CorePath(OoxmlPackage p) => p.PartByRelType(RelCore) ?? "docProps/core.xml";
    public static string AppPath(OoxmlPackage p) => p.PartByRelType(RelApp) ?? "docProps/app.xml";
    public static string CustomPath(OoxmlPackage p) => p.PartByRelType(RelCustom) ?? "docProps/custom.xml";

    public static DocProps Read(OoxmlPackage p)
    {
        var d = new DocProps();
        foreach (var f in Fields.All) d[f.Key] = "";

        var core = p.Xml(CorePath(p));
        if (core?.Root != null)
            foreach (var (k, n) in CoreMap)
            {
                var v = core.Root.Element(n)?.Value.Trim() ?? "";
                d[k] = k is Fields.Created or Fields.Modified ? Fields.FormatStored(v) : v;
            }

        var app = p.Xml(AppPath(p));
        if (app?.Root != null)
            foreach (var (k, e) in AppMap)
                d[k] = app.Root.Element(ep + e)?.Value.Trim() ?? "";

        var cus = p.Xml(CustomPath(p));
        if (cus?.Root != null)
            foreach (var pr in cus.Root.Elements(cu + "property"))
                d.Custom.Add(((string?)pr.Attribute("name") ?? "", pr.Value.Trim()));
        return d;
    }

    /// <summary>Записывает свойства. core.xml пересобирается целиком, app.xml правится точечно.</summary>
    public static void Write(OoxmlPackage p, DocProps d, DocKind kind)
    {
        // --- core.xml ---
        var corePath = p.PartByRelType(RelCore);
        if (corePath == null) { corePath = "docProps/core.xml"; p.AddRel(RelCore, corePath); }
        p.SetOverride("/" + corePath, CtCore);

        var root = new XElement(cp + "coreProperties",
            new XAttribute(XNamespace.Xmlns + "cp", cp.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "dc", dc.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "dcterms", dcterms.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "dcmitype", dcmitype.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "xsi", xsi.NamespaceName));

        var nowUtc = DateTime.UtcNow;
        foreach (var (k, n) in CoreMap)
        {
            var v = d.Get(k).Trim();
            if (k is Fields.Created or Fields.Modified)
            {
                DateTime utc;
                if (v == "") utc = nowUtc;
                else if (!Fields.TryParseLocal(v, out utc))
                    throw new FormatException($"неверная дата в поле «{Fields.Label(k)}»: «{v}» (нужно {Fields.DateFormat})");
                root.Add(new XElement(n, new XAttribute(xsi + "type", "dcterms:W3CDTF"),
                    utc.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture)));
                continue;
            }
            if (k == Fields.Revision && v != "" && !int.TryParse(v, out _))
                throw new FormatException($"редакция должна быть числом: «{v}»");
            if (v != "") root.Add(new XElement(n, v));
        }
        p.SetXml(corePath, new XDocument(root));

        // --- app.xml ---
        var appPath = p.PartByRelType(RelApp);
        XDocument app;
        if (appPath == null || !p.Has(appPath))
        {
            if (appPath == null) { appPath = "docProps/app.xml"; p.AddRel(RelApp, appPath); }
            app = NewApp(kind);
        }
        else app = p.Xml(appPath)!;
        p.SetOverride("/" + appPath, CtApp);

        var ar = app.Root!;
        foreach (var (k, e) in AppMap)
        {
            var v = d.Get(k).Trim();
            var el = ar.Element(ep + e);
            if (v == "") { el?.Remove(); continue; }
            if (k == Fields.TotalTime && !int.TryParse(v, out _))
                throw new FormatException($"время правки должно быть числом минут: «{v}»");
            if (el == null) { el = new XElement(ep + e); ar.Add(el); }
            el.Value = v;
        }
        p.SetXml(appPath, app);
    }

    static XDocument NewApp(DocKind kind)
    {
        var root = new XElement(ep + "Properties",
            new XAttribute("xmlns", ep.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "vt", vt.NamespaceName));
        if (kind == DocKind.Word) root.Add(new XElement(ep + "Template", "Normal.dotm"));
        if (kind != DocKind.Excel) root.Add(new XElement(ep + "TotalTime", "0"));
        root.Add(new XElement(ep + "Application", Fields.ExpectedApplication(kind)),
            new XElement(ep + "DocSecurity", "0"), new XElement(ep + "ScaleCrop", "false"),
            new XElement(ep + "LinksUpToDate", "false"), new XElement(ep + "SharedDoc", "false"),
            new XElement(ep + "HyperlinksChanged", "false"), new XElement(ep + "AppVersion", "16.0000"));
        return new XDocument(root);
    }

    /// <summary>Удаляет docProps/custom.xml вместе со связью и типом содержимого.</summary>
    public static int RemoveCustom(OoxmlPackage p)
    {
        var path = CustomPath(p);
        int count = p.Xml(path)?.Root?.Elements(cu + "property").Count() ?? 0;
        bool had = p.Has(path) || p.PartByRelType(RelCustom) != null;
        p.Remove(path);
        p.RemoveRelByType(RelCustom);
        p.RemoveOverride("/" + path);
        return had ? Math.Max(count, 1) : 0;
    }
}
