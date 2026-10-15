using System.Xml.Linq;

namespace MetaCleaner.Core;

public sealed class SaveOptions
{
    /// <summary>true — перезаписать исходные файлы; false — копии в папку <see cref="OutputFolderName"/>.</summary>
    public bool InPlace { get; set; }
    public bool MakeBackup { get; set; } = true;
    public string OutputFolderName { get; set; } = "cleaned";

    public bool RemoveCustom { get; set; } = true;
    public bool ReplaceReviewers { get; set; } = true;
    /// <summary>Имя для примечаний/исправлений; пусто — автор документа.</summary>
    public string Reviewer { get; set; } = "";
    public bool StripImageMeta { get; set; } = true;
    public bool CleanEmbedded { get; set; } = true;
    public bool OfficeLikeZip { get; set; } = true;
    public bool StripAiUtm { get; set; } = true;
    public bool RemoveZoneId { get; set; } = true;
    public bool SetFileTimes { get; set; } = true;
    public bool OfficeResave { get; set; }
    /// <summary>Правка текста: тире, Markdown, фразы чата, шаблоны, alt-текст.</summary>
    public TextFixOptions Text { get; set; } = new();
}

public sealed class SaveResult
{
    public string OutputPath { get; set; } = "";
    public string? BackupPath { get; set; }
    public List<string> Actions { get; } = new();
}

public static class Cleaner
{
    public static string OutputPathFor(string src, string root, SaveOptions o)
    {
        if (o.InPlace) return src;
        string rel = Path.GetRelativePath(root, src);
        if (rel.StartsWith("..") || Path.IsPathRooted(rel)) rel = Path.GetFileName(src);
        var target = Path.Combine(root, o.OutputFolderName, rel);
        return Path.Combine(Path.GetDirectoryName(target)!, TextFixer.FixFileName(Path.GetFileName(target), o.Text));
    }

    /// <summary>Записывает свойства и чистит файл. При ошибке исходник остаётся прежним.</summary>
    public static SaveResult SaveFile(string src, string root, DocProps props, SaveOptions o, OfficeSession? office)
    {
        var kind = Fields.KindFromPath(src) ?? throw new InvalidOperationException("неподдерживаемый тип файла");
        var res = new SaveResult { OutputPath = OutputPathFor(src, root, o) };
        var target = res.OutputPath;
        var tmp = target + ".mctmp";

        // проверить даты и числа до каких-либо изменений файлов
        PropsIO.Write(OoxmlPackage.Load(src), props, kind);

        if (o.InPlace)
        {
            if (o.MakeBackup)
            {
                var bak = src + ".bak";
                if (File.Exists(bak)) bak = $"{src}.{DateTime.Now:yyyyMMdd-HHmmss}.bak";
                File.Copy(src, bak);
                res.BackupPath = bak;
            }
        }
        else
        {
            Directory.CreateDirectory(Path.GetDirectoryName(target)!);
            File.Copy(src, target, overwrite: true);
        }

        try
        {
            if (o.RemoveZoneId && ZoneId.Remove(target)) res.Actions.Add("снята отметка «Загружено из интернета»");

            if (o.OfficeResave && office != null && OfficeSession.IsAvailable(kind))
            {
                office.Resave(target, kind);
                res.Actions.Add($"пересохранён через {Fields.KindName(kind)}");
            }

            var p = OoxmlPackage.Load(target);
            CleanPackage(p, kind, props, o, res, nested: false);

            var bytes = p.ToBytes(o.OfficeLikeZip);
            OoxmlPackage.Validate(bytes);
            File.WriteAllBytes(tmp, bytes);
            File.Move(tmp, target, overwrite: true);
            if (o.OfficeLikeZip) res.Actions.Add("архив пересобран как у Office");

            if (o.RemoveZoneId) ZoneId.Remove(target);
            if (o.SetFileTimes) SetTimes(target, props);

            var fixedName = TextFixer.FixFileName(Path.GetFileName(src), o.Text);
            if (fixedName != Path.GetFileName(src))
            {
                if (o.InPlace)
                {
                    var renamed = Path.Combine(Path.GetDirectoryName(src)!, fixedName);
                    if (!File.Exists(renamed))
                    {
                        File.Move(src, renamed);
                        res.OutputPath = renamed;
                        res.Actions.Add($"файл переименован: {fixedName}");
                    }
                    else res.Actions.Add($"не переименован: файл «{fixedName}» уже существует");
                }
                else res.Actions.Add($"имя файла: {fixedName}");
            }
        }
        catch
        {
            try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
            if (o.InPlace && res.BackupPath != null)
            {
                try { File.Copy(res.BackupPath, src, overwrite: true); } catch { }
            }
            else if (!o.InPlace)
            {
                try { File.Delete(target); } catch { }
            }
            throw;
        }
        return res;
    }

    static void CleanPackage(OoxmlPackage p, DocKind kind, DocProps props, SaveOptions o, SaveResult res, bool nested)
    {
        PropsIO.Write(p, props, kind);
        if (!nested) res.Actions.Add("свойства записаны");

        if (!nested && (o.Text.AnyText || o.Text.CleanAltText))
        {
            var tr = new TextFixReport();
            TextFixer.Apply(p, kind, props, o.Text, tr);
            if (tr.Dashes > 0) res.Actions.Add($"заменено длинных тире: {tr.Dashes}");
            if (tr.Markdown > 0) res.Actions.Add($"убрано разметки Markdown: {tr.Markdown}");
            if (tr.Placeholders > 0) res.Actions.Add($"заполнено шаблонов: {tr.Placeholders}");
            if (tr.AltTexts > 0) res.Actions.Add($"очищено alt-текстов и имён картинок: {tr.AltTexts}");
            if (tr.RemovedParagraphs.Count > 0)
            {
                res.Actions.Add($"удалено служебных фраз чата: {tr.RemovedParagraphs.Count}");
                foreach (var s in tr.RemovedParagraphs.Take(30))
                    res.Actions.Add("    − «" + (s.Length > 150 ? s[..150] + "…" : s) + "»");
            }
            if (tr.Unfilled.Count > 0)
                res.Actions.Add("не заполнены (нет данных в свойствах): " + string.Join(", ", tr.Unfilled.Distinct().Take(6)));
        }

        if (o.RemoveCustom)
        {
            int n = PropsIO.RemoveCustom(p);
            if (n > 0 && !nested) res.Actions.Add($"удалено пользовательских свойств: {n}");
        }

        if (o.ReplaceReviewers)
        {
            var who = o.Reviewer.Trim() != "" ? o.Reviewer.Trim() : props.Get(Fields.Creator);
            if (who != "")
            {
                int n = Reviewers.Replace(p, kind, who);
                if (n > 0 && !nested) res.Actions.Add($"заменено имён в примечаниях/исправлениях: {n}");
            }
        }

        if (o.StripImageMeta)
        {
            int n = 0;
            foreach (var e in p.Matching(x => x.Contains("/media/", StringComparison.OrdinalIgnoreCase) && ImageMeta.IsImage(x)).ToList())
            {
                var nb = ImageMeta.Strip(e.Data, out bool changed);
                if (changed) { e.Data = nb; n++; }
            }
            if (n > 0) res.Actions.Add((nested ? "  во встроенном объекте: " : "") + $"очищено изображений: {n}");
        }

        if (o.StripAiUtm)
        {
            int n = StripUtm(p);
            if (n > 0 && !nested) res.Actions.Add($"убраны ИИ-метки из ссылок: {n}");
        }

        if (o.CleanEmbedded && !nested)
        {
            int n = 0;
            foreach (var e in p.Matching(x => x.Contains("/embeddings/", StringComparison.OrdinalIgnoreCase)).ToList())
            {
                var sk = Fields.KindFromPath(e.Name);
                if (sk == null) continue;
                try
                {
                    var sub = OoxmlPackage.Load(e.Data);
                    var sp = PropsIO.Read(sub);
                    foreach (var k in new[] { Fields.Creator, Fields.LastModifiedBy, Fields.Company, Fields.Manager,
                                 Fields.Created, Fields.Modified })
                        sp[k] = props.Get(k);
                    sp[Fields.Description] = "";
                    sp[Fields.Keywords] = "";
                    sp[Fields.Application] = Fields.ExpectedApplication(sk.Value);
                    sp[Fields.AppVersion] = "16.0000";
                    CleanPackage(sub, sk.Value, sp, o, res, nested: true);
                    e.Data = sub.ToBytes(o.OfficeLikeZip);
                    n++;
                }
                catch { /* не OOXML или повреждён — оставляем как есть */ }
            }
            if (n > 0) res.Actions.Add($"очищено встроенных объектов (диаграмм): {n}");
        }
    }

    /// <summary>Удаляет utm_source=chatgpt.com и подобные метки из внешних ссылок.</summary>
    static int StripUtm(OoxmlPackage p)
    {
        int count = 0;
        foreach (var e in p.Matching(n => n.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)).ToList())
        {
            if (!OoxmlPackage.ContainsAscii(e.Data, "utm_source")) continue;
            var x = p.Xml(e.Name);
            if (x?.Root == null) continue;
            bool changed = false;
            foreach (var rel in x.Root.Elements(XName.Get("Relationship", OoxmlPackage.NsRel)))
            {
                var t = (string?)rel.Attribute("Target");
                if (t == null || !Inspector.HasAiUtm(t)) continue;
                rel.SetAttributeValue("Target", RemoveAiUtm(t));
                changed = true;
                count++;
            }
            if (changed) p.SetXml(e.Name, x);
        }
        return count;
    }

    public static string RemoveAiUtm(string url)
    {
        int q = url.IndexOf('?');
        if (q < 0) return url;
        string rest = url[(q + 1)..], frag = "";
        int h = rest.IndexOf('#');
        if (h >= 0) { frag = rest[h..]; rest = rest[..h]; }
        var parts = rest.Split('&').Where(s => s != "" && !Signatures.AiUtmParam.IsMatch(s)).ToArray();
        return url[..q] + (parts.Length > 0 ? "?" + string.Join("&", parts) : "") + frag;
    }

    static void SetTimes(string path, DocProps props)
    {
        try
        {
            if (Fields.TryParseLocal(props.Get(Fields.Created), out var c))
                File.SetCreationTimeUtc(path, c);
            if (Fields.TryParseLocal(props.Get(Fields.Modified), out var m))
            {
                File.SetLastWriteTimeUtc(path, m);
                File.SetLastAccessTimeUtc(path, m);
            }
        }
        catch { /* на некоторых ФС время создания не меняется */ }
    }
}

public static class Scanner
{
    public static List<string> FindFiles(string folder, bool recursive, string outputFolderName = "cleaned")
    {
        var opt = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
        var outDir = Path.Combine(Path.GetFullPath(folder), outputFolderName) + Path.DirectorySeparatorChar;
        return Directory.EnumerateFiles(folder, "*.*", new EnumerationOptions
            {
                RecurseSubdirectories = recursive, IgnoreInaccessible = true, AttributesToSkip = FileAttributes.System,
            })
            .Where(f => Fields.KindFromPath(f) != null)
            .Where(f => !Path.GetFileName(f).StartsWith("~$"))
            .Where(f => !Path.GetFullPath(f).StartsWith(outDir, StringComparison.OrdinalIgnoreCase))
            .OrderBy(f => f, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
    }
}
