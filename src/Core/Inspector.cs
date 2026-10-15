using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace MetaCleaner.Core;

/// <summary>Поиск следов генерации и автоматической сборки в документе.</summary>
public static class Inspector
{
    const string Auto = "Авто", Resave = "Пересохранить в Office", Manual = "Вручную", None = "—",
        FixText = "Авто (правка текста)";

    static readonly XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static readonly XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public static InspectResult InspectFile(string path)
    {
        var kind = Fields.KindFromPath(path) ?? DocKind.Word;
        var res = new InspectResult { Path = path, Kind = kind };
        OoxmlPackage p;
        try { p = OoxmlPackage.Load(path); }
        catch (Exception ex)
        {
            res.Error = ex.Message;
            res.Findings.Add(new Finding
            {
                Severity = Severity.High, Category = "Файл",
                Message = $"Не удалось открыть как документ Office: {ex.Message}", Fix = Manual,
            });
            foreach (var f in Fields.All) res.Props[f.Key] = "";
            return res;
        }
        Inspect(p, res);
        CheckZone(path, res);
        if (Path.GetFileName(path).Contains('—'))
            Add(res, Severity.Info, "Файл", "Длинное тире «—» в имени файла", FixText);
        return res;
    }

    public static void Inspect(OoxmlPackage p, InspectResult res)
    {
        res.Props = PropsIO.Read(p);
        CheckProps(p, res.Props, res.Kind, res, "");
        CheckStructure(p, res);
        CheckZip(p, res);
        CheckReviewers(p, res);
        CheckImages(p, res, "");
        CheckEmbedded(p, res);
        CheckLinks(p, res);
        CheckContent(p, res);
        CheckText(p, res);
    }

    static void Add(InspectResult r, Severity s, string cat, string msg, string fix) =>
        r.Findings.Add(new Finding { Severity = s, Category = cat, Message = msg, Fix = fix });

    // ---------------- свойства ----------------

    static readonly string[] MetaKeys =
    {
        Fields.Creator, Fields.LastModifiedBy, Fields.Description, Fields.Company, Fields.Manager,
        Fields.Template, Fields.Application, Fields.Category, Fields.ContentStatus,
    };

    static readonly string[] ContentKeys = { Fields.Title, Fields.Subject, Fields.Keywords };

    static void CheckProps(OoxmlPackage p, DocProps d, DocKind kind, InspectResult r, string prefix)
    {
        const string cat = "Свойства";
        bool main = prefix == "";
        void Mark(string key) { if (main) r.SuspiciousFields.Add(key); }

        if (!p.Has(PropsIO.CorePath(p)))
            Add(r, Severity.High, cat, prefix + "Нет основных свойств (docProps/core.xml) — Office всегда их создаёт", Auto);
        if (!p.Has(PropsIO.AppPath(p)))
            Add(r, Severity.High, cat, prefix + "Нет расширенных свойств (docProps/app.xml) — Office всегда их создаёт", Auto);

        foreach (var k in MetaKeys)
        {
            var v = d.Get(k);
            if (v == "") continue;
            if (Signatures.Generator.IsMatch(v))
            { Add(r, Severity.High, cat, $"{prefix}«{Fields.Label(k)}» = «{v}» — след программы-генератора", Auto); Mark(k); }
            else if (Signatures.Ai.IsMatch(v))
            { Add(r, Severity.High, cat, $"{prefix}«{Fields.Label(k)}» = «{v}» — упоминание ИИ", Auto); Mark(k); }
        }
        foreach (var k in ContentKeys)
        {
            var v = d.Get(k);
            if (v == "") continue;
            if (Signatures.Generator.IsMatch(v) || Signatures.Ai.IsMatch(v))
            { Add(r, Severity.Medium, cat, $"{prefix}«{Fields.Label(k)}» = «{v}» — проверьте, уместно ли", Auto); Mark(k); }
        }

        if (!main) return; // для встроенных объектов достаточно проверки значений

        var desc = d.Get(Fields.Description);
        if (desc != "" && !r.SuspiciousFields.Contains(Fields.Description))
        { Add(r, Severity.Low, cat, $"Заполнен комментарий: «{Trunc(desc, 80)}»", Auto); Mark(Fields.Description); }

        if (d.Get(Fields.Creator) == "")
        { Add(r, Severity.Medium, cat, "Не указан автор", Auto); Mark(Fields.Creator); }
        if (d.Get(Fields.LastModifiedBy) == "")
        { Add(r, Severity.Low, cat, "Не указано, кем сохранено", Auto); Mark(Fields.LastModifiedBy); }

        var app = d.Get(Fields.Application);
        if (app == "")
        { Add(r, Severity.High, cat, "Не указано приложение — Microsoft Office всегда его записывает", Auto); Mark(Fields.Application); }
        else if (!app.Contains("Microsoft", StringComparison.OrdinalIgnoreCase))
        { Add(r, Severity.Medium, cat, $"Создан не в Microsoft Office: «{app}»", Auto); Mark(Fields.Application); }
        if (d.Get(Fields.AppVersion) == "")
        { Add(r, Severity.Medium, cat, "Не указана версия приложения", Auto); Mark(Fields.AppVersion); }

        var rev = d.Get(Fields.Revision);
        if (kind != DocKind.Excel)
        {
            if (rev == "") { Add(r, Severity.Medium, cat, "Нет номера редакции", Auto); Mark(Fields.Revision); }
            else if (rev == "1") { Add(r, Severity.Low, cat, "Редакция 1 — документ ни разу не пересохранялся", Auto); Mark(Fields.Revision); }

            var tt = d.Get(Fields.TotalTime);
            if (tt is "" or "0")
            { Add(r, Severity.Medium, cat, "Нулевое время правки — документ не редактировался вручную", Auto); Mark(Fields.TotalTime); }
        }

        bool hasC = Fields.TryParseLocal(d.Get(Fields.Created), out var c);
        bool hasM = Fields.TryParseLocal(d.Get(Fields.Modified), out var m);
        if (!hasC) { Add(r, Severity.Medium, cat, "Нет даты создания", Auto); Mark(Fields.Created); }
        if (!hasM) { Add(r, Severity.Medium, cat, "Нет даты изменения", Auto); Mark(Fields.Modified); }
        if (hasC && hasM)
        {
            if (c > m) { Add(r, Severity.Medium, cat, "Дата создания позже даты изменения", Auto); Mark(Fields.Created); }
            else if (c == m) { Add(r, Severity.Low, cat, "Даты создания и изменения совпадают до секунды", Auto); Mark(Fields.Modified); }
        }
        var tpl = hasC ? Signatures.TemplateDates.FirstOrDefault(t => Math.Abs((t.Utc - c).TotalMinutes) < 1) : default;
        if (tpl.Tool != null)
        { Add(r, Severity.High, cat, $"Дата создания совпадает со встроенным шаблоном {tpl.Tool}", Auto); Mark(Fields.Created); }
        else if (hasC && c.Year < 2015)
        { Add(r, Severity.Low, cat, $"Очень старая дата создания ({c.Year}) — возможно, взята из шаблона генератора", Auto); Mark(Fields.Created); }
        if (hasM && m > DateTime.UtcNow.AddDays(1))
        { Add(r, Severity.Medium, cat, "Дата изменения в будущем", Auto); Mark(Fields.Modified); }

        if (d.Custom.Count > 0)
        {
            bool ai = d.Custom.Any(x => Signatures.Ai.IsMatch(x.Name + " " + x.Value) || Signatures.Generator.IsMatch(x.Name + " " + x.Value));
            var names = string.Join(", ", d.Custom.Take(6).Select(x => x.Name)) + (d.Custom.Count > 6 ? "…" : "");
            Add(r, ai ? Severity.High : Severity.Medium, cat, $"Пользовательские свойства ({d.Custom.Count}): {names}", Auto);
        }
    }

    // ---------------- структура ----------------

    static void CheckStructure(OoxmlPackage p, InspectResult r)
    {
        const string cat = "Структура";
        switch (r.Kind)
        {
            case DocKind.Word:
                if (!p.Has("word/settings.xml")) Add(r, Severity.High, cat, "Нет word/settings.xml — Word всегда его создаёт", Resave);
                if (!p.Has("word/fontTable.xml")) Add(r, Severity.Medium, cat, "Нет таблицы шрифтов (fontTable.xml)", Resave);
                if (!p.Has("word/webSettings.xml")) Add(r, Severity.Low, cat, "Нет webSettings.xml", Resave);
                if (!p.Has("word/styles.xml")) Add(r, Severity.Medium, cat, "Нет стилей (styles.xml)", Resave);
                if (!p.Matching(n => n.StartsWith("word/theme/", StringComparison.OrdinalIgnoreCase)).Any())
                    Add(r, Severity.Medium, cat, "Нет темы оформления (theme)", Resave);
                var doc = p.Get("word/document.xml");
                if (doc != null && !OoxmlPackage.ContainsAscii(doc, "w:rsidR"))
                    Add(r, Severity.Medium, cat, "Нет RSID-меток правки — документ ни разу не редактировался в Word", Resave);
                var set = p.Get("word/settings.xml");
                if (set != null && !OoxmlPackage.ContainsAscii(set, "docId"))
                    Add(r, Severity.Low, cat, "Нет идентификатора документа (w14/w15:docId)", Resave);
                if (set != null && !OoxmlPackage.ContainsAscii(set, "w:rsids"))
                    Add(r, Severity.Low, cat, "В настройках нет списка сессий правки (w:rsids)", Resave);
                break;
            case DocKind.PowerPoint:
                if (!p.Has("ppt/presProps.xml")) Add(r, Severity.Medium, cat, "Нет presProps.xml", Resave);
                if (!p.Has("ppt/viewProps.xml")) Add(r, Severity.Medium, cat, "Нет viewProps.xml", Resave);
                if (!p.Has("ppt/tableStyles.xml")) Add(r, Severity.Low, cat, "Нет tableStyles.xml", Resave);
                if (p.PartByRelType(PropsIO.RelThumb) == null)
                    Add(r, Severity.Medium, cat, "Нет миниатюры (thumbnail) — PowerPoint всегда её сохраняет", Resave);
                break;
            case DocKind.Excel:
                if (!p.Matching(n => n.StartsWith("xl/theme/", StringComparison.OrdinalIgnoreCase)).Any())
                    Add(r, Severity.Low, cat, "Нет темы оформления (theme)", Resave);
                if (!p.Has("xl/styles.xml")) Add(r, Severity.Medium, cat, "Нет стилей (styles.xml)", Resave);
                break;
        }
    }

    static void CheckZip(OoxmlPackage p, InspectResult r)
    {
        const string cat = "Архив";
        var real = p.Entries.FirstOrDefault(e => e.Time.Year > 1980);
        if (real != null)
            Add(r, Severity.Medium, cat,
                $"Внутри архива реальные даты файлов (напр. {real.Time.LocalDateTime:yyyy-MM-dd HH:mm}); Office ставит 1980-01-01", Auto);
        if (p.Entries.Count > 0 && !string.Equals(p.Entries[0].Name, "[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
            Add(r, Severity.Low, cat, "Порядок файлов в архиве не как у Office", Auto);
    }

    // ---------------- содержимое ----------------

    static void CheckReviewers(OoxmlPackage p, InspectResult r)
    {
        var names = Reviewers.Collect(p, r.Kind);
        if (names.Count == 0) return;
        bool bad = names.Any(n => Signatures.Ai.IsMatch(n) || Signatures.Generator.IsMatch(n));
        Add(r, bad ? Severity.High : Severity.Low, "Примечания",
            $"Авторы примечаний/исправлений: {string.Join(", ", names.Take(5))}", Auto);
    }

    static void CheckImages(OoxmlPackage p, InspectResult r, string prefix)
    {
        int count = 0; bool ai = false; var kinds = new HashSet<string>();
        foreach (var e in p.Matching(n => n.Contains("/media/", StringComparison.OrdinalIgnoreCase) && ImageMeta.IsImage(n)))
        {
            var d = ImageMeta.Describe(e.Data);
            if (d.Count == 0) continue;
            count++;
            foreach (var (t, isAi) in d) { kinds.Add(t); ai |= isAi; }
        }
        if (count > 0)
            Add(r, ai ? Severity.High : Severity.Medium, "Изображения",
                $"{prefix}Метаданные в {count} изобр.: {string.Join("; ", kinds.Take(4))}", Auto);
    }

    static void CheckEmbedded(OoxmlPackage p, InspectResult r)
    {
        foreach (var e in p.Matching(n => n.Contains("/embeddings/", StringComparison.OrdinalIgnoreCase)))
        {
            var k = Fields.KindFromPath(e.Name);
            if (k == null) continue;
            try
            {
                var sub = OoxmlPackage.Load(e.Data);
                var sp = PropsIO.Read(sub);
                var prefix = $"Встроенный объект {Path.GetFileName(e.Name)}: ";
                CheckProps(sub, sp, k.Value, r, prefix);
                CheckImages(sub, r, prefix);
                var who = sp.Get(Fields.Creator);
                if (who != "" && !Signatures.Generator.IsMatch(who) && !Signatures.Ai.IsMatch(who))
                    Add(r, Severity.Info, "Свойства", $"{prefix}автор «{who}»", Auto);
            }
            catch { /* не OOXML — пропускаем */ }
        }
    }

    static void CheckLinks(OoxmlPackage p, InspectResult r)
    {
        int utm = 0; var domains = new List<string>();
        foreach (var e in p.Matching(n => n.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
        {
            if (!OoxmlPackage.ContainsAscii(e.Data, "External")) continue;
            var x = p.Xml(e.Name);
            foreach (var rel in x?.Root?.Elements(XName.Get("Relationship", OoxmlPackage.NsRel)) ?? Enumerable.Empty<XElement>())
            {
                if ((string?)rel.Attribute("TargetMode") != "External") continue;
                var t = (string?)rel.Attribute("Target") ?? "";
                if (HasAiUtm(t)) utm++;
                else if (Signatures.AiDomain.IsMatch(t)) domains.Add(Trunc(t, 60));
            }
        }
        if (utm > 0)
            Add(r, Severity.High, "Ссылки", $"{utm} ссыл. с меткой utm_source от ИИ-сервиса (например, chatgpt.com)", Auto);
        if (domains.Count > 0)
            Add(r, Severity.High, "Ссылки", $"Ссылки на ИИ-сервисы: {string.Join(", ", domains.Distinct().Take(3))}", Manual);
    }

    public static bool HasAiUtm(string url)
    {
        int q = url.IndexOf('?');
        if (q < 0) return false;
        var query = url[(q + 1)..];
        int h = query.IndexOf('#');
        if (h >= 0) query = query[..h];
        return query.Split('&').Any(s => Signatures.AiUtmParam.IsMatch(s));
    }

    static void CheckContent(OoxmlPackage p, InspectResult r)
    {
        if (r.Kind == DocKind.Word)
        {
            int hidden = p.Matching(n => n.StartsWith("word/") && n.EndsWith(".xml"))
                .Sum(e => CountAscii(e.Data, "<w:vanish/>") + CountAscii(e.Data, "<w:vanish "));
            if (hidden > 0)
                Add(r, Severity.Medium, "Содержимое", $"Скрытый текст: {hidden} фрагм. — проверьте, что там", Manual);
        }

        var odd = new List<string>();
        bool paths = false;
        foreach (var e in p.Entries.Where(en => TextFixer.IsAltPart(en.Name)))
        {
            if (!OoxmlPackage.ContainsAscii(e.Data, "descr=") && !OoxmlPackage.ContainsAscii(e.Data, "name=")) continue;
            var x = p.Xml(e.Name);
            if (x?.Root == null) continue;
            foreach (var el in x.Root.Descendants().Where(d => d.Name.LocalName is "docPr" or "cNvPr"))
                foreach (var attr in new[] { "descr", "title", "name" })
                {
                    var v = (string?)el.Attribute(attr);
                    if (string.IsNullOrEmpty(v) || !TextFixer.IsSuspiciousName(v)) continue;
                    odd.Add(Trunc(v, 50));
                    paths |= Signatures.PathLike.IsMatch(v);
                }
        }
        if (odd.Count > 0)
            Add(r, paths ? Severity.High : Severity.Medium, "Содержимое",
                (paths ? "Путь к файлу с компьютера генерации в alt-тексте/имени картинки: " : "Подозрительные имена/alt-текст объектов: ") +
                string.Join(", ", odd.Distinct().Take(4)), FixText);

        if (r.Kind == DocKind.PowerPoint)
        {
            int notes = 0;
            foreach (var e in p.Matching(n => n.StartsWith("ppt/notesSlides/", StringComparison.OrdinalIgnoreCase) && n.EndsWith(".xml")))
            {
                var x = p.Xml(e.Name);
                var text = string.Concat(x?.Descendants(a + "t").Select(t => t.Value) ?? Enumerable.Empty<string>());
                if (text.Trim().Length > 3) notes++;
            }
            if (notes > 0)
                Add(r, Severity.Info, "Содержимое", $"Заметки докладчика на {notes} слайд(ах) — просмотрите их", Manual);
        }
    }

    // ---------------- текст ----------------

    public static List<string> ExtractParagraphs(OoxmlPackage p, DocKind kind)
    {
        var list = new List<string>();
        foreach (var e in p.Entries)
        {
            var n = e.Name.ToLowerInvariant();
            bool take = kind switch
            {
                DocKind.Word => n.StartsWith("word/") && n.EndsWith(".xml") &&
                                (n.Contains("document") || n.Contains("header") || n.Contains("footer") ||
                                 n.Contains("footnotes") || n.Contains("endnotes") || n.Contains("comments.xml")),
                DocKind.PowerPoint => (n.StartsWith("ppt/slides/slide") || n.StartsWith("ppt/notesslides/")) && n.EndsWith(".xml"),
                _ => n == "xl/sharedstrings.xml" || (n.StartsWith("xl/comments") && n.EndsWith(".xml")),
            };
            if (!take) continue;
            XDocument? x;
            try { x = p.Xml(e.Name); } catch { continue; }
            if (x?.Root == null) continue;
            switch (kind)
            {
                case DocKind.Word:
                    foreach (var para in x.Root.Descendants(w + "p"))
                        list.Add(string.Concat(para.Descendants(w + "t").Select(t => t.Value)));
                    break;
                case DocKind.PowerPoint:
                    foreach (var para in x.Root.Descendants(a + "p"))
                        list.Add(string.Concat(para.Descendants(a + "t").Select(t => t.Value)));
                    break;
                default:
                    foreach (var si in x.Root.Descendants().Where(el => el.Name.LocalName is "si" or "comment"))
                        list.Add(string.Concat(si.Descendants().Where(t => t.Name.LocalName == "t").Select(t => t.Value)));
                    break;
            }
        }
        return list.Where(s => s.Trim() != "").ToList();
    }

    static void CheckText(OoxmlPackage p, InspectResult r)
    {
        const string cat = "Текст";
        var paras = ExtractParagraphs(p, r.Kind);
        if (paras.Count == 0) return;

        static bool IsChat(string t)
        {
            var s = t.Trim();
            return (s.Length < 400 && Signatures.ChatPreamble.IsMatch(s)) || Signatures.TrailingOffer.IsMatch(t);
        }

        void Rule(Func<string, bool> test, Severity s, string what, string fix)
        {
            var hits = paras.Where(test).ToList();
            if (hits.Count == 0) return;
            var ex = string.Join(" | ", hits.Take(2).Select(t => "«" + Trunc(t.Trim(), 60) + "»"));
            Add(r, s, cat, $"{what}: {hits.Count} абз. Например: {ex}", fix);
        }

        Rule(IsChat, Severity.Medium, "Служебные фразы чата (удаляются абзац или последнее предложение)", FixText);
        Rule(t => !IsChat(t) && Signatures.AiPhrases.IsMatch(t), Severity.Medium, "Упоминания ИИ и типичные фразы ответов", Manual);
        Rule(t => Signatures.Placeholder.IsMatch(t), Severity.Medium,
            "Незаполненные шаблоны вида [Ваше имя] (имя, организация и дата подставляются из свойств)", FixText);
        Rule(t => Signatures.Markdown.IsMatch(t), Severity.Low, "Остатки разметки Markdown (**, ##, ```, `код`)", FixText);
        Rule(t => Signatures.Emoji.IsMatch(t), Severity.Low, "Эмодзи в тексте", Manual);

        int dashes = paras.Sum(t => t.Count(ch => ch == '—'));
        if (dashes > 0)
            Add(r, Severity.Info, cat, $"Длинных тире «—»: {dashes} (заменяются на «–» или «-»)", FixText);
    }

    // ---------------- отметка загрузки ----------------

    static void CheckZone(string path, InspectResult r)
    {
        var z = ZoneId.Read(path);
        if (z == null) return;
        var urls = ZoneId.Urls(z);
        bool ai = urls.Any(u => Signatures.AiDomain.IsMatch(u));
        var where = urls.Count > 0 ? ": " + string.Join(", ", urls.Select(u => Trunc(u, 60))) : "";
        Add(r, ai ? Severity.High : Severity.Medium, "Windows",
            $"Отметка «Загружено из интернета»{where}", Auto);
    }

    // ---------------- утилиты ----------------

    static int CountAscii(byte[] data, string s)
    {
        var pat = System.Text.Encoding.ASCII.GetBytes(s);
        int count = 0, pos = 0;
        var span = data.AsSpan();
        while (true)
        {
            int i = span[pos..].IndexOf(pat);
            if (i < 0) return count;
            count++;
            pos += i + pat.Length;
        }
    }

    static string Trunc(string s, int n) => s.Length > n ? s[..n] + "…" : s;
}
