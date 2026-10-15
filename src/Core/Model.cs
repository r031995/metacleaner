using System.Globalization;

namespace MetaCleaner.Core;

public enum DocKind { Word, PowerPoint, Excel }

public enum Severity { Info = 0, Low = 1, Medium = 2, High = 3 }

/// <summary>Ключи свойств документа и их подписи в интерфейсе.</summary>
public static class Fields
{
    public const string Title = "title", Subject = "subject", Creator = "creator", Keywords = "keywords",
        Description = "description", LastModifiedBy = "lastModifiedBy", Category = "category",
        ContentStatus = "contentStatus", Revision = "revision", Created = "created", Modified = "modified",
        Application = "application", AppVersion = "appVersion", Company = "company", Manager = "manager",
        Template = "template", TotalTime = "totalTime";

    public static readonly (string Key, string Label)[] All =
    {
        (Creator, "Автор"), (LastModifiedBy, "Кем сохранено"), (Description, "Комментарий"),
        (Title, "Название"), (Subject, "Тема"), (Keywords, "Ключевые слова"), (Category, "Категория"),
        (ContentStatus, "Состояние"), (Company, "Организация"), (Manager, "Руководитель"),
        (Application, "Приложение"), (AppVersion, "Версия прил."), (Template, "Шаблон"),
        (Revision, "Редакция"), (TotalTime, "Время правки, мин"), (Created, "Создан"), (Modified, "Изменён"),
    };

    public static string Label(string key) => All.FirstOrDefault(f => f.Key == key).Label ?? key;

    public const string DateFormat = "yyyy-MM-dd HH:mm";

    static readonly string[] InputFormats =
    {
        "yyyy-MM-dd HH:mm", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd", "dd.MM.yyyy HH:mm", "dd.MM.yyyy HH:mm:ss", "dd.MM.yyyy",
    };

    /// <summary>Разбирает дату, введённую в местном времени, и возвращает её в UTC.</summary>
    public static bool TryParseLocal(string s, out DateTime utc)
    {
        utc = default;
        if (!DateTime.TryParseExact(s.Trim(), InputFormats, CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeLocal, out var local)) return false;
        utc = local.ToUniversalTime();
        return true;
    }

    /// <summary>Переводит дату из core.xml (UTC, W3CDTF) в местное время для показа.</summary>
    public static string FormatStored(string raw)
    {
        raw = raw.Trim();
        if (raw == "") return "";
        if (DateTime.TryParse(raw, CultureInfo.InvariantCulture,
                DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out var dt))
            return dt.ToLocalTime().ToString(DateFormat, CultureInfo.InvariantCulture);
        return raw;
    }

    public static string KindName(DocKind k) => k switch
    {
        DocKind.Word => "Word", DocKind.PowerPoint => "PowerPoint", _ => "Excel",
    };

    public static string ExpectedApplication(DocKind k) => k switch
    {
        DocKind.Word => "Microsoft Office Word",
        DocKind.PowerPoint => "Microsoft Office PowerPoint",
        _ => "Microsoft Excel",
    };

    public static DocKind? KindFromPath(string path) => Path.GetExtension(path).ToLowerInvariant() switch
    {
        ".docx" or ".docm" or ".dotx" => DocKind.Word,
        ".pptx" or ".pptm" or ".potx" => DocKind.PowerPoint,
        ".xlsx" or ".xlsm" or ".xltx" => DocKind.Excel,
        _ => null,
    };
}

/// <summary>Набор свойств документа: ключ из <see cref="Fields"/> → значение.</summary>
public sealed class DocProps : Dictionary<string, string>
{
    public DocProps() : base(StringComparer.Ordinal) { }

    /// <summary>Пользовательские свойства (docProps/custom.xml), только для показа.</summary>
    public List<(string Name, string Value)> Custom { get; } = new();

    public string Get(string key) => TryGetValue(key, out var v) ? v ?? "" : "";

    public DocProps Clone()
    {
        var d = new DocProps();
        foreach (var kv in this) d[kv.Key] = kv.Value;
        d.Custom.AddRange(Custom);
        return d;
    }
}

public sealed class Finding
{
    public Severity Severity { get; init; }
    public string Category { get; init; } = "";
    public string Message { get; init; } = "";
    /// <summary>Как исправляется: «Авто», «Пересохранить в Office», «Вручную».</summary>
    public string Fix { get; init; } = "";

    public static string SeverityName(Severity s) => s switch
    {
        Severity.High => "Высокий", Severity.Medium => "Средний", Severity.Low => "Низкий", _ => "Инфо",
    };
}

public sealed class InspectResult
{
    public string Path { get; init; } = "";
    public DocKind Kind { get; init; }
    public DocProps Props { get; set; } = new();
    public List<Finding> Findings { get; } = new();
    public HashSet<string> SuspiciousFields { get; } = new();
    public string? Error { get; set; }

    public Severity Score => Findings.Where(f => f.Severity >= Severity.Low)
        .Select(f => f.Severity).DefaultIfEmpty(Severity.Info).Max();
}
