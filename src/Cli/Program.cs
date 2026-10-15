using System.Text;
using MetaCleaner.Core;

// Консольная версия: пакетная проверка и чистка без окна.
//   metaclean inspect <файл|папка> [-r]
//   metaclean clean   <файл|папка> [-r] [--author ФИО] [--modified-by ФИО] [--comment текст]
//                     [--company X] [--reviewer ФИО] [--in-place] [--no-backup] [--office-app]
//                     [--plausible] [--resave] [--fix-text] [--dash endash|hyphen]

Console.OutputEncoding = Encoding.UTF8;
if (args.Length < 2 || args[0] is not ("inspect" or "clean"))
{
    Console.WriteLine("Использование:\n  metaclean inspect <файл|папка> [-r]\n  metaclean clean <файл|папка> [-r] [--author ФИО] " +
                      "[--modified-by ФИО] [--comment текст] [--company X] [--reviewer ФИО] [--in-place] [--no-backup] " +
                      "[--office-app] [--plausible] [--resave] [--fix-text] [--dash endash|hyphen]");
    return 1;
}

string cmd = args[0], target = args[1];
var opts = new Dictionary<string, string?>();
for (int i = 2; i < args.Length; i++)
{
    var a = args[i];
    if (a.StartsWith("--") && i + 1 < args.Length && !args[i + 1].StartsWith("-") &&
        a is "--author" or "--modified-by" or "--comment" or "--company" or "--reviewer" or "--dash")
        opts[a] = args[++i];
    else opts[a] = null;
}
bool recursive = opts.ContainsKey("-r");

List<(string File, string Root)> files;
if (Directory.Exists(target))
    files = Scanner.FindFiles(target, recursive).Select(f => (f, target)).ToList();
else if (File.Exists(target))
    files = new() { (target, Path.GetDirectoryName(Path.GetFullPath(target))!) };
else { Console.WriteLine($"Не найдено: {target}"); return 1; }

void Print(InspectResult r)
{
    Console.WriteLine($"\n=== {r.Path}  [{Fields.KindName(r.Kind)}]  оценка: {Finding.SeverityName(r.Score)}");
    foreach (var (k, label) in Fields.All)
        Console.WriteLine($"   {label,-20} {r.Props.Get(k)}");
    foreach (var c in r.Props.Custom) Console.WriteLine($"   [custom] {c.Name} = {c.Value}");
    foreach (var f in r.Findings.OrderByDescending(f => f.Severity))
        Console.WriteLine($"   {Finding.SeverityName(f.Severity),-8} {f.Category,-12} {f.Message}  ({f.Fix})");
}

if (cmd == "inspect")
{
    foreach (var (f, _) in files) Print(Inspector.InspectFile(f));
    return 0;
}

var so = new SaveOptions
{
    InPlace = opts.ContainsKey("--in-place"),
    MakeBackup = !opts.ContainsKey("--no-backup"),
    Reviewer = opts.GetValueOrDefault("--reviewer") ?? "",
    OfficeResave = opts.ContainsKey("--resave"),
};
if (opts.ContainsKey("--fix-text"))
    so.Text = new TextFixOptions
    {
        FixDashes = true, RenameFiles = true, RemoveMarkdown = true, RemoveChatPhrases = true,
        FillPlaceholders = true, CleanAltText = true,
        DashReplacement = opts.GetValueOrDefault("--dash") == "hyphen" ? "-" : "–",
    };
var rnd = new Random();
OfficeSession? office = so.OfficeResave ? new OfficeSession() : null;
int ok = 0, fail = 0;
try
{
    foreach (var (f, root) in files)
    {
        var before = Inspector.InspectFile(f);
        if (before.Error != null) { Console.WriteLine($"ПРОПУСК {f}: {before.Error}"); fail++; continue; }
        var p = before.Props.Clone();
        void SetIf(string opt, string key) { if (opts.TryGetValue(opt, out var v) && v != null) p[key] = v; }
        SetIf("--author", Fields.Creator);
        SetIf("--modified-by", Fields.LastModifiedBy);
        if (opts.ContainsKey("--author") && !opts.ContainsKey("--modified-by")) p[Fields.LastModifiedBy] = p[Fields.Creator];
        p[Fields.Description] = opts.GetValueOrDefault("--comment") ?? "";
        SetIf("--company", Fields.Company);
        if (opts.ContainsKey("--office-app"))
        {
            p[Fields.Application] = Fields.ExpectedApplication(before.Kind);
            p[Fields.AppVersion] = "16.0000";
            if (before.Kind == DocKind.Word && p.Get(Fields.Template) == "") p[Fields.Template] = "Normal.dotm";
        }
        if (opts.ContainsKey("--plausible"))
        {
            if (before.Kind != DocKind.Excel)
            {
                p[Fields.Revision] = rnd.Next(3, 26).ToString();
                p[Fields.TotalTime] = rnd.Next(25, 241).ToString();
            }
        }
        try
        {
            var r = Cleaner.SaveFile(f, root, p, so, office);
            Console.WriteLine($"OK  {f} -> {r.OutputPath}");
            foreach (var a in r.Actions) Console.WriteLine($"    - {a}");
            Print(Inspector.InspectFile(r.OutputPath));
            ok++;
        }
        catch (Exception ex) { Console.WriteLine($"ОШИБКА {f}: {ex.Message}"); fail++; }
    }
}
finally { office?.Dispose(); }
Console.WriteLine($"\nГотово: {ok} успешно, {fail} с ошибками.");
return fail == 0 ? 0 : 2;
