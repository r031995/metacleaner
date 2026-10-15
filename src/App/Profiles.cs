using System.Text.Encodings.Web;
using System.Text.Json;

namespace MetaCleaner.App;

/// <summary>Сохранённый набор значений массовой замены и опций.</summary>
public sealed class Profile
{
    public string Name { get; set; } = "";
    public Dictionary<string, string> Values { get; set; } = new();
    public Dictionary<string, bool> Checks { get; set; } = new();
}

public static class ProfileStore
{
    static string FilePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MetaCleaner", "profiles.json");

    static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static List<Profile> Load()
    {
        try
        {
            if (File.Exists(FilePath))
                return JsonSerializer.Deserialize<List<Profile>>(File.ReadAllText(FilePath), Json) ?? new();
        }
        catch { /* испорченный файл — начинаем с пустого списка */ }
        return new();
    }

    public static void Save(List<Profile> list)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
        File.WriteAllText(FilePath, JsonSerializer.Serialize(list, Json));
    }
}
