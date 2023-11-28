using System.Text.Json;

namespace Libraries;

public static class JsonLoader
{
    public static async Task UploadAsync(List<FanData>? fanCollection, string? pathJsonFile)
    {
        if (fanCollection == null)
        {
            throw new ArgumentNullException(nameof(fanCollection));
        }

        if (pathJsonFile == null)
        {
            throw new ArgumentNullException(nameof(pathJsonFile));
        }

        if (File.Exists(pathJsonFile))
        {
            Console.WriteLine("Файл уже существует. Хотите перезаписать его? (Y/N)");
            var response = Console.ReadLine();
            if (response?.ToUpper() != "Y")
            {
                return;
            }
        }

        var options = new JsonSerializerOptions
        {
            AllowTrailingCommas = true,
            WriteIndented = true
        };


        await using (var fileStream = new FileStream(pathJsonFile, FileMode.Create))
        {
            await JsonSerializer.SerializeAsync(fileStream, fanCollection, options);
        }

        Console.WriteLine(JsonSerializer.Serialize(fanCollection));
    }

    public static async Task<List<FanData>?> DownloadAsync(string? pathJsonFile)
    {
        List<FanData>? restoredFanData = null;

        var options = new JsonSerializerOptions
        {
            AllowTrailingCommas = true,
            WriteIndented = true
        };

        if (File.Exists(pathJsonFile))
        {
            try
            {
                await using Stream streamJson = File.OpenRead(pathJsonFile);
                restoredFanData = await JsonSerializer.DeserializeAsync<List<FanData>>(streamJson, options);
                Console.WriteLine(restoredFanData);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при десериализации данных: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine("Файл *.json не найден");
        }

        return restoredFanData;
    }
}
