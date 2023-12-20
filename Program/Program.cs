using Libraries;

//Ввод адреса к файлу Excel
Console.WriteLine(
    "Укажите путь файла Excel (по умолчанию: \"D:\\3. Таблица ограничений параметров по подбору оборудования v1.3.xlsm\"):"
);
string? pathExcelFile = null;
var input = Console.ReadLine();

if (input == "")
{
    pathExcelFile = "D:\\3. Таблица ограничений параметров по подбору оборудования v1.3.xlsm";
}

//ввод номера последней строки в Excel
Console.WriteLine("Укажите номер последней строки в файле Excel (по умолчанию: последняя заполненная строка):");
input = Console.ReadLine();
var result = int.TryParse(input, out var lastRow);
if (!result && !string.IsNullOrEmpty(input))
{
    throw new ArgumentException(
        "Значение 'Номер последней строки' не является числом."
    );
}


//Ввод адреса к файлу Json
Console.WriteLine(
    "Укажите путь файла Json (по умолчанию: \"C:\\My ProjectCSharp\\OSUprogram2\\OSUprogram2\\Libraries\\Fans.json\"):"
);
string? pathJsonFile = null;
input = Console.ReadLine();

if (input == "")
{
    pathJsonFile = "C:\\My ProjectCSharp\\OSUprogram2\\OSUprogram2\\Libraries\\Fans.json";
}

var fansList = ExcelReader.Load(pathExcelFile, lastRow);

JsonLoader.Upload(fansList, pathJsonFile);

switch (JsonLoader.Response)
{
    case "y" :
        Console.WriteLine($"В файл *.json загружено {fansList?.Count} вентилятора.");
        break;
    default:
        Console.WriteLine("Файл *.json НЕ был обновлен.");
        break;
}



// fansList = JsonLoader.Download();
