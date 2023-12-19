using OfficeOpenXml;

namespace Libraries;

public static class ExcelReader
{
    public static List<FanData>? Load(string? pathExcelFile, int lastRow = 0)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Открытие существующего файла
        var package = new ExcelPackage(
            new FileInfo(pathExcelFile ?? throw new InvalidOperationException("Файл не найден."))
        );

        // Получение листа по индексу или имени
        // ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(); // по первый или последний открытый лист
        var worksheet = package.Workbook.Worksheets[0]; // по индексу

        int rowCount;

        if (lastRow > 0)
        {
            rowCount = lastRow;
        }
        else
        {
            rowCount = worksheet.Dimension.End.Row;
        }

        var fanCollection = new FanCollection
        {
            Fans = new List<FanData>()
        };

        for (int row = 7, colNoise = 62; row <= rowCount; row++)
        {
            var fanData = new FanData
            {
                Version = worksheet.Cells[row, 1].Value.ToString()!,
                Size = worksheet.Cells[row, 2].Value.ToString()!,
                ImpellerRotationDirection = worksheet.Cells[row, 7].Value!.ToString()!,
                NominalImpellerRotationSpeed = Convert.ToInt32(worksheet.Cells[row, 10].Value!),
                ImpellerRotationSpeed = Convert.ToInt32(worksheet.Cells[row, 11].Value!),
                MinVolumeFlow = Convert.ToInt32(worksheet.Cells[row, 15].Value!),
                MaxVolumeFlow = Convert.ToInt32(worksheet.Cells[row, 16].Value!),
                TotalPressureCoefficients = new PolynomialType
                {
                    SixthCoefficient = Convert.ToDouble(worksheet.Cells[row, 35].Value!),
                    FifthCoefficient = Convert.ToDouble(worksheet.Cells[row, 36].Value!),
                    FourthCoefficient = Convert.ToDouble(worksheet.Cells[row, 37].Value!),
                    ThirdCoefficient = Convert.ToDouble(worksheet.Cells[row, 38].Value!),
                    SecondCoefficient = Convert.ToDouble(worksheet.Cells[row, 39].Value!),
                    FirstCoefficient = Convert.ToDouble(worksheet.Cells[row, 40].Value!),
                    ZeroCoefficient = Convert.ToDouble(worksheet.Cells[row, 41].Value!)
                },
                PowerCoefficients = new PolynomialType
                {
                    SixthCoefficient = Convert.ToDouble(worksheet.Cells[row, 47].Value!),
                    FifthCoefficient = Convert.ToDouble(worksheet.Cells[row, 48].Value!),
                    FourthCoefficient = Convert.ToDouble(worksheet.Cells[row, 49].Value!),
                    ThirdCoefficient = Convert.ToDouble(worksheet.Cells[row, 50].Value!),
                    SecondCoefficient = Convert.ToDouble(worksheet.Cells[row, 51].Value!),
                    FirstCoefficient = Convert.ToDouble(worksheet.Cells[row, 52].Value!),
                    ZeroCoefficient = Convert.ToDouble(worksheet.Cells[row, 53].Value!)
                },
                InletCrossSection = Convert.ToDouble(worksheet.Cells[row, 56].Value!),
                NominalPower = Convert.ToDouble(worksheet.Cells[row, 9].Value!),
                OctaveNoiseCoefficients63 = new PolynomialType
                {
                    SixthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise].Value!),
                    FifthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 1].Value!),
                    FourthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 2].Value!),
                    ThirdCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 3].Value!),
                    SecondCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 4].Value!),
                    FirstCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 5].Value!),
                    ZeroCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 6].Value!)
                },
                OctaveNoiseCoefficients125 = new PolynomialType
                {
                    SixthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 7].Value!),
                    FifthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 8].Value!),
                    FourthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 9].Value!),
                    ThirdCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 10].Value!),
                    SecondCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 11].Value!),
                    FirstCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 12].Value!),
                    ZeroCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 13].Value!)
                },
                OctaveNoiseCoefficients250 = new PolynomialType
                {
                    SixthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 14].Value!),
                    FifthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 15].Value!),
                    FourthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 16].Value!),
                    ThirdCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 17].Value!),
                    SecondCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 18].Value!),
                    FirstCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 19].Value!),
                    ZeroCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 20].Value!)
                },
                OctaveNoiseCoefficients500 = new PolynomialType
                {
                    SixthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 21].Value!),
                    FifthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 22].Value!),
                    FourthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 23].Value!),
                    ThirdCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 24].Value!),
                    SecondCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 25].Value!),
                    FirstCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 26].Value!),
                    ZeroCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 27].Value!)
                },
                OctaveNoiseCoefficients1000 = new PolynomialType
                {
                    SixthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 28].Value!),
                    FifthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 29].Value!),
                    FourthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 30].Value!),
                    ThirdCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 31].Value!),
                    SecondCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 32].Value!),
                    FirstCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 33].Value!),
                    ZeroCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 34].Value!)
                },
                OctaveNoiseCoefficients2000 = new PolynomialType
                {
                    SixthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 35].Value!),
                    FifthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 36].Value!),
                    FourthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 37].Value!),
                    ThirdCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 38].Value!),
                    SecondCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 39].Value!),
                    FirstCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 40].Value!),
                    ZeroCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 41].Value!)
                },
                OctaveNoiseCoefficients4000 = new PolynomialType
                {
                    SixthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 42].Value!),
                    FifthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 43].Value!),
                    FourthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 44].Value!),
                    ThirdCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 45].Value!),
                    SecondCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 46].Value!),
                    FirstCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 47].Value!),
                    ZeroCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 48].Value!)
                },
                OctaveNoiseCoefficients8000 = new PolynomialType
                {
                    SixthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 49].Value!),
                    FifthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 50].Value!),
                    FourthCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 51].Value!),
                    ThirdCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 52].Value!),
                    SecondCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 53].Value!),
                    FirstCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 54].Value!),
                    ZeroCoefficient = Convert.ToDouble(worksheet.Cells[row, colNoise + 55].Value!)
                }
            };
            fanCollection.Fans.Add(fanData);
        }

        // Закрытие файла
        package.Dispose();

        return fanCollection.Fans;
    }
}
