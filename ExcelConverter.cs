using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SpravkoBot_AsSapfir
{
public class ExcelConverter
{
    private const string DefaultPassword = "1234";
    private readonly string excelPath;

    public ExcelConverter(string excelPath)
    {
        this.excelPath = excelPath ?? throw new ArgumentNullException(nameof(excelPath));
        // Устанавливаем лицензию EPPlus (бесплатно для некоммерческого использования)
        ExcelPackage.License.SetNonCommercialOrganization("test");
    }

    public string ConvertToCsv()
    {
        if (string.IsNullOrEmpty(excelPath))
        {
            throw new ArgumentException("Путь к файлу не может быть пустым", nameof(excelPath));
        }

        if (!File.Exists(excelPath))
        {
            throw new FileNotFoundException("Excel-файл не найден!", excelPath);
        }

        string csvPath = Path.ChangeExtension(excelPath, ".csv");

        try
        {
            using (var package = OpenExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets[0]; // Первый лист (индекс 0 в EPPlus)

                if (worksheet == null)
                {
                    throw new Exception("В файле нет доступных листов");
                }

                PrepareWorksheet(worksheet);
                SaveAsCsv(worksheet, csvPath);
            }

            // TryDeleteExcelFile();
            return csvPath;
        }
        catch (Exception ex)
        {
            throw new Exception($"Ошибка при конвертации Excel в CSV: {ex.Message}", ex);
        }
    }

    private ExcelPackage OpenExcelPackage()
    {
        try
        {
            var fileInfo = new FileInfo(excelPath);
            return new ExcelPackage(fileInfo, DefaultPassword);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при открытии файла: {ex.Message}");
            throw;
        }
    }

    private void PrepareWorksheet(ExcelWorksheet worksheet)
    {
        try
        {
            // EPPlus автоматически обрабатывает защищенные листы при наличии пароля
            // Удаляем автофильтры, если они есть
            if (worksheet.AutoFilter != null)
            {
                worksheet.AutoFilter.ClearAll();
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Ошибка при подготовке листа: {ex.Message}", ex);
        }
    }

    private void SaveAsCsv(ExcelWorksheet worksheet, string csvPath)
    {
        try
        {
            // Создаем StringBuilder для построения CSV
            var csvBuilder = new StringBuilder();
            int rowCount = worksheet.Dimension?.Rows ?? 0;
            int colCount = worksheet.Dimension?.Columns ?? 0;

            if (rowCount == 0 || colCount == 0)
            {
                throw new Exception("Лист пустой или не удалось определить размеры");
            }

            // Проходим по всем строкам и столбцам
            for (int row = 1; row <= rowCount; row++)
            {
                var rowData = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? "";
                    // Экранируем кавычки и обрабатываем запятые
                    string escapedValue = $"\"{cellValue.Replace("\"", "\"\"")}\"";
                    rowData.Add(escapedValue);
                }
                csvBuilder.AppendLine(string.Join(",", rowData));
            }

            // Проверяем директорию и записываем файл
            string directory = Path.GetDirectoryName(csvPath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            if (File.Exists(csvPath))
            {
                File.Delete(csvPath);
            }

            File.WriteAllText(csvPath, csvBuilder.ToString(), Encoding.UTF8);
            Console.WriteLine($"Сохранено как CSV: {csvPath}");
        }
        catch (Exception ex)
        {
            throw new Exception($"Ошибка при сохранении CSV: {ex.Message}", ex);
        }
    }

    private void TryDeleteExcelFile()
    {
        try
        {
            if (File.Exists(excelPath))
            {
                File.Delete(excelPath);
                Console.WriteLine($"Исходный файл {excelPath} удален.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Не удалось удалить Excel-файл: {ex.Message}");
        }
    }
}
}