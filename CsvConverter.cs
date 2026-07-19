using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace CSV.Convertor
{
    public static class CsvConverter
    {
        private const char DEFAULT_DELIMITER = ';';

        private static readonly Dictionary<string, Type> ColumnTypes = new Dictionary<string, Type>
        {
            { "Art.Id", typeof(string) },
            { "EAN code", typeof(string) },
            { "Artist", typeof(string) },
            { "Title", typeof(string) },
            { "Additional Info", typeof(string) },
            { "Units", typeof(double) },  // ИЗМЕНЕНО: теперь double, т.к. могут быть дробные значения
            { "Media", typeof(string) },
            { "List Price", typeof(double) },
            { "Customer Price", typeof(double) },
            { "Currency", typeof(string) },
            { "Origin", typeof(string) },
            { "Availability", typeof(string) },
            { "StockState", typeof(string) },
            { "Label", typeof(string) },
            { "Is Bertus Label", typeof(bool) },
            { "Catalogue Number", typeof(string) },
            { "Genre", typeof(string) },
            { "Release Date", typeof(DateTime) },
            { "PLC Status", typeof(string) },
            { "Date Added", typeof(DateTime) }
        };

        public static void Convert(string csvPath, string xlsxPath, char? delimiter = null, IProgress<int>? progress = null)
        {
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            CloseExcelProcessesForFile(xlsxPath);

            Encoding encoding = DetectEncoding(csvPath);

            char actualDelimiter = delimiter ?? DetectDelimiter(csvPath, encoding);

            string[] lines = File.ReadAllLines(csvPath, encoding);
            int totalLines = lines.Length;

            if (totalLines == 0)
                throw new Exception("CSV файл пуст!");

            using (var package = new OfficeOpenXml.ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                string[] headers = ParseCsvLine(lines[0], actualDelimiter);

                // Заголовки
                for (int j = 0; j < headers.Length; j++)
                {
                    string header = CleanValue(headers[j]);
                    var cell = worksheet.Cells[1, j + 1];
                    cell.Value = header;

                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }

                // Данные
                for (int i = 1; i < totalLines; i++)
                {
                    string line = lines[i];
                    string[] values = ParseCsvLine(line, actualDelimiter);

                    for (int j = 0; j < values.Length && j < headers.Length; j++)
                    {
                        string cleanValue = CleanValue(values[j]);
                        string header = CleanValue(headers[j]);

                        object convertedValue = ConvertValue(cleanValue, header);

                        var cell = worksheet.Cells[i + 1, j + 1];
                        cell.Value = convertedValue;

                        // Форматирование в зависимости от типа
                        if (convertedValue is double || convertedValue is int)
                        {
                            // Для колонки Units: показываем десятичные знаки только если они есть
                            if (header == "Units")
                            {
                                // #,##0.## покажет 1.5 → 1.5, 2 → 2, 3.75 → 3.75
                                cell.Style.Numberformat.Format = "#,##0.##";
                            }
                            else
                            {
                                cell.Style.Numberformat.Format = "#,##0.00";
                            }
                        }
                        else if (convertedValue is DateTime)
                        {
                            cell.Style.Numberformat.Format = "yyyy-MM-dd";
                        }

                        cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }

                    if (progress != null && i % Math.Max(1, (totalLines - 1) / 100) == 0)
                    {
                        int percent = (int)((i) * 100.0 / (totalLines - 1));
                        progress.Report(Math.Min(percent, 100));
                    }
                }

                progress?.Report(100);

                // Закрепляем первую строку И первый столбец
                worksheet.View.FreezePanes(2, 2);

                // Устанавливаем активную ячейку в B2
                worksheet.View.SelectedRange = "B2";

                // Автофильтр и автоширина
                if (worksheet.Dimension != null)
                {
                    worksheet.Cells[1, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns].AutoFilter = true;
                    worksheet.Cells[1, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns].AutoFitColumns();
                }

                SaveWithRetry(package, xlsxPath, 3);
            }
        }

        /// <summary>
        /// Сохраняет файл с повторными попытками
        /// </summary>
        private static void SaveWithRetry(OfficeOpenXml.ExcelPackage package, string filePath, int maxRetries = 3)
        {
            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    if (File.Exists(filePath))
                    {
                        try
                        {
                            File.Delete(filePath);
                        }
                        catch
                        {
                            // Пробуем перезаписать
                        }
                    }

                    package.SaveAs(new FileInfo(filePath));
                    return;
                }
                catch (Exception ex) when (attempt < maxRetries)
                {
                    System.Threading.Thread.Sleep(500);
                    CloseExcelProcessesForFile(filePath);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Не удалось сохранить файл после {maxRetries} попыток: {ex.Message}", ex);
                }
            }
        }

        /// <summary>
        /// Закрывает все процессы Excel
        /// </summary>
        private static void CloseExcelProcessesForFile(string filePath)
        {
            try
            {
                var processes = Process.GetProcessesByName("EXCEL");
                foreach (var process in processes)
                {
                    try
                    {
                        process.Kill();
                        process.WaitForExit(1000);
                    }
                    catch
                    {
                        // Игнорируем ошибки
                    }
                }
            }
            catch
            {
                // Игнорируем ошибки
            }
        }

        /// <summary>
        /// Преобразует значение в соответствии с типом колонки
        /// </summary>
        private static object ConvertValue(string value, string header)
        {
            if (string.IsNullOrEmpty(value))
            {
                if (ColumnTypes.TryGetValue(header, out Type? type))
                {
                    if (type == typeof(int) || type == typeof(double) || type == typeof(bool))
                        return null;
                }
                return value;
            }

            if (!ColumnTypes.TryGetValue(header, out Type? columnType))
                return value;

            try
            {
                if (columnType == typeof(int))
                {
                    if (int.TryParse(value, out int result))
                        return result;
                    return null;
                }
                else if (columnType == typeof(double))
                {
                    string normalizedValue = value.Replace(",", ".");
                    if (double.TryParse(normalizedValue, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out double result))
                        return result;
                    return null;
                }
                else if (columnType == typeof(DateTime))
                {
                    if (DateTime.TryParse(value, out DateTime result))
                        return result;
                    return null;
                }
                else if (columnType == typeof(bool))
                {
                    if (bool.TryParse(value, out bool result))
                        return result;
                    if (value.Equals("True", StringComparison.OrdinalIgnoreCase))
                        return true;
                    if (value.Equals("False", StringComparison.OrdinalIgnoreCase))
                        return false;
                    return null;
                }
            }
            catch
            {
                return value;
            }

            return value;
        }

        /// <summary>
        /// Парсит строку CSV с учётом кавычек
        /// </summary>
        private static string[] ParseCsvLine(string line, char delimiter)
        {
            var result = new List<string>();
            bool inQuotes = false;
            int start = 0;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        i++;
                        continue;
                    }
                    inQuotes = !inQuotes;
                }
                else if (!inQuotes && c == delimiter)
                {
                    result.Add(line.Substring(start, i - start));
                    start = i + 1;
                }
            }

            result.Add(line.Substring(start));
            return result.ToArray();
        }

        /// <summary>
        /// Очищает значение от префикса '=' и лишних кавычек
        /// </summary>
        private static string CleanValue(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            if (value.StartsWith("="))
                value = value.Substring(1);

            if (value.StartsWith("\"") && value.EndsWith("\""))
                value = value.Substring(1, value.Length - 2);

            value = value.Replace("\"\"", "\"");

            return value;
        }

        /// <summary>
        /// Определяет кодировку файла по BOM
        /// </summary>
        private static Encoding DetectEncoding(string filePath)
        {
            byte[] bytes = File.ReadAllBytes(filePath);

            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
                return Encoding.UTF8;

            if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
                return Encoding.BigEndianUnicode;

            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
                return Encoding.Unicode;

            try
            {
                return Encoding.UTF8;
            }
            catch
            {
                return Encoding.GetEncoding(1251);
            }
        }

        /// <summary>
        /// Определяет разделитель по первой строке
        /// </summary>
        private static char DetectDelimiter(string filePath, Encoding encoding)
        {
            string firstLine = File.ReadLines(filePath, encoding).FirstOrDefault();

            if (string.IsNullOrEmpty(firstLine))
                return DEFAULT_DELIMITER;

            int commaCount = firstLine.Count(c => c == ',');
            int semicolonCount = firstLine.Count(c => c == ';');
            int tabCount = firstLine.Count(c => c == '\t');

            if (semicolonCount >= commaCount && semicolonCount >= tabCount)
                return ';';

            if (tabCount >= commaCount && tabCount >= semicolonCount)
                return '\t';

            return ',';
        }
    }
}