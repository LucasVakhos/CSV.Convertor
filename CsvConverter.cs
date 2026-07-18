using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;

namespace CSV.Convertor
{
    public static class CsvConverter
    {
        private const char DEFAULT_DELIMITER = ';';

        public static void Convert(string csvPath, string xlsxPath, IProgress<int>? progress = null)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            Encoding encoding = DetectEncoding(csvPath);
            char delimiter = DetectDelimiter(csvPath, encoding);

            string[] lines = File.ReadAllLines(csvPath, encoding);
            int totalLines = lines.Length;

            if (totalLines == 0)
                throw new Exception("CSV файл пуст!");

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                for (int i = 0; i < totalLines; i++)
                {
                    string line = lines[i];
                    string[] values = ParseCsvLine(line, delimiter);

                    for (int j = 0; j < values.Length; j++)
                    {
                        string cleanValue = CleanValue(values[j]);
                        worksheet.Cells[i + 1, j + 1].Value = cleanValue;
                    }

                    // Обновляем прогресс (с проверкой на null)
                    if (progress != null && i % Math.Max(1, totalLines / 100) == 0)
                    {
                        int percent = (int)((i + 1) * 100.0 / totalLines);
                        progress.Report(Math.Min(percent, 100));
                    }
                }

                // Отправляем 100% (если progress не null)
                progress?.Report(100);

                if (worksheet.Dimension != null)
                {
                    worksheet.Cells[1, 1, totalLines, worksheet.Dimension.Columns].AutoFitColumns();
                }

                package.SaveAs(new FileInfo(xlsxPath));
            }
        }

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