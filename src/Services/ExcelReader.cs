
// ExcelReader.cs — adds ReadFirstSheet() and keeps ReadHeaders(); CSV + XLSX supported
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace AssetDataValidationTool.Services
{
    public static class ExcelReader
    {
        /// <summary>
        /// Returns the first row of the first sheet as headers.
        /// Supports .csv and .xlsx
        /// </summary>
        public static List<string> ReadHeaders(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return new List<string>();

            var ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".csv")
            {
                using var sr = new StreamReader(path);
                var first = sr.ReadLine();
                if (string.IsNullOrEmpty(first)) return new List<string>();
                return first.Split(',').Select(s => s.Trim()).ToList();
            }

            try
            {
                using var doc = SpreadsheetDocument.Open(path, false);
                var wbPart = doc.WorkbookPart;
                if (wbPart == null) return new List<string>();

                var sheet = wbPart.Workbook?.Sheets?.Elements<Sheet>()?.FirstOrDefault();
                if (sheet == null) return new List<string>();

                // Guard relationship id
                var relId = sheet.Id?.Value;
                if (string.IsNullOrEmpty(relId)) return new List<string>();

                var wsPart = wbPart.GetPartById(relId) as WorksheetPart;
                if (wsPart == null) return new List<string>();

                var firstRow = wsPart.Worksheet?.Descendants<Row>()?.FirstOrDefault();
                if (firstRow == null) return new List<string>();

                SharedStringTable? sst = wbPart.SharedStringTablePart?.SharedStringTable;
                var headers = new List<string>();

                foreach (var cell in firstRow.Elements<Cell>())
                {
                    string? text = null;
                    if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                    {
                        if (int.TryParse(cell.CellValue?.Text, out var sstIndex) && sst != null)
                        {
                            var item = sst.ElementAtOrDefault(sstIndex);
                            text = item?.InnerText;
                        }
                    }
                    else
                    {
                        text = cell.CellValue?.Text;
                    }
                    headers.Add(text?.Trim() ?? string.Empty);
                }

                // Trim trailing blanks
                for (int i = headers.Count - 1; i >= 0; i--)
                {
                    if (string.IsNullOrEmpty(headers[i])) headers.RemoveAt(i);
                    else break;
                }

                return headers;
            }
            catch
            {
                return new List<string>();
            }
        }

        /// <summary>
        /// Reads the first worksheet (or CSV) into headers + row dictionaries.
        /// </summary>
        public static (List<string> headers, List<Dictionary<string, string>> rows) ReadFirstSheet(string path)
        {
            var headers = ReadHeaders(path);
            var rows = new List<Dictionary<string, string>>();

            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path) || headers.Count == 0)
                return (headers, rows);

            var ext = Path.GetExtension(path).ToLowerInvariant();

            if (ext == ".csv")
            {
                using var sr = new StreamReader(path);
                // Skip header
                _ = sr.ReadLine();
                string? line;
                while ((line = sr.ReadLine()) != null)
                {
                    var parts = line.Split(',');
                    var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    for (int i = 0; i < headers.Count; i++)
                    {
                        var val = i < parts.Length ? parts[i] : string.Empty;
                        dict[headers[i]] = val;
                    }
                    rows.Add(dict);
                }
                return (headers, rows);
            }

            try
            {
                using var doc = SpreadsheetDocument.Open(path, false);
                var wbPart = doc.WorkbookPart;
                if (wbPart == null) return (headers, rows);

                var sheet = wbPart.Workbook?.Sheets?.Elements<Sheet>()?.FirstOrDefault();
                if (sheet == null) return (headers, rows);

                var relId = sheet.Id?.Value;
                if (string.IsNullOrEmpty(relId)) return (headers, rows);

                var wsPart = wbPart.GetPartById(relId) as WorksheetPart;
                if (wsPart == null) return (headers, rows);

                SharedStringTable? sst = wbPart.SharedStringTablePart?.SharedStringTable;

                // All subsequent rows after the first
                foreach (var row in wsPart.Worksheet?.Descendants<Row>()?.Skip(1) ?? Enumerable.Empty<Row>())
                {
                    var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    int colIndex = 0;
                    foreach (var cell in row.Elements<Cell>())
                    {
                        // Derive column index from cell reference (e.g., "C5" -> 2)
                        colIndex = GetColumnIndexFromReference(cell.CellReference?.Value) ?? colIndex;

                        string? text = null;
                        if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                        {
                            if (int.TryParse(cell.CellValue?.Text, out var sstIndex) && sst != null)
                                text = sst.ElementAtOrDefault(sstIndex)?.InnerText;
                        }
                        else
                        {
                            text = cell.CellValue?.Text;
                        }

                        if (colIndex >= 0 && colIndex < headers.Count)
                            dict[headers[colIndex]] = text ?? string.Empty;

                        colIndex++;
                    }

                    // Ensure all headers exist
                    foreach (var h in headers)
                        if (!dict.ContainsKey(h)) dict[h] = string.Empty;

                    rows.Add(dict);
                }

                return (headers, rows);
            }
            catch
            {
                return (headers, rows);
            }
        }

        private static int? GetColumnIndexFromReference(string? cellRef)
        {
            if (string.IsNullOrEmpty(cellRef)) return null;
            // Extract letters
            int idx = 0;
            foreach (char c in cellRef)
            {
                if (char.IsLetter(c))
                    idx = idx * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
                else break;
            }
            return idx > 0 ? idx - 1 : (int?)null;
        }
    }
}
