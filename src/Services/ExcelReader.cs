using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace AssetDataValidationTool.Services
{
    internal static class ExcelReader
    {
        /// <summary>
        /// Reads the first worksheet (or CSV) and returns headers + rows.
        /// headerRowIndex is 1-based (1 = first row contains headers).
        /// </summary>
        public static (List<string> Headers, List<Dictionary<string, string>> Rows)
            ReadFirstSheet(string filePath, int headerRowIndex = 1)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return (new List<string>(), new List<Dictionary<string, string>>());

            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext == ".csv")
            {
                return ReadCsv(filePath);
            }

            if (!File.Exists(filePath))
                return (new List<string>(), new List<Dictionary<string, string>>());

            using (var doc = SpreadsheetDocument.Open(filePath, false))
            {
                var wbPart = doc.WorkbookPart;
                if (wbPart == null || wbPart.Workbook == null || wbPart.Workbook.Sheets == null)
                    return (new List<string>(), new List<Dictionary<string, string>>());

                var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault();
                if (sheet == null)
                    return (new List<string>(), new List<Dictionary<string, string>>());

                var wsPart = wbPart.GetPartById(sheet.Id!) as WorksheetPart;
                if (wsPart == null || wsPart.Worksheet == null)
                    return (new List<string>(), new List<Dictionary<string, string>>());

                var sst = wbPart.SharedStringTablePart?.SharedStringTable;
                var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
                if (sheetData == null)
                    return (new List<string>(), new List<Dictionary<string, string>>());

                var rows = sheetData.Elements<Row>().ToList();
                if (rows.Count == 0)
                    return (new List<string>(), new List<Dictionary<string, string>>());

                // Header row (1-based index)
                if (headerRowIndex < 1 || headerRowIndex > rows.Count)
                    headerRowIndex = 1;

                var headerRow = rows[headerRowIndex - 1];
                var headerCells = headerRow.Elements<Cell>().ToList();
                var headers = headerCells.Select(c => GetCellValue(c, sst)).ToList();

                // Normalize empty/missing headers to generated names (H1, H2, ...)
                for (int i = 0; i < headers.Count; i++)
                {
                    if (string.IsNullOrWhiteSpace(headers[i]))
                        headers[i] = $"H{i + 1}";
                }

                var dataRows = new List<Dictionary<string, string>>();
                for (int i = headerRowIndex; i < rows.Count; i++)
                {
                    var r = rows[i];
                    var cells = r.Elements<Cell>().ToList();
                    var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                    // Align by position (first N cells → first N headers)
                    for (int j = 0; j < headers.Count; j++)
                    {
                        string val = string.Empty;
                        if (j < cells.Count)
                        {
                            val = GetCellValue(cells[j], sst) ?? string.Empty;
                        }
                        dict[headers[j]] = val.Trim();
                    }

                    // Keep only non-empty rows
                    if (dict.Values.Any(v => !string.IsNullOrWhiteSpace(v)))
                        dataRows.Add(dict);
                }

                return (headers, dataRows);
            }
        }

        /// <summary>
        /// Simple CSV reader with support for quoted fields, commas inside quotes, and escaped quotes.
        /// </summary>
        private static (List<string> Headers, List<Dictionary<string, string>> Rows) ReadCsv(string filePath)
        {
            var lines = File.ReadAllLines(filePath);
            if (lines.Length == 0)
                return (new List<string>(), new List<Dictionary<string, string>>());

            var headers = SplitCsvLine(lines[0]);
            // Normalize empty headers
            for (int i = 0; i < headers.Count; i++)
            {
                if (string.IsNullOrWhiteSpace(headers[i]))
                    headers[i] = $"H{i + 1}";
            }

            var rows = new List<Dictionary<string, string>>();
            for (int i = 1; i < lines.Length; i++)
            {
                var cols = SplitCsvLine(lines[i]);
                var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                for (int j = 0; j < headers.Count; j++)
                {
                    var v = j < cols.Count ? cols[j] : string.Empty;
                    dict[headers[j]] = (v ?? string.Empty).Trim();
                }

                if (dict.Values.Any(v => !string.IsNullOrWhiteSpace(v)))
                    rows.Add(dict);
            }
            return (headers, rows);
        }

        /// <summary>
        /// Splits a single CSV line respecting quotes and commas inside quotes.
        /// Supports escaping quotes by doubling them ("").
        /// </summary>
        private static List<string> SplitCsvLine(string line)
        {
            var result = new List<string>();
            if (line == null)
            {
                result.Add(string.Empty);
                return result;
            }

            bool inQuotes = false;
            var cur = new StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char ch = line[i];

                if (ch == '"')
                {
                    if (inQuotes)
                    {
                        // If next char is also a quote, this is an escaped quote.
                        if (i + 1 < line.Length && line[i + 1] == '"')
                        {
                            cur.Append('"');
                            i++; // skip the second quote
                        }
                        else
                        {
                            // End quote
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        // Begin quote only if at field start or after comma
                        inQuotes = true;
                    }
                }
                else if (ch == ',' && !inQuotes)
                {
                    result.Add(cur.ToString());
                    cur.Clear();
                }
                else
                {
                    cur.Append(ch);
                }
            }

            result.Add(cur.ToString());
            return result;
        }

        /// <summary>
        /// Returns the text value for an OpenXML Cell, considering SharedString, InlineString, Boolean, and Number.
        /// </summary>
        private static string GetCellValue(Cell cell, SharedStringTable? sst)
        {
            if (cell == null)
                return string.Empty;

            // Inline string?
            if (cell.DataType != null && cell.DataType.Value == CellValues.InlineString && cell.InlineString != null)
            {
                return cell.InlineString.Text?.Text ?? string.Empty;
            }

            // Raw value text (often for numbers/booleans or when DataType is null)
            var raw = cell.CellValue?.Text ?? string.Empty;

            if (cell.DataType == null)
            {
                // No explicit data type: treat as raw string/number as-is
                return raw;
            }

            var dt = cell.DataType.Value;

            // Avoid switch expressions to keep analyzers happy (CS9135)
            if (dt == CellValues.SharedString)
            {
                if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var idx)
                    && sst != null
                    && idx >= 0
                    && idx < sst.ChildElements.Count)
                {
                    var ssi = sst.ElementAt(idx);
                    // Shared string items may contain runs; InnerText flattens them.
                    return ssi?.InnerText ?? string.Empty;
                }
                return string.Empty;
            }
            else if (dt == CellValues.Boolean)
            {
                // OpenXML uses "1" or "0" for booleans
                return (raw == "1").ToString();
            }
            else
            {
                // String, Number, Date (stored as number), Error, etc. — return raw
                return raw;
            }
        }
    }
}
