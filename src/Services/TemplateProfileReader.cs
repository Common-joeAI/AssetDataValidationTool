using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace AssetDataValidationTool.Services
{
    /// <summary>
    /// Reads a "Data Validation - <AssetClass>.xlsx" workbook and attempts to
    /// 1) infer the Asset Class from the filename
    /// 2) extract the expected source labels from a worksheet named "Process"
    ///    by locating a column whose header contains "source", "input", or "file".
    /// </summary>
    internal static class TemplateProfileReader
    {
        public static (string? AssetClass, List<string> SourceLabels) ExtractFromValidationWorkbook(string filePath)
        {
            string? assetClass = TryGetAssetClassFromFilename(filePath);
            var labels = new List<string>();

            if (!File.Exists(filePath)) return (assetClass, labels);

            try
            {
                using var doc = SpreadsheetDocument.Open(filePath, false);
                var wbPart = doc.WorkbookPart!;
                var sst = wbPart.SharedStringTablePart?.SharedStringTable;

                // Find a sheet named "Process" (case-insensitive). Fallback to first sheet.
                Sheet? sheet = wbPart.Workbook.Sheets!.Elements<Sheet>()
                    .FirstOrDefault(s => string.Equals(s.Name?.Value, "Process", StringComparison.OrdinalIgnoreCase));
                if (sheet == null)
                    sheet = wbPart.Workbook.Sheets!.Elements<Sheet>().FirstOrDefault();
                if (sheet == null) return (assetClass, labels);

                var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id!);
                var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
                if (sheetData == null) return (assetClass, labels);

                var allRows = sheetData.Elements<Row>().ToList();
                if (allRows.Count == 0) return (assetClass, labels);

                // Build a simple table: each row is the sequential string values of its cells.
                var table = new List<List<string>>();
                foreach (var r in allRows)
                {
                    var cells = r.Elements<Cell>().ToList();
                    var list = new List<string>();
                    foreach (var c in cells)
                    {
                        list.Add(GetCellValue(c, sst));
                    }
                    table.Add(list);
                }

                // Heuristic: find a header cell (in the first ~15 rows) whose text contains "source", "input", or "file".
                int headerRowIndex = -1;
                int headerColIndex = -1;
                for (int i = 0; i < Math.Min(15, table.Count); i++)
                {
                    var row = table[i];
                    for (int j = 0; j < row.Count; j++)
                    {
                        var text = row[j]?.Trim() ?? string.Empty;
                        if (text.Length == 0) continue;
                        var lower = text.ToLowerInvariant();
                        if (lower.Contains("source") || lower.Contains("input") || lower.Contains("file"))
                        {
                            headerRowIndex = i;
                            headerColIndex = j;
                            break;
                        }
                    }
                    if (headerRowIndex >= 0) break;
                }

                if (headerRowIndex >= 0 && headerColIndex >= 0)
                {
                    // Collect non-empty values under that column until a blank stretch (10 blanks) or end.
                    int blanks = 0;
                    for (int r = headerRowIndex + 1; r < table.Count; r++)
                    {
                        var row = table[r];
                        string val = (headerColIndex < row.Count ? row[headerColIndex] : string.Empty)?.Trim() ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(val))
                        {
                            blanks++;
                            if (blanks >= 10) break;
                            continue;
                        }
                        blanks = 0;

                        // Avoid header repeats and notes
                        if (val.Length > 0 && val.Length <= 200)
                        {
                            labels.Add(val);
                        }
                    }
                }

                // Deduplicate while preserving order
                labels = labels
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Select(s => s.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            catch
            {
                // swallow and return what we found
            }

            return (assetClass, labels);
        }

        private static string? TryGetAssetClassFromFilename(string filePath)
        {
            var name = Path.GetFileNameWithoutExtension(filePath);
            // Examples: "Data Validation - Computers", "Data Validation - Windows Server"
            var m = Regex.Match(name, @"Data\s*Validation\s*-\s*(.+)$", RegexOptions.IgnoreCase);
            if (m.Success)
            {
                return m.Groups[1].Value.Trim();
            }
            return null;
        }

        private static string GetCellValue(Cell cell, SharedStringTable? sst)
        {
            var value = cell.CellValue?.Text ?? string.Empty;
            if (cell.DataType == null) return value;

            var dt = cell.DataType.Value;
            if (dt == CellValues.SharedString)
            {
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var idx)
                    && sst != null
                    && idx >= 0
                    && idx < sst.ChildElements.Count)
                {
                    return sst.ElementAt(idx)?.InnerText ?? string.Empty;
                }
                return string.Empty;
            }
            else if (dt == CellValues.Boolean)
            {
                return (value == "1").ToString();
            }
            else if (dt == CellValues.InlineString && cell.InlineString != null)
            {
                return cell.InlineString.Text?.Text ?? string.Empty;
            }
            else
            {
                return value;
            }
        }
    }
}
