using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using AssetDataValidationTool.Models;

namespace AssetDataValidationTool.Services
{
    internal static class ReportGenerator
    {
        public static string GenerateExcelReport(ValidationResults results, string outputFolder)
        {
            Directory.CreateDirectory(outputFolder);
            var ts = DateTime.Now.ToString("yyyyMMdd_HHmm");
            var fileName = $"ValidationReport_{results.AssetClass}_{ts}.xlsx";
            var path = Path.Combine(outputFolder, fileName);

            using (var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                var wbPart = doc.AddWorkbookPart();
                wbPart.Workbook = new Workbook();
                var sheets = new Sheets();
                wbPart.Workbook.AppendChild(sheets);

                uint sheetId = 1;

                // Summary
                {
                    var wsPart = wbPart.AddNewPart<WorksheetPart>();
                    wsPart.Worksheet = new Worksheet(new SheetData());
                    var sheet = new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = sheetId++, Name = "Summary" };
                    sheets.Append(sheet);

                    var sd = wsPart.Worksheet.GetFirstChild<SheetData>()!;
                    AppendRow(sd, new string[] { "Asset Class", results.AssetClass });
                    AppendRow(sd, new string[] { "Data Point", results.DataPoint });
                    AppendRow(sd, new string[] { "Sources", string.Join(" | ", results.Sources.Select(s => $"{s.DisplayName}:{Path.GetFileName(s.FilePath)}")) });
                    AppendRow(sd, Array.Empty<string>());
                    AppendRow(sd, new string[] { "Total Keys", results.Presence.Count.ToString() });
                    AppendRow(sd, new string[] { "Matched In All Files", results.MatchesAll.Count.ToString() });
                    AppendRow(sd, new string[] { "Conflict Count", results.Conflicts.Count.ToString() });
                }

                // KeyPresence
                {
                    var wsPart = wbPart.AddNewPart<WorksheetPart>();
                    wsPart.Worksheet = new Worksheet(new SheetData());
                    var sheet = new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = sheetId++, Name = "KeyPresence" };
                    sheets.Append(sheet);

                    var sd = wsPart.Worksheet.GetFirstChild<SheetData>()!;
                    var header = new List<string> { results.DataPoint };
                    header.AddRange(results.Sources.Select(s => s.DisplayName));
                    AppendRow(sd, header);

                    foreach (var p in results.Presence)
                    {
                        var row = new List<string> { p.Key };
                        foreach (var src in results.Sources)
                        {
                            var present = p.PresenceByFile.TryGetValue(src.DisplayName, out var pr) && pr;
                            row.Add(present ? "Yes" : "No");
                        }
                        AppendRow(sd, row);
                    }
                }

                // Conflicts
                {
                    var wsPart = wbPart.AddNewPart<WorksheetPart>();
                    wsPart.Worksheet = new Worksheet(new SheetData());
                    var sheet = new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = sheetId++, Name = "Conflicts" };
                    sheets.Append(sheet);
                    var sd = wsPart.Worksheet.GetFirstChild<SheetData>()!;

                    var header = new List<string> { results.DataPoint, "Column" };
                    header.AddRange(results.Sources.Select(s => s.DisplayName));
                    AppendRow(sd, header);

                    foreach (var c in results.Conflicts
                        .OrderBy(c => c.Key, StringComparer.OrdinalIgnoreCase)
                        .ThenBy(c => c.Column, StringComparer.OrdinalIgnoreCase))
                    {
                        var row = new List<string> { c.Key, c.Column };
                        foreach (var src in results.Sources)
                        {
                            c.ValuesByFile.TryGetValue(src.DisplayName, out var v);
                            row.Add(v ?? string.Empty);
                        }
                        AppendRow(sd, row);
                    }
                }

                // MatchesAll
                {
                    var wsPart = wbPart.AddNewPart<WorksheetPart>();
                    wsPart.Worksheet = new Worksheet(new SheetData());
                    var sheet = new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = sheetId++, Name = "MatchesAll" };
                    sheets.Append(sheet);
                    var sd = wsPart.Worksheet.GetFirstChild<SheetData>()!;

                    AppendRow(sd, new string[] { results.DataPoint });
                    foreach (var k in results.MatchesAll.OrderBy(s => s, StringComparer.OrdinalIgnoreCase))
                    {
                        AppendRow(sd, new string[] { k });
                    }
                }

                // MissingByFile
                {
                    var wsPart = wbPart.AddNewPart<WorksheetPart>();
                    wsPart.Worksheet = new Worksheet(new SheetData());
                    var sheet = new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = sheetId++, Name = "MissingByFile" };
                    sheets.Append(sheet);
                    var sd = wsPart.Worksheet.GetFirstChild<SheetData>()!;

                    AppendRow(sd, new string[] { results.DataPoint, "MissingFrom" });
                    foreach (var kvp in results.MissingByFile)
                    {
                        foreach (var key in kvp.Value.OrderBy(s => s, StringComparer.OrdinalIgnoreCase))
                        {
                            AppendRow(sd, new string[] { key, kvp.Key });
                        }
                    }
                }


                // FieldMapping (auto-inferred column mappings to Baseline)
                {
                    var baseline = results.Sources.FirstOrDefault(s => s.DisplayName.Equals("Baseline", StringComparison.OrdinalIgnoreCase)) ?? results.Sources.First();
                    var wsPart = wbPart.AddNewPart<WorksheetPart>();
                    wsPart.Worksheet = new Worksheet(new SheetData());
                    var sheet = new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = sheetId++, Name = "FieldMapping" };
                    sheets.Append(sheet);
                    var sd = wsPart.Worksheet.GetFirstChild<SheetData>()!;

                    AppendRow(sd, new string[] { "Baseline", "OtherSource", "BaselineColumn", "MappedColumn", "MatchScore" });

                    var pkMap = results.PrimaryKeyBySource ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    // Build index by PK for each source
                    var indexBySource = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>(StringComparer.OrdinalIgnoreCase);
                    foreach (var src in results.Sources)
                    {
                        var pkCol = pkMap.TryGetValue(src.DisplayName, out var pk) ? pk : results.DataPoint;
                        var idx = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
                        foreach (var row in src.Rows)
                        {
                            if (!row.TryGetValue(pkCol, out var key)) continue;
                            key = (key ?? string.Empty).Trim();
                            if (string.IsNullOrEmpty(key)) continue;
                            if (!idx.ContainsKey(key)) idx[key] = row;
                        }
                        indexBySource[src.DisplayName] = idx;
                    }

                    foreach (var other in results.Sources.Where(s => !s.DisplayName.Equals(baseline.DisplayName, StringComparison.OrdinalIgnoreCase)))
                    {
                        // Determine best column mapping by value overlap
                        var baseIdx = indexBySource[baseline.DisplayName];
                        var otherIdx = indexBySource[other.DisplayName];
                        var commonKeys = baseIdx.Keys.Intersect(otherIdx.Keys, StringComparer.OrdinalIgnoreCase).ToList();
                        if (commonKeys.Count == 0) continue;

                        var usedOtherCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                        foreach (var bCol in baseline.Headers)
                        {
                            if (string.IsNullOrWhiteSpace(bCol)) continue;
                            double bestScore = 0.0;
                            string? bestOther = null;
                            foreach (var oCol in other.Headers)
                            {
                                if (usedOtherCols.Contains(oCol)) continue;
                                int same = 0, total = 0;
                                foreach (var k in commonKeys)
                                {
                                    var bRow = baseIdx[k];
                                    var oRow = otherIdx[k];
                                    var bv = bRow.ContainsKey(bCol) ? (bRow[bCol] ?? "").Trim() : "";
                                    var ov = oRow.ContainsKey(oCol) ? (oRow[oCol] ?? "").Trim() : "";
                                    if (string.IsNullOrEmpty(bv) && string.IsNullOrEmpty(ov)) continue;
                                    total++;
                                    if (string.Equals(bv, ov, StringComparison.OrdinalIgnoreCase)) same++;
                                }
                                if (total > 0)
                                {
                                    var score = (double)same / total;
                                    if (score > bestScore)
                                    {
                                        bestScore = score;
                                        bestOther = oCol;
                                    }
                                }
                            }
                            if (bestOther != null && bestScore >= 0.6) // threshold
                            {
                                usedOtherCols.Add(bestOther);
                                AppendRow(sd, new string[] { baseline.DisplayName, other.DisplayName, bCol, bestOther, bestScore.ToString("0.00") });
                            }
                        }
                    }
                }

                // Deltas (mismatched values across sources based on inferred mapping)
                {
                    var baseline = results.Sources.FirstOrDefault(s => s.DisplayName.Equals("Baseline", StringComparison.OrdinalIgnoreCase)) ?? results.Sources.First();
                    var wsPart = wbPart.AddNewPart<WorksheetPart>();
                    wsPart.Worksheet = new Worksheet(new SheetData());
                    var sheet = new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = sheetId++, Name = "Deltas" };
                    sheets.Append(sheet);
                    var sd = wsPart.Worksheet.GetFirstChild<SheetData>()!;

                    // Header row
                    var header = new List<string> { "Key", "Column", baseline.DisplayName };
                    header.AddRange(results.Sources.Where(s => !s.DisplayName.Equals(baseline.DisplayName, StringComparison.OrdinalIgnoreCase)).Select(s => s.DisplayName));
                    AppendRow(sd, header.ToArray());

                    var pkMap = results.PrimaryKeyBySource ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    // Build index by PK for each source
                    var indexBySource = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>(StringComparer.OrdinalIgnoreCase);
                    foreach (var src in results.Sources)
                    {
                        var pkCol = pkMap.TryGetValue(src.DisplayName, out var pk) ? pk : results.DataPoint;
                        var idx = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
                        foreach (var row in src.Rows)
                        {
                            if (!row.TryGetValue(pkCol, out var key)) continue;
                            key = (key ?? string.Empty).Trim();
                            if (string.IsNullOrEmpty(key)) continue;
                            if (!idx.ContainsKey(key)) idx[key] = row;
                        }
                        indexBySource[src.DisplayName] = idx;
                    }

                    // Build column mapping baseline -> other for each other source using same heuristic
                    var mapBySource = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
                    foreach (var other in results.Sources.Where(s => !s.DisplayName.Equals(baseline.DisplayName, StringComparison.OrdinalIgnoreCase)))
                    {
                        var baseIdx = indexBySource[baseline.DisplayName];
                        var otherIdx = indexBySource[other.DisplayName];
                        var commonKeys = baseIdx.Keys.Intersect(otherIdx.Keys, StringComparer.OrdinalIgnoreCase).ToList();
                        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var bCol in baseline.Headers)
                        {
                            double bestScore = 0.0;
                            string? bestOther = null;
                            foreach (var oCol in other.Headers)
                            {
                                int same = 0, total = 0;
                                foreach (var k in commonKeys)
                                {
                                    var bRow = baseIdx[k];
                                    var oRow = otherIdx[k];
                                    var bv = bRow.ContainsKey(bCol) ? (bRow[bCol] ?? "").Trim() : "";
                                    var ov = oRow.ContainsKey(oCol) ? (oRow[oCol] ?? "").Trim() : "";
                                    if (string.IsNullOrEmpty(bv) && string.IsNullOrEmpty(ov)) continue;
                                    total++;
                                    if (string.Equals(bv, ov, StringComparison.OrdinalIgnoreCase)) same++;
                                }
                                if (total > 0)
                                {
                                    var score = (double)same / total;
                                    if (score > bestScore)
                                    {
                                        bestScore = score;
                                        bestOther = oCol;
                                    }
                                }
                            }
                            if (bestOther != null && bestScore >= 0.6) map[bCol] = bestOther;
                        }
                        mapBySource[other.DisplayName] = map;
                    }

                    // Gather all keys
                    var allKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var kv in indexBySource) foreach (var k in kv.Value.Keys) allKeys.Add(k);

                    foreach (var key in allKeys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
                    {
                        var baseIdx = indexBySource[baseline.DisplayName];
                        var baseHas = baseIdx.ContainsKey(key);
                        foreach (var bCol in baseline.Headers)
                        {
                            var baseVal = baseHas && baseIdx[key].ContainsKey(bCol) ? (baseIdx[key][bCol] ?? "") : "";
                            var values = new List<string> { key, bCol, baseVal };
                            bool anyMismatch = false;
                            foreach (var other in results.Sources.Where(s => !s.DisplayName.Equals(baseline.DisplayName, StringComparison.OrdinalIgnoreCase)))
                            {
                                var otherIdx = indexBySource[other.DisplayName];
                                string v = "";
                                if (otherIdx.ContainsKey(key))
                                {
                                    var map = mapBySource[other.DisplayName];
                                    if (map.TryGetValue(bCol, out var otherCol) && otherIdx[key].ContainsKey(otherCol))
                                    {
                                        v = otherIdx[key][otherCol] ?? "";
                                    }
                                }
                                values.Add(v);
                                if (!string.Equals(v.Trim(), (baseVal ?? "").Trim(), StringComparison.OrdinalIgnoreCase))
                                {
                                    // Consider mismatch only if at least one has a non-empty value
                                    if (!(string.IsNullOrWhiteSpace(v) && string.IsNullOrWhiteSpace(baseVal))) anyMismatch = true;
                                }
                            }
                            if (anyMismatch)
                            {
                                AppendRow(sd, values.ToArray());
                            }
                        }
                    }
                }


                // === SUMMARY BUMP ADDED ===
                {
                    // Compute counts for Summary + DeltasSummary
                    var pkMap = results.PrimaryKeyBySource ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    var baseline = results.Sources.FirstOrDefault(s => s.DisplayName.Equals("Baseline", StringComparison.OrdinalIgnoreCase)) ?? results.Sources.First();
                    // Rebuild indices
                    var indexBySource = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>(StringComparer.OrdinalIgnoreCase);
                    foreach (var src in results.Sources)
                    {
                        var pkCol = pkMap.TryGetValue(src.DisplayName, out var pk) ? pk : results.DataPoint;
                        var idx = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
                        foreach (var row in src.Rows)
                        {
                            if (!row.TryGetValue(pkCol, out var key)) continue;
                            key = (key ?? string.Empty).Trim();
                            if (string.IsNullOrEmpty(key)) continue;
                            if (!idx.ContainsKey(key)) idx[key] = row;
                        }
                        indexBySource[src.DisplayName] = idx;
                    }
                    // Build mappings baseline->other (same heuristic as FieldMapping)
                    var mapBySource = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
                    foreach (var other in results.Sources.Where(s => !s.DisplayName.Equals(baseline.DisplayName, StringComparison.OrdinalIgnoreCase)))
                    {
                        var baseIdx = indexBySource[baseline.DisplayName];
                        var otherIdx = indexBySource[other.DisplayName];
                        var commonKeys = baseIdx.Keys.Intersect(otherIdx.Keys, StringComparer.OrdinalIgnoreCase).ToList();
                        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var bCol in baseline.Headers)
                        {
                            double bestScore = 0.0; string? bestOther = null;
                            foreach (var oCol in other.Headers)
                            {
                                int same = 0, total = 0;
                                foreach (var k in commonKeys)
                                {
                                    var bRow = baseIdx[k];
                                    var oRow = otherIdx[k];
                                    var bv = bRow.ContainsKey(bCol) ? (bRow[bCol] ?? "").Trim() : "";
                                    var ov = oRow.ContainsKey(oCol) ? (oRow[oCol] ?? "").Trim() : "";
                                    if (string.IsNullOrEmpty(bv) && string.IsNullOrEmpty(ov)) continue;
                                    total++; if (string.Equals(bv, ov, StringComparison.OrdinalIgnoreCase)) same++;
                                }
                                if (total > 0)
                                {
                                    var score = (double)same / total;
                                    if (score > bestScore) { bestScore = score; bestOther = oCol; }
                                }
                            }
                            if (bestOther != null && bestScore >= 0.6) map[bCol] = bestOther;
                        }
                        mapBySource[other.DisplayName] = map;
                    }

                    // Count mismatched cells per other source and total
                    var mismatchCountBySource = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    foreach (var s in results.Sources.Where(s => !s.DisplayName.Equals(baseline.DisplayName, StringComparison.OrdinalIgnoreCase)))
                        mismatchCountBySource[s.DisplayName] = 0;
                    int totalMismatchCells = 0;
                    // Iterate keys
                    var allKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var kv in indexBySource) foreach (var k in kv.Value.Keys) allKeys.Add(k);
                    foreach (var key in allKeys)
                    {
                        var baseIdx = indexBySource[baseline.DisplayName];
                        var baseHas = baseIdx.ContainsKey(key);
                        foreach (var bCol in baseline.Headers)
                        {
                            var baseVal = baseHas && baseIdx[key].ContainsKey(bCol) ? (baseIdx[key][bCol] ?? "") : "";
                            foreach (var other in results.Sources.Where(s => !s.DisplayName.Equals(baseline.DisplayName, StringComparison.OrdinalIgnoreCase)))
                            {
                                var otherIdx = indexBySource[other.DisplayName];
                                string v = "";
                                if (otherIdx.ContainsKey(key))
                                {
                                    var map = mapBySource[other.DisplayName];
                                    if (map.TryGetValue(bCol, out var otherCol) && otherIdx[key].ContainsKey(otherCol))
                                    { v = otherIdx[key][otherCol] ?? ""; }
                                }
                                // mismatch?
                                if (!string.Equals(v.Trim(), (baseVal ?? "").Trim(), StringComparison.OrdinalIgnoreCase))
                                {
                                    if (!(string.IsNullOrWhiteSpace(v) && string.IsNullOrWhiteSpace(baseVal)))
                                    {
                                        mismatchCountBySource[other.DisplayName] = mismatchCountBySource.GetValueOrDefault(other.DisplayName) + 1;
                                        totalMismatchCells++;
                                    }
                                }
                            }
                        }
                    }

                    // Count mapped field pairs total
                    int mappedPairs = 0;
                    foreach (var kv in mapBySource) mappedPairs += kv.Value.Count;

                    // Append to Summary sheet
                    var summaryWs = GetWorksheetPartByName(wbPart, "Summary");
                    if (summaryWs != null)
                    {
                        var sd = summaryWs.Worksheet.GetFirstChild<SheetData>()!;
                        AppendRow(sd, Array.Empty<string>());
                        AppendRow(sd, new string[] { "Metrics" });
                        AppendRow(sd, new string[] { "Mapped Field Pairs", mappedPairs.ToString() });
                        AppendRow(sd, new string[] { "Total Delta Cells", totalMismatchCells.ToString() });
                        foreach (var kv in mismatchCountBySource)
                        {
                            AppendRow(sd, new string[] { $"Delta Cells - {kv.Key}", kv.Value.ToString() });
                        }
                    }

                    // DeltasSummary sheet (for chart source)
                    {
                        var wsPart2 = wbPart.AddNewPart<WorksheetPart>();
                        wsPart2.Worksheet = new Worksheet(new SheetData());
                        var sheet2 = new Sheet { Id = wbPart.GetIdOfPart(wsPart2), SheetId = sheetId++, Name = "DeltasSummary" };
                        sheets.Append(sheet2);
                        var sd2 = wsPart2.Worksheet.GetFirstChild<SheetData>()!;
                        AppendRow(sd2, new string[] { "Source", "MismatchedCells" });
                        foreach (var kv in mismatchCountBySource)
                        {
                            AppendRow(sd2, new string[] { kv.Key, kv.Value.ToString() });
                        }
                    }

                    // Skip chart creation - it's causing XML corruption
                    // We'll just create a simple sheet instead
                    var chartsWs = wbPart.AddNewPart<WorksheetPart>();
                    chartsWs.Worksheet = new Worksheet(new SheetData());
                    var sheetCharts = new Sheet { Id = wbPart.GetIdOfPart(chartsWs), SheetId = sheetId++, Name = "Charts" };
                    sheets.Append(sheetCharts);
                    var chartsSd = chartsWs.Worksheet.GetFirstChild<SheetData>()!;
                    AppendRow(chartsSd, new string[] { "Charts are disabled to prevent Excel corruption" });
                    AppendRow(chartsSd, new string[] { "Please refer to DeltasSummary sheet for data" });
                }

                // Source previews
                foreach (var src in results.Sources)
                {
                    var wsPart = wbPart.AddNewPart<WorksheetPart>();
                    wsPart.Worksheet = new Worksheet(new SheetData());
                    var safeName = MakeSheetName($"Source_{src.DisplayName}");
                    var sheet = new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = sheetId++, Name = safeName };
                    sheets.Append(sheet);
                    var sd = wsPart.Worksheet.GetFirstChild<SheetData>()!;

                    AppendRow(sd, src.Headers.ToArray());
                    foreach (var row in src.Rows.Take(100))
                    {
                        var values = src.Headers.Select(h => row.TryGetValue(h, out var v) ? v : string.Empty).ToArray();
                        AppendRow(sd, values);
                    }
                }

                wbPart.Workbook.Save();
            }

            results.ReportFilePath = path;
            return path;
        }

        private static void AppendRow(SheetData sd, IEnumerable<string> values)
        {
            var r = new Row();
            foreach (var v in values)
            {
                // Sanitize the value to ensure it's valid XML
                string safeValue = SanitizeForXml(v ?? string.Empty);
                var c = new Cell { DataType = CellValues.String, CellValue = new CellValue(safeValue) };
                r.Append(c);
            }
            sd.Append(r);
        }

        private static string SanitizeForXml(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            // Remove invalid XML characters
            var validXml = new System.Text.StringBuilder();
            foreach (char c in text)
            {
                if (IsValidXmlChar(c))
                    validXml.Append(c);
            }
            return validXml.ToString();
        }

        private static bool IsValidXmlChar(char c)
        {
            // XML 1.0 valid character ranges
            return c == 0x9 || c == 0xA || c == 0xD ||
                  (c >= 0x20 && c <= 0xD7FF) ||
                  (c >= 0xE000 && c <= 0xFFFD);
        }

        private static WorksheetPart? GetWorksheetPartByName(WorkbookPart workbookPart, string sheetName)
        {
            var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>();
            if (sheets == null) return null;

            // Safer null checking
            var theSheet = sheets.FirstOrDefault(s => s.Name?.Value != null &&
                                                s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

            // Ensure both sheet and ID exist
            if (theSheet?.Id?.Value == null) return null;

            return (WorksheetPart)workbookPart.GetPartById(theSheet.Id.Value);
        }

        private static string MakeSheetName(string name)
        {
            // Excel sheet name restrictions: no \ / * [ ] : ? and <=31 chars
            var invalid = new char[] { '\\', '/', '*', '[', ']', ':', '?' };
            foreach (var c in invalid) name = name.Replace(c.ToString(), "-");
            return name.Length > 31 ? name.Substring(0, 31) : name;
        }
    }
}
