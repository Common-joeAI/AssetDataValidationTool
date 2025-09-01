using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
                    var indexBySource = new Dictionary<string, Dictionary<string, Dictionary<string,string>>>(StringComparer.OrdinalIgnoreCase);
                    foreach (var src in results.Sources)
                    {
                        var pkCol = pkMap.TryGetValue(src.DisplayName, out var pk) ? pk : results.DataPoint;
                        var idx = new Dictionary<string, Dictionary<string,string>>(StringComparer.OrdinalIgnoreCase);
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
                    var indexBySource = new Dictionary<string, Dictionary<string, Dictionary<string,string>>>(StringComparer.OrdinalIgnoreCase);
                    foreach (var src in results.Sources)
                    {
                        var pkCol = pkMap.TryGetValue(src.DisplayName, out var pk) ? pk : results.DataPoint;
                        var idx = new Dictionary<string, Dictionary<string,string>>(StringComparer.OrdinalIgnoreCase);
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
                    var mapBySource = new Dictionary<string, Dictionary<string,string>>(StringComparer.OrdinalIgnoreCase);
                    foreach (var other in results.Sources.Where(s => !s.DisplayName.Equals(baseline.DisplayName, StringComparison.OrdinalIgnoreCase)))
                    {
                        var baseIdx = indexBySource[baseline.DisplayName];
                        var otherIdx = indexBySource[other.DisplayName];
                        var commonKeys = baseIdx.Keys.Intersect(otherIdx.Keys, StringComparer.OrdinalIgnoreCase).ToList();
                        var map = new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
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
                    var indexBySource = new Dictionary<string, Dictionary<string, Dictionary<string,string>>>(StringComparer.OrdinalIgnoreCase);
                    foreach (var src in results.Sources)
                    {
                        var pkCol = pkMap.TryGetValue(src.DisplayName, out var pk) ? pk : results.DataPoint;
                        var idx = new Dictionary<string, Dictionary<string,string>>(StringComparer.OrdinalIgnoreCase);
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
                    var mapBySource = new Dictionary<string, Dictionary<string,string>>(StringComparer.OrdinalIgnoreCase);
                    foreach (var other in results.Sources.Where(s => !s.DisplayName.Equals(baseline.DisplayName, StringComparison.OrdinalIgnoreCase)))
                    {
                        var baseIdx = indexBySource[baseline.DisplayName];
                        var otherIdx = indexBySource[other.DisplayName];
                        var commonKeys = baseIdx.Keys.Intersect(otherIdx.Keys, StringComparer.OrdinalIgnoreCase).ToList();
                        var map = new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
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
                    var mismatchCountBySource = new Dictionary<string,int>(StringComparer.OrdinalIgnoreCase);
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

                    // Charts sheet with a basic bar chart of mismatches by source
                    try
                    {
                        var chartsWs = wbPart.AddNewPart<WorksheetPart>();
                        chartsWs.Worksheet = new Worksheet();
                        var sheetCharts = new Sheet { Id = wbPart.GetIdOfPart(chartsWs), SheetId = sheetId++, Name = "Charts" };
                        sheets.Append(sheetCharts);

                        var drawingsPart = chartsWs.AddNewPart<DrawingsPart>();
                        chartsWs.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = chartsWs.GetIdOfPart(drawingsPart) });
                        chartsWs.Worksheet.Save();

                        var chartPart = drawingsPart.AddNewPart<DocumentFormat.OpenXml.Packaging.ChartPart>();
                        chartPart.ChartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace();
                        var chartSpace = chartPart.ChartSpace;
                        chartSpace.Append(new DocumentFormat.OpenXml.Drawing.Charts.EditingLanguage() { Val = "en-US" });

                        var chart = chartSpace.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Chart());
                        var plotArea = chart.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.PlotArea());
                        plotArea.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Layout());
                        var barChart = plotArea.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.BarChart(
                            new DocumentFormat.OpenXml.Drawing.Charts.BarDirection() { Val = DocumentFormat.OpenXml.Drawing.Charts.BarDirectionValues.Column },
                            new DocumentFormat.OpenXml.Drawing.Charts.BarGrouping() { Val = DocumentFormat.OpenXml.Drawing.Charts.BarGroupingValues.Clustered }
                        ));

                        var series = new DocumentFormat.OpenXml.Drawing.Charts.BarChartSeries(
                            new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = (uint)0 },
                            new DocumentFormat.OpenXml.Drawing.Charts.Order() { Val = (uint)0 },
                            new DocumentFormat.OpenXml.Drawing.Charts.SeriesText(new DocumentFormat.OpenXml.Drawing.Charts.NumericValue() { Text = "Delta Cells by Source" })
                        );

                        // Categories (sources) and Values (counts) from DeltasSummary!A2:A{N} and B2:B{N}
                        int n = mismatchCountBySource.Count;
                        if (n < 1) n = 1;
                        string catRef = $"DeltasSummary!$A$2:$A${n + 1}";
                        string valRef = $"DeltasSummary!$B$2:$B${n + 1}";

                        var cat = new DocumentFormat.OpenXml.Drawing.Charts.CategoryAxisData(
                            new DocumentFormat.OpenXml.Drawing.Charts.StringReference()
                            {
                                Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula(catRef),
                                StringCache = new DocumentFormat.OpenXml.Drawing.Charts.StringCache()
                            });
                        var val = new DocumentFormat.OpenXml.Drawing.Charts.Values(
                            new DocumentFormat.OpenXml.Drawing.Charts.NumberReference()
                            {
                                Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula(valRef),
                                NumberingCache = new DocumentFormat.OpenXml.Drawing.Charts.NumberingCache()
                            });

                        series.Append(cat);
                        series.Append(val);
                        barChart.Append(series);
                        barChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisId() { Val = 48650112u });
                        barChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisId() { Val = 48672768u });

                        // Category axis
                        var catAx = plotArea.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis(
                            new DocumentFormat.OpenXml.Drawing.Charts.AxisId() { Val = 48650112u },
                            new DocumentFormat.OpenXml.Drawing.Charts.Scaling(new DocumentFormat.OpenXml.Drawing.Charts.Orientation() { Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax }),
                            new DocumentFormat.OpenXml.Drawing.Charts.Delete() { Val = false },
                            new DocumentFormat.OpenXml.Drawing.Charts.AxisPosition() { Val = DocumentFormat.OpenXml.Drawing.Charts.AxisPositionValues.Bottom },
                            new DocumentFormat.OpenXml.Drawing.Charts.CrossingAxis() { Val = 48672768u },
                            new DocumentFormat.OpenXml.Drawing.Charts.Crosses() { Val = DocumentFormat.OpenXml.Drawing.Charts.CrossesValues.AutoZero }
                        ));

                        // Value axis
                        var valAx = plotArea.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.ValueAxis(
                            new DocumentFormat.OpenXml.Drawing.Charts.AxisId() { Val = 48672768u },
                            new DocumentFormat.OpenXml.Drawing.Charts.Scaling(new DocumentFormat.OpenXml.Drawing.Charts.Orientation() { Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax }),
                            new DocumentFormat.OpenXml.Drawing.Charts.Delete() { Val = false },
                            new DocumentFormat.OpenXml.Drawing.Charts.AxisPosition() { Val = DocumentFormat.OpenXml.Drawing.Charts.AxisPositionValues.Left },
                            new DocumentFormat.OpenXml.Drawing.Charts.CrossingAxis() { Val = 48650112u },
                            new DocumentFormat.OpenXml.Drawing.Charts.Crosses() { Val = DocumentFormat.OpenXml.Drawing.Charts.CrossesValues.AutoZero }
                        ));

                        chart.Append(new DocumentFormat.OpenXml.Drawing.Charts.PlotVisibleOnly() { Val = true });
                        chartSpace.Save();

                        // Place the chart in the sheet via anchor
                        var wsDr = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
                        drawingsPart.WorksheetDrawing = wsDr;

                        var twoCellAnchor = new DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor();
                        twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(
                            new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("1"),
                            new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"),
                            new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("1"),
                            new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0")
                        ));
                        twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(
                            new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("16"),
                            new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"),
                            new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("24"),
                            new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0")
                        ));

                        var graphicFrame = new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame();
                        graphicFrame.Macro = string.Empty;
                        graphicFrame.NonVisualGraphicFrameProperties = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
                            new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() { Id = 2u, Name = "DeltaChart" },
                            new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()
                        );
                        graphicFrame.Transform = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Transform(
                            new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                            new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 0L, Cy = 0L }
                        );
                        var graphic = new DocumentFormat.OpenXml.Drawing.Graphic();
                        graphic.GraphicData = new DocumentFormat.OpenXml.Drawing.GraphicData(
                            new DocumentFormat.OpenXml.Drawing.Charts.ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) }
                        );
                        graphic.GraphicData.Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart";

                        graphicFrame.Append(graphic);
                        twoCellAnchor.Append(graphicFrame);
                        twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData());
                        wsDr.Append(twoCellAnchor);
                        drawingsPart.WorksheetDrawing.Save();
                    }
                    catch { /* chart creation is best-effort; ignore if not supported */}
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
                var c = new Cell { DataType = CellValues.String, CellValue = new CellValue(v ?? string.Empty) };
                r.Append(c);
            }
            sd.Append(r);
        }

        
        private static WorksheetPart? GetWorksheetPartByName(WorkbookPart workbookPart, string sheetName)
        {
            var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>();
            if (sheets == null) return null;
            var theSheet = sheets.FirstOrDefault(s => s.Name != null && s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            if (theSheet == null || theSheet.Id == null) return null;
            return (WorksheetPart)workbookPart.GetPartById(theSheet.Id);
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
