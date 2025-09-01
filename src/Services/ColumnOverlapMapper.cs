// src/Services/ColumnOverlapMapper.cs
using System;
using System.Collections.Generic;
using System.Linq;

namespace AssetDataValidationTool.Services
{
    internal static class ColumnOverlapMapper
    {
        public static List<(string SourceA, string ColumnA, string SourceB, string ColumnB, int Overlap, double Ratio)>
            SuggestMappings(Dictionary<string, List<Dictionary<string, string>>> tables, int minOverlap = 5, double minRatio = 0.05)
        {
            // NOTE: keep the tuple element names here so .Overlap/.Ratio work
            var result = new List<(string SourceA, string ColumnA, string SourceB, string ColumnB, int Overlap, double Ratio)>();
            var sources = tables.Keys.ToList();

            for (int i = 0; i < sources.Count; i++)
                for (int j = i + 1; j < sources.Count; j++)
                {
                    var a = sources[i];
                    var b = sources[j];
                    var rowsA = tables[a];
                    var rowsB = tables[b];
                    if (rowsA.Count == 0 || rowsB.Count == 0) continue;

                    var headersA = rowsA.SelectMany(r => r.Keys).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                    var headersB = rowsB.SelectMany(r => r.Keys).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

                    foreach (var ha in headersA)
                    {
                        var setA = new HashSet<string>(rowsA.Select(r => ValueNormalizer.NormalizeValue(ha, r.TryGetValue(ha, out var v) ? v : ""))
                                                             .Where(s => !string.IsNullOrWhiteSpace(s)));
                        if (setA.Count == 0) continue;

                        foreach (var hb in headersB)
                        {
                            var setB = new HashSet<string>(rowsB.Select(r => ValueNormalizer.NormalizeValue(hb, r.TryGetValue(hb, out var v) ? v : ""))
                                                                 .Where(s => !string.IsNullOrWhiteSpace(s)));
                            if (setB.Count == 0) continue;

                            int overlap = setA.Intersect(setB).Count();
                            double denom = Math.Max(setA.Count, setB.Count);
                            double ratio = denom > 0 ? (double)overlap / denom : 0.0;

                            if (overlap >= minOverlap && ratio >= minRatio)
                                result.Add((a, ha, b, hb, overlap, Math.Round(ratio, 4)));
                        }
                    }
                }

            return result
                .OrderByDescending(t => t.Overlap)
                .ThenByDescending(t => t.Ratio)
                .ToList();
        }
    }
}
