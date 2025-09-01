using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AssetDataValidationTool.Models;

namespace AssetDataValidationTool.Services
{
    internal static class Validator
    {
        public static ValidationResults Validate(string assetClass, string dataPoint, List<(string displayName, string filePath)> sources)
        {
            var result = new ValidationResults
            {
                AssetClass = assetClass,
                DataPoint = dataPoint
            };

            // Load sources
            foreach (var (displayName, filePath) in sources)
            {
                var (headers, rows) = ExcelReader.ReadFirstSheet(filePath);
                result.Sources.Add(new SourceTable
                {
                    FilePath = filePath,
                    DisplayName = displayName,
                    Headers = headers,
                    Rows = rows
                });
            }

            // Initialize primary key map if not provided
            if (result.PrimaryKeyBySource == null)
            {
                result.PrimaryKeyBySource = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            // Build key sets per source using the appropriate column for each source
            var keysByFile = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (var src in result.Sources)
            {
                var hs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Get the column name to use for this source
                string columnToUse;
                if (result.PrimaryKeyBySource.TryGetValue(src.DisplayName, out var mappedColumn))
                {
                    // Use the mapped column if it exists
                    columnToUse = mappedColumn;
                }
                else
                {
                    // Otherwise use dataPoint directly if it exists as a column
                    columnToUse = dataPoint;
                    // Store the mapping for future use
                    result.PrimaryKeyBySource[src.DisplayName] = dataPoint;
                }

                foreach (var row in src.Rows)
                {
                    if (!row.TryGetValue(columnToUse, out var key)) key = string.Empty;
                    if (!string.IsNullOrWhiteSpace(key)) hs.Add(key.Trim());
                }
                keysByFile[src.DisplayName] = hs;
            }

            // Presence table
            var allKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var set in keysByFile.Values) allKeys.UnionWith(set);

            foreach (var key in allKeys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
            {
                var presence = new KeyPresence { Key = key };
                foreach (var kvp in keysByFile)
                {
                    presence.PresenceByFile[kvp.Key] = kvp.Value.Contains(key);
                }
                result.Presence.Add(presence);
            }

            // Matches in ALL files
            result.MatchesAll = new HashSet<string>(
                allKeys.Where(k => keysByFile.All(kvp => kvp.Value.Contains(k))),
                StringComparer.OrdinalIgnoreCase);

            // Missing by file
            foreach (var kvp in keysByFile)
            {
                var missing = new HashSet<string>(allKeys.Where(k => !kvp.Value.Contains(k)), StringComparer.OrdinalIgnoreCase);
                result.MissingByFile[kvp.Key] = missing;
            }

            // Conflicts: For keys present in >= 2 files, compare common columns
            foreach (var key in allKeys)
            {
                var rowsByFile = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
                foreach (var src in result.Sources)
                {
                    // Get the mapped column name for this source
                    string columnToUse = result.PrimaryKeyBySource[src.DisplayName];

                    var row = src.Rows.FirstOrDefault(r => r.TryGetValue(columnToUse, out var v) &&
                                                          string.Equals(v?.Trim(), key, StringComparison.OrdinalIgnoreCase));
                    if (row != null) rowsByFile[src.DisplayName] = row;
                }
                if (rowsByFile.Count < 2) continue;

                var commonColumns = result.Sources.Select(s => s.Headers).Aggregate((a, b) => a.Intersect(b, StringComparer.OrdinalIgnoreCase).ToList());

                // Get all columns used as primary keys
                var pkColumns = new HashSet<string>(result.PrimaryKeyBySource.Values, StringComparer.OrdinalIgnoreCase);

                foreach (var col in commonColumns)
                {
                    // Skip this column if it's used as a primary key in any source
                    if (pkColumns.Contains(col))
                        continue;

                    string? firstVal = null;
                    bool differs = false;
                    var valuesByFile = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
                    foreach (var f in rowsByFile)
                    {
                        var val = f.Value.TryGetValue(col, out var v) ? (v ?? string.Empty).Trim() : string.Empty;
                        valuesByFile[f.Key] = val;
                        if (firstVal == null) firstVal = val;
                        else if (!string.Equals(firstVal, val, StringComparison.OrdinalIgnoreCase))
                        {
                            differs = true;
                        }
                    }
                    if (differs)
                    {
                        result.Conflicts.Add(new Conflict
                        {
                            Key = key,
                            Column = col,
                            ValuesByFile = valuesByFile
                        });
                    }
                }
            }

            return result;
        }
    }
}
