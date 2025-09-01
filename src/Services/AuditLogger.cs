using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AssetDataValidationTool.Services
{
    internal static class AuditLogger
    {
        public static string WriteAuditLog(string outputFolder, string assetClass, string dataPoint, IEnumerable<(string displayName, string filePath)> sources)
        {
            Directory.CreateDirectory(outputFolder);
            var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var path = Path.Combine(outputFolder, $"audit_{ts}.log");
            var lines = new List<string>
            {
                $"Timestamp: {DateTime.Now:O}",
                $"Username: {Environment.UserName}",
                $"Machine: {Environment.MachineName}",
                $"AssetClass: {assetClass}",
                $"DataPoint: {dataPoint}",
                "Sources:"
            };
            lines.AddRange(sources.Select(s => $"  - {s.displayName}: {s.filePath}"));
            File.WriteAllLines(path, lines);
            return path;
        }
    }
}
