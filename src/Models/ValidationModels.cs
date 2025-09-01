using System;
using System.Collections.Generic;

namespace AssetDataValidationTool.Models
{
    public class SourceTable
    {
        public string FilePath { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public List<string> Headers { get; set; } = new();
        public List<Dictionary<string, string>> Rows { get; set; } = new();
    }

    public class KeyPresence
    {
        public string Key { get; set; } = string.Empty;
        public Dictionary<string, bool> PresenceByFile { get; set; } = new();
    }

    public class Conflict
    {
        public string Key { get; set; } = string.Empty;
        public string Column { get; set; } = string.Empty;
        public Dictionary<string, string?> ValuesByFile { get; set; } = new();
    }

    public class ValidationResults
    {
        public string AssetClass { get; set; } = string.Empty;
        public string DataPoint { get; set; } = string.Empty;
        public List<SourceTable> Sources { get; set; } = new();
        public List<KeyPresence> Presence { get; set; } = new();
        public List<Conflict> Conflicts { get; set; } = new();
        public HashSet<string> MatchesAll { get; set; } = new();
        public Dictionary<string, HashSet<string>> MissingByFile { get; set; } = new();
        public string ReportFilePath { get; set; } = string.Empty;
        public string AuditLogPath { get; set; } = string.Empty;
        public string ZipPackagePath { get; set; } = string.Empty;
    }
}
