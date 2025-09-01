using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace AssetDataValidationTool.Services
{
    internal static class Packager
    {
        public static string CreateZip(string assetClass, string reportFile, IEnumerable<string> sourceFiles, string auditLogPath, string outputFolder)
        {
            Directory.CreateDirectory(outputFolder);
            var date = DateTime.Now.ToString("yyyyMMdd");
            var username = Environment.UserName;
            var zipName = $"{assetClass}-{date}-{username}.zip".Replace(' ', '_');
            var zipPath = Path.Combine(outputFolder, zipName);

            if (File.Exists(zipPath)) File.Delete(zipPath);

            using (var archive = ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                archive.CreateEntryFromFile(reportFile, Path.GetFileName(reportFile));
                foreach (var f in sourceFiles.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    if (File.Exists(f))
                    {
                        archive.CreateEntryFromFile(f, Path.Combine("sources", Path.GetFileName(f)));
                    }
                }
                archive.CreateEntryFromFile(auditLogPath, Path.GetFileName(auditLogPath));
            }
            return zipPath;
        }
    }
}
