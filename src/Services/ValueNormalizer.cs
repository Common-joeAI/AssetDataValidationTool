// src/Services/ValueNormalizer.cs
using System;
using System.Text.RegularExpressions;

namespace AssetDataValidationTool.Services
{
    internal static class ValueNormalizer
    {
        public static string NormalizePk(string header, string value)
            => NormalizeByHeader(header, value, pkMode: true);

        public static string NormalizeValue(string header, string value)
            => NormalizeByHeader(header, value, pkMode: false);

        private static string NormalizeByHeader(string header, string value, bool pkMode)
        {
            header = (header ?? string.Empty).Trim();
            value  = (value  ?? string.Empty).Trim();

            value = Regex.Replace(value, @"\s+", " ");
            value = value.Trim('.', '-', '_');
            var upperHeader = header.ToUpperInvariant();

            if (upperHeader.Contains("MAC"))
            {
                var mac = Regex.Replace(value, @"[^0-9A-Fa-f]", "");
                return mac.ToUpperInvariant();
            }

            if (upperHeader.Contains("IP"))
            {
                var m = Regex.Match(value, @"^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\s*$");
                if (m.Success)
                {
                    int o1 = int.Parse(m.Groups[1].Value);
                    int o2 = int.Parse(m.Groups[2].Value);
                    int o3 = int.Parse(m.Groups[3].Value);
                    int o4 = int.Parse(m.Groups[4].Value);
                    return $"{o1}.{o2}.{o3}.{o4}";
                }
                return value;
            }

            if (upperHeader.Contains("HOST") || upperHeader.Contains("NAME"))
                return value.ToLowerInvariant();

            if (upperHeader.Contains("SERIAL") || upperHeader.Contains("S\\N") || upperHeader.Contains("ASSET TAG") || upperHeader.Contains("ASSET_TAG"))
            {
                var s = Regex.Replace(value, @"[\s\-]", "");
                return s.ToUpperInvariant();
            }

            return pkMode ? value.ToUpperInvariant() : value;
        }
    }
}
