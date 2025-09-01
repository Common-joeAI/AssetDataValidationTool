// src/Services/OpenXmlHelpers.cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace AssetDataValidationTool.Services
{
    internal static class OpenXmlHelpers
    {
        public static SharedStringTable EnsureSharedStringTable(WorkbookPart wbPart)
        {
            // Get existing part or create one
            var sstPart = wbPart.SharedStringTablePart ?? wbPart.AddNewPart<SharedStringTablePart>();
            if (sstPart.SharedStringTable == null)
                sstPart.SharedStringTable = new SharedStringTable();
            return sstPart.SharedStringTable;
        }

        public static int GetSharedStringIndex(WorkbookPart wbPart, string text)
        {
            var sst = EnsureSharedStringTable(wbPart);
            int index = 0;
            foreach (SharedStringItem item in sst.Elements<SharedStringItem>())
            {
                if (item.InnerText == (text ?? string.Empty)) return index;
                index++;
            }
            sst.AppendChild(new SharedStringItem(new Text(text ?? string.Empty)));
            sst.Save();
            return index;
        }

        public static (WorksheetPart wsPart, Sheet sheet) AddWorksheet(SpreadsheetDocument doc, string requestedName, HashSet<string>? usedNames = null)
        {
            var wbPart = doc.WorkbookPart!;
            var sheets = wbPart.Workbook.Sheets ?? (wbPart.Workbook.Sheets = new Sheets());

            string name = SanitizeSheetName(requestedName);
            usedNames ??= new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int disambiguator = 2;
            string baseName = name;

            while (usedNames.Contains(name) || sheets.Elements<Sheet>().Any(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase)))
            {
                name = baseName + " (" + disambiguator.ToString() + ")";
                if (name.Length > 31) name = name.Substring(0, 31);
                disambiguator++;
            }
            usedNames.Add(name);

            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet(new SheetData());
            wsPart.Worksheet.Save();

            uint newSheetId = 1;
            if (sheets.Elements<Sheet>().Any())
                newSheetId = sheets.Elements<Sheet>().Select(s => (uint)s.SheetId.Value).Max() + 1;

            var relId = wbPart.GetIdOfPart(wsPart);
            var sheet = new Sheet { Name = name, SheetId = newSheetId, Id = relId };
            sheets.Append(sheet);
            wbPart.Workbook.Save();
            return (wsPart, sheet);
        }

        public static string SanitizeSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) name = "Sheet";
            const string invalidChars = ":\\/?*[]";
            var sb = new StringBuilder(name.Length);
            foreach (var ch in name)
            {
                if (invalidChars.IndexOf(ch) >= 0) continue;
                sb.Append(ch);
            }
            name = sb.ToString().Trim();
            if (name.Length > 31) name = name.Substring(0, 31);
            if (name.Length == 0) name = "Sheet";
            return name;
        }

        public static SheetData GetSheetData(WorksheetPart wsPart)
        {
            var sd = wsPart.Worksheet.Elements<SheetData>().FirstOrDefault();
            if (sd == null)
            {
                sd = new SheetData();
                wsPart.Worksheet.Append(sd);
            }
            return sd;
        }

        public static Cell WriteTextCell(WorkbookPart wbPart, Row row, string columnRef, string text)
        {
            var cell = new Cell
            {
                CellReference = columnRef + row.RowIndex,
                DataType = CellValues.SharedString,
                CellValue = new CellValue(GetSharedStringIndex(wbPart, text ?? string.Empty).ToString())
            };
            row.Append(cell);
            return cell;
        }

        public static Cell WriteNumberCell(Row row, string columnRef, double number)
        {
            var cell = new Cell
            {
                CellReference = columnRef + row.RowIndex,
                DataType = null, // numeric
                CellValue = new CellValue(number.ToString(System.Globalization.CultureInfo.InvariantCulture))
            };
            row.Append(cell);
            return cell;
        }
    }
}
