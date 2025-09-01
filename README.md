# Asset Data Validation Automation Tool

Windows Forms app (.NET 8, C#) for monthly asset data validation.

## Build & Run
- Open `AssetDataValidationTool.sln` in Visual Studio 2022+
- `Restore` NuGet packages
- Build & run

## How to use
1. Choose **Asset Class**.
2. Choose **Data Point** (key column for matching across files).
3. For each required source, click **Browse** and select a .xlsx or .csv file.
4. Click **Validate**.
   - Generates an Excel report in `Documents/AssetDataValidationOutput/<AssetClass>/ValidationReport_<AssetClass>_<timestamp>.xlsx`.
   - Writes an audit log in the same folder.
5. Click **Package Zip** to create `assetclass-yyyyMMdd-username.zip` containing:
   - The report
   - The selected source files (under `/sources`)
   - The audit log

## Configure required files per Asset Class
Edit `Config/assetclasses.json`. Example:
```json
{
  "Computers": ["Baseline","Discovery"],
  "Printers": ["Baseline","Discovery"]
}
```
Add/remove classes and labels as needed. Labels appear on the UI and in the report.

## Notes
- `.xlsx` reading/writing uses the **DocumentFormat.OpenXml** SDK (local library; no external network calls).
- CSV is also supported for inputs.
- The report includes:
  - **Summary**
  - **KeyPresence** (Yes/No per file for each key)
  - **Conflicts** (keys where same column values differ between sources)
  - **MatchesAll** (keys present in every source file)
  - **MissingByFile** (keys missing from specific files)
  - **Source_*** preview sheets (first 100 rows) to help mimic input structure

## Output Variables (where to find them in code)
- `ValidationResults.AssetClass`
- `ValidationResults.DataPoint`
- `ValidationResults.Sources`
- `ValidationResults.Presence`
- `ValidationResults.Conflicts`
- `ValidationResults.ReportFilePath`
- `ValidationResults.AuditLogPath`
- `ValidationResults.ZipPackagePath`

## Compatibility
- Windows 10/11
- .NET 8 Desktop Runtime
