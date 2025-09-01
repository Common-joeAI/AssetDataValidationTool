# Asset Data Validation Automation Tool

A fast, local-first Windows app for comparing **baseline** vs **discovery** data across multiple sources (CSV / Excel). It normalizes headers, lets you pick **Primary Keys (PKs)** per source, auto-inferrs **field mappings**, and exports a rich **Excel report** with deltas, summaries, and audit logs.

> Built with C# / WinForms and the Open XML SDK. No external services required — all processing happens on your machine.

---

## ✨ Highlights

- **Per‑source Primary Key selector** (join on hostnames, serials, asset tags, etc.).  
- **Auto field mapping** between sources using value overlap (with a sensible similarity threshold).  
- **Delta detection** across sources against a chosen Baseline.  
- **One‑click Excel report** with multiple tabs (see below) + **zip packaging** for hand‑off.  
- **Profile loader** from “_Data Validation – &lt;AssetClass&gt;.xlsx_” to auto-populate source labels.  
- **Simple config JSON** to define asset classes & required sources.  
- **Local audit log** and reproducible output folders per asset class.

---

## 📦 Output: What the report contains

The generated workbook includes (at least) these tabs:

- **Summary** – key counts, totals, and high‑level metrics.  
- **KeyPresence** – which keys exist in which sources.  
- **Conflicts** – rows/columns where values disagree.  
- **MatchesAll** – keys present (and consistent) in all sources.  
- **MissingByFile** – keys missing from specific files.  
- **FieldMapping** – inferred column mapping baseline ⟶ other source (with a MatchScore).  
- **Deltas** – per‑key, per‑column mismatches vs Baseline using the inferred mappings.  
- **DeltasSummary** – mismatch counts by source (used for charts).  
- **Charts** – a simple column chart (“Delta Cells by Source”). _Chart creation is best‑effort via Open XML; the data tables are the source of truth._

**Notes**  
- The current mapping heuristic pairs columns by **value overlap** across rows sharing the same selected PK. Default threshold: **0.60**.  
- Baseline is determined by the source named **“Baseline”** (case‑insensitive). If none, the first source is used as Baseline.

---

## 🖥️ App Walkthrough

1. **Asset Class** – choose an asset class (e.g., “Computers”).  
2. **Data Point** – pick a default data point (e.g., “Hostname”).  
3. **Source Files** – for each required source label, browse to a file (`.csv` or `.xlsx`).  
4. **PK** – select the **Primary Key** for each source from its header dropdown.  
5. **Validate** – runs normalization, presence checks, field mapping, and delta detection.  
6. **Open Report** – view the workbook; or **Package Zip** to bundle outputs for hand‑off.

Outputs are written to:
```
%USERPROFILE%\Documents\AssetDataValidationOutput\<AssetClass>\
```

You’ll also get an **audit log** capturing the run context (asset class, data point, file list, timestamps).

---

## 🛠️ Tech & Project Structure

- **UI**: `Forms/MainForm.cs`  
  - Table‑layout grid for sources: `Label | File | Browse | Status | PK: | PK Combo`  
  - PK dropdown populated from the file’s header row.
- **Models**: `Models/ValidationModels.cs` (results, presence/conflict models), `Models/InputRequirement.cs` (config input shape).
- **Validation**: `Services/Validator.cs` – composes sources into a unified model, key presence, conflicts, etc.
- **Excel I/O**:  
  - `Services/ExcelReader.cs` – reads headers/values from `.xlsx` (Open XML) & `.csv`.  
  - `Services/ReportGenerator.cs` – writes the multi‑sheet workbook, field mappings, deltas, and charts.  
  - `Services/TemplateProfileReader.cs` – parses “Data Validation – &lt;AssetClass&gt;.xlsx” to discover source labels.
- **Packaging**: `Services/Packager.cs` – creates a deployment zip with report, audit log, and source manifests.

**Dependencies**
- [.NET 6+ for Windows](https://dotnet.microsoft.com/) (WinForms)  
- [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml) (Open XML SDK)  

> The code uses modern C# features (e.g., target‑typed `new()`), so use **Visual Studio 2022+** or `dotnet` SDK with the Windows desktop workload.

---

## ⚙️ Config

The app looks for a JSON config at:
```
.\Config\assetclasses.json
```

**Shape (new style):**
```jsonc
{
  "Computers": [
    { "label": "Baseline",  "description": "Authoritative source", "patterns": ["*cmdb*.xlsx", "*mecm*.csv"] },
    { "label": "Discovery", "description": "Latest network scan",   "patterns": ["*nmap*.csv", "*nessus*.csv"] },
    { "label": "Active Directory" },
    { "label": "Ham Export" }
  ],
  "Switches": [
    { "label": "Baseline" },
    { "label": "Discovery" }
  ]
}
```
- `label` (required): the friendly name that appears in the UI.  
- `description` (optional): tooltip text.  
- `patterns` (optional): filename wildcards to hint “name looks right / unexpected”.

**Back‑compat (old style)** – if you already have a legacy config like:
```json
{ "Computers": ["Baseline", "Discovery"] }
```
it will be auto‑interpreted.

**Profile Loader (optional)**  
Click **“Load From Validation Workbook”** and select a file named:
```
Data Validation - <AssetClass>.xlsx
```
The app scans it to produce the source labels for that asset class, then updates your dropdown on the fly.

---

## 🚀 Build & Run

### Visual Studio
1. Open the solution in **Visual Studio 2022** (Windows).  
2. Ensure the project targets `net6.0-windows` (or your desired WinForms target).  
3. Restore NuGet packages (make sure `DocumentFormat.OpenXml` is installed).  
4. Build & Run.

### CLI
```bash
# from the solution folder
dotnet build
dotnet run --project src/AssetDataValidationTool
```

> If you’re on .NET Framework instead, keep using VS; language version must support C# 9/10 features or adjust the syntax accordingly.

---

## 🔍 Troubleshooting

- **Read/Write errors**: close any open copy of the output workbook before re‑running.  
- **Headers not appearing in PK dropdown**: ensure your CSV/Excel has a header row; for Excel, the first row in the first worksheet is used.  
- **Charts missing**: Excel can still open the workbook — the chart is best‑effort via Open XML; the **DeltasSummary** sheet holds the data used for plotting.  
- **UI looks cramped**: this branch uses a `TableLayoutPanel` grid so controls stretch with window width. Try resizing — the file textbox and PK dropdown grow with the window.

---

## 🧭 Roadmap / Ideas

- Toggleable threshold slider for field‑mapping similarity.  
- Save/restore per‑source PK choices per asset class.  
- Optional “Top N mismatching columns” summary + chart.  
- Column normalization helpers (trimming, casings, date parsing) with user presets.  
- Export to CSV alongside Excel.

---

## 🤝 Contributing

Pull requests welcome! If you’re proposing a larger change, open an issue first to discuss scope. Please keep PRs focused and include a short test dataset or screenshots when UI changes are involved.

---

## 📄 License

**Choose a license** (e.g., MIT) and place it at the repo root as `LICENSE`.  
If unspecified, all rights reserved by default.

---

## 📸 Screenshots

Place images in `docs/` and reference here, e.g.:

```
docs/
  screenshot-main.png
```

Then in this README:

```md
![Main window](docs/screenshot-main.png)
```

---

## 🙌 Credits

- Open XML SDK team for the excellent `DocumentFormat.OpenXml` library.
- Everyone contributing test files and ideas that shaped the PK mapping & deltas views.
