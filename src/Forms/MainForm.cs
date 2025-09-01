// MainForm.cs — conservative version (no target-typed new, safe strings)
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Drawing;
using AssetDataValidationTool.Models;
using AssetDataValidationTool.Services;

namespace AssetDataValidationTool.Forms
{
    public class MainForm : Form
    {
        private readonly ComboBox cmbAssetClass = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 300 };
        private readonly ComboBox cmbDataPoint = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 300 };
        private readonly Button btnValidate = new Button { Text = "Validate", Enabled = false, Width = 120, Height = 36 };
        private readonly Button btnZip = new Button { Text = "Package Zip", Enabled = false, Width = 120, Height = 36 };
        private readonly Button btnOpenOutput = new Button { Text = "Open Output Folder", Width = 160, Height = 36 };
        private readonly Button btnLoadProfile = new Button { Text = "Load From Validation Workbook", Width = 260, Height = 32 };
        private readonly Button btnOpenConfig = new Button { Text = "Open Config", Width = 120, Height = 32 };
        private readonly Button btnApiSettings = new Button { Text = "API Settings (Future)", Width = 180, Height = 32 };

        private readonly TableLayoutPanel tblFiles = new TableLayoutPanel();
        private readonly StatusStrip status = new StatusStrip();
        private readonly ToolStripStatusLabel statusLabel = new ToolStripStatusLabel("Ready");
        private readonly ToolTip tip = new ToolTip();

        private readonly string outputRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "AssetDataValidationOutput");
        private Dictionary<string, List<InputRequirement>> requiredFilesByAsset = new Dictionary<string, List<InputRequirement>>(StringComparer.OrdinalIgnoreCase);
        private readonly List<(InputRequirement req, Label lbl, TextBox pathBox, Button browseBtn, Label status, ComboBox pkBox)> fileInputs = new List<(InputRequirement, Label, TextBox, Button, Label, ComboBox)>();
        private ValidationResults? lastResults; // Made nullable

        public MainForm()
        {
            Text = "Asset Data Validation Automation Tool";
            StartPosition = FormStartPosition.CenterScreen;
            Width = 1100; Height = 720;
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Segoe UI", 10f);

            status.Items.Add(statusLabel);
            status.Dock = DockStyle.Bottom;
            Controls.Add(status);

            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 6, Padding = new Padding(12) };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);

            // Row 1
            var row1 = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 5, AutoSize = true, Margin = new Padding(0, 0, 0, 6) };
            row1.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            row1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 320));
            row1.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            row1.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            row1.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            row1.Controls.Add(new Label { Text = "Asset Class:", AutoSize = true, Padding = new Padding(0, 8, 8, 0) }, 0, 0);
            row1.Controls.Add(cmbAssetClass, 1, 0);
            row1.Controls.Add(btnOpenConfig, 2, 0);
            row1.Controls.Add(btnLoadProfile, 3, 0);
            row1.Controls.Add(btnApiSettings, 4, 0);
            root.Controls.Add(row1, 0, 0);

            btnOpenConfig.Click += (s, e) => OpenConfig();
            btnLoadProfile.Click += (s, e) => LoadFromValidationWorkbook();
            btnApiSettings.Click += (s, e) => { try { using (var f = new ApiSettingsForm()) f.ShowDialog(this); } catch { } };

            var explainer = new Label
            {
                Text = "Baseline = authoritative export (CMDB/MECM/etc).  Discovery = latest scan (SNMP/Nessus/Nmap).",
                AutoSize = true,
                ForeColor = Color.DimGray,
                Margin = new Padding(0, 0, 0, 8)
            };
            root.Controls.Add(explainer, 0, 1);

            // Row 2
            var row2 = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 2, AutoSize = true, Margin = new Padding(0, 0, 0, 8) };
            row2.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            row2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 320));
            row2.Controls.Add(new Label { Text = "Data Point:", AutoSize = true, Padding = new Padding(0, 8, 8, 0) }, 0, 0);
            row2.Controls.Add(cmbDataPoint, 1, 0);
            root.Controls.Add(row2, 0, 2);

            // Files group
            var grp = new GroupBox { Text = "Source Files (hover labels for hints; ✓ name looks right / ! unexpected):", Dock = DockStyle.Fill, Padding = new Padding(10), Margin = new Padding(0, 0, 0, 8) };
            tblFiles.Dock = DockStyle.Fill; tblFiles.AutoSize = false; tblFiles.AutoScroll = true;
            tblFiles.ColumnCount = 6;
            tblFiles.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            tblFiles.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            tblFiles.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            tblFiles.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            tblFiles.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            tblFiles.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 240));
            grp.Controls.Add(tblFiles);
            root.Controls.Add(grp, 0, 3);

            var divider = new Panel { Height = 1, Dock = DockStyle.Top, BackColor = SystemColors.ControlDark, Margin = new Padding(0, 8, 0, 8) };
            root.Controls.Add(divider, 0, 4);

            // Buttons
            var buttons = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 3, AutoSize = true };
            buttons.Controls.Add(btnValidate, 0, 0);
            buttons.Controls.Add(btnZip, 1, 0);
            buttons.Controls.Add(btnOpenOutput, 2, 0);
            root.Controls.Add(buttons, 0, 5);

            cmbAssetClass.SelectedIndexChanged += (s, e) => RebuildFileInputs();
            cmbDataPoint.SelectedIndexChanged += (s, e) => RefreshValidateButton();
            btnValidate.Click += (s, e) => RunValidation();
            btnZip.Click += (s, e) => PackageZip();
            btnOpenOutput.Click += (s, e) => OpenOutputFolder();

            LoadConfig();
            LoadDefaults();
        }

        private string ConfigPath { get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config", "assetclasses.json"); } }

        private void LoadDefaults()
        {
            string[] defaults = new string[] { "Hostname", "IP Address", "MAC Address", "Serial Number", "Asset Tag", "Device Type", "Operating System", "Location", "Owner" };
            cmbDataPoint.Items.Clear();
            cmbDataPoint.Items.AddRange(defaults);
            if (cmbDataPoint.Items.Count > 0) cmbDataPoint.SelectedIndex = 0;
        }

        private void LoadConfig()
        {
            try
            {
                var dir = Path.GetDirectoryName(ConfigPath);
                if (string.IsNullOrEmpty(dir))
                    throw new InvalidOperationException("Config path directory could not be resolved.");
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                if (!File.Exists(ConfigPath))
                {
                    var minimal = @"{""Computers"": [ { ""label"": ""Baseline"" }, { ""label"": ""Discovery"" } ]}";
                    File.WriteAllText(ConfigPath, minimal);
                }

                var json = File.ReadAllText(ConfigPath);
                var opts = new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true };

                Dictionary<string, List<InputRequirement>>? newShape = null;
                try { newShape = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, List<InputRequirement>>>(json, opts); } catch { }

                if (newShape != null && newShape.Count > 0)
                {
                    requiredFilesByAsset = new Dictionary<string, List<InputRequirement>>(newShape, StringComparer.OrdinalIgnoreCase);
                }
                else
                {
                    Dictionary<string, List<string>>? oldShape = null;
                    try { oldShape = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, List<string>>>(json, opts); } catch { }

                    requiredFilesByAsset = new Dictionary<string, List<InputRequirement>>(StringComparer.OrdinalIgnoreCase);
                    if (oldShape != null)
                    {
                        foreach (var kv in oldShape)
                        {
                            var list = new List<InputRequirement>();
                            foreach (var s in kv.Value) list.Add(new InputRequirement { Label = s });
                            requiredFilesByAsset[kv.Key] = list;
                        }
                    }
                }

                cmbAssetClass.Items.Clear();
                cmbAssetClass.Items.AddRange(requiredFilesByAsset.Keys.OrderBy(s => s).ToArray());
                if (cmbAssetClass.Items.Count > 0) cmbAssetClass.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Failed to load config: " + ex.Message);
            }
        }

        private void OpenConfig()
        {
            try { Process.Start(new ProcessStartInfo { FileName = ConfigPath, UseShellExecute = true }); }
            catch (Exception ex) { MessageBox.Show(this, "Unable to open config: " + ex.Message); }
        }

        private void LoadFromValidationWorkbook()
        {
            using (var ofd = new OpenFileDialog { Filter = "Excel Workbook|*.xlsx|All files|*.*", Title = "Pick a 'Data Validation - <AssetClass>.xlsx' workbook" })
            {
                if (ofd.ShowDialog(this) != DialogResult.OK) return;

                var result = TemplateProfileReader.ExtractFromValidationWorkbook(ofd.FileName);
                var assetClass = result.Item1;
                var labels = result.Item2;

                if (string.IsNullOrWhiteSpace(assetClass))
                {
                    MessageBox.Show(this, "Could not infer Asset Class from filename. Expected: Data Validation - <AssetClass>.xlsx");
                    return;
                }

                if (labels == null || labels.Count == 0)
                {
                    if (!requiredFilesByAsset.ContainsKey(assetClass))
                        requiredFilesByAsset[assetClass] = new List<InputRequirement>
                        {
                            new InputRequirement { Label = "Baseline" },
                            new InputRequirement { Label = "Discovery" }
                        };
                }
                else
                {
                    requiredFilesByAsset[assetClass] = labels.Select(s => new InputRequirement
                    {
                        Label = s,
                        Description = "Source defined in " + Path.GetFileName(ofd.FileName),
                        Patterns = null
                    }).ToList();
                }

                if (!cmbAssetClass.Items.Contains(assetClass)) cmbAssetClass.Items.Add(assetClass);
                cmbAssetClass.SelectedItem = assetClass;

                RebuildFileInputs();
                statusLabel.Text = "Loaded profile for " + assetClass + " (" + (labels != null ? labels.Count : 0) + " source label(s))";
            }
        }

        private void RebuildFileInputs()
        {
            tblFiles.SuspendLayout();
            tblFiles.Controls.Clear();
            tblFiles.RowStyles.Clear();
            tblFiles.RowCount = 0;
            fileInputs.Clear();

            var asset = cmbAssetClass.SelectedItem?.ToString() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(asset)) { tblFiles.ResumeLayout(); return; }

            List<InputRequirement>? reqs;
            if (!requiredFilesByAsset.TryGetValue(asset, out reqs) || reqs == null || reqs.Count == 0)
            {
                reqs = new List<InputRequirement> { new InputRequirement { Label = "Baseline" }, new InputRequirement { Label = "Discovery" } };
            }

            // Set up consistent column widths
            tblFiles.ColumnStyles.Clear();
            tblFiles.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150)); // Label column - fixed width
            tblFiles.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));  // TextBox column - fills available space
            tblFiles.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // Browse button
            tblFiles.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // Status label
            tblFiles.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // PK label
            tblFiles.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 240)); // PK dropdown - fixed width

            int row = 0;
            foreach (var req in reqs)
            {
                tblFiles.RowCount++;
                tblFiles.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                // Create controls with consistent margins and anchoring
                var lbl = new Label
                {
                    Text = req.Label + ":",
                    AutoSize = true,
                    Margin = new Padding(0, 6, 8, 6),
                    Anchor = AnchorStyles.Left | AnchorStyles.Top,
                    TextAlign = ContentAlignment.MiddleLeft
                };

                var txt = new TextBox
                {
                    Margin = new Padding(0, 3, 8, 3),
                    Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
                    MinimumSize = new Size(300, 0),
                    Dock = DockStyle.Fill
                };

                var btn = new Button
                {
                    Text = "Browse...",
                    AutoSize = true,
                    Margin = new Padding(0, 2, 8, 2),
                    Anchor = AnchorStyles.Left | AnchorStyles.Top
                };

                var statusLbl = new Label
                {
                    Text = "—",
                    AutoSize = true,
                    ForeColor = Color.DimGray,
                    Margin = new Padding(0, 6, 8, 6),
                    Anchor = AnchorStyles.Left | AnchorStyles.Top,
                    TextAlign = ContentAlignment.MiddleLeft
                };

                var pkCap = new Label
                {
                    Text = "PK:",
                    AutoSize = true,
                    Margin = new Padding(0, 6, 8, 6),
                    Anchor = AnchorStyles.Left | AnchorStyles.Top,
                    TextAlign = ContentAlignment.MiddleLeft
                };

                var pkBox = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Width = 220,
                    Margin = new Padding(0, 3, 0, 3),
                    Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
                    Dock = DockStyle.Fill
                };

                tip.SetToolTip(lbl, req.Description ?? "");
                tip.SetToolTip(txt, req.Description ?? "");
                tip.SetToolTip(statusLbl, req.Description ?? "");

                btn.Click += (s, e) => BrowseForFile(txt, statusLbl, req, pkBox);
                txt.TextChanged += (s, e) => { if (File.Exists(txt.Text)) PopulatePkOptions(pkBox, txt.Text); };

                // Add controls to the table with proper alignment
                tblFiles.Controls.Add(lbl, 0, row);
                tblFiles.Controls.Add(txt, 1, row);
                tblFiles.Controls.Add(btn, 2, row);
                tblFiles.Controls.Add(statusLbl, 3, row);
                tblFiles.Controls.Add(pkCap, 4, row);
                tblFiles.Controls.Add(pkBox, 5, row);

                fileInputs.Add((req, lbl, txt, btn, statusLbl, pkBox));
                row++;
            }

            tblFiles.ResumeLayout();
            RefreshValidateButton();
        }


        private static void PopulatePkOptions(ComboBox pkBox, string filePath)
        {
            try
            {
                pkBox.Items.Clear();
                if (File.Exists(filePath))
                {
                    var headers = ExcelReader.ReadHeaders(filePath);
                    foreach (var h in headers) pkBox.Items.Add(h);
                    if (pkBox.Items.Count > 0) pkBox.SelectedIndex = 0;
                }
            }
            catch { }
        }

        private void BrowseForFile(TextBox target, Label statusLabelForRow, InputRequirement req, ComboBox pkBox)
        {
            using (var ofd = new OpenFileDialog { Filter = "Excel/CSV|*.xlsx;*.csv|All files|*.*" })
            {
                if (ofd.ShowDialog(this) == DialogResult.OK)
                {
                    target.Text = ofd.FileName;
                    PopulatePkOptions(pkBox, ofd.FileName);

                    bool matches = MatchesAnyPattern(ofd.FileName, req.Patterns ?? new List<string>()); // Fixed null handling
                    if (req.Patterns != null && req.Patterns.Count > 0)
                    {
                        statusLabelForRow.Text = matches ? "✓ name looks right" : "! unexpected name";
                        statusLabelForRow.ForeColor = matches ? Color.Green : Color.OrangeRed;
                        target.BackColor = matches ? Color.Honeydew : Color.LemonChiffon;
                    }
                    else
                    {
                        statusLabelForRow.Text = "—";
                        statusLabelForRow.ForeColor = Color.DimGray;
                        target.BackColor = SystemColors.Window;
                    }
                    RefreshValidateButton();
                }
            }
        }

        private void RefreshValidateButton()
        {
            bool allSelected = fileInputs.All(f =>
                !string.IsNullOrWhiteSpace(f.pathBox.Text) &&
                File.Exists(f.pathBox.Text) &&
                f.pkBox.SelectedItem != null);
            btnValidate.Enabled = allSelected && cmbDataPoint.SelectedItem != null;
        }

        private async void RunValidation()
        {
            if (!btnValidate.Enabled) return;

            try
            {
                var asset = cmbAssetClass.SelectedItem?.ToString() ?? "Unknown"; // Fixed null handling
                var dataPoint = cmbDataPoint.SelectedItem?.ToString() ?? "Hostname"; // Fixed null handling
                var selected = fileInputs.Select(f => (displayName: f.req.Label, filePath: f.pathBox.Text)).ToList();

                statusLabel.Text = "Validating...";
                UseWaitCursor = true;
                Application.DoEvents();

                var results = Validator.Validate(asset, dataPoint, selected);

                // Fixed: Convert nullable dictionary to non-nullable
                var pkMap = fileInputs.ToDictionary(
                    f => f.req.Label,
                    f => f.pkBox.SelectedItem?.ToString() ?? dataPoint);
                results.PrimaryKeyBySource = pkMap;

                var outFolder = Path.Combine(outputRoot, asset.Replace(' ', '_'));
                Directory.CreateDirectory(outFolder);

                var report = ReportGenerator.GenerateExcelReport(results, outFolder);
                var audit = AuditLogger.WriteAuditLog(outFolder, asset, dataPoint, selected);

                results.ReportFilePath = report;
                results.AuditLogPath = audit;
                lastResults = results;

                try
                {
                    var apiSettings = ApiSettingsStore.Load();
                    if (apiSettings.Enabled)
                    {
                        using (var api = new HttpApiClient(apiSettings))
                        {
                            await TryPushAsync(api, report, selected);
                        }
                    }
                }
                catch { }

                btnZip.Enabled = true;
                statusLabel.Text = "Report created: " + report;
                UseWaitCursor = false;

                if (MessageBox.Show(this, "Open report?", "Validation Complete", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    Process.Start(new ProcessStartInfo { FileName = report, UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                UseWaitCursor = false;
                statusLabel.Text = "Error";
                MessageBox.Show(this, "Validation failed: " + ex.Message);
            }
        }

        private static System.Threading.Tasks.Task TryPushAsync(HttpApiClient api, string reportPath, List<(string displayName, string filePath)> selected)
        {
            var s = ApiSettingsStore.Load();
            if (!s.Enabled) return System.Threading.Tasks.Task.CompletedTask;

            return System.Threading.Tasks.Task.Run(async () =>
            {
                foreach (var item in selected)
                {
                    try { await api.UploadSourceAsync(item.displayName, item.filePath).ConfigureAwait(false); } catch { }
                }
                try { await api.UploadReportAsync(reportPath).ConfigureAwait(false); } catch { }
            });
        }

        private void PackageZip()
        {
            if (lastResults == null)
            {
                MessageBox.Show(this, "No validation results to package. Run Validate first.");
                return;
            }

            try
            {
                var asset = lastResults.AssetClass;
                var outFolder = Path.Combine(outputRoot, asset.Replace(' ', '_'));

                var zip = Packager.CreateZip(
                    asset,
                    lastResults.ReportFilePath,
                    lastResults.Sources.Select(s => s.FilePath),
                    lastResults.AuditLogPath,
                    outFolder
                );

                lastResults.ZipPackagePath = zip;
                statusLabel.Text = "Packaged: " + zip;

                if (MessageBox.Show(this, "Open zip folder?", "Packaging Complete", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    Process.Start(new ProcessStartInfo { FileName = outFolder, UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Packaging failed: " + ex.Message);
            }
        }

        private void OpenOutputFolder()
        {
            try
            {
                Directory.CreateDirectory(outputRoot);
                Process.Start(new ProcessStartInfo { FileName = outputRoot, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Unable to open output folder: " + ex.Message);
            }
        }

        private static bool MatchesAnyPattern(string filePath, List<string> patterns)
        {
            if (patterns == null || patterns.Count == 0) return true;
            var name = Path.GetFileName(filePath);
            foreach (var p in patterns) if (WildcardIsMatch(name, p)) return true;
            return false;
        }

        private static bool WildcardIsMatch(string text, string wildcard)
        {
            if (string.IsNullOrEmpty(wildcard)) return false;
            string pattern = "^" + Regex.Escape(wildcard).Replace(@"\*", ".*").Replace(@"\?", ".") + "$";
            return Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase);
        }
    }
}
