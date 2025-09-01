using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Drawing;
using AssetDataValidationTool.Models;
using AssetDataValidationTool.Services;

namespace AssetDataValidationTool.Forms
{
    public class MainForm : Form
    {
        private ComboBox cmbAssetClass = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 300 };
        private ComboBox cmbDataPoint = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 300 };
        private FlowLayoutPanel filesPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoScroll = true, WrapContents = false, Width = 620, Height = 280 };
        private Button btnValidate = new Button { Text = "Validate", Enabled = false, Width = 120, Height = 36 };
        private Button btnZip = new Button { Text = "Package Zip", Enabled = false, Width = 120, Height = 36 };
        private Button btnOpenOutput = new Button { Text = "Open Output Folder", Width = 160, Height = 36 };
        private Button btnLoadProfile = new Button { Text = "Load From Validation Workbook", Width = 260, Height = 32 };
        private StatusStrip status = new StatusStrip();
        private ToolStripStatusLabel statusLabel = new ToolStripStatusLabel("Ready");
        private ToolTip tip = new ToolTip();

        private string outputRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "AssetDataValidationOutput");
        private Dictionary<string, List<InputRequirement>> requiredFilesByAsset = new(StringComparer.OrdinalIgnoreCase);
        private List<(InputRequirement req, Label lbl, TextBox pathBox, Button browseBtn, Label status)> fileInputs = new();
        private ValidationResults? lastResults;

        public MainForm()
        {
            Text = "Asset Data Validation Automation Tool";
            Width = 820;
            Height = 640;
            StartPosition = FormStartPosition.CenterScreen;

            status.Items.Add(statusLabel);
            status.Dock = DockStyle.Bottom;

            var top = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                AutoScroll = true,
                WrapContents = false,
                Padding = new Padding(12)
            };
            Controls.Add(top);
            Controls.Add(status);

            // Row: Asset + Config buttons
            var row1 = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Width = 760, Height = 44 };
            row1.Controls.Add(new Label { Text = "Asset Class:", AutoSize = true, Padding = new Padding(0, 12, 8, 0) });
            row1.Controls.Add(cmbAssetClass);

            var btnEditConfig = new Button { Text = "Open Config", Width = 120, Height = 32 };
            btnEditConfig.Click += (s, e) => OpenConfig();
            row1.Controls.Add(btnEditConfig);

            btnLoadProfile.Click += (s, e) => LoadFromValidationWorkbook();
            row1.Controls.Add(btnLoadProfile);
            top.Controls.Add(row1);

            // Explainer
            top.Controls.Add(new Label
            {
                Text = "Baseline = authoritative export (CMDB/MECM/etc).  Discovery = latest scan (SNMP/Nessus/Nmap).",
                AutoSize = true,
                ForeColor = Color.DimGray,
                Padding = new Padding(0, 2, 0, 8)
            });

            // Row: Data Point
            var row2 = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Width = 760, Height = 44 };
            row2.Controls.Add(new Label { Text = "Data Point:", AutoSize = true, Padding = new Padding(0, 12, 8, 0) });
            row2.Controls.Add(cmbDataPoint);
            top.Controls.Add(row2);

            // Files header
            top.Controls.Add(new Label { Text = "Source Files (hover labels for hints; ✓ name looks right / ! unexpected):", AutoSize = true, Padding = new Padding(0, 6, 0, 4) });
            top.Controls.Add(filesPanel);

            // Buttons row
            var buttonsRow = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Width = 760, Height = 50 };
            buttonsRow.Controls.Add(btnValidate);
            buttonsRow.Controls.Add(btnZip);
            buttonsRow.Controls.Add(btnOpenOutput);
            top.Controls.Add(buttonsRow);

            // Events
            cmbAssetClass.SelectedIndexChanged += (s, e) => RebuildFileInputs();
            cmbDataPoint.SelectedIndexChanged += (s, e) => RefreshValidateButton();
            btnValidate.Click += (s, e) => RunValidation();
            btnZip.Click += (s, e) => PackageZip();
            btnOpenOutput.Click += (s, e) => OpenOutputFolder();

            LoadConfig();
            LoadDefaults();
        }

        private string ConfigPath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config", "assetclasses.json");

        private void LoadDefaults()
        {
            string[] defaultDataPoints = new[]
            {
                "Hostname","IP Address","MAC Address","Serial Number","Asset Tag","Device Type","Operating System","Location","Owner"
            };
            cmbDataPoint.Items.Clear();
            cmbDataPoint.Items.AddRange(defaultDataPoints);
            if (cmbDataPoint.Items.Count > 0) cmbDataPoint.SelectedIndex = 0;
        }

        private void LoadConfig()
        {
            try
            {
                if (!File.Exists(ConfigPath))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(ConfigPath)!);
                    var minimal = "{\n  \"Computers\": [ { \"label\": \"Baseline\" }, { \"label\": \"Discovery\" } ]\n}";
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

                    requiredFilesByAsset = new(StringComparer.OrdinalIgnoreCase);
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
                MessageBox.Show(this, $"Failed to load config: {ex.Message}");
            }
        }

        private void OpenConfig()
        {
            try
            {
                Process.Start(new ProcessStartInfo { FileName = ConfigPath, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Unable to open config: {ex.Message}");
            }
        }

        /// <summary>
        /// Load asset class + source labels from a "Data Validation - <AssetClass>.xlsx" workbook.
        /// </summary>
        private void LoadFromValidationWorkbook()
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "Excel Workbook|*.xlsx|All files|*.*",
                Title = "Pick a 'Data Validation - <AssetClass>.xlsx' workbook"
            };
            if (ofd.ShowDialog(this) != DialogResult.OK) return;

            var (assetClass, labels) = TemplateProfileReader.ExtractFromValidationWorkbook(ofd.FileName);

            if (string.IsNullOrWhiteSpace(assetClass))
            {
                MessageBox.Show(this, "Could not infer Asset Class from filename. Expected: Data Validation - <AssetClass>.xlsx");
                return;
            }

            if (labels == null || labels.Count == 0)
            {
                // Still allow setting the asset class; user can manually edit config later
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
                    Description = $"Source defined in {Path.GetFileName(ofd.FileName)}",
                    Patterns = null
                }).ToList();
            }

            // Refresh dropdown (add if missing and select it)
            if (!cmbAssetClass.Items.Contains(assetClass))
                cmbAssetClass.Items.Add(assetClass);
            cmbAssetClass.SelectedItem = assetClass;

            RebuildFileInputs();
            statusLabel.Text = $"Loaded profile for {assetClass} ({labels.Count} source label(s))";
        }

        private void RebuildFileInputs()
        {
            filesPanel.Controls.Clear();
            fileInputs.Clear();

            var asset = cmbAssetClass.SelectedItem?.ToString() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(asset)) return;

            if (!requiredFilesByAsset.TryGetValue(asset, out var reqs) || reqs == null || reqs.Count == 0)
            {
                reqs = new List<InputRequirement> { new InputRequirement { Label = "Baseline" }, new InputRequirement { Label = "Discovery" } };
            }

            foreach (var req in reqs)
            {
                var row = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Width = 740, Height = 38, Padding = new Padding(0, 2, 0, 0) };
                var lbl = new Label { Text = $"{req.Label}:", AutoSize = true, Width = 260, Padding = new Padding(0, 10, 6, 0) };
                var txt = new TextBox { Width = 360, ReadOnly = true };
                var btn = new Button { Text = "Browse", Width = 100, Height = 28 };
                var statusLbl = new Label { Text = "—", AutoSize = true, Padding = new Padding(8, 10, 0, 0), ForeColor = Color.DimGray };

                // Tooltip (description)
                var tt = (req.Description ?? "Select the appropriate file.");
                tip.SetToolTip(lbl, tt); tip.SetToolTip(txt, tt); tip.SetToolTip(statusLbl, tt);

                btn.Click += (s, e) => BrowseForFile(txt, statusLbl, req);

                row.Controls.Add(lbl);
                row.Controls.Add(txt);
                row.Controls.Add(btn);
                row.Controls.Add(statusLbl);

                filesPanel.Controls.Add(row);
                fileInputs.Add((req, lbl, txt, btn, statusLbl));
            }

            RefreshValidateButton();
        }

        private void BrowseForFile(TextBox target, Label statusLabelForRow, InputRequirement req)
        {
            using var ofd = new OpenFileDialog { Filter = "Excel/CSV|*.xlsx;*.csv|All files|*.*" };
            if (ofd.ShowDialog(this) == DialogResult.OK)
            {
                target.Text = ofd.FileName;

                // If patterns were provided, evaluate; otherwise neutral display
                bool matches = MatchesAnyPattern(ofd.FileName, req.Patterns);
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

        private void RefreshValidateButton()
        {
            var allSelected = fileInputs.All(f => !string.IsNullOrWhiteSpace(f.pathBox.Text) && File.Exists(f.pathBox.Text));
            var hasDataPoint = cmbDataPoint.SelectedItem != null;
            btnValidate.Enabled = allSelected && hasDataPoint;
        }

        private void RunValidation()
        {
            if (!btnValidate.Enabled) return;

            try
            {
                var asset = cmbAssetClass.SelectedItem?.ToString() ?? "Unknown";
                var dataPoint = cmbDataPoint.SelectedItem?.ToString() ?? "Hostname";
                var selected = fileInputs.Select(f => (displayName: f.req.Label, filePath: f.pathBox.Text)).ToList();

                statusLabel.Text = "Validating...";
                UseWaitCursor = true;
                Application.DoEvents();

                var results = Validator.Validate(asset, dataPoint, selected);

                var outFolder = Path.Combine(outputRoot, asset.Replace(' ', '_'));
                Directory.CreateDirectory(outFolder);

                var report = ReportGenerator.GenerateExcelReport(results, outFolder);
                var audit = AuditLogger.WriteAuditLog(outFolder, asset, dataPoint, selected);

                results.ReportFilePath = report;
                results.AuditLogPath = audit;
                lastResults = results;

                btnZip.Enabled = true;
                statusLabel.Text = $"Report created: {report}";
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
                statusLabel.Text = $"Packaged: {zip}";

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
                MessageBox.Show(this, $"Unable to open output folder: {ex.Message}");
            }
        }

        private static bool MatchesAnyPattern(string filePath, List<string>? patterns)
        {
            if (patterns == null || patterns.Count == 0) return true;
            var name = Path.GetFileName(filePath);
            foreach (var p in patterns) if (WildcardIsMatch(name, p)) return true;
            return false;
        }

        private static bool WildcardIsMatch(string text, string wildcard)
        {
            if (string.IsNullOrEmpty(wildcard)) return false;
            string pattern = "^" + Regex.Escape(wildcard).Replace("\\*", ".*").Replace("\\?", ".") + "$";
            return Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase);
        }
    }
}
