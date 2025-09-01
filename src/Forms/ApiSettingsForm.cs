using AssetDataValidationTool.Models;
using AssetDataValidationTool.Services;
using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace AssetDataValidationTool.Forms
{
    public class ApiSettingsForm : Form
    {
        private TabControl tabs = new TabControl { Dock = DockStyle.Fill };
        private ToolTip tip = new ToolTip();

        // Global tab controls
        private CheckBox chkEnabled = new CheckBox { Text = "Enable experimental API integration", AutoSize = true };
        private TextBox txtBaseUrl = new TextBox { Width = 420 };
        private ComboBox cmbAuth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 200 };
        private TextBox txtApiKey = new TextBox { Width = 420 };
        private TextBox txtUser = new TextBox { Width = 200 };
        private TextBox txtPass = new TextBox { Width = 200, UseSystemPasswordChar = true };
        private TextBox txtHeaders = new TextBox { Width = 420, Height = 80, Multiline = true, ScrollBars = ScrollBars.Vertical };
        private TextBox txtSourceEndpoint = new TextBox { Width = 420 };
        private TextBox txtReportEndpoint = new TextBox { Width = 420 };
        private NumericUpDown numTimeout = new NumericUpDown { Minimum = 5, Maximum = 600, Value = 60, Width = 80 };

        // Buttons
        private Button btnSave = new Button { Text = "Save", Width = 100, Height = 32 };
        private Button btnCancel = new Button { Text = "Cancel", Width = 100, Height = 32 };

        // Per-service controls (minimal)
        // ServiceNow
        private CheckBox snEnabled = new CheckBox { Text = "Enable ServiceNow", AutoSize = true };
        private TextBox snUrl = new TextBox { Width = 420 };
        private TextBox snUser = new TextBox { Width = 200 };
        private TextBox snPass = new TextBox { Width = 200, UseSystemPasswordChar = true };
        private TextBox snTable = new TextBox { Width = 260 };
        private TextBox snQuery = new TextBox { Width = 420 };
        private NumericUpDown snPage = new NumericUpDown { Minimum = 1, Maximum = 10000, Value = 200, Width = 80 };

        // Nessus
        private CheckBox neEnabled = new CheckBox { Text = "Enable Nessus", AutoSize = true };
        private TextBox neUrl = new TextBox { Width = 420 };
        private TextBox neAccess = new TextBox { Width = 260 };
        private TextBox neSecret = new TextBox { Width = 260 };
        private NumericUpDown neTimeout = new NumericUpDown { Minimum = 5, Maximum = 600, Value = 60, Width = 80 };

        // Absolute
        private CheckBox abEnabled = new CheckBox { Text = "Enable Absolute", AutoSize = true };
        private TextBox abUrl = new TextBox { Width = 420 };
        private TextBox abKey = new TextBox { Width = 420 };
        private NumericUpDown abTimeout = new NumericUpDown { Minimum = 5, Maximum = 600, Value = 60, Width = 80 };

        // Active Directory
        private CheckBox adEnabled = new CheckBox { Text = "Enable Active Directory", AutoSize = true };
        private TextBox adPath = new TextBox { Width = 420 };
        private TextBox adUser = new TextBox { Width = 200 };
        private TextBox adPass = new TextBox { Width = 200, UseSystemPasswordChar = true };
        private TextBox adFilter = new TextBox { Width = 420 };
        private TextBox adAttrs = new TextBox { Width = 420 };

        // Azure AD
        private CheckBox aaEnabled = new CheckBox { Text = "Enable Azure AD", AutoSize = true };
        private TextBox aaTenant = new TextBox { Width = 380 };
        private TextBox aaClient = new TextBox { Width = 380 };
        private TextBox aaSecret = new TextBox { Width = 380, UseSystemPasswordChar = true };
        private TextBox aaScopes = new TextBox { Width = 420 };

        // Rapid7
        private CheckBox r7Enabled = new CheckBox { Text = "Enable Rapid7", AutoSize = true };
        private TextBox r7Url = new TextBox { Width = 420 };
        private TextBox r7User = new TextBox { Width = 200 };
        private TextBox r7Pass = new TextBox { Width = 200, UseSystemPasswordChar = true };
        private TextBox r7Key = new TextBox { Width = 420 };
        private NumericUpDown r7Timeout = new NumericUpDown { Minimum = 5, Maximum = 600, Value = 60, Width = 80 };

        public ApiSettingsForm()
        {
            Text = "API Settings (Future)";
            StartPosition = FormStartPosition.CenterParent;
            Width = 840;
            Height = 640;
            Font = new Font("Segoe UI", 10f);

            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(8) };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);

            // Set up tooltips with correct variable names
            tip.SetToolTip(txtBaseUrl, "Base URL");
            tip.SetToolTip(cmbAuth, "Auth Type");
            tip.SetToolTip(txtApiKey, "API Key");
            tip.SetToolTip(txtUser, "Username");
            tip.SetToolTip(txtPass, "Password");
            tip.SetToolTip(txtSourceEndpoint, "Source upload endpoint");
            tip.SetToolTip(txtReportEndpoint, "Report upload endpoint");

            // Create tabs - Fixed: Create TabPage objects properly
            tabs.TabPages.Add(new TabPage("Global") { Controls = { BuildGlobalPage() } });
            tabs.TabPages.Add(new TabPage("ServiceNow") { Controls = { BuildServiceNowPage() } });
            tabs.TabPages.Add(new TabPage("Nessus") { Controls = { BuildNessusPage() } });
            tabs.TabPages.Add(new TabPage("Absolute") { Controls = { BuildAbsolutePage() } });
            tabs.TabPages.Add(new TabPage("Active Directory") { Controls = { BuildAdPage() } });
            tabs.TabPages.Add(new TabPage("Azure AD") { Controls = { BuildAzureAdPage() } });
            tabs.TabPages.Add(new TabPage("Rapid7") { Controls = { BuildRapid7Page() } });

            root.Controls.Add(tabs, 0, 0);

            var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill, AutoSize = true };
            buttons.Controls.Add(btnSave);
            buttons.Controls.Add(btnCancel);
            root.Controls.Add(buttons, 0, 1);

            btnCancel.Click += (s, e) => DialogResult = DialogResult.Cancel;
            btnSave.Click += (s, e) => SaveSettings();

            LoadSettings();
        }

        private Control BuildGlobalPage()
        {
            var grid = NewGrid();
            int r = 0;
            grid.Controls.Add(chkEnabled, 0, r); grid.SetColumnSpan(chkEnabled, 2); r++;
            grid.Controls.Add(new Label { Text = "Base URL:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            grid.Controls.Add(txtBaseUrl, 1, r); r++;
            grid.Controls.Add(new Label { Text = "Auth Type:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            cmbAuth.Items.AddRange(new object[] { ApiAuthType.None, ApiAuthType.ApiKey, ApiAuthType.BearerToken, ApiAuthType.Basic });
            grid.Controls.Add(cmbAuth, 1, r); r++;
            grid.Controls.Add(new Label { Text = "API Key / Bearer:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            grid.Controls.Add(txtApiKey, 1, r); r++;
            grid.Controls.Add(new Label { Text = "Username / Password:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            var up = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true };
            up.Controls.Add(txtUser); up.Controls.Add(txtPass);
            grid.Controls.Add(up, 1, r); r++;
            grid.Controls.Add(new Label { Text = "Default Headers (Key: Value per line):", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            grid.Controls.Add(txtHeaders, 1, r); r++;
            grid.Controls.Add(new Label { Text = "Source Endpoint:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            grid.Controls.Add(txtSourceEndpoint, 1, r); r++;
            grid.Controls.Add(new Label { Text = "Report Endpoint:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            grid.Controls.Add(txtReportEndpoint, 1, r); r++;
            grid.Controls.Add(new Label { Text = "Timeout (sec):", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            grid.Controls.Add(numTimeout, 1, r); r++;
            return grid;
        }

        private Control BuildServiceNowPage()
        {
            var g = NewGrid();
            int r = 0;
            g.Controls.Add(snEnabled, 0, r); g.SetColumnSpan(snEnabled, 2); r++;
            g.Controls.Add(new Label { Text = "Instance URL:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(snUrl, 1, r); r++;
            g.Controls.Add(new Label { Text = "Username / Password:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            var up = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true };
            up.Controls.Add(snUser); up.Controls.Add(snPass);
            g.Controls.Add(up, 1, r); r++;
            g.Controls.Add(new Label { Text = "Table:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(snTable, 1, r); r++;
            g.Controls.Add(new Label { Text = "Query:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(snQuery, 1, r); r++;
            g.Controls.Add(new Label { Text = "Page Size:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(snPage, 1, r); r++;
            return g;
        }

        private Control BuildNessusPage()
        {
            var g = NewGrid();
            int r = 0;
            g.Controls.Add(neEnabled, 0, r); g.SetColumnSpan(neEnabled, 2); r++;
            g.Controls.Add(new Label { Text = "Base URL:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(neUrl, 1, r); r++;
            g.Controls.Add(new Label { Text = "Access Key / Secret Key:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            var ks = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true };
            ks.Controls.Add(neAccess); ks.Controls.Add(neSecret);
            g.Controls.Add(ks, 1, r); r++;
            g.Controls.Add(new Label { Text = "Timeout (sec):", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(neTimeout, 1, r); r++;
            return g;
        }

        private Control BuildAbsolutePage()
        {
            var g = NewGrid();
            int r = 0;
            g.Controls.Add(abEnabled, 0, r); g.SetColumnSpan(abEnabled, 2); r++;
            g.Controls.Add(new Label { Text = "Base URL:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(abUrl, 1, r); r++;
            g.Controls.Add(new Label { Text = "API Key (Bearer):", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(abKey, 1, r); r++;
            g.Controls.Add(new Label { Text = "Timeout (sec):", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(abTimeout, 1, r); r++;
            return g;
        }

        private Control BuildAdPage()
        {
            var g = NewGrid();
            int r = 0;
            g.Controls.Add(adEnabled, 0, r); g.SetColumnSpan(adEnabled, 2); r++;
            g.Controls.Add(new Label { Text = "LDAP Path:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(adPath, 1, r); r++;
            g.Controls.Add(new Label { Text = "Username / Password:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            var up = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true };
            up.Controls.Add(adUser); up.Controls.Add(adPass);
            g.Controls.Add(up, 1, r); r++;
            g.Controls.Add(new Label { Text = "Filter:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(adFilter, 1, r); r++;
            g.Controls.Add(new Label { Text = "Attributes (comma‑sep):", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(adAttrs, 1, r); r++;
            return g;
        }

        private Control BuildAzureAdPage()
        {
            var g = NewGrid();
            int r = 0;
            g.Controls.Add(aaEnabled, 0, r); g.SetColumnSpan(aaEnabled, 2); r++;
            g.Controls.Add(new Label { Text = "Tenant ID:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(aaTenant, 1, r); r++;
            g.Controls.Add(new Label { Text = "Client ID:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(aaClient, 1, r); r++;
            g.Controls.Add(new Label { Text = "Client Secret:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(aaSecret, 1, r); r++;
            g.Controls.Add(new Label { Text = "Scopes (space‑sep):", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(aaScopes, 1, r); r++;
            return g;
        }

        private Control BuildRapid7Page()
        {
            var g = NewGrid();
            int r = 0;
            g.Controls.Add(r7Enabled, 0, r); g.SetColumnSpan(r7Enabled, 2); r++;
            g.Controls.Add(new Label { Text = "Base URL:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(r7Url, 1, r); r++;
            g.Controls.Add(new Label { Text = "Username / Password:", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            var up = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true };
            up.Controls.Add(r7User); up.Controls.Add(r7Pass);
            g.Controls.Add(up, 1, r); r++;
            g.Controls.Add(new Label { Text = "API Key (optional):", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(r7Key, 1, r); r++;
            g.Controls.Add(new Label { Text = "Timeout (sec):", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, r);
            g.Controls.Add(r7Timeout, 1, r); r++;
            return g;
        }

        private TableLayoutPanel NewGrid()
        {
            var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 10, Padding = new Padding(12), AutoScroll = true };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            return grid;
        }

        private void LoadSettings()
        {
            var s = ApiSettingsStore.Load();
            // Global
            chkEnabled.Checked = s.Enabled;
            txtBaseUrl.Text = s.BaseUrl;
            cmbAuth.SelectedItem = s.AuthType;
            txtApiKey.Text = s.ApiKey;
            txtUser.Text = s.Username;
            txtPass.Text = s.Password;
            // Fixed: Handle nullable values in DefaultHeaders
            txtHeaders.Text = string.Join(Environment.NewLine, s.DefaultHeaders.Select(kv => kv.Key + ": " + (kv.Value ?? "")));
            txtSourceEndpoint.Text = s.SourceEndpoint;
            txtReportEndpoint.Text = s.ReportEndpoint;
            numTimeout.Value = s.TimeoutSeconds;

            // ServiceNow
            snEnabled.Checked = s.ServiceNow.Enabled;
            snUrl.Text = s.ServiceNow.InstanceUrl;
            snUser.Text = s.ServiceNow.Username;
            snPass.Text = s.ServiceNow.Password;
            snTable.Text = s.ServiceNow.Table;
            snQuery.Text = s.ServiceNow.Query;
            snPage.Value = s.ServiceNow.PageSize;

            // Nessus
            neEnabled.Checked = s.Nessus.Enabled;
            neUrl.Text = s.Nessus.BaseUrl;
            neAccess.Text = s.Nessus.AccessKey;
            neSecret.Text = s.Nessus.SecretKey;
            neTimeout.Value = s.Nessus.TimeoutSeconds;

            // Absolute
            abEnabled.Checked = s.Absolute.Enabled;
            abUrl.Text = s.Absolute.BaseUrl;
            abKey.Text = s.Absolute.ApiKey;
            abTimeout.Value = s.Absolute.TimeoutSeconds;

            // AD
            adEnabled.Checked = s.ActiveDirectory.Enabled;
            adPath.Text = s.ActiveDirectory.LdapPath;
            adUser.Text = s.ActiveDirectory.Username;
            adPass.Text = s.ActiveDirectory.Password;
            adFilter.Text = s.ActiveDirectory.Filter;
            adAttrs.Text = string.Join(",", s.ActiveDirectory.Attributes ?? Array.Empty<string>());

            // Azure AD
            aaEnabled.Checked = s.AzureAd.Enabled;
            aaTenant.Text = s.AzureAd.TenantId;
            aaClient.Text = s.AzureAd.ClientId;
            aaSecret.Text = s.AzureAd.ClientSecret;
            aaScopes.Text = string.Join(" ", s.AzureAd.Scopes ?? Array.Empty<string>());

            // Rapid7
            r7Enabled.Checked = s.Rapid7.Enabled;
            r7Url.Text = s.Rapid7.BaseUrl;
            r7User.Text = s.Rapid7.Username;
            r7Pass.Text = s.Rapid7.Password;
            r7Key.Text = s.Rapid7.ApiKey;
            r7Timeout.Value = s.Rapid7.TimeoutSeconds;
        }

        private void SaveSettings()
        {
            var s = new ApiSettings
            {
                // Global
                Enabled = chkEnabled.Checked,
                BaseUrl = txtBaseUrl.Text?.Trim() ?? "",
                AuthType = (ApiAuthType)(cmbAuth.SelectedItem ?? ApiAuthType.None),
                ApiKey = txtApiKey.Text ?? "",
                Username = txtUser.Text ?? "",
                Password = txtPass.Text ?? "",
                DefaultHeaders = ParseHeaders(txtHeaders.Text), // This now returns Dictionary<string, string?>
                SourceEndpoint = txtSourceEndpoint.Text?.Trim() ?? "",
                ReportEndpoint = txtReportEndpoint.Text?.Trim() ?? "",
                TimeoutSeconds = (int)numTimeout.Value,

                // Per-service
                ServiceNow = new ServiceNowSettings
                {
                    Enabled = snEnabled.Checked,
                    InstanceUrl = snUrl.Text?.Trim() ?? "",
                    Username = snUser.Text ?? "",
                    Password = snPass.Text ?? "",
                    Table = snTable.Text?.Trim() ?? "",
                    Query = snQuery.Text ?? "",
                    PageSize = (int)snPage.Value
                },
                Nessus = new NessusSettings
                {
                    Enabled = neEnabled.Checked,
                    BaseUrl = neUrl.Text?.Trim() ?? "",
                    AccessKey = neAccess.Text ?? "",
                    SecretKey = neSecret.Text ?? "",
                    TimeoutSeconds = (int)neTimeout.Value
                },
                Absolute = new AbsoluteSettings
                {
                    Enabled = abEnabled.Checked,
                    BaseUrl = abUrl.Text?.Trim() ?? "",
                    ApiKey = abKey.Text ?? "",
                    TimeoutSeconds = (int)abTimeout.Value
                },
                ActiveDirectory = new ActiveDirectorySettings
                {
                    Enabled = adEnabled.Checked,
                    LdapPath = adPath.Text?.Trim() ?? "",
                    Username = adUser.Text ?? "",
                    Password = adPass.Text ?? "",
                    Filter = adFilter.Text ?? "",
                    Attributes = (adAttrs.Text ?? "").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(s => s.Trim()).Where(s => s.Length > 0).ToArray()
                },
                AzureAd = new AzureAdSettings
                {
                    Enabled = aaEnabled.Checked,
                    TenantId = aaTenant.Text?.Trim() ?? "",
                    ClientId = aaClient.Text?.Trim() ?? "",
                    ClientSecret = aaSecret.Text ?? "",
                    Scopes = (aaScopes.Text ?? "").Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                },
                Rapid7 = new Rapid7Settings
                {
                    Enabled = r7Enabled.Checked,
                    BaseUrl = r7Url.Text?.Trim() ?? "",
                    Username = r7User.Text ?? "",
                    Password = r7Pass.Text ?? "",
                    ApiKey = r7Key.Text ?? "",
                    TimeoutSeconds = (int)r7Timeout.Value
                }
            };

            ApiSettingsStore.Save(s);
            DialogResult = DialogResult.OK;
        }

        // Updated to return Dictionary<string, string?> to match ApiSettings
        private static Dictionary<string, string?> ParseHeaders(string text)
        {
            var dict = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(text)) return dict;
            var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var raw in lines)
            {
                var line = raw.Trim();
                var idx = line.IndexOf(':');
                if (idx > 0)
                {
                    var k = line.Substring(0, idx).Trim();
                    var v = line.Substring(idx + 1).Trim();
                    if (!string.IsNullOrWhiteSpace(k)) dict[k] = v;
                }
            }
            return dict;
        }
    }
}
