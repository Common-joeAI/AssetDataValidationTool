using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading.Tasks;
using AssetDataValidationTool.Models;

namespace AssetDataValidationTool.Services.Integrations
{
    /// <summary>
    /// Azure AD (Entra ID) via Microsoft Graph scaffold.
    /// NOTE: In a future iteration, use MSAL to get a token and set Bearer headers.
    /// </summary>
    public class AzureAdClient : ISourceProvider, IDisposable
    {
        private readonly AzureAdSettings _cfg;
        private readonly HttpClient _http;

        public AzureAdClient(AzureAdSettings cfg)
        {
            _cfg = cfg;
            _http = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
            _http.BaseAddress = new Uri("https://graph.microsoft.com/");
        }

        public async Task<List<SourceTable>> FetchAsync(string assetClass, string dataPoint)
        {
            var list = new List<SourceTable>();
            if (!_cfg.Enabled) return list;

            // TODO: acquire token with MSAL and set Authorization header
            // _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            // Scaffold GET: devices
            try
            {
                var res = await _http.GetAsync("v1.0/devices?$top=20").ConfigureAwait(false);
                if (!res.IsSuccessStatusCode) return list;
                var stream = await res.Content.ReadAsStreamAsync().ConfigureAwait(false);
                using var doc = await JsonDocument.ParseAsync(stream).ConfigureAwait(false);

                var tbl = new SourceTable { DisplayName = "Azure AD", FilePath = "graph://devices", Headers = new List<string>(), Rows = new List<Dictionary<string,string>>() };
                if (doc.RootElement.TryGetProperty("value", out var arr) && arr.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in arr.EnumerateArray())
                    {
                        var row = new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var prop in item.EnumerateObject())
                        {
                            row[prop.Name] = prop.Value.ValueKind == JsonValueKind.String ? prop.Value.GetString() ?? "" : prop.Value.ToString();
                            if (!tbl.Headers.Contains(prop.Name)) tbl.Headers.Add(prop.Name);
                        }
                        tbl.Rows.Add(row);
                    }
                }
                list.Add(tbl);
            }
            catch
            {
                // ignore in scaffold
            }

            return list;
        }

        public void Dispose() => _http?.Dispose();
    }
}
