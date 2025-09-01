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
    /// ServiceNow REST client scaffold. Uses table API to fetch CI records.
    /// </summary>
    public class ServiceNowClient : ISourceProvider, IDisposable
    {
        private readonly ServiceNowSettings _cfg;
        private readonly HttpClient _http;

        public ServiceNowClient(ServiceNowSettings cfg)
        {
            _cfg = cfg;
            _http = new HttpClient { Timeout = TimeSpan.FromSeconds(Math.Max(5, cfg.PageSize > 0 ? 60 : 60)) };
            if (!string.IsNullOrWhiteSpace(cfg.InstanceUrl))
                _http.BaseAddress = new Uri(cfg.InstanceUrl.TrimEnd('/') + "/");
            if (!string.IsNullOrEmpty(cfg.Username))
            {
                var token = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes($"{cfg.Username}:{cfg.Password}"));
                _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", token);
            }
            _http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        public async Task<List<SourceTable>> FetchAsync(string assetClass, string dataPoint)
        {
            var list = new List<SourceTable>();
            if (!_cfg.Enabled) return list;
            if (string.IsNullOrWhiteSpace(_cfg.Table)) return list;

            // NOTE: Basic scaffold for future implementation
            // GET /api/now/table/{table}?sysparm_limit=...&sysparm_query=...
            var url = $"api/now/table/{_cfg.Table}?sysparm_limit={_cfg.PageSize}";
            if (!string.IsNullOrWhiteSpace(_cfg.Query)) url += $"&sysparm_query={Uri.EscapeDataString(_cfg.Query)}";

            try
            {
                var res = await _http.GetAsync(url).ConfigureAwait(false);
                if (!res.IsSuccessStatusCode) return list;
                using var s = await res.Content.ReadAsStreamAsync().ConfigureAwait(false);
                using var doc = await JsonDocument.ParseAsync(s).ConfigureAwait(false);

                var tbl = new SourceTable { DisplayName = "ServiceNow", FilePath = "servicenow://", Headers = new List<string>(), Rows = new List<Dictionary<string,string>>() };

                if (doc.RootElement.TryGetProperty("result", out var result) && result.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in result.EnumerateArray())
                    {
                        var row = new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var prop in item.EnumerateObject())
                        {
                            var v = prop.Value.ValueKind == JsonValueKind.String ? prop.Value.GetString() : prop.Value.ToString();
                            row[prop.Name] = v ?? "";
                            if (!tbl.Headers.Contains(prop.Name)) tbl.Headers.Add(prop.Name);
                        }
                        tbl.Rows.Add(row);
                    }
                }
                list.Add(tbl);
            }
            catch
            {
                // swallow in scaffold mode
            }

            return list;
        }

        public void Dispose() => _http?.Dispose();
    }
}
