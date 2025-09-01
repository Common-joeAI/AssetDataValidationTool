using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using AssetDataValidationTool.Models;

namespace AssetDataValidationTool.Services.Integrations
{
    /// <summary>
    /// Rapid7 InsightVM / Nexpose scaffold.
    /// </summary>
    public class Rapid7Client : ISourceProvider, IDisposable
    {
        private readonly Rapid7Settings _cfg;
        private readonly HttpClient _http;

        public Rapid7Client(Rapid7Settings cfg)
        {
            _cfg = cfg;
            _http = new HttpClient { Timeout = TimeSpan.FromSeconds(Math.Max(5, cfg.TimeoutSeconds)) };
            if (!string.IsNullOrWhiteSpace(cfg.BaseUrl))
                _http.BaseAddress = new Uri(cfg.BaseUrl.TrimEnd('/') + "/");

            if (!string.IsNullOrWhiteSpace(cfg.ApiKey))
                _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", cfg.ApiKey);
            else if (!string.IsNullOrWhiteSpace(cfg.Username))
            {
                var token = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes($"{cfg.Username}:{cfg.Password}"));
                _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", token);
            }
            _http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        public Task<List<SourceTable>> FetchAsync(string assetClass, string dataPoint)
        {
            var list = new List<SourceTable>();
            if (!_cfg.Enabled) return Task.FromResult(list);
            // Scaffold: implement assets endpoint in future
            return Task.FromResult(list);
        }

        public void Dispose() => _http?.Dispose();
    }
}
