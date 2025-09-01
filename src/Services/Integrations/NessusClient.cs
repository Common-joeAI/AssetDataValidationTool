using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using AssetDataValidationTool.Models;

namespace AssetDataValidationTool.Services.Integrations
{
    /// <summary>
    /// Nessus API scaffold (Tenable). Supports API key headers.
    /// </summary>
    public class NessusClient : ISourceProvider, IDisposable
    {
        private readonly NessusSettings _cfg;
        private readonly HttpClient _http;

        public NessusClient(NessusSettings cfg)
        {
            _cfg = cfg;
            var handler = new HttpClientHandler();
            _http = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(Math.Max(5, cfg.TimeoutSeconds)) };
            if (!string.IsNullOrWhiteSpace(cfg.BaseUrl))
                _http.BaseAddress = new Uri(cfg.BaseUrl.TrimEnd('/') + "/");
            // Tenable X-ApiKeys: accessKey=...; secretKey=...
            if (!string.IsNullOrWhiteSpace(cfg.AccessKey) && !string.IsNullOrWhiteSpace(cfg.SecretKey))
            {
                _http.DefaultRequestHeaders.Add("X-ApiKeys", $"accessKey={cfg.AccessKey}; secretKey={cfg.SecretKey}");
            }
            _http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        public Task<List<SourceTable>> FetchAsync(string assetClass, string dataPoint)
        {
            var list = new List<SourceTable>();
            if (!_cfg.Enabled) return Task.FromResult(list);
            // Scaffold: call an assets endpoint in future; return empty for now
            return Task.FromResult(list);
        }

        public void Dispose() => _http?.Dispose();
    }
}
