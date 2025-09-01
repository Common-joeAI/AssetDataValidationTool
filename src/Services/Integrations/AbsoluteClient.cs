using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using AssetDataValidationTool.Models;

namespace AssetDataValidationTool.Services.Integrations
{
    /// <summary>
    /// Absolute (MDM/endpoint) API scaffold.
    /// </summary>
    public class AbsoluteClient : ISourceProvider, IDisposable
    {
        private readonly AbsoluteSettings _cfg;
        private readonly HttpClient _http;

        public AbsoluteClient(AbsoluteSettings cfg)
        {
            _cfg = cfg;
            _http = new HttpClient { Timeout = TimeSpan.FromSeconds(Math.Max(5, cfg.TimeoutSeconds)) };
            if (!string.IsNullOrWhiteSpace(cfg.BaseUrl))
                _http.BaseAddress = new Uri(cfg.BaseUrl.TrimEnd('/') + "/");
            if (!string.IsNullOrWhiteSpace(cfg.ApiKey))
                _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", cfg.ApiKey);
            _http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        public Task<List<SourceTable>> FetchAsync(string assetClass, string dataPoint)
        {
            var list = new List<SourceTable>();
            if (!_cfg.Enabled) return Task.FromResult(list);
            // Scaffold: implement endpoints as needed
            return Task.FromResult(list);
        }

        public void Dispose() => _http?.Dispose();
    }
}
