using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using AssetDataValidationTool.Models;

namespace AssetDataValidationTool.Services
{
    /// <summary>
    /// Minimal HTTP client scaffold. Real implementation can be filled in later.
    /// </summary>
    public class HttpApiClient : IApiClient, IDisposable
    {
        private readonly ApiSettings _settings;
        private readonly HttpClient _http;

        public HttpApiClient(ApiSettings settings)
        {
            _settings = settings;
            var handler = new HttpClientHandler();
            _http = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(Math.Max(5, settings.TimeoutSeconds))
            };
            if (!string.IsNullOrWhiteSpace(settings.BaseUrl))
                _http.BaseAddress = new Uri(settings.BaseUrl, UriKind.Absolute);

            // auth
            switch (settings.AuthType)
            {
                case ApiAuthType.ApiKey:
                    if (!string.IsNullOrWhiteSpace(settings.ApiKey))
                        _http.DefaultRequestHeaders.Add("X-API-Key", settings.ApiKey);
                    break;
                case ApiAuthType.BearerToken:
                    if (!string.IsNullOrWhiteSpace(settings.ApiKey))
                        _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey);
                    break;
                case ApiAuthType.Basic:
                    if (!string.IsNullOrWhiteSpace(settings.Username))
                    {
                        var token = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes($"{settings.Username}:{settings.Password}"));
                        _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", token);
                    }
                    break;
                case ApiAuthType.None:
                default: break;
            }
            // you can add default headers later; left simple for scaffold
        }

        public async Task<bool> UploadSourceAsync(string label, string filePath)
        {
            if (!_settings.Enabled || string.IsNullOrWhiteSpace(_settings.SourceEndpoint)) return false;
            using var form = new MultipartFormDataContent();
            form.Add(new StringContent(label ?? string.Empty), "label");
            form.Add(new StreamContent(File.OpenRead(filePath)), "file", Path.GetFileName(filePath));
            var res = await _http.PostAsync(_settings.SourceEndpoint, form).ConfigureAwait(false);
            return res.IsSuccessStatusCode;
        }

        public async Task<bool> UploadReportAsync(string reportPath)
        {
            if (!_settings.Enabled || string.IsNullOrWhiteSpace(_settings.ReportEndpoint)) return false;
            using var form = new MultipartFormDataContent();
            form.Add(new StreamContent(File.OpenRead(reportPath)), "file", Path.GetFileName(reportPath));
            var res = await _http.PostAsync(_settings.ReportEndpoint, form).ConfigureAwait(false);
            return res.IsSuccessStatusCode;
        }

        public void Dispose() => _http?.Dispose();
    }
}
