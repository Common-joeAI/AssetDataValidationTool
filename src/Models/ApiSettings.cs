// ApiSettings.cs — null-friendly headers and sanitizer
using System;
using System.Collections.Generic;
using System.Linq;

namespace AssetDataValidationTool.Models
{
    public enum ApiAuthType { None, ApiKey, BearerToken, Basic }

    public class ApiSettings
    {
        public bool Enabled { get; set; } = false;
        public string BaseUrl { get; set; } = string.Empty;
        public ApiAuthType AuthType { get; set; } = ApiAuthType.None;
        public string ApiKey { get; set; } = string.Empty;
        public string Username { get; set; } = string.Empty;
        public string Password { get; set; } = string.Empty;

        // Allow nullable values from config; sanitize on load before use
        public Dictionary<string, string?> DefaultHeaders { get; set; } = new(StringComparer.OrdinalIgnoreCase);

        public string SourceEndpoint { get; set; } = string.Empty;
        public string ReportEndpoint { get; set; } = string.Empty;
        public int TimeoutSeconds { get; set; } = 60;
        public bool ValidateServerCertificate { get; set; } = true;

        public ServiceNowSettings ServiceNow { get; set; } = new();
        public NessusSettings Nessus { get; set; } = new();
        public AbsoluteSettings Absolute { get; set; } = new();
        public ActiveDirectorySettings ActiveDirectory { get; set; } = new();
        public AzureAdSettings AzureAd { get; set; } = new();
        public Rapid7Settings Rapid7 { get; set; } = new();

        public void Sanitize()
        {
            if (DefaultHeaders != null)
            {
                // Create a new dictionary with non-nullable string values
                var sanitized = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
                foreach (var kv in DefaultHeaders)
                {
                    sanitized[kv.Key] = kv.Value ?? string.Empty;
                }
                DefaultHeaders = sanitized;
            }
        }
    }

    public class ServiceNowSettings
    {
        public bool Enabled { get; set; } = false;
        public string InstanceUrl { get; set; } = string.Empty;
        public string Username { get; set; } = string.Empty;
        public string Password { get; set; } = string.Empty;
        public string Table { get; set; } = "cmdb_ci_computer";
        public string Query { get; set; } = string.Empty;
        public int PageSize { get; set; } = 200;
    }

    public class NessusSettings
    {
        public bool Enabled { get; set; } = false;
        public string BaseUrl { get; set; } = string.Empty;
        public string AccessKey { get; set; } = string.Empty;
        public string SecretKey { get; set; } = string.Empty;
        public int TimeoutSeconds { get; set; } = 60;
    }

    public class AbsoluteSettings
    {
        public bool Enabled { get; set; } = false;
        public string BaseUrl { get; set; } = string.Empty;
        public string ApiKey { get; set; } = string.Empty;
        public int TimeoutSeconds { get; set; } = 60;
    }

    public class ActiveDirectorySettings
    {
        public bool Enabled { get; set; } = false;
        public string LdapPath { get; set; } = "LDAP://dc=example,dc=com";
        public string Username { get; set; } = string.Empty;
        public string Password { get; set; } = string.Empty;
        public string Filter { get; set; } = "(objectClass=computer)";
        public string[] Attributes { get; set; } = new[] { "cn", "dNSHostName", "operatingSystem" };
    }

    public class AzureAdSettings
    {
        public bool Enabled { get; set; } = false;
        public string TenantId { get; set; } = string.Empty;
        public string ClientId { get; set; } = string.Empty;
        public string ClientSecret { get; set; } = string.Empty;
        public string[] Scopes { get; set; } = new[] { "https://graph.microsoft.com/.default" };
    }

    public class Rapid7Settings
    {
        public bool Enabled { get; set; } = false;
        public string BaseUrl { get; set; } = string.Empty;
        public string Username { get; set; } = string.Empty;
        public string Password { get; set; } = string.Empty;
        public string ApiKey { get; set; } = string.Empty;
        public int TimeoutSeconds { get; set; } = 60;
    }
}
