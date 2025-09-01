using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.DirectoryServices;
using AssetDataValidationTool.Models;

namespace AssetDataValidationTool.Services.Integrations
{
    /// <summary>
    /// On‑prem Active Directory (LDAP) scaffold. Requires Windows + DirectoryServices.
    /// </summary>
    public class ActiveDirectoryClient : ISourceProvider
    {
        private readonly ActiveDirectorySettings _cfg;

        public ActiveDirectoryClient(ActiveDirectorySettings cfg) { _cfg = cfg; }

        public Task<List<SourceTable>> FetchAsync(string assetClass, string dataPoint)
        {
            var list = new List<SourceTable>();
            if (!_cfg.Enabled) return Task.FromResult(list);

            try
            {
                using var entry = new DirectoryEntry(_cfg.LdapPath, _cfg.Username, _cfg.Password);
                using var searcher = new DirectorySearcher(entry) { Filter = _cfg.Filter, PageSize = 500 };
                if (_cfg.Attributes != null)
                {
                    searcher.PropertiesToLoad.Clear();
                    foreach (var a in _cfg.Attributes) searcher.PropertiesToLoad.Add(a);
                }

                var tbl = new SourceTable { DisplayName = "Active Directory", FilePath = "ldap://", Headers = new List<string>(), Rows = new List<Dictionary<string,string>>() };
                using var results = searcher.FindAll();
                foreach (SearchResult r in results)
                {
                    var row = new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
                    foreach (string prop in r.Properties.PropertyNames)
                    {
                        var vals = r.Properties[prop];
                        var s = vals != null && vals.Count > 0 ? Convert.ToString(vals[0]) : "";
                        row[prop] = s ?? "";
                        if (!tbl.Headers.Contains(prop)) tbl.Headers.Add(prop);
                    }
                    tbl.Rows.Add(row);
                }
                list.Add(tbl);
            }
            catch
            {
                // suppress in scaffold
            }

            return Task.FromResult(list);
        }
    }
}
