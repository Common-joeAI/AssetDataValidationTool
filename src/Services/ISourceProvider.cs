using System.Collections.Generic;
using System.Threading.Tasks;
using AssetDataValidationTool.Models;

namespace AssetDataValidationTool.Services
{
    /// <summary>
    /// Abstraction for pulling source data programmatically (alternative to local files).
    /// Implementations should shape data into SourceTable(s).
    /// </summary>
    public interface ISourceProvider
    {
        Task<List<SourceTable>> FetchAsync(string assetClass, string dataPoint);
    }
}
