using System.Threading.Tasks;

namespace AssetDataValidationTool.Services
{
    /// <summary>
    /// Abstraction for pushing generated outputs to a destination (e.g., ticketing, storage).
    /// </summary>
    public interface IReportSink
    {
        Task<bool> UploadReportAsync(string reportPath);
    }
}
