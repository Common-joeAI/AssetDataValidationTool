using System.Threading.Tasks;

namespace AssetDataValidationTool.Services
{
    public interface IApiClient
    {
        Task<bool> UploadSourceAsync(string label, string filePath);
        Task<bool> UploadReportAsync(string reportPath);
    }
}
