using FinanceReport.Models;

namespace FinanceReport.Services.Interface
{
    public interface IExcelProcessingService
    {
        public Task<List<Dictionary<string, string>>> ProcessExcelData(RequestModel request, string filePath);

        public byte[] GenerateExcelFile(List<Dictionary<string, string>> result);
    }
}
