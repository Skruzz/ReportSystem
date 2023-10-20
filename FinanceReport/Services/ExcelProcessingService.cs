using FinanceReport.Models;
using FinanceReport.Services.Interface;
using Microsoft.Extensions.Caching.Memory;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace FinanceReport.Services
{
    public class ExcelProcessingService : IExcelProcessingService
    {
        private readonly ILogger<ExcelProcessingService> _logger;
        private readonly IMemoryCache _memoryCache;

        public ExcelProcessingService(ILogger<ExcelProcessingService> logger, IMemoryCache cache)
        {
            _logger = logger;
            _memoryCache = cache;
        }

        public async Task<List<Dictionary<string, string>>> ProcessExcelData(RequestModel request, string filePath)
        {
            try
            {
                string cacheKey = $"ExcelData_{request.WorksheetName}_{filePath}";
                string cacheFieldCountKey = $"ExcelData_FieldCount";

                if (_memoryCache.TryGetValue(cacheKey, out List<Dictionary<string, string>> cachedResult))
                {
                    //Check if field count same
                    if(_memoryCache.TryGetValue(cacheFieldCountKey, out int cachedCount))
                    {
                        if(cachedCount == request.FieldMappings.Count)
                        {
                            // If the result is in the cache, return it                    
                            return cachedResult;
                        }
                    }               
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                string worksheetName = request.WorksheetName;
                List<FieldMapping> fieldMappings = request.FieldMappings;

                var result = new List<Dictionary<string, string>>();
                var dataLock = new object(); // Lock for thread-safe list modification                

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var workbook = package.Workbook;
                    var worksheet = workbook.Worksheets[worksheetName];

                    if (worksheet == null)
                    {
                        _logger.LogError(null, "Worksheet not found.");
                        throw new Exception("Worksheet not found.");
                    }

                    int colStart = 6;

                    await Task.WhenAll(Enumerable.Range(colStart, worksheet.Dimension.End.Column + 1)
                   .Select(async colNumber =>
                   {
                       var companyData = new Dictionary<string, string>();
                       var cellCompany = worksheet.Cells[2, colNumber]; // Assuming the company names are in row 2 as per Excel format
                       string companyName = cellCompany.Text.Trim();

                       if (!string.IsNullOrEmpty(companyName))
                       {
                           companyData["companyName"] = companyName;

                           foreach (var mapping in fieldMappings)
                           {
                               try
                               {
                                   var cell = worksheet.Cells[mapping.RowNumber, colNumber];
                                   string formattedCellValue = (cell.Text.Trim() == "$-") ? "" : cell.Text;
                                   companyData[mapping.FieldName.ToLowerInvariant()] = formattedCellValue;
                               }
                               catch (Exception ex)
                               {
                                   _logger.LogError(ex, "An exception occurred");
                               }
                           }

                           lock (dataLock)
                           {
                               result.Add(companyData);
                           }
                       }
                   }));

                }
                // Sort the list by 'companyName' key
                result = result.OrderBy(dict => dict["companyName"]).ToList();

                // Store the result in the cache for future use for eg. 30 min
                // we can adjust the cache duration as needed
                _memoryCache.Set(cacheKey, result, TimeSpan.FromMinutes(30));
                _memoryCache.Set(cacheFieldCountKey, request.FieldMappings.Count, TimeSpan.FromMinutes(30));

                return result;

            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "An exception occurred");
                throw;
            }
        }

        public byte[] GenerateExcelFile(List<Dictionary<string, string>> result)
        {
            try
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Results");

                    // Create headers
                    int headerRow = 1;
                    foreach (var key in result[0].Keys)
                    {
                        worksheet.Cells[headerRow, result[0].Keys.ToList().IndexOf(key) + 1].Value = key;
                        worksheet.Cells[headerRow, result[0].Keys.ToList().IndexOf(key) + 1].Style.Font.Bold = true;
                        worksheet.Cells[headerRow, result[0].Keys.ToList().IndexOf(key) + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[headerRow, result[0].Keys.ToList().IndexOf(key) + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[headerRow, result[0].Keys.ToList().IndexOf(key) + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[headerRow, result[0].Keys.ToList().IndexOf(key) + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    }

                    // Populate data
                    int dataRow = 2;
                    foreach (var companyData in result)
                    {
                        foreach (var key in companyData.Keys)
                        {
                            worksheet.Cells[dataRow, companyData.Keys.ToList().IndexOf(key) + 1].Value = companyData[key];
                            worksheet.Cells[dataRow, companyData.Keys.ToList().IndexOf(key) + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            worksheet.Cells[dataRow, companyData.Keys.ToList().IndexOf(key) + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }
                        dataRow++;
                    }

                    // Auto-size columns
                    worksheet.Cells.AutoFitColumns();

                    // Save the Excel file
                    return package.GetAsByteArray();
                }
            }
            catch (Exception ex)
            {

                _logger.LogError(ex, "An exception occurred");
                throw;
            }
        }

    }
}
