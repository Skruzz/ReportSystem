using FinanceReport.Models;
using FinanceReport.Services.Interface;
using Microsoft.AspNetCore.Mvc;
using Microsoft.VisualBasic;
using OfficeOpenXml;
using Serilog;

namespace FinanceReport.Controllers
{
    [Route("api/report")]
    [ApiController]
    public class ExcelHandlerController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        private readonly ILogger<ExcelHandlerController> _logger;
        private readonly IExcelProcessingService _excelProcessingService;

        public ExcelHandlerController(ILogger<ExcelHandlerController> logger, IConfiguration configuration, IExcelProcessingService excelProcessingService)
        {
            _logger = logger;
            _configuration = configuration;
            _excelProcessingService = excelProcessingService;
        }

        [HttpPost("extractdata")]
        [ProducesResponseType(400)]
        [ProducesResponseType(500)]
        [ProducesResponseType(200)]
        public async Task<IActionResult> ProcessExcelData([FromBody] RequestModel request)
        {
            try
            {
                // Validate the request
                if (request == null)
                {
                    return BadRequest("Invalid request.");
                }

                // Load the Excel file
                string filePath = Path.Combine(Directory.GetCurrentDirectory(), _configuration.GetValue<string>("FileSettings:FilePath"));

                //get the final result
                var result = await _excelProcessingService.ProcessExcelData(request, filePath);

                // Return the result data
                return Ok(result);

            }
            catch (Exception ex)
            {
                // Handle exceptions
                _logger.LogError(ex, "An exception occurred");
                return StatusCode(StatusCodes.Status500InternalServerError, "An error occurred while processing your request.");
            }
        }


        [HttpPost("download-excel")]      
        [ProducesResponseType(400)]
        [ProducesResponseType(500)]
        [ProducesResponseType(200)]
        public async Task<IActionResult> DownloadExcelFile([FromBody] RequestModel request)
        {
            try
            {
                // Load the Excel file
                string filePath = Path.Combine(Directory.GetCurrentDirectory(), _configuration.GetValue<string>("FileSettings:FilePath"));

                // Assuming you already have the result data
                var result = await _excelProcessingService.ProcessExcelData(request, filePath);

                if (result == null || result.Count == 0)
                {
                    // Handle the case when there is no data to export
                    return BadRequest("No data to export.");
                }

                byte[] excelBytes = _excelProcessingService.GenerateExcelFile(result);
                if (excelBytes != null)
                {
                    return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "report.xlsx");
                }
                else
                {
                    // Handle the case when Excel generation fails
                    return StatusCode(StatusCodes.Status500InternalServerError, "Failed to generate the Excel file.");
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions
                _logger.LogError(ex, "An exception occurred");

                return StatusCode(StatusCodes.Status500InternalServerError, "An error occurred while processing your request.");
            }
        }
    }
}
