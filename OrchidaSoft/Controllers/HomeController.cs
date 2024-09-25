using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OrchidaSoft.Models;
using OrchidaSoft.Services;
using System.Diagnostics;

namespace OrchidaSoft.Controllers
{
    public class HomeController : Controller
    {
        private readonly ExcelService _excelService;

        public HomeController(ExcelService excelService)
        {
            _excelService = excelService;
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }
        [HttpPost("upload")]
        public async Task<IActionResult> Upload(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }

            if (!file.FileName.EndsWith(".xlsx"))
            {
                return BadRequest("Invalid file format. Please upload an Excel file.");
            }

            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                stream.Position = 0; // Reset stream position for reading

                var resultStream = await _excelService.ProcessExcelAsync(stream);
                var fileName = $"Modified_Taxes_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                return File(resultStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }
    }


}

