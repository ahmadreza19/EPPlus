using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Test_Project.Models;

namespace Test_Project.Controllers
{
    public class DataController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly IConfiguration _configuration;

        public DataController(ApplicationDbContext context, IConfiguration configuration)
        {
            _context = context;
            _configuration = configuration;
        }

       
        [HttpPost]
        public async Task<IActionResult> Import()
        {
            var folderPath = _configuration.GetValue<string>("ExcelFolderPath");
            var fileName = _configuration.GetValue<string>("ExcelFileName");
            var filePath = Path.Combine(folderPath, fileName);

            if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
                return Content("فایل اکسل پیدا نشد.");

            // Set the license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++) // شروع از ردیف دوم فرض بر اینکه اولین ردیف هدر است
                    {
                        var data = new DataModel
                        {
                            Column1 = worksheet.Cells[row, 1].Text,
                            Column2 = worksheet.Cells[row, 2].Text,
                            // سایر ستون‌ها بر اساس ساختار اکسل شما
                        };
                        _context.DataModels.Add(data);
                    }

                    await _context.SaveChangesAsync();
                }
            }

            return RedirectToAction("Index");
        }


        public IActionResult Index()
        {
            return View(_context.DataModels.ToList());
        }
    }
}

