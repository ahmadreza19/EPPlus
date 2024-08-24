using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using Quartz;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Test_Project.Models;

namespace Test_Project.Models
{
    public class ExcelImportJob : IJob
    {
        private readonly IServiceProvider _serviceProvider;
        private readonly IConfiguration _configuration;
        private readonly ILogger<ExcelImportJob> _logger;

        public ExcelImportJob(IServiceProvider serviceProvider, IConfiguration configuration, ILogger<ExcelImportJob> logger)
        {
            _serviceProvider = serviceProvider;
            _configuration = configuration;
            _logger = logger;
        }

        public async Task Execute(IJobExecutionContext context)
        {
            var folderPath = _configuration.GetValue<string>("ExcelFolderPath");
            var fileName = _configuration.GetValue<string>("ExcelFileName");
            var filePath = Path.Combine(folderPath, fileName);

            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                _logger.LogError("فایل اکسل پیدا نشد.");
                return;
            }

            using (var scope = _serviceProvider.CreateScope())
            {
                var dbContext = scope.ServiceProvider.GetRequiredService<ApplicationDbContext>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var package = new ExcelPackage(stream))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        var rowCount = worksheet.Dimension.Rows;

                        for (int row = 2; row <= rowCount; row++) // شروع از ردیف دوم فرض بر اینکه اولین ردیف هدر است
                        {
                            var column1 = worksheet.Cells[row, 1].Text;
                            var column2 = worksheet.Cells[row, 2].Text;

                            // بررسی تکراری بودن رکورد
                            bool exists = await dbContext.DataModels
                                .AnyAsync(e => e.Column1 == column1 && e.Column2 == column2);

                            if (!exists)
                            {
                                var data = new DataModel
                                {
                                    Column1 = column1,
                                    Column2 = column2,
                                    // سایر ستون‌ها بر اساس ساختار اکسل شما
                                };
                                dbContext.DataModels.Add(data);
                            }
                        }

                        await dbContext.SaveChangesAsync();
                    }
                }
            }

            _logger.LogInformation("اطلاعات اکسل با موفقیت وارد پایگاه داده شد.");
        }
    }

 


}
