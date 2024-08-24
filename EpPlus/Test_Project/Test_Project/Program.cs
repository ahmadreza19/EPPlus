using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using Quartz;
using Quartz.Spi;
using Test_Project.Models;

var builder = WebApplication.CreateBuilder(args);

// Set the ExcelPackage license context.
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

// Add services to the container.
builder.Services.AddControllersWithViews();

// Configure DbContext with SQL Server.
builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));

// Add IConfiguration to DI.
builder.Services.AddSingleton<IConfiguration>(builder.Configuration);

// Configure Quartz.NET
builder.Services.AddQuartz(q =>
{
    q.UseMicrosoftDependencyInjectionJobFactory(); // استفاده از متد صحیح

    // Create a job and trigger
    var jobKey = new JobKey("ExcelImportJob");
    q.AddJob<ExcelImportJob>(opts => opts.WithIdentity(jobKey));

    q.AddTrigger(opts => opts
     .ForJob(jobKey)
     .WithIdentity("ExcelImportJob-trigger")
     .WithCronSchedule("0 * * * * ?")); // تنظیم برای اجرا هر 5 دقیقه یک بار
});



builder.Services.AddQuartzHostedService(q => q.WaitForJobsToComplete = true);

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();
