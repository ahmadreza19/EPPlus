using Microsoft.EntityFrameworkCore;

namespace Test_Project.Models
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
        }

        public DbSet<DataModel> DataModels { get; set; }
    }
   

}
