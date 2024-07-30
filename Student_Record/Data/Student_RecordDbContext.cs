using Microsoft.EntityFrameworkCore;
using Student_Record.Models;

namespace Student_Record.Data
{
    public class Student_RecordDbContext : DbContext
    {
        public Student_RecordDbContext()
        {
        }

        public Student_RecordDbContext(DbContextOptions options) : base(options)
        {
        }

        public DbSet<Students> Students { get; set; }
    }
}
