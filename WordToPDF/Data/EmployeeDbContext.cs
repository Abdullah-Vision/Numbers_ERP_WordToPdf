using Microsoft.EntityFrameworkCore;
using WordToPDF.Models;

namespace WordToPDF.Data
{
    // SQL Server context - serves the PDF converter flow only.
    // Tables: Employee, Education (1-to-many) and WordTemplate in EmployeeDB.
    // The Oracle ApplicationDbContext remains untouched for all other flows.
    public class EmployeeDbContext : DbContext
    {
        public EmployeeDbContext(DbContextOptions<EmployeeDbContext> options) : base(options)
        {

        }

        public DbSet<Employee> Employees { get; set; }
        public DbSet<Education> Educations { get; set; }
        public DbSet<WordTemplate> WordTemplates { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<Employee>(entity =>
            {
                entity.ToTable("Employee"); // SQL Server table name
                entity.HasKey(e => e.Id);
            });

            modelBuilder.Entity<Education>(entity =>
            {
                entity.ToTable("Education"); // SQL Server table name
                entity.HasKey(e => e.Id);

                // One-to-many: Employee -> Education
                entity.HasOne(edu => edu.Employee)
                      .WithMany(e => e.EducationRecords)
                      .HasForeignKey(edu => edu.EmployeeId);
            });

            modelBuilder.Entity<WordTemplate>(entity =>
            {
                entity.ToTable("WordTemplate"); // SQL Server table name
                entity.HasKey(e => e.Id);
            });

            // Keyless shape used by the USP_GetEmployeeWithEducationFlat stored procedure
            modelBuilder.Entity<EmployeeEducationDto>().HasNoKey().ToView(null);
        }
    }
}
