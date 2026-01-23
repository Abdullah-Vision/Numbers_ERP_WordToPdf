using Microsoft.EntityFrameworkCore;
using WordToPDF.Models;

namespace WordToPDF.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options) 
        {

        }

        public DbSet<Employee> Employees { get; set; }
        public DbSet<Education> Educations { get; set; }

        public DbSet<WordTemplate> WordTemplates { get; set; }


        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);


            modelBuilder.Entity<WordTemplate>(entity =>
            {
                entity.ToTable("WORDTEMPLATES"); // Oracle table name
                entity.HasKey(e => e.Id);

                entity.Property(e => e.Id).HasColumnName("ID");
                entity.Property(e => e.TemplateName).HasColumnName("TEMPLATENAME");
                entity.Property(e => e.FilePath).HasColumnName("FILEPATH");
            });

            modelBuilder.Entity<Employee>().HasData(
                new Employee
                {
                    Id = 1,
                    FullName = "Alice Wonderland",
                    DateOfBirth = new DateTime(1998, 05, 15),
                    Gender = "Female",
                    Address = "123 Rabbit Hole, Apt 4B",
                    City = "Wonderland City",
                    StateProvince = "Fantasy",
                    ZipPostalCode = "90210",
                    Country = "Wonderland",
                    PhoneNumber = "555-123-4567",
                    EmailAddress = "alice.w@example.com"
                },
                new Employee
                {
                    Id = 2,
                    FullName = "Bob The Builder",
                    DateOfBirth = new DateTime(1990, 11, 20),
                    Gender = "Male",
                    Address = "456 Construction Site",
                    City = "Toolsville",
                    StateProvince = "Workshire",
                    ZipPostalCode = "10001",
                    Country = "Buildland",
                    PhoneNumber = "555-987-6543",
                    EmailAddress = "bob.b@example.com"
                },
                new Employee
                {
                    Id = 3,
                    FullName = "Charlie Chaplin",
                    DateOfBirth = new DateTime(1889, 04, 16),
                    Gender = "Male",
                    Address = "789 Silent Film Studio",
                    City = "Hollywood",
                    StateProvince = "California",
                    ZipPostalCode = "90028",
                    Country = "USA",
                    PhoneNumber = "555-111-2222",
                    EmailAddress = "charlie.c@example.com"
                }
            );

            // Configure the relationship between Employee and Education
            modelBuilder.Entity<Employee>()
                .HasMany(e => e.EducationRecords)
                .WithOne(edu => edu.Employee)
                .HasForeignKey(edu => edu.EmployeeId);

            modelBuilder.Entity<EmployeeEducationDto>().HasNoKey().ToView(null);

            // Seed Education data
            modelBuilder.Entity<Education>().HasData(
               
                new Education { Id = 1, EmployeeId = 1, Degree = "B.Sc. in Fictional Studies", University = "Wonderland University", YearGraduated = 2020 },
                new Education { Id = 2, EmployeeId = 1, Degree = "M.A. in Logic & Illogic", University = "Looking Glass Institute", YearGraduated = 2022 },

               
                new Education { Id = 3, EmployeeId = 2, Degree = "Associate in Construction Tech", University = "Bob's Community College", YearGraduated = 2010 },
                new Education { Id = 4, EmployeeId = 2, Degree = "Certificate in Heavy Machinery", University = "Toolsville Polytechnic", YearGraduated = 2012 },

               
                new Education { Id = 5, EmployeeId = 3, Degree = "Diploma in Physical Comedy", University = "Slapstick Academy", YearGraduated = 1908 },
                new Education { Id = 6, EmployeeId = 3, Degree = "Honorary Doctorate of Humour", University = "Hollywood Arts College", YearGraduated = 1975 }
            );
        }
    }
}
