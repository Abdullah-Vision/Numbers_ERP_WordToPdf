using System.ComponentModel.DataAnnotations.Schema;

namespace WordToPDF.Models
{
    public class Education
    {
        public int Id { get; set; }
        public int EmployeeId { get; set; } 
        public string Degree { get; set; }
        public string University { get; set; }
        public int YearGraduated { get; set; }

      
        [ForeignKey("EmployeeId")]
        public Employee Employee { get; set; }
    }
}