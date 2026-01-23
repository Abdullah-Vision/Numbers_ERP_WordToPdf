using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace WordToPDF.Models
{
    public class WordTemplate
    {
        public int Id { get; set; }

        [Required]
        [MaxLength(255)]
        public string TemplateName { get; set; }

        [Required]
        [MaxLength(500)]
        public string FilePath { get; set; } 
    }
}