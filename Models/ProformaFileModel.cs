using System.ComponentModel.DataAnnotations;

namespace Asset_Management.Models
{
    public class ProformaFileModel
    {
        [Key]
        public int id_file { get; set; }
        public int id_gatepass { get; set; }
        public string? document_type { get; set; }
        public string? filename { get; set; }
        public string? created_by { get; set; }
        public DateTime record_date { get; set; }
    }
}