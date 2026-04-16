using System.ComponentModel.DataAnnotations;

namespace Asset_Management.Models
{
    public class RateEuroModel
    {
        [Key]
        public int id_exc { get; set; }
        public string? month_no { get; set; }
        public string? year_no { get; set; }
        public string? currency { get; set; }
        public double? exc_rate { get; set; }
    }
}
