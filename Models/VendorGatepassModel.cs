using System.ComponentModel.DataAnnotations;

namespace Asset_Management.Models
{
    public class VendorGatepassModel
    {
        [Key]
        public int? id_vendor { get; set; }
        public string? vendor_code { get; set; }
        public string? vendor_name { get; set; }
        public string? vendor_address { get; set; }
        public string? vendor_postal_code { get; set; }
        public string? vendor_location { get; set; }
        public string? vendor_batam { get; set; }
        public string? vendor_phone { get; set; }
        public string? vendor_email { get; set; }
        public DateTime? record_date { get; set; }
    }
}