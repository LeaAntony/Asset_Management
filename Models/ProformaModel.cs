using System.ComponentModel.DataAnnotations;

namespace Asset_Management.Models
{
    public class ProformaModel
    {
        [Key]
        public int? id_proforma { get; set; }
        public string? proforma_no { get; set; }
        public string? attn_to { get; set; }
        public string? street { get; set; }
        public string? city { get; set; }
        public string? country { get; set; }
        public string? postal_code { get; set; }
        public string? phone_no { get; set; }
        public string? email { get; set; }
        public string? file_attach { get; set; }
        public string? file_support { get; set; }
        public string? ship_mode { get; set; }
        public string? courier_charges { get; set; }
        public string? courier_name { get; set; }
        public string? courier_account_no { get; set; }
        public string? freight_charges { get; set; }
        public string? invoice_payment { get; set; }
        public string? requested_by { get; set; }
        public DateTime? record_date { get; set; }
        public int? id_gatepass { get; set; }
        public string? gatepass_no { get; set; }
    }
}