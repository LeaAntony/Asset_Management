using System.ComponentModel.DataAnnotations;

namespace Asset_Management.Models
{
    public class ReconAssetModel
    {
        [Key]
        public int? id_det { get; set; }
        public int? id_order { get; set; }
        public string? order_no { get; set; }
        public string? date_requested { get; set; }
        public int? status_disposal { get; set; }
        public string? status_desc { get; set; }
        public string? filename_doc { get; set; }
        public string? filename_bap { get; set; }
        public string? category { get; set; }
        public string? asset_no { get; set; }
        public string? asset_subnumber { get; set; }
        public string? asset_desc { get; set; }
        public string? asset_class { get; set; }
        public string? cost_center { get; set; }
        public string? capitalized_on { get; set; }
        public double? gross_value_usd { get; set; }
        public double? book_value_usd { get; set; }
        public double? gross_value_eur { get; set; }
        public double? book_value_eur { get; set; }
        public string? doc_no { get; set; }
        public string? doc_date { get; set; }
        public int? doc_month { get; set; }
        public string? doc_month_name { get; set; }
        public int? doc_year { get; set; }
        public int? doc_validated { get; set; }
    }
}
