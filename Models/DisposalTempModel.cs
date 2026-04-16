using System.ComponentModel.DataAnnotations;

namespace Asset_Management.Models
{
    public class DisposalTempModel
    {
        [Key]
        public int? id_temp { get; set; }
        public string? category { get; set; }
        public string? asset_no { get; set; }
        public string? asset_subnumber { get; set; }
        public string? asset_desc { get; set; }
        public string? asset_class { get; set; }
        public string? cost_center { get; set; }
        public string? capitalized_on { get; set; }
        public double? current_apc { get; set; }
        public double? curr_bk_val { get; set; }
        public string? nbv_remark { get; set; }
        public string? naming_output { get; set; }
        public string? gp_before { get; set; }
        public decimal? hs_code { get; set; }
        public int? is_validated { get; set; }

    }
}
