using System.ComponentModel.DataAnnotations;

namespace Asset_Management.Models
{
    public class AUCDetailModel
    {
        [Key]
        public int? id_auc { get; set; }
        public int? id_asset { get; set; }
        public string? asset_no { get; set; }
        public string? asset_subnumber { get; set; }
        public string? asset_desc { get; set; }
        public string? cost_center { get; set; }
        public string? plant { get; set; }
        public string? capitalized_on { get; set; }
        public string? sesa_owner { get; set; }
        public string? name { get; set; }
        public string? department { get; set; }
        public string? current_apc { get; set; }
        public string? aging_days { get; set; }
        public string? est_date_proj { get; set; }
        public string? status_po { get; set; }
        public string? status_project { get; set; }
        public string? validated_by { get; set; }
        public string? validated_remark { get; set; }
        public string? auc_category { get; set; }
        public string? approval_fbp_by { get; set; }
        public string? approval_fbp_status { get; set; }
        public string? date_approval_fbp { get; set; }
        public string? reclass_asset { get; set; }
        public double? summ_curr_apc { get; set; }
    }
}
