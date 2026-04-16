using System.ComponentModel.DataAnnotations;

namespace Asset_Management.Models
{
    public class AssetCountListModel
    {
        [Key]
        public int id_count { get; set; }
        public string? asset_no { get; set; }
        public string? asset_subnumber { get; set; }
        public string? asset_desc { get; set; }
        public string? asset_class { get; set; }
        public string? asset_class_desc { get; set; }
        public string? cost_center { get; set; }
        public string? capitalized_on { get; set; }
        public string? department { get; set; }
        public string? count_year { get; set; }
        public int? count_month { get; set; }
        public string? count_month_name { get; set; }
        public int? is_counted { get; set; }
        public int? status_disposal { get; set; }
        public string? counted_by { get; set; }
        public string? date_counted { get; set; }
        public string? existence { get; set; }
        public string? good_condition { get; set; }
        public string? still_in_operation { get; set; }
        public string? tagging_available { get; set; }
        public string? applicable_of_tagging { get; set; }
        public string? correct_naming { get; set; }
        public string? correct_location { get; set; }
        public string? retagging { get; set; }
        public string? file_imgs { get; set; }
        public string? file_imgs_retagging { get; set; }
        public string? disposal_no { get; set; }
        public string? new_asset_name { get; set; }
        public string? new_asset_location { get; set; }
        public int? is_validated { get; set; }
        public string? recount_by { get; set; }
        public string? recount_remark { get; set; }
        public string? counted_name { get; set; }
        public double? curr_bk_val { get; set; }
        public string? status_remark { get; set; }
    }

    public class ExportAssetCountListModel
    {
        public string? schedule_year { get; set; }
        public string? schedule_month { get; set; }
        public string? count_date { get; set; }
        public string? validated { get; set; }
        public string? asset_no { get; set; }
        public string? asset_subnumber { get; set; }
        public string? asset_desc { get; set; }
        public string? asset_class { get; set; }
        public string? cost_center { get; set; }
        public string? capitalized_on { get; set; }
        public string? department { get; set; }
        public string? existence { get; set; }
        public string? good_condition { get; set; }
        public string? still_in_operation { get; set; }
        public string? tagging_available { get; set; }
        public string? applicable_of_tagging { get; set; }
        public string? correct_naming { get; set; }
        public string? correct_location { get; set; }
        public string? retagging { get; set; }
        public string? file_imgs { get; set; }
        public string? counted_by { get; set; }
        public string? action_remark { get; set; }
    }
}
