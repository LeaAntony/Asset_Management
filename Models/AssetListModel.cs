using System.ComponentModel.DataAnnotations;

namespace Asset_Management.Models
{
    public class AssetListModel
    {
        [Key]
        public int? id_asset { get; set; }
        public string? asset_no { get; set; }
        public string? asset_subnumber { get; set; }
        public string? asset_desc { get; set; }
        public string? asset_class { get; set; }
        public string? asset_class_desc { get; set; }
        public string? cost_center { get; set; }
        public string? cc_desc { get; set; }
        public string? cc_grouping { get; set; }
        public string? cc_plant { get; set; }
        public string? capitalized_on { get; set; }
        public double? apc_fy_start { get; set; }
        public double? acquisition { get; set; }
        public double? retirement { get; set; }
        public double? transfer { get; set; }
        public double? current_apc { get; set; }
        public double? dep_fy_start { get; set; }
        public double? dep_for_year { get; set; }
        public double? dep_retir { get; set; }
        public double? dep_transfer { get; set; }
        public double? accumul_dep { get; set; }
        public double? bk_val_fy { get; set; }
        public double? curr_bk_val { get; set; }
        public string? currency { get; set; }
        public string? department { get; set; }
        public string? plant { get; set; }
        public string? sesa_owner { get; set; }
        public string? name_owner { get; set; }
        public string? tagging_status { get; set; }
        public int? aging_days_tagging { get; set; }
        public string? vendor_name { get; set; }
        public string? file_tag_status { get; set; }
        public string? file_tag { get; set; }
        public string? tag_validated { get; set; }
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
        public string? file_count { get; set; }
        public string? disposal_no { get; set; }
        public string? new_asset_name { get; set; }
        public string? new_asset_location { get; set; }
        public string? status_gatepass { get; set; }
        public string? aging_days { get; set; }
        public int? is_validated { get; set; }
        public string? owner { get; set; }
        public string? asset_location { get; set; }
        public string? asset_location_address { get; set; }
        public string? filename_doc { get; set; }

    }
}
