using System.ComponentModel.DataAnnotations;

namespace Asset_Management.Models
{
    public class DisposalListModel
    {
        [Key]
        public int? id_order { get; set; }
        public string? order_no { get; set; }
        public string? date_of_disposal { get; set; }
        public DateTime? date_of_disposal_history { get; set; }
        public string? dept_name { get; set; }
        public string? filename_doc { get; set; }
        public string? remark { get; set; }
        public string? requestor { get; set; }
        public string? requestor_name { get; set; }
        public string? date_requested { get; set; }
        public string? open_or_close { get; set; }
        public int? status_disposal { get; set; }
        public string? disposal_purpose { get; set; }
        public string? approval_hod { get; set; }
        public string? approval_hod_status { get; set; }
        public string? approval_hod_name { get; set; }
        public string? approval_fbp { get; set; }
        public string? approval_fbp_status { get; set; }
        public string? approval_fbp_name { get; set; }
        public string? approval_ph { get; set; }
        public string? approval_ph_status { get; set; }
        public string? approval_ph_name { get; set; }
        public string? approval_cfbp { get; set; }
        public string? approval_cfbp_status { get; set; }
        public string? approval_cfbp_name { get; set; }
        public string? approval_cvp { get; set; }
        public string? approval_cvp_status { get; set; }
        public string? approval_cvp_name { get; set; }
        public string? approval_fvp { get; set; }
        public string? approval_fvp_status { get; set; }
        public string? approval_fvp_name { get; set; }
        public string? approval_svp { get; set; }
        public string? approval_svp_status { get; set; }
        public string? approval_svp_name { get; set; }
        public string? approval_fsvp { get; set; }
        public string? approval_fsvp_status { get; set; }
        public string? approval_fsvp_name { get; set; }
        public string? status_desc { get; set; }
        public int is_bap_rejected { get; set; }
        public string? last_rejection_reason { get; set; }
        public string? filename_bap { get; set; }
        public int has_gatepass { get; set; }
        public double? aging_days { get; set; }
        public double? book_value_usd { get; set; }
        public double? gross_value_usd { get; set; }
        public double? book_value_eur { get; set; }
        public double? gross_value_eur { get; set; }
        public int? id_order_consol { get; set; }
        public double? consol_book_value_usd { get; set; }
        public double? consol_gross_value_usd { get; set; }
        public double? consol_book_value_eur { get; set; }
        public double? consol_gross_value_eur { get; set; }
        public string? consol_order_no { get; set; }
        public string? supporting_document { get; set; }
    }

    public class DisposalData
    {
        public string asset_no { get; set; }
        public string asset_subnumber { get; set; }
        public string category { get; set; }
        public string nbv_remark { get; set; } = "";
    }

    public class DisposalUploadResultModel
    {
        public int RowNumber { get; set; }
        public string AssetNo { get; set; }
        public string Subnumber { get; set; }
        public string Category { get; set; }
        public string Status { get; set; }
        public string ErrorMessage { get; set; }
        public bool RequireNbvRemark { get; set; }
        public string NbvRemark { get; set; }
        public double CurrentBookValue { get; set; }
        public int IsCounted { get; set; }
        public string CountYear { get; set; }
    }
}
