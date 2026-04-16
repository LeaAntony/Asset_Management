using System.ComponentModel.DataAnnotations;

namespace Asset_Management.Models
{
    public class GatePassModel
    {
        [Key]
        public int? id_gatepass { get; set; }
        public string? gatepass_no { get; set; }
        public string? id_detail { get; set; }
        public string? category { get; set; }
        public string? type { get; set; }
        public string? employee_name { get; set; }
        public string? csr_to { get; set; }
        public string? return_date { get; set; }
        public string? vendor_code { get; set; }
        public string? vendor_name { get; set; }
        public string? vendor_address { get; set; }
        public string? recipient_phone { get; set; }
        public string? recipient_email { get; set; }
        public string? vehicle_no { get; set; }
        public string? driver_name { get; set; }
        public string? security_guard { get; set; }
        public string? remark { get; set; }
        public string? created_by { get; set; }
        public string? create_date { get; set; }
        public string? image_before { get; set; }
        public string? status_gatepass { get; set; }
        public string? status_remark { get; set; }
        public string? image_after { get; set; }
        public int? id_proforma { get; set; }
        public string? proforma_fin_status { get; set; }
        public string? shipping_status { get; set; }
        public string? supporting_documents { get; set; }
        public int? has_doc_no { get; set; }
        public int? has_overdue { get; set; }

        public string? asset_no { get; set; }
        public string? asset_subnumber { get; set; }
        public string? asset_desc { get; set; }
        public string? asset_class { get; set; }
        public string? cost_center { get; set; }
        public string? capitalized_on { get; set; }
        public decimal? hs_code { get; set; }

        public string? approver_no { get; set; }
        public string? approver { get; set; }
        public string? approval_status { get; set; }
        public string? date_approval { get; set; }

        public string? approval_hod { get; set; }
        public string? approval_hod_status { get; set; }
        public string? approval_fbp { get; set; }
        public string? approval_fbp_status { get; set; }
        public string? approval_ph { get; set; }
        public string? approval_ph_status { get; set; }
        public string? security_name { get; set; }
        public string? approval_security_status { get; set; }

        public string? approval_finance { get; set; }
        public string? approval_finance_status { get; set; }
        public string? approval_pic { get; set; }
        public string? approval_pic_status { get; set; }
        public string? requestor_name { get; set; }
        public string? confirmation_hod { get; set; }
        public string? confirmation_hod_status { get; set; }

        public string? new_pic { get; set; }
        public string? new_pic_name { get; set; }
        public string? naming_output { get; set; }
        public string? location { get; set; }
        public string Imgs { get; set; }
        public string? actual_return_date { get; set; }

        public string? proforma_no { get; set; }
        public string? attn_to { get; set; }
        public string? street { get; set; }
        public string? city { get; set; }
        public string? country { get; set; }
        public string? postal_code { get; set; }
        public string? phone_no { get; set; }
        public string? email { get; set; }
        public string? coo { get; set; }
        public string? ship_mode { get; set; }
        public string? courier_charges { get; set; }
        public string? courier_name { get; set; }
        public string? courier_account_no { get; set; }
        public string? freight_charges { get; set; }
        public string? incoterms { get; set; }
        public string? invoice_payment { get; set; }
        public string? file_attach { get; set; }
        public string? file_support { get; set; }
        public string? file_peb { get; set; }

        public string? plant { get; set; }
        public string? shipment_date { get; set; }
        public string? dhl_awb { get; set; }

        public string? box_no { get; set; }

        public int? id_file { get; set; }
        public string? document_type { get; set; }
        public string? fin_filename { get; set; }
        public string? fin_created_by { get; set; }
        public DateTime? fin_record_date { get; set; }

        public int? id_invoice { get; set; }
        public string? invoice_no { get; set; }
        public string? invoice_currency { get; set; }
        public DateTime? invoice_date { get; set; }
        public string? invoice_by { get; set; }
        public string? approval_invoice_status { get; set; }
        public string? approval_invoice_sesa { get; set; }
    }

    public class GatePassItem
    {
        public string asset_no { get; set; }
        public string asset_subnumber { get; set; }
        public string image_before { get; set; }
        public string image_after { get; set; }
    }

    public class GPOpenModel
    {
        public string? id_gatepass { get; set; }
        public string? gatepass_no { get; set; }
        public string? category { get; set; }
        public string? gatepass_date { get; set; }
        public string? requestor_name { get; set; }
        public string? return_date { get; set; }
        public string? destination { get; set; }
        public string? new_pic_name { get; set; }
        public string? status_gatepass { get; set; }
    }

    public class GPModel
    {
        public string? id_gatepass { get; set; }
        public string? gatepass_no { get; set; }
        public string? category { get; set; }
        public string? create_date { get; set; }
        public string? requestor_name { get; set; }
        public string? return_date { get; set; }
        public string? status_gatepass { get; set; }
        public string? department { get; set; }
        public string? aging_days { get; set; }
    }

    public class GatePassInvoiceAssetModel
    {
        public int id_detail { get; set; }
        public string? asset_no { get; set; }
        public string? asset_subnumber { get; set; }
        public string? asset_desc { get; set; }
        public string? asset_class { get; set; }
        public string? cost_center { get; set; }
        public string? capitalized_on { get; set; }
        public int? id_invoice_detail { get; set; }
        public decimal amount { get; set; }
    }
}