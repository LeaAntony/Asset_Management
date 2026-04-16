using System.ComponentModel.DataAnnotations;

namespace Asset_Management.Models
{
    public class InvoiceExportModel
    {
        public InvoiceHeaderModel? Header { get; set; }
        public List<InvoiceDetailModel>? Details { get; set; }
    }

    public class InvoiceHeaderModel
    {
        public string? invoice_no { get; set; }
        public DateTime invoice_date { get; set; }
        public string? invoice_currency { get; set; }
        public string? gatepass_no { get; set; }
        public string? vendor_name { get; set; }
        public string? vendor_address { get; set; }
        public string? created_by { get; set; }
        public string? created_by_name { get; set; }
        public string? approval_by_name { get; set; }
        public DateTime approval_date { get; set; }
    }

    public class InvoiceDetailModel
    {
        public int id_invoice_detail { get; set; } 
        public string? asset_no { get; set; }
        public string? asset_subnumber { get; set; }
        public string? asset_desc { get; set; }
        public string? asset_class { get; set; }
        public decimal amount { get; set; }
    }

    public class InvoiceDetailViewModel
    {
        public int? id_gatepass { get; set; }
        public string? invoice_main_file { get; set; }
        public string? invoice_secondary_file { get; set; }
        public string? invoice_payment { get; set; }
        public string? invoice_currency { get; set; }
        public List<GatePassInvoiceAssetModel>? assets { get; set; }
    }

}