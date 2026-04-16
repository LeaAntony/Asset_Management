using System.ComponentModel.DataAnnotations;

namespace Asset_Management.Models
{
    public class ProformaShippingModel
    {
        [Key]
        public int id_shipping { get; set; }
        public int? id_proforma { get; set; }
        public string? plant { get; set; }
        public DateTime? shipment_date { get; set; }
        public string? shipment_type { get; set; }
        public string? dhl_awb { get; set; }
        public string? created_by { get; set; }
        public DateTime? record_date { get; set; }
    }

    public class ShippingBoxModel
    {
        public int id_box { get; set; }
        public string? box_no { get; set; }
        public decimal? length_cm { get; set; }
        public decimal? width_cm { get; set; }
        public decimal? height_cm { get; set; }
        public decimal? gross_weight_kg { get; set; }
        public decimal? net_weight_kg { get; set; }
        public int asset_count { get; set; }
    }

    public class TempShippingBoxModel
    {
        [Key]
        public int id_temp { get; set; }
        public string? created_by { get; set; }
        public int? id_gatepass { get; set; }
        public string? box_no { get; set; }
        public decimal? length_cm { get; set; }
        public decimal? width_cm { get; set; }
        public decimal? height_cm { get; set; }
        public decimal? gross_weight_kg { get; set; }
        public decimal? net_weight_kg { get; set; }
        public string? asset_list { get; set; }
        public DateTime? create_date { get; set; }
    }

    public class ShippingCreateViewModel
    {
        public int id_gatepass { get; set; }
        public string? gatepass_no { get; set; }
        public string? shipping_plant { get; set; }
        public string? plant { get; set; }
        public DateTime? shipment_date { get; set; }
        public string? shipment_type { get; set; }
        public string? dhl_awb { get; set; }
        public List<ShippingAssetModel>? assets { get; set; }
        public List<TempShippingBoxModel>? temp_boxes { get; set; }
        public List<ShippingBoxModel>? saved_boxes { get; set; }
        public bool is_shipping_saved { get; set; }
        public bool is_data_complete_in_db { get; set; }
    }

    public class ShippingAssetModel
    {
        public int id_detail { get; set; }
        public string? asset_no { get; set; }
        public string? asset_subnumber { get; set; }
        public string? asset_desc { get; set; }
        public decimal? hs_code { get; set; }
        public int? id_box { get; set; }
        public bool is_assigned { get; set; }
        public bool is_selected { get; set; }
    }

    public class ShippingBoxDisplayModel
    {
        public List<ShippingBoxModel>? saved_boxes { get; set; }
        public List<TempShippingBoxModel>? temp_boxes { get; set; }
    }

    public class ShippingExportModel
    {
        public string? proforma_no { get; set; }
        public string? plant { get; set; }
        public string? tax { get; set; }
        public DateTime? shipment_date { get; set; }
        public string? shipment_type { get; set; }
        public string? dhl_awb { get; set; }
        public string? vendor_name { get; set; }
        public string? street { get; set; }
        public string? city { get; set; }
        public string? country { get; set; }
        public string? postal_code { get; set; }
        public string? phone_no { get; set; }
        public string? attn_to { get; set; }
        public string? coo { get; set; }
        public string? ship_mode { get; set; }
        public string? courier_account_no { get; set; }
        public string? courier_charges { get; set; }
        public string? freight_charges { get; set; }
        public string? incoterms { get; set; }
        public string? invoice_payment { get; set; }
        public string? requestor_name { get; set; }
        public string? requestor_plant { get; set; }
        public int total_boxes { get; set; }
        public int total_assets { get; set; }
        public string? remark { get; set; }
        public List<ShippingBoxExportModel>? boxes { get; set; }
    }

    public class ShippingBoxExportModel
    {
        public string? box_no { get; set; }
        public decimal length_cm { get; set; }
        public decimal width_cm { get; set; }
        public decimal height_cm { get; set; }
        public decimal gross_weight_kg { get; set; }
        public decimal net_weight_kg { get; set; }
        public List<ShippingAssetExportModel>? assets { get; set; }
    }

    public class ShippingAssetExportModel
    {
        public string? asset_no { get; set; }
        public string? asset_subnumber { get; set; }
        public string? asset_desc { get; set; }
        public decimal current_apc { get; set; }
        public decimal? hs_code { get; set; }
    }

    public class ShippingNonAssetCreateViewModel
    {
        public int id_gatepass { get; set; }
        public string? gatepass_no { get; set; }
        public string? shipping_plant { get; set; }
        public string? plant { get; set; }
        public DateTime? shipment_date { get; set; }
        public string? shipment_type { get; set; }
        public string? dhl_awb { get; set; }
        public List<ShippingNonAssetModel>? items { get; set; }
        public List<TempShippingNonAssetBoxModel>? temp_boxes { get; set; }
        public List<ShippingNonAssetBoxModel>? saved_boxes { get; set; }
        public bool is_shipping_saved { get; set; }
        public bool is_data_complete_in_db { get; set; }
    }

    public class ShippingNonAssetModel
    {
        public int id_detail { get; set; }
        public string? po_or_asset_no { get; set; }
        public string? description { get; set; }
        public decimal qty { get; set; }
        public string? uom { get; set; }
        public decimal price_value { get; set; }
        public decimal? hs_code { get; set; }
        public int? id_box { get; set; }
        public bool is_assigned { get; set; }
        public bool is_selected { get; set; }
    }

    public class TempShippingNonAssetBoxModel
    {
        [Key]
        public int id_temp { get; set; }
        public string? created_by { get; set; }
        public int? id_gatepass { get; set; }
        public string? box_no { get; set; }
        public decimal? length_cm { get; set; }
        public decimal? width_cm { get; set; }
        public decimal? height_cm { get; set; }
        public decimal? gross_weight_kg { get; set; }
        public decimal? net_weight_kg { get; set; }
        public string? item_list { get; set; }
        public DateTime? create_date { get; set; }
    }

    public class ShippingNonAssetBoxModel
    {
        public int id_box { get; set; }
        public string? box_no { get; set; }
        public decimal length_cm { get; set; }
        public decimal width_cm { get; set; }
        public decimal height_cm { get; set; }
        public decimal gross_weight_kg { get; set; }
        public decimal net_weight_kg { get; set; }
        public int item_count { get; set; }
    }

    public class ShippingNonAssetBoxDisplayModel
    {
        public List<ShippingNonAssetBoxModel>? saved_boxes { get; set; }
        public List<TempShippingNonAssetBoxModel>? temp_boxes { get; set; }
    }

    public class ShippingNonAssetExportModel
    {
        public string? proforma_no { get; set; }
        public string? plant { get; set; }
        public string? tax { get; set; }
        public DateTime? shipment_date { get; set; }
        public string? shipment_type { get; set; }
        public string? dhl_awb { get; set; }
        public string? vendor_name { get; set; }
        public string? street { get; set; }
        public string? city { get; set; }
        public string? country { get; set; }
        public string? postal_code { get; set; }
        public string? phone_no { get; set; }
        public string? attn_to { get; set; }
        public string? coo { get; set; }
        public string? ship_mode { get; set; }
        public string? courier_account_no { get; set; }
        public string? courier_charges { get; set; }
        public string? freight_charges { get; set; }
        public string? incoterms { get; set; }
        public string? invoice_payment { get; set; }
        public string? requestor_name { get; set; }
        public string? requestor_plant { get; set; }
        public int total_boxes { get; set; }
        public int total_items { get; set; }
        public string? remark { get; set; }
        public List<ShippingNonAssetBoxExportModel>? boxes { get; set; }
    }

    public class ShippingNonAssetBoxExportModel
    {
        public string? box_no { get; set; }
        public decimal length_cm { get; set; }
        public decimal width_cm { get; set; }
        public decimal height_cm { get; set; }
        public decimal gross_weight_kg { get; set; }
        public decimal net_weight_kg { get; set; }
        public List<ShippingNonAssetItemExportModel>? items { get; set; }
    }

    public class ShippingNonAssetItemExportModel
    {
        public string? po_or_asset_no { get; set; }
        public string? description { get; set; }
        public decimal qty { get; set; }
        public string? uom { get; set; }
        public decimal price_value { get; set; }
        public decimal? hs_code { get; set; }
    }
}