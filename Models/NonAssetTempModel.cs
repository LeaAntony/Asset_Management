namespace Asset_Management.Models
{
    public class NonAssetTempModel
    {
        public int id_temp { get; set; }
        public string sesa_id { get; set; }
        public string po_or_asset_no { get; set; }
        public string description { get; set; }
        public decimal qty { get; set; }
        public string uom { get; set; }
        public decimal price_value { get; set; }
        public decimal? hs_code { get; set; }
        public string gp_before { get; set; }
        public string gp_after { get; set; }
        public DateTime record_date { get; set; }
    }
}