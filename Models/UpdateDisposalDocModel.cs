namespace Asset_Management.Models
{
    public class UpdateDisposalDocModel
    {
        public List<UpdateAssetDocModel> data_docs { get; set; }
        public int? id_order { get; set; }
    }
}
