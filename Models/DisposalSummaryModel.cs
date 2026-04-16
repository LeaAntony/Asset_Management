using System.Data;

namespace Asset_Management.Models
{
    public class DisposalSummaryModel
    {
        public int? pending { get; set; }
        public int? waiting_approval { get; set; }
        public int? waiting_bap { get; set; }
        public int? waiting_doc_number { get; set; }
        public int? rejected { get; set; }
        public int? complete { get; set; }
    }
}
