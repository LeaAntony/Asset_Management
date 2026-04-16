using Microsoft.EntityFrameworkCore;
using Asset_Management.Models;

namespace Asset_Management.Function
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options)
        {

        }
        public DbSet<AssetListModel> v_asset { get; set; }
        public DbSet<AssetCountListModel> v_asset_count { get; set; }
        public DbSet<DisposalListModel> v_disposal_header { get; set; }
        public DbSet<DisposalDetailModel> v_disposal_detail { get; set; }
        public DbSet<ReconAssetModel> v_recon_asset { get; set; }
        public DbSet<RateEuroModel> tbl_exc_rate_eur { get; set; }
        public DbSet<UserDetailModel> mst_users { get; set; }
        public DbSet<AssetAUCModel> v_asset_auc { get; set; }
        public DbSet<VendorGatepassModel> mst_vendor_gatepass { get; set; }
        public DbSet<GatePassModel> v_gatepass { get; set; }
        public DbSet<GatePassNonAssetModel> v_gatepass_non_asset { get; set; }


    }
}
