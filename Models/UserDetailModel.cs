using System.ComponentModel.DataAnnotations;
namespace Asset_Management.Models
{
    public class UserDetailModel
    {
        [Key]
        public int? id_user { get; set; }
        public string? sesa_id { get; set; }
        public string? name { get; set; }
        public string? email { get; set; }
        public string? level { get; set; }
        public string? role { get; set; }
        public string? department { get; set; }
        public string? plant { get; set; }
        public string? manager_sesa_id { get; set; }
        public string? manager_name { get; set; }
        public string? manager_email { get; set; }
        public string? other_dept { get; set; }
        public int? role_manage_user { get; set; }

        public string? delegated_to { get; set; }
        public string? delegated_name { get; set; }

    }
}
