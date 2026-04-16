using DocumentFormat.OpenXml.Spreadsheet;
using iText.StyledXmlParser.Jsoup.Select;
using Microsoft.AspNetCore.Mvc;
using Asset_Management.Models;
using System.Data;
using System.Data.SqlClient;
using System.Drawing.Drawing2D;
using System.Text;

namespace Asset_Management.Function
{
    public class DatabaseAccessLayer
    {
        public string ConnectionString = "Data Source=localhost;Initial Catalog=Asset_Management;Integrated Security=true;Persist Security Info=True;MultipleActiveResultSets=true";

        public List<string> GetDepartment()
        {
            List<string> deptList = new List<string>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_DEPARTMENT", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            deptList.Add(reader["dept_name"].ToString() ?? "");
                        }
                    }
                }

                conn.Close();
            }
            return deptList;
        }
        public List<LotFamilyModel> GetLotFamily()
        {
            List<LotFamilyModel> dataList = new List<LotFamilyModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_LOT_FAMILY", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            LotFamilyModel data = new LotFamilyModel();
                            data.id = reader["id"].ToString();
                            data.lot = reader["lot"].ToString();
                            data.family = reader["family"].ToString();
                            data.station_area = reader["station_area"].ToString();
                            data.family_station_area = data.family + "/" + data.station_area;
                            data.cost_center = reader["cost_center"].ToString();
                            data.family_description = reader["family_description"].ToString();
                            data.station_area_description = reader["station_area_description"].ToString();
                            data.family_station_area_description = data.family_description + "/" + data.station_area_description;
                            dataList.Add(data);
                        }
                    }
                }

                conn.Close();
            }
            return dataList;
        }
        public List<LotMachineModel> GetLocMachine()
        {
            List<LotMachineModel> dataList = new List<LotMachineModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_LOT_MACHINE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            LotMachineModel data = new LotMachineModel();
                            data.lot = reader["lot"].ToString();
                            data.machine_name = reader["machine_name"].ToString();
                            dataList.Add(data);
                        }
                    }
                }

                conn.Close();
            }
            return dataList;
        }
        public List<AssetListModel> GetMotherAssetList(string search)
        {
            List<AssetListModel> assetList = new List<AssetListModel>();

            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_MOTHER_ASSET_LIST", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@search", search);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            AssetListModel asset = new AssetListModel
                            {
                                asset_no = reader["asset_no"].ToString(),
                                asset_subnumber = reader["asset_subnumber"].ToString(),
                                asset_desc = reader["asset_desc"].ToString(),
                                cost_center = reader["cost_center"].ToString()
                            };
                            assetList.Add(asset);
                        }
                    }
                }
            }

            return assetList;
        }
        public List<VendorModel> GetVendor()
        {
            List<VendorModel> dataList = new List<VendorModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_VENDOR", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            VendorModel data = new VendorModel();
                            data.vendor_code = reader["vendor_code"].ToString();
                            data.vendor_name = reader["vendor_name"].ToString();
                            dataList.Add(data);
                        }
                    }
                }

                conn.Close();
            }
            return dataList;
        }
        public List<ProductFamilyModel> GetProductFamily()
        {
            List<ProductFamilyModel> dataList = new List<ProductFamilyModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_PRODUCT_FAMILY", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            ProductFamilyModel data = new ProductFamilyModel();
                            data.cost_center_desc = reader["cost_center_desc"].ToString();
                            dataList.Add(data);
                        }
                    }
                }

                conn.Close();
            }
            return dataList;
        }
        public List<UserDetailModel> GetUserDetail(string sesa_id)
        {
            List<UserDetailModel> dataList = new List<UserDetailModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_USER_DETAIL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            UserDetailModel row = new UserDetailModel();
                            row.sesa_id = reader["sesa_id"].ToString();
                            row.name = reader["name"].ToString();
                            row.email = reader["email"].ToString();
                            row.level = reader["level"].ToString();
                            row.role = reader["role"].ToString();
                            row.department = reader["department"].ToString();
                            row.manager_sesa_id = reader["manager_sesa_id"].ToString();
                            row.manager_name = reader["manager_name"].ToString();
                            row.manager_email = reader["manager_email"].ToString();
                            row.other_dept = reader["other_dept"].ToString();
                            row.role_manage_user = Convert.ToInt32(reader["role_manage_user"]);
                            dataList.Add(row);
                        }
                    }
                }

                conn.Close();
            }
            return dataList;
        }
        public List<UserDetailModel> GetApproverList(string search = "", string excludeSesaId = "")
        {
            List<UserDetailModel> approverList = new List<UserDetailModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_APPROVER_LIST", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@search", "%" + search + "%");
                    cmd.Parameters.AddWithValue("@excludeSesaId", excludeSesaId ?? "");
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            UserDetailModel row = new UserDetailModel();
                            row.sesa_id = reader["sesa_id"].ToString();
                            row.name = reader["name"].ToString();
                            row.level = reader["level"].ToString();
                            row.role = reader["role"].ToString();
                            approverList.Add(row);
                        }
                    }
                }
                conn.Close();
            }
            return approverList;
        }
        public UserDetailModel GetUserDelegationInfo(string sesaId)
        {
            UserDetailModel userInfo = new UserDetailModel();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_USER_DELEGATION_INFO", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@sesa_id", sesaId);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        userInfo.sesa_id = reader["sesa_id"].ToString();
                        userInfo.name = reader["name"].ToString();
                        userInfo.delegated_to = reader["delegated_to"].ToString();
                        userInfo.delegated_name = reader["delegated_name"].ToString();
                    }
                }
                conn.Close();
            }
            return userInfo;
        }
        public string SetDelegation(string originalSesaId, string delegatedTo)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_DELEGATION", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@original_sesa_id", originalSesaId);
                    cmd.Parameters.AddWithValue("@delegated_to", delegatedTo);

                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        return reader["result"].ToString();
                    }
                }
            }
            return "error;Unknown error";
        }

        public string RemoveDelegation(string sesaId)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("DELETE_DELEGATION", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@sesa_id", sesaId);

                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        return reader["result"].ToString();
                    }
                }
            }
            return "error;Unknown error";
        }

        public List<string> GetAllOriginalApproversForDelegated(string sesa_id)
        {
            List<string> originalApprovers = new List<string>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_ORIGINAL_APPROVER_DELEGATION", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@delegated_sesa_id", sesa_id);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            originalApprovers.Add(reader["sesa_id"].ToString());
                        }
                    }
                }
            }
            return originalApprovers;
        }

        public bool CanViewApproval(string current_user_sesa_id, string approver_sesa_id)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_APPROVAL_LIST", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@current_user_sesa_id", current_user_sesa_id);
                    cmd.Parameters.AddWithValue("@approver_sesa_id", approver_sesa_id);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        return Convert.ToBoolean(reader["result"]);
                    }
                }
            }
            return false;
        }

        public string GetGatepassApproverAtLevel(int id_gatepass, string approval_level)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GATEPASS_APPROVER_LEVEL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@approval_level", approval_level);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        return reader["result"].ToString();
                    }
                }
            }
            return "";
        }

        public string GetGatepassConfirmerHOD(int id_gatepass)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GATEPASS_CONFIRMER_HOD", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    var result = cmd.ExecuteScalar();
                    return result?.ToString() ?? "";
                }
            }
        }

        public bool IsDelegatedApprover(string delegatedUserId, string originalApprover)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_DELEGATED_APPROVER", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@delegated_user_id", delegatedUserId);
                    cmd.Parameters.AddWithValue("@original_approver", originalApprover);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        return Convert.ToBoolean(reader["result"]);
                    }
                }
            }
            return false;
        }

        public List<AssetListModel> GetAsset(string asset_search)
        {
            List<AssetListModel> assetList = new List<AssetListModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_SEARCH_ASSET", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@asset_search", asset_search);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            AssetListModel row = new AssetListModel();
                            row.asset_no = reader["asset_no"].ToString();
                            row.asset_subnumber = reader["asset_subnumber"].ToString();
                            row.asset_desc = reader["asset_desc"].ToString();
                            row.asset_class = reader["asset_class"].ToString();
                            row.cost_center = reader["cost_center"].ToString();
                            row.capitalized_on = Utility.ConvertToDateFormat(reader["capitalized_on"].ToString() ?? "");
                            row.current_apc = Convert.ToDouble(reader["current_apc"]);
                            row.curr_bk_val = Convert.ToDouble(reader["curr_bk_val"]);
                            row.status_disposal = Convert.ToInt32(reader["status_disposal"]);
                            row.is_counted = reader["is_counted"] != DBNull.Value ? Convert.ToInt32(reader["is_counted"]) : 0;
                            row.count_year = reader["count_year"]?.ToString();
                            assetList.Add(row);
                        }
                    }
                }
                conn.Close();
            }
            return assetList;
        }
        public List<AssetListModel> GetAssetNoTagging(string sesa_owner = "")
        {
            List<AssetListModel> assetList = new List<AssetListModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_ASSET_NO_TAGGING", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@sesa_owner", sesa_owner);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            AssetListModel row = new AssetListModel();
                            row.asset_no = reader["asset_no"].ToString();
                            row.asset_subnumber = reader["asset_subnumber"].ToString();
                            row.department = reader["department"].ToString();
                            row.asset_desc = reader["asset_desc"].ToString();
                            row.asset_class = reader["asset_class"].ToString();
                            row.cost_center = reader["cost_center"].ToString();
                            row.capitalized_on = Utility.ConvertToDateFormat(reader["capitalized_on"].ToString() ?? "");
                            row.name_owner = reader["name_owner"].ToString();
                            row.vendor_name = reader["vendor_name"].ToString();
                            row.tag_validated = Convert.ToString(reader["tag_validated"]);
                            assetList.Add(row);
                        }
                    }
                }
                conn.Close();
            }
            return assetList;
        }

        public List<AssetListModel> GetAssetNoTaggingByClass(string class_desc)
        {
            List<AssetListModel> assetList = new List<AssetListModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_ASSET_NO_TAGGING_BY_CLASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@class_desc", class_desc);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            AssetListModel row = new AssetListModel();
                            row.asset_no = reader["asset_no"].ToString();
                            row.asset_subnumber = reader["asset_subnumber"].ToString();
                            row.asset_desc = reader["asset_desc"].ToString();
                            row.asset_class = reader["asset_class"].ToString();
                            row.asset_class_desc = reader["asset_class_desc"].ToString();
                            row.department = reader["department"].ToString();
                            row.plant = reader["plant"].ToString();
                            row.cost_center = reader["cost_center"].ToString();
                            row.capitalized_on = Utility.ConvertToDateFormat(reader["capitalized_on"].ToString() ?? "");
                            row.aging_days_tagging = Convert.ToInt32(reader["aging_days_tagging"]);
                            assetList.Add(row);
                        }
                    }
                }

                conn.Close();
            }
            return assetList;
        }
        public List<AssetListModel> GetAssetNoTaggingByDepartment(string department)
        {
            List<AssetListModel> assetList = new List<AssetListModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_ASSET_NO_TAGGING_BY_DEPARTMENT", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@department", department);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            AssetListModel row = new AssetListModel();
                            row.asset_no = reader["asset_no"].ToString();
                            row.asset_subnumber = reader["asset_subnumber"].ToString();
                            row.asset_desc = reader["asset_desc"].ToString();
                            row.asset_class = reader["asset_class"].ToString();
                            row.asset_class_desc = reader["asset_class_desc"].ToString();
                            row.department = reader["department"].ToString();
                            row.plant = reader["plant"].ToString();
                            row.owner = reader["owner"].ToString();
                            row.cost_center = reader["cost_center"].ToString();
                            row.capitalized_on = Utility.ConvertToDateFormat(reader["capitalized_on"].ToString() ?? "");
                            row.aging_days_tagging = Convert.ToInt32(reader["aging_days_tagging"]);
                            row.current_apc = Convert.ToDouble(reader["current_apc"]);
                            row.curr_bk_val = Convert.ToDouble(reader["curr_bk_val"]);
                            assetList.Add(row);
                        }
                    }
                }

                conn.Close();
            }
            return assetList;
        }

        public List<AssetListModel> GetAssetNeedCount(string sesa_id)
        {
            List<AssetListModel> assetList = new List<AssetListModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_ASSET_NEED_COUNT", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            AssetListModel row = new AssetListModel();
                            row.asset_no = reader["asset_no"].ToString();
                            row.asset_subnumber = reader["asset_subnumber"].ToString();
                            row.asset_desc = reader["asset_desc"].ToString();
                            row.asset_class = reader["asset_class"].ToString();
                            row.cost_center = reader["cost_center"].ToString();
                            row.cc_desc = reader["cc_desc"].ToString();
                            row.cc_grouping = reader["cc_grouping"].ToString();
                            row.cc_plant = reader["cc_plant"].ToString();
                            row.count_year = reader["count_year"].ToString();
                            row.name_owner = reader["name_owner"].ToString();
                            row.department = reader["department"].ToString();
                            row.count_month = reader["count_month"] == DBNull.Value ? (int?)null : Convert.ToInt32(reader["count_month"]);
                            row.count_month_name = reader["count_month_name"].ToString();
                            row.capitalized_on = Utility.ConvertToDateFormat(reader["capitalized_on"].ToString() ?? "");
                            row.is_counted = Convert.ToInt32(reader["is_counted"]);
                            assetList.Add(row);
                        }
                    }
                }
                conn.Close();
            }
            return assetList;
        }

        public List<AssetListModel> GetAssetCountDetail(string count_month, string count_year, string? department = null)
        {
            List<AssetListModel> assetList = new List<AssetListModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_ASSET_COUNT_DETAIL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@count_month", count_month);
                    cmd.Parameters.AddWithValue("@count_year", count_year);
                    cmd.Parameters.AddWithValue("@department", string.IsNullOrWhiteSpace(department) ? (object)DBNull.Value : department);

                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            AssetListModel row = new AssetListModel();
                            row.count_year = reader["count_year"].ToString();
                            row.count_month = Convert.IsDBNull(reader["count_month"]) ? 0 : Convert.ToInt32(reader["count_month"]);
                            row.count_month_name = reader["count_month_name"].ToString();
                            row.asset_no = reader["asset_no"].ToString();
                            row.asset_subnumber = reader["asset_subnumber"].ToString();
                            row.asset_desc = reader["asset_desc"].ToString();
                            row.asset_class = reader["asset_class"].ToString();
                            row.department = reader["department"].ToString();
                            row.plant = reader["plant"].ToString();
                            row.owner = reader["owner"].ToString();
                            row.current_apc = Convert.IsDBNull(reader["current_apc"]) ? 0.0 : Convert.ToDouble(reader["current_apc"]);
                            row.curr_bk_val = Convert.IsDBNull(reader["curr_bk_val"]) ? 0.0 : Convert.ToDouble(reader["curr_bk_val"]);
                            row.cost_center = reader["cost_center"].ToString();
                            row.capitalized_on = Utility.ConvertToDateFormat(reader["capitalized_on"].ToString() ?? "");
                            row.is_counted = Convert.ToInt32(reader["is_counted"]);
                            row.counted_by = reader["counted_by"].ToString();
                            row.date_counted = reader["date_counted"].ToString();
                            row.existence = reader["existence"].ToString();
                            row.good_condition = reader["good_condition"].ToString();
                            row.still_in_operation = reader["still_in_operation"].ToString();
                            row.tagging_available = reader["tagging_available"].ToString();
                            row.applicable_of_tagging = reader["applicable_of_tagging"].ToString();
                            row.correct_naming = reader["correct_naming"].ToString();
                            row.correct_location = reader["correct_location"].ToString();
                            row.retagging = reader["retagging"].ToString();
                            row.file_imgs = reader["file_imgs"].ToString();
                            row.is_validated = Convert.ToInt32(reader["is_validated"]);
                            assetList.Add(row);
                        }
                    }
                }
                conn.Close();
            }
            return assetList;
        }
        public List<AssetListModel> GetAssetTaggingValidation()
        {
            List<AssetListModel> assetList = new List<AssetListModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_ASSET_TAGGING_VALIDATION", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            AssetListModel row = new AssetListModel();
                            row.asset_no = reader["asset_no"].ToString();
                            row.asset_subnumber = reader["asset_subnumber"].ToString();
                            row.asset_desc = reader["asset_desc"].ToString();
                            row.asset_class = reader["asset_class"].ToString();
                            row.cost_center = reader["cost_center"].ToString();
                            row.capitalized_on = Utility.ConvertToDateFormat(reader["capitalized_on"].ToString() ?? "");
                            row.tag_validated = Convert.ToString(reader["tag_validated"]);
                            row.file_tag = reader["file_tag"].ToString();
                            row.name_owner = reader["name_owner"].ToString();
                            assetList.Add(row);
                        }
                    }
                }
                conn.Close();
            }
            return assetList;
        }

        public List<AssetListModel> GET_ASSET_NO_TAGGING()
        {
            List<AssetListModel> assetList = new List<AssetListModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_ASSET_NO_TAGGING_ALL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            AssetListModel row = new AssetListModel();
                            row.asset_no = reader["asset_no"].ToString();
                            row.asset_subnumber = reader["asset_subnumber"].ToString();
                            row.asset_desc = reader["asset_desc"].ToString();
                            row.asset_class = reader["asset_class"].ToString();
                            row.cost_center = reader["cost_center"].ToString();
                            row.plant = reader["plant"].ToString();
                            row.department = reader["department"].ToString();
                            row.capitalized_on = Utility.ConvertToDateFormat(reader["capitalized_on"].ToString() ?? "");
                            row.vendor_name = reader["vendor_name"].ToString();
                            row.sesa_owner = reader["sesa_owner"].ToString();
                            row.aging_days = reader["aging_days"].ToString();
                            assetList.Add(row);
                        }
                    }
                }
                conn.Close();
            }
            return assetList;
        }

        public int GetTotalNoTagging(string sesa_owner = "")
        {
            int totalNoTagging = 0;
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_TOTAL_NO_TAGGING", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@sesa_owner", sesa_owner ?? "");
                    totalNoTagging = (int)cmd.ExecuteScalar();
                }
                conn.Close();
            }
            return totalNoTagging;
        }

        public int GetTotalAssetNeedCount(string sesa_id)
        {
            int totalAssetCount = 0;
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_TOTAL_ASSET_NEED_COUNT", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id ?? "");
                    totalAssetCount = (int)cmd.ExecuteScalar();
                }
                conn.Close();
            }
            return totalAssetCount;
        }


        public int GetTotalNoTaggingWaitingValidation(string sesa_owner = "")
        {
            int totalNoTagging = 0;
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_TOTAL_NO_TAGGING_WAITING_VALIDATION", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@sesa_owner", sesa_owner ?? "");
                    totalNoTagging = (int)cmd.ExecuteScalar();
                }
                conn.Close();
            }
            return totalNoTagging;
        }

        public IActionResult GetNoTaggingByClass()
        {
            var chartData = new ChartModel();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_NO_TAGGING_BY_CLASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            chartData.labels.Add(reader["class"].ToString());
                            chartData.values.Add(Convert.ToDouble(reader["total_asset"]));

                        }
                    }
                }

                conn.Close();
            }
            return new JsonResult(new
            {
                labels = chartData.labels,
                values = chartData.values
            });
        }
        public IActionResult GetNoTaggingByDepartment()
        {
            var chartData = new ChartModel();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_NO_TAGGING_BY_DEPARTMENT", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            chartData.labels.Add(reader["department"].ToString());
                            chartData.values.Add(Convert.ToDouble(reader["total_asset"]));

                        }
                    }
                }

                conn.Close();
            }
            return new JsonResult(new
            {
                labels = chartData.labels,
                values = chartData.values
            });
        }
        public List<TableModel> GetNoTaggingByDepartmentDetail()
        {
            List<TableModel> dataList = new List<TableModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_NO_TAGGING_BY_DEPARTMENT", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            TableModel row = new TableModel();
                            row.label = reader["department"].ToString();
                            row.value = Convert.ToDouble(reader["total_asset"]);
                            dataList.Add(row);

                        }
                    }
                }

                conn.Close();
            }
            return dataList;
        }
        public IActionResult GetTaggingRatio()
        {
            int totalTagging = 0;
            int totalNoTagging = 0;
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_TAGGING_RATIO", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        if (reader.Read())
                        {
                            totalTagging = reader.IsDBNull(0) ? 0 : reader.GetInt32(0);
                            totalNoTagging = reader.IsDBNull(1) ? 0 : reader.GetInt32(1);
                        }
                    }
                }
                conn.Close();
            }
            return new JsonResult(new
            {
                totalTagging = totalTagging,
                totalNoTagging = totalNoTagging
            });
        }

        public List<string> GetCountDepartment()
        {
            List<string> deptList = new List<string>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_COUNT_DEPARTMENT", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            deptList.Add(reader["department"].ToString() ?? "");
                        }
                    }
                }
                conn.Close();
            }
            return deptList;
        }

        public IActionResult GetCountResult(string year_no, string? department = null)
        {
            var chartData = new ChartModel();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_COUNT_RESULT_V2", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@year_no", year_no ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@department", string.IsNullOrWhiteSpace(department) ? (object)DBNull.Value : department);

                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            chartData.labels.Add(reader["label"].ToString());
                            chartData.values.Add(Convert.ToDouble(reader["counted"]));
                            chartData.values2.Add(Convert.ToDouble(reader["not_counted"]));
                        }
                    }
                }
                conn.Close();
            }
            return new JsonResult(new
            {
                labels = chartData.labels,
                values = chartData.values,
                values2 = chartData.values2
            });
        }
        public IActionResult GetCountByDepartment(string count_year)
        {
            var chartData = new ChartModel();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_COUNT_BY_DEPARTMENT", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@count_year", count_year);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            chartData.labels.Add(reader["department"].ToString());
                            chartData.values.Add(Convert.ToDouble(reader["total_need_count"]));

                        }
                    }
                }

                conn.Close();
            }
            return new JsonResult(new
            {
                labels = chartData.labels,
                values = chartData.values
            });
        }
        public List<TableModel> GetCountByDepartmentList(string count_year, string month)
        {
            List<TableModel> dataList = new List<TableModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_COUNT_BY_DEPARTMENT", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@count_year", count_year);
                    cmd.Parameters.AddWithValue("@month", month);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            TableModel row = new TableModel();
                            row.label = reader["department"].ToString();
                            row.value = Convert.ToDouble(reader["total_need_count"]);
                            dataList.Add(row);

                        }
                    }
                }

                conn.Close();
            }
            return dataList;
        }
        public List<AssetListModel> GetAssetCountDepartmentDetail(string count_year, string department, string month)
        {
            List<AssetListModel> assetList = new List<AssetListModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_ASSET_COUNT_DEPARTMENT_DETAIL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@count_year", count_year);
                    cmd.Parameters.AddWithValue("@department", department);
                    cmd.Parameters.AddWithValue("@month", month);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            AssetListModel row = new AssetListModel();
                            row.count_year = reader["count_year"].ToString();
                            if (reader["count_month"] != DBNull.Value)
                            {
                                row.count_month = Convert.ToInt32(reader["count_month"]);
                            }
                            else
                            {
                                row.count_month = null;
                            }
                            row.count_month_name = reader["count_month_name"].ToString();
                            row.asset_no = reader["asset_no"].ToString();
                            row.asset_subnumber = reader["asset_subnumber"].ToString();
                            row.asset_desc = reader["asset_desc"].ToString();
                            row.asset_class = reader["asset_class"].ToString();
                            row.department = reader["department"].ToString();
                            row.plant = reader["plant"].ToString();
                            row.cost_center = reader["cost_center"].ToString();
                            row.owner = reader["owner"].ToString();
                            row.capitalized_on = Utility.ConvertToDateFormat(reader["capitalized_on"].ToString() ?? "");
                            row.is_counted = Convert.ToInt32(reader["is_counted"]);
                            row.counted_by = reader["counted_by"].ToString();
                            row.date_counted = reader["date_counted"].ToString();
                            row.existence = reader["existence"].ToString();
                            row.good_condition = reader["good_condition"].ToString();
                            row.still_in_operation = reader["still_in_operation"].ToString();
                            row.tagging_available = reader["tagging_available"].ToString();
                            row.applicable_of_tagging = reader["applicable_of_tagging"].ToString();
                            row.correct_naming = reader["correct_naming"].ToString();
                            row.correct_location = reader["correct_location"].ToString();
                            row.retagging = reader["retagging"].ToString();
                            row.file_imgs = reader["file_imgs"].ToString();
                            row.is_validated = Convert.ToInt32(reader["is_validated"]);
                            row.current_apc = Convert.ToDouble(reader["current_apc"]);
                            row.curr_bk_val = Convert.ToDouble(reader["curr_bk_val"]);
                            assetList.Add(row);
                        }
                    }
                }

                conn.Close();
            }
            return assetList;
        }

        public string UpdateTaggingAsset(string asset_no, string asset_subnumber, string file_tagging, string tagged_by)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_TAGGING_ASSET", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@asset_no", asset_no);
                    cmd.Parameters.AddWithValue("@asset_subnumber", asset_subnumber);
                    cmd.Parameters.AddWithValue("@file_tagging", file_tagging);
                    cmd.Parameters.AddWithValue("@tagged_by", tagged_by);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }
        public string ConfirmTagging(string asset_no, string asset_subnumber, string confirm_type)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("CONFIRM_TAGGING", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@asset_no", asset_no);
                    cmd.Parameters.AddWithValue("@asset_subnumber", asset_subnumber);
                    cmd.Parameters.AddWithValue("@confirm_type", confirm_type);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }
        public List<AssetCountListModel> GetCountedResult(string count_year, string asset_no, string asset_subnumber)
        {
            List<AssetCountListModel> assetList = new List<AssetCountListModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_COUNTED_RESULT", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@count_year", count_year);
                    cmd.Parameters.AddWithValue("@asset_no", asset_no);
                    cmd.Parameters.AddWithValue("@asset_subnumber", asset_subnumber);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            AssetCountListModel row = new AssetCountListModel();
                            row.existence = reader["existence"].ToString();
                            row.good_condition = reader["good_condition"].ToString();
                            row.still_in_operation = reader["still_in_operation"].ToString();
                            row.tagging_available = reader["tagging_available"].ToString();
                            row.applicable_of_tagging = reader["applicable_of_tagging"].ToString();
                            row.correct_naming = reader["correct_naming"].ToString();
                            row.correct_location = reader["correct_location"].ToString();
                            row.retagging = reader["retagging"].ToString();
                            row.file_imgs = reader["file_imgs"].ToString();
                            row.recount_by = reader["recount_by"].ToString();
                            assetList.Add(row);
                        }
                    }
                }
                conn.Close();
            }
            return assetList;
        }

        public string SubmitAssetCount(string count_year, string asset_no, string asset_subnumber, string existence, string good_condition, string still_in_operation, string tagging_available, string applicable_of_tagging, string correct_naming, string correct_location, string retagging, string filename, string counted_by)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("SUBMIT_ASSET_COUNT", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@count_year", count_year);
                    cmd.Parameters.AddWithValue("@asset_no", asset_no);
                    cmd.Parameters.AddWithValue("@asset_subnumber", asset_subnumber);
                    cmd.Parameters.AddWithValue("@existence", existence);
                    cmd.Parameters.AddWithValue("@good_condition", good_condition);
                    cmd.Parameters.AddWithValue("@still_in_operation", still_in_operation);
                    cmd.Parameters.AddWithValue("@tagging_available", tagging_available);
                    cmd.Parameters.AddWithValue("@applicable_of_tagging", applicable_of_tagging);
                    cmd.Parameters.AddWithValue("@correct_naming", correct_naming);
                    cmd.Parameters.AddWithValue("@correct_location", correct_location);
                    cmd.Parameters.AddWithValue("@retagging", retagging);
                    cmd.Parameters.AddWithValue("@filename", filename);
                    cmd.Parameters.AddWithValue("@counted_by", counted_by);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }
        public string UPDATE_RETAGGING_IMG(string id_count, string filename)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE tbl_asset_count SET file_imgs_retagging = @filename WHERE id_count = @id_count", conn))
                {
                    cmd.Parameters.AddWithValue("@id_count", id_count);
                    cmd.Parameters.AddWithValue("@filename", filename);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }
        public string Recount(int id_count, string sesa_id, string recount_remark = "")
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("RECOUNT", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_count", id_count);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.Parameters.AddWithValue("@recount_remark", recount_remark);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }
        public string ValidateCount(int id_count, string existence, string good_condition, string still_in_operation, string tagging, string applicable_of_tagging, string correct_naming, string correct_location, string retagging, string sesa_id)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("VALIDATE_COUNT_V2", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_count", id_count);
                    cmd.Parameters.AddWithValue("@existence", existence);
                    cmd.Parameters.AddWithValue("@good_condition", good_condition);
                    cmd.Parameters.AddWithValue("@still_in_operation", still_in_operation);
                    cmd.Parameters.AddWithValue("@tagging", tagging);
                    cmd.Parameters.AddWithValue("@applicable_of_tagging", applicable_of_tagging);
                    cmd.Parameters.AddWithValue("@correct_naming", correct_naming);
                    cmd.Parameters.AddWithValue("@correct_location", correct_location);
                    cmd.Parameters.AddWithValue("@retagging", retagging);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }
        public string UpdateCountAssetName(int id_count, string asset_name, string sesa_id)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_COUNT_ASSET_NAME", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_count", id_count);
                    cmd.Parameters.AddWithValue("@asset_name", asset_name);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }
        public string UpdateCountAssetLocation(int id_count, string asset_location, string sesa_id)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_COUNT_ASSET_LOCATION", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_count", id_count);
                    cmd.Parameters.AddWithValue("@asset_location", asset_location);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }
        public string UpdateUser(string sesa_id, string full_name, string email, string manager_sesa_id, string manager_name)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_USER", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.Parameters.AddWithValue("@full_name", full_name);
                    cmd.Parameters.AddWithValue("@email", email);
                    cmd.Parameters.AddWithValue("@manager_sesa_id", manager_sesa_id);
                    cmd.Parameters.AddWithValue("@manager_name", manager_name);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }
        
        public UserDetailModel? ValidateManualLogin(string sesa_id, string password)
        {
            string passwordHash = new Authentication().MD5Hash(password);

            using SqlConnection conn = new SqlConnection(ConnectionString);
            conn.Open();
            using SqlCommand cmd = new SqlCommand("GET_MANUAL_LOGIN", conn)
            {
                CommandType = CommandType.StoredProcedure
            };
            cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
            cmd.Parameters.AddWithValue("@password_hash", passwordHash);

            using SqlDataReader reader = cmd.ExecuteReader();
            if (!reader.Read())
                return null;

            return new UserDetailModel
            {
                sesa_id = reader["sesa_id"].ToString(),
                name = reader["name"].ToString(),
                email = reader["email"].ToString(),
                level = reader["level"].ToString(),
                role = reader["role"].ToString(),
                department = reader["department"].ToString(),
                manager_sesa_id = reader["manager_sesa_id"].ToString(),
                role_manage_user = Convert.ToInt32(reader["role_manage_user"])
            };
        }
        public UserDetailModel? ValidateSecurityLogin(string sesa_id, string password)
        {
            string passwordHash = new Authentication().MD5Hash(password);

            using SqlConnection conn = new SqlConnection(ConnectionString);
            conn.Open();
            using SqlCommand cmd = new SqlCommand("GET_SECURITY_LOGIN", conn)
            {
                CommandType = CommandType.StoredProcedure
            };
            cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
            cmd.Parameters.AddWithValue("@password_hash", passwordHash);

            using SqlDataReader reader = cmd.ExecuteReader();
            if (!reader.Read())
                return null;

            return new UserDetailModel
            {
                sesa_id = reader["sesa_id"].ToString(),
                name = reader["name"].ToString(),
                email = reader["email"].ToString(),
                level = reader["level"].ToString(),
                role = reader["role"].ToString(),
                department = reader["department"].ToString(),
                manager_sesa_id = reader["manager_sesa_id"].ToString(),
                role_manage_user = Convert.ToInt32(reader["role_manage_user"])
            };
        }
        public string CheckOldPass(string sesa_id, string old_pass)
        {
            var hashpassword = new Authentication();
            string passwordHash = hashpassword.MD5Hash(old_pass);
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("SELECT COUNT(*) FROM mst_users WHERE sesa_id=@sesa_id AND password=@passwordHash", conn))
                {
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.Parameters.AddWithValue("@passwordHash", passwordHash);

                    int rowCount = (int)cmd.ExecuteScalar();

                    if (rowCount > 0)
                    {
                        return "success";
                    }
                    else
                    {
                        return "failed";
                    }
                }
            }
        }
        public string SaveNewPass(string sesa_id, string new_pass)
        {
            var hashpassword = new Authentication();
            string passwordHash = hashpassword.MD5Hash(new_pass);
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE mst_users SET password=@password WHERE sesa_id=@sesa_id", conn))
                {
                    cmd.Parameters.AddWithValue("@password", passwordHash);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }
        
        public AssetListModel ASSET_TAGGING(string assetNo, string assetSubnumber)
        {
            AssetListModel pdf = new AssetListModel();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_ASSET_TAGGING", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@asset_no", assetNo);
                    cmd.Parameters.AddWithValue("@asset_subnumber", assetSubnumber);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            pdf.file_tag = reader["file_tag"].ToString();
                        }
                    }
                }
                conn.Close();
            }
            return pdf;
        }

        public string DELETE_TEMP_GATEPASS_SESA(string sesa_id)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("DELETE_TEMP_GATEPASS_SESA", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string ADD_TEMP_GATEPASS(string asset_no, string asset_subnumber, string sesa_id, string naming_output, string filename)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("INSERT_TEMP_GATEPASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@asset_no", asset_no);
                    cmd.Parameters.AddWithValue("@asset_subnumber", asset_subnumber);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.Parameters.AddWithValue("@naming_output", naming_output);
                    cmd.Parameters.AddWithValue("@filename", filename);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string ADD_TEMP_GATEPASS_DISPOSAL(string asset_no, string asset_subnumber, string sesa_id, string naming_output)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("INSERT_TEMP_GATEPASS_DISPOSAL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@asset_no", asset_no);
                    cmd.Parameters.AddWithValue("@asset_subnumber", asset_subnumber);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.Parameters.AddWithValue("@naming_output", naming_output);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string UPDATE_TEMP_GATEPASS_IMAGE(int id_temp, string filename)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_TEMP_GATEPASS_IMAGE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_temp", id_temp);
                    cmd.Parameters.AddWithValue("@filename", filename);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public List<DisposalTempModel> GET_TEMP_GATEPASS(string sesa_id)
        {
            List<DisposalTempModel> disList = new List<DisposalTempModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_TEMP_GATEPASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                DisposalTempModel row = new DisposalTempModel();
                                row.id_temp = Convert.ToInt32(reader["id_temp"]);
                                row.asset_no = reader["asset_no"].ToString();
                                row.asset_subnumber = reader["asset_subnumber"].ToString();
                                row.asset_desc = reader["asset_desc"].ToString();
                                row.naming_output = reader["naming_output"].ToString();
                                row.gp_before = reader["gp_before"].ToString();
                                row.hs_code = reader["hs_code"] == DBNull.Value ? null : Convert.ToDecimal(reader["hs_code"]);
                                row.current_apc = Convert.ToDouble(reader["current_apc"]);
                                row.curr_bk_val = Convert.ToDouble(reader["curr_bk_val"]);
                                disList.Add(row);
                            }
                        }
                    }
                }
                conn.Close();
            }
            return disList;
        }

        public List<string> GET_SECURITY_GUARD()
        {
            List<string> catList = new List<string>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_SECURITY_GUARD", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            catList.Add(reader["security_name"].ToString() ?? "");
                        }
                    }
                }
                conn.Close();
            }
            return catList;
        }

        public List<UserDetailModel> GET_NEW_PIC()
        {
            List<UserDetailModel> picList = new List<UserDetailModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_NEW_PIC", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            picList.Add(new UserDetailModel
                            {
                                sesa_id = reader["sesa_id"].ToString() ?? "",
                                name = reader["name"].ToString() ?? ""
                            });
                        }
                    }
                }
                conn.Close();
            }
            return picList;
        }

        public List<UserDetailModel> GET_EMPLOYEE()
        {
            List<UserDetailModel> employeeList = new List<UserDetailModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_EMPLOYEE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            employeeList.Add(new UserDetailModel
                            {
                                sesa_id = reader["sesa_id"].ToString() ?? "",
                                name = reader["name"].ToString() ?? ""
                            });
                        }
                    }
                }
                conn.Close();
            }
            return employeeList;
        }

        public List<string> GET_DEPARTMENT_LIST()
        {
            List<string> deptList = new List<string>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_DEPARTMENT_LIST", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            deptList.Add(reader["department"].ToString() ?? "");
                        }
                    }
                }
                conn.Close();
            }
            return deptList;
        }

        public List<string> GET_PLANT_LIST()
        {
            List<string> plantList = new List<string>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_PLANT_LIST", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            plantList.Add(reader["plant"].ToString() ?? "");
                        }
                    }
                }
                conn.Close();
            }
            return plantList;
        }

        public object SearchVendorGatepass(string search, int page = 1)
        {
            int pageSize = 30;
            int offset = (page - 1) * pageSize;

            var vendorList = new List<object>();
            int totalCount = 0;

            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();

                using (SqlCommand cmd = new SqlCommand("GET_SEARCH_VENDOR_GATEPASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@search", search ?? "");
                    cmd.Parameters.AddWithValue("@page", page);
                    cmd.Parameters.AddWithValue("@pageSize", pageSize);

                    SqlParameter totalCountParam = new SqlParameter("@totalCount", SqlDbType.Int);
                    totalCountParam.Direction = ParameterDirection.Output;
                    cmd.Parameters.Add(totalCountParam);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            vendorList.Add(new
                            {
                                vendorCode = reader["vendor_code"].ToString() ?? "",
                                vendorName = reader["vendor_name"].ToString() ?? ""
                            });
                        }
                    }

                    totalCount = (int)totalCountParam.Value;
                }

                conn.Close();
            }

            return new
            {
                items = vendorList,
                total_count = totalCount
            };
        }

        public string SAVE_GATEPASS(string create_by, string create_date, string category, string return_date, string new_pic, string type, string location, string employee, string csr, string vendor_code, string vendor_name
            , string vendor_phone, string vendor_address, string vendor_email, string security_guard, string shipping_plant, string vehicle_no, string driver_name, string remark, int? id_proforma = null, int? id_order = null, string supporting_documents = null)
        {
            string id_gatepass = "";
            string order_no = "";
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("INSERT_NEW_GATEPASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@category", string.IsNullOrEmpty(category) ? (object)DBNull.Value : category);
                    cmd.Parameters.AddWithValue("@return_date", string.IsNullOrEmpty(return_date) ? (object)DBNull.Value : return_date);
                    cmd.Parameters.AddWithValue("@new_pic", string.IsNullOrEmpty(new_pic) ? (object)DBNull.Value : new_pic);
                    cmd.Parameters.AddWithValue("@location", string.IsNullOrEmpty(location) ? (object)DBNull.Value : location);
                    cmd.Parameters.AddWithValue("@vendor_code", string.IsNullOrEmpty(vendor_code) ? (object)DBNull.Value : vendor_code);
                    cmd.Parameters.AddWithValue("@vendor_name", string.IsNullOrEmpty(vendor_name) ? (object)DBNull.Value : vendor_name);
                    cmd.Parameters.AddWithValue("@vendor_address", string.IsNullOrEmpty(vendor_address) ? (object)DBNull.Value : vendor_address);
                    cmd.Parameters.AddWithValue("@vendor_phone", string.IsNullOrEmpty(vendor_phone) ? (object)DBNull.Value : vendor_phone);
                    cmd.Parameters.AddWithValue("@vendor_email", string.IsNullOrEmpty(vendor_email) ? (object)DBNull.Value : vendor_email);
                    cmd.Parameters.AddWithValue("@vehicle_no", string.IsNullOrEmpty(vehicle_no) ? (object)DBNull.Value : vehicle_no);
                    cmd.Parameters.AddWithValue("@driver_name", string.IsNullOrEmpty(driver_name) ? (object)DBNull.Value : driver_name);
                    cmd.Parameters.AddWithValue("@security_guard", string.IsNullOrEmpty(security_guard) ? (object)DBNull.Value : security_guard);
                    cmd.Parameters.AddWithValue("@shipping_plant", string.IsNullOrEmpty(shipping_plant) ? (object)DBNull.Value : shipping_plant);
                    cmd.Parameters.AddWithValue("@remark", string.IsNullOrEmpty(remark) ? (object)DBNull.Value : remark);
                    cmd.Parameters.AddWithValue("@created_by", string.IsNullOrEmpty(create_by) ? (object)DBNull.Value : create_by);
                    cmd.Parameters.AddWithValue("@type", string.IsNullOrEmpty(type) ? (object)DBNull.Value : type);
                    cmd.Parameters.AddWithValue("@employee", string.IsNullOrEmpty(employee) ? (object)DBNull.Value : employee);
                    cmd.Parameters.AddWithValue("@csr_to", string.IsNullOrEmpty(csr) ? (object)DBNull.Value : csr);
                    cmd.Parameters.AddWithValue("@id_proforma", (object)id_proforma ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@id_order", (object)id_order ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@supporting_documents", string.IsNullOrEmpty(supporting_documents) ? (object)DBNull.Value : supporting_documents);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            id_gatepass = reader["id_gatepass"].ToString() ?? "";
                            order_no = reader["order_no"].ToString() ?? "";
                        }
                    }
                    return id_gatepass + ";" + order_no;
                }
            }
        }
        public string SAVE_PROFORMA(
            string attn_to, string street, string city, string country,
            string postal_code, string phone_no, string email, string coo, string file_attach, string file_support,
            string ship_mode, string courier_charges, string courier_name, string courier_account_no,
            string freight_charges, string incoterms, string invoice_payment, string requested_by)
        {
            string id_proforma = "";
            string proforma_no = "";

            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("INSERT_NEW_PROFORMA_HEADER", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@attn_to", string.IsNullOrEmpty(attn_to) ? (object)DBNull.Value : attn_to);
                    cmd.Parameters.AddWithValue("@street", string.IsNullOrEmpty(street) ? (object)DBNull.Value : street);
                    cmd.Parameters.AddWithValue("@city", string.IsNullOrEmpty(city) ? (object)DBNull.Value : city);
                    cmd.Parameters.AddWithValue("@country", string.IsNullOrEmpty(country) ? (object)DBNull.Value : country);
                    cmd.Parameters.AddWithValue("@postal_code", string.IsNullOrEmpty(postal_code) ? (object)DBNull.Value : postal_code);
                    cmd.Parameters.AddWithValue("@phone_no", string.IsNullOrEmpty(phone_no) ? (object)DBNull.Value : phone_no);
                    cmd.Parameters.AddWithValue("@email", string.IsNullOrEmpty(email) ? (object)DBNull.Value : email);
                    cmd.Parameters.AddWithValue("@coo", string.IsNullOrEmpty(coo) ? (object)DBNull.Value : coo);
                    cmd.Parameters.AddWithValue("@file_attach", string.IsNullOrEmpty(file_attach) ? (object)DBNull.Value : file_attach);
                    cmd.Parameters.AddWithValue("@file_support", string.IsNullOrEmpty(file_support) ? (object)DBNull.Value : file_support);
                    cmd.Parameters.AddWithValue("@ship_mode", string.IsNullOrEmpty(ship_mode) ? (object)DBNull.Value : ship_mode);
                    cmd.Parameters.AddWithValue("@courier_charges", string.IsNullOrEmpty(courier_charges) ? (object)DBNull.Value : courier_charges);
                    cmd.Parameters.AddWithValue("@courier_name", string.IsNullOrEmpty(courier_name) ? (object)DBNull.Value : courier_name);
                    cmd.Parameters.AddWithValue("@courier_account_no", string.IsNullOrEmpty(courier_account_no) ? (object)DBNull.Value : courier_account_no);
                    cmd.Parameters.AddWithValue("@freight_charges", string.IsNullOrEmpty(freight_charges) ? (object)DBNull.Value : freight_charges);
                    cmd.Parameters.AddWithValue("@incoterms", string.IsNullOrEmpty(incoterms) ? (object)DBNull.Value : incoterms);
                    cmd.Parameters.AddWithValue("@invoice_payment", string.IsNullOrEmpty(invoice_payment) ? (object)DBNull.Value : invoice_payment);
                    cmd.Parameters.AddWithValue("@requested_by", string.IsNullOrEmpty(requested_by) ? (object)DBNull.Value : requested_by);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            id_proforma = reader["id_proforma"].ToString() ?? "";
                            proforma_no = reader["proforma_no"].ToString() ?? "";
                        }
                    }
                }
                conn.Close();
            }
            return id_proforma + ";" + proforma_no;
        }
        public string UPDATE_HS_CODE_TEMP(string asset_no, string asset_subnumber, string hs_code, string sesa_id)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_HS_CODE_TEMP", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    decimal? hsCodeValue = null;
                    if (!string.IsNullOrEmpty(hs_code) && decimal.TryParse(hs_code, out decimal parsed))
                    {
                        hsCodeValue = parsed;
                    }

                    cmd.Parameters.AddWithValue("@asset_no", asset_no);
                    cmd.Parameters.AddWithValue("@asset_subnumber", asset_subnumber);
                    cmd.Parameters.AddWithValue("@hs_code", hsCodeValue ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string EditGatepass(int id_gatepass, string return_date)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_GATEPASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@return_date", return_date);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string CancelGatepass(int id_gatepass, string cancelled_by, string cancel_reason)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("CANCEL_GATEPASS", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        cmd.Parameters.AddWithValue("@cancelled_by", cancelled_by);
                        cmd.Parameters.AddWithValue("@cancel_reason", string.IsNullOrEmpty(cancel_reason) ? (object)DBNull.Value : cancel_reason);

                        conn.Open();
                        var result = cmd.ExecuteScalar()?.ToString() ?? "error;Unknown error occurred";
                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                return $"error;{ex.Message}";
            }
        }

        public string UPDATE_IMG_GATEPASS(int id_gatepass, string filename_doc)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_IMG_GATEPASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@filename_doc", filename_doc);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string UPDATE_AFTER_GATEPASS_RETURN(int id_gatepass)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_AFTER_GATEPASS_RETURN_PIC", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string UPDATE_AFTER_GATEPASS_PIC(int id_gatepass, string ret_file, string id_detail)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(ConnectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("UPDATE_AFTER_GATEPASS_PIC", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        cmd.Parameters.AddWithValue("@ret_file", ret_file);
                        cmd.Parameters.AddWithValue("@id_detail", id_detail);
                        cmd.ExecuteNonQuery();
                        return "success";
                    }
                }
            }
            catch (Exception ex)
            {
                return $"error: {ex.Message}";
            }
        }

        public string UPDATE_AFTER_GATEPASS(int id_gatepass, string ret_file, string id_detail)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_AFTER_GATEPASS_RETURN_V1", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_detail", id_detail);
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@ret_file", ret_file);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public List<GatePassModel> GET_IMG_GATEPASS(string id_gatepass)
        {
            List<GatePassModel> images = new List<GatePassModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_IMG_GATEPASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            GatePassModel image = new GatePassModel
                            {
                                gatepass_no = reader["gatepass_no"].ToString(),
                                category = reader["category"].ToString(),
                                asset_no = reader["asset_no"].ToString(),
                                asset_subnumber = reader["asset_subnumber"].ToString(),
                                image_before = reader["gp_before"] != DBNull.Value ? reader["gp_before"].ToString() : "NotAvail.png",
                                image_after = reader["gp_after"] != DBNull.Value ? reader["gp_after"].ToString() : "NotAvail.png"
                            };
                            images.Add(image);
                        }
                    }
                }
                conn.Close();
            }
            return images;
        }

        public List<GatePassModel> GET_GATEPASS_DETAIL(int id_gatepass)
        {
            List<GatePassModel> detList = new List<GatePassModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GATEPASS_DETAIL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);

                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            GatePassModel row = new GatePassModel();
                            row.id_gatepass = reader["id_gatepass"] != DBNull.Value ? Convert.ToInt32(reader["id_gatepass"]) : (int?)null;
                            row.gatepass_no = reader["gatepass_no"] != DBNull.Value ? reader["gatepass_no"].ToString() : null;
                            row.category = reader["category"] != DBNull.Value ? reader["category"].ToString() : null;
                            row.asset_no = reader["asset_no"] != DBNull.Value ? reader["asset_no"].ToString() : null;
                            row.asset_subnumber = reader["asset_subnumber"] != DBNull.Value ? reader["asset_subnumber"].ToString() : null;
                            row.asset_desc = reader["asset_desc"] != DBNull.Value ? reader["asset_desc"].ToString() : null;
                            row.asset_class = reader["asset_class"] != DBNull.Value ? reader["asset_class"].ToString() : null;
                            row.cost_center = reader["cost_center"] != DBNull.Value ? reader["cost_center"].ToString() : null;
                            row.capitalized_on = reader["capitalized_on"] != DBNull.Value ? DateTime.Parse(reader["capitalized_on"].ToString()).ToString("dd MMM yyyy") : null;
                            row.naming_output = reader["naming_output"] != DBNull.Value ? reader["naming_output"].ToString() : null;
                            row.actual_return_date = reader["actual_return_date"] != DBNull.Value ? DateTime.Parse(reader["actual_return_date"].ToString()).ToString("dd MMM yyyy") : null;
                            row.hs_code = reader["hs_code"] != DBNull.Value ? Convert.ToDecimal(reader["hs_code"]) : null;
                            row.box_no = reader["box_no"] != DBNull.Value ? reader["box_no"].ToString() : null;
                            row.id_proforma = reader["id_proforma"] != DBNull.Value ? Convert.ToInt32(reader["id_proforma"]) : (int?)null;
                            row.proforma_fin_status = reader["proforma_fin_status"] != DBNull.Value ? reader["proforma_fin_status"].ToString() : null;
                            row.shipping_status = reader["shipping_status"] != DBNull.Value ? reader["shipping_status"].ToString() : null;
                            row.supporting_documents = reader["supporting_documents"] != DBNull.Value ? reader["supporting_documents"].ToString() : null;
                            row.proforma_no = reader["proforma_no"] != DBNull.Value ? reader["proforma_no"].ToString() : null;
                            row.attn_to = reader["attn_to"] != DBNull.Value ? reader["attn_to"].ToString() : null;
                            row.street = reader["street"] != DBNull.Value ? reader["street"].ToString() : null;
                            row.city = reader["city"] != DBNull.Value ? reader["city"].ToString() : null;
                            row.country = reader["country"] != DBNull.Value ? reader["country"].ToString() : null;
                            row.postal_code = reader["postal_code"] != DBNull.Value ? reader["postal_code"].ToString() : null;
                            row.phone_no = reader["phone_no"] != DBNull.Value ? reader["phone_no"].ToString() : null;
                            row.email = reader["email"] != DBNull.Value ? reader["email"].ToString() : null;
                            row.coo = reader["coo"] != DBNull.Value ? reader["coo"].ToString() : null;
                            row.ship_mode = reader["ship_mode"] != DBNull.Value ? reader["ship_mode"].ToString() : null;
                            row.courier_charges = reader["courier_charges"] != DBNull.Value ? reader["courier_charges"].ToString() : null;
                            row.courier_name = reader["courier_name"] != DBNull.Value ? reader["courier_name"].ToString() : null;
                            row.courier_account_no = reader["courier_account_no"] != DBNull.Value ? reader["courier_account_no"].ToString() : null;
                            row.freight_charges = reader["freight_charges"] != DBNull.Value ? reader["freight_charges"].ToString() : null;
                            row.incoterms = reader["incoterms"] != DBNull.Value ? reader["incoterms"].ToString() : null;
                            row.invoice_payment = reader["invoice_payment"] != DBNull.Value ? reader["invoice_payment"].ToString() : null;
                            row.file_attach = reader["file_attach"] != DBNull.Value ? reader["file_attach"].ToString() : null;
                            row.file_support = reader["file_support"] != DBNull.Value ? reader["file_support"].ToString() : null;
                            row.file_peb = reader["file_peb"] != DBNull.Value ? reader["file_peb"].ToString() : null;
                            row.id_file = reader["id_file"] != DBNull.Value ? Convert.ToInt32(reader["id_file"]) : (int?)null;
                            row.document_type = reader["document_type"] != DBNull.Value ? reader["document_type"].ToString() : null;
                            row.fin_filename = reader["fin_filename"] != DBNull.Value ? reader["fin_filename"].ToString() : null;
                            row.fin_created_by = reader["fin_created_by"] != DBNull.Value ? reader["fin_created_by"].ToString() : null;
                            row.fin_record_date = reader["fin_record_date"] != DBNull.Value ? Convert.ToDateTime(reader["fin_record_date"]) : (DateTime?)null;
                            row.plant = reader["plant"] != DBNull.Value ? reader["plant"].ToString() : null;
                            row.shipment_date = reader["shipment_date"] != DBNull.Value ? Convert.ToDateTime(reader["shipment_date"]).ToString("dd MMM yyyy") : null;
                            row.dhl_awb = reader["dhl_awb"] != DBNull.Value ? reader["dhl_awb"].ToString() : null;

                            detList.Add(row);
                        }
                    }
                }

                conn.Close();
            }
            return detList;
        }
        public List<GatePassModel> GATEPASS_STATUS(int id_gatepass)
        {
            List<GatePassModel> dataList = new List<GatePassModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GATEPASS_STATUS_V1", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            GatePassModel row = new GatePassModel();
                            row.approver_no = reader["approver_no"].ToString();
                            row.approver = reader["approver"].ToString();
                            row.approval_status = reader["approval_status"].ToString();
                            if (reader["date_approval"] != DBNull.Value)
                            {
                                row.date_approval = DateTime.Parse(reader["date_approval"].ToString()).ToString("dd MMM yyyy");
                            }
                            else
                            {
                                row.date_approval = null;
                            }

                            string baseRemark = reader["remark"].ToString();
                            string delegatedBy = reader["delegated_by"].ToString();
                            string delegatedByName = reader["delegated_by_name"].ToString();

                            if (!string.IsNullOrEmpty(delegatedBy) && !string.IsNullOrEmpty(delegatedByName))
                            {
                                if (!string.IsNullOrEmpty(baseRemark))
                                {
                                    row.remark = baseRemark + "\n\nApproved by: " + delegatedByName;
                                }
                                else
                                {
                                    row.remark = "Approved by: " + delegatedByName;
                                }
                            }
                            else
                            {
                                row.remark = baseRemark;
                            }

                            dataList.Add(row);
                        }
                    }
                }

                conn.Close();
            }
            return dataList;
        }

        public List<GatePassModel> GET_GATEPASS_HEADER(int id_gatepass)
        {
            List<GatePassModel> disList = new List<GatePassModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GATEPASS_HEADER", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            GatePassModel row = new GatePassModel();
                            row.gatepass_no = reader["gatepass_no"].ToString();
                            row.created_by = reader["created_by"].ToString();
                            row.category = reader["category"].ToString();
                            if (reader["return_date"] != DBNull.Value)
                            {
                                row.return_date = DateTime.Parse(reader["return_date"].ToString()).ToString("dd MMM yyyy");
                            }
                            else
                            {
                                row.return_date = "-";
                            }
                            row.create_date = DateTime.Parse(reader["create_date"].ToString()).ToString("dd MMM yyyy");
                            row.remark = reader["remark"].ToString();
                            row.location = reader["location"].ToString();
                            row.new_pic_name = reader["new_pic_name"].ToString();
                            row.type = reader["type"].ToString();
                            row.csr_to = reader["csr_to"].ToString();
                            row.employee_name = reader["employee_name"].ToString();
                            disList.Add(row);
                        }
                    }
                }
                conn.Close();
            }
            return disList;
        }

        public decimal GET_TOTAL_CURRENT_APC(int id_gatepass)
        {
            decimal totalApc = 0;
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_TOTAL_CURRENT_APC", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    var result = cmd.ExecuteScalar();
                    if (result != DBNull.Value)
                    {
                        totalApc = Convert.ToDecimal(result);
                    }
                }
                conn.Close();
            }
            return totalApc;
        }

        public string DELETE_TEMP_GATEPASS(int id_temp)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("DELETE_TEMP_GATEPASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_temp", id_temp);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string UPDATE_APPROVAL_GATEPASS(int id_gatepass, string approval_level, string approval_status, string approval_by, string remark)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();

                string actualApprover = GetGatepassApproverAtLevel(id_gatepass, approval_level);
                bool isDelegated = IsDelegatedApprover(approval_by, actualApprover);
                string delegatedBy = null;

                if (isDelegated)
                {
                    delegatedBy = approval_by;
                }
                else
                {
                    actualApprover = approval_by;
                    delegatedBy = null;
                }

                using (SqlCommand cmd = new SqlCommand("UPDATE_APPROVAL_GATEPASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@approval_level", approval_level);
                    cmd.Parameters.AddWithValue("@approval_status", approval_status);
                    cmd.Parameters.AddWithValue("@approval_by", actualApprover);
                    cmd.Parameters.AddWithValue("@remark", remark);
                    cmd.Parameters.AddWithValue("@delegated_by", delegatedBy ?? (object)DBNull.Value);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string CONFIRM_GATEPASS_RETURN(int id_gatepass, string approval_level, string approval_status, string approval_by, string remark)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();

                string actualConfirmer = GetGatepassConfirmerHOD(id_gatepass);
                bool isDelegated = IsDelegatedApprover(approval_by, actualConfirmer);
                string delegatedBy = null;

                if (isDelegated)
                {
                    delegatedBy = approval_by;
                }
                else
                {
                    actualConfirmer = approval_by;
                    delegatedBy = null;
                }

                using (SqlCommand cmd = new SqlCommand("CONFIRM_GATEPASS_RETURN_V1", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@approval_level", approval_level);
                    cmd.Parameters.AddWithValue("@approval_status", approval_status);
                    cmd.Parameters.AddWithValue("@approval_by", actualConfirmer);
                    cmd.Parameters.AddWithValue("@remark", remark);
                    cmd.Parameters.AddWithValue("@delegated_by", delegatedBy ?? (object)DBNull.Value);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }
        public string GET_STATUS_GATEPASS(int id_gatepass)
        {
            string status_gatepass = string.Empty;
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_STATUS_GATEPASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    status_gatepass = (string)cmd.ExecuteScalar();
                }
                conn.Close();
            }
            return status_gatepass;
        }

        public string UPDATE_SECURITY_GATEPASS(int id_gatepass, string approval_security_status, string security_name, string approval_remark)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_SECURITY_GATEPASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@approval_security_status", approval_security_status);
                    cmd.Parameters.AddWithValue("@security_name", security_name);
                    cmd.Parameters.AddWithValue("@approval_remark", approval_remark);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string UPDATE_GATEPASS_RETURN(int id_gatepass, string approval_security_status, string security_name, string approval_remark)
        {
            string gatepass_no = "";
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_GATEPASS_RETURN", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@approval_security_status", approval_security_status);
                    cmd.Parameters.AddWithValue("@security_name", security_name);
                    cmd.Parameters.AddWithValue("@approval_remark", approval_remark);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            gatepass_no = reader["gatepass_no"].ToString() ?? "";
                        }
                    }
                    return gatepass_no;
                }
            }
        }

        public string UPDATE_FINANCE_GATEPASS(int id_gatepass, string approval_status, string approval_remark, string sesa_id)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_FINANCE_GATEPASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@approval_status", approval_status);
                    cmd.Parameters.AddWithValue("@approval_remark", approval_remark);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }
        public string DELETE_PROFORMA_FIN_FILE(int id_file, int id_gatepass)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("DELETE_PROFORMA_FIN_FILE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_file", id_file);
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string result = reader["result"].ToString();
                            string filename = reader["filename"].ToString();
                            if (result == "success" && !string.IsNullOrEmpty(filename))
                            {
                                string filePath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Upload", "Proforma_Finance", filename);
                                if (System.IO.File.Exists(filePath))
                                {
                                    System.IO.File.Delete(filePath);
                                }
                            }
                            return result;
                        }
                    }
                }
                conn.Close();
            }
            return "Error: No result returned";
        }

        public string UPLOAD_PROFORMA_FIN_FILES(int id_gatepass, string document_type, string filename, string created_by)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("INSERT_PROFORMA_FIN_FILES", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@document_type", document_type);
                    cmd.Parameters.AddWithValue("@filename", filename);
                    cmd.Parameters.AddWithValue("@created_by", created_by);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return reader["result"].ToString();
                        }
                    }
                }
                conn.Close();
            }
            return "error;Unknown error";
        }

        public string UPDATE_PROFORMA_FIN_STATUS(int id_gatepass, string completed_by_sesa_id = null)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("UPDATE_PROFORMA_FIN_STATUS", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        cmd.Parameters.AddWithValue("@completed_by_sesa_id", completed_by_sesa_id ?? (object)DBNull.Value);

                        conn.Open();
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                return reader["result"].ToString();
                            }
                        }
                    }
                }
                return "success";
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
        }

        public List<ProformaFileModel> GET_PROFORMA_FIN_FILES(int id_gatepass)
        {
            List<ProformaFileModel> fileList = new List<ProformaFileModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_PROFORMA_FIN_FILES", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            ProformaFileModel file = new ProformaFileModel();
                            file.id_file = Convert.ToInt32(reader["id_file"]);
                            file.document_type = reader["document_type"].ToString();
                            file.filename = reader["filename"].ToString();
                            file.record_date = Convert.ToDateTime(reader["record_date"]);
                            fileList.Add(file);
                        }
                    }
                }
                conn.Close();
            }
            return fileList;
        }

        public PebFileModel GET_PEB_FILE(int id_gatepass)
        {
            PebFileModel file = null;
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_PEB_FILE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        file = new PebFileModel();
                        string filenames = reader["file_peb"].ToString();
                        file.filename = filenames;
                        file.filenames = string.IsNullOrEmpty(filenames) ? new List<string>() : filenames.Split(';').ToList();
                    }
                }
                conn.Close();
            }
            return file;
        }

        public string UPLOAD_PEB_FILE(int id_gatepass, string newFilenames, string created_by)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("INSERT_PROFORMA_PEB_FILE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@newFilenames", newFilenames);
                    cmd.Parameters.AddWithValue("@created_by", created_by ?? (object)DBNull.Value);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string result = reader["result"].ToString();
                            if (result == "success")
                            {
                                return "success";
                            }
                            else
                            {
                                string errorMessage = reader["error_message"]?.ToString() ?? "Unknown error";
                                return $"error;{errorMessage}";
                            }
                        }
                    }
                    return "error;No result returned";
                }
            }
        }

        public string DELETE_PEB_FILE(int id_gatepass, string filenameToDelete = null)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("DELETE_PEB_FILE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@filenameToDelete", string.IsNullOrEmpty(filenameToDelete) ? (object)DBNull.Value : filenameToDelete);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string result = reader["result"].ToString();
                            string filesToDeleteStr = reader["files_to_delete"].ToString();
                            if (result == "success" && !string.IsNullOrEmpty(filesToDeleteStr))
                            {
                                List<string> filesToDelete = filesToDeleteStr.Split(';').ToList();
                                foreach (string filename in filesToDelete)
                                {
                                    if (!string.IsNullOrEmpty(filename))
                                    {
                                        string filePath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Upload", "PEB", filename);
                                        if (System.IO.File.Exists(filePath))
                                        {
                                            System.IO.File.Delete(filePath);
                                        }
                                    }
                                }
                            }
                            return result;
                        }
                    }
                }
                conn.Close();
            }
            return "Error: No result returned";
        }

        public string COMPLETE_PEB_UPLOAD(int id_gatepass, string completed_by_sesa_id = null)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("UPDATE_PROFORMA_PEB_STATUS", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        cmd.Parameters.AddWithValue("@completed_by_sesa_id", completed_by_sesa_id ?? (object)DBNull.Value);

                        conn.Open();
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                return reader["result"].ToString();
                            }
                        }
                    }
                }
                return "success";
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
        }
        public string ADD_NEW_USER(string sesa_id, string name, string email, string department, string plant, string level, string role, string manager_sesa_id)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("ADD_NEW_USER", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.Parameters.AddWithValue("@name", name);
                    cmd.Parameters.AddWithValue("@email", email);
                    cmd.Parameters.AddWithValue("@department", department);
                    cmd.Parameters.AddWithValue("@plant", plant);
                    cmd.Parameters.AddWithValue("@level", level);
                    cmd.Parameters.AddWithValue("@role", role);
                    cmd.Parameters.AddWithValue("@manager_sesa_id", manager_sesa_id);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string status = reader["Status"].ToString();
                            string message = reader["Message"].ToString();
                            return status + ";" + message;
                        }
                    }
                }
            }
            return "error;Unknown error occurred.";
        }

        public string NEW_PIC_CONFIRM(string approval_pic_status, string sesa_id, string approval_remark, int id_gatepass_pic)
        {
            string gatepass_no = "";
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_NEWPIC_CONFIRM_V1", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass_pic);
                    cmd.Parameters.AddWithValue("@approval_pic_status", approval_pic_status);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.Parameters.AddWithValue("@approval_remark", approval_remark);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            gatepass_no = reader["gatepass_no"].ToString() ?? "";
                        }
                    }
                    return gatepass_no;
                }
            }
        }

        public string UPDATE_USER(string sesa_id, string name, string email, string department, string plant, string level, string role, string manager_sesa_id)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_USER_DATA", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.Parameters.AddWithValue("@name", name);
                    cmd.Parameters.AddWithValue("@email", email);
                    cmd.Parameters.AddWithValue("@department", department);
                    cmd.Parameters.AddWithValue("@plant", plant);
                    cmd.Parameters.AddWithValue("@level", level);
                    cmd.Parameters.AddWithValue("@role", role);
                    cmd.Parameters.AddWithValue("@manager_sesa_id", manager_sesa_id);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }
        public List<AssetListModel> GET_ASSET_GATEPASS(string asset_search)
        {
            List<AssetListModel> assetList = new List<AssetListModel>();

            using (var conn = new SqlConnection(ConnectionString))
            {
                conn.Open();

                using (var cmd = new SqlCommand("GET_SEARCH_ASSET_GATEPASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@asset_search", asset_search);

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var row = new AssetListModel
                            {
                                asset_no = reader["asset_no"].ToString(),
                                asset_subnumber = reader["asset_subnumber"].ToString(),
                                asset_desc = reader["asset_desc"].ToString(),
                                status_gatepass = reader["status_gatepass"].ToString()
                            };
                            assetList.Add(row);
                        }
                    }
                }
            }

            return assetList;
        }

        public List<GatePassModel> GetGatepassInfo(string id_gatepass)
        {
            List<GatePassModel> gpDetail = new List<GatePassModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                string query = "GET_GATEPASS_INFO";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);

                conn.Open();
                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        var data = new GatePassModel();
                        data.gatepass_no = reader["gatepass_no"].ToString();
                        data.category = reader["category"].ToString();
                        data.type = reader["type"].ToString();
                        data.employee_name = reader["employee_name"].ToString();
                        data.csr_to = reader["csr_to"].ToString();
                        data.return_date = Utility.ConvertToDateFormat(Convert.ToString(reader["create_date"]));
                        data.vendor_code = reader["vendor_code"].ToString();
                        data.vendor_name = reader["vendor_name"].ToString();
                        data.vendor_address = reader["vendor_address"].ToString();
                        data.recipient_phone = reader["recipient_phone"].ToString();
                        data.recipient_email = reader["recipient_email"].ToString();
                        data.vehicle_no = reader["vehicle_no"].ToString();
                        data.driver_name = reader["driver_name"].ToString();
                        data.security_guard = reader["security_guard"].ToString();
                        data.remark = reader["remark"].ToString();
                        data.image_before = reader["image_before"].ToString();
                        data.image_after = reader["image_after"].ToString();
                        data.created_by = reader["created_by"].ToString();
                        data.new_pic = reader["new_pic"].ToString();
                        data.location = reader["location"].ToString();
                        data.create_date = Utility.ConvertToDateFormat(Convert.ToString(reader["create_date"]));
                        gpDetail.Add(data);
                    }
                    conn.Close();
                }
            }
            return gpDetail;
        }

        public List<ApprovalModel> GetGatepassApproval(string id_gatepass)
        {
            List<ApprovalModel> dashData = new List<ApprovalModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                string query = "GET_GATEPASS_APPROVAL";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);

                conn.Open();
                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        var data = new ApprovalModel();
                        data.approval_level = reader["approval_level"].ToString();
                        data.approval_by = reader["approval_by"].ToString();
                        data.approval_status = reader["approval_status"].ToString();
                        data.approval_name = reader["approval_name"].ToString();
                        data.date_approval = reader["date_approval"].ToString();
                        data.remark = reader["remark"].ToString();

                        dashData.Add(data);
                    }
                    conn.Close();
                    return dashData;
                }

                else
                {
                    conn.Close();
                    return null;
                }
            }
        }

        public List<GPOpenModel> GP_OPEN_DETAIL(string department)
        {
            var assetList = new List<GPOpenModel>();

            using (var conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (var cmd = new SqlCommand("GP_OPEN_DETAIL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@department", department);

                    using (var reader = cmd.ExecuteReader())
                    {
                        var dataTable = new DataTable();
                        dataTable.Load(reader);

                        foreach (DataRow row in dataTable.Rows)
                        {
                            var model = new GPOpenModel
                            {
                                id_gatepass = row["id_gatepass"].ToString(),
                                gatepass_no = row["gatepass_no"].ToString(),
                                category = row["category"].ToString(),
                                gatepass_date = row["gatepass_date"].ToString(),
                                requestor_name = row["requestor_name"].ToString(),
                                return_date = row["return_date"].ToString(),
                                destination = row["destination"].ToString(),
                                new_pic_name = row["new_pic_name"].ToString(),
                                status_gatepass = row["status_gatepass"].ToString()
                            };
                            assetList.Add(model);
                        }
                    }
                }
            }

            return assetList;
        }

        public List<GPOpenModel> GP_PROGRESS_DETAIL(string status, string category)
        {
            List<GPOpenModel> assetList = new List<GPOpenModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GP_PROGRESS_DETAIL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@status", status);
                    cmd.Parameters.AddWithValue("@category", category);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            GPOpenModel row = new GPOpenModel();
                            row.id_gatepass = reader["id_gatepass"].ToString();
                            row.gatepass_no = reader["gatepass_no"].ToString();
                            row.category = reader["category"].ToString();
                            row.gatepass_date = reader["gatepass_date"].ToString();
                            row.requestor_name = reader["requestor_name"].ToString();
                            row.return_date = reader["return_date"].ToString();
                            row.destination = reader["destination"].ToString();
                            row.new_pic_name = reader["new_pic_name"].ToString();
                            row.status_gatepass = reader["status_gatepass"].ToString();
                            assetList.Add(row);
                        }
                    }
                }

                conn.Close();
            }
            return assetList;
        }

        public List<GPOpenModel> GP_OPEN_NON_ASSET_DETAIL(string department)
        {
            List<GPOpenModel> assetList = new List<GPOpenModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GP_OPEN_NON_ASSET_DETAIL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@department", department);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            GPOpenModel row = new GPOpenModel();
                            row.id_gatepass = reader["id_gatepass"].ToString();
                            row.gatepass_no = reader["gatepass_no"].ToString();
                            row.category = reader["category"].ToString();
                            row.gatepass_date = reader["gatepass_date"].ToString();
                            row.requestor_name = reader["requestor_name"].ToString();
                            row.return_date = reader["return_date"].ToString();
                            row.destination = reader["destination"].ToString();
                            row.new_pic_name = "-";
                            row.status_gatepass = reader["status_gatepass"].ToString();
                            assetList.Add(row);
                        }
                    }
                }
                conn.Close();
            }
            return assetList;
        }

        public List<GPOpenModel> GP_PROGRESS_NON_ASSET_DETAIL(string status, string category)
        {
            List<GPOpenModel> assetList = new List<GPOpenModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GP_PROGRESS_NON_ASSET_DETAIL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@status", status);
                    cmd.Parameters.AddWithValue("@category", category);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            GPOpenModel row = new GPOpenModel();
                            row.id_gatepass = reader["id_gatepass"].ToString();
                            row.gatepass_no = reader["gatepass_no"].ToString();
                            row.category = reader["category"].ToString();
                            row.gatepass_date = reader["gatepass_date"].ToString();
                            row.requestor_name = reader["requestor_name"].ToString();
                            row.return_date = reader["return_date"].ToString();
                            row.destination = reader["destination"].ToString();
                            row.new_pic_name = "-";
                            row.status_gatepass = reader["status_gatepass"].ToString();
                            assetList.Add(row);
                        }
                    }
                }
                conn.Close();
            }
            return assetList;
        }

        public List<GPModel> GET_GP_DETAIL_AGING(string department, string aging_days)
        {
            List<GPModel> assetList = new List<GPModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GP_DETAIL_AGING", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@department", department);
                    cmd.Parameters.AddWithValue("@aging_days", aging_days);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            GPModel row = new GPModel();
                            row.id_gatepass = reader["id_gatepass"].ToString();
                            row.gatepass_no = reader["gatepass_no"].ToString();
                            row.category = reader["category"].ToString();
                            row.create_date = reader["create_date"].ToString();
                            row.requestor_name = reader["requestor_name"].ToString();
                            row.return_date = reader["return_date"].ToString();
                            row.status_gatepass = reader["status_gatepass"].ToString();
                            row.department = reader["department"].ToString();
                            row.aging_days = reader["aging_days"].ToString();
                            assetList.Add(row);
                        }
                    }
                }

                conn.Close();
            }
            return assetList;
        }

        public List<GPModel> GET_GP_DETAIL_RETURN(string type)
        {
            List<GPModel> assetList = new List<GPModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GP_DETAIL_RETURN", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@type", type);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            GPModel row = new GPModel();
                            row.id_gatepass = reader["id_gatepass"].ToString();
                            row.gatepass_no = reader["gatepass_no"].ToString();
                            row.category = reader["category"].ToString();
                            row.create_date = reader["create_date"].ToString();
                            row.requestor_name = reader["requestor_name"].ToString();
                            row.return_date = reader["return_date"].ToString();
                            row.status_gatepass = reader["status_gatepass"].ToString();
                            row.department = reader["department"].ToString();
                            row.aging_days = reader["aging_days"].ToString();
                            assetList.Add(row);
                        }
                    }
                }

                conn.Close();
            }
            return assetList;
        }

        public List<GPModel> GET_GP_DETAIL_TOTAL(string type)
        {
            List<GPModel> assetList = new List<GPModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GP_DETAIL_TOTAL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@type", type);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            GPModel row = new GPModel();
                            row.id_gatepass = reader["id_gatepass"].ToString();
                            row.gatepass_no = reader["gatepass_no"].ToString();
                            row.category = reader["category"].ToString();
                            row.create_date = reader["create_date"].ToString();
                            row.requestor_name = reader["requestor_name"].ToString();
                            row.return_date = reader["return_date"].ToString();
                            row.status_gatepass = reader["status_gatepass"].ToString();
                            row.department = reader["department"].ToString();
                            row.aging_days = reader["aging_days"].ToString();
                            assetList.Add(row);
                        }
                    }
                }

                conn.Close();
            }
            return assetList;
        }

        public List<GPModel> GET_GP_NA_DETAIL_AGING(string department, string aging_days)
        {
            List<GPModel> assetList = new List<GPModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GP_NA_DETAIL_AGING", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@department", department);
                    cmd.Parameters.AddWithValue("@aging_days", aging_days);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            GPModel row = new GPModel();
                            row.id_gatepass = reader["id_gatepass"].ToString();
                            row.gatepass_no = reader["gatepass_no"].ToString();
                            row.category = reader["category"].ToString();
                            row.create_date = reader["create_date"].ToString();
                            row.requestor_name = reader["requestor_name"].ToString();
                            row.return_date = reader["return_date"].ToString();
                            row.status_gatepass = reader["status_gatepass"].ToString();
                            row.department = reader["department"].ToString();
                            row.aging_days = reader["aging_days"].ToString();
                            assetList.Add(row);
                        }
                    }
                }
                conn.Close();
            }
            return assetList;
        }

        public List<GPModel> GET_GP_NA_DETAIL_RETURN(string type)
        {
            List<GPModel> assetList = new List<GPModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GP_NA_DETAIL_RETURN", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@type", type);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            GPModel row = new GPModel();
                            row.id_gatepass = reader["id_gatepass"].ToString();
                            row.gatepass_no = reader["gatepass_no"].ToString();
                            row.category = reader["category"].ToString();
                            row.create_date = reader["create_date"].ToString();
                            row.requestor_name = reader["requestor_name"].ToString();
                            row.return_date = reader["return_date"].ToString();
                            row.status_gatepass = reader["status_gatepass"].ToString();
                            row.department = reader["department"].ToString();
                            row.aging_days = reader["aging_days"].ToString();
                            assetList.Add(row);
                        }
                    }
                }
                conn.Close();
            }
            return assetList;
        }

        public List<GPModel> GET_GP_NA_DETAIL_TOTAL(string type)
        {
            List<GPModel> assetList = new List<GPModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GP_NA_DETAIL_TOTAL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@type", type);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            GPModel row = new GPModel();
                            row.id_gatepass = reader["id_gatepass"].ToString();
                            row.gatepass_no = reader["gatepass_no"].ToString();
                            row.category = reader["category"].ToString();
                            row.create_date = reader["create_date"].ToString();
                            row.requestor_name = reader["requestor_name"].ToString();
                            row.return_date = reader["return_date"].ToString();
                            row.status_gatepass = reader["status_gatepass"].ToString();
                            row.department = reader["department"].ToString();
                            row.aging_days = reader["aging_days"].ToString();
                            assetList.Add(row);
                        }
                    }
                }
                conn.Close();
            }
            return assetList;
        }

        public List<AssetCountListModel> GET_PROPOSE_DISPOSE(List<int> selectedIds)
        {
            List<AssetCountListModel> materials = new List<AssetCountListModel>();

            if (selectedIds == null || !selectedIds.Any())
            {
                return materials;
            }

            string selectedIdsStr = string.Join(",", selectedIds);

            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_PROPOSE_DISPOSE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@selectedIds", selectedIdsStr);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            AssetCountListModel mat = new AssetCountListModel();
                            mat.asset_no = reader["asset_no"].ToString();
                            mat.asset_subnumber = reader["asset_subnumber"].ToString();
                            mat.asset_desc = reader["asset_desc"].ToString();
                            mat.curr_bk_val = Convert.ToDouble(reader["curr_bk_val"]);
                            materials.Add(mat);
                        }
                    }
                }
                conn.Close();
            }
            return materials;
        }

        public string GetURL(string file_desc)
        {
            string url = "";
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_URL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@file_desc", file_desc);
                    url = (string)cmd.ExecuteScalar();
                }
                conn.Close();
            }
            return url;
        }

        public List<GatePassModel> GET_LIST_IMG(string id_gatepass)
        {
            List<GatePassModel> materials = new List<GatePassModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_LIST_IMG", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", int.Parse(id_gatepass));
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            GatePassModel mat = new GatePassModel
                            {
                                id_gatepass = reader["id_gatepass"] != DBNull.Value ? (int?)Convert.ToInt32(reader["id_gatepass"]) : null,
                                id_detail = reader["id_detail"].ToString(),
                                asset_no = reader["asset_no"].ToString(),
                                asset_subnumber = reader["asset_subnumber"].ToString(),
                                image_after = reader["gp_after"].ToString(),
                                return_date = reader["return_date"].ToString(),
                            };
                            materials.Add(mat);
                        }
                    }
                }
                conn.Close();
            }
            return materials;
        }

        public List<ShippingAssetModel> GetShippingAssets(int id_gatepass, string sesa_id = "")
        {
            List<ShippingAssetModel> assetList = new List<ShippingAssetModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_SHIPPING_ASSETS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id ?? (object)DBNull.Value);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        ShippingAssetModel row = new ShippingAssetModel();
                        row.id_detail = Convert.ToInt32(reader["id_detail"]);
                        row.asset_no = reader["asset_no"].ToString();
                        row.asset_subnumber = reader["asset_subnumber"].ToString();
                        row.asset_desc = reader["asset_desc"].ToString();
                        row.hs_code = reader["hs_code"] == DBNull.Value ? null : Convert.ToDecimal(reader["hs_code"]);
                        row.id_box = Convert.ToInt32(reader["id_box"]);
                        row.is_assigned = Convert.ToBoolean(reader["is_assigned"]);
                        row.is_selected = false;
                        assetList.Add(row);
                    }
                }
                conn.Close();
            }
            return assetList;
        }

        public string CreateTempBox(TempShippingBoxModel box)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("INSERT_SHIPPING_TEMP_BOX", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@created_by", box.created_by ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@id_gatepass", box.id_gatepass ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@box_no", box.box_no ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@asset_list", box.asset_list ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@length_cm", box.length_cm ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@width_cm", box.width_cm ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@height_cm", box.height_cm ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@gross_weight_kg", box.gross_weight_kg ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@net_weight_kg", box.net_weight_kg ?? (object)DBNull.Value);

                    SqlParameter newIdParam = new SqlParameter("@new_id", SqlDbType.Int);
                    newIdParam.Direction = ParameterDirection.Output;
                    cmd.Parameters.Add(newIdParam);

                    SqlParameter resultParam = new SqlParameter("@result", SqlDbType.NVarChar, 500);
                    resultParam.Direction = ParameterDirection.Output;
                    cmd.Parameters.Add(resultParam);

                    cmd.ExecuteNonQuery();

                    string result = resultParam.Value?.ToString() ?? "success";

                    if (result == "success")
                    {
                        int newId = Convert.ToInt32(newIdParam.Value);
                        return $"success;{newId}";
                    }

                    return result;
                }
            }
        }

        public List<TempShippingBoxModel> GetTempBoxes(string sesa_id, int id_gatepass)
        {
            List<TempShippingBoxModel> boxList = new List<TempShippingBoxModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_SHIPPING_TEMP_BOXES", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@created_by", sesa_id);
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);

                    using SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        TempShippingBoxModel row = new TempShippingBoxModel();
                        row.id_temp = Convert.ToInt32(reader["id_temp"]);
                        row.created_by = reader["created_by"].ToString();
                        row.id_gatepass = reader["id_gatepass"] != DBNull.Value ? Convert.ToInt32(reader["id_gatepass"]) : 0;
                        row.box_no = reader["box_no"].ToString();
                        row.asset_list = reader["asset_list"].ToString();
                        row.length_cm = reader["length_cm"] != DBNull.Value ? Convert.ToDecimal(reader["length_cm"]) : 0;
                        row.width_cm = reader["width_cm"] != DBNull.Value ? Convert.ToDecimal(reader["width_cm"]) : 0;
                        row.height_cm = reader["height_cm"] != DBNull.Value ? Convert.ToDecimal(reader["height_cm"]) : 0;
                        row.gross_weight_kg = reader["gross_weight_kg"] != DBNull.Value ? Convert.ToDecimal(reader["gross_weight_kg"]) : 0;
                        row.net_weight_kg = reader["net_weight_kg"] != DBNull.Value ? Convert.ToDecimal(reader["net_weight_kg"]) : 0;
                        row.create_date = reader["record_date"] != DBNull.Value ? Convert.ToDateTime(reader["record_date"]) : DateTime.Now;

                        boxList.Add(row);
                    }
                }
                conn.Close();
            }
            return boxList;
        }

        public string DeleteTempBox(int id_temp)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("DELETE_SHIPPING_TEMP_BOX", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_temp", id_temp);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string ClearTempBoxes(string sesa_id, int id_gatepass)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("CLEAR_SHIPPING_TEMP_BOXES", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@created_by", sesa_id);
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string SaveShipping(ShippingCreateViewModel model, string sesa_id)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("INSERT_SHIPPING_V2", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", model.id_gatepass);
                    cmd.Parameters.AddWithValue("@plant", model.plant);
                    cmd.Parameters.AddWithValue("@shipment_date", model.shipment_date);
                    string shipmentTypeToSave = !string.IsNullOrEmpty(model.shipment_type) ? model.shipment_type : model.shipment_type;
                    cmd.Parameters.AddWithValue("@shipment_type", shipmentTypeToSave ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@dhl_awb", model.dhl_awb ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    var tempBoxes = GetTempBoxes(sesa_id, model.id_gatepass);
                    string tempBoxesJson = System.Text.Json.JsonSerializer.Serialize(tempBoxes);
                    cmd.Parameters.AddWithValue("@temp_boxes_json", tempBoxesJson);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string result = reader["result"].ToString();
                            string shippingId = reader["shipping_id"].ToString();
                            string proformaNo = reader["proforma_no"].ToString();
                            return $"{result};{shippingId};{proformaNo}";
                        }
                    }
                    return "error;Unknown error";
                }
            }
        }

        public string SaveShippingForce(ShippingCreateViewModel model, string sesa_id, bool forceDelete = false)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("INSERT_SHIPPING_FORCE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", model.id_gatepass);
                    cmd.Parameters.AddWithValue("@plant", model.plant);
                    cmd.Parameters.AddWithValue("@shipment_date", model.shipment_date);
                    cmd.Parameters.AddWithValue("@shipment_type", model.shipment_type ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@dhl_awb", model.dhl_awb ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.Parameters.AddWithValue("@force_delete", forceDelete);

                    var tempBoxes = GetTempBoxes(sesa_id, model.id_gatepass);
                    string tempBoxesJson = System.Text.Json.JsonSerializer.Serialize(tempBoxes);
                    cmd.Parameters.AddWithValue("@temp_boxes_json", tempBoxesJson);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string result = reader["result"].ToString();
                            string shippingId = reader["shipping_id"].ToString();
                            string proformaNo = reader["proforma_no"].ToString();
                            return $"{result};{shippingId};{proformaNo}";
                        }
                    }
                    return "error;Unknown error";
                }
            }
        }

        public ShippingBoxDisplayModel GetCombinedShippingBoxes(int id_gatepass, string sesa_id)
        {
            var result = new ShippingBoxDisplayModel();
            result.saved_boxes = GetSavedShippingBoxes(id_gatepass);
            result.temp_boxes = GetTempBoxes(sesa_id, id_gatepass);
            return result;
        }

        public string SubmitShipping(int id_gatepass, string sesa_id)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("UPDATE_SHIPPING", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        cmd.Parameters.AddWithValue("@sesa_id", sesa_id);

                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows && reader.Read())
                            {
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    if (reader.GetName(i).Equals("result", StringComparison.OrdinalIgnoreCase))
                                    {
                                        return reader["result"].ToString();
                                    }
                                }

                                return "success";
                            }
                            else
                            {
                                return "error;No result returned from stored procedure";
                            }
                        }
                    }
                }
                catch (SqlException sqlEx)
                {
                    return $"error;SQL Error: {sqlEx.Message}";
                }
                catch (Exception ex)
                {
                    return $"error;{ex.Message}";
                }
            }
        }

        public List<ShippingBoxModel> GetSavedShippingBoxes(int id_gatepass)
        {
            List<ShippingBoxModel> boxList = new List<ShippingBoxModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_SAVED_SHIPPING_BOXES", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        ShippingBoxModel row = new ShippingBoxModel();
                        row.id_box = Convert.ToInt32(reader["id_box"]);
                        row.box_no = reader["box_no"].ToString();
                        row.length_cm = reader["length_cm"] != DBNull.Value ? Convert.ToDecimal(reader["length_cm"]) : 0;
                        row.width_cm = reader["width_cm"] != DBNull.Value ? Convert.ToDecimal(reader["width_cm"]) : 0;
                        row.height_cm = reader["height_cm"] != DBNull.Value ? Convert.ToDecimal(reader["height_cm"]) : 0;
                        row.gross_weight_kg = reader["gross_weight_kg"] != DBNull.Value ? Convert.ToDecimal(reader["gross_weight_kg"]) : 0;
                        row.net_weight_kg = reader["net_weight_kg"] != DBNull.Value ? Convert.ToDecimal(reader["net_weight_kg"]) : 0;
                        row.asset_count = Convert.ToInt32(reader["asset_count"]);
                        boxList.Add(row);
                    }
                }
                conn.Close();
            }
            return boxList;
        }

        public ShippingCreateViewModel GetShippingData(int id_gatepass, string sesa_id = "")
        {
            var model = new ShippingCreateViewModel();
            model.id_gatepass = id_gatepass;

            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_SHIPPING_DATA", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        model.gatepass_no = reader["gatepass_no"].ToString();
                        model.shipping_plant = reader["shipping_plant"].ToString();
                        model.plant = reader["plant"].ToString();
                        model.dhl_awb = reader["dhl_awb"].ToString();
                        if (reader["shipment_date"] != DBNull.Value)
                            model.shipment_date = Convert.ToDateTime(reader["shipment_date"]);

                        model.is_shipping_saved = reader["id_shipping"] != DBNull.Value;
                        string courierName = reader["courier_name"]?.ToString();
                        string existingShipmentType = reader["shipment_type"]?.ToString();

                        if (!string.IsNullOrEmpty(existingShipmentType))
                        {
                            model.shipment_type = existingShipmentType;
                        }
                        else if (!string.IsNullOrEmpty(courierName))
                        {
                            model.shipment_type = courierName;
                        }
                        else
                        {
                            model.shipment_type = "";
                        }

                        if (model.is_shipping_saved)
                        {
                            string dbPlant = reader["plant"]?.ToString();
                            DateTime? dbShipmentDate = reader["shipment_date"] != DBNull.Value ? Convert.ToDateTime(reader["shipment_date"]) : null;
                            string dbShipmentType = model.shipment_type;
                            string dbDhlAwb = reader["dhl_awb"]?.ToString();

                            model.is_data_complete_in_db = !string.IsNullOrEmpty(dbPlant) &&
                                                          dbShipmentDate.HasValue &&
                                                          !string.IsNullOrEmpty(dbShipmentType) &&
                                                          !string.IsNullOrEmpty(dbDhlAwb);
                        }
                        else
                        {
                            model.is_data_complete_in_db = false;
                        }
                    }
                }

                model.assets = GetShippingAssets(id_gatepass, sesa_id);

                if (model.is_shipping_saved)
                {
                    model.saved_boxes = GetSavedShippingBoxes(id_gatepass);
                }
                else
                {
                    model.temp_boxes = GetTempBoxes(sesa_id, id_gatepass);
                }

                conn.Close();
            }

            return model;
        }

        public string DeleteSavedShippingBox(int id_box)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("DELETE_SAVED_SHIPPING_BOX", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_box", id_box);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return reader["result"].ToString();
                        }
                    }
                    return "error;Unknown error";
                }
            }
        }

        public ShippingExportModel GetShippingExportData(int id_gatepass)
        {
            var model = new ShippingExportModel();

            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();

                using (SqlCommand cmd = new SqlCommand("GET_SHIPPING_EXPORT_DATA", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        model.proforma_no = reader["proforma_no"].ToString();
                        model.plant = reader["plant"].ToString();
                        model.tax = reader["tax"].ToString();
                        model.shipment_date = reader["shipment_date"] != DBNull.Value ? Convert.ToDateTime(reader["shipment_date"]) : null;
                        model.shipment_type = reader["shipment_type"].ToString();
                        model.dhl_awb = reader["dhl_awb"].ToString();
                        model.vendor_name = reader["vendor_name"].ToString();
                        model.street = reader["street"].ToString();
                        model.city = reader["city"].ToString();
                        model.country = reader["country"].ToString();
                        model.postal_code = reader["postal_code"].ToString();
                        model.phone_no = reader["phone_no"].ToString();
                        model.attn_to = reader["attn_to"].ToString();
                        model.coo = reader["coo"].ToString();
                        model.ship_mode = reader["ship_mode"].ToString();
                        model.courier_account_no = reader["courier_account_no"].ToString();
                        model.courier_charges = reader["courier_charges"].ToString();
                        model.freight_charges = reader["freight_charges"].ToString();
                        model.incoterms = reader["incoterms"].ToString();
                        model.invoice_payment = reader["invoice_payment"].ToString();
                        model.requestor_name = reader["requestor_name"].ToString();
                        model.requestor_plant = reader["requestor_plant"].ToString();
                        model.total_boxes = Convert.ToInt32(reader["total_boxes"]);
                        model.total_assets = Convert.ToInt32(reader["total_assets"]);
                        model.remark = reader["remark"].ToString();
                    }
                }

                model.boxes = new List<ShippingBoxExportModel>();
                using (SqlCommand cmd = new SqlCommand("GET_SHIPPING_EXPORT_BOXES", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var box = new ShippingBoxExportModel();
                        box.box_no = reader["box_no"].ToString();
                        box.length_cm = Convert.ToDecimal(reader["length_cm"]);
                        box.width_cm = Convert.ToDecimal(reader["width_cm"]);
                        box.height_cm = Convert.ToDecimal(reader["height_cm"]);
                        box.gross_weight_kg = Convert.ToDecimal(reader["gross_weight_kg"]);
                        box.net_weight_kg = Convert.ToDecimal(reader["net_weight_kg"]);

                        var boxId = Convert.ToInt32(reader["id_box"]);
                        box.assets = GetAssetsForBox(boxId, conn);

                        model.boxes.Add(box);
                    }
                }

                conn.Close();
            }

            return model;
        }

        private List<ShippingAssetExportModel> GetAssetsForBox(int boxId, SqlConnection conn)
        {
            var assets = new List<ShippingAssetExportModel>();

            using (SqlCommand cmd = new SqlCommand("GET_ASSETS_FOR_BOX", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@id_box", boxId);
                using SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var asset = new ShippingAssetExportModel();
                    asset.asset_no = reader["asset_no"].ToString();
                    asset.asset_subnumber = reader["asset_subnumber"].ToString();
                    asset.asset_desc = reader["asset_desc"].ToString();
                    asset.current_apc = Convert.ToDecimal(reader["current_apc"]);
                    asset.hs_code = reader["hs_code"] != DBNull.Value ? Convert.ToDecimal(reader["hs_code"]) : null;
                    assets.Add(asset);
                }
            }

            return assets;
        }
        public string UpdateShippingAssetHSCode(int id_detail, string asset_no, string asset_subnumber, string hs_code)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_SHIPPING_ASSET_HS_CODE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_detail", id_detail);
                    cmd.Parameters.AddWithValue("@asset_no", asset_no);
                    cmd.Parameters.AddWithValue("@asset_subnumber", asset_subnumber);

                    if (string.IsNullOrEmpty(hs_code))
                    {
                        cmd.Parameters.AddWithValue("@hs_code", DBNull.Value);
                    }
                    else
                    {
                        if (decimal.TryParse(hs_code, out decimal hsCodeValue))
                        {
                            cmd.Parameters.AddWithValue("@hs_code", hsCodeValue);
                        }
                        else
                        {
                            return "error;Invalid HS Code format";
                        }
                    }

                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public List<string> GET_UOM()
        {
            List<string> uomList = new List<string>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_UOM", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            uomList.Add(reader["uom"].ToString() ?? "");
                        }
                    }
                }
                conn.Close();
            }
            return uomList;
        }

        public string DELETE_TEMP_NON_ASSET_SESA(string sesa_id)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("DELETE_TEMP_NON_ASSET_SESA", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string ADD_TEMP_NON_ASSET(string sesa_id, string po_or_asset_no, string description, decimal qty, string uom, decimal price_value, string filename_before, decimal? hs_code)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("INSERT_TEMP_NON_ASSET", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.Parameters.AddWithValue("@po_or_asset_no", (object)po_or_asset_no ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@description", (object)description ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@qty", qty);
                    cmd.Parameters.AddWithValue("@uom", (object)uom ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@price_value", price_value);
                    cmd.Parameters.AddWithValue("@filename_before", (object)filename_before ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@hs_code", hs_code.HasValue ? hs_code.Value : (object)DBNull.Value);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public List<NonAssetTempModel> GET_TEMP_NON_ASSET(string sesa_id)
        {
            var list = new List<NonAssetTempModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_TEMP_NON_ASSET", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    using (SqlDataReader r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            list.Add(new NonAssetTempModel
                            {
                                id_temp = Convert.ToInt32(r["id_temp"]),
                                sesa_id = r["sesa_id"].ToString() ?? "",
                                po_or_asset_no = r["po_or_asset_no"].ToString() ?? "",
                                description = r["description"].ToString() ?? "",
                                qty = r["qty"] != DBNull.Value ? Convert.ToDecimal(r["qty"]) : 0,
                                uom = r["uom"].ToString() ?? "",
                                price_value = r["price_value"] != DBNull.Value ? Convert.ToDecimal(r["price_value"]) : 0,
                                hs_code = r["hs_code"] == DBNull.Value ? null : Convert.ToDecimal(r["hs_code"]),
                                gp_before = r["gp_before"].ToString() ?? "",
                                gp_after = r["gp_after"].ToString() ?? "",
                                record_date = r["record_date"] != DBNull.Value ? Convert.ToDateTime(r["record_date"]) : DateTime.Now
                            });
                        }
                    }
                }
                conn.Close();
            }
            return list;
        }

        public string DELETE_TEMP_NON_ASSET(int id_temp)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("DELETE_TEMP_NON_ASSET", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_temp", id_temp);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string SAVE_GATEPASS_NON_ASSET(
            string created_by,
            string category,
            string return_date,
            string vendor_code,
            string vendor_name,
            string vendor_address,
            string vendor_phone,
            string vendor_email,
            string security_guard,
            string shipping_plant,
            string vehicle_no,
            string driver_name,
            string remark,
            int? id_proforma = null)
        {
            string id_gatepass = "";
            string gatepass_no = "";
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("INSERT_NEW_NON_ASSET_GATEPASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@category", string.IsNullOrEmpty(category) ? (object)DBNull.Value : category);
                    cmd.Parameters.AddWithValue("@return_date", string.IsNullOrEmpty(return_date) ? (object)DBNull.Value : return_date);
                    cmd.Parameters.AddWithValue("@vendor_code", string.IsNullOrEmpty(vendor_code) ? (object)DBNull.Value : vendor_code);
                    cmd.Parameters.AddWithValue("@vendor_name", string.IsNullOrEmpty(vendor_name) ? (object)DBNull.Value : vendor_name);
                    cmd.Parameters.AddWithValue("@vendor_address", string.IsNullOrEmpty(vendor_address) ? (object)DBNull.Value : vendor_address);
                    cmd.Parameters.AddWithValue("@vendor_phone", string.IsNullOrEmpty(vendor_phone) ? (object)DBNull.Value : vendor_phone);
                    cmd.Parameters.AddWithValue("@vendor_email", string.IsNullOrEmpty(vendor_email) ? (object)DBNull.Value : vendor_email);
                    cmd.Parameters.AddWithValue("@vehicle_no", string.IsNullOrEmpty(vehicle_no) ? (object)DBNull.Value : vehicle_no);
                    cmd.Parameters.AddWithValue("@driver_name", string.IsNullOrEmpty(driver_name) ? (object)DBNull.Value : driver_name);
                    cmd.Parameters.AddWithValue("@security_guard", string.IsNullOrEmpty(security_guard) ? (object)DBNull.Value : security_guard);
                    cmd.Parameters.AddWithValue("@shipping_plant", string.IsNullOrEmpty(shipping_plant) ? (object)DBNull.Value : shipping_plant);
                    cmd.Parameters.AddWithValue("@remark", string.IsNullOrEmpty(remark) ? (object)DBNull.Value : remark);
                    cmd.Parameters.AddWithValue("@created_by", string.IsNullOrEmpty(created_by) ? (object)DBNull.Value : created_by);
                    cmd.Parameters.AddWithValue("@id_proforma", (object)id_proforma ?? DBNull.Value);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            id_gatepass = reader["id_gatepass"].ToString() ?? "";
                            gatepass_no = reader["gatepass_no"].ToString() ?? "";
                        }
                    }
                }
                conn.Close();
            }
            return id_gatepass + ";" + gatepass_no;
        }

        public string SAVE_PROFORMA_NON_ASSET(
            string attn_to, string street, string city, string country,
            string postal_code, string phone_no, string email, string coo, string file_attach, string file_support,
            string ship_mode, string courier_charges, string courier_name, string courier_account_no,
            string freight_charges, string incoterms, string invoice_payment, string requested_by)
        {
            string id_proforma = "";
            string proforma_no = "";

            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("INSERT_NEW_NON_ASSET_PROFORMA_HEADER", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@attn_to", string.IsNullOrEmpty(attn_to) ? (object)DBNull.Value : attn_to);
                    cmd.Parameters.AddWithValue("@street", string.IsNullOrEmpty(street) ? (object)DBNull.Value : street);
                    cmd.Parameters.AddWithValue("@city", string.IsNullOrEmpty(city) ? (object)DBNull.Value : city);
                    cmd.Parameters.AddWithValue("@country", string.IsNullOrEmpty(country) ? (object)DBNull.Value : country);
                    cmd.Parameters.AddWithValue("@postal_code", string.IsNullOrEmpty(postal_code) ? (object)DBNull.Value : postal_code);
                    cmd.Parameters.AddWithValue("@phone_no", string.IsNullOrEmpty(phone_no) ? (object)DBNull.Value : phone_no);
                    cmd.Parameters.AddWithValue("@email", string.IsNullOrEmpty(email) ? (object)DBNull.Value : email);
                    cmd.Parameters.AddWithValue("@coo", string.IsNullOrEmpty(coo) ? (object)DBNull.Value : coo);
                    cmd.Parameters.AddWithValue("@file_attach", string.IsNullOrEmpty(file_attach) ? (object)DBNull.Value : file_attach);
                    cmd.Parameters.AddWithValue("@file_support", string.IsNullOrEmpty(file_support) ? (object)DBNull.Value : file_support);
                    cmd.Parameters.AddWithValue("@ship_mode", string.IsNullOrEmpty(ship_mode) ? (object)DBNull.Value : ship_mode);
                    cmd.Parameters.AddWithValue("@courier_charges", string.IsNullOrEmpty(courier_charges) ? (object)DBNull.Value : courier_charges);
                    cmd.Parameters.AddWithValue("@courier_name", string.IsNullOrEmpty(courier_name) ? (object)DBNull.Value : courier_name);
                    cmd.Parameters.AddWithValue("@courier_account_no", string.IsNullOrEmpty(courier_account_no) ? (object)DBNull.Value : courier_account_no);
                    cmd.Parameters.AddWithValue("@freight_charges", string.IsNullOrEmpty(freight_charges) ? (object)DBNull.Value : freight_charges);
                    cmd.Parameters.AddWithValue("@incoterms", string.IsNullOrEmpty(incoterms) ? (object)DBNull.Value : incoterms);
                    cmd.Parameters.AddWithValue("@invoice_payment", string.IsNullOrEmpty(invoice_payment) ? (object)DBNull.Value : invoice_payment);
                    cmd.Parameters.AddWithValue("@requested_by", string.IsNullOrEmpty(requested_by) ? (object)DBNull.Value : requested_by);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            id_proforma = reader["id_proforma"].ToString() ?? "";
                            proforma_no = reader["proforma_no"].ToString() ?? "";
                        }
                    }
                }
                conn.Close();
            }
            return id_proforma + ";" + proforma_no;
        }

        public List<Dictionary<string, object>> GET_IMG_NON_ASSET_GATEPASS(int id_gatepass)
        {
            var list = new List<Dictionary<string, object>>();
            using (var conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (var cmd = new SqlCommand("GET_IMG_NON_ASSET_GATEPASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using (var rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            var row = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                            {
                                ["gatepass_no"] = rdr["gatepass_no"]?.ToString(),
                                ["po_or_asset_no"] = rdr["po_or_asset_no"]?.ToString(),
                                ["description"] = rdr["description"]?.ToString(),
                                ["qty"] = rdr["qty"] == DBNull.Value ? 0 : Convert.ToDecimal(rdr["qty"]),
                                ["uom"] = rdr["uom"]?.ToString(),
                                ["gp_before"] = rdr["gp_before"] != DBNull.Value ? rdr["gp_before"].ToString() : "NotAvail.png",
                                ["gp_after"] = rdr["gp_after"] != DBNull.Value ? rdr["gp_after"].ToString() : "NotAvail.png"
                            };
                            list.Add(row);
                        }
                    }
                }
                conn.Close();
            }
            return list;
        }

        public string EDIT_GATEPASS_NON_ASSET(int id_gatepass, string return_date)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_GATEPASS_NON_ASSET", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@return_date", return_date);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string CANCEL_GATEPASS_NON_ASSET(int id_gatepass, string cancelled_by, string cancel_reason)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("CANCEL_GATEPASS_NON_ASSET", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        cmd.Parameters.AddWithValue("@cancelled_by", cancelled_by);
                        cmd.Parameters.AddWithValue("@cancel_reason", string.IsNullOrEmpty(cancel_reason) ? (object)DBNull.Value : cancel_reason);

                        conn.Open();
                        var result = cmd.ExecuteScalar()?.ToString() ?? "error;Unknown error occurred";
                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                return $"error;{ex.Message}";
            }
        }

        public List<GatePassNonAssetModel> GET_LIST_NON_ASSET_IMG(string id_gatepass)
        {
            List<GatePassNonAssetModel> materials = new List<GatePassNonAssetModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_LIST_NON_ASSET_IMG", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass ?? (object)DBNull.Value);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            GatePassNonAssetModel mat = new GatePassNonAssetModel
                            {
                                id_gatepass = reader["id_gatepass"] != DBNull.Value ? Convert.ToInt32(reader["id_gatepass"]) : 0,
                                id_detail = reader["id_detail"].ToString(),
                                po_or_asset_no = reader["po_or_asset_no"].ToString(),
                                image_after = reader["gp_after"].ToString(),
                                return_date = reader["return_date"].ToString(),
                            };
                            materials.Add(mat);
                        }
                    }
                }
                conn.Close();
            }
            return materials;
        }

        public string UPDATE_AFTER_NON_ASSET_GATEPASS(int id_gatepass, string ret_file, string id_detail)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_AFTER_NON_ASSET_GATEPASS_RETURN", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_detail", id_detail);
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@ret_file", ret_file);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public Dictionary<string, object> GET_SHOW_NON_ASSET_DETAIL(int id_gatepass)
        {
            var dict = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            using (var conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (var cmd = new SqlCommand("GET_SHOW_NON_ASSET_DETAIL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using (var rdr = cmd.ExecuteReader())
                    {
                        if (rdr.Read())
                        {
                            dict["vendor_name"] = rdr["vendor_name"]?.ToString();
                            dict["vendor_address"] = rdr["vendor_address"]?.ToString();
                            dict["recipient_phone"] = rdr["recipient_phone"]?.ToString();
                            dict["recipient_email"] = rdr["recipient_email"]?.ToString();
                            dict["security_guard"] = rdr["security_guard"]?.ToString();
                            dict["vehicle_no"] = rdr["vehicle_no"]?.ToString();
                            dict["driver_name"] = rdr["driver_name"]?.ToString();
                        }
                    }
                }
                conn.Close();
            }
            return dict;
        }

        public List<GatePassNonAssetModel> GET_GATEPASS_NON_ASSET_DETAIL(int id_gatepass)
        {
            List<GatePassNonAssetModel> detList = new List<GatePassNonAssetModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GATEPASS_NON_ASSET_DETAIL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);

                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            GatePassNonAssetModel row = new GatePassNonAssetModel();
                            row.id_gatepass = Convert.ToInt32(reader["id_gatepass"]);
                            row.gatepass_no = reader["gatepass_no"].ToString();
                            row.category = reader["category"].ToString();
                            row.po_or_asset_no = reader["po_or_asset_no"].ToString();
                            row.description = reader["description"].ToString();
                            row.qty = reader["qty"] != DBNull.Value ? Convert.ToDecimal(reader["qty"]) : 0;
                            row.uom = reader["uom"].ToString();
                            row.price_value = reader["price_value"] != DBNull.Value ? Convert.ToDecimal(reader["price_value"]) : 0;
                            row.hs_code = reader["hs_code"] != DBNull.Value ? Convert.ToDecimal(reader["hs_code"]) : null;
                            row.box_no = reader["box_no"]?.ToString();
                            row.id_proforma = reader["id_proforma"] != DBNull.Value ? Convert.ToInt32(reader["id_proforma"]) : null;
                            row.proforma_fin_status = reader["proforma_fin_status"]?.ToString();
                            row.shipping_status = reader["shipping_status"]?.ToString();
                            row.proforma_no = reader["proforma_no"]?.ToString();
                            row.attn_to = reader["attn_to"]?.ToString();
                            row.street = reader["street"]?.ToString();
                            row.city = reader["city"]?.ToString();
                            row.country = reader["country"]?.ToString();
                            row.postal_code = reader["postal_code"]?.ToString();
                            row.phone_no = reader["phone_no"]?.ToString();
                            row.email = reader["email"]?.ToString();
                            row.coo = reader["coo"]?.ToString();
                            row.ship_mode = reader["ship_mode"]?.ToString();
                            row.courier_charges = reader["courier_charges"]?.ToString();
                            row.courier_name = reader["courier_name"]?.ToString();
                            row.courier_account_no = reader["courier_account_no"]?.ToString();
                            row.freight_charges = reader["freight_charges"]?.ToString();
                            row.incoterms = reader["incoterms"]?.ToString();
                            row.invoice_payment = reader["invoice_payment"]?.ToString();
                            row.file_attach = reader["file_attach"]?.ToString();
                            row.file_support = reader["file_support"]?.ToString();
                            row.file_peb = reader["file_peb"]?.ToString();
                            row.plant = reader["plant"]?.ToString();
                            if (reader["shipment_date"] != DBNull.Value)
                            {
                                row.shipment_date = Convert.ToDateTime(reader["shipment_date"]).ToString("dd MMM yyyy");
                            }
                            row.dhl_awb = reader["dhl_awb"]?.ToString();
                            row.id_file = reader["id_file"] != DBNull.Value ? Convert.ToInt32(reader["id_file"]) : null;
                            row.document_type = reader["document_type"]?.ToString();
                            row.fin_filename = reader["fin_filename"]?.ToString();
                            row.fin_created_by = reader["fin_created_by"]?.ToString();
                            row.fin_record_date = reader["fin_record_date"] != DBNull.Value ? Convert.ToDateTime(reader["fin_record_date"]) : null;

                            detList.Add(row);
                        }
                    }
                }

                conn.Close();
            }
            return detList;
        }

        public List<GatePassNonAssetModel> GetGatepassNonAssetInfo(string id_gatepass)
        {
            List<GatePassNonAssetModel> gatepassList = new List<GatePassNonAssetModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GATEPASS_NON_ASSET_INFO", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                GatePassNonAssetModel row = new GatePassNonAssetModel();
                                row.id_gatepass = Convert.ToInt32(reader["id_gatepass"]);
                                row.gatepass_no = reader["gatepass_no"].ToString();
                                row.category = reader["category"].ToString();
                                row.status_gatepass = reader["status_gatepass"].ToString();
                                row.return_date = reader["return_date"].ToString();
                                row.vendor_code = reader["vendor_code"].ToString();
                                row.vendor_name = reader["vendor_name"].ToString();
                                row.vendor_address = reader["vendor_address"].ToString();
                                row.recipient_phone = reader["recipient_phone"].ToString();
                                row.recipient_email = reader["recipient_email"].ToString();
                                row.vehicle_no = reader["vehicle_no"].ToString();
                                row.driver_name = reader["driver_name"].ToString();
                                row.security_guard = reader["security_guard"].ToString();
                                row.remark = reader["remark"].ToString();
                                row.created_by = reader["created_by"].ToString();
                                row.create_date = reader["create_date"].ToString();
                                row.image_before = reader["image_before"].ToString();
                                row.image_after = reader["image_after"].ToString();
                                gatepassList.Add(row);
                            }
                        }
                    }
                }
                conn.Close();
            }
            return gatepassList;
        }

        public List<ApprovalModel> GetGatepassNonAssetApproval(string id_gatepass)
        {
            List<ApprovalModel> approvalList = new List<ApprovalModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GATEPASS_NON_ASSET_APPROVAL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            ApprovalModel row = new ApprovalModel();
                            row.id_approval = reader["id_approval"].ToString();
                            row.id_gatepass = reader["id_gatepass"].ToString();
                            row.approval_no = reader["approval_no"].ToString();
                            row.approval_level = reader["approval_level"].ToString();
                            row.approval_by = reader["approval_by"].ToString();
                            row.approval_status = reader["approval_status"].ToString();
                            row.date_approval = reader["date_approval"].ToString();
                            row.approval_name = reader["approval_name"].ToString();
                            approvalList.Add(row);
                        }
                    }
                }
                conn.Close();
            }
            return approvalList;
        }

        public List<GatePassNonAssetModel> GET_GATEPASS_NON_ASSET_STATUS(int id_gatepass)
        {
            List<GatePassNonAssetModel> dataList = new List<GatePassNonAssetModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GATEPASS_NON_ASSET_STATUS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            GatePassNonAssetModel row = new GatePassNonAssetModel();
                            row.approver_no = reader["approver_no"].ToString();
                            row.approver = reader["approver"].ToString();
                            row.approval_status = reader["approval_status"].ToString();
                            if (reader["date_approval"] != DBNull.Value)
                            {
                                row.date_approval = DateTime.Parse(reader["date_approval"].ToString()).ToString("dd MMM yyyy");
                            }
                            else
                            {
                                row.date_approval = null;
                            }

                            string baseRemark = reader["remark"].ToString();
                            string delegatedBy = reader["delegated_by"].ToString();
                            string delegatedByName = reader["delegated_by_name"].ToString();

                            if (!string.IsNullOrEmpty(delegatedBy) && !string.IsNullOrEmpty(delegatedByName))
                            {
                                if (!string.IsNullOrEmpty(baseRemark))
                                {
                                    row.remark = baseRemark + "\n\nApproved by: " + delegatedByName;
                                }
                                else
                                {
                                    row.remark = "Approved by: " + delegatedByName;
                                }
                            }
                            else
                            {
                                row.remark = baseRemark;
                            }

                            dataList.Add(row);
                        }
                    }
                }

                conn.Close();
            }
            return dataList;
        }

        public List<GatePassNonAssetModel> GET_GATEPASS_NON_ASSET_HEADER(int id_gatepass)
        {
            List<GatePassNonAssetModel> disList = new List<GatePassNonAssetModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GATEPASS_NON_ASSET_HEADER", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            GatePassNonAssetModel row = new GatePassNonAssetModel();
                            row.gatepass_no = reader["gatepass_no"].ToString();
                            row.created_by = reader["created_by"].ToString();
                            row.category = reader["category"].ToString();
                            if (reader["return_date"] != DBNull.Value)
                            {
                                row.return_date = DateTime.Parse(reader["return_date"].ToString()).ToString("dd MMM yyyy");
                            }
                            else
                            {
                                row.return_date = "-";
                            }
                            row.create_date = DateTime.Parse(reader["create_date"].ToString()).ToString("dd MMM yyyy");
                            row.remark = reader["remark"].ToString();
                            disList.Add(row);
                        }
                    }
                }
                conn.Close();
            }
            return disList;
        }

        public decimal GET_TOTAL_NON_ASSET_VALUE(int id_gatepass)
        {
            decimal totalValue = 0;
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_TOTAL_NON_ASSET_VALUE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    var result = cmd.ExecuteScalar();
                    if (result != DBNull.Value)
                    {
                        totalValue = Convert.ToDecimal(result);
                    }
                }
                conn.Close();
            }
            return totalValue;
        }

        public string UPDATE_APPROVAL_NON_ASSET_GATEPASS(int id_gatepass, string approval_level, string approval_status, string approval_by, string remark)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();

                string actualApprover = GetGatepassNonAssetApproverAtLevel(id_gatepass, approval_level);
                bool isDelegated = IsDelegatedApprover(approval_by, actualApprover);
                string delegatedBy = null;

                if (isDelegated)
                {
                    delegatedBy = approval_by;
                }
                else
                {
                    actualApprover = approval_by;
                    delegatedBy = null;
                }

                using (SqlCommand cmd = new SqlCommand("UPDATE_APPROVAL_NON_ASSET_GATEPASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@approval_level", approval_level);
                    cmd.Parameters.AddWithValue("@approval_status", approval_status);
                    cmd.Parameters.AddWithValue("@approval_by", actualApprover);
                    cmd.Parameters.AddWithValue("@remark", remark);
                    cmd.Parameters.AddWithValue("@delegated_by", delegatedBy ?? (object)DBNull.Value);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string CONFIRM_GATEPASS_NON_ASSET_RETURN(int id_gatepass, string approval_level, string approval_status, string approval_by, string remark)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();

                string actualConfirmer = GetGatepassNonAssetConfirmerHOD(id_gatepass);
                bool isDelegated = IsDelegatedApprover(approval_by, actualConfirmer);
                string delegatedBy = null;

                if (isDelegated)
                {
                    delegatedBy = approval_by;
                }
                else
                {
                    actualConfirmer = approval_by;
                    delegatedBy = null;
                }

                using (SqlCommand cmd = new SqlCommand("CONFIRM_GATEPASS_NON_ASSET_RETURN", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@approval_level", approval_level);
                    cmd.Parameters.AddWithValue("@approval_status", approval_status);
                    cmd.Parameters.AddWithValue("@approval_by", actualConfirmer);
                    cmd.Parameters.AddWithValue("@remark", remark);
                    cmd.Parameters.AddWithValue("@delegated_by", delegatedBy ?? (object)DBNull.Value);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string GetGatepassNonAssetApproverAtLevel(int id_gatepass, string approval_level)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GATEPASS_NON_ASSET_APPROVER_LEVEL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@approval_level", approval_level);
                    var result = cmd.ExecuteScalar();
                    return result?.ToString() ?? "";
                }
            }
        }

        public string GetGatepassNonAssetConfirmerHOD(int id_gatepass)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GATEPASS_NON_ASSET_CONFIRMER_HOD", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    var result = cmd.ExecuteScalar();
                    return result?.ToString() ?? "";
                }
            }
        }

        public string UPDATE_FINANCE_NON_ASSET_GATEPASS(int id_gatepass, string approval_status, string approval_remark, string sesa_id)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_FINANCE_NON_ASSET_GATEPASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@approval_status", approval_status);
                    cmd.Parameters.AddWithValue("@approval_remark", approval_remark);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string UPLOAD_PROFORMA_NON_ASSET_FIN_FILES(int id_gatepass, string document_type, string filename, string created_by)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("INSERT_PROFORMA_NON_ASSET_FIN_FILES", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@document_type", document_type);
                    cmd.Parameters.AddWithValue("@filename", filename);
                    cmd.Parameters.AddWithValue("@created_by", created_by);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return reader["result"].ToString();
                        }
                    }
                    return "error;Unknown error";
                }
            }
        }

        public string DELETE_PROFORMA_NON_ASSET_FIN_FILE(int id_file, int id_gatepass)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("DELETE_PROFORMA_NON_ASSET_FIN_FILE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_file", id_file);
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string result = reader["result"].ToString();
                            string filename = reader["filename"].ToString();
                            if (result == "success" && !string.IsNullOrEmpty(filename))
                            {
                                string filePath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Upload", "Proforma_Finance", filename);
                                if (System.IO.File.Exists(filePath))
                                {
                                    System.IO.File.Delete(filePath);
                                }
                            }
                            return result;
                        }
                    }
                    return "Error: No result returned";
                }
            }
        }

        public string UPDATE_PROFORMA_NON_ASSET_FIN_STATUS(int id_gatepass, string completed_by_sesa_id = null)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(ConnectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("UPDATE_PROFORMA_NON_ASSET_FIN_STATUS", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        cmd.Parameters.AddWithValue("@completed_by_sesa_id", completed_by_sesa_id ?? (object)DBNull.Value);

                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            do
                            {
                                if (reader.HasRows)
                                {
                                    while (reader.Read())
                                    {
                                        for (int i = 0; i < reader.FieldCount; i++)
                                        {
                                            if (reader.GetName(i).Equals("result", StringComparison.OrdinalIgnoreCase))
                                            {
                                                return reader["result"].ToString();
                                            }
                                        }
                                    }
                                }
                            } while (reader.NextResult());
                        }
                    }
                }
                return "success";
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
        }

        public List<ProformaFileModel> GET_PROFORMA_NON_ASSET_FIN_FILES(int id_gatepass)
        {
            List<ProformaFileModel> fileList = new List<ProformaFileModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_PROFORMA_NON_ASSET_FIN_FILES", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            ProformaFileModel file = new ProformaFileModel();
                            file.id_file = Convert.ToInt32(reader["id_file"]);
                            file.document_type = reader["document_type"].ToString();
                            file.filename = reader["filename"].ToString();
                            file.record_date = Convert.ToDateTime(reader["record_date"]);
                            fileList.Add(file);
                        }
                    }
                }
                conn.Close();
            }
            return fileList;
        }

        public PebFileModel GET_PEB_NON_ASSET_FILE(int id_gatepass)
        {
            PebFileModel file = null;
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_PEB_NON_ASSET_FILE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        file = new PebFileModel();
                        string filenames = reader["file_peb"].ToString();
                        file.filename = filenames;
                        file.filenames = string.IsNullOrEmpty(filenames) ? new List<string>() : filenames.Split(';').ToList();
                    }
                }
                conn.Close();
            }
            return file;
        }

        public string UPLOAD_PEB_NON_ASSET_FILE(int id_gatepass, string newFilenames, string created_by)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("INSERT_PROFORMA_NON_ASSET_PEB_FILE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@newFilenames", newFilenames);
                    cmd.Parameters.AddWithValue("@created_by", created_by ?? (object)DBNull.Value);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string result = reader["result"].ToString();
                            if (result == "success")
                            {
                                return "success";
                            }
                            else
                            {
                                string errorMessage = reader["error_message"]?.ToString() ?? "Unknown error";
                                return $"error;{errorMessage}";
                            }
                        }
                    }
                    return "error;No result returned";
                }
            }
        }

        public string DELETE_PEB_NON_ASSET_FILE(int id_gatepass, string filenameToDelete = null)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("DELETE_PEB_NON_ASSET_FILE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@filenameToDelete", string.IsNullOrEmpty(filenameToDelete) ? (object)DBNull.Value : filenameToDelete);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string result = reader["result"].ToString();
                            string filesToDeleteStr = reader["files_to_delete"].ToString();
                            if (result == "success" && !string.IsNullOrEmpty(filesToDeleteStr))
                            {
                                List<string> filesToDelete = filesToDeleteStr.Split(';').ToList();
                                foreach (string filename in filesToDelete)
                                {
                                    if (!string.IsNullOrEmpty(filename))
                                    {
                                        string filePath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Upload", "PEB", filename);
                                        if (System.IO.File.Exists(filePath))
                                        {
                                            System.IO.File.Delete(filePath);
                                        }
                                    }
                                }
                            }
                            return result;
                        }
                    }
                    return "Error: No result returned";
                }
            }
        }

        public string COMPLETE_PEB_NON_ASSET_UPLOAD(int id_gatepass, string completed_by_sesa_id = null)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("UPDATE_PROFORMA_NON_ASSET_PEB_STATUS", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        cmd.Parameters.AddWithValue("@completed_by_sesa_id", completed_by_sesa_id ?? (object)DBNull.Value);

                        conn.Open();
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            do
                            {
                                if (reader.HasRows)
                                {
                                    while (reader.Read())
                                    {
                                        for (int i = 0; i < reader.FieldCount; i++)
                                        {
                                            if (reader.GetName(i).Equals("result", StringComparison.OrdinalIgnoreCase))
                                            {
                                                return reader["result"].ToString();
                                            }
                                        }
                                    }
                                }
                            } while (reader.NextResult());
                        }
                    }
                }
                return "success";
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
        }

        public List<ShippingNonAssetModel> GET_SHIPPING_NON_ASSETS(int id_gatepass, string sesa_id = "")
        {
            List<ShippingNonAssetModel> itemList = new List<ShippingNonAssetModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_SHIPPING_NON_ASSETS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id ?? (object)DBNull.Value);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        ShippingNonAssetModel row = new ShippingNonAssetModel();
                        row.id_detail = Convert.ToInt32(reader["id_detail"]);
                        row.po_or_asset_no = reader["po_or_asset_no"].ToString();
                        row.description = reader["description"].ToString();
                        row.qty = reader["qty"] != DBNull.Value ? Convert.ToDecimal(reader["qty"]) : 0;
                        row.uom = reader["uom"].ToString();
                        row.price_value = reader["price_value"] != DBNull.Value ? Convert.ToDecimal(reader["price_value"]) : 0;
                        row.hs_code = reader["hs_code"] == DBNull.Value ? null : Convert.ToDecimal(reader["hs_code"]);
                        row.id_box = Convert.ToInt32(reader["id_box"]);
                        row.is_assigned = Convert.ToBoolean(reader["is_assigned"]);
                        row.is_selected = false;
                        itemList.Add(row);
                    }
                }
                conn.Close();
            }
            return itemList;
        }

        public string CREATE_TEMP_NON_ASSET_BOX(TempShippingNonAssetBoxModel box)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("INSERT_SHIPPING_NON_ASSET_TEMP_BOX", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@created_by", box.created_by ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@id_gatepass", box.id_gatepass ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@box_no", box.box_no ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@item_list", box.item_list ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@length_cm", box.length_cm ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@width_cm", box.width_cm ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@height_cm", box.height_cm ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@gross_weight_kg", box.gross_weight_kg ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@net_weight_kg", box.net_weight_kg ?? (object)DBNull.Value);

                    SqlParameter newIdParam = new SqlParameter("@new_id", SqlDbType.Int);
                    newIdParam.Direction = ParameterDirection.Output;
                    cmd.Parameters.Add(newIdParam);

                    SqlParameter resultParam = new SqlParameter("@result", SqlDbType.NVarChar, 500);
                    resultParam.Direction = ParameterDirection.Output;
                    cmd.Parameters.Add(resultParam);

                    cmd.ExecuteNonQuery();

                    string result = resultParam.Value?.ToString() ?? "success";

                    if (result == "success")
                    {
                        int newId = Convert.ToInt32(newIdParam.Value);
                        return $"success;{newId}";
                    }

                    return result;
                }
            }
        }

        public List<TempShippingNonAssetBoxModel> GET_TEMP_NON_ASSET_BOXES(string sesa_id, int id_gatepass)
        {
            List<TempShippingNonAssetBoxModel> boxList = new List<TempShippingNonAssetBoxModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_SHIPPING_NON_ASSET_TEMP_BOXES", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@created_by", sesa_id);
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);

                    using SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        TempShippingNonAssetBoxModel row = new TempShippingNonAssetBoxModel();
                        row.id_temp = Convert.ToInt32(reader["id_temp"]);
                        row.created_by = reader["created_by"].ToString();
                        row.id_gatepass = reader["id_gatepass"] != DBNull.Value ? Convert.ToInt32(reader["id_gatepass"]) : 0;
                        row.box_no = reader["box_no"].ToString();
                        row.item_list = reader["item_list"].ToString();
                        row.length_cm = reader["length_cm"] != DBNull.Value ? Convert.ToDecimal(reader["length_cm"]) : 0;
                        row.width_cm = reader["width_cm"] != DBNull.Value ? Convert.ToDecimal(reader["width_cm"]) : 0;
                        row.height_cm = reader["height_cm"] != DBNull.Value ? Convert.ToDecimal(reader["height_cm"]) : 0;
                        row.gross_weight_kg = reader["gross_weight_kg"] != DBNull.Value ? Convert.ToDecimal(reader["gross_weight_kg"]) : 0;
                        row.net_weight_kg = reader["net_weight_kg"] != DBNull.Value ? Convert.ToDecimal(reader["net_weight_kg"]) : 0;
                        row.create_date = reader["record_date"] != DBNull.Value ? Convert.ToDateTime(reader["record_date"]) : DateTime.Now;

                        boxList.Add(row);
                    }
                }
                conn.Close();
            }
            return boxList;
        }

        public string DELETE_TEMP_NON_ASSET_BOX(int id_temp)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("DELETE_SHIPPING_NON_ASSET_TEMP_BOX", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_temp", id_temp);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string SAVE_SHIPPING_NON_ASSET(ShippingNonAssetCreateViewModel model, string sesa_id)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("INSERT_SHIPPING_NON_ASSET", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", model.id_gatepass);
                    cmd.Parameters.AddWithValue("@plant", model.plant);
                    cmd.Parameters.AddWithValue("@shipment_date", model.shipment_date);
                    string shipmentTypeToSave = !string.IsNullOrEmpty(model.shipment_type) ? model.shipment_type : model.shipment_type;
                    cmd.Parameters.AddWithValue("@shipment_type", shipmentTypeToSave ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@dhl_awb", model.dhl_awb ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    var tempBoxes = GET_TEMP_NON_ASSET_BOXES(sesa_id, model.id_gatepass);
                    string tempBoxesJson = System.Text.Json.JsonSerializer.Serialize(tempBoxes);
                    cmd.Parameters.AddWithValue("@temp_boxes_json", tempBoxesJson);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string result = reader["result"].ToString();
                            string shippingId = reader["shipping_id"].ToString();
                            string proformaNo = reader["proforma_no"].ToString();
                            return $"{result};{shippingId};{proformaNo}";
                        }
                    }
                    return "error;Unknown error";
                }
            }
        }

        public string SAVE_SHIPPING_NON_ASSET_FORCE(ShippingNonAssetCreateViewModel model, string sesa_id, bool forceDelete = false)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("INSERT_SHIPPING_NON_ASSET_FORCE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", model.id_gatepass);
                    cmd.Parameters.AddWithValue("@plant", model.plant);
                    cmd.Parameters.AddWithValue("@shipment_date", model.shipment_date);
                    cmd.Parameters.AddWithValue("@shipment_type", model.shipment_type ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@dhl_awb", model.dhl_awb ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    cmd.Parameters.AddWithValue("@force_delete", forceDelete);

                    var tempBoxes = GET_TEMP_NON_ASSET_BOXES(sesa_id, model.id_gatepass);
                    string tempBoxesJson = System.Text.Json.JsonSerializer.Serialize(tempBoxes);
                    cmd.Parameters.AddWithValue("@temp_boxes_json", tempBoxesJson);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string result = reader["result"].ToString();
                            string shippingId = reader["shipping_id"].ToString();
                            string proformaNo = reader["proforma_no"].ToString();
                            return $"{result};{shippingId};{proformaNo}";
                        }
                    }
                    return "error;Unknown error";
                }
            }
        }

        public string SUBMIT_SHIPPING_NON_ASSET(int id_gatepass, string sesa_id)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("UPDATE_SHIPPING_NON_ASSET", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        cmd.Parameters.AddWithValue("@sesa_id", sesa_id);

                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows && reader.Read())
                            {
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    if (reader.GetName(i).Equals("result", StringComparison.OrdinalIgnoreCase))
                                    {
                                        return reader["result"].ToString();
                                    }
                                }

                                return "success";
                            }
                            else
                            {
                                return "error;No result returned from stored procedure";
                            }
                        }
                    }
                }
                catch (SqlException sqlEx)
                {
                    return $"error;SQL Error: {sqlEx.Message}";
                }
                catch (Exception ex)
                {
                    return $"error;{ex.Message}";
                }
            }
        }

        public ShippingNonAssetCreateViewModel GET_SHIPPING_NON_ASSET_DATA(int id_gatepass, string sesa_id = "")
        {
            var model = new ShippingNonAssetCreateViewModel();
            model.id_gatepass = id_gatepass;

            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_SHIPPING_NON_ASSET_DATA", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        model.gatepass_no = reader["gatepass_no"].ToString();
                        model.shipping_plant = reader["shipping_plant"].ToString();
                        model.plant = reader["plant"].ToString();
                        model.dhl_awb = reader["dhl_awb"].ToString();
                        if (reader["shipment_date"] != DBNull.Value)
                            model.shipment_date = Convert.ToDateTime(reader["shipment_date"]);

                        model.is_shipping_saved = reader["id_shipping"] != DBNull.Value;
                        string courierName = reader["courier_name"]?.ToString();
                        string existingShipmentType = reader["shipment_type"]?.ToString();

                        if (!string.IsNullOrEmpty(existingShipmentType))
                        {
                            model.shipment_type = existingShipmentType;
                        }
                        else if (!string.IsNullOrEmpty(courierName))
                        {
                            model.shipment_type = courierName;
                        }
                        else
                        {
                            model.shipment_type = "";
                        }

                        if (model.is_shipping_saved)
                        {
                            string dbPlant = reader["plant"]?.ToString();
                            DateTime? dbShipmentDate = reader["shipment_date"] != DBNull.Value ? Convert.ToDateTime(reader["shipment_date"]) : null;
                            string dbShipmentType = model.shipment_type;
                            string dbDhlAwb = reader["dhl_awb"]?.ToString();

                            model.is_data_complete_in_db = !string.IsNullOrEmpty(dbPlant) &&
                                                          dbShipmentDate.HasValue &&
                                                          !string.IsNullOrEmpty(dbShipmentType) &&
                                                          !string.IsNullOrEmpty(dbDhlAwb);
                        }
                        else
                        {
                            model.is_data_complete_in_db = false;
                        }
                    }
                }

                model.items = GET_SHIPPING_NON_ASSETS(id_gatepass, sesa_id);

                if (model.is_shipping_saved)
                {
                    model.saved_boxes = GET_SAVED_NON_ASSET_SHIPPING_BOXES(id_gatepass);
                }
                else
                {
                    model.temp_boxes = GET_TEMP_NON_ASSET_BOXES(sesa_id, id_gatepass);
                }

                conn.Close();
            }

            return model;
        }

        public List<ShippingNonAssetBoxModel> GET_SAVED_NON_ASSET_SHIPPING_BOXES(int id_gatepass)
        {
            List<ShippingNonAssetBoxModel> boxList = new List<ShippingNonAssetBoxModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_SAVED_NON_ASSET_SHIPPING_BOXES", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        ShippingNonAssetBoxModel row = new ShippingNonAssetBoxModel();
                        row.id_box = Convert.ToInt32(reader["id_box"]);
                        row.box_no = reader["box_no"].ToString();
                        row.length_cm = reader["length_cm"] != DBNull.Value ? Convert.ToDecimal(reader["length_cm"]) : 0;
                        row.width_cm = reader["width_cm"] != DBNull.Value ? Convert.ToDecimal(reader["width_cm"]) : 0;
                        row.height_cm = reader["height_cm"] != DBNull.Value ? Convert.ToDecimal(reader["height_cm"]) : 0;
                        row.gross_weight_kg = reader["gross_weight_kg"] != DBNull.Value ? Convert.ToDecimal(reader["gross_weight_kg"]) : 0;
                        row.net_weight_kg = reader["net_weight_kg"] != DBNull.Value ? Convert.ToDecimal(reader["net_weight_kg"]) : 0;
                        row.item_count = Convert.ToInt32(reader["item_count"]);
                        boxList.Add(row);
                    }
                }
                conn.Close();
            }
            return boxList;
        }

        public string DELETE_SAVED_NON_ASSET_SHIPPING_BOX(int id_box)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("DELETE_SAVED_NON_ASSET_SHIPPING_BOX", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_box", id_box);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return reader["result"].ToString();
                        }
                    }
                    return "error;Unknown error";
                }
            }
        }

        public ShippingNonAssetExportModel GET_SHIPPING_NON_ASSET_EXPORT_DATA(int id_gatepass)
        {
            var model = new ShippingNonAssetExportModel();

            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();

                using (SqlCommand cmd = new SqlCommand("GET_SHIPPING_NON_ASSET_EXPORT_DATA", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        model.proforma_no = reader["proforma_no"].ToString();
                        model.plant = reader["plant"].ToString();
                        model.tax = reader["tax"].ToString();
                        model.shipment_date = reader["shipment_date"] != DBNull.Value ? Convert.ToDateTime(reader["shipment_date"]) : null;
                        model.shipment_type = reader["shipment_type"].ToString();
                        model.dhl_awb = reader["dhl_awb"].ToString();
                        model.vendor_name = reader["vendor_name"].ToString();
                        model.street = reader["street"].ToString();
                        model.city = reader["city"].ToString();
                        model.country = reader["country"].ToString();
                        model.postal_code = reader["postal_code"].ToString();
                        model.phone_no = reader["phone_no"].ToString();
                        model.attn_to = reader["attn_to"].ToString();
                        model.coo = reader["coo"].ToString();
                        model.ship_mode = reader["ship_mode"].ToString();
                        model.courier_account_no = reader["courier_account_no"].ToString();
                        model.courier_charges = reader["courier_charges"].ToString();
                        model.freight_charges = reader["freight_charges"].ToString();
                        model.incoterms = reader["incoterms"].ToString();
                        model.invoice_payment = reader["invoice_payment"].ToString();
                        model.requestor_name = reader["requestor_name"].ToString();
                        model.requestor_plant = reader["requestor_plant"].ToString();
                        model.total_boxes = Convert.ToInt32(reader["total_boxes"]);
                        model.total_items = Convert.ToInt32(reader["total_items"]);
                        model.remark = reader["remark"].ToString();
                    }
                }

                model.boxes = new List<ShippingNonAssetBoxExportModel>();
                using (SqlCommand cmd = new SqlCommand("GET_SHIPPING_NON_ASSET_EXPORT_BOXES", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    using SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var box = new ShippingNonAssetBoxExportModel();
                        box.box_no = reader["box_no"].ToString();
                        box.length_cm = Convert.ToDecimal(reader["length_cm"]);
                        box.width_cm = Convert.ToDecimal(reader["width_cm"]);
                        box.height_cm = Convert.ToDecimal(reader["height_cm"]);
                        box.gross_weight_kg = Convert.ToDecimal(reader["gross_weight_kg"]);
                        box.net_weight_kg = Convert.ToDecimal(reader["net_weight_kg"]);

                        var boxId = Convert.ToInt32(reader["id_box"]);
                        box.items = GET_NON_ASSETS_FOR_BOX(boxId, conn);

                        model.boxes.Add(box);
                    }
                }

                conn.Close();
            }

            return model;
        }

        private List<ShippingNonAssetItemExportModel> GET_NON_ASSETS_FOR_BOX(int boxId, SqlConnection conn)
        {
            var items = new List<ShippingNonAssetItemExportModel>();

            using (SqlCommand cmd = new SqlCommand("GET_NON_ASSETS_FOR_BOX", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@id_box", boxId);
                using SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var item = new ShippingNonAssetItemExportModel();
                    item.po_or_asset_no = reader["po_or_asset_no"].ToString();
                    item.description = reader["description"].ToString();
                    item.qty = Convert.ToDecimal(reader["qty"]);
                    item.uom = reader["uom"].ToString();
                    item.price_value = Convert.ToDecimal(reader["price_value"]);
                    item.hs_code = reader["hs_code"] != DBNull.Value ? Convert.ToDecimal(reader["hs_code"]) : null;
                    items.Add(item);
                }
            }

            return items;
        }

        public string UPDATE_SHIPPING_NON_ASSET_ITEM_HS_CODE(int id_detail, string po_or_asset_no, string hs_code)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_SHIPPING_NON_ASSET_ITEM_HS_CODE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_detail", id_detail);
                    cmd.Parameters.AddWithValue("@po_or_asset_no", po_or_asset_no);

                    if (string.IsNullOrEmpty(hs_code))
                    {
                        cmd.Parameters.AddWithValue("@hs_code", DBNull.Value);
                    }
                    else
                    {
                        if (decimal.TryParse(hs_code, out decimal hsCodeValue))
                        {
                            cmd.Parameters.AddWithValue("@hs_code", hsCodeValue);
                        }
                        else
                        {
                            return "error;Invalid HS Code format";
                        }
                    }

                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string UPDATE_SECURITY_NON_ASSET_GATEPASS(int id_gatepass, string approval_security_status, string security_name, string approval_remark)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_SECURITY_NON_ASSET_GATEPASS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@approval_security_status", approval_security_status);
                    cmd.Parameters.AddWithValue("@security_name", security_name);
                    cmd.Parameters.AddWithValue("@approval_remark", approval_remark);
                    cmd.ExecuteNonQuery();
                    return "success";
                }
            }
        }

        public string GET_STATUS_NON_ASSET_GATEPASS(int id_gatepass)
        {
            string status_gatepass = string.Empty;
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_STATUS_GATEPASS_NON_ASSET", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    status_gatepass = (string)cmd.ExecuteScalar();
                }
                conn.Close();
            }
            return status_gatepass;
        }

        public string UPDATE_GATEPASS_NON_ASSET_RETURN(int id_gatepass, string approval_security_status, string security_name, string approval_remark)
        {
            string gatepass_no = "";
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_GATEPASS_NON_ASSET_RETURN", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@approval_security_status", approval_security_status);
                    cmd.Parameters.AddWithValue("@security_name", security_name);
                    cmd.Parameters.AddWithValue("@approval_remark", approval_remark);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            gatepass_no = reader["gatepass_no"].ToString() ?? "";
                        }
                    }
                    return gatepass_no;
                }
            }
        }

        public List<GatePassInvoiceAssetModel> GET_GATEPASS_INVOICE_ASSETS(int id_gatepass)
        {
            List<GatePassInvoiceAssetModel> assetList = new List<GatePassInvoiceAssetModel>();
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GATEPASS_INVOICE_ASSETS", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);

                    using SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        GatePassInvoiceAssetModel row = new GatePassInvoiceAssetModel();
                        row.id_detail = Convert.ToInt32(reader["id_detail"]);
                        row.asset_no = reader["asset_no"].ToString();
                        row.asset_subnumber = reader["asset_subnumber"].ToString();
                        row.asset_desc = reader["asset_desc"].ToString();
                        row.asset_class = reader["asset_class"].ToString();
                        row.cost_center = reader["cost_center"].ToString();
                        row.capitalized_on = reader["capitalized_on"] != DBNull.Value
                            ? Convert.ToDateTime(reader["capitalized_on"]).ToString("dd MMM yyyy")
                            : "";
                        row.id_invoice_detail = reader["id_invoice_detail"] != DBNull.Value
                            ? Convert.ToInt32(reader["id_invoice_detail"])
                            : (int?)null;
                        row.amount = Convert.ToDecimal(reader["amount"]);
                        assetList.Add(row);
                    }
                }
                conn.Close();
                    }
            return assetList;
        }

        public string INSERT_GATEPASS_INVOICE_DETAIL(int id_gatepass, string selected_assets, decimal amount, string sesa_id)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("INSERT_GATEPASS_INVOICE_DETAIL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@selected_assets", selected_assets);
                    cmd.Parameters.AddWithValue("@amount", amount);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);

                    SqlParameter resultParam = new SqlParameter("@result", SqlDbType.NVarChar, 500);
                    resultParam.Direction = ParameterDirection.Output;
                    cmd.Parameters.Add(resultParam);

                    cmd.ExecuteNonQuery();

                    return resultParam.Value?.ToString() ?? "success";
                }
            }
        }

        public string DELETE_GATEPASS_INVOICE_DETAIL(int id_invoice_detail, string sesa_id)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("DELETE_GATEPASS_INVOICE_DETAIL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_invoice_detail", id_invoice_detail);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);

                    SqlParameter resultParam = new SqlParameter("@result", SqlDbType.NVarChar, 500);
                    resultParam.Direction = ParameterDirection.Output;
                    cmd.Parameters.Add(resultParam);

                    cmd.ExecuteNonQuery();

                    return resultParam.Value?.ToString() ?? "success";
                }
            }
        }

        public string UPDATE_GATEPASS_INVOICE_DETAIL(int id_invoice_detail, decimal amount, string sesa_id)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_GATEPASS_INVOICE_DETAIL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_invoice_detail", id_invoice_detail);
                    cmd.Parameters.AddWithValue("@amount", amount);
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);

                    SqlParameter resultParam = new SqlParameter("@result", SqlDbType.NVarChar, 500);
                    resultParam.Direction = ParameterDirection.Output;
                    cmd.Parameters.Add(resultParam);

                    cmd.ExecuteNonQuery();

                    return resultParam.Value?.ToString() ?? "success";
                }
            }
        }

        public string INSERT_GATEPASS_INVOICE(int id_gatepass, string invoice_currency, DateTime invoice_date, string created_by, string invoice_main_file, string invoice_secondary_file)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand("INSERT_GATEPASS_INVOICE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                    cmd.Parameters.AddWithValue("@invoice_currency", string.IsNullOrEmpty(invoice_currency) ? (object)DBNull.Value : invoice_currency);
                    cmd.Parameters.AddWithValue("@invoice_date", invoice_date);
                    cmd.Parameters.AddWithValue("@created_by", created_by);
                    cmd.Parameters.AddWithValue("@invoice_main_file", invoice_main_file);
                    cmd.Parameters.AddWithValue("@invoice_secondary_file", string.IsNullOrEmpty(invoice_secondary_file) ? (object)DBNull.Value : invoice_secondary_file);

                    conn.Open();
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string result = reader["result"].ToString();
                            string invoice_no = reader["invoice_no"].ToString();
                            string error_message = reader["error_message"].ToString();

                            if (result == "success")
                            {
                                return $"success;{invoice_no}";
                            }
                            else
                            {
                                return $"error;{error_message}";
                            }
                        }
                    }
                }
            }
            return "error;Unknown error occurred";
        }

        public InvoiceExportModel GET_GATEPASS_INVOICE_EXPORT_DATA(int id_invoice)
        {
            var model = new InvoiceExportModel();

            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_GATEPASS_INVOICE_EXPORT_DATA", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_invoice", id_invoice);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            model.Header = new InvoiceHeaderModel
                            {
                                invoice_no = reader["invoice_no"].ToString(),
                                invoice_currency = reader["invoice_currency"] != DBNull.Value
                                    ? reader["invoice_currency"].ToString()
                                    : "IDR",
                                invoice_date = Convert.ToDateTime(reader["invoice_date"]),
                                gatepass_no = reader["gatepass_no"].ToString(),
                                vendor_name = reader["vendor_name"].ToString(),
                                vendor_address = reader["vendor_address"]?.ToString() ?? "",
                                created_by = reader["created_by"].ToString(),
                                created_by_name = reader["created_by_name"].ToString(),
                                approval_by_name = reader["approval_by_name"].ToString(),
                                approval_date = Convert.ToDateTime(reader["approval_date"]),
                            };
                        }

                        if (reader.NextResult())
                        {
                            model.Details = new List<InvoiceDetailModel>();
                            while (reader.Read())
                            {
                                model.Details.Add(new InvoiceDetailModel
                                {
                                    id_invoice_detail = Convert.ToInt32(reader["id_invoice_detail"]),
                                    asset_no = reader["asset_no"].ToString(),
                                    asset_subnumber = reader["asset_subnumber"].ToString(),
                                    asset_desc = reader["asset_desc"].ToString(),
                                    asset_class = reader["asset_class"].ToString(),
                                    amount = Convert.ToDecimal(reader["amount"])
                                });
                            }
                        }
                    }
                }
                conn.Close();
            }

            return model;
        }

        public string UPDATE_APPROVAL_GATEPASS_INVOICE(int id_invoice, string approval_status, string approval_remark, string approval_sesa)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_APPROVAL_GATEPASS_INVOICE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_invoice", id_invoice);
                    cmd.Parameters.AddWithValue("@approval_status", approval_status);
                    cmd.Parameters.AddWithValue("@approval_remark", approval_remark ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@approval_sesa", approval_sesa);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string result = reader["result"].ToString();
                            string message = reader["message"].ToString();
                            string invoice_no = reader["invoice_no"]?.ToString() ?? "";
                            return $"{result};{message};{invoice_no}";
                        }
                    }
                    return "error;Unknown error;";
                }
            }
        }

        public string UPDATE_GATEPASS_INVOICE_PAYMENT(int id_invoice, string invoice_payment_file, string uploaded_by)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("UPDATE_GATEPASS_INVOICE_PAYMENT", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_invoice", id_invoice);
                    cmd.Parameters.AddWithValue("@invoice_payment_file", invoice_payment_file);
                    cmd.Parameters.AddWithValue("@uploaded_by", uploaded_by);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        string resultValue = "error;No result returned";

                        do
                        {
                            if (reader.HasRows && reader.Read())
                            {
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    if (reader.GetName(i).Equals("result", StringComparison.OrdinalIgnoreCase))
                                    {
                                        string result = reader["result"].ToString();
                                        if (result == "success")
                                        {
                                            resultValue = "success";
                                        }
                                        else
                                        {
                                            string errorMessage = reader["error_message"]?.ToString() ?? "Unknown error";
                                            resultValue = $"error;{errorMessage}";
                                        }
                                        break;
                                    }
                                }
                            }
                        } while (reader.NextResult());

                        return resultValue;
                    }
                }
            }
        }

        public InvoiceDetailViewModel GET_INVOICE_DETAIL(int id_invoice)
        {
            var result = new InvoiceDetailViewModel();

            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_INVOICE_DETAIL", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_invoice", id_invoice);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            result.id_gatepass = reader["id_gatepass"] != DBNull.Value
                                ? Convert.ToInt32(reader["id_gatepass"])
                                : (int?)null;
                            result.invoice_main_file = reader["invoice_main_file"] != DBNull.Value
                                ? reader["invoice_main_file"].ToString()
                                : null;
                            result.invoice_secondary_file = reader["invoice_secondary_file"] != DBNull.Value
                                ? reader["invoice_secondary_file"].ToString()
                                : null;
                            result.invoice_payment = reader["invoice_payment"] != DBNull.Value
                                ? reader["invoice_payment"].ToString()
                                : null;
                            result.invoice_currency = reader["invoice_currency"] != DBNull.Value
                                ? reader["invoice_currency"].ToString()
                                : null;
                        }

                        if (reader.NextResult())
                        {
                            result.assets = new List<GatePassInvoiceAssetModel>();
                            while (reader.Read())
                            {
                                GatePassInvoiceAssetModel asset = new GatePassInvoiceAssetModel();
                                asset.id_detail = Convert.ToInt32(reader["id_detail"]);
                                asset.asset_no = reader["asset_no"].ToString();
                                asset.asset_subnumber = reader["asset_subnumber"].ToString();
                                asset.asset_desc = reader["asset_desc"].ToString();
                                asset.asset_class = reader["asset_class"].ToString();
                                asset.cost_center = reader["cost_center"].ToString();
                                asset.capitalized_on = reader["capitalized_on"] != DBNull.Value
                                    ? Convert.ToDateTime(reader["capitalized_on"]).ToString("dd MMM yyyy")
                                    : "";
                                asset.id_invoice_detail = reader["id_invoice_detail"] != DBNull.Value
                                    ? Convert.ToInt32(reader["id_invoice_detail"])
                                    : (int?)null;
                                asset.amount = Convert.ToDecimal(reader["amount"]);
                                result.assets.Add(asset);
                            }
                        }
                    }
                }
                conn.Close();
            }

            return result;
        }

        public string GET_INVOICE_PAYMENT_FILE(int id_invoice)
        {
            string filename = null;
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("GET_INVOICE_PAYMENT_FILE", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_invoice", id_invoice);

                    var result = cmd.ExecuteScalar();
                    filename = result?.ToString();
                }
                conn.Close();
            }
            return filename;
        }

    }
}
