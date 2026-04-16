using System.Data;
using System.IO;
using Microsoft.AspNetCore.Http;
using System.Collections.Generic;
using System.Data.SqlClient;
using Microsoft.AspNetCore.Components;
using System;
using Asset_Management.Function;
using System.Threading.Tasks;

namespace Asset_Management.Service
{
    public class ImportExportFactory
    {

        public readonly string ConnectionString = new DatabaseAccessLayer().ConnectionString;
        private readonly FileManagementService _fileManagement;
        private readonly ExcelServiceProvider _excelService;

        public ImportExportFactory(FileManagementService fileManagement,
            ExcelServiceProvider excelService)
        {
            _fileManagement = fileManagement;
            _excelService = excelService;
        }

        public ImportExportFactory() :
            this(new FileManagementService(),
                new ExcelServiceProvider())
        {
        }

        public void ImportAsset(IFormFile file, string id_login, string sesa_id)
        {
            string query = "DELETE FROM temp_asset WHERE uploaded_by='" + sesa_id + "'";
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                using (var cmd = new SqlCommand(query, conn))
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }
            }

            var uploadedFilePath = _fileManagement.UploadFile(file);

            if (uploadedFilePath == "Upload Fail")
            {
                return;
            }
            var fileInfo = new FileInfo(uploadedFilePath);

            var dataTable = _excelService.Excel_To_DataTable(fileInfo);

            BulkInsertAsset(dataTable, id_login, sesa_id);
        }
        public void BulkInsertAsset(DataTable tbl, string id_login, string sesa_id)
        {
            tbl.Columns.Add("id_login", typeof(string));
            tbl.Columns.Add("uploaded_by", typeof(string));

            foreach (DataRow row in tbl.Rows)
            {
                row["id_login"] = id_login;
                row["uploaded_by"] = sesa_id;
            }

            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(conn))
                {
                    sqlBulkCopy.DestinationTableName = "dbo.temp_asset";

                    sqlBulkCopy.ColumnMappings.Add("Column 1", "asset_no");
                    sqlBulkCopy.ColumnMappings.Add("Column 2", "asset_subnumber");
                    sqlBulkCopy.ColumnMappings.Add("Column 3", "asset_desc");
                    sqlBulkCopy.ColumnMappings.Add("Column 4", "asset_class");
                    sqlBulkCopy.ColumnMappings.Add("Column 5", "cost_center");
                    sqlBulkCopy.ColumnMappings.Add("Column 6", "capitalized_on");
                    sqlBulkCopy.ColumnMappings.Add("Column 7", "apc_fy_start");
                    sqlBulkCopy.ColumnMappings.Add("Column 8", "acquisition");
                    sqlBulkCopy.ColumnMappings.Add("Column 9", "retirement");
                    sqlBulkCopy.ColumnMappings.Add("Column 10", "transfer");
                    sqlBulkCopy.ColumnMappings.Add("Column 11", "current_apc");
                    sqlBulkCopy.ColumnMappings.Add("Column 12", "dep_fy_start");
                    sqlBulkCopy.ColumnMappings.Add("Column 13", "dep_for_year");
                    sqlBulkCopy.ColumnMappings.Add("Column 14", "dep_retir");
                    sqlBulkCopy.ColumnMappings.Add("Column 15", "dep_transfer");
                    sqlBulkCopy.ColumnMappings.Add("Column 16", "accumul_dep");
                    sqlBulkCopy.ColumnMappings.Add("Column 17", "bk_val_fy");
                    sqlBulkCopy.ColumnMappings.Add("Column 18", "curr_bk_val");
                    sqlBulkCopy.ColumnMappings.Add("Column 19", "currency");
                    sqlBulkCopy.ColumnMappings.Add("Column 20", "department");
                    sqlBulkCopy.ColumnMappings.Add("Column 21", "plant");
                    sqlBulkCopy.ColumnMappings.Add("Column 22", "remarks");
                    sqlBulkCopy.ColumnMappings.Add("Column 23", "sesa_owner");
                    sqlBulkCopy.ColumnMappings.Add("Column 24", "tagging_status");
                    sqlBulkCopy.ColumnMappings.Add("Column 25", "vendor_name");
                    sqlBulkCopy.ColumnMappings.Add("Column 26", "asset_location");
                    sqlBulkCopy.ColumnMappings.Add("Column 27", "asset_location_address");
                    sqlBulkCopy.ColumnMappings.Add("id_login", "id_login");
                    sqlBulkCopy.ColumnMappings.Add("uploaded_by", "uploaded_by");

                    conn.Open();
                    sqlBulkCopy.WriteToServer(tbl);
                    conn.Close();
                }
            }

            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("INSERT_ASSET_LIST", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@id_login", id_login);
                cmd.Parameters.AddWithValue("@uploaded_by", sesa_id);
                cmd.ExecuteNonQuery();
                conn.Close();
            }
        }
        public string ImportFileRename(IFormFile file_doc, string filename, string subfolder)
        {
            var uploadedFilePath = _fileManagement.UploadFileRename(file_doc, filename, subfolder);
            return uploadedFilePath;
        }

    }
}
