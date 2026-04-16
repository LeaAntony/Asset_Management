using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office.Word;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Asset_Management.Function;
using Asset_Management.Models;
using Asset_Management.Service;
using System.Data;
using System.Data.SqlClient;
using System.Linq.Dynamic.Core;
using System.Security.Claims;

namespace Asset_Management.Controllers
{
    [Authorize(Policy = "RequireAny")]
    public class RequestorController : Controller
    {
        private readonly ApplicationDbContext _context;

        private readonly ImportExportFactory _importexportFactory;
        private readonly ILogger<RequestorController> _logger;

        public IActionResult NewRequest()
        {
            return View();
        }

        private string DbConnection()
        {
            var dbAccess = new DatabaseAccessLayer();
            string dbString = dbAccess.ConnectionString;
            return dbString;
        }

        public RequestorController(ImportExportFactory importexportFactory, ILogger<RequestorController> logger, ApplicationDbContext context)
        {
            this._context = context;
            _importexportFactory = importexportFactory;
            _logger = logger;
        }
        public IActionResult Profile()
        {
            return this.CheckSession(() =>
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                List<UserDetailModel> userList = db.GetUserDetail(sesa_id);
                return View(userList);
            });
        }

        [HttpGet]
        public IActionResult GetAssetListBadge()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            int totalNoTagging = db.GetTotalNoTagging(sesa_id);
            int totalAssetNeedCount = db.GetTotalAssetNeedCount(sesa_id);
            int total = totalNoTagging + totalAssetNeedCount;
            return Json(new { count = total });
        }

        public IActionResult GetGatepassListBadge()
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                List<UserDetailModel> userDetail = db.GetUserDetail(sesa_id);
                var user = userDetail.FirstOrDefault();
                string currentDate = DateTime.Now.ToString();
                if (user == null)
                {
                    return Ok(new { count = 0 });
                }
                var count = (from GatepassList in _context.v_gatepass
                             where GatepassList.created_by == sesa_id
                             && GatepassList.status_gatepass != "COMPLETED"
                             && GatepassList.status_gatepass != "REJECT"
                             && GatepassList.status_gatepass != "CANCEL"
                             && ((GatepassList.has_overdue == 1)
                             || (GatepassList.status_gatepass == "TRANSFER"
                             && (GatepassList.category == "TRANSFER PLANT TO PLANT"
                             || GatepassList.category == "TRANSFER INSIDE PLANT")))
                             select GatepassList).Count();
                return Ok(new { count = count });
            }
            catch (Exception)
            {
                return Ok(new { count = 0 });
            }
        }

        public IActionResult GetPicConfirmBadge()
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                List<UserDetailModel> userDetail = db.GetUserDetail(sesa_id);
                var user = userDetail.FirstOrDefault();
                string currentDate = DateTime.Now.ToString();
                if (user == null)
                {
                    return Ok(new { count = 0 });
                }
                var count = (from PicConfirm in _context.v_gatepass
                             where PicConfirm.new_pic == sesa_id
                             && PicConfirm.status_gatepass == "TRANSFER"
                             && (PicConfirm.category == "TRANSFER PLANT TO PLANT"
                             || PicConfirm.category == "TRANSFER INSIDE PLANT")
                             select PicConfirm).Count();
                return Ok(new { count = count });
            }
            catch (Exception)
            {
                return Ok(new { count = 0 });
            }
        }

        public IActionResult GetGatepassNonAssetListBadge()
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                List<UserDetailModel> userDetail = db.GetUserDetail(sesa_id);
                var user = userDetail.FirstOrDefault();
                string currentDate = DateTime.Now.ToString();
                if (user == null)
                {
                    return Ok(new { count = 0 });
                }
                var count = (from GatepassNonAssetList in _context.v_gatepass_non_asset
                             where GatepassNonAssetList.has_overdue == 1
                             && GatepassNonAssetList.created_by == sesa_id
                             && GatepassNonAssetList.status_gatepass != "COMPLETED"
                             && GatepassNonAssetList.status_gatepass != "REJECT"
                             && GatepassNonAssetList.status_gatepass != "CANCEL"
                             select GatepassNonAssetList).Count();
                return Ok(new { count = count });
            }
            catch (Exception)
            {
                return Ok(new { count = 0 });
            }
        }
        public IActionResult GatePass(string assets = "")
        {
            return this.CheckSession(() =>
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                string delTemp = db.DELETE_TEMP_GATEPASS_SESA(sesa_id);
                List<UserDetailModel> userDetail = db.GetUserDetail(sesa_id);
                List<string> catList = db.GET_SECURITY_GUARD();
                List<UserDetailModel> picList = db.GET_NEW_PIC();
                List<LotFamilyModel> lotFamilyList = db.GetLotFamily();
                List<LotMachineModel> lotMachineList = db.GetLocMachine();
                List<ProductFamilyModel> productFamilyList = db.GetProductFamily();
                List<UserDetailModel> employeeList = db.GET_EMPLOYEE();

                ViewBag.lotFamilyList = lotFamilyList;
                ViewBag.catList = catList;
                ViewBag.picList = picList;
                ViewBag.employeeList = employeeList;
                ViewBag.assets = assets;
                ViewBag.userDetail = userDetail;
                ViewBag.lotMachineList = lotMachineList;
                ViewBag.productFamilyList = productFamilyList;
                return View(userDetail);
            });
        }
        public IActionResult GatePassNew(string assets = "")
        {
            return this.CheckSession(() =>
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                string delTemp = db.DELETE_TEMP_GATEPASS_SESA(sesa_id);
                List<UserDetailModel> userDetail = db.GetUserDetail(sesa_id);
                List<string> catList = db.GET_SECURITY_GUARD();
                List<UserDetailModel> picList = db.GET_NEW_PIC();
                List<LotFamilyModel> lotFamilyList = db.GetLotFamily();
                List<LotMachineModel> lotMachineList = db.GetLocMachine();
                List<ProductFamilyModel> productFamilyList = db.GetProductFamily();
                List<UserDetailModel> employeeList = db.GET_EMPLOYEE();

                ViewBag.lotFamilyList = lotFamilyList;
                ViewBag.catList = catList;
                ViewBag.picList = picList;
                ViewBag.employeeList = employeeList;
                ViewBag.assets = assets;
                ViewBag.userDetail = userDetail;
                ViewBag.lotMachineList = lotMachineList;
                ViewBag.productFamilyList = productFamilyList;
                return View(userDetail);
            });
        }
        public IActionResult GatePassList(string gatepass_no = null)
        {
            return this.CheckSession(() =>
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                List<UserDetailModel> userDetail = db.GetUserDetail(sesa_id);
                ViewBag.gatepass_no = gatepass_no;
                return View(userDetail);
            });
        }

        public IActionResult AssetList()
        {
            return this.CheckSession(() =>
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                List<UserDetailModel> userDetail = db.GetUserDetail(sesa_id);
                int totalNoTagging = db.GetTotalNoTagging(sesa_id);
                int totalAssetNeedCount = db.GetTotalAssetNeedCount(sesa_id);
                ViewBag.totalNoTagging = totalNoTagging;
                ViewBag.totalAssetNeedCount = totalAssetNeedCount;
                return View(userDetail);
            });
        }
        public IActionResult GetAssetNoTagging()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<AssetListModel> dataList = db.GetAssetNoTagging(sesa_id);
            return PartialView("_TableAssetNoTagging", dataList);
        }
        public IActionResult GetAssetNeedCount()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<AssetListModel> dataList = db.GetAssetNeedCount(sesa_id);
            return PartialView("_TableAssetNeedCount", dataList);
        }

        public IActionResult DOWNLOAD_ASSET_COUNT()
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                DateTime currentDateTime = DateTime.Now;
                string formattedDateTime = currentDateTime.ToString("yyyy.MM.dd - HH:mm:ss");

                DataSet ds = GET_ASSET_COUNT();
                DataTable dt = ds.Tables[0];
                var ws = wb.Worksheets.Add(dt);

                foreach (var column in ws.Columns())
                {
                    column.AdjustToContents();
                }

                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ASSET NEED COUNT " + formattedDateTime + ".xlsx");
                }
            }
        }
        private DataSet GET_ASSET_COUNT()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            DataSet ds = new DataSet();

            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                string query = @"DECLARE @departments NVARCHAR(MAX);
                SELECT @departments = other_dept 
                FROM [v_detail_user] 
                WHERE sesa_id = @sesa_id;

                SELECT 
                    a.count_year, 
                    DateName(month, DateAdd(month, a.count_month, -1)) AS count_month, 
                    a.asset_no, 
                    a.asset_subnumber, 
                    b.asset_desc, 
                    b.asset_class, 
                    b.cost_center, 
                    b.capitalized_on,
                    b.department,
                    u.name AS asset_owner
                FROM 
                    tbl_asset_count a 
                JOIN 
                    tbl_asset b ON a.asset_no = b.asset_no AND a.asset_subnumber = b.asset_subnumber 
                LEFT JOIN 
                    mst_users u ON b.sesa_owner = u.sesa_id 
                WHERE 
                    a.is_counted = 0 
                    AND b.department IN (
                        SELECT TRIM(value) 
                        FROM STRING_SPLIT(@departments, ',')
                    );";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@sesa_id", sesa_id);
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(ds);
                    }
                }
            }

            return ds;
        }

        public IActionResult DOWNLOAD_ASSET_TAGGING()
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                DateTime currentDateTime = DateTime.Now;
                string formattedDateTime = currentDateTime.ToString("yyyy.MM.dd - HH:mm:ss");

                DataSet ds = GET_ASSET_TAGGING();
                DataTable dt = ds.Tables[0];
                var ws = wb.Worksheets.Add(dt);

                foreach (var column in ws.Columns())
                {
                    column.AdjustToContents();
                }

                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ASSET NO TAGGING " + formattedDateTime + ".xlsx");
                }
            }
        }
        private DataSet GET_ASSET_TAGGING()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            DataSet ds = new DataSet();

            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                string query = @"SELECT asset_no, asset_subnumber, asset_desc, asset_class, cost_center, capitalized_on
                         FROM v_asset 
                         WHERE tagging_status='NO' AND (@sesa_owner='' OR sesa_owner=@sesa_owner)";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@sesa_owner", sesa_id);
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(ds);
                    }
                }
            }

            return ds;
        }

        public IActionResult GetAssetList()
        {
            try
            {
                var draw = Request.Form["draw"].FirstOrDefault();
                var start = Request.Form["start"].FirstOrDefault();
                var length = Request.Form["length"].FirstOrDefault();
                var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][data]"].FirstOrDefault();
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                var searchValue = Request.Form["search[value]"].FirstOrDefault();
                var column0Value = Request.Form["columns[0][search][value]"];
                var column1Value = Request.Form["columns[1][search][value]"];
                var column2Value = Request.Form["columns[2][search][value]"];
                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;
                int recordsTotal = 0;
                var mstData = (from AssetList in _context.v_asset
                               select
                                   new
                                   {
                                       AssetList.id_asset,
                                       AssetList.asset_no,
                                       AssetList.asset_subnumber,
                                       AssetList.asset_desc,
                                       AssetList.asset_class,
                                       AssetList.cost_center,
                                       AssetList.cc_desc,
                                       AssetList.cc_grouping,
                                       AssetList.cc_plant,
                                       capitalized_on = Convert.ToString(AssetList.capitalized_on),
                                       AssetList.apc_fy_start,
                                       AssetList.acquisition,
                                       AssetList.retirement,
                                       AssetList.transfer,
                                       AssetList.current_apc,
                                       AssetList.dep_fy_start,
                                       AssetList.dep_for_year,
                                       AssetList.dep_retir,
                                       AssetList.dep_transfer,
                                       AssetList.accumul_dep,
                                       AssetList.bk_val_fy,
                                       AssetList.curr_bk_val,
                                       AssetList.currency,
                                       AssetList.department,
                                       AssetList.plant,
                                       AssetList.sesa_owner,
                                       AssetList.name_owner,
                                       AssetList.tagging_status,
                                       AssetList.vendor_name,
                                       AssetList.asset_location,
                                       AssetList.asset_location_address,
                                       AssetList.file_tag,
                                       AssetList.file_tag_status,
                                       AssetList.filename_doc
                                   });

                if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDirection)))
                {
                    mstData = mstData.OrderBy(sortColumn + " " + sortColumnDirection);
                }
                if (!string.IsNullOrEmpty(searchValue))
                {
                    mstData = mstData.Where(m => m.asset_no.Contains(searchValue)
                                                || m.asset_subnumber.Contains(searchValue)
                                                || m.asset_desc.Contains(searchValue)
                                                || m.asset_class.Contains(searchValue)
                                                || m.cost_center.Contains(searchValue));
                }
                for (int i = 0; i < 31; i++)
                {
                    var searchColVal = Request.Form["columns[" + i.ToString() + "][search][value]"];
                    var fieldName = Request.Form["columns[" + i.ToString() + "][data]"].FirstOrDefault();
                    if (!string.IsNullOrEmpty(searchColVal))
                    {
                        if (fieldName == "asset_no")
                        {
                            mstData = mstData.Where(m => m.asset_no.Contains(searchColVal));
                        }
                        else if (fieldName == "asset_subnumber")
                        {
                            mstData = mstData.Where(m => m.asset_subnumber.Contains(searchColVal));
                        }
                        else if (fieldName == "asset_desc")
                        {
                            mstData = mstData.Where(m => m.asset_desc.Contains(searchColVal));
                        }
                        else if (fieldName == "asset_class")
                        {
                            mstData = mstData.Where(m => m.asset_class.Contains(searchColVal));
                        }
                        else if (fieldName == "cost_center")
                        {
                            mstData = mstData.Where(m => m.cost_center.Contains(searchColVal));
                        }
                        else if (fieldName == "cc_desc")
                        {
                            mstData = mstData.Where(m => m.cc_desc.Contains(searchColVal));
                        }
                        else if (fieldName == "cc_grouping")
                        {
                            mstData = mstData.Where(m => m.cc_grouping.Contains(searchColVal));
                        }
                        else if (fieldName == "cc_plant")
                        {
                            mstData = mstData.Where(m => m.cc_plant.Contains(searchColVal));
                        }
                        else if (fieldName == "capitalized_on")
                        {
                            mstData = mstData.Where(m => m.capitalized_on.Contains(searchColVal));
                        }
                        else if (fieldName == "apc_fy_start")
                        {
                            mstData = mstData.Where(m => m.apc_fy_start.ToString().Contains(searchColVal));
                        }
                        else if (fieldName == "acquisition")
                        {
                            mstData = mstData.Where(m => m.acquisition.ToString().Contains(searchColVal));
                        }
                        else if (fieldName == "retirement")
                        {
                            mstData = mstData.Where(m => m.retirement.ToString().Contains(searchColVal));
                        }
                        else if (fieldName == "transfer")
                        {
                            mstData = mstData.Where(m => m.transfer.ToString().Contains(searchColVal));
                        }
                        else if (fieldName == "current_apc")
                        {
                            mstData = mstData.Where(m => m.current_apc.ToString().Contains(searchColVal));
                        }
                        else if (fieldName == "dep_fy_start")
                        {
                            mstData = mstData.Where(m => m.dep_fy_start.ToString().Contains(searchColVal));
                        }
                        else if (fieldName == "dep_for_year")
                        {
                            mstData = mstData.Where(m => m.dep_for_year.ToString().Contains(searchColVal));
                        }
                        else if (fieldName == "dep_retir")
                        {
                            mstData = mstData.Where(m => m.dep_retir.ToString().Contains(searchColVal));
                        }
                        else if (fieldName == "dep_transfer")
                        {
                            mstData = mstData.Where(m => m.dep_transfer.ToString().Contains(searchColVal));
                        }
                        else if (fieldName == "accumul_dep")
                        {
                            mstData = mstData.Where(m => m.accumul_dep.ToString().Contains(searchColVal));
                        }
                        else if (fieldName == "bk_val_fy")
                        {
                            mstData = mstData.Where(m => m.bk_val_fy.ToString().Contains(searchColVal));
                        }
                        else if (fieldName == "curr_bk_val")
                        {
                            mstData = mstData.Where(m => m.curr_bk_val.ToString().Contains(searchColVal));
                        }
                        else if (fieldName == "currency")
                        {
                            mstData = mstData.Where(m => m.currency.Contains(searchColVal));
                        }
                        else if (fieldName == "department")
                        {
                            mstData = mstData.Where(m => m.department.Contains(searchColVal));
                        }
                        else if (fieldName == "plant")
                        {
                            mstData = mstData.Where(m => m.plant.Contains(searchColVal));
                        }
                        else if (fieldName == "name_owner")
                        {
                            mstData = mstData.Where(m => m.name_owner.Contains(searchColVal));
                        }
                        else if (fieldName == "tagging_status")
                        {
                            mstData = mstData.Where(m => m.tagging_status.Contains(searchColVal));
                        }
                        else if (fieldName == "vendor_name")
                        {
                            mstData = mstData.Where(m => m.vendor_name.Contains(searchColVal));
                        }
                        else if (fieldName == "file_tag_status")
                        {
                            mstData = mstData.Where(m => m.file_tag_status.Contains(searchColVal));
                        }
                        else if (fieldName == "asset_location")
                        {
                            mstData = mstData.Where(m => m.asset_location.Contains(searchColVal));
                        }
                        else if (fieldName == "asset_location_address")
                        {
                            mstData = mstData.Where(m => m.asset_location_address.Contains(searchColVal));
                        }
                        else if (fieldName == "filename_doc")
                        {
                            mstData = mstData.Where(m => m.filename_doc.Contains(searchColVal));
                        }
                    }
                }
                recordsTotal = mstData.Count();
                var data = mstData.Skip(skip).Take(pageSize).ToList();
                var jsonData = new { draw = draw, recordsFiltered = recordsTotal, recordsTotal = recordsTotal, data = data };
                return Ok(jsonData);
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        
        public IActionResult UpdateTaggingAsset(IFormFileCollection file_tagging, string asset_no, string asset_subnumber)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            var email = new EmailController();
            var exp = new ExportController();

            string filename = asset_no + "_" + asset_subnumber + "_TAG.pdf";
            string ret_file = exp.CreateAssetTag(file_tagging, filename, asset_no, asset_subnumber);
            string[] res = ret_file.Split(';');
            if (res[0] == "ERROR")
            {
                return Content(res[1], "text/plain");
            }

            string status_msg = db.UpdateTaggingAsset(asset_no, asset_subnumber, res[1], sesa_id);
            string send_email = "success";

            return Content(status_msg + ";" + send_email + ";Succesfully updated!", "text/plain");
        }
        public IActionResult GetCountedResult(string count_year, string asset_no, string asset_subnumber)
        {
            var db = new DatabaseAccessLayer();
            List<AssetCountListModel> resList = db.GetCountedResult(count_year, asset_no, asset_subnumber);
            return Json(resList);
        }
        [HttpPost]
        public IActionResult SubmitAssetCount(IFormFileCollection file_imgs, string count_year, string asset_no, string asset_subnumber, string existence, string good_condition, string still_in_operation, string tagging_available, string applicable_of_tagging, string correct_naming, string correct_location, string retagging)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            var email = new EmailController();
            var exp = new ExportController();
            string filename = "";

            if (file_imgs != null && file_imgs.Count != 0)
            {
                filename = count_year + "_" + asset_no + "_" + asset_subnumber + ".pdf";
                string ret_file = exp.CreateAssetCount(file_imgs, filename, count_year, asset_no, asset_subnumber);
                string[] res = ret_file.Split(';');
                if (res[0] == "ERROR")
                {
                    return Content(res[1], "text/plain");
                }
            }
            string status_msg = db.SubmitAssetCount(count_year, asset_no, asset_subnumber, existence, good_condition, still_in_operation, tagging_available, applicable_of_tagging, correct_naming, correct_location, retagging, filename, sesa_id);
            string send_email = "success";

            return Content(status_msg + ";" + send_email + ";Succesfully updated!", "text/plain");
        }
        public IActionResult UPDATE_RETAGGING_IMG(IFormFileCollection file_imgs, string count_year, string asset_no, string asset_subnumber, string id_count)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            var email = new EmailController();
            var exp = new ExportController();
            string filename = "";

            if (file_imgs != null && file_imgs.Count != 0)
            {
                filename = count_year + "_" + asset_no + "_" + asset_subnumber + "_RETAGGING" + ".pdf";
                string ret_file = exp.CreateAssetRetagging(file_imgs, filename, count_year, asset_no, asset_subnumber);
                string[] res = ret_file.Split(';');
                if (res[0] == "ERROR")
                {
                    return Content(res[1], "text/plain");
                }
            }
            string status_msg = db.UPDATE_RETAGGING_IMG(id_count, filename);
            string send_email = "success";

            return Content(status_msg + ";" + send_email + ";Succesfully updated!", "text/plain");
        }

        [HttpPost]
        public IActionResult Recount(int id_count)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            var email = new EmailController();
            var exp = new ExportController();

            string status_msg = db.Recount(id_count, sesa_id);

            return Content(status_msg, "text/plain");
        }
        [HttpPost]
        public IActionResult UpdateCountAssetName(int id_count, string asset_name)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();

            string status_msg = db.UpdateCountAssetName(id_count, asset_name, sesa_id);

            return Content(status_msg, "text/plain");
        }
        [HttpPost]
        public IActionResult UpdateCountAssetLocation(int id_count, string asset_location)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();

            string status_msg = db.UpdateCountAssetLocation(id_count, asset_location, sesa_id);

            return Content(status_msg, "text/plain");
        }
        [HttpGet]
        public IActionResult SearchAsset(string asset_no)
        {
            var db = new DatabaseAccessLayer();
            List<AssetListModel> assetList = db.GetAsset(asset_no);
            return Json(new { items = assetList });
        }

        public IActionResult DOWNLOAD_ASSET_LIST()
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                DateTime currentDateTime = DateTime.Now;
                string formattedDateTime = currentDateTime.ToString("yyyy.MM.dd - HH:mm:ss");

                DataSet ds = GET_ASSET_LIST();
                DataTable dt = ds.Tables[0];
                var ws = wb.Worksheets.Add(dt);

                foreach (var column in ws.Columns())
                {
                    column.AdjustToContents();
                }

                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ASSET LIST " + formattedDateTime + ".xlsx");
                }
            }
        }
        private DataSet GET_ASSET_LIST()
        {
            DataSet ds = new DataSet();

            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                string query = @"
                SELECT 
                    a.[asset_no] AS [Asset No],
                    a.[asset_subnumber] AS [Subnumber],
                    a.[asset_desc] AS [Asset Desc],
                    a.[asset_class] AS [Asset Class],
                    a.[cost_center] AS [Cost Center],
                    a.[capitalized_on] AS [Capitalized On],
                    a.[apc_fy_start] AS [APC FY Start],
                    a.[acquisition] AS [Acquisition],
                    a.[retirement] AS [Retirement],
                    a.[transfer] AS [Transfer],
                    a.[current_apc] AS [Current APC],
                    a.[dep_fy_start] AS [Dep. FY Start],
                    a.[dep_for_year] AS [Dep. For Year],
                    a.[dep_retir] AS [Dep.Retir],
                    a.[dep_transfer] AS [Dep.Transfer],
                    a.[accumul_dep] AS [Accumul.Dep],
                    a.[bk_val_fy] AS [Bk.Val.FY Start],
                    a.[curr_bk_val] AS [Curr.Bk.Val],
                    a.[currency] AS [Currency],
                    a.[department] AS [Department],
                    a.[plant] AS [Plant],
                    u.[name] AS [Owner],
                    a.[tagging_status]  AS [Tagging Status],
                    a.[vendor_name]  AS [Vendor Name]
                FROM 
                    [Asset_Management].[dbo].[tbl_asset] a
                LEFT JOIN 
                    [Asset_Management].[dbo].[mst_users] u
                ON 
                    a.sesa_owner = u.sesa_id;";

                using (SqlCommand cmd = new SqlCommand(query))
                {
                    cmd.Connection = conn;
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(ds);
                    }
                }
            }

            return ds;
        }

        [HttpGet]
        public IActionResult ASSET_TAGGING(string assetNo, string assetSubnumber)
        {
            var db = new DatabaseAccessLayer();
            AssetListModel pdf = db.ASSET_TAGGING(assetNo, assetSubnumber);

            if (pdf == null || string.IsNullOrEmpty(pdf.file_tag))
            {
                return Json(new { success = false, message = "No data available." });
            }

            string filePath = "";
            if (pdf.file_tag.Contains("_TAG"))
            {
                filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/upload/Asset", pdf.file_tag);
            }
            else
            {
                filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/upload/Count", pdf.file_tag);
            }

            if (!System.IO.File.Exists(filePath))
            {
                return Json(new { success = false, message = "File not found." });
            }

            byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
            return File(fileBytes, "application/pdf");
        }


        public async Task<IActionResult> ADD_TEMP_GATEPASS(string asset_no, string asset_subnumber, string naming_output, IFormFile asset_imgs)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            var exp = new ExportController();

            string imageResult = "No file uploaded";
            string filenameSave = string.Empty;
            if (asset_imgs != null && asset_imgs.Length > 0)
            {
                string dateSuffix = DateTime.Now.ToString("yyMMdd");
                string filename = $"{asset_no}_{asset_subnumber}_{dateSuffix}_BEFORE";
                imageResult = exp.CREATE_IMAGE_GP(asset_imgs, filename);

                if (imageResult.StartsWith("OK;"))
                {
                    filenameSave = imageResult.Split(';')[1];
                }
            }

            string addTemp = db.ADD_TEMP_GATEPASS(asset_no, asset_subnumber, sesa_id, naming_output, filenameSave);
            return Content(addTemp, "text/plain");
        }

        [HttpPost]
        public IActionResult ADD_TEMP_GATEPASS_DISPOSAL(string asset_no, string asset_subnumber, string naming_output)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();

            string addTemp = db.ADD_TEMP_GATEPASS_DISPOSAL(asset_no, asset_subnumber, sesa_id, naming_output);
            return Content(addTemp, "text/plain");
        }

        [HttpPost]
        public async Task<IActionResult> UPLOAD_TEMP_GATEPASS_IMAGE(int id_temp, IFormFile asset_img, string asset_no, string asset_subnumber)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            var exp = new ExportController();

            if (asset_img == null || asset_img.Length == 0)
            {
                return Content("error;No file uploaded", "text/plain");
            }

            string dateSuffix = DateTime.Now.ToString("yyMMdd");
            string filename = $"{asset_no}_{asset_subnumber}_{dateSuffix}_BEFORE";
            string imageResult = exp.CREATE_IMAGE_GP(asset_img, filename);

            if (imageResult.StartsWith("OK;"))
            {
                string filenameSave = imageResult.Split(';')[1];
                string updateResult = db.UPDATE_TEMP_GATEPASS_IMAGE(id_temp, filenameSave);
                return Content(updateResult, "text/plain");
            }
            else
            {
                return Content("error;" + imageResult, "text/plain");
            }
        }

        public IActionResult GET_TEMP_GATEPASS()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<DisposalTempModel> dataList = db.GET_TEMP_GATEPASS(sesa_id);
            return Json(dataList);
        }

        [HttpGet]
        public IActionResult GET_VENDOR_DATA_GATEPASS(string vendor_code)
        {
            string query = @"
                            SELECT 
                               [vendor_name]
                              ,[vendor_address]
                              ,[vendor_phone]
                              ,[vendor_email]
                              ,[vendor_batam]
                            FROM 
                                mst_vendor_gatepass 
                            WHERE 
                                vendor_code = @vendor_code OR @vendor_code IS NULL";

            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    if (vendor_code == null)
                    {
                        cmd.Parameters.AddWithValue("@vendor_code", DBNull.Value);
                    }
                    else
                    {
                        cmd.Parameters.AddWithValue("@vendor_code", vendor_code);
                    }

                    conn.Open();

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            var data = new
                            {
                                vendor_name = reader["vendor_name"].ToString(),
                                vendor_address = reader["vendor_address"].ToString(),
                                vendor_phone = reader["vendor_phone"].ToString(),
                                vendor_email = reader["vendor_email"].ToString(),
                                vendor_batam = reader["vendor_batam"].ToString()
                            };

                            return Json(data);
                        }
                    }
                }
            }

            return Json(null);
        }

        public IActionResult ADD_NEW_GATEPASS_WITH_PROFORMA(
            string create_by, string create_date, string category, string new_pic, string type, string location, string employee, string csr, string return_date, string vendor_code, string vendor_name,
            string vendor_phone, string vendor_address, string vendor_email, string security_guard, string shipping_plant, string vehicle_no, string driver_name, string remark, string prof_attn_to, string prof_street, string prof_city, string prof_country, string prof_postal_code, string prof_phone, string prof_email, string prof_coo,
            IFormFile file_attach, IFormFileCollection file_support, IFormFileCollection supporting_documents, string ship_mode, string charged_by, string courier_name, string courier_acc_no,
            string freight_charges, string incoterms, string invoice_payment, string gross_val, int? id_order = null)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;

            if (string.IsNullOrEmpty(sesa_id))
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }

            var db = new DatabaseAccessLayer();
            var email = new EmailController();
            var exp = new ExportController();

            string status_msg = "success";
            int? id_proforma = null;

            try
            {
                string supporting_docs_files = "";
                if ((type == "VENDOR" || type == "CSR") && (supporting_documents == null || supporting_documents.Count == 0))
                {
                    return Content("Supporting documents are required for VENDOR or CSR type.", "text/plain");
                }

                if (supporting_documents != null && supporting_documents.Count > 0)
                {
                    if (supporting_documents.Count < 1 || supporting_documents.Count > 3)
                    {
                        return Content("Please upload between 1 and 3 supporting documents.", "text/plain");
                    }

                    string upload_result = exp.UPLOAD_SUPPORTING_DOCUMENTS(supporting_documents);
                    string[] upload_parts = upload_result.Split(';');

                    if (upload_parts[0] == "OK")
                    {
                        supporting_docs_files = string.Join(";", upload_parts.Skip(1));
                    }
                    else
                    {
                        return Content($"File upload error: {upload_parts[1]}", "text/plain");
                    }
                }

                bool needProforma = false;
                if (!string.IsNullOrEmpty(category) && !string.IsNullOrEmpty(vendor_code))
                {
                    if (category == "TRANSFER TO VENDOR" || category == "REPAIR TO VENDOR" ||
                        (category == "FA SALES - WITHOUT DISMANTLING" && type == "VENDOR"))
                    {
                        string vendorBatam = "Batam";
                        string query = @"
                        SELECT vendor_batam 
                        FROM mst_vendor_gatepass 
                        WHERE vendor_code = @vendor_code";
                        using (SqlConnection conn = new SqlConnection(DbConnection()))
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@vendor_code", vendor_code);
                            conn.Open();
                            var result = cmd.ExecuteScalar();
                            vendorBatam = result?.ToString() ?? "Batam";
                        }
                        needProforma = vendorBatam == "Non Batam";
                    }
                }

                string file_attach_name = "";
                string file_support_names = "";

                if (needProforma)
                {
                    double grossValue = 0;

                    string cleanGrossVal = gross_val?.Trim() ?? "0";

                    cleanGrossVal = System.Text.RegularExpressions.Regex.Replace(cleanGrossVal, @"[^\d,.]", "");

                    if (cleanGrossVal.Contains(",") && cleanGrossVal.Contains("."))
                    {
                        int lastCommaIndex = cleanGrossVal.LastIndexOf(',');
                        int lastDotIndex = cleanGrossVal.LastIndexOf('.');

                        if (lastCommaIndex < lastDotIndex)
                        {
                            cleanGrossVal = cleanGrossVal.Replace(",", "");
                        }
                        else
                        {
                            cleanGrossVal = cleanGrossVal.Replace(".", "").Replace(",", ".");
                        }
                    }
                    else if (cleanGrossVal.Contains(","))
                    {
                        var parts = cleanGrossVal.Split(',');
                        if (parts.Length == 2 && parts[1].Length == 3 && parts[0].Length > 0)
                        {
                            cleanGrossVal = cleanGrossVal.Replace(",", "");
                        }
                        else
                        {
                            cleanGrossVal = cleanGrossVal.Replace(",", ".");
                        }
                    }
                    else if (cleanGrossVal.Contains("."))
                    {
                        var parts = cleanGrossVal.Split('.');
                        if (parts.Length == 2 && parts[1].Length == 3 && parts[0].Length > 0)
                        {
                            cleanGrossVal = cleanGrossVal.Replace(".", "");
                        }
                    }

                    bool parseSuccess = double.TryParse(cleanGrossVal, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out grossValue);

                    if (!parseSuccess)
                    {
                        return Content($"Invalid gross value format: '{gross_val}'. Please check the value.", "text/plain");
                    }

                    if (grossValue > 2500)
                    {
                        if (file_attach == null)
                        {
                            return Content("File attachment is required for gross value above 2500 USD.", "text/plain");
                        }
                    }

                    if (grossValue <= 2500 && file_attach != null)
                    {
                        return Content("File attachment is only allowed if gross value above 2500 USD.", "text/plain");
                    }

                    if (file_attach != null || (file_support != null && file_support.Count > 0))
                    {
                        string file_result = exp.UPLOAD_PROFORMA_FILES(file_attach, file_support);
                        string[] file_res = file_result.Split(';');

                        if (file_res[0] == "OK")
                        {
                            file_attach_name = file_res.Length > 1 ? file_res[1] : "";
                            if (file_res.Length > 2)
                            {
                                file_support_names = string.Join(";", file_res.Skip(2));
                            }
                        }
                        else
                        {
                            return Content($"File upload error: {file_res[1]}", "text/plain");
                        }
                    }

                    string courier_charges = charged_by;
                    string ret_proforma = db.SAVE_PROFORMA(prof_attn_to, prof_street, prof_city, prof_country, prof_postal_code,
                        prof_phone, prof_email, prof_coo, file_attach_name, file_support_names, ship_mode, courier_charges,
                        courier_name, courier_acc_no, freight_charges, incoterms, invoice_payment, sesa_id);

                    string[] proforma_res = ret_proforma.Split(';');
                    if (int.TryParse(proforma_res[0], out int idProformaVal))
                    {
                        id_proforma = idProformaVal;
                    }
                }

                string ret_order = db.SAVE_GATEPASS(create_by, create_date, category, return_date, new_pic, type, location, employee, csr, vendor_code, vendor_name, vendor_phone, vendor_address, vendor_email, security_guard, shipping_plant, vehicle_no, driver_name, remark, id_proforma, id_order, supporting_docs_files);
                string[] order = ret_order.Split(';');
                int id_gatepass = Convert.ToInt32(order[0]);

                string send_email = "success";
                if (category != "FA SALES - WITHOUT DISMANTLING" || (category == "FA SALES - WITHOUT DISMANTLING" && id_proforma != null) || (category == "FA SALES - WITHOUT DISMANTLING" && id_order != null))
                {
                    send_email = email.EMAIL_REQ_GATEPASS(id_gatepass);
                }

                return Content(status_msg + ";" + send_email + ";" + order[1], "text/plain");
            }
            catch (Exception ex)
            {
                return Content($"An error occurred: {ex.Message}", "text/plain");
            }
        }

        [HttpPost]
        public IActionResult UPDATE_HS_CODE_TEMP(string asset_no, string asset_subnumber, string hs_code)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            string result = db.UPDATE_HS_CODE_TEMP(asset_no, asset_subnumber, hs_code, sesa_id);
            return Content(result, "text/plain");
        }

        public IActionResult ADD_NEW_GATEPASS(string create_by, string create_date, string category, string new_pic, string type, string location, string employee, string csr, string return_date, string vendor_code, string vendor_name
            , string vendor_phone, string vendor_address, string vendor_email, string security_guard, string shipping_plant, string vehicle_no, string driver_name, string remark, IFormFileCollection supporting_documents, int? id_order = null)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;

            if (string.IsNullOrEmpty(sesa_id))
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }

            var db = new DatabaseAccessLayer();
            var email = new EmailController();
            var exp = new ExportController();

            string status_msg = "success";

            try
            {
                string supporting_docs_files = "";
                if ((type == "VENDOR" || type == "CSR") && (supporting_documents == null || supporting_documents.Count == 0))
                {
                    return Content("Supporting documents are required for VENDOR or CSR type.", "text/plain");
                }

                if (supporting_documents != null && supporting_documents.Count > 0)
                {
                    if (supporting_documents.Count < 1 || supporting_documents.Count > 3)
                    {
                        return Content("Please upload between 1 and 3 supporting documents.", "text/plain");
                    }

                    string upload_result = exp.UPLOAD_SUPPORTING_DOCUMENTS(supporting_documents);
                    string[] upload_parts = upload_result.Split(';');

                    if (upload_parts[0] == "OK")
                    {
                        supporting_docs_files = string.Join(";", upload_parts.Skip(1));
                    }
                    else
                    {
                        return Content($"File upload error: {upload_parts[1]}", "text/plain");
                    }
                }

                string ret_order = db.SAVE_GATEPASS(create_by, create_date, category, return_date, new_pic, type, location, employee, csr, vendor_code, vendor_name, vendor_phone, vendor_address, vendor_email, security_guard, shipping_plant, vehicle_no, driver_name, remark, null, id_order, supporting_docs_files);
                string[] order = ret_order.Split(';');

                string send_email = "success";
                if (category != "FA SALES - WITHOUT DISMANTLING")
                {
                    send_email = email.EMAIL_REQ_GATEPASS(Convert.ToInt32(order[0]));
                }

                return Content(status_msg + ";" + send_email + ";" + order[1], "text/plain");
            }
            catch (Exception ex)
            {
                return Content($"An error occurred: {ex.Message}", "text/plain");
            }
        }

        public IActionResult GET_GATEPASS_LIST()
        {
            try
            {
                var draw = Request.Form["draw"].FirstOrDefault();
                var start = Request.Form["start"].FirstOrDefault();
                var length = Request.Form["length"].FirstOrDefault();
                var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][data]"].FirstOrDefault();
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                var searchValue = Request.Form["search[value]"].FirstOrDefault();
                var status = Request.Form["status"].FirstOrDefault();

                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;
                int recordsTotal = 0;
                var created_by = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;

                var mstData = _context.v_gatepass.Select(gatepass => new
                {
                    gatepass.id_gatepass,
                    gatepass.gatepass_no,
                    gatepass.category,
                    gatepass.type,
                    gatepass.return_date,
                    gatepass.vendor_code,
                    gatepass.vendor_name,
                    gatepass.vendor_address,
                    gatepass.recipient_phone,
                    gatepass.recipient_email,
                    gatepass.vehicle_no,
                    gatepass.driver_name,
                    gatepass.security_guard,
                    gatepass.remark,
                    gatepass.created_by,
                    create_date = Convert.ToString(gatepass.create_date),
                    gatepass.status_gatepass,
                    gatepass.status_remark,
                    gatepass.image_after,
                    gatepass.id_proforma,
                    gatepass.proforma_no
                });

                if (!string.IsNullOrEmpty(created_by))
                {
                    mstData = mstData.Where(m => m.created_by == created_by);
                }

                if (!string.IsNullOrEmpty(status))
                {
                    if (status == "OPEN")
                    {
                        mstData = mstData.Where(m => m.status_gatepass != "COMPLETED" && !m.status_gatepass.Contains("REJECT") && !m.status_gatepass.Contains("CANCEL"));
                    }
                    else if (status == "CLOSE")
                    {
                        mstData = mstData.Where(m => m.status_gatepass == "COMPLETED" || m.status_gatepass.Contains("REJECT") || m.status_gatepass.Contains("CANCEL"));
                    }
                    else
                    {
                        mstData = mstData.Where(m => m.status_gatepass == status);
                    }
                }

                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    mstData = mstData.OrderBy($"{sortColumn} {sortColumnDirection}");
                }

                if (!string.IsNullOrEmpty(searchValue))
                {
                    mstData = mstData.Where(m => m.gatepass_no.Contains(searchValue)
                        || m.category.Contains(searchValue)
                        || m.vendor_code.Contains(searchValue)
                        || m.vendor_name.Contains(searchValue)
                        || m.vendor_address.Contains(searchValue));
                }

                for (int i = 0; i < 14; i++)
                {
                    var searchColVal = Request.Form[$"columns[{i}][search][value]"].FirstOrDefault();
                    var fieldName = Request.Form[$"columns[{i}][data]"].FirstOrDefault();

                    if (!string.IsNullOrEmpty(searchColVal) && !string.IsNullOrEmpty(fieldName))
                    {
                        mstData = mstData.Where($"{fieldName}.Contains(@0)", searchColVal);
                    }
                }

                recordsTotal = mstData.Count();
                var data = mstData.Skip(skip).Take(pageSize).ToList();
                var jsonData = new { draw, recordsFiltered = recordsTotal, recordsTotal, data };
                return Ok(jsonData);
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpGet]
        public IActionResult GET_IMG_GATEPASS(string id_gatepass)
        {
            var db = new DatabaseAccessLayer();
            List<GatePassModel> images = db.GET_IMG_GATEPASS(id_gatepass);
            return Json(images);
        }

        public IActionResult GET_GATEPASS_DETAIL(int id_gatepass)
        {
            try
            {
                var db = new DatabaseAccessLayer();

                var detailList = db.GET_GATEPASS_DETAIL(id_gatepass);
                var statusList = db.GATEPASS_STATUS(id_gatepass);
                var headerList = db.GET_GATEPASS_HEADER(id_gatepass);

                var financeFilesList = new List<ProformaFileModel>();
                var processedFileIds = new HashSet<int>();

                foreach (var detail in detailList)
                {
                    if (detail.id_file.HasValue && !processedFileIds.Contains(detail.id_file.Value))
                    {
                        financeFilesList.Add(new ProformaFileModel
                        {
                            id_file = detail.id_file.Value,
                            document_type = detail.document_type,
                            filename = detail.fin_filename,
                            created_by = detail.fin_created_by,
                            record_date = detail.fin_record_date ?? DateTime.Now
                        });
                        processedFileIds.Add(detail.id_file.Value);
                    }
                }

                var uniqueDetailList = detailList
                    .GroupBy(x => new { x.asset_no, x.asset_subnumber })
                    .Select(g => g.First())
                    .ToList();

                ViewBag.appList = statusList;
                ViewBag.disList = headerList;
                ViewBag.financeFiles = financeFilesList;
                ViewBag.totalApc = db.GET_TOTAL_CURRENT_APC(id_gatepass);

                return PartialView("_TableGatePassDetail", uniqueDetailList);
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        public IActionResult DELETE_TEMP_GATEPASS(int id_temp)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            string delTemp = db.DELETE_TEMP_GATEPASS(id_temp);
            return Content(delTemp, "text/plain");
        }

        public IActionResult GET_SHOW_DETAIL(string id_gatepass)
        {
            string query = @"
                    SELECT 
                           [vendor_name]
                          ,[vendor_address]
                          ,[recipient_phone]
                          ,[recipient_email]
                          ,[vehicle_no]
                          ,[driver_name]
                          ,[security_guard]
                      FROM [dbo].[tbl_gatepass_header]
                      WHERE 
                          id_gatepass = @id_gatepass OR @id_gatepass IS NULL";

            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@id_gatepass", (object)id_gatepass ?? DBNull.Value);

                    conn.Open();

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            var data = new
                            {
                                vendor_name = reader["vendor_name"]?.ToString() ?? "-",
                                vendor_address = reader["vendor_address"]?.ToString() ?? "-",
                                recipient_phone = reader["recipient_phone"]?.ToString() ?? "-",
                                recipient_email = reader["recipient_email"]?.ToString() ?? "-",
                                vehicle_no = reader["vehicle_no"]?.ToString() ?? "-",
                                driver_name = reader["driver_name"]?.ToString() ?? "-",
                                security_guard = reader["security_guard"]?.ToString() ?? "-"
                            };

                            return Json(data);
                        }
                    }
                }
            }

            return Json(null);
        }

        public IActionResult UPDATE_GATEPASS_REQ(IFormFile file_imgs, int id_gatepass, string id_detail, string asset_no, string asset_subnumber)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            string name = User.FindFirst("asset_management_name")?.Value;
            if (string.IsNullOrEmpty(sesa_id))
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }
            else
            {
                string status_msg = string.Empty;
                var db = new DatabaseAccessLayer();
                 var email = new EmailController();
                var exp = new ExportController();
                string dateSuffix = DateTime.Now.ToString("yyMMdd");
                string ret_order = asset_no + "_" + asset_subnumber + "_" + dateSuffix;
                string filename = ret_order + "_AFTER";
                string ret_file = exp.RETURN_GATEPASS(file_imgs, filename, id_gatepass);
                status_msg = db.UPDATE_AFTER_GATEPASS(id_gatepass, ret_file, id_detail);
                string send_email = email.EMAIL_REQ_GATEPASS_RETURN(id_gatepass);

                return Content(status_msg + ";" +  ret_order, "text/plain");
            }
        }

        [HttpGet]
        public IActionResult SearchAssetGatepass(string asset_no)
        {
            var db = new DatabaseAccessLayer();
            List<AssetListModel> assetList = db.GET_ASSET_GATEPASS(asset_no);
            return Json(new { items = assetList });
        }
        public JsonResult SearchVendor(string search, int page = 1)
        {
            var db = new DatabaseAccessLayer();
            var result = db.SearchVendorGatepass(search, page);
            return Json(result);
        }
        public IActionResult EditGatepass(int id_gatepass, string return_date)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            string name = User.FindFirst("asset_management_name")?.Value;
            if (string.IsNullOrEmpty(sesa_id))
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }
            else
            {
                var db = new DatabaseAccessLayer();
                string status_msg = db.EditGatepass(id_gatepass, return_date);
                return Content(status_msg, "text/plain");
            }
        }

        [HttpPost]
        public IActionResult CancelGatepass(int id_gatepass, string cancel_reason)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                string name = User.FindFirst("asset_management_name")?.Value;
                if (string.IsNullOrEmpty(sesa_id))
                {
                    return Content("Session Timeout, Please relogin!!", "text/plain");
                }
                else
                {
                    var db = new DatabaseAccessLayer();
                    var result = db.CancelGatepass(id_gatepass, sesa_id, cancel_reason);
                    return Json(result);
                }
            }
            catch (Exception ex)
            {
                return Json($"error;{ex.Message}");
            }
        }

        [HttpGet]
        public IActionResult GET_LIST_IMG(string id_gatepass, string gatepass_no)
        {
            var db = new DatabaseAccessLayer();
            List<GatePassModel> datalist = db.GET_LIST_IMG(id_gatepass);

            return PartialView("_TableGPListIMG", datalist);
        }

        public IActionResult GatePassNonAsset()
        {
            return this.CheckSession(() =>
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();

                db.DELETE_TEMP_NON_ASSET_SESA(sesa_id);

                List<UserDetailModel> userDetail = db.GetUserDetail(sesa_id);
                var securityList = db.GET_SECURITY_GUARD();
                var uomList = db.GET_UOM();

                ViewBag.securityList = securityList;
                ViewBag.uomList = uomList;

                return View(userDetail);
            });
        }

        [HttpPost]
        public IActionResult ADD_TEMP_GATEPASS_NONASSET(string po_or_asset_no, string description, decimal qty, string uom, decimal price_value, IFormFile asset_img, string hs_code)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            var exp = new ExportController();

            string filenameSave = null;
            if (asset_img != null && asset_img.Length > 0)
            {
                string dateSuffix = DateTime.Now.ToString("yyMMdd");
                string safeNo = (po_or_asset_no ?? "NONASSET").Replace("/", "_").Replace("\\", "_").Replace(" ", "_");
                string filename = $"{safeNo}_{dateSuffix}_BEFORE";
                string result = exp.CREATE_IMAGE_GP_NA(asset_img, filename);
                if (result.StartsWith("OK;"))
                {
                    filenameSave = result.Split(';')[1];
                }
                else
                {
                    return Content(result, "text/plain");
                }
            }

            decimal? hs = null;
            if (!string.IsNullOrWhiteSpace(hs_code))
            {
                if (decimal.TryParse(hs_code, out var v)) hs = v;
                else return Content("Invalid HS Code format", "text/plain");
            }

            var ret = db.ADD_TEMP_NON_ASSET(sesa_id, po_or_asset_no, description, qty, uom, price_value, filenameSave, hs);
            return Content(ret, "text/plain");
        }

        [HttpGet]
        public IActionResult GET_TEMP_GATEPASS_NONASSET()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            var list = db.GET_TEMP_NON_ASSET(sesa_id);
            return Json(list);
        }

        [HttpPost]
        public IActionResult DELETE_TEMP_GATEPASS_NONASSET(int id_temp)
        {
            var db = new DatabaseAccessLayer();
            var ret = db.DELETE_TEMP_NON_ASSET(id_temp);
            return Content(ret, "text/plain");
        }

        [HttpPost]
        public IActionResult ADD_NEW_GATEPASS_NONASSET(
            string category, string return_date,
            string vendor_code, string vendor_name, string vendor_address, string vendor_phone, string vendor_email, string vendor_batam,
            string security_guard, string shipping_plant, string vehicle_no, string driver_name, string remark,
            string prof_attn_to, string prof_street, string prof_city, string prof_country, string prof_postal_code, string prof_phone, string prof_email, string prof_coo,
            IFormFile file_attach, IFormFileCollection file_support,
            string ship_mode, string charged_by, string courier_name, string courier_acc_no,
            string freight_charges, string incoterms, string invoice_payment)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            if (string.IsNullOrEmpty(sesa_id))
                return Content("Session Timeout, Please relogin!!", "text/plain");

            var db = new DatabaseAccessLayer();
            var exp = new ExportController();
            var email = new EmailController();

            try
            {
                var items = db.GET_TEMP_NON_ASSET(sesa_id);
                if (items == null || items.Count == 0)
                    return Content("No items added.", "text/plain");

                double grossValue = Convert.ToDouble(items.Sum(x => x.qty * x.price_value));

                bool needProforma = string.Equals(vendor_batam, "Non Batam", StringComparison.OrdinalIgnoreCase);

                string file_attach_name = "";
                string file_support_names = "";

                if (needProforma)
                {
                    if (grossValue > 2500)
                    {
                        if (file_attach == null)
                            return Content("File attachment is required for gross value above 2500 USD.", "text/plain");
                    }
                    else
                    {
                        if (file_attach != null)
                            return Content("File attachment is only allowed if gross value above 2500 USD.", "text/plain");
                    }

                    if (file_attach != null || (file_support != null && file_support.Count > 0))
                    {
                        string file_result = exp.UPLOAD_PROFORMA_NON_ASSET_FILES(file_attach, file_support);
                        string[] file_res = file_result.Split(';');

                        if (file_res[0] == "OK")
                        {
                            file_attach_name = file_res.Length > 1 ? file_res[1] : "";
                            if (file_res.Length > 2)
                            {
                                file_support_names = string.Join(";", file_res.Skip(2));
                            }
                        }
                        else
                        {
                            return Content($"File upload error: {file_res[1]}", "text/plain");
                        }
                    }

                    string courier_charges = charged_by;
                    string ret_proforma = db.SAVE_PROFORMA_NON_ASSET(
                        prof_attn_to, prof_street, prof_city, prof_country, prof_postal_code,
                        prof_phone, prof_email, prof_coo,
                        file_attach_name, file_support_names,
                        ship_mode, courier_charges, courier_name, courier_acc_no,
                        freight_charges, incoterms, invoice_payment, sesa_id
                    );

                    string[] proforma_res = ret_proforma.Split(';');
                    if (!int.TryParse(proforma_res[0], out int idProformaVal))
                        return Content("Failed to save proforma", "text/plain");

                    string saveRes = db.SAVE_GATEPASS_NON_ASSET(
                        created_by: sesa_id,
                        category: category,
                        return_date: return_date,
                        vendor_code: vendor_code,
                        vendor_name: vendor_name,
                        vendor_address: vendor_address,
                        vendor_phone: vendor_phone,
                        vendor_email: vendor_email,
                        security_guard: security_guard,
                        shipping_plant: shipping_plant,
                        vehicle_no: vehicle_no,
                        driver_name: driver_name,
                        remark: remark,
                        id_proforma: idProformaVal
                    );

                    if (string.IsNullOrWhiteSpace(saveRes) || !saveRes.Contains(";"))
                        return Content(saveRes ?? "Failed to save gatepass", "text/plain");

                    var data = saveRes.Split(';');
                    int id_gpna = int.Parse(data[0]);

                    string send_email = email.EMAIL_REQ_NON_ASSET_GATEPASS(id_gpna);

                    return Content($"success;{send_email};{data[1]}", "text/plain");
                }
                else
                {
                    string saveRes = db.SAVE_GATEPASS_NON_ASSET(
                        created_by: sesa_id,
                        category: category,
                        return_date: return_date,
                        vendor_code: vendor_code,
                        vendor_name: vendor_name,
                        vendor_address: vendor_address,
                        vendor_phone: vendor_phone,
                        vendor_email: vendor_email,
                        security_guard: security_guard,
                        shipping_plant: shipping_plant,
                        vehicle_no: vehicle_no,
                        driver_name: driver_name,
                        remark: remark,
                        id_proforma: null
                    );

                    if (string.IsNullOrWhiteSpace(saveRes) || !saveRes.Contains(";"))
                        return Content(saveRes ?? "Failed to save gatepass", "text/plain");

                    var data = saveRes.Split(';');
                    int id_gpna = int.Parse(data[0]);

                    string send_email = email.EMAIL_REQ_NON_ASSET_GATEPASS(id_gpna);

                    return Content($"success;{send_email};{data[1]}", "text/plain");
                }
            }
            catch (Exception ex)
            {
                return Content("error;" + ex.Message, "text/plain");
            }
        }

        public IActionResult GatePassNonAssetList(string gatepass_no = null)
        {
            return this.CheckSession(() =>
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                List<UserDetailModel> userDetail = db.GetUserDetail(sesa_id);
                ViewBag.gatepass_no = gatepass_no;
                return View(userDetail);
            });
        }

        [HttpPost]
        public IActionResult GET_GATEPASS_NONASSET_LIST()
        {
            try
            {
                var draw = Request.Form["draw"].FirstOrDefault();
                var start = Request.Form["start"].FirstOrDefault();
                var length = Request.Form["length"].FirstOrDefault();
                var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][data]"].FirstOrDefault();
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                var searchValue = Request.Form["search[value]"].FirstOrDefault();
                var status = Request.Form["status"].FirstOrDefault();

                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;

                var created_by = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;

                var mstData = _context.v_gatepass_non_asset.Select(x => new
                {
                    x.id_gatepass,
                    x.gatepass_no,
                    x.category,
                    x.return_date,
                    x.vendor_code,
                    x.vendor_name,
                    x.vendor_address,
                    x.created_by,
                    create_date = Convert.ToString(x.create_date),
                    x.status_gatepass,
                    x.status_remark,
                    x.proforma_no
                });

                if (!string.IsNullOrEmpty(created_by))
                {
                    mstData = mstData.Where(m => m.created_by == created_by);
                }

                if (!string.IsNullOrEmpty(status))
                {
                    if (status == "OPEN")
                    {
                        mstData = mstData.Where(m => m.status_gatepass != "COMPLETED" && !m.status_gatepass.Contains("REJECT") && !m.status_gatepass.Contains("CANCEL"));
                    }
                    else if (status == "CLOSE")
                    {
                        mstData = mstData.Where(m => m.status_gatepass == "COMPLETED" || m.status_gatepass.Contains("REJECT") || m.status_gatepass.Contains("CANCEL"));
                    }
                    else
                    {
                        mstData = mstData.Where(m => m.status_gatepass == status);
                    }
                }

                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    mstData = mstData.OrderBy($"{sortColumn} {sortColumnDirection}");
                }

                if (!string.IsNullOrEmpty(searchValue))
                {
                    mstData = mstData.Where(m =>
                        m.gatepass_no.Contains(searchValue) ||
                        m.category.Contains(searchValue) ||
                        (m.vendor_code ?? "").Contains(searchValue) ||
                        (m.vendor_name ?? "").Contains(searchValue) ||
                        (m.vendor_address ?? "").Contains(searchValue));
                }

                for (int i = 0; i < 14; i++)
                {
                    var searchColVal = Request.Form[$"columns[{i}][search][value]"].FirstOrDefault();
                    var fieldName = Request.Form[$"columns[{i}][data]"].FirstOrDefault();

                    if (!string.IsNullOrEmpty(searchColVal) && !string.IsNullOrEmpty(fieldName))
                    {
                        mstData = mstData.Where($"{fieldName}.Contains(@0)", searchColVal);
                    }
                }

                var recordsTotal = mstData.Count();
                var data = mstData.Skip(skip).Take(pageSize).ToList();
                var jsonData = new { draw, recordsFiltered = recordsTotal, recordsTotal, data };
                return Ok(jsonData);
            }
            catch
            {
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpGet]
        public IActionResult GET_IMG_GATEPASS_NONASSET(int id_gatepass)
        {
            try
            {
                var db = new DatabaseAccessLayer { ConnectionString = DbConnection() };
                var rows = db.GET_IMG_NON_ASSET_GATEPASS(id_gatepass);
                return new JsonResult(rows);
            }
            catch (Exception ex)
            {
                return new JsonResult(new { error = ex.Message });
            }
        }

        [HttpPost]
        public IActionResult GET_SHOW_DETAIL_NONASSET(int id_gatepass)
        {
            try
            {
                var db = new DatabaseAccessLayer();
                var data = db.GET_SHOW_NON_ASSET_DETAIL(id_gatepass);
                return Json(data);
            }
            catch (Exception ex)
            {
                return Json(new { error = ex.Message });
            }
        }

        [HttpGet]
        public IActionResult GET_GATEPASS_NONASSET_DETAIL(int id_gatepass)
        {
            try
            {
                var db = new DatabaseAccessLayer();

                var detailList = db.GET_GATEPASS_NON_ASSET_DETAIL(id_gatepass);
                var statusList = db.GET_GATEPASS_NON_ASSET_STATUS(id_gatepass);
                var headerList = db.GET_GATEPASS_NON_ASSET_HEADER(id_gatepass);

                var financeFilesList = new List<ProformaFileModel>();
                var processedFileIds = new HashSet<int>();

                foreach (var detail in detailList)
                {
                    if (detail.id_file.HasValue && !processedFileIds.Contains(detail.id_file.Value))
                    {
                        financeFilesList.Add(new ProformaFileModel
                        {
                            id_file = detail.id_file.Value,
                            document_type = detail.document_type,
                            filename = detail.fin_filename,
                            created_by = detail.fin_created_by,
                            record_date = detail.fin_record_date ?? DateTime.Now
                        });
                        processedFileIds.Add(detail.id_file.Value);
                    }
                }

                var uniqueDetailList = detailList
                    .GroupBy(x => new { x.po_or_asset_no, x.description })
                    .Select(g => g.First())
                    .ToList();

                ViewBag.appList = statusList;
                ViewBag.disList = headerList;
                ViewBag.financeFiles = financeFilesList;
                ViewBag.totalValue = db.GET_TOTAL_NON_ASSET_VALUE(id_gatepass);

                return PartialView("_TableGatePassNonAssetDetail", uniqueDetailList);
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        public IActionResult EditGatepassNonAsset(int id_gatepass, string return_date)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            string name = User.FindFirst("asset_management_name")?.Value;
            if (string.IsNullOrEmpty(sesa_id))
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }
            else
            {
                var db = new DatabaseAccessLayer();
                string status_msg = db.EDIT_GATEPASS_NON_ASSET(id_gatepass, return_date);
                return Content(status_msg, "text/plain");
            }
        }

        [HttpPost]
        public IActionResult CancelGatepassNonAsset(int id_gatepass, string cancel_reason)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                string name = User.FindFirst("asset_management_name")?.Value;
                if (string.IsNullOrEmpty(sesa_id))
                {
                    return Content("Session Timeout, Please relogin!!", "text/plain");
                }
                else
                {
                    var db = new DatabaseAccessLayer();
                    var result = db.CANCEL_GATEPASS_NON_ASSET(id_gatepass, sesa_id, cancel_reason);
                    return Json(result);
                }
            }
            catch (Exception ex)
            {
                return Json($"error;{ex.Message}");
            }
        }

        [HttpGet]
        public IActionResult GET_LIST_IMG_NON_ASSET(string id_gatepass, string gatepass_no)
        {
            var db = new DatabaseAccessLayer();
            List<GatePassNonAssetModel> datalist = db.GET_LIST_NON_ASSET_IMG(id_gatepass);

            return PartialView("_TableGPListIMGNonAsset", datalist);
        }

        [HttpPost]
        public IActionResult UPDATE_GATEPASS_NON_ASSET_REQ(IFormFile file_imgs, int id_gatepass, string id_detail, string po_or_asset_no)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            string name = User.FindFirst("asset_management_name")?.Value;
            if (string.IsNullOrEmpty(sesa_id))
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }
            else
            {
                string status_msg = string.Empty;
                var db = new DatabaseAccessLayer();
                var email = new EmailController();
                var exp = new ExportController();
                string dateSuffix = DateTime.Now.ToString("yyMMdd");
                string safeNo = (po_or_asset_no ?? "NONASSET").Replace("/", "_").Replace("\\", "_").Replace(" ", "_");
                string ret_order = safeNo + "_" + dateSuffix;

                string filename = ret_order + "_AFTER";
                string ret_file = exp.RETURN_GATEPASS_NON_ASSET(file_imgs, filename, id_gatepass);

                status_msg = db.UPDATE_AFTER_NON_ASSET_GATEPASS(id_gatepass, ret_file, id_detail);

                string send_email = email.EMAIL_REQ_NON_ASSET_GATEPASS_RETURN(id_gatepass);

                return Content(status_msg + ";" + ret_order + ";" + send_email, "text/plain");
            }
        }

    }
}
