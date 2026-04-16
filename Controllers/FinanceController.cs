using ClosedXML.Excel;
using iText.Html2pdf;
using iText.IO.Source;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using OfficeOpenXml;
using Asset_Management.Function;
using Asset_Management.Models;
using Asset_Management.Service;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq.Dynamic.Core;
using System.Linq.Expressions;
using System.Security.Claims;
using System.Text;
using static iText.IO.Codec.TiffWriter;


namespace Asset_Management.Controllers
{
    [Authorize(Policy = "RequireFinanceApprover")]
    public class FinanceController : Controller
    {
        private readonly ApplicationDbContext _context;

        private readonly ImportExportFactory _importexportFactory;
        private readonly ILogger<FinanceController> _logger;

        private string DbConnection()
        {
            var dbAccess = new DatabaseAccessLayer();
            string dbString = dbAccess.ConnectionString;
            return dbString;
        }

        public FinanceController(ImportExportFactory importexportFactory, ILogger<FinanceController> logger, ApplicationDbContext context)
        {
            this._context = context;
            _logger = logger;
            _importexportFactory = importexportFactory;
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

        public IActionResult GetAssetListBadge()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            int totalNoTaggingWaitingValidation = db.GetTotalNoTaggingWaitingValidation();
            int total = totalNoTaggingWaitingValidation;
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
                             where GatepassList.status_gatepass == "Waiting Confirmation by FIN SS"
                             || GatepassList.status_gatepass == "Waiting Proforma Documents"
                             || GatepassList.status_gatepass == "Waiting PEB Documents"
                             select GatepassList).Count();
                return Ok(new { count = count });
            }
            catch (Exception)
            {
                return Ok(new { count = 0 });
            }
        }

        public IActionResult GetCreateInvoiceBadge()
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
                var count = (from CreateInvoice in _context.v_gatepass
                             where CreateInvoice.status_gatepass == "Waiting Invoice from Finance"
                             select CreateInvoice).Count();
                return Ok(new { count = count });
            }
            catch (Exception)
            {
                return Ok(new { count = 0 });
            }
        }

        public IActionResult GetinvoiceListBadge()
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
                var count = (from CreateInvoice in _context.v_gatepass
                             where CreateInvoice.status_gatepass == "Waiting Payment"
                             select CreateInvoice).Count();
                return Ok(new { count = count });
            }
            catch (Exception)
            {
                return Ok(new { count = 0 });
            }
        }

        public IActionResult GetInvoiceApprovalBadge()
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
                var count = (from InvoiceApproval in _context.v_gatepass
                             where InvoiceApproval.approval_invoice_sesa == sesa_id
                             && InvoiceApproval.approval_invoice_status == "OPEN"
                             && InvoiceApproval.status_gatepass == "Waiting Accounting Manager Approval"
                             select InvoiceApproval).Count();
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
                             where GatepassNonAssetList.status_gatepass == "Waiting Confirmation by FIN SS"
                             || GatepassNonAssetList.status_gatepass == "Waiting Proforma Documents"
                             || GatepassNonAssetList.status_gatepass == "Waiting PEB Documents"
                             select GatepassNonAssetList).Count();
                return Ok(new { count = count });
            }
            catch (Exception)
            {
                return Ok(new { count = 0 });
            }
        }

        public IActionResult Gatepass(string gatepass_no = null)
        {
            return this.CheckSession(() =>
            {
                var db = new DatabaseAccessLayer();
                ViewBag.sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                ViewBag.role = User.FindFirst("asset_management_role")?.Value;
                List<string> catList = db.GET_SECURITY_GUARD();
                ViewBag.catList = catList;
                ViewBag.gatepass_no = gatepass_no;
                return View();
            });
        }
        public IActionResult AssetList()
        {
            return this.CheckSession(() =>
            {
                var db = new DatabaseAccessLayer();
                int totalNoTagging = db.GetTotalNoTagging();
                int totalNoTaggingWaitingValidation = db.GetTotalNoTaggingWaitingValidation();
                ViewBag.totalNoTagging = totalNoTagging;
                ViewBag.totalNoTaggingWaitingValidation = totalNoTaggingWaitingValidation;
                return View();
            });
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
                                       AssetList.name_owner,
                                       AssetList.tagging_status,
                                       AssetList.file_tag,
                                       AssetList.vendor_name,
                                       AssetList.asset_location,
                                       AssetList.asset_location_address,
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
                for(int i = 0; i < 31; i++)
                {
                    var searchColVal = Request.Form["columns["+i.ToString()+"][search][value]"];
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
        
        [HttpPost]
        public IActionResult ValidateCount(int id_count, string existence, string good_condition, string still_in_operation, string tagging, string applicable_of_tagging, string correct_naming, string correct_location, string retagging)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();

            string status_msg = db.ValidateCount(id_count, existence, good_condition, still_in_operation, tagging, applicable_of_tagging, correct_naming, correct_location, retagging, sesa_id);

            return Content(status_msg, "text/plain");
        }
        public IActionResult GetAssetTaggingValidation()
        {
            var db = new DatabaseAccessLayer();
            List<AssetListModel> dataList = db.GetAssetTaggingValidation();
            return PartialView("_TableAssetTaggingValidation", dataList);
        }
        public IActionResult GET_ASSET_NO_TAGGING()
        {
            var db = new DatabaseAccessLayer();
            List<AssetListModel> dataList = db.GET_ASSET_NO_TAGGING();
            return PartialView("_TableAssetNoTagging", dataList);
        }
        [HttpPost]
        public IActionResult ConfirmTagging(string asset_no, string asset_subnumber, string confirm_type)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            string updateAsset = db.ConfirmTagging(asset_no, asset_subnumber, confirm_type);
            return Content(updateAsset, "text/plain");
        }
        public IActionResult UploadAsset(IFormFile FormFile)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            if (sesa_id == "")
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }
            else
            {
                DateTime now = DateTime.Now;
                string id_login = now.ToString("yyMMddHHmmssfff");
                _importexportFactory.ImportAsset(FormFile, id_login, sesa_id);

                return Content("Upload Success!!", "text/plain");

            }
        }
        
        public IActionResult GET_GATEPASS_HEADER()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            string role = User.FindFirst("asset_management_role")?.Value;
            try
            {
                var status = Request.Form["status"].FirstOrDefault();
                var draw = Request.Form["draw"].FirstOrDefault();
                var start = Request.Form["start"].FirstOrDefault();
                var length = Request.Form["length"].FirstOrDefault();
                var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][data]"].FirstOrDefault();
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                var searchValue = Request.Form["search[value]"].FirstOrDefault();
                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;

                var mstData = from gatepass in _context.v_gatepass
                              where ((status == "OPEN"
                                      && gatepass.status_gatepass != "COMPLETED"
                                      && !gatepass.status_gatepass.Contains("REJECT")
                                      && !gatepass.status_gatepass.Contains("Cancel"))
                                     ||
                                     (status == "CLOSE"
                                     && (gatepass.status_gatepass == "COMPLETED"
                                     || gatepass.status_gatepass.Contains("REJECT")
                                     || gatepass.status_gatepass.Contains("Cancel"))))
                              select new
                              {
                                  gatepass.id_gatepass,
                                  gatepass.gatepass_no,
                                  gatepass.category,
                                  gatepass.type,
                                  gatepass.created_by,
                                  create_date = Convert.ToString(gatepass.create_date),
                                  return_date = Convert.ToString(gatepass.return_date),
                                  gatepass.vendor_name,
                                  gatepass.status_gatepass,
                                  gatepass.status_remark,
                                  gatepass.approval_hod,
                                  gatepass.approval_hod_status,
                                  gatepass.approval_fbp,
                                  gatepass.approval_fbp_status,
                                  gatepass.approval_ph,
                                  gatepass.approval_ph_status,
                                  gatepass.security_name,
                                  gatepass.approval_security_status,
                                  gatepass.approval_finance,
                                  gatepass.approval_finance_status,
                                  gatepass.requestor_name,
                                  gatepass.image_after,
                                  gatepass.id_proforma,
                                  gatepass.proforma_fin_status,
                                  gatepass.shipping_status,
                                  gatepass.proforma_no
                              };

                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    mstData = mstData.OrderBy($"{sortColumn} {sortColumnDirection}");
                }

                if (!string.IsNullOrEmpty(searchValue))
                {
                    mstData = mstData.Where(m =>
                        m.gatepass_no.Contains(searchValue) ||
                        m.proforma_no.Contains(searchValue) ||
                        m.status_gatepass.Contains(searchValue) ||
                        m.category.Contains(searchValue) ||
                        m.type.Contains(searchValue) ||
                        m.created_by.Contains(searchValue) ||
                        m.create_date.Contains(searchValue) ||
                        m.return_date.Contains(searchValue) ||
                        m.vendor_name.Contains(searchValue) ||
                        m.status_remark.Contains(searchValue) ||
                        m.approval_hod.Contains(searchValue) ||
                        m.approval_hod_status.Contains(searchValue) ||
                        m.approval_fbp.Contains(searchValue) ||
                        m.approval_fbp_status.Contains(searchValue) ||
                        m.approval_ph.Contains(searchValue) ||
                        m.approval_ph_status.Contains(searchValue) ||
                        m.security_name.Contains(searchValue) ||
                        m.approval_security_status.Contains(searchValue) ||
                        m.approval_finance.Contains(searchValue) ||
                        m.approval_finance_status.Contains(searchValue) ||
                        m.requestor_name.Contains(searchValue) ||
                        m.image_after.Contains(searchValue) ||
                        m.proforma_fin_status.Contains(searchValue) ||
                        m.shipping_status.Contains(searchValue));
                }

                for (int i = 0; i < 10; i++)
                {
                    var searchColVal = Request.Form["columns[" + i.ToString() + "][search][value]"];
                    var fieldName = Request.Form["columns[" + i.ToString() + "][data]"].FirstOrDefault();
                    if (!string.IsNullOrEmpty(searchColVal))
                    {
                        if (fieldName == "gatepass_no")
                        {
                            mstData = mstData.Where(m => m.gatepass_no.Contains(searchColVal));
                        }
                        else if (fieldName == "proforma_no")
                        {
                            mstData = mstData.Where(m => m.proforma_no.Contains(searchColVal));
                        }
                        else if (fieldName == "category")
                        {
                            mstData = mstData.Where(m => m.category.Contains(searchColVal));
                        }
                        else if (fieldName == "requestor_name")
                        {
                            mstData = mstData.Where(m => m.requestor_name.Contains(searchColVal));
                        }
                        else if (fieldName == "create_date")
                        {
                            mstData = mstData.Where(m => m.create_date.Contains(searchColVal));
                        }
                        else if (fieldName == "return_date")
                        {
                            mstData = mstData.Where(m => m.return_date.Contains(searchColVal));
                        }
                        else if (fieldName == "vendor_name")
                        {
                            mstData = mstData.Where(m => m.vendor_name.Contains(searchColVal));
                        }
                        else if (fieldName == "status_gatepass")
                        {
                            mstData = mstData.Where(m => m.status_gatepass.Contains(searchColVal));
                        }
                        else if (fieldName == "status_remark")
                        {
                            mstData = mstData.Where(m => m.status_remark.Contains(searchColVal));
                        }
                    }
                }

                int recordsTotal = mstData.Count();
                var data = mstData.Skip(skip).Take(pageSize).ToList();
                var jsonData = new { draw, recordsFiltered = recordsTotal, recordsTotal, data };
                return Ok(jsonData);
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Internal server error. Please try again later.");
            }
        }

        public IActionResult UPDATE_FINANCE_GATEPASS(int id_gatepass, string approval_status, string approval_remark)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            string name = User.FindFirst("asset_management_name")?.Value;
            string level = User.FindFirst("asset_management_level")?.Value;
            if (sesa_id == "")
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }
            else
            {
                var db = new DatabaseAccessLayer();
                var email = new EmailController();
                string ret_order = db.UPDATE_FINANCE_GATEPASS(id_gatepass, approval_status, approval_remark, sesa_id);
                string send_email = "";
                if (approval_status == "APPROVED")
                {
                    send_email = email.GATEPASS_COMPLETED(id_gatepass, name, level);
                }
                else if (approval_status == "REJECTED")
                {
                    send_email = email.GATEPASS_REJECTED(id_gatepass);
                }
                return Content("success;" + send_email, "text/plain");
            }
        }

        [HttpPost]
        public IActionResult UPLOAD_PROFORMA_FIN_FILES(IFormFile file, int id_gatepass, string file_type)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            if (sesa_id == "")
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }

            try
            {
                if (file != null && file.Length > 0)
                {
                    var db = new DatabaseAccessLayer();

                    var existingFiles = db.GET_PROFORMA_FIN_FILES(id_gatepass);
                    var duplicateFile = existingFiles.FirstOrDefault(f => f.document_type == file_type);

                    if (duplicateFile != null)
                    {
                        return Content($"A file for document type '{file_type}' already exists. Please delete the existing file first or choose a different document type.", "text/plain");
                    }

                    var exp = new ExportController();
                    string file_result = exp.UPLOAD_PROFORMA_FIN_FILES(file, file_type);
                    string[] file_res = file_result.Split(';');

                    if (file_res[0] == "OK")
                    {
                        string finalFileName = file_res[1];
                        string result = db.UPLOAD_PROFORMA_FIN_FILES(id_gatepass, file_type, finalFileName, sesa_id);

                        return Content("success", "text/plain");
                    }
                    else
                    {
                        return Content($"File upload error: {file_res[1]}", "text/plain");
                    }
                }
                return Content("No file selected", "text/plain");
            }
            catch (Exception ex)
            {
                return Content($"Error: {ex.Message}", "text/plain");
            }
        }

        [HttpPost]
        public IActionResult DELETE_PROFORMA_FIN_FILE(int id_file, int id_gatepass)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            if (sesa_id == "")
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }

            try
            {
                var db = new DatabaseAccessLayer();
                string result = db.DELETE_PROFORMA_FIN_FILE(id_file, id_gatepass);
                return Content(result, "text/plain");
            }
            catch (Exception ex)
            {
                return Content($"Error: {ex.Message}", "text/plain");
            }
        }

        [HttpPost]
        public IActionResult COMPLETE_PROFORMA_FIN_UPLOAD(int id_gatepass)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;

            if (sesa_id == "")
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }

            try
            {
                var db = new DatabaseAccessLayer();

                string result = db.UPDATE_PROFORMA_FIN_STATUS(id_gatepass, sesa_id);

                return Content("success", "text/plain");
            }
            catch (Exception ex)
            {
                return Content($"Error: {ex.Message}", "text/plain");
            }
        }

        [HttpGet]
        public IActionResult GET_PROFORMA_FIN_FILES(int id_gatepass)
        {
            var db = new DatabaseAccessLayer();
            var files = db.GET_PROFORMA_FIN_FILES(id_gatepass);
            return Json(files);
        }

        [HttpGet]
        public IActionResult GET_PEB_FILE(int id_gatepass)
        {
            var db = new DatabaseAccessLayer();
            var file = db.GET_PEB_FILE(id_gatepass);
            return Json(file);
        }

        [HttpPost]
        public IActionResult UPLOAD_PEB_FILE(IFormFileCollection files, int id_gatepass)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            if (sesa_id == "")
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }

            try
            {
                if (files != null && files.Count > 0)
                {
                    var exp = new ExportController();
                    string file_result = exp.UPLOAD_PEB_FILE(files);
                    string[] file_res = file_result.Split(';');

                    if (file_res[0] == "OK")
                    {
                        string combinedFilenames = file_res[1];
                        var db = new DatabaseAccessLayer();
                        string result = db.UPLOAD_PEB_FILE(id_gatepass, combinedFilenames, sesa_id);
                        return Content("success", "text/plain");
                    }
                    else
                    {
                        return Content($"File upload error: {file_res[1]}", "text/plain");
                    }
                }
                return Content("No files selected", "text/plain");
            }
            catch (Exception ex)
            {
                return Content($"Error: {ex.Message}", "text/plain");
            }
        }

        [HttpPost]
        public IActionResult DELETE_PEB_FILE(int id_gatepass, string filename = null)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            if (sesa_id == "")
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }

            try
            {
                var db = new DatabaseAccessLayer();
                string result = db.DELETE_PEB_FILE(id_gatepass, filename);
                return Content(result, "text/plain");
            }
            catch (Exception ex)
            {
                return Content($"Error: {ex.Message}", "text/plain");
            }
        }

        [HttpPost]
        public IActionResult COMPLETE_PEB_UPLOAD(int id_gatepass)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            if (sesa_id == "")
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }

            try
            {
                var db = new DatabaseAccessLayer();

                string result = db.COMPLETE_PEB_UPLOAD(id_gatepass, sesa_id);

                return Content("success", "text/plain");
            }
            catch (Exception ex)
            {
                return Content($"Error: {ex.Message}", "text/plain");
            }
        }

        public IActionResult MstData_VendorGatepass()
        {
            string level = User.FindFirst("asset_management_level")?.Value;
            if (level == "finance")
            {
                return View();
            }
            else
            {
                return RedirectToAction("SignOut", "Login");
            }
        }

        public IActionResult GetVendorGatepassList()
        {
            try
            {
                var draw = Request.Form["draw"].FirstOrDefault();
                var start = Request.Form["start"].FirstOrDefault();
                var length = Request.Form["length"].FirstOrDefault();
                var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][data]"].FirstOrDefault();
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                var searchValue = Request.Form["search[value]"].FirstOrDefault();

                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;
                int recordsTotal = 0;

                var mstData = (from VendorGatepass in _context.mst_vendor_gatepass
                               select new
                               {
                                   VendorGatepass.id_vendor,
                                   VendorGatepass.vendor_code,
                                   VendorGatepass.vendor_name,
                                   VendorGatepass.vendor_address,
                                   VendorGatepass.vendor_postal_code,
                                   VendorGatepass.vendor_location,
                                   VendorGatepass.vendor_batam,
                                   VendorGatepass.vendor_phone,
                                   VendorGatepass.vendor_email,
                                   record_date = VendorGatepass.record_date != null ?
                                                VendorGatepass.record_date.Value.ToString("yyyy-MM-dd") : ""
                               });

                if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDirection)))
                {
                    mstData = mstData.OrderBy(sortColumn + " " + sortColumnDirection);
                }

                if (!string.IsNullOrEmpty(searchValue))
                {
                    mstData = mstData.Where(m => m.vendor_code.Contains(searchValue)
                                                || m.vendor_name.Contains(searchValue)
                                                || m.vendor_address.Contains(searchValue)
                                                || m.vendor_location.Contains(searchValue));
                }

                for (int i = 0; i < 8; i++)
                {
                    var searchColVal = Request.Form["columns[" + i.ToString() + "][search][value]"];
                    var fieldName = Request.Form["columns[" + i.ToString() + "][data]"].FirstOrDefault();
                    if (!string.IsNullOrEmpty(searchColVal))
                    {
                        if (fieldName == "vendor_code")
                        {
                            mstData = mstData.Where(m => m.vendor_code.Contains(searchColVal));
                        }
                        else if (fieldName == "vendor_name")
                        {
                            mstData = mstData.Where(m => m.vendor_name.Contains(searchColVal));
                        }
                        else if (fieldName == "vendor_address")
                        {
                            mstData = mstData.Where(m => m.vendor_address.Contains(searchColVal));
                        }
                        else if (fieldName == "vendor_postal_code")
                        {
                            mstData = mstData.Where(m => m.vendor_postal_code.Contains(searchColVal));
                        }
                        else if (fieldName == "vendor_location")
                        {
                            mstData = mstData.Where(m => m.vendor_location.Contains(searchColVal));
                        }
                        else if (fieldName == "vendor_batam")
                        {
                            mstData = mstData.Where(m => m.vendor_batam.Contains(searchColVal));
                        }
                        else if (fieldName == "vendor_phone")
                        {
                            mstData = mstData.Where(m => m.vendor_phone.Contains(searchColVal));
                        }
                        else if (fieldName == "vendor_email")
                        {
                            mstData = mstData.Where(m => m.vendor_email.Contains(searchColVal));
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

        public JsonResult AddVendorGatepass(string vendor_code, string vendor_name, string vendor_address, string vendor_postal_code, string vendor_location, string vendor_batam, string vendor_phone, string vendor_email)
        {
            try
            {
                Models.VendorGatepassModel mstvendor = new Models.VendorGatepassModel();
                mstvendor.vendor_code = vendor_code;
                mstvendor.vendor_name = vendor_name;
                mstvendor.vendor_address = vendor_address;
                mstvendor.vendor_postal_code = vendor_postal_code;
                mstvendor.vendor_location = vendor_location;
                mstvendor.vendor_batam = vendor_batam;
                mstvendor.vendor_phone = vendor_phone;
                mstvendor.vendor_email = vendor_email;

                string query = "ADD_VENDOR_GATEPASS";
                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@vendor_code", vendor_code);
                        cmd.Parameters.AddWithValue("@vendor_name", vendor_name);
                        cmd.Parameters.AddWithValue("@vendor_address", vendor_address);
                        cmd.Parameters.AddWithValue("@vendor_postal_code", vendor_postal_code);
                        cmd.Parameters.AddWithValue("@vendor_location", vendor_location);
                        cmd.Parameters.AddWithValue("@vendor_batam", vendor_batam);
                        cmd.Parameters.AddWithValue("@vendor_phone", vendor_phone);
                        cmd.Parameters.AddWithValue("@vendor_email", vendor_email);

                        conn.Open();
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                string status = reader["status"].ToString();
                                string message = reader["message"].ToString();

                                return Json(new { status = status, message = message });
                            }
                        }
                        conn.Close();
                    }
                }
                return Json(new { status = "error", message = "Unknown error occurred" });
            }
            catch (Exception ex)
            {
                return Json(new { status = "error", message = ex.Message });
            }
        }

        public JsonResult EditVendorGatepass(int id, string vendor_code, string vendor_name, string vendor_address, string vendor_postal_code, string vendor_location, string vendor_batam, string vendor_phone, string vendor_email)
        {
            try
            {
                string originalVendorCode = "";
                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("SELECT vendor_code FROM mst_vendor_gatepass WHERE id_vendor = @id_vendor", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_vendor", id);
                        var result = cmd.ExecuteScalar();
                        originalVendorCode = result?.ToString() ?? "";
                    }
                    conn.Close();
                }

                string query = "EDIT_VENDOR_GATEPASS";
                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@id_vendor", id);
                        cmd.Parameters.AddWithValue("@vendor_code", originalVendorCode);
                        cmd.Parameters.AddWithValue("@vendor_name", vendor_name);
                        cmd.Parameters.AddWithValue("@vendor_address", vendor_address);
                        cmd.Parameters.AddWithValue("@vendor_postal_code", vendor_postal_code);
                        cmd.Parameters.AddWithValue("@vendor_location", vendor_location);
                        cmd.Parameters.AddWithValue("@vendor_batam", vendor_batam);
                        cmd.Parameters.AddWithValue("@vendor_phone", vendor_phone);
                        cmd.Parameters.AddWithValue("@vendor_email", vendor_email);

                        conn.Open();
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                string status = reader["status"].ToString();
                                string message = reader["message"].ToString();

                                return Json(new { status = status, message = message });
                            }
                        }
                        conn.Close();
                    }
                }
                return Json(new { status = "error", message = "Unknown error occurred" });
            }
            catch (Exception ex)
            {
                return Json(new { status = "error", message = ex.Message });
            }
        }

        public JsonResult DeleteVendorGatepass(int id)
        {
            int Execute;
            string storedProcedure = "DELETE_VENDOR_GATEPASS";
            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                using (SqlCommand cmd = new SqlCommand(storedProcedure, conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id_vendor", id);
                    conn.Open();
                    Execute = cmd.ExecuteNonQuery();
                    conn.Close();
                }
            }

            return Json(Execute);
        }

        [HttpGet]
        public IActionResult GetPlant(string plant)
        {
            List<ProductFamilyModel> data = new List<ProductFamilyModel>();
            string query = "SELECT DISTINCT plant FROM mst_product_family WHERE plant LIKE '%" + plant + "%' ORDER BY plant ASC";
            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                using (SqlCommand cmd = new SqlCommand(query))
                {
                    cmd.Connection = conn;
                    conn.Open();

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var data_list = new ProductFamilyModel();
                            data_list.Text = reader["plant"].ToString();
                            data_list.id = reader["plant"].ToString();

                            data.Add(data_list);
                        }
                    }
                    conn.Close();
                }
            }

            return Json(new { items = data });
        }

        public IActionResult GatepassNonAsset(string gatepass_no = null)
        {
            string sesa_id = HttpContext.Session.GetString("sesa_id") ?? "";
            string role = HttpContext.Session.GetString("role") ?? "";

            ViewBag.sesa_id = sesa_id.ToUpper();
            ViewBag.role = role.ToUpper();
            ViewBag.gatepass_no = gatepass_no;

            return View();
        }

        [HttpPost]
        public IActionResult GET_GATEPASS_NONASSET_HEADER()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            string role = User.FindFirst("asset_management_role")?.Value;
            try
            {
                var status = Request.Form["status"].FirstOrDefault();
                var draw = Request.Form["draw"].FirstOrDefault();
                var start = Request.Form["start"].FirstOrDefault();
                var length = Request.Form["length"].FirstOrDefault();
                var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][data]"].FirstOrDefault();
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                var searchValue = Request.Form["search[value]"].FirstOrDefault();
                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;

                var mstData = from gatepass in _context.v_gatepass_non_asset
                              where ((status == "OPEN"
                                      && gatepass.status_gatepass != "COMPLETED"
                                      && !gatepass.status_gatepass.Contains("REJECT")
                                      && !gatepass.status_gatepass.Contains("Cancel"))
                                     ||
                                     (status == "CLOSE"
                                     && (gatepass.status_gatepass == "COMPLETED"
                                     || gatepass.status_gatepass.Contains("REJECT")
                                     || gatepass.status_gatepass.Contains("Cancel"))))
                              select new
                              {
                                  gatepass.id_gatepass,
                                  gatepass.gatepass_no,
                                  gatepass.category,
                                  gatepass.created_by,
                                  create_date = Convert.ToString(gatepass.create_date),
                                  return_date = Convert.ToString(gatepass.return_date),
                                  gatepass.vendor_name,
                                  gatepass.status_gatepass,
                                  gatepass.status_remark,
                                  gatepass.approval_hod,
                                  gatepass.approval_hod_status,
                                  gatepass.approval_fbp,
                                  gatepass.approval_fbp_status,
                                  gatepass.approval_ph,
                                  gatepass.approval_ph_status,
                                  gatepass.security_name,
                                  gatepass.approval_security_status,
                                  gatepass.approval_finance,
                                  gatepass.approval_finance_status,
                                  gatepass.requestor_name,
                                  gatepass.image_after,
                                  gatepass.id_proforma,
                                  gatepass.proforma_fin_status,
                                  gatepass.shipping_status,
                                  gatepass.proforma_no
                              };

                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    mstData = mstData.OrderBy($"{sortColumn} {sortColumnDirection}");
                }

                if (!string.IsNullOrEmpty(searchValue))
                {
                    mstData = mstData.Where(m => m.gatepass_no.Contains(searchValue)
                                                || m.category.Contains(searchValue)
                                                || m.create_date.Contains(searchValue)
                                                || m.return_date.Contains(searchValue)
                                                || m.vendor_name.Contains(searchValue)
                                                || m.status_gatepass.Contains(searchValue)
                                                || m.status_remark.Contains(searchValue)
                                                || m.requestor_name.Contains(searchValue)
                                                || m.proforma_no.Contains(searchValue));
                }

                for (int i = 0; i < 10; i++)
                {
                    var searchColVal = Request.Form["columns[" + i.ToString() + "][search][value]"];
                    var fieldName = Request.Form["columns[" + i.ToString() + "][data]"].FirstOrDefault();
                    if (!string.IsNullOrEmpty(searchColVal))
                    {
                        if (fieldName == "gatepass_no")
                        {
                            mstData = mstData.Where(m => m.gatepass_no.Contains(searchColVal));
                        }
                        else if (fieldName == "proforma_no")
                        {
                            mstData = mstData.Where(m => m.proforma_no.Contains(searchColVal));
                        }
                        else if (fieldName == "category")
                        {
                            mstData = mstData.Where(m => m.category.Contains(searchColVal));
                        }
                        else if (fieldName == "requestor_name")
                        {
                            mstData = mstData.Where(m => m.requestor_name.Contains(searchColVal));
                        }
                        else if (fieldName == "create_date")
                        {
                            mstData = mstData.Where(m => m.create_date.Contains(searchColVal));
                        }
                        else if (fieldName == "return_date")
                        {
                            mstData = mstData.Where(m => m.return_date.Contains(searchColVal));
                        }
                        else if (fieldName == "vendor_name")
                        {
                            mstData = mstData.Where(m => m.vendor_name.Contains(searchColVal));
                        }
                        else if (fieldName == "status_gatepass")
                        {
                            mstData = mstData.Where(m => m.status_gatepass.Contains(searchColVal));
                        }
                        else if (fieldName == "status_remark")
                        {
                            mstData = mstData.Where(m => m.status_remark.Contains(searchColVal));
                        }
                    }
                }

                int recordsTotal = mstData.Count();
                var data = mstData.Skip(skip).Take(pageSize).ToList();
                var jsonData = new { draw, recordsFiltered = recordsTotal, recordsTotal, data };
                return Ok(jsonData);
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Internal server error. Please try again later.");
            }
        }

        [HttpPost]
        public IActionResult UPDATE_FINANCE_GATEPASS_NONASSET(int id_gatepass, string approval_status, string approval_remark)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            string name = User.FindFirst("asset_management_name")?.Value;
            string level = User.FindFirst("asset_management_level")?.Value;

            if (string.IsNullOrEmpty(sesa_id))
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }

            try
            {
                var db = new DatabaseAccessLayer();
                var email = new EmailController();

                string ret_order = db.UPDATE_FINANCE_NON_ASSET_GATEPASS(id_gatepass, approval_status, approval_remark, sesa_id);
                string send_email = "";

                if (approval_status == "APPROVED")
                {
                    send_email = email.GATEPASS_COMPLETED_NON_ASSET(id_gatepass, name);
                }
                else if (approval_status == "REJECTED")
                {
                    send_email = email.GATEPASS_REJECTED_NON_ASSET(id_gatepass);
                }

                return Content("success;" + send_email, "text/plain");
            }
            catch (Exception ex)
            {
                return Content($"error;{ex.Message}", "text/plain");
            }
        }

        [HttpGet]
        public IActionResult GET_PROFORMA_FIN_FILES_NONASSET(int id_gatepass)
        {
            try
            {
                var db = new DatabaseAccessLayer();
                var files = db.GET_PROFORMA_NON_ASSET_FIN_FILES(id_gatepass);
                return Json(files);
            }
            catch (Exception ex)
            {
                return Json(new { error = ex.Message });
            }
        }

        [HttpPost]
        public IActionResult UPLOAD_PROFORMA_FIN_FILES_NONASSET(IFormFile file, int id_gatepass, string file_type)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            if (string.IsNullOrEmpty(sesa_id))
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }

            try
            {
                if (file != null && file.Length > 0)
                {
                    var db = new DatabaseAccessLayer();

                    var existingFiles = db.GET_PROFORMA_NON_ASSET_FIN_FILES(id_gatepass);
                    var duplicateFile = existingFiles.FirstOrDefault(f => f.document_type == file_type);

                    if (duplicateFile != null)
                    {
                        return Content($"A file for document type '{file_type}' already exists. Please delete the existing file first or choose a different document type.", "text/plain");
                    }

                    var exp = new ExportController();
                    string file_result = exp.UPLOAD_PROFORMA_NON_ASSET_FIN_FILES(file, file_type);
                    string[] file_res = file_result.Split(';');

                    if (file_res[0] == "OK")
                    {
                        string finalFileName = file_res[1];
                        string result = db.UPLOAD_PROFORMA_NON_ASSET_FIN_FILES(id_gatepass, file_type, finalFileName, sesa_id);

                        return Content("success", "text/plain");
                    }
                    else
                    {
                        return Content($"File upload error: {file_res[1]}", "text/plain");
                    }
                }
                return Content("No file selected", "text/plain");
            }
            catch (Exception ex)
            {
                return Content($"Error: {ex.Message}", "text/plain");
            }
        }

        [HttpPost]
        public IActionResult DELETE_PROFORMA_FIN_FILE_NONASSET(int id_file, int id_gatepass)
        {
            try
            {
                var db = new DatabaseAccessLayer();
                string result = db.DELETE_PROFORMA_NON_ASSET_FIN_FILE(id_file, id_gatepass);
                return Json(result);
            }
            catch (Exception ex)
            {
                return Json($"error;{ex.Message}");
            }
        }

        [HttpPost]
        public IActionResult COMPLETE_PROFORMA_FIN_UPLOAD_NONASSET(int id_gatepass)
        {
            try
            {
                var db = new DatabaseAccessLayer();
                string sesa_id = HttpContext.Session.GetString("sesa_id") ?? "";
                string result = db.UPDATE_PROFORMA_NON_ASSET_FIN_STATUS(id_gatepass, sesa_id);
                return Json(result);
            }
            catch (Exception ex)
            {
                return Json($"error;{ex.Message}");
            }
        }

        [HttpGet]
        public IActionResult GET_PEB_FILE_NONASSET(int id_gatepass)
        {
            try
            {
                var db = new DatabaseAccessLayer();
                var file = db.GET_PEB_NON_ASSET_FILE(id_gatepass);
                return Json(file);
            }
            catch (Exception ex)
            {
                return Json(new { error = ex.Message });
            }
        }

        [HttpPost]
        public IActionResult UPLOAD_PEB_FILE_NONASSET(IFormFileCollection files, int id_gatepass)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            if (string.IsNullOrEmpty(sesa_id))
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }

            try
            {
                if (files != null && files.Count > 0)
                {
                    var exp = new ExportController();
                    string file_result = exp.UPLOAD_PEB_NON_ASSET_FILE(files);
                    string[] file_res = file_result.Split(';');

                    if (file_res[0] == "OK")
                    {
                        string combinedFilenames = file_res[1];
                        var db = new DatabaseAccessLayer();
                        string result = db.UPLOAD_PEB_NON_ASSET_FILE(id_gatepass, combinedFilenames, sesa_id);
                        return Content("success", "text/plain");
                    }
                    else
                    {
                        return Content($"File upload error: {file_res[1]}", "text/plain");
                    }
                }
                return Content("No files selected", "text/plain");
            }
            catch (Exception ex)
            {
                return Content($"Error: {ex.Message}", "text/plain");
            }
        }

        [HttpPost]
        public IActionResult DELETE_PEB_FILE_NONASSET(int id_gatepass, string filename = null)
        {
            try
            {
                var db = new DatabaseAccessLayer();
                string result = db.DELETE_PEB_NON_ASSET_FILE(id_gatepass, filename);
                return Json(result);
            }
            catch (Exception ex)
            {
                return Json($"error;{ex.Message}");
            }
        }

        [HttpPost]
        public IActionResult COMPLETE_PEB_UPLOAD_NONASSET(int id_gatepass)
        {
            try
            {
                string sesa_id = HttpContext.Session.GetString("sesa_id") ?? "";
                var db = new DatabaseAccessLayer();
                string result = db.COMPLETE_PEB_NON_ASSET_UPLOAD(id_gatepass, sesa_id);
                return Json(result);
            }
            catch (Exception ex)
            {
                return Json($"error;{ex.Message}");
            }
        }

        public IActionResult CreateInvoice(string gatepass_no = null)
        {
            return this.CheckSession(() =>
            {
                var db = new DatabaseAccessLayer();
                ViewBag.sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                ViewBag.role = User.FindFirst("asset_management_role")?.Value;
                ViewBag.gatepass_no = gatepass_no;
                return View();
            });
        }

        [HttpPost]
        public IActionResult GET_CREATE_INVOICE()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            string role = User.FindFirst("asset_management_role")?.Value;
            try
            {
                var draw = Request.Form["draw"].FirstOrDefault();
                var start = Request.Form["start"].FirstOrDefault();
                var length = Request.Form["length"].FirstOrDefault();
                var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][data]"].FirstOrDefault();
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                var searchValue = Request.Form["search[value]"].FirstOrDefault();
                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;
                var gatepassWithNonRejectedInvoice = (from invoice in _context.v_gatepass
                                                      where invoice.approval_invoice_status != "REJECTED"
                                                      select invoice.id_gatepass).Distinct();

                var mstData = from gatepass in _context.v_gatepass
                              where gatepass.category == "FA SALES - WITHOUT DISMANTLING"
                              && gatepass.status_gatepass == "Waiting Invoice from Finance"
                              && gatepass.has_doc_no == 1
                              && (gatepass.id_invoice == null || !gatepassWithNonRejectedInvoice.Contains(gatepass.id_gatepass))
                              select new
                              {
                                  gatepass.id_gatepass,
                                  gatepass.gatepass_no,
                                  gatepass.category,
                                  gatepass.type,
                                  gatepass.created_by,
                                  create_date = Convert.ToString(gatepass.create_date),
                                  return_date = Convert.ToString(gatepass.return_date),
                                  gatepass.vendor_name,
                                  gatepass.status_gatepass,
                                  gatepass.status_remark,
                                  gatepass.requestor_name,
                                  gatepass.proforma_no
                              };

                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    mstData = mstData.OrderBy($"{sortColumn} {sortColumnDirection}");
                }

                if (!string.IsNullOrEmpty(searchValue))
                {
                    mstData = mstData.Where(m => m.gatepass_no.Contains(searchValue)
                        || m.proforma_no.Contains(searchValue)
                        || m.vendor_name.Contains(searchValue));
                }

                int recordsTotal = mstData.Count();
                var data = mstData.Skip(skip).Take(pageSize).ToList();
                var jsonData = new { draw, recordsFiltered = recordsTotal, recordsTotal, data };
                return Ok(jsonData);
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Internal server error. Please try again later.");
            }
        }

        [HttpGet]
        public IActionResult GET_GATEPASS_INVOICE_ASSETS(int id_gatepass)
        {
            var db = new DatabaseAccessLayer();
            var assets = db.GET_GATEPASS_INVOICE_ASSETS(id_gatepass);
            return Json(assets);
        }

        [HttpPost]
        public IActionResult INSERT_GATEPASS_INVOICE_DETAIL(int id_gatepass, string selected_assets, decimal amount)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            if (string.IsNullOrEmpty(sesa_id))
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }

            try
            {
                var db = new DatabaseAccessLayer();
                string result = db.INSERT_GATEPASS_INVOICE_DETAIL(id_gatepass, selected_assets, amount, sesa_id);
                return Content(result, "text/plain");
            }
            catch (Exception ex)
            {
                return Content($"error;{ex.Message}", "text/plain");
            }
        }

        [HttpPost]
        public IActionResult DELETE_GATEPASS_INVOICE_DETAIL(int id_invoice_detail)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            if (string.IsNullOrEmpty(sesa_id))
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }

            try
            {
                var db = new DatabaseAccessLayer();
                string result = db.DELETE_GATEPASS_INVOICE_DETAIL(id_invoice_detail, sesa_id);
                return Content(result, "text/plain");
            }
            catch (Exception ex)
            {
                return Content($"error;{ex.Message}", "text/plain");
            }
        }

        [HttpPost]
        public IActionResult UPDATE_GATEPASS_INVOICE_DETAIL(int id_invoice_detail, decimal amount)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            if (string.IsNullOrEmpty(sesa_id))
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }

            try
            {
                var db = new DatabaseAccessLayer();
                string result = db.UPDATE_GATEPASS_INVOICE_DETAIL(id_invoice_detail, amount, sesa_id);
                return Content(result, "text/plain");
            }
            catch (Exception ex)
            {
                return Content($"error;{ex.Message}", "text/plain");
            }
        }

        [HttpPost]
        public IActionResult INSERT_GATEPASS_INVOICE(int id_gatepass, string invoice_currency, DateTime invoice_date, IFormFile main_file, IFormFile secondary_file)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            if (string.IsNullOrEmpty(sesa_id))
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }

            try
            {
                if (string.IsNullOrEmpty(invoice_currency))
                {
                    return Content("error;Invoice currency is required", "text/plain");
                }

                if (main_file == null || main_file.Length == 0)
                {
                    return Content("error;Main invoice file is required", "text/plain");
                }

                var exp = new ExportController();
                string mainFileResult = exp.UPLOAD_INVOICE_FILE(main_file, "main");
                string[] mainFileRes = mainFileResult.Split(';');

                if (mainFileRes[0] != "OK")
                {
                    return Content($"error;Main file upload failed: {mainFileRes[1]}", "text/plain");
                }

                string mainFileName = mainFileRes[1];
                string secondaryFileName = null;

                if (secondary_file != null && secondary_file.Length > 0)
                {
                    string secondaryFileResult = exp.UPLOAD_INVOICE_FILE(secondary_file, "secondary");
                    string[] secondaryFileRes = secondaryFileResult.Split(';');

                    if (secondaryFileRes[0] != "OK")
                    {
                        return Content($"error;Secondary file upload failed: {secondaryFileRes[1]}", "text/plain");
                    }

                    secondaryFileName = secondaryFileRes[1];
                }

                var db = new DatabaseAccessLayer();
                var email = new EmailController();

                string result = db.INSERT_GATEPASS_INVOICE(id_gatepass, invoice_currency, invoice_date, sesa_id, mainFileName, secondaryFileName);

                if (result.StartsWith("success"))
                {
                    string[] resultParts = result.Split(';');
                    if (resultParts.Length > 1)
                    {
                        string invoice_no = resultParts[1];

                        int? id_invoice = null;
                        using (SqlConnection conn = new SqlConnection(DbConnection()))
                        {
                            string query = "SELECT id_invoice FROM tbl_gatepass_invoice_header WHERE invoice_no = @invoice_no";
                            using (SqlCommand cmd = new SqlCommand(query, conn))
                            {
                                cmd.Parameters.AddWithValue("@invoice_no", invoice_no);
                                conn.Open();
                                var queryResult = cmd.ExecuteScalar();
                                if (queryResult != null && queryResult != DBNull.Value)
                                {
                                    id_invoice = Convert.ToInt32(queryResult);
                                }
                                conn.Close();
                            }
                        }

                        if (id_invoice.HasValue)
                        {
                            string send_email = email.EMAIL_REQ_INVOICE_APPROVAL(id_invoice.Value);
                            return Content($"{result};Email: {send_email}", "text/plain");
                        }
                        else
                        {
                            return Content($"{result};Error: Invoice ID not found", "text/plain");
                        }
                    }
                }

                return Content(result, "text/plain");
            }
            catch (Exception ex)
            {
                return Content($"error;{ex.Message}", "text/plain");
            }
        }

        public IActionResult InvoiceList(string invoice_no = null)
        {
            return this.CheckSession(() =>
            {
                ViewBag.sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                ViewBag.role = User.FindFirst("asset_management_role")?.Value;
                ViewBag.invoice_no = invoice_no;
                return View();
            });
        }

        [HttpPost]
        public IActionResult GET_INVOICE_LIST()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            string role = User.FindFirst("asset_management_role")?.Value;
            try
            {
                var draw = Request.Form["draw"].FirstOrDefault();
                var start = Request.Form["start"].FirstOrDefault();
                var length = Request.Form["length"].FirstOrDefault();
                var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][data]"].FirstOrDefault();
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                var searchValue = Request.Form["search[value]"].FirstOrDefault();
                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;

                var mstData = from gatepass in _context.v_gatepass
                              where gatepass.category == "FA SALES - WITHOUT DISMANTLING"
                              && gatepass.id_invoice != null
                              select new
                              {
                                  gatepass.id_gatepass,
                                  gatepass.gatepass_no,
                                  gatepass.id_invoice,
                                  gatepass.invoice_no,
                                  gatepass.vendor_name,
                                  gatepass.invoice_date,
                                  gatepass.status_gatepass,
                                  gatepass.status_remark,
                                  gatepass.approval_invoice_status,
                                  gatepass.invoice_by,
                                  gatepass.requestor_name,
                                  gatepass.create_date
                              };

                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    mstData = mstData.OrderBy($"{sortColumn} {sortColumnDirection}");
                }

                if (!string.IsNullOrEmpty(searchValue))
                {
                    mstData = mstData.Where(m => m.invoice_no.Contains(searchValue)
                        || m.gatepass_no.Contains(searchValue)
                        || m.vendor_name.Contains(searchValue));
                }

                int recordsTotal = mstData.Count();
                var dataList = mstData.Skip(skip).Take(pageSize).ToList();

                var formattedData = dataList.Select(d => new
                {
                    d.id_gatepass,
                    d.gatepass_no,
                    d.id_invoice,
                    d.invoice_no,
                    d.vendor_name,
                    invoice_date = d.invoice_date?.ToString("dd MMM yyyy") ?? "",
                    d.status_gatepass,
                    d.status_remark,
                    d.approval_invoice_status,
                    d.invoice_by,
                    d.requestor_name,
                    created_date = d.create_date
                }).ToList();

                var jsonData = new { draw, recordsFiltered = recordsTotal, recordsTotal, data = formattedData };
                return Ok(jsonData);
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Internal server error. Please try again later.");
            }
        }

        [HttpGet]
        public IActionResult GET_INVOICE_DETAIL(int id_invoice)
        {
            try
            {
                var db = new DatabaseAccessLayer();
                var result = db.GET_INVOICE_DETAIL(id_invoice);

                return Json(new
                {
                    assets = result.assets ?? new List<GatePassInvoiceAssetModel>(),
                    invoice_main_file = result.invoice_main_file,
                    invoice_secondary_file = result.invoice_secondary_file,
                    invoice_payment = result.invoice_payment,
                    invoice_currency = result.invoice_currency
                });
            }
            catch (Exception ex)
            {
                return Json(new
                {
                    assets = new List<GatePassInvoiceAssetModel>(),
                    invoice_main_file = (string)null,
                    invoice_secondary_file = (string)null,
                    invoice_payment = (string)null,
                    invoice_currency = (string)null
                });
            }
        }

        public IActionResult InvoicePDF(int id_invoice)
        {
            return RedirectToAction("InvoicePDF", "Export", new { id_invoice });
        }

        public IActionResult InvoiceApprovalList(string invoice_no = null)
        {
            return this.CheckSession(() =>
            {
                ViewBag.sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                ViewBag.role = User.FindFirst("asset_management_role")?.Value;
                ViewBag.invoice_no = invoice_no;
                return View();
            });
        }

        [HttpPost]
        public IActionResult GET_INVOICE_APPROVAL_LIST()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            try
            {
                var draw = Request.Form["draw"].FirstOrDefault();
                var start = Request.Form["start"].FirstOrDefault();
                var length = Request.Form["length"].FirstOrDefault();
                var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][data]"].FirstOrDefault();
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                var searchValue = Request.Form["search[value]"].FirstOrDefault();
                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;

                var mstData = from gatepass in _context.v_gatepass
                              where gatepass.category == "FA SALES - WITHOUT DISMANTLING"
                              && gatepass.status_gatepass == "Waiting Accounting Manager Approval"
                              && gatepass.id_invoice != null
                              && gatepass.approval_invoice_status == "OPEN"
                              && gatepass.approval_invoice_sesa == sesa_id
                              select new
                              {
                                  gatepass.id_gatepass,
                                  gatepass.gatepass_no,
                                  gatepass.id_invoice,
                                  gatepass.invoice_no,
                                  gatepass.vendor_name,
                                  gatepass.invoice_currency,
                                  gatepass.invoice_date,
                                  gatepass.approval_invoice_status,
                                  gatepass.invoice_by,
                                  gatepass.requestor_name,
                                  gatepass.create_date
                              };

                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    mstData = mstData.OrderBy($"{sortColumn} {sortColumnDirection}");
                }

                if (!string.IsNullOrEmpty(searchValue))
                {
                    mstData = mstData.Where(m => m.invoice_no.Contains(searchValue)
                        || m.gatepass_no.Contains(searchValue)
                        || m.vendor_name.Contains(searchValue));
                }

                int recordsTotal = mstData.Count();
                var dataList = mstData.Skip(skip).Take(pageSize).ToList();

                var formattedData = dataList.Select(d => new
                {
                    d.id_gatepass,
                    d.gatepass_no,
                    d.id_invoice,
                    d.invoice_no,
                    d.vendor_name,
                    d.invoice_currency,
                    invoice_date = d.invoice_date?.ToString("dd MMM yyyy") ?? "",
                    d.approval_invoice_status,
                    d.invoice_by,
                    d.requestor_name,
                    created_date = d.create_date
                }).ToList();

                var jsonData = new { draw, recordsFiltered = recordsTotal, recordsTotal, data = formattedData };
                return Ok(jsonData);
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Internal server error. Please try again later.");
            }
        }

        [HttpPost]
        public IActionResult UPDATE_APPROVAL_GATEPASS_INVOICE(int id_invoice, string approval_status, string approval_remark)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            string name = User.FindFirst("asset_management_name")?.Value;

            if (string.IsNullOrEmpty(sesa_id))
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }

            try
            {
                var db = new DatabaseAccessLayer();
                var email = new EmailController();

                string result = db.UPDATE_APPROVAL_GATEPASS_INVOICE(id_invoice, approval_status, approval_remark, sesa_id);
                string send_email = "";

                if (approval_status == "APPROVED")
                {
                    send_email = email.EMAIL_INVOICE_APPROVED(id_invoice, name);
                }
                else if (approval_status == "REJECTED")
                {
                    send_email = email.EMAIL_INVOICE_REJECTED(id_invoice);
                }

                return Content($"{result};{send_email}", "text/plain");
            }
            catch (Exception ex)
            {
                return Content($"error;{ex.Message};", "text/plain");
            }
        }

        [HttpPost]
        public IActionResult UPLOAD_INVOICE_PAYMENT(IFormFile file, int id_invoice)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            if (string.IsNullOrEmpty(sesa_id))
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }

            try
            {
                if (file == null || file.Length == 0)
                {
                    return Content("error;No file selected", "text/plain");
                }

                var exp = new ExportController();
                string file_result = exp.UPLOAD_INVOICE_PAYMENT_FILE(file);
                string[] file_res = file_result.Split(';');

                if (file_res[0] == "OK")
                {
                    string finalFileName = file_res[1];
                    var db = new DatabaseAccessLayer();
                    string result = db.UPDATE_GATEPASS_INVOICE_PAYMENT(id_invoice, finalFileName, sesa_id);

                    if (result == "success")
                    {
                        return Content("success", "text/plain");
                    }
                    else
                    {
                        return Content(result, "text/plain");
                    }
                }
                else
                {
                    return Content($"error;File upload failed: {file_res[1]}", "text/plain");
                }
            }
            catch (Exception ex)
            {
                return Content($"error;{ex.Message}", "text/plain");
            }
        }

        [HttpGet]
        public IActionResult GET_INVOICE_PAYMENT_FILE(int id_invoice)
        {
            try
            {
                var db = new DatabaseAccessLayer();
                string filename = db.GET_INVOICE_PAYMENT_FILE(id_invoice);
                return Json(new { filename = filename });
            }
            catch (Exception ex)
            {
                return Json(new { error = ex.Message });
            }
        }

        private string GetMimeType(string extension)
        {
            return extension switch
            {
                ".msg" => "application/vnd.ms-outlook",
                ".pdf" => "application/pdf",
                ".jpg" or ".jpeg" => "image/jpeg",
                ".png" => "image/png",
                ".doc" => "application/msword",
                ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".xls" => "application/vnd.ms-excel",
                ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                _ => "application/octet-stream"
            };
        }

        [HttpGet]
        public IActionResult DOWNLOAD_INVOICE_FILE(string fileName, string fileType)
        {
            try
            {
                string uploadPath = fileType.ToLower() == "payment"
                    ? System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Upload", "Gatepass_Invoice_Payment")
                    : System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Upload", "Gatepass_Invoice");

                string fullPath = System.IO.Path.Combine(uploadPath, fileName);

                if (!System.IO.File.Exists(fullPath))
                {
                    return NotFound("File not found");
                }

                byte[] fileBytes = System.IO.File.ReadAllBytes(fullPath);

                string extension = System.IO.Path.GetExtension(fileName).ToLower();

                string mimeType = GetMimeType(extension);

                Response.Headers["Content-Disposition"] = $"attachment; filename=\"{fileName}\"";
                return File(fileBytes, mimeType, fileName);
            }
            catch (Exception ex)
            {
                return Content($"Error downloading file: {ex.Message}", "text/plain");
            }
        }

    }
}
