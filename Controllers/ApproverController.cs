using System.Data.SqlClient;
using Asset_Management.Service;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Mvc;
using Asset_Management.Models;
using Asset_Management.Function;
using System.Data;
using System.Linq.Dynamic.Core;
using Microsoft.AspNetCore.Authorization;
using System.Security.Claims;
using OfficeOpenXml;
using System.Linq;

namespace Asset_Management.Controllers
{
    public class ApproverController : Controller
    {
        private readonly ApplicationDbContext _context;

        private readonly ILogger<ApproverController> _logger;

        private string DbConnection()
        {
            var dbAccess = new DatabaseAccessLayer();
            string dbString = dbAccess.ConnectionString;
            return dbString;
        }

        public ApproverController(ILogger<ApproverController> logger, ApplicationDbContext context)
        {
            this._context = context;
            _logger = logger;
        }
        
        [Authorize(Policy = "RequireApprover")]
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

        [Authorize(Policy = "RequireApprover")]
        public IActionResult GetGatepassApprovalListBadge()
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
                var count = (from GatepassApprovalList in _context.v_gatepass
                             where (GatepassApprovalList.status_gatepass == "OPEN"
                             && ((GatepassApprovalList.approval_hod == sesa_id && GatepassApprovalList.approval_hod_status == "OPEN")
                             || (GatepassApprovalList.approval_fbp == sesa_id && GatepassApprovalList.approval_fbp_status == "OPEN")
                             || (GatepassApprovalList.approval_ph == sesa_id && GatepassApprovalList.approval_ph_status == "OPEN")))
                             || (GatepassApprovalList.status_gatepass == "Waiting Confirmation by HOD"
                             && GatepassApprovalList.approval_hod == sesa_id)
                             select GatepassApprovalList).Count();
                return Ok(new { count = count });
            }
            catch (Exception)
            {
                return Ok(new { count = 0 });
            }
        }

        [Authorize(Policy = "RequireApprover")]
        public IActionResult GetGatepassNonAssetApprovalListBadge()
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
                var count = (from GatepassNonAssetApprovalList in _context.v_gatepass_non_asset
                             where (GatepassNonAssetApprovalList.status_gatepass == "OPEN"
                             && ((GatepassNonAssetApprovalList.approval_hod == sesa_id && GatepassNonAssetApprovalList.approval_hod_status == "OPEN")
                             || (GatepassNonAssetApprovalList.approval_fbp == sesa_id && GatepassNonAssetApprovalList.approval_fbp_status == "OPEN")
                             || (GatepassNonAssetApprovalList.approval_ph == sesa_id && GatepassNonAssetApprovalList.approval_ph_status == "OPEN")))
                             || (GatepassNonAssetApprovalList.status_gatepass == "Waiting Confirmation by HOD"
                             && GatepassNonAssetApprovalList.approval_hod == sesa_id)
                             select GatepassNonAssetApprovalList).Count();
                return Ok(new { count = count });
            }
            catch (Exception)
            {
                return Ok(new { count = 0 });
            }
        }

        [Authorize(Policy = "RequireApprover")]
        public IActionResult GatePassList(string gatepass_no = null)
        {
            return this.CheckSession(() =>
            {
                ViewBag.sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                ViewBag.role = User.FindFirst("asset_management_role")?.Value;
                ViewBag.gatepass_no = gatepass_no;
                return View();
            });
        }

        [Authorize(Policy = "RequireApprover")]
        public IActionResult GET_GATEPASS_HEADER()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            string role = User.FindFirst("asset_management_role")?.Value;

            try
            {
                var db = new DatabaseAccessLayer();
                var status = Request.Form["status"].FirstOrDefault();
                var draw = Request.Form["draw"].FirstOrDefault();
                var start = Request.Form["start"].FirstOrDefault();
                var length = Request.Form["length"].FirstOrDefault();
                var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][data]"].FirstOrDefault();
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                var searchValue = Request.Form["search[value]"].FirstOrDefault();
                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;

                List<string> originalApprovers = db.GetAllOriginalApproversForDelegated(sesa_id);
                List<string> allowedGatepassStatuses = new List<string>();
                List<string> allowedApprovalStatuses = new List<string>();

                if (status == "OPEN APPROVAL")
                {
                    allowedGatepassStatuses.Add("OPEN");
                    allowedGatepassStatuses.Add("Waiting Confirmation by HOD");
                    allowedApprovalStatuses.Add("OPEN");
                }
                else if (status == "OPEN")
                {
                    allowedGatepassStatuses.Add("OPEN");
                    allowedGatepassStatuses.Add("Waiting Security Validation");
                    allowedGatepassStatuses.Add("Waiting Invoice from Finance");
                    allowedGatepassStatuses.Add("TRANSFER");
                    allowedGatepassStatuses.Add("Waiting Confirmation by FIN SS");
                    allowedGatepassStatuses.Add("Waiting Proforma Documents");
                    allowedGatepassStatuses.Add("Waiting Shipping Process");
                    allowedGatepassStatuses.Add("Waiting PEB Documents");
                    allowedApprovalStatuses.AddRange(new[] { "OPEN", "APPROVED" });
                }
                else
                {
                    allowedGatepassStatuses.AddRange(new[] { "COMPLETED", "REJECTED" });
                    allowedApprovalStatuses.AddRange(new[] { "APPROVED", "REJECTED" });
                }

                var mstData = from gatepass in _context.v_gatepass
                              where allowedGatepassStatuses.Contains(gatepass.status_gatepass) &&
                                    (
                                  (gatepass.approval_hod == sesa_id && allowedApprovalStatuses.Contains(gatepass.approval_hod_status)) ||
                                  (gatepass.approval_fbp == sesa_id && allowedApprovalStatuses.Contains(gatepass.approval_fbp_status)) ||
                                  (gatepass.confirmation_hod == sesa_id && allowedApprovalStatuses.Contains(gatepass.confirmation_hod_status)) ||
                                  (gatepass.approval_ph == sesa_id && allowedApprovalStatuses.Contains(gatepass.approval_ph_status)) ||
                                  (originalApprovers.Contains(gatepass.approval_hod) && allowedApprovalStatuses.Contains(gatepass.approval_hod_status)) ||
                                  (originalApprovers.Contains(gatepass.approval_fbp) && allowedApprovalStatuses.Contains(gatepass.approval_fbp_status)) ||
                                  (originalApprovers.Contains(gatepass.confirmation_hod) && allowedApprovalStatuses.Contains(gatepass.confirmation_hod_status)) ||
                                  (originalApprovers.Contains(gatepass.approval_ph) && allowedApprovalStatuses.Contains(gatepass.approval_ph_status))
                              )
                              select new
                              {
                                  gatepass.id_gatepass,
                                  gatepass.gatepass_no,
                                  gatepass.category,
                                  gatepass.created_by,
                                  gatepass.type,
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
                                  gatepass.confirmation_hod,
                                  gatepass.confirmation_hod_status,
                                  gatepass.requestor_name,
                                  gatepass.proforma_no,
                                  can_approve_hod = (gatepass.approval_hod == sesa_id || originalApprovers.Contains(gatepass.approval_hod)),
                                  can_approve_fbp = (gatepass.approval_fbp == sesa_id || originalApprovers.Contains(gatepass.approval_fbp)),
                                  can_approve_ph = (gatepass.approval_ph == sesa_id || originalApprovers.Contains(gatepass.approval_ph)),
                                  can_confirm_hod = (gatepass.confirmation_hod == sesa_id || originalApprovers.Contains(gatepass.confirmation_hod)),
                                  is_delegated_hod = originalApprovers.Contains(gatepass.approval_hod),
                                  is_delegated_fbp = originalApprovers.Contains(gatepass.approval_fbp),
                                  is_delegated_ph = originalApprovers.Contains(gatepass.approval_ph),
                                  is_delegated_confirm_hod = originalApprovers.Contains(gatepass.confirmation_hod),
                                  original_approver_hod = gatepass.approval_hod,
                                  original_approver_fbp = gatepass.approval_fbp,
                                  original_approver_ph = gatepass.approval_ph,
                                  original_approver_confirm_hod = gatepass.confirmation_hod
                              };

                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    mstData = mstData.OrderBy($"{sortColumn} {sortColumnDirection}");
                }

                if (!string.IsNullOrEmpty(searchValue))
                {
                    mstData = mstData.Where(m => m.gatepass_no.Contains(searchValue) || m.status_gatepass.Contains(searchValue) || m.proforma_no.Contains(searchValue));
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
        [Authorize(Policy = "RequireApprover")]
        public IActionResult UPDATE_APPROVAL_GATEPASS(int id_gatepass, string approval_level, string approval_status, string approval_remark)
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
                string ret_order = db.UPDATE_APPROVAL_GATEPASS(id_gatepass, approval_level, approval_status, sesa_id, approval_remark);
                string send_email_approved = "";
                string send_email = "";
                if (approval_status == "APPROVED")
                {
                    string actualApprover = db.GetGatepassApproverAtLevel(id_gatepass, approval_level);
                    bool isDelegationInvolved = db.IsDelegatedApprover(sesa_id, actualApprover);
                    string originalApprover = isDelegationInvolved ? actualApprover : "";
                    send_email_approved = email.GATEPASS_APPROVED(id_gatepass, name, level, isDelegationInvolved, originalApprover, sesa_id);
                    string status_gatepass = db.GET_STATUS_GATEPASS(id_gatepass);
                    if (status_gatepass == "OPEN")
                    {
                        send_email = email.EMAIL_REQ_GATEPASS(id_gatepass, isDelegationInvolved, originalApprover);
                    }
                    else
                    {
                        send_email = email.GATEPASS_COMPLETE(id_gatepass);
                    }
                }
                else if (approval_status == "REJECTED")
                {
                    send_email = email.GATEPASS_REJECTED(id_gatepass);
                }
                return Content("success;" + send_email, "text/plain");
            }
        }

        [Authorize(Policy = "RequireApprover")]
        public IActionResult CONFIRM_GATEPASS_RETURN(int id_gatepass, string approval_level, string approval_status, string approval_remark)
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
                string ret_order = db.CONFIRM_GATEPASS_RETURN(id_gatepass, approval_level, approval_status, sesa_id, approval_remark);
                string send_email_approved = "";
                string actualConfirmer = db.GetGatepassConfirmerHOD(id_gatepass);
                bool isDelegationInvolved = db.IsDelegatedApprover(sesa_id, actualConfirmer);
                string originalApprover = isDelegationInvolved ? actualConfirmer : "";
                string send_email = email.GATEPASS_APPROVED(id_gatepass, name, level, isDelegationInvolved, originalApprover, sesa_id);
                string send_email_req = email.EMAIL_REQ_GATEPASS_RETURN_FINANCE_CONFIRMATION(id_gatepass);
                return Content("success;" + send_email, "text/plain");
            }
        }

        [Authorize(Policy = "RequireApprover")]
        public IActionResult GatePassNonAssetList(string gatepass_no = null)
        {
            return this.CheckSession(() =>
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                string role = User.FindFirst("asset_management_role")?.Value;
                var db = new DatabaseAccessLayer();
                List<UserDetailModel> userDetail = db.GetUserDetail(sesa_id);
                ViewBag.gatepass_no = gatepass_no;
                ViewBag.role = role;
                ViewBag.sesa_id = sesa_id;
                return View(userDetail);
            });
        }

        [Authorize(Policy = "RequireApprover")]
        [HttpPost]
        public IActionResult GET_GATEPASS_NONASSET_HEADER()
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                string role = User.FindFirst("asset_management_role")?.Value;
                var db = new DatabaseAccessLayer();

                var draw = Request.Form["draw"].FirstOrDefault();
                var start = Request.Form["start"].FirstOrDefault();
                var length = Request.Form["length"].FirstOrDefault();
                var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][data]"].FirstOrDefault();
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                var searchValue = Request.Form["search[value]"].FirstOrDefault();
                var status = Request.Form["status"].FirstOrDefault();

                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;

                List<string> originalApprovers = db.GetAllOriginalApproversForDelegated(sesa_id);
                List<string> allowedGatepassStatuses = new List<string>();
                List<string> allowedApprovalStatuses = new List<string>();

                if (status == "OPEN APPROVAL")
                {
                    allowedGatepassStatuses.Add("OPEN");
                    allowedGatepassStatuses.Add("Waiting Confirmation by HOD");
                    allowedApprovalStatuses.Add("OPEN");
                }
                else if (status == "OPEN")
                {
                    allowedGatepassStatuses.Add("OPEN");
                    allowedGatepassStatuses.Add("Waiting Security Validation");
                    allowedGatepassStatuses.Add("TRANSFER");
                    allowedGatepassStatuses.Add("Waiting Confirmation by FIN SS");
                    allowedGatepassStatuses.Add("Waiting Proforma Documents");
                    allowedGatepassStatuses.Add("Waiting Shipping Process");
                    allowedGatepassStatuses.Add("Waiting PEB Documents");
                    allowedApprovalStatuses.AddRange(new[] { "OPEN", "APPROVED" });
                }
                else
                {
                    allowedGatepassStatuses.AddRange(new[] { "COMPLETED", "REJECTED" });
                    allowedApprovalStatuses.AddRange(new[] { "APPROVED", "REJECTED" });
                }

                var mstData = from gatepass in _context.v_gatepass_non_asset
                              where allowedGatepassStatuses.Contains(gatepass.status_gatepass) &&
                                    (
                                        (gatepass.approval_hod == sesa_id && allowedApprovalStatuses.Contains(gatepass.approval_hod_status)) ||
                                        (gatepass.approval_fbp == sesa_id && allowedApprovalStatuses.Contains(gatepass.approval_fbp_status)) ||
                                        (gatepass.confirmation_hod == sesa_id && allowedApprovalStatuses.Contains(gatepass.confirmation_hod_status)) ||
                                        (gatepass.approval_ph == sesa_id && allowedApprovalStatuses.Contains(gatepass.approval_ph_status)) ||
                                        (originalApprovers.Contains(gatepass.approval_hod) && allowedApprovalStatuses.Contains(gatepass.approval_hod_status)) ||
                                        (originalApprovers.Contains(gatepass.approval_fbp) && allowedApprovalStatuses.Contains(gatepass.approval_fbp_status)) ||
                                        (originalApprovers.Contains(gatepass.confirmation_hod) && allowedApprovalStatuses.Contains(gatepass.confirmation_hod_status)) ||
                                        (originalApprovers.Contains(gatepass.approval_ph) && allowedApprovalStatuses.Contains(gatepass.approval_ph_status))
                                    )
                              select new
                              {
                                  gatepass.id_gatepass,
                                  gatepass.gatepass_no,
                                  gatepass.category,
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
                                  gatepass.requestor_name,
                                  create_date = Convert.ToString(gatepass.create_date),
                                  gatepass.status_gatepass,
                                  gatepass.status_remark,
                                  gatepass.proforma_no,
                                  gatepass.approval_hod,
                                  gatepass.approval_hod_status,
                                  gatepass.approval_fbp,
                                  gatepass.approval_fbp_status,
                                  gatepass.approval_ph,
                                  gatepass.approval_ph_status,
                                  gatepass.confirmation_hod,
                                  gatepass.confirmation_hod_status,
                                  can_approve_hod = (gatepass.approval_hod == sesa_id || originalApprovers.Contains(gatepass.approval_hod)),
                                  can_approve_fbp = (gatepass.approval_fbp == sesa_id || originalApprovers.Contains(gatepass.approval_fbp)),
                                  can_approve_ph = (gatepass.approval_ph == sesa_id || originalApprovers.Contains(gatepass.approval_ph)),
                                  can_confirm_hod = (gatepass.confirmation_hod == sesa_id || originalApprovers.Contains(gatepass.confirmation_hod)),
                                  is_delegated_hod = originalApprovers.Contains(gatepass.approval_hod),
                                  is_delegated_fbp = originalApprovers.Contains(gatepass.approval_fbp),
                                  is_delegated_ph = originalApprovers.Contains(gatepass.approval_ph),
                                  is_delegated_confirm_hod = originalApprovers.Contains(gatepass.confirmation_hod),
                                  original_approver_hod = gatepass.approval_hod,
                                  original_approver_fbp = gatepass.approval_fbp,
                                  original_approver_ph = gatepass.approval_ph,
                                  original_approver_confirm_hod = gatepass.confirmation_hod
                              };

                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    mstData = mstData.OrderBy($"{sortColumn} {sortColumnDirection}");
                }

                if (!string.IsNullOrEmpty(searchValue))
                {
                    mstData = mstData.Where(m => m.gatepass_no.Contains(searchValue)
                        || m.category.Contains(searchValue)
                        || (m.vendor_code ?? "").Contains(searchValue)
                        || (m.vendor_name ?? "").Contains(searchValue));
                }

                var recordsTotal = mstData.Count();
                var data = mstData.Skip(skip).Take(pageSize).ToList();
                var jsonData = new { draw, recordsFiltered = recordsTotal, recordsTotal, data };
                return Ok(jsonData);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in GET_GATEPASS_NONASSET_HEADER");
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [Authorize(Policy = "RequireApprover")]
        [HttpPost]
        public IActionResult UPDATE_APPROVAL_GATEPASS_NONASSET(int id_gatepass, string approval_level, string approval_status, string approval_remark)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            string name = User.FindFirst("asset_management_name")?.Value;
            if (sesa_id == "")
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }
            else
            {
                var db = new DatabaseAccessLayer();
                var email = new EmailController();
                string ret_order = db.UPDATE_APPROVAL_NON_ASSET_GATEPASS(id_gatepass, approval_level, approval_status, sesa_id, approval_remark);
                string send_email_approved = "";
                string send_email = "";
                if (approval_status == "APPROVED")
                {
                    string actualApprover = db.GetGatepassNonAssetApproverAtLevel(id_gatepass, approval_level);
                    bool isDelegationInvolved = db.IsDelegatedApprover(sesa_id, actualApprover);
                    string originalApprover = isDelegationInvolved ? actualApprover : "";
                    send_email_approved = email.GATEPASS_APPROVED_NON_ASSET(id_gatepass, name, isDelegationInvolved, originalApprover, sesa_id);

                    using (SqlConnection conn = new SqlConnection(db.ConnectionString))
                    {
                        conn.Open();
                        using (SqlCommand cmd = new SqlCommand(@"
                    SELECT COUNT(*) FROM tbl_gatepass_non_asset_approval 
                    WHERE id_gatepass = @id AND approval_status = 'OPEN'", conn))
                        {
                            cmd.Parameters.AddWithValue("@id", id_gatepass);
                            int openApprovals = (int)cmd.ExecuteScalar();

                            if (openApprovals > 0)
                            {
                                send_email = email.EMAIL_REQ_NON_ASSET_GATEPASS(id_gatepass, isDelegationInvolved, originalApprover);
                            }
                            else
                            {
                                send_email = email.GATEPASS_COMPLETE_NON_ASSET(id_gatepass);
                            }
                        }
                    }
                }
                else if (approval_status == "REJECTED")
                {
                    send_email = email.GATEPASS_REJECTED_NON_ASSET(id_gatepass);
                }
                return Content("success;" + send_email, "text/plain");
            }
        }

        [Authorize(Policy = "RequireApprover")]
        [HttpPost]
        public IActionResult CONFIRM_GATEPASS_NONASSET_RETURN(int id_gatepass, string approval_level, string approval_status, string approval_remark)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            string name = User.FindFirst("asset_management_name")?.Value;
            if (sesa_id == "")
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }
            else
            {
                var db = new DatabaseAccessLayer();
                var email = new EmailController();
                string ret_order = db.CONFIRM_GATEPASS_NON_ASSET_RETURN(id_gatepass, approval_level, approval_status, sesa_id, approval_remark);
                string actualConfirmer = db.GetGatepassNonAssetConfirmerHOD(id_gatepass);
                bool isDelegationInvolved = db.IsDelegatedApprover(sesa_id, actualConfirmer);
                string originalApprover = isDelegationInvolved ? actualConfirmer : "";
                string send_email = email.GATEPASS_APPROVED_NON_ASSET(id_gatepass, name, isDelegationInvolved, originalApprover, sesa_id);
                return Content("success;" + send_email, "text/plain");
            }
        }


    }
}
