using Asset_Management.Function;
using Asset_Management.Models;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.Data.SqlClient;
using System.Linq.Dynamic.Core;
using System.Security.Claims;

namespace Asset_Management.Controllers
{
    [Authorize(Policy = "RequireFinanceAdmin")]
    public class AdminController : Controller
    {
        private readonly ApplicationDbContext _context;
        private string DbConnection()
        {
            var dbAccess = new DatabaseAccessLayer();
            string dbString = dbAccess.ConnectionString;
            return dbString;
        }
        public AdminController(ApplicationDbContext context)
        {
            _context = context;
        }
        public IActionResult Index()
        {
            return View();
        }
        public IActionResult UserList()
        {
            var db = new DatabaseAccessLayer();
            List<string> deptList = db.GET_DEPARTMENT_LIST();
            List<string> plantList = db.GET_PLANT_LIST();
            ViewBag.deptList = deptList;
            ViewBag.plantList = plantList;
            return View();
        }
        public IActionResult GetUserList()
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
                var mstData = (from UserDetail in _context.mst_users
                               select
                                   new
                                   {
                                       UserDetail.id_user,
                                       UserDetail.sesa_id,
                                       UserDetail.name,
                                       UserDetail.level,
                                       UserDetail.role,
                                       UserDetail.email,
                                       UserDetail.department,
                                       UserDetail.plant
                                   });

                if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDirection)))
                {
                    mstData = mstData.OrderBy(sortColumn + " " + sortColumnDirection);
                }
                if (!string.IsNullOrEmpty(searchValue))
                {
                    mstData = mstData.Where(m => m.sesa_id.Contains(searchValue)
                                                || m.name.Contains(searchValue)
                                                || m.level.Contains(searchValue)
                                                || m.role.Contains(searchValue)
                                                || m.department.Contains(searchValue)
                                                || m.plant.Contains(searchValue));
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
        public IActionResult ADD_NEW_USER(string sesa_id, string name, string email, string department, string plant, string level, string role, string manager_sesa_id, string password = "")
        {
            var db = new DatabaseAccessLayer();
            string status_msg;

            if (!string.IsNullOrEmpty(password))
            {
                status_msg = db.ADD_NEW_USER_WITH_PASSWORD(sesa_id, name, email, department, plant, level, role, manager_sesa_id, password);
            }
            else
            {
                status_msg = db.ADD_NEW_USER(sesa_id, name, email, department, plant, level, role, manager_sesa_id);
            }

            return Content(status_msg, "text/plain");
        }

        [HttpPost]
        public IActionResult DELETE_USER(string id_user)
        {
            var db = new DatabaseAccessLayer();
            int rowsAffected = db.DELETE_USER_BY_ID(id_user);
            return Json(rowsAffected);
        }

        [HttpGet]
        public IActionResult GET_DETAIL_USER(string id_user)
        {
            var db = new DatabaseAccessLayer();
            var user = db.GET_USER_BY_ID(id_user);
            if (user != null)
            {
                return Json(new
                {
                    sesa_id = user.sesa_id,
                    name = user.name,
                    email = user.email,
                    department = user.department,
                    plant = user.plant,
                    level = user.level,
                    role = user.role,
                    manager_sesa_id = user.manager_sesa_id
                });
            }
            return NotFound();
        }

        [HttpPost]
        public async Task<IActionResult> UPDATE_USER(string e_sesa_id, string e_name, string e_email, string e_department, string e_plant, string e_level, string e_role, string e_manager_sesa_id, string e_password = "")
        {
            var db = new DatabaseAccessLayer();
            string status_msg = db.UPDATE_USER(e_sesa_id, e_name, e_email, e_department, e_plant, e_level, e_role, e_manager_sesa_id);

            if (status_msg != "success")
                return Content(status_msg + ";;Failed to update!", "text/plain");

            if (!string.IsNullOrEmpty(e_password))
            {
                db.SaveNewPass(e_sesa_id, e_password);
            }

            string currentSesaId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            bool isSelf = !string.IsNullOrEmpty(currentSesaId) &&
                          string.Equals(currentSesaId, e_sesa_id, StringComparison.OrdinalIgnoreCase);

            if (isSelf)
            {
                var claimsIdentity = (ClaimsIdentity)User.Identity;

                var claimsToUpdate = new Dictionary<string, string>
                {
                    { "asset_management_name",  e_name  ?? "" },
                    { "asset_management_level", e_level ?? "" },
                    { "asset_management_role",  e_role  ?? e_level },
                };

                foreach (var (claimType, newValue) in claimsToUpdate)
                {
                    var existing = claimsIdentity.FindFirst(claimType);
                    if (existing != null)
                        claimsIdentity.RemoveClaim(existing);
                    claimsIdentity.AddClaim(new Claim(claimType, newValue));
                }

                await HttpContext.SignInAsync(
                    CookieAuthenticationDefaults.AuthenticationScheme,
                    new ClaimsPrincipal(claimsIdentity));

                return Content("success;cookie_updated;Successfully Updated! Your level/role has been refreshed.", "text/plain");
            }

            return Content("success;;Successfully Updated!", "text/plain");
        }

        [HttpGet]
        public IActionResult GetApproverList(string search = "", int page = 1, string excludeSesaId = "")
        {
            var db = new DatabaseAccessLayer();
            var approvers = db.GetApproverList(search, excludeSesaId);
            return Json(approvers);
        }

        [HttpGet]
        public IActionResult GetUserDelegationInfo(string sesa_id)
        {
            var db = new DatabaseAccessLayer();
            var userInfo = db.GetUserDelegationInfo(sesa_id);
            return Json(userInfo);
        }

        [HttpPost]
        public IActionResult SetDelegation(string original_sesa_id, string delegated_to)
        {
            var db = new DatabaseAccessLayer();
            string status_msg = db.SetDelegation(original_sesa_id, delegated_to);
            return Content(status_msg, "text/plain");
        }

        [HttpPost]
        public IActionResult RemoveDelegation(string sesa_id)
        {
            var db = new DatabaseAccessLayer();
            string status_msg = db.RemoveDelegation(sesa_id);
            return Content(status_msg, "text/plain");
        }
    }
}
