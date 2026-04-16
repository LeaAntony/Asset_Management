using iText.Kernel.Pdf.Canvas.Parser.ClipperLib;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Asset_Management.Function;
using Asset_Management.Models;
using System.Data.SqlClient;
using System.Data;
using System.Diagnostics;
using System.Security.Claims;

namespace Asset_Management.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IConfiguration _configuration;
        private readonly ApplicationDbContext _context;

        private string DbConnection()
        {
            var dbAccess = new DatabaseAccessLayer();
            string dbString = dbAccess.ConnectionString;
            return dbString;
        }

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration, ApplicationDbContext context)
        {
            _logger = logger;
            _configuration = configuration;
            _context = context;
        }

        [AllowAnonymous]
        public IActionResult Index()
        {
            return View();
        }

        [Authorize]
        public async Task<IActionResult> Login()
        {
            string name = User.FindFirst("asset_management_name")?.Value;
            string level = User.FindFirst("asset_management_level")?.Value;
            if (name == null || level == null)
            {
                return RedirectToAction("Index", "Auth");
            }
            else if (level == "no_access")
            {
                var claimsIdentity = (ClaimsIdentity)User.Identity;

                var existLevel = claimsIdentity?.FindFirst("asset_management_level");
                if (existLevel != null)
                {
                    claimsIdentity.RemoveClaim(existLevel);
                }
                await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(claimsIdentity));
            }

            return RedirectToAction("Open", "Home");
        }
        [AllowAnonymous]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> ManualLogin(string sesa_id, string password)
        {
            if (string.IsNullOrWhiteSpace(sesa_id) || string.IsNullOrWhiteSpace(password))
                return Json(new { success = false, message = "SESA ID and password are required." });

            try
            {
                string sesaIdNormalized = sesa_id.Trim().ToUpper();

                var user = new DatabaseAccessLayer().ValidateManualLogin(sesaIdNormalized, password);
                if (user is null)
                    return Json(new { success = false, message = "Invalid SESA ID or password." });

                var claims = new List<Claim>
        {
            new Claim(ClaimTypes.NameIdentifier, sesaIdNormalized),
            new Claim("asset_management_name", user.name  ?? ""),
            new Claim("asset_management_level", user.level),
            new Claim("asset_management_role", user.role  ?? ""),
            new Claim("asset_management_role_manage_user", user.role_manage_user.ToString()),
        };

                var claimsIdentity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
                await HttpContext.SignInAsync(
                    CookieAuthenticationDefaults.AuthenticationScheme,
                    new ClaimsPrincipal(claimsIdentity));

                return Json(new { success = true, redirectUrl = Url.Action("Open", "Home") });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Manual login error for SESA ID: {SesaId}", sesa_id);
                return Json(new { success = false, message = "An error occurred. Please try again." });
            }
        }

        [AllowAnonymous]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> SecurityLogin(string sesa_id, string password)
        {
            if (string.IsNullOrWhiteSpace(sesa_id) || string.IsNullOrWhiteSpace(password))
                return Json(new { success = false, message = "SESA ID and password are required." });

            try
            {
                string sesaIdNormalized = sesa_id.Trim().ToUpper();

                var user = new DatabaseAccessLayer().ValidateSecurityLogin(sesaIdNormalized, password);
                if (user is null)
                    return Json(new { success = false, message = "Invalid SESA ID or password." });

                var claims = new List<Claim>
                {
                    new Claim(ClaimTypes.NameIdentifier, sesaIdNormalized),
                    new Claim("asset_management_name", user.name ?? ""),
                    new Claim("asset_management_level", user.level),
                    new Claim("asset_management_role", user.role ?? ""),
                    new Claim("asset_management_role_manage_user", user.role_manage_user.ToString()),
                };

                var claimsIdentity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
                await HttpContext.SignInAsync(
                    CookieAuthenticationDefaults.AuthenticationScheme,
                    new ClaimsPrincipal(claimsIdentity));

                return Json(new { success = true, redirectUrl = Url.Action("Gatepass", "Security") });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Security manual login error for SESA ID: {SesaId}", sesa_id);
                return Json(new { success = false, message = "An error occurred. Please try again." });
            }
        }

        public IActionResult Dashboard()
        {
            return View();
        }

        public IActionResult DashboardNonAsset()
        {
            return View();
        }

        public IActionResult GetNoTaggingByClass()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            var chartData = db.GetNoTaggingByClass();
            return chartData;
        }
        public IActionResult GetNoTaggingByDepartment()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            var chartData = db.GetNoTaggingByDepartment();
            return chartData;
        }
        public IActionResult GetNoTaggingByDepartmentDetail()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<TableModel> dataList = db.GetNoTaggingByDepartmentDetail();
            return PartialView("_TableAssetNoTaggingDetail", dataList);
        }
        public IActionResult GetAssetNoTaggingByClass(string class_desc)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<AssetListModel> dataList = db.GetAssetNoTaggingByClass(class_desc);
            return PartialView("_TableAssetNoTagging", dataList);
        }
        public IActionResult GetAssetNoTaggingByDepartment(string department)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<AssetListModel> dataList = db.GetAssetNoTaggingByDepartment(department);
            return PartialView("_TableAssetNoTagging", dataList);
        }
        public IActionResult GetTaggingRatio()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            var chartData = db.GetTaggingRatio();
            return chartData;
        }
        [HttpGet]
        public IActionResult GetCountYearRange()
        {
            try
            {
                var years = _context.v_asset_count
                    .Where(x => x.count_year != null && x.count_year != "")
                    .Select(x => x.count_year)
                    .Distinct()
                    .ToList();

                if (years == null || !years.Any())
                {
                    int currentYear = DateTime.Now.Year;
                    return Json(new { startYear = currentYear, endYear = currentYear });
                }

                var validYears = years
                    .Where(y => int.TryParse(y, out _))
                    .Select(y => int.Parse(y))
                    .ToList();

                if (!validYears.Any())
                {
                    int currentYear = DateTime.Now.Year;
                    return Json(new { startYear = currentYear, endYear = currentYear });
                }

                int startYear = validYears.Min();
                int endYear = validYears.Max();

                return Json(new { startYear, endYear });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting count year range");
                int currentYear = DateTime.Now.Year;
                return Json(new { startYear = currentYear, endYear = currentYear });
            }
        }
        public IActionResult GetCountResult(string year_no, string? department)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            var chartData = db.GetCountResult(year_no, department);
            return chartData;
        }
        [HttpGet]
        public IActionResult GetCountDepartmentList()
        {
            var db = new DatabaseAccessLayer();
            var list = db.GetCountDepartment();
            return Json(list);
        }
        public IActionResult GetAssetCountDetail(string count_month, string count_year, string? department = null)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<AssetListModel> dataList = db.GetAssetCountDetail(count_month, count_year, department);

            ViewBag.department = string.IsNullOrWhiteSpace(department) ? "All Departments" : department;

            bool isYearOnly = string.IsNullOrEmpty(count_month) ||
                              (int.TryParse(count_month, out int yearCheck) && count_month.Length == 4);

            ViewBag.isYearOnly = isYearOnly;
            ViewBag.countYear = count_year;

            return PartialView("_TableAssetCountDetail", dataList);
        }
        public IActionResult GetCountByDepartment(string count_year, string month)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<TableModel> dataList = db.GetCountByDepartmentList(count_year, month);
            return Json(dataList);
        }
        public IActionResult GetAssetCountDepartmentDetail(string count_year, string department, string month)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<AssetListModel> dataList = db.GetAssetCountDepartmentDetail(count_year, department, month);
            ViewBag.department = department;
            return PartialView("_TableAssetCountDepartmentDetail", dataList);
        }
        public async Task<IActionResult> Open()
        {
            string user_level = User.FindFirst("asset_management_level")?.Value;
            if (user_level != null)
            {
                switch (user_level.ToLower())
                {
                    case "requestor":
                        return RedirectToAction("Dashboard", "Home");
                    case "approver":
                        return RedirectToAction("Dashboard", "Home");
                    case "finance":
                        return RedirectToAction("Dashboard", "Home");
                    case "security":
                        return RedirectToAction("Gatepass", "Security");
                    case "admin":
                        return RedirectToAction("UserList", "Admin");
                    default:
                        return RedirectToAction("Index", "Home");
                }
            }
            else
            {
                return RedirectToAction("Index", "Home");
            }
        }

        public async Task<IActionResult> ChangeLevel(string level = "approver", string role = "hod")
        {
            var identity = User.Identity as ClaimsIdentity;
            var existingClaim = identity?.FindFirst("asset_management_level");
            if (existingClaim != null)
            {
                identity.RemoveClaim(existingClaim);
                identity.AddClaim(new Claim("asset_management_level", level));
            }
            else
            {
                identity.AddClaim(new Claim("asset_management_level", level));
            }

            var existingClaim2 = identity?.FindFirst("asset_management_role");
            if (existingClaim2 != null)
            {
                identity.RemoveClaim(existingClaim2);
                identity.AddClaim(new Claim("asset_management_role", role));
            }
            else
            {
                identity.AddClaim(new Claim("asset_management_role", role));
            }
            var claimsPrincipal = new ClaimsPrincipal(identity);

            await HttpContext.SignInAsync(claimsPrincipal);

            return Ok(new
            {
                level = User.FindFirst("asset_management_level")?.Value,
                role = User.FindFirst("asset_management_role")?.Value
            });
        }
        public IActionResult Unauthorize()
        {
            return View();
        }
        public async Task<IActionResult> Logout()
        {
            var claimsIdentity = (ClaimsIdentity)User.Identity;

            var existName = claimsIdentity?.FindFirst("asset_management_name");
            if (existName != null)
            {
                claimsIdentity.RemoveClaim(existName);
            }
            var existLevel = claimsIdentity?.FindFirst("asset_management_level");
            if (existLevel != null)
            {
                claimsIdentity.RemoveClaim(existLevel);
            }
            var existRole = claimsIdentity?.FindFirst("asset_management_role");
            if (existRole != null)
            {
                claimsIdentity.RemoveClaim(existRole);
            }
            var existRoleManageUser = claimsIdentity?.FindFirst("asset_management_role_manage_user");
            if (existRoleManageUser != null)
            {
                claimsIdentity.RemoveClaim(existRoleManageUser);
            }
            await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(claimsIdentity));

            HttpContext.Session.Clear();
            foreach (var cookie in Request.Cookies.Keys)
            {
                Response.Cookies.Delete(cookie);
            }
            return RedirectToAction("Index", "Home");
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [HttpGet]
        public IActionResult GET_CHART_BY_AGING()
        {
            var data = new List<dynamic>();

            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                var command = new SqlCommand("GET_CHART_BY_AGING", conn)
                {
                    CommandType = CommandType.StoredProcedure
                };
                conn.Open();
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        data.Add(new
                        {
                            days = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                            count = reader.IsDBNull(1) ? 0 : reader.GetInt32(1),
                        });
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(data));
            return Json(data);
        }

        [HttpGet]
        public IActionResult GET_CHART_BY_LINE()
        {
            var data = new List<dynamic>();

            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                var command = new SqlCommand("GET_CHART_BY_LINE", conn)
                {
                    CommandType = CommandType.StoredProcedure
                };
                conn.Open();
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        data.Add(new
                        {
                            month_year = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                            count = reader.IsDBNull(1) ? 0 : reader.GetInt32(1),
                        });
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(data));
            return Json(data);
        }

        [HttpGet]
        public IActionResult GET_CHART_GATEPASS_OPEN()
        {
            var data = new List<dynamic>();

            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                var command = new SqlCommand("GET_CHART_GATEPASS_OPEN", conn)
                {
                    CommandType = CommandType.StoredProcedure
                };
                conn.Open();
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        data.Add(new
                        {
                            department = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                            total = reader.IsDBNull(1) ? 0 : reader.GetInt32(1),
                        });
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(data));
            return Json(data);
        }

        [HttpGet]
        public IActionResult GET_CHART_GATEPASS_PROGRESS()
        {
            var data = new List<dynamic>();

            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                var command = new SqlCommand("GET_CHART_GATEPASS_PROGRESS", conn)
                {
                    CommandType = CommandType.StoredProcedure
                };
                conn.Open();
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        data.Add(new
                        {
                            category = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                            not_completed = reader.IsDBNull(1) ? 0 : reader.GetInt32(1),
                            completed = reader.IsDBNull(2) ? 0 : reader.GetInt32(2),
                        });
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(data));
            return Json(data);
        }

        [HttpGet]
        public IActionResult GET_CHART_GATEPASS_NON_ASSET_OPEN()
        {
            var data = new List<dynamic>();

            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                var command = new SqlCommand("GET_CHART_GATEPASS_NON_ASSET_OPEN", conn)
                {
                    CommandType = CommandType.StoredProcedure
                };
                conn.Open();
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        data.Add(new
                        {
                            department = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                            total = reader.IsDBNull(1) ? 0 : reader.GetInt32(1),
                        });
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(data));
            return Json(data);
        }

        [HttpGet]
        public IActionResult GET_CHART_GATEPASS_NON_ASSET_PROGRESS()
        {
            var data = new List<dynamic>();

            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                var command = new SqlCommand("GET_CHART_GATEPASS_NON_ASSET_PROGRESS", conn)
                {
                    CommandType = CommandType.StoredProcedure
                };
                conn.Open();
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        data.Add(new
                        {
                            category = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                            not_completed = reader.IsDBNull(1) ? 0 : reader.GetInt32(1),
                            completed = reader.IsDBNull(2) ? 0 : reader.GetInt32(2),
                        });
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(data));
            return Json(data);
        }

        [HttpGet]
        public IActionResult CHART_TOTAL_GATEPASS()
        {
            var data = new List<dynamic>();

            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                var command = new SqlCommand("CHART_TOTAL_GATEPASS", conn)
                {
                    CommandType = CommandType.StoredProcedure
                };
                conn.Open();
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        data.Add(new
                        {
                            total = reader.IsDBNull(0) ? 0 : reader.GetInt32(0),
                            total_return = reader.IsDBNull(1) ? 0 : reader.GetInt32(1),
                            not_return = reader.IsDBNull(2) ? 0 : reader.GetInt32(2),
                        });
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(data));
            return Json(data);
        }

        [HttpGet]
        public IActionResult CHART_GATEPASS_RETURN()
        {
            var data = new List<dynamic>();

            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                var command = new SqlCommand("CHART_GATEPASS_RETURN", conn)
                {
                    CommandType = CommandType.StoredProcedure
                };
                conn.Open();
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        data.Add(new
                        {
                            at_supplier = reader.IsDBNull(0) ? 0 : reader.GetInt32(0),
                            waiting_req = reader.IsDBNull(1) ? 0 : reader.GetInt32(1),
                            waiting_hod = reader.IsDBNull(2) ? 0 : reader.GetInt32(2),
                            waiting_fin = reader.IsDBNull(3) ? 0 : reader.GetInt32(3)
                        });
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(data));
            return Json(data);
        }

        [HttpGet]
        public IActionResult CHART_GATEPASS_AGING()
        {
            var data = new List<dynamic>();

            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                var command = new SqlCommand("CHART_GATEPASS_AGING", conn)
                {
                    CommandType = CommandType.StoredProcedure
                };
                conn.Open();
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        data.Add(new
                        {
                            dept = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                            less30days = reader.IsDBNull(1) ? 0 : reader.GetInt32(1),
                            between30_60 = reader.IsDBNull(2) ? 0 : reader.GetInt32(2),
                            morethan60 = reader.IsDBNull(3) ? 0 : reader.GetInt32(3)
                        });
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(data)); 
            return Json(data);
        }

        [HttpGet]
        public IActionResult CHART_TOTAL_GATEPASS_NON_ASSET()
        {
            var data = new List<dynamic>();

            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                var command = new SqlCommand("CHART_TOTAL_GATEPASS_NON_ASSET", conn)
                {
                    CommandType = CommandType.StoredProcedure
                };
                conn.Open();
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        data.Add(new
                        {
                            total = reader.IsDBNull(0) ? 0 : reader.GetInt32(0),
                            total_return = reader.IsDBNull(1) ? 0 : reader.GetInt32(1),
                            not_return = reader.IsDBNull(2) ? 0 : reader.GetInt32(2),
                        });
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(data));
            return Json(data);
        }

        [HttpGet]
        public IActionResult CHART_GATEPASS_NON_ASSET_RETURN()
        {
            var data = new List<dynamic>();

            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                var command = new SqlCommand("CHART_GATEPASS_NON_ASSET_RETURN", conn)
                {
                    CommandType = CommandType.StoredProcedure
                };
                conn.Open();
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        data.Add(new
                        {
                            at_supplier = reader.IsDBNull(0) ? 0 : reader.GetInt32(0),
                            waiting_req = reader.IsDBNull(1) ? 0 : reader.GetInt32(1),
                            waiting_hod = reader.IsDBNull(2) ? 0 : reader.GetInt32(2),
                            waiting_fin = reader.IsDBNull(3) ? 0 : reader.GetInt32(3)
                        });
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(data));
            return Json(data);
        }

        [HttpGet]
        public IActionResult CHART_GATEPASS_NON_ASSET_AGING()
        {
            var data = new List<dynamic>();

            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                var command = new SqlCommand("CHART_GATEPASS_NON_ASSET_AGING", conn)
                {
                    CommandType = CommandType.StoredProcedure
                };
                conn.Open();
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        data.Add(new
                        {
                            dept = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                            less30days = reader.IsDBNull(1) ? 0 : reader.GetInt32(1),
                            between30_60 = reader.IsDBNull(2) ? 0 : reader.GetInt32(2),
                            morethan60 = reader.IsDBNull(3) ? 0 : reader.GetInt32(3)
                        });
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(data));
            return Json(data);
        }

        [HttpGet]
        public IActionResult GET_CHART_AUC_DEPT()
        {
            var data = new List<dynamic>();

            using (SqlConnection conn = new SqlConnection(DbConnection()))
            {
                var command = new SqlCommand("GET_CHART_AUC_DEPT", conn)
                {
                    CommandType = CommandType.StoredProcedure
                };
                conn.Open();
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        data.Add(new
                        {
                            department = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                            asset_count = reader.IsDBNull(1) ? 0 : reader.GetInt32(1),
                        });
                    }
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(data)); 
            return Json(data);
        }

        public IActionResult GP_OPEN_DETAIL(string department)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<GPOpenModel> dataList = db.GP_OPEN_DETAIL(department);
            ViewBag.department = department;
            return PartialView("_TableGPOpenDetail", dataList);
        }

        public IActionResult GP_PROGRESS_DETAIL(string status, string category)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<GPOpenModel> dataList = db.GP_PROGRESS_DETAIL(status, category);
            ViewBag.category = category;
            ViewBag.status = status;
            return PartialView("_TableGPProgressDetail", dataList);
        }

        public IActionResult GP_OPEN_NON_ASSET_DETAIL(string department)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<GPOpenModel> dataList = db.GP_OPEN_NON_ASSET_DETAIL(department);
            ViewBag.department = department;
            return PartialView("_TableGPOpenNonAssetDetail", dataList);
        }

        public IActionResult GP_PROGRESS_NON_ASSET_DETAIL(string status, string category)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<GPOpenModel> dataList = db.GP_PROGRESS_NON_ASSET_DETAIL(status, category);
            ViewBag.category = category;
            ViewBag.status = status;
            return PartialView("_TableGPProgressNonAssetDetail", dataList);
        }

        public IActionResult GET_GP_DETAIL_AGING(string department, string aging_days)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<GPModel> dataList = db.GET_GP_DETAIL_AGING(department, aging_days);
            ViewBag.department = department;
            ViewBag.aging_days = aging_days;
            return PartialView("_TableGPDetailAging", dataList);
        }

        public IActionResult GET_GP_DETAIL_RETURN(string type)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<GPModel> dataList = db.GET_GP_DETAIL_RETURN(type);
            ViewBag.type = type;
            return PartialView("_TableGPDetailReturn", dataList);
        }

        public IActionResult GET_GP_DETAIL_TOTAL(string type)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<GPModel> dataList = db.GET_GP_DETAIL_TOTAL(type);
            ViewBag.type = type;
            return PartialView("_TableGPDetailTotal", dataList);
        }

        public IActionResult GET_GP_NA_DETAIL_AGING(string department, string aging_days)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<GPModel> dataList = db.GET_GP_NA_DETAIL_AGING(department, aging_days);
            ViewBag.department = department;
            ViewBag.aging_days = aging_days;
            return PartialView("_TableGPNADetailAging", dataList);
        }

        public IActionResult GET_GP_NA_DETAIL_RETURN(string type)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<GPModel> dataList = db.GET_GP_NA_DETAIL_RETURN(type);
            ViewBag.type = type;
            return PartialView("_TableGPNADetailReturn", dataList);
        }

        public IActionResult GET_GP_NA_DETAIL_TOTAL(string type)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            var db = new DatabaseAccessLayer();
            List<GPModel> dataList = db.GET_GP_NA_DETAIL_TOTAL(type);
            ViewBag.type = type;
            return PartialView("_TableGPNADetailTotal", dataList);
        }

    }
}
