using Asset_Management.Function;
using Asset_Management.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data.SqlClient;

namespace Asset_Management.Controllers
{
    public class UserController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
        public IActionResult ChangeNewPass(string sesa_id, string old_pass, string new_pass)
        {
            string id_login = HttpContext.Session.GetString("id_login") ?? "";
            string plant = HttpContext.Session.GetString("plant") ?? "";
            if (sesa_id == "")
            {
                return Content("Session Timeout, Please relogin!!", "text/plain");
            }
            else
            {
                var db = new DatabaseAccessLayer();

                string checkOldPass = db.CheckOldPass(sesa_id, old_pass);
                if (checkOldPass == "failed")
                {
                    return Content("Wrong old password", "text/plain");
                }

                string status_msg = db.SaveNewPass(sesa_id, new_pass);

                return Content(status_msg, "text/plain");
            }
        }

    }
}
