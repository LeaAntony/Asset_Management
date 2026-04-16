using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Asset_Management.Models;
using System.Diagnostics;
using System.Security.Claims;

namespace Asset_Management.Controllers
{
    [Authorize]
    public class AccountController : Controller
    {
        public IActionResult AccessDenied() {
            return Redirect("/"); 
        }
    }
}
