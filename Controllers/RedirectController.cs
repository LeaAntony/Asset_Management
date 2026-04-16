using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.Security.Claims;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;

namespace Asset_Management.Controllers
{
    public class RedirectController : Controller
    {
        public IActionResult GatePassList(string gatepass_no = null)
        {
            string level = User.FindFirst("asset_management_level")?.Value;
            string role = User.FindFirst("asset_management_role")?.Value;

            if (!User.Identity.IsAuthenticated || level?.ToLower() == null)
            {
                var returnUrl = Request.Path + Request.QueryString;
                var encodedReturnUrl = returnUrl.Replace("&", "[AND]");
                return Challenge(new AuthenticationProperties
                {
                    RedirectUri = Url.Action("Index", "Auth", new { originalPath = encodedReturnUrl })
                });
            }

            if (level?.ToLower() == "requestor")
            {
                if (role?.ToLower() == "shipping")
                {
                    return RedirectToAction("ShippingList", "Shipping", new { gatepass_no = gatepass_no });
                }
                else
                {
                    return RedirectToAction("GatePassList", "Requestor", new { gatepass_no = gatepass_no });
                }
            }
            else if (level?.ToLower() == "approver")
            {
                return RedirectToAction("GatePassList", "Approver", new { gatepass_no = gatepass_no });
            }
            else if (level?.ToLower() == "finance")
            {
                return RedirectToAction("Gatepass", "Finance", new { gatepass_no = gatepass_no });
            }
            else if (level?.ToLower() == "security")
            {
                return RedirectToAction("Gatepass", "Security", new { gatepass_no = gatepass_no });
            }

            return RedirectToAction("Index", "Home");
        }

        public IActionResult GatePassNonAssetList(string gatepass_no = null)
        {
            string level = User.FindFirst("asset_management_level")?.Value;
            string role = User.FindFirst("asset_management_role")?.Value;

            if (!User.Identity.IsAuthenticated || level?.ToLower() == null)
            {
                var returnUrl = Request.Path + Request.QueryString;
                var encodedReturnUrl = returnUrl.Replace("&", "[AND]");
                return Challenge(new AuthenticationProperties
                {
                    RedirectUri = Url.Action("Index", "Auth", new { originalPath = encodedReturnUrl })
                });
            }

            if (level?.ToLower() == "requestor")
            {
                if (role?.ToLower() == "shipping")
                {
                    return RedirectToAction("ShippingNonAssetList", "Shipping", new { gatepass_no = gatepass_no });
                }
                else
                {
                    return RedirectToAction("GatePassNonAssetList", "Requestor", new { gatepass_no = gatepass_no });
                }
            }
            else if (level?.ToLower() == "approver")
            {
                return RedirectToAction("GatePassNonAssetList", "Approver", new { gatepass_no = gatepass_no });
            }
            else if (level?.ToLower() == "finance")
            {
                return RedirectToAction("GatepassNonAsset", "Finance", new { gatepass_no = gatepass_no });
            }
            else if (level?.ToLower() == "security")
            {
                return RedirectToAction("GatepassNonAsset", "Security", new { gatepass_no = gatepass_no });
            }

            return RedirectToAction("Index", "Home");
        }

        public IActionResult DisposalList(string order_no = null)
        {
            string level = User.FindFirst("asset_management_level")?.Value;
            string role = User.FindFirst("asset_management_role")?.Value;

            if (!User.Identity.IsAuthenticated || level?.ToLower() == null)
            {
                var returnUrl = Request.Path + Request.QueryString;
                var encodedReturnUrl = returnUrl.Replace("&", "[AND]");
                return Challenge(new AuthenticationProperties
                {
                    RedirectUri = Url.Action("Index", "Auth", new { originalPath = encodedReturnUrl })
                });
            }

            if (level?.ToLower() == "requestor")
            {
                return RedirectToAction("DisposalList", "Requestor", new { order_no = order_no });
            }
            else if (level?.ToLower() == "approver")
            {
                return RedirectToAction("DisposalList", "Approver", new { order_no = order_no });
            }
            else if (level?.ToLower() == "finance")
            {
                return RedirectToAction("DisposalList", "Finance", new { order_no = order_no });
            }

            return RedirectToAction("Index", "Home");
        }

        public IActionResult InvoiceList(string invoice_no = null)
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            string level = User.FindFirst("asset_management_level")?.Value;

            if (!User.Identity.IsAuthenticated || level?.ToLower() == null)
            {
                var returnUrl = Request.Path + Request.QueryString;
                var encodedReturnUrl = returnUrl.Replace("&", "[AND]");
                return Challenge(new AuthenticationProperties
                {
                    RedirectUri = Url.Action("Index", "Auth", new { originalPath = encodedReturnUrl })
                });
            }

            if (sesa_id?.ToLower() == "SESA100001")
            {
                return RedirectToAction("InvoiceApprovalList", "Finance", new { invoice_no = invoice_no });
            }
            else if (level?.ToLower() == "finance")
            {
                return RedirectToAction("InvoiceList", "Finance", new { invoice_no = invoice_no });
            }

            return RedirectToAction("Index", "Home");
        }
    }
}