using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using Asset_Management.Function;
using Asset_Management.Models;
using System.Linq.Dynamic.Core;
using System.Security.Claims;

namespace Asset_Management.Controllers
{
    [Authorize(Policy = "RequireAny")]
    public class ShippingController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly ILogger<ShippingController> _logger;
        private readonly IWebHostEnvironment _webHostEnvironment;

        public ShippingController(ILogger<ShippingController> logger, ApplicationDbContext context, IWebHostEnvironment webHostEnvironment)
        {
            this._context = context;
            _logger = logger;
            _webHostEnvironment = webHostEnvironment;
        }

        private string DbConnection()
        {
            var dbAccess = new DatabaseAccessLayer();
            string dbString = dbAccess.ConnectionString;
            return dbString;
        }

        public IActionResult GetShippingListBadge()
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
                if (user.role != "shipping")
                {
                    return Ok(new { count = 0 });
                }
                var count = (from ShippingList in _context.v_gatepass
                             where ShippingList.status_gatepass == "Waiting Shipping Process"
                             select ShippingList).Count();
                return Ok(new { count = count });
            }
            catch (Exception)
            {
                return Ok(new { count = 0 });
            }
        }

        public IActionResult GetShippingNonAssetListBadge()
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
                if (user.role != "shipping")
                {
                    return Ok(new { count = 0 });
                }
                var count = (from ShippingNonAssetList in _context.v_gatepass_non_asset
                             where ShippingNonAssetList.status_gatepass == "Waiting Shipping Process"
                             select ShippingNonAssetList).Count();
                return Ok(new { count = count });
            }
            catch (Exception)
            {
                return Ok(new { count = 0 });
            }
        }

        public IActionResult ShippingList(string gatepass_no = null)
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

        public IActionResult GET_SHIPPING_LIST()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            string role = User.FindFirst("asset_management_role")?.Value;
            try
            {
                var status = Request.Form["status"].FirstOrDefault() ?? "OPEN";
                var draw = Request.Form["draw"].FirstOrDefault();
                var start = Request.Form["start"].FirstOrDefault();
                var length = Request.Form["length"].FirstOrDefault();
                var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][data]"].FirstOrDefault();
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                var searchValue = Request.Form["search[value]"].FirstOrDefault();
                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;

                var mstData = from gatepass in _context.v_gatepass
                              where status == "OPEN" ? gatepass.shipping_status == "OPEN" : gatepass.shipping_status == "CLOSED"
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
                                  gatepass.shipping_status,
                                  gatepass.proforma_fin_status,
                                  gatepass.proforma_no
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
                _logger.LogError(ex, "Error in GET_SHIPPING_LIST: {Message}", ex.Message);
                return StatusCode(500, "Internal server error. Please try again later.");
            }
        }

        [HttpGet]
        public IActionResult GET_IMG_GATEPASS(string id_gatepass)
        {
            try
            {
                var db = new DatabaseAccessLayer();
                var images = db.GET_IMG_GATEPASS(id_gatepass);
                return Ok(images);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting gatepass images");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpGet]
        public IActionResult GetShippingAssets(int id_gatepass)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                var assets = db.GetShippingAssets(id_gatepass, sesa_id);
                return Ok(assets);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting shipping assets");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpPost]
        public IActionResult CreateTempBox([FromBody] TempShippingBoxModel box)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrEmpty(sesa_id))
                {
                    return BadRequest(new { success = false, message = "User not authenticated" });
                }

                box.created_by = sesa_id;

                var db = new DatabaseAccessLayer();
                var result = db.CreateTempBox(box);

                if (result.StartsWith("success"))
                {
                    return Ok(new { success = true, message = "Box created successfully", id_temp = result.Split(';')[1] });
                }

                return BadRequest(new { success = false, message = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating temp box");
                return StatusCode(500, new { success = false, message = "Internal server error: " + ex.Message });
            }
        }

        [HttpGet]
        public IActionResult GetTempBoxes(int id_gatepass)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                var boxes = db.GetTempBoxes(sesa_id, id_gatepass);
                return Ok(boxes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting temp boxes");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpPost]
        public IActionResult DeleteTempBox(int id_temp)
        {
            try
            {
                var db = new DatabaseAccessLayer();
                var result = db.DeleteTempBox(id_temp);

                if (result == "success")
                {
                    return Ok(new { success = true, message = "Box deleted successfully" });
                }

                return BadRequest(new { success = false, message = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting temp box");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpPost]
        public IActionResult UpdateHSCode(int id_detail, string asset_no, string asset_subnumber, string hs_code)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                var result = db.UpdateShippingAssetHSCode(id_detail, asset_no, asset_subnumber, hs_code);

                if (result == "success")
                {
                    return Ok(new { success = true, message = "HS Code updated successfully" });
                }

                return BadRequest(new { success = false, message = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating HS Code");
                return StatusCode(500, new { success = false, message = "Internal server error" });
            }
        }

        [HttpPost]
        public IActionResult SaveShipping([FromBody] ShippingCreateViewModel model)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                var result = db.SaveShipping(model, sesa_id);

                var parts = result.Split(';');
                string status = parts[0];

                if (status == "success")
                {
                    return Ok(new
                    {
                        success = true,
                        message = "Shipping data saved successfully",
                        id_shipping = parts[1],
                        proforma_no = parts[2]
                    });
                }
                else if (status == "plant_changed_with_boxes")
                {
                    return Ok(new
                    {
                        success = false,
                        requireConfirmation = true,
                        message = parts[1],
                        type = "plant_changed"
                    });
                }
                else
                {
                    return BadRequest(new { success = false, message = parts[1] });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving shipping");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpPost]
        public IActionResult SaveShippingForce([FromBody] SaveShippingForceRequest request)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                var result = db.SaveShippingForce(request.Model, sesa_id, request.ForceDelete);

                if (result.StartsWith("success"))
                {
                    var parts = result.Split(';');
                    return Ok(new
                    {
                        success = true,
                        message = "Shipping data saved successfully",
                        id_shipping = parts[1],
                        proforma_no = parts[2]
                    });
                }

                return BadRequest(new { success = false, message = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error force saving shipping");
                return StatusCode(500, "Internal server error");
            }
        }

        public class SaveShippingForceRequest
        {
            public ShippingCreateViewModel Model { get; set; }
            public bool ForceDelete { get; set; }
        }

        [HttpPost]
        public IActionResult SubmitShipping(int id_gatepass)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                var result = db.SubmitShipping(id_gatepass, sesa_id);

                if (result == "success")
                {
                    return Ok(new { success = true, message = "Shipping submitted successfully" });
                }

                return BadRequest(new { success = false, message = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error submitting shipping");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpGet]
        public IActionResult GetShippingData(int id_gatepass)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                var shippingData = db.GetShippingData(id_gatepass, sesa_id);
                bool isDataCompleteInDb = false;
                if (shippingData.is_shipping_saved)
                {
                    isDataCompleteInDb = !string.IsNullOrEmpty(shippingData.plant) &&
                                       shippingData.shipment_date.HasValue &&
                                       !string.IsNullOrEmpty(shippingData.shipment_type) &&
                                       !string.IsNullOrEmpty(shippingData.dhl_awb);
                }

                var response = new
                {
                    id_gatepass = shippingData.id_gatepass,
                    gatepass_no = shippingData.gatepass_no,
                    shipping_plant = shippingData.shipping_plant,
                    plant = shippingData.plant,
                    shipment_date = shippingData.shipment_date,
                    shipment_type = shippingData.shipment_type,
                    dhl_awb = shippingData.dhl_awb,
                    is_shipping_saved = shippingData.is_shipping_saved,
                    is_data_complete_in_db = isDataCompleteInDb,
                    assets = shippingData.assets,
                    temp_boxes = shippingData.temp_boxes,
                    saved_boxes = shippingData.saved_boxes
                };

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting shipping data");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpPost]
        public IActionResult DeleteSavedBox(int id_box)
        {
            try
            {
                var db = new DatabaseAccessLayer();
                var result = db.DeleteSavedShippingBox(id_box);

                if (result == "success")
                {
                    return Ok(new { success = true, message = "Saved box deleted successfully" });
                }

                return BadRequest(new { success = false, message = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting saved box");
                return StatusCode(500, new { success = false, message = "Internal server error" });
            }
        }

        [HttpGet]
        public IActionResult GetCombinedBoxes(int id_gatepass)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                var combinedBoxes = db.GetCombinedShippingBoxes(id_gatepass, sesa_id);
                return Ok(combinedBoxes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting combined boxes");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpGet]
        public IActionResult ExportShippingData(int id_gatepass)
        {
            try
            {
                var db = new DatabaseAccessLayer();
                var shippingData = db.GetShippingExportData(id_gatepass);

                if (shippingData == null || string.IsNullOrEmpty(shippingData.proforma_no))
                {
                    return NotFound("Shipping data not found");
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Invoice & Packing List");

                    CreateShippingExcelLayout(worksheet, shippingData);

                    string fileName = $"Proforma_{shippingData.proforma_no}_{DateTime.Now:yyyyMMdd}.xlsx";

                    byte[] fileContents = package.GetAsByteArray();
                    return File(fileContents, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error exporting shipping data");
                return StatusCode(500, "Internal server error");
            }
        }

        private void CreateShippingExcelLayout(ExcelWorksheet worksheet, ShippingExportModel data)
        {
            var blueBg = System.Drawing.Color.LightBlue;
            var greenBg = System.Drawing.Color.PaleGreen;
            var yellowBg = System.Drawing.Color.Khaki;
            var stabiloBg = System.Drawing.Color.Yellow;

            string logoPath = Path.Combine(_webHostEnvironment.WebRootPath, "images", "logo", "aml_black.png");
            if (System.IO.File.Exists(logoPath))
            {
                var picture = worksheet.Drawings.AddPicture("SELogo", new FileInfo(logoPath));
                picture.SetPosition(0, 0, 0, 0);
                picture.SetSize(45, 45);
            }
            else
            {
                worksheet.Cells["A1"].Value = "Asset Management";
                worksheet.Cells["A1"].Style.Font.Bold = true;
                worksheet.Cells["A1"].Style.Font.Size = 14;
                worksheet.Cells["A1"].Style.Font.Color.SetColor(System.Drawing.Color.Green);
            }

            if (data.plant == "JHO")
            {
                worksheet.Cells["A3"].Value = "IDN - JAKARTA HEAD OFFICE";
                worksheet.Cells["A4"].Value = "Jl. Gatot Subroto No. 52, Jakarta Selatan";
                worksheet.Cells["A5"].Value = "Jakarta 12930, Indonesia";
                worksheet.Cells["A6"].Value = "Tel. (62) 21 7919 7919, Fax. (62) 21 7919 7920";
            }
            else if (data.plant == "BMS 1")
            {
                worksheet.Cells["A3"].Value = "IDN - BATAM MANUFACTURING SITE 1";
                worksheet.Cells["A4"].Value = "Batamindo Industrial Park, Lot 101, Mukakuning";
                worksheet.Cells["A5"].Value = "Batam Island, Indonesia - 29433";
                worksheet.Cells["A6"].Value = "Tel. (62) 770 611101, Fax. (62) 770 611102";
            }
            else if (data.plant == "BMS 2")
            {
                worksheet.Cells["A3"].Value = "IDN - BATAM MANUFACTURING SITE 2";
                worksheet.Cells["A4"].Value = "Batamindo Industrial Park, Lot 205, Mukakuning";
                worksheet.Cells["A5"].Value = "Batam Island, Indonesia - 29433";
                worksheet.Cells["A6"].Value = "Tel. (62) 770 611205, Fax. (62) 770 611206";
            }
            else if (data.plant == "SDH")
            {
                worksheet.Cells["A3"].Value = "IDN - SURABAYA DISTRIBUTION HUB";
                worksheet.Cells["A4"].Value = "Jl. Raya Surabaya-Sidoarjo, Gedangan";
                worksheet.Cells["A5"].Value = "Surabaya, Indonesia - 61254";
                worksheet.Cells["A6"].Value = "Tel. (62) 31 7453 7453, Fax. (62) 31 7453 7454";
            }
            else if (data.plant == "BRC")
            {
                worksheet.Cells["A3"].Value = "IDN - BANDUNG R&D CENTER";
                worksheet.Cells["A4"].Value = "Jl. Benda No. 1A, Kemang, Jakarta Selatan";
                worksheet.Cells["A5"].Value = "Jakarta 12560, Indonesia";
                worksheet.Cells["A6"].Value = "Tel. (62) 21 7247 7247, Fax. (62) 21 7247 7248";
            }
            else if (data.plant == "MRO")
            {
                worksheet.Cells["A3"].Value = "IDN - MEDAN REGIONAL OFFICE";
                worksheet.Cells["A4"].Value = "Jl. Industri Raya No. 88, Jakarta Timur";
                worksheet.Cells["A5"].Value = "Jakarta 13640, Indonesia";
                worksheet.Cells["A6"].Value = "Tel. (62) 21 8602 8602, Fax. (62) 21 8602 8603";
            }

            worksheet.Cells["A3:A6"].Style.Font.Color.SetColor(System.Drawing.Color.DarkBlue);

            worksheet.Cells["L8:M8"].Merge = true;
            worksheet.Cells["L8"].Value = "Tax ID: " + $"{data.tax}";
            worksheet.Cells["L8"].Style.Font.Bold = true;
            worksheet.Cells["L8"].Style.Font.Color.SetColor(System.Drawing.Color.DarkBlue);
            worksheet.Cells["L8"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells["L8"].Style.Fill.BackgroundColor.SetColor(blueBg);

            worksheet.Cells["A8:E8"].Merge = true;
            worksheet.Cells["A8"].Value = $"Contact name:";
            worksheet.Cells["A8"].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
            worksheet.Cells["A8"].Style.Font.Bold = true;
            worksheet.Cells["A8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

            worksheet.Cells["F8"].Value = $"{data.requestor_name}";
            worksheet.Cells["F8"].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
            worksheet.Cells["F8"].Style.Font.Bold = true;

            worksheet.Cells["F9"].Value = $"{data.requestor_plant}";
            worksheet.Cells["F9"].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
            worksheet.Cells["F9"].Style.Font.Bold = true;

            worksheet.Cells["J3:N4"].Merge = true;
            worksheet.Cells["J3"].Value = "Invoice & Packing List";
            worksheet.Cells["J3"].Style.Font.Bold = true;
            worksheet.Cells["J3"].Style.Font.Size = 16;
            worksheet.Cells["J3"].Style.Font.Color.SetColor(System.Drawing.Color.Green);
            worksheet.Cells["J3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells["J3"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            worksheet.Cells["K6:N6"].Merge = true;
            worksheet.Cells["K6"].Value = "Duplicata de Facture";
            worksheet.Cells["K6"].Style.Font.Italic = true;
            worksheet.Cells["K6"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            worksheet.Cells["P1:S1"].Merge = true;
            worksheet.Cells["P1"].Value = "Proforma Invoice";
            worksheet.Cells["P1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells["P1"].Style.Fill.BackgroundColor.SetColor(yellowBg);
            worksheet.Cells["P1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells["P1"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["P1"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["Q1"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["R1"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["S1"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["S1"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            worksheet.Cells["P2:S2"].Merge = true;
            worksheet.Cells["P2"].Value = data.proforma_no;
            worksheet.Cells["P2"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells["P2"].Style.Fill.BackgroundColor.SetColor(yellowBg);
            worksheet.Cells["P2"].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
            worksheet.Cells["P2"].Style.Font.Bold = true;
            worksheet.Cells["P2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells["P2"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["P2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["Q2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["R2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["S2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["S2"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            worksheet.Cells["O3"].Value = "Ship to:";
            worksheet.Cells["O3"].Style.Font.Bold = true;
            worksheet.Cells["O3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            worksheet.Cells["P3"].Value = data.vendor_name;
            worksheet.Cells["P3"].Style.Font.Bold = true;

            worksheet.Cells["P4"].Value = data.street;
            worksheet.Cells["P4"].Style.Font.Bold = true;

            worksheet.Cells["P5"].Value = data.city;
            worksheet.Cells["P5"].Style.Font.Bold = true;

            worksheet.Cells["P6"].Value = $"{data.country}, {data.postal_code}";
            worksheet.Cells["P6"].Style.Font.Bold = true;

            worksheet.Cells["O7"].Value = "Telp:";
            worksheet.Cells["O7"].Style.Font.Bold = true;
            worksheet.Cells["O7"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            worksheet.Cells["P7"].Value = data.phone_no;
            worksheet.Cells["P7"].Style.Font.Bold = true;

            worksheet.Cells["O8"].Value = "Attn. To:";
            worksheet.Cells["O8"].Style.Font.Bold = true;
            worksheet.Cells["O8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            worksheet.Cells["P8"].Value = data.attn_to;
            worksheet.Cells["P8"].Style.Font.Bold = true;

            int currentRow = 10;

            worksheet.Cells[currentRow, 2, currentRow, 7].Merge = true;
            worksheet.Cells[currentRow, 2].Value = "Shipment of:";

            worksheet.Cells[currentRow, 8, currentRow, 9].Merge = true;
            worksheet.Cells[currentRow, 8].Value = "Shipment Date:";

            worksheet.Cells[currentRow, 10, currentRow, 11].Merge = true;
            worksheet.Cells[currentRow, 10].Value = "Shipment Type:";

            worksheet.Cells[currentRow, 12].Value = "DHL AWB";
            worksheet.Cells[currentRow, 12].Style.Font.Color.SetColor(System.Drawing.Color.Red);
            worksheet.Cells[currentRow, 12].Style.Font.Bold = true;

            worksheet.Cells[currentRow, 13].Value = "Currency:";

            worksheet.Cells[currentRow, 14, currentRow, 15].Merge = true;
            worksheet.Cells[currentRow, 14].Value = "Freight Term:";

            var headerRange = worksheet.Cells[currentRow, 2, currentRow, 15];
            headerRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            headerRange.Style.Fill.BackgroundColor.SetColor(blueBg);
            headerRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            headerRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            headerRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            headerRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            headerRange.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            headerRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            currentRow++;

            worksheet.Cells[currentRow, 2, currentRow, 7].Merge = true;
            worksheet.Cells[currentRow, 2].Value = $"{data.total_boxes} BOX    {data.total_assets} Items";

            worksheet.Cells[currentRow, 8, currentRow, 9].Merge = true;
            worksheet.Cells[currentRow, 8].Value = data.shipment_date?.ToString("dd-MMM-yy");

            worksheet.Cells[currentRow, 10, currentRow, 11].Merge = true;
            worksheet.Cells[currentRow, 10].Value = $"{data.ship_mode}";

            worksheet.Cells[currentRow, 12].Value = data.dhl_awb;
            worksheet.Cells[currentRow, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 12].Style.Fill.BackgroundColor.SetColor(yellowBg);

            worksheet.Cells[currentRow, 13].Value = "USD";

            worksheet.Cells[currentRow, 14, currentRow, 15].Merge = true;
            worksheet.Cells[currentRow, 14].Value = $"{data.freight_charges}\n({data.incoterms})";
            worksheet.Cells[currentRow, 14].Style.WrapText = true;

            var dataRange = worksheet.Cells[currentRow, 2, currentRow, 11];
            dataRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            dataRange.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            dataRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            dataRange.Style.Fill.BackgroundColor.SetColor(greenBg);

            var dataRange2 = worksheet.Cells[currentRow, 13, currentRow, 15];
            dataRange2.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange2.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange2.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange2.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            dataRange2.Style.Fill.BackgroundColor.SetColor(greenBg);

            var dataRange3 = worksheet.Cells[currentRow, 10, currentRow, 15];
            dataRange3.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange3.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange3.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange3.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange3.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            dataRange3.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            dataRange3.Style.Font.Color.SetColor(System.Drawing.Color.Blue);
            dataRange3.Style.Font.Bold = true;

            currentRow += 2;

            worksheet.Cells[currentRow, 1, currentRow + 1, 2].Merge = true;
            worksheet.Cells[currentRow, 1].Value = "Handling Unit";
            var handlingUnitRange = worksheet.Cells[currentRow, 1, currentRow + 1, 2];
            handlingUnitRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            handlingUnitRange.Style.Fill.BackgroundColor.SetColor(greenBg);

            worksheet.Cells[currentRow, 3, currentRow, 7].Merge = true;
            worksheet.Cells[currentRow, 3].Value = "Dimension (cm)";

            worksheet.Cells[currentRow, 8, currentRow + 1, 8].Merge = true;
            worksheet.Cells[currentRow, 8].Value = "Gross Weight";
            var grossWeightRange = worksheet.Cells[currentRow, 8, currentRow + 1, 8];
            grossWeightRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            grossWeightRange.Style.Fill.BackgroundColor.SetColor(greenBg);

            worksheet.Cells[currentRow, 9, currentRow + 1, 9].Merge = true;
            worksheet.Cells[currentRow, 9].Value = "Net Weight";

            worksheet.Cells[currentRow, 10, currentRow + 1, 10].Merge = true;
            worksheet.Cells[currentRow, 10].Value = "S/N";
            var snRange = worksheet.Cells[currentRow, 10, currentRow + 1, 10];
            snRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            snRange.Style.Fill.BackgroundColor.SetColor(greenBg);

            worksheet.Cells[currentRow, 11, currentRow + 1, 11].Merge = true;
            worksheet.Cells[currentRow, 11].Value = "Customer PO. #";

            worksheet.Cells[currentRow, 12, currentRow + 1, 12].Merge = true;
            worksheet.Cells[currentRow, 12].Value = "Product";
            var productRange = worksheet.Cells[currentRow, 12, currentRow + 1, 12];
            productRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            productRange.Style.Fill.BackgroundColor.SetColor(greenBg);

            worksheet.Cells[currentRow, 13, currentRow + 1, 15].Merge = true;
            worksheet.Cells[currentRow, 13].Value = "Description";
            var descriptionRange = worksheet.Cells[currentRow, 13, currentRow + 1, 15];
            descriptionRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            descriptionRange.Style.Fill.BackgroundColor.SetColor(greenBg);

            worksheet.Cells[currentRow, 16, currentRow + 1, 16].Merge = true;
            worksheet.Cells[currentRow, 16].Value = "QTY";

            worksheet.Cells[currentRow, 17, currentRow + 1, 17].Merge = true;
            worksheet.Cells[currentRow, 17].Value = "Unit Price";
            var unitPriceRange = worksheet.Cells[currentRow, 17, currentRow + 1, 17];
            unitPriceRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            unitPriceRange.Style.Fill.BackgroundColor.SetColor(greenBg);

            worksheet.Cells[currentRow, 18, currentRow + 1, 18].Merge = true;
            worksheet.Cells[currentRow, 18].Value = "Um";

            worksheet.Cells[currentRow, 19, currentRow + 1, 19].Merge = true;
            worksheet.Cells[currentRow, 19].Value = "Extended Price";
            var extendedPriceRange = worksheet.Cells[currentRow, 19, currentRow + 1, 19];
            extendedPriceRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            extendedPriceRange.Style.Fill.BackgroundColor.SetColor(greenBg);

            worksheet.Cells[currentRow, 20, currentRow + 1, 20].Merge = true;
            worksheet.Cells[currentRow, 20].Value = "HS CODE";

            currentRow++;

            worksheet.Cells[currentRow, 3].Value = "L";
            worksheet.Cells[currentRow, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            worksheet.Cells[currentRow, 4].Value = "x";
            worksheet.Cells[currentRow, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            worksheet.Cells[currentRow, 5].Value = "W";
            worksheet.Cells[currentRow, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            worksheet.Cells[currentRow, 6].Value = "x";
            worksheet.Cells[currentRow, 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            worksheet.Cells[currentRow, 7].Value = "H";
            worksheet.Cells[currentRow, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            var allHeaderRange = worksheet.Cells[currentRow - 1, 1, currentRow, 20];
            allHeaderRange.Style.Font.Bold = true;
            allHeaderRange.Style.WrapText = true;
            allHeaderRange.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            allHeaderRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            for (int row = currentRow - 1; row <= currentRow; row++)
            {
                for (int col = 1; col <= 20; col++)
                {
                    worksheet.Cells[row, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[row, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[row, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[row, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }
            }

            currentRow++;

            decimal totalVolume = 0;
            decimal totalNetWeight = 0;
            int totalQty = 0;
            decimal totalExtendedPrice = 0;
            var boxMergeInfo = new List<(int startRow, int endRow, string boxNo)>();

            if (data.boxes != null)
            {
                foreach (var box in data.boxes)
                {
                    decimal lengthCm = box.length_cm;
                    decimal widthCm = box.width_cm;
                    decimal heightCm = box.height_cm;
                    if (lengthCm > 0 && widthCm > 0 && heightCm > 0)
                    {
                        decimal lengthM = lengthCm / 100m;
                        decimal widthM = widthCm / 100m;
                        decimal heightM = heightCm / 100m;

                        decimal boxVolume = lengthM * widthM * heightM;
                        totalVolume += boxVolume;
                    }

                    totalNetWeight += box.net_weight_kg;

                    int boxSerialNumber = 1;
                    int boxStartRow = currentRow;

                    if (box.assets != null && box.assets.Any())
                    {
                        foreach (var asset in box.assets)
                        {
                            for (int col = 1; col <= 20; col++)
                            {
                                worksheet.Cells[currentRow, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[currentRow, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[currentRow, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                            }

                            if (asset == box.assets.First())
                            {
                                worksheet.Cells[currentRow, 3].Value = (double)lengthCm;
                                worksheet.Cells[currentRow, 3].Style.Numberformat.Format = "0.0";

                                worksheet.Cells[currentRow, 4].Value = "x";
                                worksheet.Cells[currentRow, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                worksheet.Cells[currentRow, 5].Value = (double)widthCm;
                                worksheet.Cells[currentRow, 5].Style.Numberformat.Format = "0.0";

                                worksheet.Cells[currentRow, 6].Value = "x";
                                worksheet.Cells[currentRow, 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                worksheet.Cells[currentRow, 7].Value = (double)heightCm;
                                worksheet.Cells[currentRow, 7].Style.Numberformat.Format = "0.0";

                                worksheet.Cells[currentRow, 8].Value = (double)box.gross_weight_kg;
                                worksheet.Cells[currentRow, 8].Style.Numberformat.Format = "0.0";

                                worksheet.Cells[currentRow, 9].Value = (double)box.net_weight_kg;
                                worksheet.Cells[currentRow, 9].Style.Numberformat.Format = "0.0";
                            }

                            worksheet.Cells[currentRow, 10].Value = boxSerialNumber++;

                            worksheet.Cells[currentRow, 11].Value = data.invoice_payment;

                            worksheet.Cells[currentRow, 12].Value = $"{asset.asset_no}/{asset.asset_subnumber}";

                            try
                            {
                                worksheet.Cells[currentRow, 13, currentRow, 15].Merge = true;
                                worksheet.Cells[currentRow, 13].Value = asset.asset_desc;
                                worksheet.Cells[currentRow, 13, currentRow, 15].Style.WrapText = true;
                            }
                            catch (InvalidOperationException)
                            {
                                worksheet.Cells[currentRow, 13].Value = asset.asset_desc;
                                worksheet.Cells[currentRow, 13].Style.WrapText = true;
                            }

                            if (!string.IsNullOrEmpty(asset.asset_desc))
                            {
                                int charactersPerLine = 26;
                                string[] explicitLines = asset.asset_desc.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                                int totalLines = 0;

                                foreach (string line in explicitLines)
                                {
                                    if (string.IsNullOrWhiteSpace(line))
                                    {
                                        totalLines += 1;
                                    }
                                    else
                                    {
                                        int linesForThisSegment = Math.Max(1, (int)Math.Ceiling((double)line.Length / charactersPerLine));
                                        totalLines += linesForThisSegment;
                                    }
                                }

                                if (explicitLines.Length <= 1)
                                {
                                    totalLines = Math.Max(1, (int)Math.Ceiling((double)asset.asset_desc.Length / charactersPerLine));
                                }

                                double calculatedHeight = Math.Max(18, totalLines * 18);
                                worksheet.Row(currentRow).Height = calculatedHeight;
                                worksheet.Row(currentRow).CustomHeight = true;
                            }
                            else
                            {
                                worksheet.Row(currentRow).Height = 20;
                                worksheet.Row(currentRow).CustomHeight = true;
                            }

                            worksheet.Cells[currentRow, 16].Value = 1;

                            worksheet.Cells[currentRow, 17].Value = (double)asset.current_apc;
                            worksheet.Cells[currentRow, 17].Style.Numberformat.Format = "#,##0.00";

                            worksheet.Cells[currentRow, 18].Value = "PC";

                            worksheet.Cells[currentRow, 19].Value = (double)asset.current_apc;
                            worksheet.Cells[currentRow, 19].Style.Numberformat.Format = "#,##0.00";

                            worksheet.Cells[currentRow, 20].Value = asset.hs_code?.ToString() ?? "";

                            totalQty++;
                            totalExtendedPrice += asset.current_apc;

                            currentRow++;
                        }

                        int boxEndRow = currentRow - 1;
                        boxMergeInfo.Add((boxStartRow, boxEndRow, box.box_no));

                        for (int row = boxStartRow; row <= boxEndRow; row++)
                        {
                            worksheet.Cells[row, 10].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[row, 10].Style.Fill.BackgroundColor.SetColor(greenBg);
                        }

                        for (int row = boxStartRow; row <= boxEndRow; row++)
                        {
                            worksheet.Cells[row, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[row, 12].Style.Fill.BackgroundColor.SetColor(greenBg);
                        }

                        for (int row = boxStartRow; row <= boxEndRow; row++)
                        {
                            var descCells = worksheet.Cells[row, 13, row, 15];
                            descCells.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            descCells.Style.Fill.BackgroundColor.SetColor(greenBg);
                        }

                        for (int row = boxStartRow; row <= boxEndRow; row++)
                        {
                            worksheet.Cells[row, 17].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[row, 17].Style.Fill.BackgroundColor.SetColor(greenBg);

                            worksheet.Cells[row, 19].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[row, 19].Style.Fill.BackgroundColor.SetColor(greenBg);
                        }
                    }
                }
            }

            foreach (var (startRow, endRow, boxNo) in boxMergeInfo)
            {
                try
                {
                    if (startRow <= endRow)
                    {
                        worksheet.Cells[startRow, 1, endRow, 2].Merge = true;
                        worksheet.Cells[startRow, 1].Value = boxNo;
                        var handlingUnitCells = worksheet.Cells[startRow, 1, endRow, 2];
                        handlingUnitCells.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        handlingUnitCells.Style.Fill.BackgroundColor.SetColor(greenBg);

                        worksheet.Cells[startRow, 3, endRow, 3].Merge = true;
                        worksheet.Cells[startRow, 4, endRow, 4].Merge = true;
                        worksheet.Cells[startRow, 5, endRow, 5].Merge = true;
                        worksheet.Cells[startRow, 6, endRow, 6].Merge = true;
                        worksheet.Cells[startRow, 7, endRow, 7].Merge = true;

                        worksheet.Cells[startRow, 8, endRow, 8].Merge = true;
                        var grossWeightCells = worksheet.Cells[startRow, 8, endRow, 8];
                        grossWeightCells.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        grossWeightCells.Style.Fill.BackgroundColor.SetColor(greenBg);

                        worksheet.Cells[startRow, 9, endRow, 9].Merge = true;
                    }
                }
                catch (InvalidOperationException ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Merge conflict for box {boxNo} rows {startRow}-{endRow}: {ex.Message}");
                }
            }

            worksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
            worksheet.Cells[currentRow, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 1].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 10].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 10].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 12].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 13, currentRow, 15].Merge = true;
            worksheet.Cells[currentRow, 13].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 13].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 17].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 17].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 19].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 19].Style.Fill.BackgroundColor.SetColor(greenBg);

            for (int col = 1; col <= 20; col++)
            {
                worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }

            currentRow++;

            if (!string.IsNullOrEmpty(data.coo))
            {
                try
                {
                    worksheet.Cells[currentRow, 8, currentRow, 15].Merge = true;
                    worksheet.Cells[currentRow, 8].Value = $"COO: {data.coo}";
                    worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(stabiloBg);
                    worksheet.Cells[currentRow, 8].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                    worksheet.Cells[currentRow, 8].Style.Font.Bold = true;
                    worksheet.Cells[currentRow, 8].Style.WrapText = true;

                    for (int col = 8; col <= 15; col++)
                    {
                        worksheet.Cells[currentRow, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    worksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                    worksheet.Cells[currentRow, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 1].Style.Fill.BackgroundColor.SetColor(greenBg);
                    worksheet.Cells[currentRow, 17].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 17].Style.Fill.BackgroundColor.SetColor(greenBg);
                    worksheet.Cells[currentRow, 19].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 19].Style.Fill.BackgroundColor.SetColor(greenBg);

                    for (int col = 1; col <= 20; col++)
                    {
                        worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    currentRow++;
                }
                catch (InvalidOperationException)
                {
                    worksheet.Cells[currentRow, 8].Value = $"COO: {data.coo}";
                    worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(yellowBg);
                    worksheet.Cells[currentRow, 8].Style.Font.Bold = true;
                    currentRow++;
                }
            }

            if (!string.IsNullOrEmpty(data.courier_account_no) &&
                !string.IsNullOrEmpty(data.courier_charges) &&
                data.courier_charges.Equals("Courier", StringComparison.OrdinalIgnoreCase) ||
                data.courier_charges.Equals("Receiver", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    worksheet.Cells[currentRow, 8, currentRow, 15].Merge = true;
                    worksheet.Cells[currentRow, 8].Value = $"ACC: {data.courier_account_no}";
                    worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(stabiloBg);
                    worksheet.Cells[currentRow, 8].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                    worksheet.Cells[currentRow, 8].Style.Font.Bold = true;
                    worksheet.Cells[currentRow, 8].Style.WrapText = true;

                    for (int col = 8; col <= 15; col++)
                    {
                        worksheet.Cells[currentRow, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    worksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                    worksheet.Cells[currentRow, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 1].Style.Fill.BackgroundColor.SetColor(greenBg);
                    worksheet.Cells[currentRow, 17].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 17].Style.Fill.BackgroundColor.SetColor(greenBg);
                    worksheet.Cells[currentRow, 19].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 19].Style.Fill.BackgroundColor.SetColor(greenBg);

                    for (int col = 1; col <= 20; col++)
                    {
                        worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    currentRow++;
                }
                catch (InvalidOperationException)
                {
                    worksheet.Cells[currentRow, 8].Value = $"ACC: {data.courier_account_no}";
                    worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(yellowBg);
                    worksheet.Cells[currentRow, 8].Style.Font.Bold = true;
                    currentRow++;
                }
            }

            for (int col = 8; col <= 15; col++)
            {
                worksheet.Cells[currentRow - 1, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }

            if (!string.IsNullOrEmpty(data.remark))
            {
                try
                {
                    worksheet.Cells[currentRow, 8, currentRow, 15].Merge = true;
                    worksheet.Cells[currentRow, 8].Value = data.remark;
                    worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(stabiloBg);
                    worksheet.Cells[currentRow, 8].Style.Font.Bold = true;

                    for (int col = 8; col <= 15; col++)
                    {
                        worksheet.Cells[currentRow, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    worksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                    worksheet.Cells[currentRow, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 1].Style.Fill.BackgroundColor.SetColor(greenBg);
                    worksheet.Cells[currentRow, 17].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 17].Style.Fill.BackgroundColor.SetColor(greenBg);
                    worksheet.Cells[currentRow, 19].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 19].Style.Fill.BackgroundColor.SetColor(greenBg);

                    for (int col = 1; col <= 20; col++)
                    {
                        worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    currentRow++;
                }
                catch (InvalidOperationException)
                {
                    worksheet.Cells[currentRow, 8].Value = data.remark;
                    worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(yellowBg);
                    worksheet.Cells[currentRow, 8].Style.Font.Bold = true;
                    currentRow++;
                }
            }

            worksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
            worksheet.Cells[currentRow, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 1].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 10].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 10].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 12].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 13, currentRow, 15].Merge = true;
            worksheet.Cells[currentRow, 13].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 13].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 17].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 17].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 19].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 19].Style.Fill.BackgroundColor.SetColor(greenBg);

            for (int col = 1; col <= 20; col++)
            {
                worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }

            currentRow++;

            try
            {
                worksheet.Cells[currentRow, 8, currentRow, 15].Merge = true;
                worksheet.Cells[currentRow, 8].Value = "Items not for sale";
                worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(stabiloBg);
                worksheet.Cells[currentRow, 8].Style.Font.Color.SetColor(System.Drawing.Color.Green);
                worksheet.Cells[currentRow, 8].Style.Font.Bold = true;
                worksheet.Cells[currentRow, 8].Style.WrapText = true;

                for (int col = 8; col <= 15; col++)
                {
                    worksheet.Cells[currentRow, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                currentRow++;

                worksheet.Cells[currentRow, 8, currentRow, 15].Merge = true;
                worksheet.Cells[currentRow, 8].Value = "Value declare for custom purpose only";
                worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(stabiloBg);
                worksheet.Cells[currentRow, 8].Style.Font.Color.SetColor(System.Drawing.Color.Green);
                worksheet.Cells[currentRow, 8].Style.Font.Bold = true;
                worksheet.Cells[currentRow, 8].Style.WrapText = true;

                for (int col = 8; col <= 15; col++)
                {
                    worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                currentRow++;

                worksheet.Cells[currentRow, 8, currentRow, 15].Merge = true;
                worksheet.Cells[currentRow, 8].Value = "Free of Charge";
                worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(stabiloBg);
                worksheet.Cells[currentRow, 8].Style.Font.Color.SetColor(System.Drawing.Color.Green);
                worksheet.Cells[currentRow, 8].Style.Font.Bold = true;
                worksheet.Cells[currentRow, 8].Style.WrapText = true;

                for (int col = 8; col <= 15; col++)
                {
                    worksheet.Cells[currentRow, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                for (int remarkRow = currentRow - 2; remarkRow <= currentRow; remarkRow++)
                {
                    worksheet.Cells[remarkRow, 1, remarkRow, 2].Merge = true;
                    worksheet.Cells[remarkRow, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[remarkRow, 1].Style.Fill.BackgroundColor.SetColor(greenBg);
                    worksheet.Cells[remarkRow, 17].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[remarkRow, 17].Style.Fill.BackgroundColor.SetColor(greenBg);
                    worksheet.Cells[remarkRow, 19].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[remarkRow, 19].Style.Fill.BackgroundColor.SetColor(greenBg);

                    for (int col = 1; col <= 20; col++)
                    {
                        worksheet.Cells[remarkRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[remarkRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }
                }

                currentRow += 1;

                worksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                worksheet.Cells[currentRow, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 1].Style.Fill.BackgroundColor.SetColor(greenBg);
                worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(greenBg);
                worksheet.Cells[currentRow, 10].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 10].Style.Fill.BackgroundColor.SetColor(greenBg);
                worksheet.Cells[currentRow, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 12].Style.Fill.BackgroundColor.SetColor(greenBg);
                worksheet.Cells[currentRow, 13, currentRow, 15].Merge = true;
                worksheet.Cells[currentRow, 13].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 13].Style.Fill.BackgroundColor.SetColor(greenBg);
                worksheet.Cells[currentRow, 17].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 17].Style.Fill.BackgroundColor.SetColor(greenBg);
                worksheet.Cells[currentRow, 19].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 19].Style.Fill.BackgroundColor.SetColor(greenBg);

                for (int col = 1; col <= 20; col++)
                {
                    worksheet.Cells[currentRow, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                currentRow += 1;
            }
            catch (InvalidOperationException)
            {
                worksheet.Cells[currentRow, 8].Value = "Items not for sale, value declared for customs purposes only, free of charge.";
                worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(yellowBg);
                worksheet.Cells[currentRow, 8].Style.Font.Bold = true;
                currentRow++;
            }

            worksheet.Cells[currentRow, 1, currentRow, 3].Merge = true;
            worksheet.Cells[currentRow, 1].Value = "Grand Total";

            worksheet.Cells[currentRow, 4, currentRow, 6].Merge = true;
            worksheet.Cells[currentRow, 4].Value = (double)totalVolume;

            worksheet.Cells[currentRow, 7].Value = "M3";

            worksheet.Cells[currentRow, 9].Value = (double)totalNetWeight;
            worksheet.Cells[currentRow, 9].Style.Numberformat.Format = "0.0";

            worksheet.Cells[currentRow, 10].Value = "Kgs";

            worksheet.Cells[currentRow, 16].Value = totalQty;

            worksheet.Cells[currentRow, 18].Value = '$';

            worksheet.Cells[currentRow, 19].Value = (double)totalExtendedPrice;
            worksheet.Cells[currentRow, 19].Style.Numberformat.Format = "#,##0.00";

            for (int col = 1; col <= 19; col++)
            {
                worksheet.Cells[currentRow, col].Style.Font.Bold = true;
            }

            currentRow++;

            for (int col = 1; col <= 19; col++)
            {
                worksheet.Cells[currentRow, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            }

            worksheet.Column(1).Width = 1.91;
            worksheet.Column(2).Width = 6.55;
            worksheet.Column(3).AutoFit();
            worksheet.Column(4).Width = 2.91;
            worksheet.Column(5).AutoFit();
            worksheet.Column(6).Width = 2.91;
            worksheet.Column(7).AutoFit();
            worksheet.Column(8).AutoFit();
            worksheet.Column(9).AutoFit();
            worksheet.Column(10).Width = 4.55;
            worksheet.Column(11).Width = 9.18;
            worksheet.Column(12).Width = 19.36;
            worksheet.Column(13).Width = 8.31;
            worksheet.Column(14).Width = 3.18;
            worksheet.Column(15).Width = 15.45;
            worksheet.Column(16).Width = 5.55;
            worksheet.Column(17).AutoFit();
            worksheet.Column(18).Width = 4.18;
            worksheet.Column(19).AutoFit();
            worksheet.Column(20).Width = 10.82;
            worksheet.Row(11).Height = 31.5;
            worksheet.Row(currentRow).Height = 2.2;

            worksheet.PrinterSettings.Orientation = OfficeOpenXml.eOrientation.Landscape;
            worksheet.PrinterSettings.FitToPage = true;
        }

        public IActionResult ShippingNonAssetList(string gatepass_no = null)
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

        public IActionResult GET_SHIPPING_NONASSET_LIST()
        {
            string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            string role = User.FindFirst("asset_management_role")?.Value;
            try
            {
                var status = Request.Form["status"].FirstOrDefault() ?? "OPEN";
                var draw = Request.Form["draw"].FirstOrDefault();
                var start = Request.Form["start"].FirstOrDefault();
                var length = Request.Form["length"].FirstOrDefault();
                var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][data]"].FirstOrDefault();
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                var searchValue = Request.Form["search[value]"].FirstOrDefault();
                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;

                var mstData = from gatepass in _context.v_gatepass_non_asset
                              where status == "OPEN" ? gatepass.shipping_status == "OPEN" : gatepass.shipping_status == "CLOSED"
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
                                  gatepass.requestor_name,
                                  gatepass.shipping_status,
                                  gatepass.proforma_fin_status,
                                  gatepass.proforma_no
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
                _logger.LogError(ex, "Error in GET_SHIPPING_NONASSET_LIST: {Message}", ex.Message);
                return StatusCode(500, "Internal server error. Please try again later.");
            }
        }

        [HttpGet]
        public IActionResult GET_IMG_GATEPASS_NONASSET(string id_gatepass)
        {
            try
            {
                var db = new DatabaseAccessLayer();
                var images = db.GET_IMG_NON_ASSET_GATEPASS(Convert.ToInt32(id_gatepass));
                return Ok(images);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting non-asset gatepass images");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpGet]
        public IActionResult GetShippingNonAssets(int id_gatepass)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                var items = db.GET_SHIPPING_NON_ASSETS(id_gatepass, sesa_id);
                return Ok(items);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting shipping non-asset items");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpPost]
        public IActionResult CreateTempNonAssetBox([FromBody] TempShippingNonAssetBoxModel box)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (string.IsNullOrEmpty(sesa_id))
                {
                    return BadRequest(new { success = false, message = "User not authenticated" });
                }

                box.created_by = sesa_id;

                var db = new DatabaseAccessLayer();
                var result = db.CREATE_TEMP_NON_ASSET_BOX(box);

                if (result.StartsWith("success"))
                {
                    return Ok(new { success = true, message = "Box created successfully", id_temp = result.Split(';')[1] });
                }

                return BadRequest(new { success = false, message = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating temp non-asset box");
                return StatusCode(500, new { success = false, message = "Internal server error: " + ex.Message });
            }
        }

        [HttpGet]
        public IActionResult GetTempNonAssetBoxes(int id_gatepass)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                var boxes = db.GET_TEMP_NON_ASSET_BOXES(sesa_id, id_gatepass);
                return Ok(boxes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting temp non-asset boxes");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpPost]
        public IActionResult DeleteTempNonAssetBox(int id_temp)
        {
            try
            {
                var db = new DatabaseAccessLayer();
                var result = db.DELETE_TEMP_NON_ASSET_BOX(id_temp);

                if (result == "success")
                {
                    return Ok(new { success = true, message = "Box deleted successfully" });
                }

                return BadRequest(new { success = false, message = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting temp non-asset box");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpPost]
        public IActionResult UpdateNonAssetHSCode(int id_detail, string po_or_asset_no, string hs_code)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                var result = db.UPDATE_SHIPPING_NON_ASSET_ITEM_HS_CODE(id_detail, po_or_asset_no, hs_code);

                if (result == "success")
                {
                    return Ok(new { success = true, message = "HS Code updated successfully" });
                }

                return BadRequest(new { success = false, message = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating non-asset HS Code");
                return StatusCode(500, new { success = false, message = "Internal server error" });
            }
        }

        [HttpPost]
        public IActionResult SaveNonAssetShipping([FromBody] ShippingNonAssetCreateViewModel model)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                var result = db.SAVE_SHIPPING_NON_ASSET(model, sesa_id);

                var parts = result.Split(';');
                string status = parts[0];

                if (status == "success")
                {
                    return Ok(new
                    {
                        success = true,
                        message = "Non-asset shipping data saved successfully",
                        id_shipping_non_asset = parts[1],
                        proforma_no = parts[2]
                    });
                }
                else if (status == "plant_changed_with_boxes")
                {
                    return Ok(new
                    {
                        success = false,
                        requireConfirmation = true,
                        message = parts[1],
                        type = "plant_changed"
                    });
                }
                else
                {
                    return BadRequest(new { success = false, message = parts[1] });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving non-asset shipping");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpPost]
        public IActionResult SaveNonAssetShippingForce([FromBody] SaveNonAssetShippingForceRequest request)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                var result = db.SAVE_SHIPPING_NON_ASSET_FORCE(request.Model, sesa_id, request.ForceDelete);

                if (result.StartsWith("success"))
                {
                    var parts = result.Split(';');
                    return Ok(new
                    {
                        success = true,
                        message = "Non-asset shipping data saved successfully",
                        id_shipping_non_asset = parts[1],
                        proforma_no = parts[2]
                    });
                }

                return BadRequest(new { success = false, message = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error force saving non-asset shipping");
                return StatusCode(500, "Internal server error");
            }
        }

        public class SaveNonAssetShippingForceRequest
        {
            public ShippingNonAssetCreateViewModel Model { get; set; }
            public bool ForceDelete { get; set; }
        }

        [HttpPost]
        public IActionResult SubmitNonAssetShipping(int id_gatepass)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                var result = db.SUBMIT_SHIPPING_NON_ASSET(id_gatepass, sesa_id);

                if (result == "success")
                {
                    return Ok(new { success = true, message = "Non-asset shipping submitted successfully" });
                }

                return BadRequest(new { success = false, message = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error submitting non-asset shipping");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpGet]
        public IActionResult GetNonAssetShippingData(int id_gatepass)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                var shippingData = db.GET_SHIPPING_NON_ASSET_DATA(id_gatepass, sesa_id);
                bool isDataCompleteInDb = false;
                if (shippingData.is_shipping_saved)
                {
                    isDataCompleteInDb = !string.IsNullOrEmpty(shippingData.plant) &&
                                       shippingData.shipment_date.HasValue &&
                                       !string.IsNullOrEmpty(shippingData.shipment_type) &&
                                       !string.IsNullOrEmpty(shippingData.dhl_awb);
                }

                var response = new
                {
                    id_gatepass = shippingData.id_gatepass,
                    shipping_plant = shippingData.shipping_plant,
                    gatepass_no = shippingData.gatepass_no,
                    plant = shippingData.plant,
                    shipment_date = shippingData.shipment_date,
                    shipment_type = shippingData.shipment_type,
                    dhl_awb = shippingData.dhl_awb,
                    is_shipping_saved = shippingData.is_shipping_saved,
                    is_data_complete_in_db = isDataCompleteInDb,
                    items = shippingData.items,
                    temp_boxes = shippingData.temp_boxes,
                    saved_boxes = shippingData.saved_boxes
                };

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting non-asset shipping data");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpPost]
        public IActionResult DeleteSavedNonAssetBox(int id_box)
        {
            try
            {
                var db = new DatabaseAccessLayer();
                var result = db.DELETE_SAVED_NON_ASSET_SHIPPING_BOX(id_box);

                if (result == "success")
                {
                    return Ok(new { success = true, message = "Saved box deleted successfully" });
                }

                return BadRequest(new { success = false, message = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting saved non-asset box");
                return StatusCode(500, new { success = false, message = "Internal server error" });
            }
        }

        [HttpGet]
        public IActionResult GetCombinedNonAssetBoxes(int id_gatepass)
        {
            try
            {
                string sesa_id = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var db = new DatabaseAccessLayer();
                var combinedBoxes = new ShippingNonAssetBoxDisplayModel
                {
                    saved_boxes = db.GET_SAVED_NON_ASSET_SHIPPING_BOXES(id_gatepass),
                    temp_boxes = db.GET_TEMP_NON_ASSET_BOXES(sesa_id, id_gatepass)
                };
                return Ok(combinedBoxes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting combined non-asset boxes");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpGet]
        public IActionResult ExportNonAssetShippingData(int id_gatepass)
        {
            try
            {
                var db = new DatabaseAccessLayer();
                var shippingData = db.GET_SHIPPING_NON_ASSET_EXPORT_DATA(id_gatepass);

                if (shippingData == null || string.IsNullOrEmpty(shippingData.proforma_no))
                {
                    return NotFound("Non-asset shipping data not found");
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Invoice & Packing List");

                    CreateNonAssetShippingExcelLayout(worksheet, shippingData);

                    string fileName = $"Proforma_NonAsset_{shippingData.proforma_no}_{DateTime.Now:yyyyMMdd}.xlsx";

                    byte[] fileContents = package.GetAsByteArray();
                    return File(fileContents, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error exporting non-asset shipping data");
                return StatusCode(500, "Internal server error");
            }
        }

        private void CreateNonAssetShippingExcelLayout(ExcelWorksheet worksheet, ShippingNonAssetExportModel data)
        {
            var blueBg = System.Drawing.Color.LightBlue;
            var greenBg = System.Drawing.Color.PaleGreen;
            var yellowBg = System.Drawing.Color.Khaki;
            var stabiloBg = System.Drawing.Color.Yellow;

            string logoPath = Path.Combine(_webHostEnvironment.WebRootPath, "images", "logo", "aml_black.png");
            if (System.IO.File.Exists(logoPath))
            {
                var picture = worksheet.Drawings.AddPicture("SELogo", new FileInfo(logoPath));
                picture.SetPosition(0, 0, 0, 0);
                picture.SetSize(150, 40);
            }
            else
            {
                worksheet.Cells["A1"].Value = "Asset Management";
                worksheet.Cells["A1"].Style.Font.Bold = true;
                worksheet.Cells["A1"].Style.Font.Size = 14;
                worksheet.Cells["A1"].Style.Font.Color.SetColor(System.Drawing.Color.Green);
            }

            if (data.plant == "JHO")
            {
                worksheet.Cells["A3"].Value = "IDN - JAKARTA HEAD OFFICE";
                worksheet.Cells["A4"].Value = "Jl. Gatot Subroto No. 52, Jakarta Selatan";
                worksheet.Cells["A5"].Value = "Jakarta 12930, Indonesia";
                worksheet.Cells["A6"].Value = "Tel. (62) 21 7919 7919, Fax. (62) 21 7919 7920";
            }
            else if (data.plant == "BMS 1")
            {
                worksheet.Cells["A3"].Value = "IDN - BATAM MANUFACTURING SITE 1";
                worksheet.Cells["A4"].Value = "Batamindo Industrial Park, Lot 101, Mukakuning";
                worksheet.Cells["A5"].Value = "Batam Island, Indonesia - 29433";
                worksheet.Cells["A6"].Value = "Tel. (62) 770 611101, Fax. (62) 770 611102";
            }
            else if (data.plant == "BMS 2")
            {
                worksheet.Cells["A3"].Value = "IDN - BATAM MANUFACTURING SITE 2";
                worksheet.Cells["A4"].Value = "Batamindo Industrial Park, Lot 205, Mukakuning";
                worksheet.Cells["A5"].Value = "Batam Island, Indonesia - 29433";
                worksheet.Cells["A6"].Value = "Tel. (62) 770 611205, Fax. (62) 770 611206";
            }
            else if (data.plant == "SDH")
            {
                worksheet.Cells["A3"].Value = "IDN - SURABAYA DISTRIBUTION HUB";
                worksheet.Cells["A4"].Value = "Jl. Raya Surabaya-Sidoarjo, Gedangan";
                worksheet.Cells["A5"].Value = "Surabaya, Indonesia - 61254";
                worksheet.Cells["A6"].Value = "Tel. (62) 31 7453 7453, Fax. (62) 31 7453 7454";
            }
            else if (data.plant == "BRC")
            {
                worksheet.Cells["A3"].Value = "IDN - BANDUNG R&D CENTER";
                worksheet.Cells["A4"].Value = "Jl. Benda No. 1A, Kemang, Jakarta Selatan";
                worksheet.Cells["A5"].Value = "Jakarta 12560, Indonesia";
                worksheet.Cells["A6"].Value = "Tel. (62) 21 7247 7247, Fax. (62) 21 7247 7248";
            }
            else if (data.plant == "MRO")
            {
                worksheet.Cells["A3"].Value = "IDN - MEDAN REGIONAL OFFICE";
                worksheet.Cells["A4"].Value = "Jl. Industri Raya No. 88, Jakarta Timur";
                worksheet.Cells["A5"].Value = "Jakarta 13640, Indonesia";
                worksheet.Cells["A6"].Value = "Tel. (62) 21 8602 8602, Fax. (62) 21 8602 8603";
            }

            worksheet.Cells["A3:A6"].Style.Font.Color.SetColor(System.Drawing.Color.DarkBlue);

            worksheet.Cells["L8:M8"].Merge = true;
            worksheet.Cells["L8"].Value = "Tax ID: " + $"{data.tax}";
            worksheet.Cells["L8"].Style.Font.Bold = true;
            worksheet.Cells["L8"].Style.Font.Color.SetColor(System.Drawing.Color.DarkBlue);
            worksheet.Cells["L8"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells["L8"].Style.Fill.BackgroundColor.SetColor(blueBg);

            worksheet.Cells["A8:E8"].Merge = true;
            worksheet.Cells["A8"].Value = $"Contact name:";
            worksheet.Cells["A8"].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
            worksheet.Cells["A8"].Style.Font.Bold = true;
            worksheet.Cells["A8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

            worksheet.Cells["F8"].Value = $"{data.requestor_name}";
            worksheet.Cells["F8"].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
            worksheet.Cells["F8"].Style.Font.Bold = true;

            worksheet.Cells["F9"].Value = $"{data.requestor_plant}";
            worksheet.Cells["F9"].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
            worksheet.Cells["F9"].Style.Font.Bold = true;

            worksheet.Cells["J3:N4"].Merge = true;
            worksheet.Cells["J3"].Value = "Invoice & Packing List";
            worksheet.Cells["J3"].Style.Font.Bold = true;
            worksheet.Cells["J3"].Style.Font.Size = 16;
            worksheet.Cells["J3"].Style.Font.Color.SetColor(System.Drawing.Color.Green);
            worksheet.Cells["J3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells["J3"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            worksheet.Cells["K6:N6"].Merge = true;
            worksheet.Cells["K6"].Value = "Duplicata de Facture";
            worksheet.Cells["K6"].Style.Font.Italic = true;
            worksheet.Cells["K6"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            worksheet.Cells["P1:S1"].Merge = true;
            worksheet.Cells["P1"].Value = "Proforma Invoice";
            worksheet.Cells["P1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells["P1"].Style.Fill.BackgroundColor.SetColor(yellowBg);
            worksheet.Cells["P1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells["P1"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["P1"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["Q1"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["R1"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["S1"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["S1"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            worksheet.Cells["P2:S2"].Merge = true;
            worksheet.Cells["P2"].Value = data.proforma_no;
            worksheet.Cells["P2"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells["P2"].Style.Fill.BackgroundColor.SetColor(yellowBg);
            worksheet.Cells["P2"].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
            worksheet.Cells["P2"].Style.Font.Bold = true;
            worksheet.Cells["P2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells["P2"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["P2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["Q2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["R2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["S2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells["S2"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            worksheet.Cells["O3"].Value = "Ship to:";
            worksheet.Cells["O3"].Style.Font.Bold = true;
            worksheet.Cells["O3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            worksheet.Cells["P3"].Value = data.vendor_name;
            worksheet.Cells["P3"].Style.Font.Bold = true;

            worksheet.Cells["P4"].Value = data.street;
            worksheet.Cells["P4"].Style.Font.Bold = true;

            worksheet.Cells["P5"].Value = data.city;
            worksheet.Cells["P5"].Style.Font.Bold = true;

            worksheet.Cells["P6"].Value = $"{data.country}, {data.postal_code}";
            worksheet.Cells["P6"].Style.Font.Bold = true;

            worksheet.Cells["O7"].Value = "Telp:";
            worksheet.Cells["O7"].Style.Font.Bold = true;
            worksheet.Cells["O7"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            worksheet.Cells["P7"].Value = data.phone_no;
            worksheet.Cells["P7"].Style.Font.Bold = true;

            worksheet.Cells["O8"].Value = "Attn. To:";
            worksheet.Cells["O8"].Style.Font.Bold = true;
            worksheet.Cells["O8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            worksheet.Cells["P8"].Value = data.attn_to;
            worksheet.Cells["P8"].Style.Font.Bold = true;

            int currentRow = 10;

            worksheet.Cells[currentRow, 2, currentRow, 7].Merge = true;
            worksheet.Cells[currentRow, 2].Value = "Shipment of:";

            worksheet.Cells[currentRow, 8, currentRow, 9].Merge = true;
            worksheet.Cells[currentRow, 8].Value = "Shipment Date:";

            worksheet.Cells[currentRow, 10, currentRow, 11].Merge = true;
            worksheet.Cells[currentRow, 10].Value = "Shipment Type:";

            worksheet.Cells[currentRow, 12].Value = "DHL AWB";
            worksheet.Cells[currentRow, 12].Style.Font.Color.SetColor(System.Drawing.Color.Red);
            worksheet.Cells[currentRow, 12].Style.Font.Bold = true;

            worksheet.Cells[currentRow, 13].Value = "Currency:";

            worksheet.Cells[currentRow, 14, currentRow, 15].Merge = true;
            worksheet.Cells[currentRow, 14].Value = "Freight Term:";

            var headerRange = worksheet.Cells[currentRow, 2, currentRow, 15];
            headerRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            headerRange.Style.Fill.BackgroundColor.SetColor(blueBg);
            headerRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            headerRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            headerRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            headerRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            headerRange.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            headerRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            currentRow++;

            worksheet.Cells[currentRow, 2, currentRow, 7].Merge = true;
            worksheet.Cells[currentRow, 2].Value = $"{data.total_boxes} BOX    {data.total_items} Items";

            worksheet.Cells[currentRow, 8, currentRow, 9].Merge = true;
            worksheet.Cells[currentRow, 8].Value = data.shipment_date?.ToString("dd-MMM-yy");

            worksheet.Cells[currentRow, 10, currentRow, 11].Merge = true;
            worksheet.Cells[currentRow, 10].Value = $"{data.ship_mode}";

            worksheet.Cells[currentRow, 12].Value = data.dhl_awb;
            worksheet.Cells[currentRow, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 12].Style.Fill.BackgroundColor.SetColor(yellowBg);

            worksheet.Cells[currentRow, 13].Value = "USD";

            worksheet.Cells[currentRow, 14, currentRow, 15].Merge = true;
            worksheet.Cells[currentRow, 14].Value = $"{data.freight_charges}\n({data.incoterms})";
            worksheet.Cells[currentRow, 14].Style.WrapText = true;

            var dataRange = worksheet.Cells[currentRow, 2, currentRow, 11];
            dataRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            dataRange.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            dataRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            dataRange.Style.Fill.BackgroundColor.SetColor(greenBg);

            var dataRange2 = worksheet.Cells[currentRow, 13, currentRow, 15];
            dataRange2.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange2.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange2.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange2.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            dataRange2.Style.Fill.BackgroundColor.SetColor(greenBg);

            var dataRange3 = worksheet.Cells[currentRow, 10, currentRow, 15];
            dataRange3.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange3.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange3.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange3.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange3.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            dataRange3.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            dataRange3.Style.Font.Color.SetColor(System.Drawing.Color.Blue);
            dataRange3.Style.Font.Bold = true;

            currentRow += 2;

            worksheet.Cells[currentRow, 1, currentRow + 1, 2].Merge = true;
            worksheet.Cells[currentRow, 1].Value = "Handling Unit";
            var handlingUnitRange = worksheet.Cells[currentRow, 1, currentRow + 1, 2];
            handlingUnitRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            handlingUnitRange.Style.Fill.BackgroundColor.SetColor(greenBg);

            worksheet.Cells[currentRow, 3, currentRow, 7].Merge = true;
            worksheet.Cells[currentRow, 3].Value = "Dimension (cm)";

            worksheet.Cells[currentRow, 8, currentRow + 1, 8].Merge = true;
            worksheet.Cells[currentRow, 8].Value = "Gross Weight";
            var grossWeightRange = worksheet.Cells[currentRow, 8, currentRow + 1, 8];
            grossWeightRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            grossWeightRange.Style.Fill.BackgroundColor.SetColor(greenBg);

            worksheet.Cells[currentRow, 9, currentRow + 1, 9].Merge = true;
            worksheet.Cells[currentRow, 9].Value = "Net Weight";

            worksheet.Cells[currentRow, 10, currentRow + 1, 10].Merge = true;
            worksheet.Cells[currentRow, 10].Value = "S/N";
            var snRange = worksheet.Cells[currentRow, 10, currentRow + 1, 10];
            snRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            snRange.Style.Fill.BackgroundColor.SetColor(greenBg);

            worksheet.Cells[currentRow, 11, currentRow + 1, 11].Merge = true;
            worksheet.Cells[currentRow, 11].Value = "Customer PO. #";

            worksheet.Cells[currentRow, 12, currentRow + 1, 12].Merge = true;
            worksheet.Cells[currentRow, 12].Value = "Product";
            var productRange = worksheet.Cells[currentRow, 12, currentRow + 1, 12];
            productRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            productRange.Style.Fill.BackgroundColor.SetColor(greenBg);

            worksheet.Cells[currentRow, 13, currentRow + 1, 15].Merge = true;
            worksheet.Cells[currentRow, 13].Value = "Description";
            var descriptionRange = worksheet.Cells[currentRow, 13, currentRow + 1, 15];
            descriptionRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            descriptionRange.Style.Fill.BackgroundColor.SetColor(greenBg);

            worksheet.Cells[currentRow, 16, currentRow + 1, 16].Merge = true;
            worksheet.Cells[currentRow, 16].Value = "QTY";

            worksheet.Cells[currentRow, 17, currentRow + 1, 17].Merge = true;
            worksheet.Cells[currentRow, 17].Value = "Unit Price";
            var unitPriceRange = worksheet.Cells[currentRow, 17, currentRow + 1, 17];
            unitPriceRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            unitPriceRange.Style.Fill.BackgroundColor.SetColor(greenBg);

            worksheet.Cells[currentRow, 18, currentRow + 1, 18].Merge = true;
            worksheet.Cells[currentRow, 18].Value = "Um";

            worksheet.Cells[currentRow, 19, currentRow + 1, 19].Merge = true;
            worksheet.Cells[currentRow, 19].Value = "Extended Price";
            var extendedPriceRange = worksheet.Cells[currentRow, 19, currentRow + 1, 19];
            extendedPriceRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            extendedPriceRange.Style.Fill.BackgroundColor.SetColor(greenBg);

            worksheet.Cells[currentRow, 20, currentRow + 1, 20].Merge = true;
            worksheet.Cells[currentRow, 20].Value = "HS CODE";

            currentRow++;

            worksheet.Cells[currentRow, 3].Value = "L";
            worksheet.Cells[currentRow, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            worksheet.Cells[currentRow, 4].Value = "x";
            worksheet.Cells[currentRow, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            worksheet.Cells[currentRow, 5].Value = "W";
            worksheet.Cells[currentRow, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            worksheet.Cells[currentRow, 6].Value = "x";
            worksheet.Cells[currentRow, 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            worksheet.Cells[currentRow, 7].Value = "H";
            worksheet.Cells[currentRow, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            var allHeaderRange = worksheet.Cells[currentRow - 1, 1, currentRow, 20];
            allHeaderRange.Style.Font.Bold = true;
            allHeaderRange.Style.WrapText = true;
            allHeaderRange.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            allHeaderRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            for (int row = currentRow - 1; row <= currentRow; row++)
            {
                for (int col = 1; col <= 20; col++)
                {
                    worksheet.Cells[row, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[row, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[row, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[row, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }
            }

            currentRow++;

            decimal totalVolume = 0;
            decimal totalNetWeight = 0;
            decimal totalQty = 0;
            decimal totalExtendedPrice = 0;
            var boxMergeInfo = new List<(int startRow, int endRow, string boxNo)>();

            if (data.boxes != null)
            {
                foreach (var box in data.boxes)
                {
                    decimal lengthCm = box.length_cm;
                    decimal widthCm = box.width_cm;
                    decimal heightCm = box.height_cm;
                    if (lengthCm > 0 && widthCm > 0 && heightCm > 0)
                    {
                        decimal lengthM = lengthCm / 100m;
                        decimal widthM = widthCm / 100m;
                        decimal heightM = heightCm / 100m;

                        decimal boxVolume = lengthM * widthM * heightM;
                        totalVolume += boxVolume;
                    }

                    totalNetWeight += box.net_weight_kg;

                    int boxSerialNumber = 1;
                    int boxStartRow = currentRow;

                    if (box.items != null && box.items.Any())
                    {
                        foreach (var asset in box.items)
                        {
                            for (int col = 1; col <= 20; col++)
                            {
                                worksheet.Cells[currentRow, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[currentRow, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[currentRow, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                            }

                            if (asset == box.items.First())
                            {
                                worksheet.Cells[currentRow, 3].Value = (double)lengthCm;
                                worksheet.Cells[currentRow, 3].Style.Numberformat.Format = "0.0";

                                worksheet.Cells[currentRow, 4].Value = "x";
                                worksheet.Cells[currentRow, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                worksheet.Cells[currentRow, 5].Value = (double)widthCm;
                                worksheet.Cells[currentRow, 5].Style.Numberformat.Format = "0.0";

                                worksheet.Cells[currentRow, 6].Value = "x";
                                worksheet.Cells[currentRow, 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                worksheet.Cells[currentRow, 7].Value = (double)heightCm;
                                worksheet.Cells[currentRow, 7].Style.Numberformat.Format = "0.0";

                                worksheet.Cells[currentRow, 8].Value = (double)box.gross_weight_kg;
                                worksheet.Cells[currentRow, 8].Style.Numberformat.Format = "0.0";

                                worksheet.Cells[currentRow, 9].Value = (double)box.net_weight_kg;
                                worksheet.Cells[currentRow, 9].Style.Numberformat.Format = "0.0";
                            }

                            worksheet.Cells[currentRow, 10].Value = boxSerialNumber++;

                            worksheet.Cells[currentRow, 11].Value = data.invoice_payment;

                            worksheet.Cells[currentRow, 12].Value = $"{asset.po_or_asset_no}";

                            try
                            {
                                worksheet.Cells[currentRow, 13, currentRow, 15].Merge = true;
                                worksheet.Cells[currentRow, 13].Value = asset.description;
                                worksheet.Cells[currentRow, 13, currentRow, 15].Style.WrapText = true;
                            }
                            catch (InvalidOperationException)
                            {
                                worksheet.Cells[currentRow, 13].Value = asset.description;
                                worksheet.Cells[currentRow, 13].Style.WrapText = true;
                            }

                            if (!string.IsNullOrEmpty(asset.description))
                            {
                                int charactersPerLine = 26;
                                string[] explicitLines = asset.description.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                                int totalLines = 0;

                                foreach (string line in explicitLines)
                                {
                                    if (string.IsNullOrWhiteSpace(line))
                                    {
                                        totalLines += 1;
                                    }
                                    else
                                    {
                                        int linesForThisSegment = Math.Max(1, (int)Math.Ceiling((double)line.Length / charactersPerLine));
                                        totalLines += linesForThisSegment;
                                    }
                                }

                                if (explicitLines.Length <= 1)
                                {
                                    totalLines = Math.Max(1, (int)Math.Ceiling((double)asset.description.Length / charactersPerLine));
                                }

                                double calculatedHeight = Math.Max(18, totalLines * 18);
                                worksheet.Row(currentRow).Height = calculatedHeight;
                                worksheet.Row(currentRow).CustomHeight = true;
                            }
                            else
                            {
                                worksheet.Row(currentRow).Height = 20;
                                worksheet.Row(currentRow).CustomHeight = true;
                            }

                            worksheet.Cells[currentRow, 16].Value = $"{asset.qty.ToString("#,##0.##")} {asset.uom}";
                            worksheet.Cells[currentRow, 16].Style.Numberformat.Format = "#,##0.00";

                            worksheet.Cells[currentRow, 17].Value = (double)asset.price_value;
                            worksheet.Cells[currentRow, 17].Style.Numberformat.Format = "#,##0.00";

                            worksheet.Cells[currentRow, 18].Value = "PC";

                            worksheet.Cells[currentRow, 19].Value = (double)(asset.price_value * asset.qty);
                            worksheet.Cells[currentRow, 19].Style.Numberformat.Format = "#,##0.00";

                            worksheet.Cells[currentRow, 20].Value = asset.hs_code?.ToString() ?? "";

                            totalQty += asset.qty;
                            totalExtendedPrice += asset.price_value * asset.qty;

                            currentRow++;
                        }

                        int boxEndRow = currentRow - 1;
                        boxMergeInfo.Add((boxStartRow, boxEndRow, box.box_no));

                        for (int row = boxStartRow; row <= boxEndRow; row++)
                        {
                            worksheet.Cells[row, 10].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[row, 10].Style.Fill.BackgroundColor.SetColor(greenBg);
                        }

                        for (int row = boxStartRow; row <= boxEndRow; row++)
                        {
                            worksheet.Cells[row, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[row, 12].Style.Fill.BackgroundColor.SetColor(greenBg);
                        }

                        for (int row = boxStartRow; row <= boxEndRow; row++)
                        {
                            var descCells = worksheet.Cells[row, 13, row, 15];
                            descCells.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            descCells.Style.Fill.BackgroundColor.SetColor(greenBg);
                        }

                        for (int row = boxStartRow; row <= boxEndRow; row++)
                        {
                            worksheet.Cells[row, 17].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[row, 17].Style.Fill.BackgroundColor.SetColor(greenBg);

                            worksheet.Cells[row, 19].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[row, 19].Style.Fill.BackgroundColor.SetColor(greenBg);
                        }
                    }
                }
            }

            foreach (var (startRow, endRow, boxNo) in boxMergeInfo)
            {
                try
                {
                    if (startRow <= endRow)
                    {
                        worksheet.Cells[startRow, 1, endRow, 2].Merge = true;
                        worksheet.Cells[startRow, 1].Value = boxNo;
                        var handlingUnitCells = worksheet.Cells[startRow, 1, endRow, 2];
                        handlingUnitCells.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        handlingUnitCells.Style.Fill.BackgroundColor.SetColor(greenBg);

                        worksheet.Cells[startRow, 3, endRow, 3].Merge = true;
                        worksheet.Cells[startRow, 4, endRow, 4].Merge = true;
                        worksheet.Cells[startRow, 5, endRow, 5].Merge = true;
                        worksheet.Cells[startRow, 6, endRow, 6].Merge = true;
                        worksheet.Cells[startRow, 7, endRow, 7].Merge = true;

                        worksheet.Cells[startRow, 8, endRow, 8].Merge = true;
                        var grossWeightCells = worksheet.Cells[startRow, 8, endRow, 8];
                        grossWeightCells.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        grossWeightCells.Style.Fill.BackgroundColor.SetColor(greenBg);

                        worksheet.Cells[startRow, 9, endRow, 9].Merge = true;
                    }
                }
                catch (InvalidOperationException ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Merge conflict for box {boxNo} rows {startRow}-{endRow}: {ex.Message}");
                }
            }

            worksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
            worksheet.Cells[currentRow, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 1].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 10].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 10].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 12].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 13, currentRow, 15].Merge = true;
            worksheet.Cells[currentRow, 13].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 13].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 17].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 17].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 19].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 19].Style.Fill.BackgroundColor.SetColor(greenBg);

            for (int col = 1; col <= 20; col++)
            {
                worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }

            currentRow++;

            if (!string.IsNullOrEmpty(data.coo))
            {
                try
                {
                    worksheet.Cells[currentRow, 8, currentRow, 15].Merge = true;
                    worksheet.Cells[currentRow, 8].Value = $"COO: {data.coo}";
                    worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(stabiloBg);
                    worksheet.Cells[currentRow, 8].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                    worksheet.Cells[currentRow, 8].Style.Font.Bold = true;
                    worksheet.Cells[currentRow, 8].Style.WrapText = true;

                    for (int col = 8; col <= 15; col++)
                    {
                        worksheet.Cells[currentRow, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    worksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                    worksheet.Cells[currentRow, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 1].Style.Fill.BackgroundColor.SetColor(greenBg);
                    worksheet.Cells[currentRow, 17].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 17].Style.Fill.BackgroundColor.SetColor(greenBg);
                    worksheet.Cells[currentRow, 19].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 19].Style.Fill.BackgroundColor.SetColor(greenBg);

                    for (int col = 1; col <= 20; col++)
                    {
                        worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    currentRow++;
                }
                catch (InvalidOperationException)
                {
                    worksheet.Cells[currentRow, 8].Value = $"COO: {data.coo}";
                    worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(yellowBg);
                    worksheet.Cells[currentRow, 8].Style.Font.Bold = true;
                    currentRow++;
                }
            }

            if (!string.IsNullOrEmpty(data.courier_account_no) &&
                !string.IsNullOrEmpty(data.courier_charges) &&
                data.courier_charges.Equals("Courier", StringComparison.OrdinalIgnoreCase) ||
                data.courier_charges.Equals("Receiver", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    worksheet.Cells[currentRow, 8, currentRow, 15].Merge = true;
                    worksheet.Cells[currentRow, 8].Value = $"ACC: {data.courier_account_no}";
                    worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(stabiloBg);
                    worksheet.Cells[currentRow, 8].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                    worksheet.Cells[currentRow, 8].Style.Font.Bold = true;
                    worksheet.Cells[currentRow, 8].Style.WrapText = true;

                    for (int col = 8; col <= 15; col++)
                    {
                        worksheet.Cells[currentRow, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    worksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                    worksheet.Cells[currentRow, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 1].Style.Fill.BackgroundColor.SetColor(greenBg);
                    worksheet.Cells[currentRow, 17].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 17].Style.Fill.BackgroundColor.SetColor(greenBg);
                    worksheet.Cells[currentRow, 19].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 19].Style.Fill.BackgroundColor.SetColor(greenBg);

                    for (int col = 1; col <= 20; col++)
                    {
                        worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    currentRow++;
                }
                catch (InvalidOperationException)
                {
                    worksheet.Cells[currentRow, 8].Value = $"ACC: {data.courier_account_no}";
                    worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(yellowBg);
                    worksheet.Cells[currentRow, 8].Style.Font.Bold = true;
                    currentRow++;
                }
            }

            for (int col = 8; col <= 15; col++)
            {
                worksheet.Cells[currentRow - 1, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }

            if (!string.IsNullOrEmpty(data.remark))
            {
                try
                {
                    worksheet.Cells[currentRow, 8, currentRow, 15].Merge = true;
                    worksheet.Cells[currentRow, 8].Value = data.remark;
                    worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(stabiloBg);
                    worksheet.Cells[currentRow, 8].Style.Font.Bold = true;

                    for (int col = 8; col <= 15; col++)
                    {
                        worksheet.Cells[currentRow, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    worksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                    worksheet.Cells[currentRow, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 1].Style.Fill.BackgroundColor.SetColor(greenBg);
                    worksheet.Cells[currentRow, 17].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 17].Style.Fill.BackgroundColor.SetColor(greenBg);
                    worksheet.Cells[currentRow, 19].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 19].Style.Fill.BackgroundColor.SetColor(greenBg);

                    for (int col = 1; col <= 20; col++)
                    {
                        worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    currentRow++;
                }
                catch (InvalidOperationException)
                {
                    worksheet.Cells[currentRow, 8].Value = data.remark;
                    worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(yellowBg);
                    worksheet.Cells[currentRow, 8].Style.Font.Bold = true;
                    currentRow++;
                }
            }

            worksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
            worksheet.Cells[currentRow, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 1].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 10].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 10].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 12].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 13, currentRow, 15].Merge = true;
            worksheet.Cells[currentRow, 13].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 13].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 17].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 17].Style.Fill.BackgroundColor.SetColor(greenBg);
            worksheet.Cells[currentRow, 19].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 19].Style.Fill.BackgroundColor.SetColor(greenBg);

            for (int col = 1; col <= 20; col++)
            {
                worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }

            currentRow++;

            try
            {
                worksheet.Cells[currentRow, 8, currentRow, 15].Merge = true;
                worksheet.Cells[currentRow, 8].Value = "Items not for sale";
                worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(stabiloBg);
                worksheet.Cells[currentRow, 8].Style.Font.Color.SetColor(System.Drawing.Color.Green);
                worksheet.Cells[currentRow, 8].Style.Font.Bold = true;
                worksheet.Cells[currentRow, 8].Style.WrapText = true;

                for (int col = 8; col <= 15; col++)
                {
                    worksheet.Cells[currentRow, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                currentRow++;

                worksheet.Cells[currentRow, 8, currentRow, 15].Merge = true;
                worksheet.Cells[currentRow, 8].Value = "Value declare for custom purpose only";
                worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(stabiloBg);
                worksheet.Cells[currentRow, 8].Style.Font.Color.SetColor(System.Drawing.Color.Green);
                worksheet.Cells[currentRow, 8].Style.Font.Bold = true;
                worksheet.Cells[currentRow, 8].Style.WrapText = true;

                for (int col = 8; col <= 15; col++)
                {
                    worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                currentRow++;

                worksheet.Cells[currentRow, 8, currentRow, 15].Merge = true;
                worksheet.Cells[currentRow, 8].Value = "Free of Charge";
                worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(stabiloBg);
                worksheet.Cells[currentRow, 8].Style.Font.Color.SetColor(System.Drawing.Color.Green);
                worksheet.Cells[currentRow, 8].Style.Font.Bold = true;
                worksheet.Cells[currentRow, 8].Style.WrapText = true;

                for (int col = 8; col <= 15; col++)
                {
                    worksheet.Cells[currentRow, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                for (int remarkRow = currentRow - 2; remarkRow <= currentRow; remarkRow++)
                {
                    worksheet.Cells[remarkRow, 1, remarkRow, 2].Merge = true;
                    worksheet.Cells[remarkRow, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[remarkRow, 1].Style.Fill.BackgroundColor.SetColor(greenBg);
                    worksheet.Cells[remarkRow, 17].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[remarkRow, 17].Style.Fill.BackgroundColor.SetColor(greenBg);
                    worksheet.Cells[remarkRow, 19].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[remarkRow, 19].Style.Fill.BackgroundColor.SetColor(greenBg);

                    for (int col = 1; col <= 20; col++)
                    {
                        worksheet.Cells[remarkRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[remarkRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }
                }

                currentRow += 1;

                worksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                worksheet.Cells[currentRow, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 1].Style.Fill.BackgroundColor.SetColor(greenBg);
                worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(greenBg);
                worksheet.Cells[currentRow, 10].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 10].Style.Fill.BackgroundColor.SetColor(greenBg);
                worksheet.Cells[currentRow, 12].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 12].Style.Fill.BackgroundColor.SetColor(greenBg);
                worksheet.Cells[currentRow, 13, currentRow, 15].Merge = true;
                worksheet.Cells[currentRow, 13].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 13].Style.Fill.BackgroundColor.SetColor(greenBg);
                worksheet.Cells[currentRow, 17].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 17].Style.Fill.BackgroundColor.SetColor(greenBg);
                worksheet.Cells[currentRow, 19].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 19].Style.Fill.BackgroundColor.SetColor(greenBg);

                for (int col = 1; col <= 20; col++)
                {
                    worksheet.Cells[currentRow, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                currentRow += 1;
            }
            catch (InvalidOperationException)
            {
                worksheet.Cells[currentRow, 8].Value = "Items not for sale, value declared for customs purposes only, free of charge.";
                worksheet.Cells[currentRow, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(yellowBg);
                worksheet.Cells[currentRow, 8].Style.Font.Bold = true;
                currentRow++;
            }

            worksheet.Cells[currentRow, 1, currentRow, 3].Merge = true;
            worksheet.Cells[currentRow, 1].Value = "Grand Total";

            worksheet.Cells[currentRow, 4, currentRow, 6].Merge = true;
            worksheet.Cells[currentRow, 4].Value = (double)totalVolume;

            worksheet.Cells[currentRow, 7].Value = "M3";

            worksheet.Cells[currentRow, 9].Value = (double)totalNetWeight;
            worksheet.Cells[currentRow, 9].Style.Numberformat.Format = "0.0";

            worksheet.Cells[currentRow, 10].Value = "Kgs";

            worksheet.Cells[currentRow, 18].Value = '$';

            worksheet.Cells[currentRow, 19].Value = (double)totalExtendedPrice;
            worksheet.Cells[currentRow, 19].Style.Numberformat.Format = "#,##0.00";

            for (int col = 1; col <= 19; col++)
            {
                worksheet.Cells[currentRow, col].Style.Font.Bold = true;
            }

            currentRow++;

            for (int col = 1; col <= 19; col++)
            {
                worksheet.Cells[currentRow, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            }

            worksheet.Column(1).Width = 1.91;
            worksheet.Column(2).Width = 6.55;
            worksheet.Column(3).AutoFit();
            worksheet.Column(4).Width = 2.91;
            worksheet.Column(5).AutoFit();
            worksheet.Column(6).Width = 2.91;
            worksheet.Column(7).AutoFit();
            worksheet.Column(8).AutoFit();
            worksheet.Column(9).AutoFit();
            worksheet.Column(10).Width = 4.55;
            worksheet.Column(11).Width = 9.18;
            worksheet.Column(12).Width = 19.36;
            worksheet.Column(13).Width = 8.31;
            worksheet.Column(14).Width = 3.18;
            worksheet.Column(15).Width = 15.45;
            worksheet.Column(16).Width = 9.55;
            worksheet.Column(17).AutoFit();
            worksheet.Column(18).Width = 4.18;
            worksheet.Column(19).AutoFit();
            worksheet.Column(20).Width = 10.82;
            worksheet.Row(11).Height = 31.5;
            worksheet.Row(currentRow).Height = 2.2;

            worksheet.PrinterSettings.Orientation = OfficeOpenXml.eOrientation.Landscape;
            worksheet.PrinterSettings.FitToPage = true;
        }

    }
}