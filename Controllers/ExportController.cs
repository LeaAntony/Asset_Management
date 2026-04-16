using Microsoft.AspNetCore.Mvc;
using System.Data.SqlClient;
using Asset_Management.Models;
using Asset_Management.Function;
using System.Data;
using Asset_Management.Service;
using Microsoft.EntityFrameworkCore;
using System.Linq.Dynamic.Core;
using iText.Html2pdf;
using iText.IO.Source;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using System.Text;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Layout;
using iText.Kernel.Utils;
using Microsoft.AspNetCore.Hosting.Server;
using System.IO;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Events;
using iText.Kernel.Pdf.Canvas.Parser;
using Org.BouncyCastle.Utilities.Zlib;
using iText.IO.Image;
using iText.Kernel.Pdf.Layer;
using static System.Runtime.InteropServices.JavaScript.JSType;
using iText.IO.Font;
using iText.Kernel.Font;
using iText.Kernel.Colors;
using iText.IO.Font.Constants;
using OfficeOpenXml;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.Formats;
using SixLabors.ImageSharp.Formats.Jpeg;
using DocumentFormat.OpenXml.Wordprocessing;
using PageSize = iText.Kernel.Geom.PageSize;


namespace Asset_Management.Controllers
{
    public class ExportController : Controller
    {
        private readonly ApplicationDbContext _context;

        private readonly ImportExportFactory _importexportFactory;
        private readonly ILogger<ExportController> _logger;

        private string DbConnection()
        {
            var dbAccess = new DatabaseAccessLayer();
            string dbString = dbAccess.ConnectionString;
            return dbString;
        }
        private float CmToPoints(float cm)
        {
            return cm * 28.35f;
        }

        public IActionResult DownloadFile(string fileName)
        {
            var filePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), fileName);

            if (System.IO.File.Exists(filePath))
            {
                var file = System.IO.File.OpenRead(filePath);
                return File(file, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }

            return NotFound();
        }
        public IActionResult CreateAssetLabel(string asset_no, string asset_subnumber)
        {
            var db = new DatabaseAccessLayer();

            string htmlToPdf = @"<html>
                                      <header>
                                        <style>
                                          body {
                                            font-family:Helvetica ;
                                            font-size:9pt;
                                          }
                                        </style>
                                      </header>
                                      <body style='margin-top:0px'>
                                        <img src='wwwroot/images/logo/aml_black.png' style='display:none' />
                                        <table style='width:100%;'>
                                          <tr>
                                            <td>Asset No.</td>
                                            <td>: </td>
                                            <td>2114094</td>
                                            <td>1</td>
                                          </tr>
                                          <tr>
                                            <td>Department</td>
                                            <td>: </td>
                                            <td colspan='2'>JHO-ATS48 & Zeph-MBC</td>
                                          </tr>
                                          <tr>
                                            <td>Supplier</td>
                                            <td>: </td>
                                            <td colspan='2'>Asset Management</td>
                                          </tr>
                                          <tr>
                                            <td>Cap. Date</td>
                                            <td>: </td>
                                            <td colspan='2'>18 Oktober 2025</td>
                                          </tr>
                                          <tr>
                                            <td>Requestor</td>
                                            <td>: </td>
                                            <td colspan='2'>Lea Antony</td>
                                          </tr>
                                          <tr>
                                            <td>Cost Center</td>
                                            <td>: </td>
                                            <td colspan='2'>>2020625</td>
                                          </tr>
                                          <tr>
                                            <td colspan='4'>LASDASD/ADSFSDAG</td>
                                          </tr>
                                        </table>
                                      </body>
                                    </html>";


            float widthCm = 7.62f;
            float heightCm = 5.08f;
            PageSize pageSize = new PageSize(CmToPoints(widthCm), CmToPoints(heightCm));

            using (var memoryStream = new MemoryStream())
            {
                PdfWriter writer = new PdfWriter(memoryStream);
                PdfDocument pdfDocument = new PdfDocument(writer);

                pdfDocument.SetDefaultPageSize(pageSize);

                using (MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(htmlToPdf)))
                {
                    HtmlConverter.ConvertToPdf(htmlStream, pdfDocument);
                }

                pdfDocument.Close();

                Response.Headers["Content-Disposition"] = "inline; filename=label.pdf";
                return File(memoryStream.ToArray(), "application/pdf");
            }
        }
        public string CreateAssetTag(IFormFileCollection imgs, string filename, string asset_no, string asset_subnumber)
        {
            try
            {
                string imageHtml = "";
                int image_no = 0;
                foreach (var img in imgs)
                {
                    image_no++;

                    if (img.Length > 0)
                    {
                        using (var memoryStream = new MemoryStream())
                        {
                            img.CopyTo(memoryStream);
                            memoryStream.Position = 0;

                            using (var image = SixLabors.ImageSharp.Image.Load(memoryStream))
                            {
                                int maxDimension = 800;
                                if (image.Width > maxDimension || image.Height > maxDimension)
                                {
                                    image.Mutate(x => x.Resize(new ResizeOptions
                                    {
                                        Size = new Size(maxDimension, maxDimension),
                                        Mode = ResizeMode.Max
                                    }));
                                }

                                using (var finalStream = new MemoryStream())
                                {
                                    image.SaveAsJpeg(finalStream, new JpegEncoder { Quality = 90 });
                                    finalStream.Position = 0;

                                    byte[] compressedImageBytes = finalStream.ToArray();
                                    string base64Image = Convert.ToBase64String(compressedImageBytes);

                                    if (image_no % 2 != 0)
                                    {
                                        if (image_no != 1)
                                        {
                                            imageHtml += "</tr>";
                                        }
                                        imageHtml += "<tr>";
                                    }

                                    imageHtml += $"<td style='border:0.5px solid black'><img src=\"data:image/jpeg;base64,{base64Image}\" style='width:100%; height:auto;' alt='Uploaded Image' /></td>";
                                }
                            }
                        }
                    }
                }
                imageHtml = "<html><body style='margin-top:100px;margin-bottom:100px'><table>" + imageHtml + "</table></html></body>";

                MemoryStream streamImage = new MemoryStream();
                HtmlConverter.ConvertToPdf(imageHtml, streamImage);

                ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
                PdfWriter writer = new PdfWriter(byteArrayOutputStream);
                PdfDocument pdfDocument = new PdfDocument(writer);

                MemoryStream streamImage2 = new MemoryStream(streamImage.ToArray());
                using (var inputPdfDoc = new PdfDocument(new PdfReader(streamImage2)))
                {
                    inputPdfDoc.CopyPagesTo(1, inputPdfDoc.GetNumberOfPages(), pdfDocument);
                }

                AddHeaderAndFooterTag(pdfDocument, asset_no, asset_subnumber);
                pdfDocument.Close();

                string filePath = System.IO.Path.Combine("wwwroot", "Upload", "Asset", filename);
                System.IO.File.WriteAllBytes(filePath, byteArrayOutputStream.ToArray());

                return "OK;" + filename;

            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        public string CreateAssetCount(IFormFileCollection imgs, string filename, string count_year, string asset_no, string asset_subnumber)
        {
            try
            {
                string imageHtml = "";
                int image_no = 0;
                foreach (var img in imgs)
                {
                    image_no++;

                    if (img.Length > 0)
                    {
                        using (var memoryStream = new MemoryStream())
                        {
                            img.CopyTo(memoryStream);
                            memoryStream.Position = 0;

                            using (var image = SixLabors.ImageSharp.Image.Load(memoryStream))
                            {
                                int maxDimension = 800;
                                if (image.Width > maxDimension || image.Height > maxDimension)
                                {
                                    image.Mutate(x => x.Resize(new ResizeOptions
                                    {
                                        Size = new Size(maxDimension, maxDimension),
                                        Mode = ResizeMode.Max
                                    }));
                                }

                                using (var finalStream = new MemoryStream())
                                {
                                    image.SaveAsJpeg(finalStream, new JpegEncoder { Quality = 90 });
                                    finalStream.Position = 0;

                                    byte[] compressedImageBytes = finalStream.ToArray();
                                    string base64Image = Convert.ToBase64String(compressedImageBytes);

                                    if (image_no % 2 != 0)
                                    {
                                        if (image_no != 1)
                                        {
                                            imageHtml += "</tr>";
                                        }
                                        imageHtml += "<tr>";
                                    }
                                    imageHtml += $"<td style='border:0.5px solid black'><img src=\"data:image/jpeg;base64,{base64Image}\" style='width:350px; height:auto;' alt='Uploaded Image' /></td>";
                                }
                            }
                        }
                    }
                }
                imageHtml = "<html><body style='margin-top:100px;margin-bottom:100px'><table>" + imageHtml + "</table></html></body>";

                MemoryStream streamImage = new MemoryStream();
                HtmlConverter.ConvertToPdf(imageHtml, streamImage);

                ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
                PdfWriter writer = new PdfWriter(byteArrayOutputStream);
                PdfDocument pdfDocument = new PdfDocument(writer);

                MemoryStream streamImage2 = new MemoryStream(streamImage.ToArray());
                using (var inputPdfDoc = new PdfDocument(new PdfReader(streamImage2)))
                {
                    inputPdfDoc.CopyPagesTo(1, inputPdfDoc.GetNumberOfPages(), pdfDocument);
                }

                AddHeaderAndFooterCount(pdfDocument,count_year,asset_no,asset_subnumber);
                pdfDocument.Close();

                string filePath = System.IO.Path.Combine("wwwroot", "Upload", "Count", filename);
                System.IO.File.WriteAllBytes(filePath, byteArrayOutputStream.ToArray());

                return "OK;" + filename;

            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        public string CreateAssetRetagging(IFormFileCollection imgs, string filename, string count_year, string asset_no, string asset_subnumber)
        {
            try
            {
                string imageHtml = "";
                int image_no = 0;
                foreach (var img in imgs)
                {
                    image_no++;

                    if (img.Length > 0)
                    {
                        using (var memoryStream = new MemoryStream())
                        {
                            img.CopyTo(memoryStream);
                            memoryStream.Position = 0;

                            using (var image = SixLabors.ImageSharp.Image.Load(memoryStream))
                            {
                                int maxDimension = 800;
                                if (image.Width > maxDimension || image.Height > maxDimension)
                                {
                                    image.Mutate(x => x.Resize(new ResizeOptions
                                    {
                                        Size = new Size(maxDimension, maxDimension),
                                        Mode = ResizeMode.Max 
                                    }));
                                }

                                using (var finalStream = new MemoryStream())
                                {
                                    image.SaveAsJpeg(finalStream, new JpegEncoder { Quality = 90 });
                                    finalStream.Position = 0;

                                    byte[] compressedImageBytes = finalStream.ToArray();
                                    string base64Image = Convert.ToBase64String(compressedImageBytes);

                                    if (image_no % 2 != 0)
                                    {
                                        if (image_no != 1)
                                        {
                                            imageHtml += "</tr>";
                                        }
                                        imageHtml += "<tr>";
                                    }
                                    imageHtml += $"<td style='border:0.5px solid black'><img src=\"data:image/jpeg;base64,{base64Image}\" style='width:350px; height:auto;' alt='Uploaded Image' /></td>";
                                }
                            }
                        }
                    }
                }
                imageHtml = "<html><body style='margin-top:100px;margin-bottom:100px'><table>" + imageHtml + "</table></html></body>";

                MemoryStream streamImage = new MemoryStream();
                HtmlConverter.ConvertToPdf(imageHtml, streamImage);

                ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
                PdfWriter writer = new PdfWriter(byteArrayOutputStream);
                PdfDocument pdfDocument = new PdfDocument(writer);

                MemoryStream streamImage2 = new MemoryStream(streamImage.ToArray());
                using (var inputPdfDoc = new PdfDocument(new PdfReader(streamImage2)))
                {
                    inputPdfDoc.CopyPagesTo(1, inputPdfDoc.GetNumberOfPages(), pdfDocument);
                }

                AddHeaderAndFooterRetagging(pdfDocument, count_year, asset_no, asset_subnumber);
                pdfDocument.Close();

                string filePath = System.IO.Path.Combine("wwwroot", "Upload", "Count", filename);
                System.IO.File.WriteAllBytes(filePath, byteArrayOutputStream.ToArray());

                return "OK;" + filename;

            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        private void AddHeaderAndFooter(PdfDocument pdfDoc, string order_no = "")
        {
            for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
            {
                PdfPage page = pdfDoc.GetPage(i);
                PdfCanvas canvas = new PdfCanvas(page);

                ImageData data = ImageDataFactory.Create("wwwroot/images/logo/aml_black.png");
                data.SetHeight(80);
                canvas.AddImageAt(data, 400, 760, false);

                canvas.SaveState();
                canvas.BeginText();
                var font = iText.Kernel.Font.PdfFontFactory.CreateFont();
                PdfFont fontBold = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                canvas.SetFontAndSize(fontBold, 12);
                string text = "ASSET DESTRUCTION REPORT";
                float textWidth = font.GetWidth(text, 12);
                float pageWidth = page.GetPageSize().GetWidth();
                float xPosition = (pageWidth - textWidth) / 2;
                float yPosition = page.GetPageSize().GetHeight() - 50;
                canvas.SetTextMatrix(xPosition, yPosition);
                canvas.ShowText(text);
                canvas.EndText();
                canvas.RestoreState();

                canvas.SaveState();
                canvas.BeginText();
                float yPositionFooter = 90;
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.SetFontAndSize(fontBold, 9);
                canvas.ShowText("Asset Management");
                yPositionFooter -= 12;
                canvas.SetFontAndSize(iText.Kernel.Font.PdfFontFactory.CreateFont(), 9);
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.ShowText("Batam");
                yPositionFooter -= 12;
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.ShowText("Indonesia");
                yPositionFooter -= 12;
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.ShowText("Tel :  (62) 770 434343");
                yPositionFooter -= 12;
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.ShowText("Fax :  (62) 770 343434");

                yPositionFooter -= 20;
                float textWidthOrderNo = font.GetWidth(order_no, 12);
                float xPositionFooter = (pageWidth - textWidthOrderNo) / 2;
                canvas.SetTextMatrix(xPositionFooter, yPositionFooter);
                canvas.ShowText(order_no);

                canvas.EndText();
                canvas.RestoreState();
            }
        }
        private void AddHeaderAndFooterCount(PdfDocument pdfDoc, string count_year = "", string asset_no = "", string asset_subnumber = "")
        {
            for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
            {
                PdfPage page = pdfDoc.GetPage(i);
                PdfCanvas canvas = new PdfCanvas(page);

                ImageData data = ImageDataFactory.Create("wwwroot/images/logo/aml_black.png");
                data.SetHeight(80);
                canvas.AddImageAt(data, 400, 760, false);

                canvas.SaveState();
                canvas.BeginText();
                var font = iText.Kernel.Font.PdfFontFactory.CreateFont();
                PdfFont fontBold = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                canvas.SetFontAndSize(fontBold, 12);
                string text = "ASSET COUNT";
                float textWidth = font.GetWidth(text, 12);
                float pageWidth = page.GetPageSize().GetWidth();
                float xPosition = (pageWidth - textWidth) / 2;
                float yPosition = page.GetPageSize().GetHeight() - 50;
                canvas.SetTextMatrix(xPosition, yPosition);
                canvas.ShowText(text);
                canvas.EndText();
                canvas.RestoreState();

                canvas.SaveState();
                canvas.BeginText();
                float yPositionFooter = 90;
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.SetFontAndSize(fontBold, 9);
                canvas.ShowText("Asset Management");
                yPositionFooter -= 12;
                canvas.SetFontAndSize(iText.Kernel.Font.PdfFontFactory.CreateFont(), 9);
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.ShowText("Batam");
                yPositionFooter -= 12;
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.ShowText("Indonesia");
                yPositionFooter -= 12;
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.ShowText("Tel :  (62) 770 434343");
                yPositionFooter -= 12;
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.ShowText("Fax :  (62) 770 343434");

                yPositionFooter -= 20;
                float textWidthOrderNo = font.GetWidth(count_year + " " + asset_no+"-"+asset_subnumber, 12);
                float xPositionFooter = (pageWidth - textWidthOrderNo) / 2;
                canvas.SetTextMatrix(xPositionFooter, yPositionFooter);
                canvas.ShowText(count_year + " " + asset_no + "-" + asset_subnumber);

                canvas.EndText();
                canvas.RestoreState();
            }
        }
        private void AddHeaderAndFooterRetagging(PdfDocument pdfDoc, string count_year = "", string asset_no = "", string asset_subnumber = "")
        {
            for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
            {
                PdfPage page = pdfDoc.GetPage(i);
                PdfCanvas canvas = new PdfCanvas(page);

                ImageData data = ImageDataFactory.Create("wwwroot/images/logo/aml_black.png");
                data.SetHeight(80);
                canvas.AddImageAt(data, 400, 760, false);

                canvas.SaveState();
                canvas.BeginText();
                var font = iText.Kernel.Font.PdfFontFactory.CreateFont();
                PdfFont fontBold = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                canvas.SetFontAndSize(fontBold, 12);
                string text = "ASSET RETAGGING";
                float textWidth = font.GetWidth(text, 12);
                float pageWidth = page.GetPageSize().GetWidth();
                float xPosition = (pageWidth - textWidth) / 2;
                float yPosition = page.GetPageSize().GetHeight() - 50;
                canvas.SetTextMatrix(xPosition, yPosition);
                canvas.ShowText(text);
                canvas.EndText();
                canvas.RestoreState();

                canvas.SaveState();
                canvas.BeginText();
                float yPositionFooter = 90;
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.SetFontAndSize(fontBold, 9);
                canvas.ShowText("Asset Management");
                yPositionFooter -= 12;
                canvas.SetFontAndSize(iText.Kernel.Font.PdfFontFactory.CreateFont(), 9);
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.ShowText("Batam");
                yPositionFooter -= 12;
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.ShowText("Indonesia");
                yPositionFooter -= 12;
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.ShowText("Tel :  (62) 770 434343");
                yPositionFooter -= 12;
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.ShowText("Fax :  (62) 770 343434");

                yPositionFooter -= 20;
                float textWidthOrderNo = font.GetWidth(count_year + " " + asset_no + "-" + asset_subnumber, 12);
                float xPositionFooter = (pageWidth - textWidthOrderNo) / 2;
                canvas.SetTextMatrix(xPositionFooter, yPositionFooter);
                canvas.ShowText(count_year + " " + asset_no + "-" + asset_subnumber);

                canvas.EndText();
                canvas.RestoreState();
            }
        }
        private void AddHeaderAndFooterTag(PdfDocument pdfDoc, string asset_no = "", string asset_subnumber = "")
        {
            for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
            {
                PdfPage page = pdfDoc.GetPage(i);
                PdfCanvas canvas = new PdfCanvas(page);

                ImageData data = ImageDataFactory.Create("wwwroot/images/logo/aml_black.png");
                data.SetHeight(80);
                canvas.AddImageAt(data, 400, 760, false);

                canvas.SaveState();
                canvas.BeginText();
                var font = iText.Kernel.Font.PdfFontFactory.CreateFont();
                PdfFont fontBold = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                canvas.SetFontAndSize(fontBold, 12);
                string text = "ASSET TAGGING";
                float textWidth = font.GetWidth(text, 12);
                float pageWidth = page.GetPageSize().GetWidth();
                float xPosition = (pageWidth - textWidth) / 2;
                float yPosition = page.GetPageSize().GetHeight() - 50;
                canvas.SetTextMatrix(xPosition, yPosition);
                canvas.ShowText(text);
                canvas.EndText();
                canvas.RestoreState();

                canvas.SaveState();
                canvas.BeginText();
                float yPositionFooter = 90;
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.SetFontAndSize(fontBold, 9);
                canvas.ShowText("Asset Management");
                yPositionFooter -= 12;
                canvas.SetFontAndSize(iText.Kernel.Font.PdfFontFactory.CreateFont(), 9);
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.ShowText("Batam");
                yPositionFooter -= 12;
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.ShowText("Indonesia");
                yPositionFooter -= 12;
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.ShowText("Tel :  (62) 770 434343");
                yPositionFooter -= 12;
                canvas.SetTextMatrix(68, yPositionFooter);
                canvas.ShowText("Fax :  (62) 770 343434");

                yPositionFooter -= 20;
                float textWidthOrderNo = font.GetWidth(asset_no + "-" + asset_subnumber, 12);
                float xPositionFooter = (pageWidth - textWidthOrderNo) / 2;
                canvas.SetTextMatrix(xPositionFooter, yPositionFooter);
                canvas.ShowText(asset_no + "-" + asset_subnumber);

                canvas.EndText();
                canvas.RestoreState();
            }
        }
        public FileResult GeneratePDF(string htmlToPdf, string pdfFilename)
        {
            using (MemoryStream stream = new MemoryStream(Encoding.ASCII.GetBytes(htmlToPdf)))
            {
                ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
                PdfWriter writer = new PdfWriter(byteArrayOutputStream);
                PdfDocument pdfDocument = new PdfDocument(writer);
                HtmlConverter.ConvertToPdf(stream, pdfDocument);

                pdfDocument.Close();
                Response.Headers["Content-Disposition"] = "inline; filename="+ pdfFilename;
                return File(byteArrayOutputStream.ToArray(), "application/pdf");
            }
        }
        public FileResult GeneratePDF_xx(string htmlToPdf, string pdfFilename)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
                PdfWriter writer = new PdfWriter(byteArrayOutputStream);
                PdfDocument pdfDocument = new PdfDocument(writer);
                HtmlConverter.ConvertToPdf(htmlToPdf, stream);
                MemoryStream stream2 = new MemoryStream(stream.ToArray());
                using (var inputPdfDoc = new PdfDocument(new PdfReader(stream2)))
                {
                    inputPdfDoc.CopyPagesTo(1, inputPdfDoc.GetNumberOfPages(), pdfDocument);
                }
                MemoryStream stream3 = new MemoryStream(stream.ToArray());
                using (var inputPdfDoc = new PdfDocument(new PdfReader(stream3)))
                {
                    inputPdfDoc.CopyPagesTo(1, inputPdfDoc.GetNumberOfPages(), pdfDocument);
                }
                AddHeaderAndFooter(pdfDocument);

                pdfDocument.Close();
                Response.Headers["Content-Disposition"] = "inline; filename=" + pdfFilename;
                return File(byteArrayOutputStream.ToArray(), "application/pdf");
            }
        }

        public ActionResult GenerateCombinedPdf_v1(string htmlToPdf, string[] existingPdfFiles, string pdfFilename)
        {
            using (MemoryStream combinedStream = new MemoryStream())
            {
                using (PdfWriter writer = new PdfWriter(combinedStream))
                {
                    using (PdfDocument pdfDocument = new PdfDocument(writer))
                    {
                        pdfDocument.SetDefaultPageSize(iText.Kernel.Geom.PageSize.A4);

                        using (MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(htmlToPdf)))
                        {
                            HtmlConverter.ConvertToPdf(htmlStream, pdfDocument);
                        }

                        PdfMerger merger = new PdfMerger(pdfDocument);

                        foreach (var pdfFilePath in existingPdfFiles)
                        {
                            if (System.IO.File.Exists(pdfFilePath))
                            {
                                using (PdfDocument srcDocument = new PdfDocument(new PdfReader(pdfFilePath)))
                                {
                                    if (!srcDocument.IsClosed())
                                    {
                                        merger.Merge(srcDocument, 1, srcDocument.GetNumberOfPages());
                                    }
                                    else
                                    {
                                        throw new InvalidOperationException("Source document is closed.");
                                    }
                                }
                            }
                            else
                            {
                                throw new FileNotFoundException($"The file '{pdfFilePath}' could not be found.");
                            }
                        }
                    }
                }
                Response.Headers["Content-Disposition"] = "inline; filename=" + pdfFilename;
                return File(combinedStream.ToArray(), "application/pdf");
            }
        }

        public IActionResult GenerateCombinedPdf(string htmlContent, List<string> pdfFilePaths, string pdfFilename)
        {
            MemoryStream combinedStream = new MemoryStream();

            using (MemoryStream htmlStream = new MemoryStream())
            {
                HtmlConverter.ConvertToPdf(htmlContent, htmlStream);
                htmlStream.Position = 0;

                using (PdfWriter writer = new PdfWriter(combinedStream))
                {
                    using (PdfDocument combinedPdf = new PdfDocument(writer))
                    {
                        using (PdfDocument tempPdf = new PdfDocument(new PdfReader(htmlStream)))
                        {
                            PdfMerger merger = new PdfMerger(combinedPdf);
                            merger.Merge(tempPdf, 1, tempPdf.GetNumberOfPages());
                        }

                        foreach (var pdfFilePath in pdfFilePaths)
                        {
                            if (System.IO.File.Exists(pdfFilePath))
                            {
                                using (PdfDocument srcPdf = new PdfDocument(new PdfReader(pdfFilePath)))
                                {
                                    PdfMerger merger = new PdfMerger(combinedPdf);
                                    merger.Merge(srcPdf, 1, srcPdf.GetNumberOfPages());
                                }
                            }
                            else
                            {
                                throw new FileNotFoundException($"The file '{pdfFilePath}' could not be found.");
                            }
                        }
                    }
                }
            }

            combinedStream.Position = 0;
            Response.Headers["Content-Disposition"] = $"inline; filename={pdfFilename}";
            return File(combinedStream.ToArray(), "application/pdf");
        }

        public IActionResult PRINT_ASSET_LABEL(string asset_no, string asset_subnumber, string department, string vendorName, DateTime capitalizedOn, string nameOwner, string costCenter, string assetDesc)
        {
            var db = new DatabaseAccessLayer();

            string formattedDate = capitalizedOn.ToString("dd MMMM yyyy");

            asset_no = asset_no == "null" ? string.Empty : asset_no;
            asset_subnumber = asset_subnumber == "null" ? string.Empty : asset_subnumber;
            department = department == "null" ? string.Empty : department;
            vendorName = vendorName == "null" ? string.Empty : vendorName;
            nameOwner = nameOwner == "null" ? string.Empty : nameOwner;
            costCenter = costCenter == "null" ? string.Empty : costCenter;
            assetDesc = assetDesc == "null" ? string.Empty : assetDesc;

            var imagePath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo", "logoblack.png");
            string htmlToPdf = $@"<html>
              <header style='margin-top:-40'>
                <style>
                  body {{
                    font-family: Helvetica;
                    font-weight: bold;
                    font-size: 3.2pt;
                    margin: 0;
                    padding: 0;
                    height: 100%;
                    width: 100%;
                    box-sizing: border-box;
                  }}
                  .logo {{
                    width: 30%;
                    height: auto;
                  }}
                  .header {{
                    display: flex;
                    align-items: center;
                    margin: 0; /* Remove margin */
                    padding: 0; /* Remove padding */
                  }}
                  .header-info {{
                    margin-left: 10px; /* Adjust the space between the logo and the text */
                  }}
                  table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin: 0; /* Remove margin */
                    padding: 0; /* Remove padding */
                  }}
                  td {{
                    padding: 0.2px;
                    vertical-align: top;
                    font-weight: bold;
                  }}
                  td.title {{
                    width: 20px;
                    min-width: 20px;
                    max-width: 20px;
                    white-space: nowrap;
                    padding-right: 1px;
                    font-weight: bold;
                  }}
                  td.value {{
                    width: 125px;
                    min-width: 80px;
                    max-width: 125px;
                    padding-left: 2px;
                    white-space: normal;
                    word-break: break-word;
                    hyphens: auto;
                  }}
                  td.desc {{
                    width: 125px;
                    min-width: 80px;
                    max-width: 125px;
                    padding:2px;
                    font-size:4pt;
                    white-space: normal;
                    word-break: break-word;
                    hyphens: auto;
                    text-align: center;
                  }}
                  .text-center {{
                    text-align: center;
                  }}
                </style>
              </header>
              <body style='margin:0; margin-top:-30; padding:0; height:150%; width:100%;'>
                <div class='header' style='width: 100%; margin-top:-30; margin-left:3px'>
                  <img class='logo' src='file:///{imagePath}'/>
                  <div class='header-info'>
                    <div style='font-weight: bold; font-size:10pt; margin-right:-20px'>{asset_no}{asset_subnumber}</div>
                  </div>
                </div>
                <table style='margin-top:-25px;margin-left:3px'>
                  <tr>
                    <td class='title' style='font-size:3pt'>Asset No.</td>
                    <td class='value' colspan='3'>: {asset_no} / {asset_subnumber}</td>
                  </tr>
                  <tr>
                    <td class='title' style='font-size:3pt'>Department</td>
                    <td class='value' colspan='3'>: {department}</td>
                  </tr>
                  <tr>
                    <td class='title' style='font-size:3pt'>Supplier</td>
                    <td class='value' colspan='3'>: {vendorName}</td>
                  </tr>
                  <tr>
                    <td class='title' style='font-size:3pt'>Cap. Date</td>
                    <td class='value' colspan='3'>: {formattedDate}</td>
                  </tr>
                  <tr>
                    <td class='title' style='font-size:3pt'>Requestor</td>
                    <td class='value' colspan='3'>: {nameOwner}</td>
                  </tr>
                  <tr>
                    <td class='title' style='font-size:3pt'>Cost Center</td>
                    <td class='value' colspan='3'>: {costCenter}</td>
                  </tr>
                  <tr>
                    <td class='desc' colspan='4'>{assetDesc}</td>
                  </tr>
                </table>
              </body>
            </html>";

            float widthCm = 5.4f;
            float heightCm = 4.5f;
            PageSize pageSize = new PageSize(CmToPoints(widthCm), CmToPoints(heightCm));

            using (var memoryStream = new MemoryStream())
            {
                PdfWriter writer = new PdfWriter(memoryStream);
                PdfDocument pdfDocument = new PdfDocument(writer);

                pdfDocument.SetDefaultPageSize(pageSize);

                using (MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(htmlToPdf)))
                {
                    HtmlConverter.ConvertToPdf(htmlStream, pdfDocument);
                }

                pdfDocument.Close();

                Response.Headers["Content-Disposition"] = "inline; filename=label.pdf";
                return File(memoryStream.ToArray(), "application/pdf");
            }
        }

        public string CREATE_GATEPASS(IFormFile file_imgs, string filename, string order_no)
        {
            try
            {
                if (file_imgs.Length > 0)
                {
                    using (var memoryStream = new MemoryStream())
                    {
                        file_imgs.CopyTo(memoryStream);
                        memoryStream.Position = 0;

                        using (var image = SixLabors.ImageSharp.Image.Load(memoryStream))
                        {
                            int maxDimension = 1000;
                            if (image.Width > maxDimension || image.Height > maxDimension)
                            {
                                image.Mutate(x => x.Resize(new ResizeOptions
                                {
                                    Size = new Size(maxDimension, maxDimension),
                                    Mode = ResizeMode.Max
                                }));
                            }

                            using (var finalStream = new MemoryStream())
                            {
                                image.SaveAsJpeg(finalStream, new JpegEncoder { Quality = 100 });
                                finalStream.Position = 0;
                                string filenameSave = filename + ".png";
                                string savePath = System.IO.Path.Combine("wwwroot", "Upload", "GatePass", filenameSave);
                                System.IO.File.WriteAllBytes(savePath, finalStream.ToArray());
                                return "OK;" + filenameSave;
                            }
                        }
                    }
                }
                return "No file uploaded";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string RETURN_GATEPASS(IFormFile file_imgs, string filename, int id_gatepass)
        {
            try
            {
                if (file_imgs.Length > 0)
                {
                    using (var memoryStream = new MemoryStream())
                    {
                        file_imgs.CopyTo(memoryStream);
                        memoryStream.Position = 0;

                        using (var image = SixLabors.ImageSharp.Image.Load(memoryStream))
                        {
                            int maxDimension = 1000;
                            if (image.Width > maxDimension || image.Height > maxDimension)
                            {
                                image.Mutate(x => x.Resize(new ResizeOptions
                                {
                                    Size = new Size(maxDimension, maxDimension),
                                    Mode = ResizeMode.Max
                                }));
                            }

                            using (var finalStream = new MemoryStream())
                            {
                                image.SaveAsJpeg(finalStream, new JpegEncoder { Quality = 100 });
                                finalStream.Position = 0;
                                string filenameSave = filename + ".png";
                                string savePath = System.IO.Path.Combine("wwwroot", "Upload", "GatePass", filenameSave);
                                System.IO.File.WriteAllBytes(savePath, finalStream.ToArray());
                                return filenameSave;
                            }
                        }
                    }
                }
                return "No file uploaded";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public IActionResult PRINT_GATEPASS_PDF(string id_gatepass)
        {
            var db = new DatabaseAccessLayer();
            List<GatePassModel> gpDet = db.GetGatepassInfo(id_gatepass);
            var gp = gpDet.First();
            List<GatePassModel> jumbobagList = db.GET_GATEPASS_DETAIL(int.Parse(id_gatepass));
            List<ApprovalModel> appList = db.GetGatepassApproval(id_gatepass);


            string[] imageList = (gp.image_before ?? "").Split(';');
            StringBuilder htmlBuilder = new StringBuilder();

            foreach (string image in imageList)
            {
                if (!string.IsNullOrEmpty(image))
                {
                    htmlBuilder.AppendFormat("<img src='wwwroot/Upload/GatePass/{0}' alt='Image' style='width: 50%; height:auto; text-align:center; margin-right:5px' />", image);
                }
            }

            string imageOutput = htmlBuilder.ToString();


            string listJumbobag = "";
            if (gp.category == "REPAIR TO VENDOR" || gp.category == "TRANSFER TO VENDOR" || gp.category == "FA SALES - WITHOUT DISMANTLING")
            {
                foreach (var row in jumbobagList)
                {
                    listJumbobag += "<tr>" +
                                    "<td style='color:black;border-bottom: 0.5px solid grey;'>" + row.asset_no + "</td>" +
                                    "<td style='color:black;border-bottom: 0.5px solid grey;'>" + row.asset_subnumber + "</td>" +
                                    "<td style='color:black;border-bottom: 0.5px solid grey;'>" + row.asset_desc + "</td>" +
                                    "</tr>";
                }
            }
            else if (gp.category == "TRANSFER PLANT TO PLANT")
            {
                foreach (var row in jumbobagList)
                {
                    listJumbobag += "<tr>" +
                                    "<td style='color:black;border-bottom: 0.5px solid grey;'>" + row.asset_no + "</td>" +
                                    "<td style='color:black;border-bottom: 0.5px solid grey;'>" + row.asset_subnumber + "</td>" +
                                    "<td style='color:black;border-bottom: 0.5px solid grey;'>" + row.asset_desc + "</td>" +
                                    "<td style='color:black;border-bottom: 0.5px solid grey;'>" + row.naming_output + "</td>" +
                                    "</tr>";
                }
            }

            string listApproval = "";
            foreach (var row in appList)
            {
                string approvalLevelText = "";
                if (row.approval_level == "HOD")
                {
                    approvalLevelText = "HEAD DEPARTMENT";
                }
                else if (row.approval_level == "FBP")
                {
                    approvalLevelText = "FINANCE MANAGER";
                }
                else if (row.approval_level == "PH")
                {
                    approvalLevelText = "PLANT HEAD";
                }
                else
                {
                    approvalLevelText = row.approval_level.ToUpper();
                }
                listApproval = listApproval + $@"<tr>
                                                    <td>{approvalLevelText}</td>
                                                    <td>{row.approval_name}</td>
                                                    <td>APPROVED</td>
                                                </tr>";
            }

            string htmlToPdf = "";
            if (gp.category == "REPAIR TO VENDOR")
            {
                htmlToPdf = $@"<html>
                                      <header>
                                        <style>
                                          body {{
                                            margin-left:20px;
                                            margin-right:20px;
                                            margin-top:0px;
                                            margin-bottom:10px;
                                            font-family:Helvetica ;
                                            font-size:9pt;
                                          }}
                                          .text-top-justify {{
                                            text-align: justify;
                                            text-justify: inter-word;
                                            vertical-align: top;
                                          }}
                                          .bg-green {{
                                            background-color: #009E4C;
                                            color: white;
                                          }}
                                          .text-green {{
                                            color: #009E4C;
                                          }}
                                          #table_jumbobag > tr > td {{
                                            border-bottom: 1px solid black;
                                          }}
                                        </style>
                                      </header>
                                      <body style=''>
                                        <table style='width:100%; text-align:center'>
                                            <tr>
                                                <td colspan='4' style='background-color:#f5f5f5;padding-top:5px;padding-bottom:5px'>
                                                    <img src='wwwroot/images/logo/aml_black.png' style='width: 20%; height:auto; text-align:center; padding-left:30%' />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='font-size:12pt'>GATE PASS {gp.image_before}</td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='font-size:12pt;color:green'>{gp.gatepass_no}</td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='font-size:12pt;color:green'>{gp.category}</td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='text-align:right'>Date : {gp.create_date}</td>
                                            </tr>
                                        </table>
                                        <table style='width:100%; text-align:center' table-bordered cellpadding='4'>
                                            <tr>
                                                <td colspan='4' style='' class='bg-green'>GATE PASS DETAIL</td>
                                            </tr>
                                            <tr style='text-align:left'>
                                                <td style='width:20%'>Deliver To</td>
                                                <td style='width:25%' class='text-green' style='font-size:10pt;'>: {gp.vendor_name}</td>
                                                <td style='width:25%'>Vehicle Police No.</td>
                                                <td style='width:25%' class='text-green'>: {gp.vehicle_no}</td>
                                            </tr>
                                            <tr style='text-align:left'>
                                                <td>Delivery Address</td>
                                                <td class='text-green'>: {gp.vendor_address}</td>
                                                <td>Driver's Name</td>
                                                <td class='text-green'>: {gp.driver_name}</td>
                                            </tr>
                                            <tr style='text-align:left'>
                                                <td>Recipient's Phone No.</td>
                                                <td class='text-green'>: {gp.recipient_phone}</td>
                                                <td>Return Date</td>
                                                <td class='text-green'>: {gp.return_date}</td>
                                            </tr>
                                            <tr style='text-align:left'>
                                                <td>Recipient's Email</td>
                                                <td class='text-green'>: {gp.recipient_email}</td>
                                                <td></td>
                                                <td></td>
                                            </tr>
                                        </table>
                                        <table id='table_jumbobag' style='width:100%;margin-top:20px;text-align:left;border:0.5px solid grey' cellpadding='7' cellspacing='0'>
                                            <tr class='bg-green'>
                                                <td>ASSET NO</td>
                                                <td>SUBNUMBER</td>
                                                <td>DESCRIPTION</td>
                                            </tr>{listJumbobag}
                                        </table>
                                        <table style='width:100%;margin-top:20px' cellpadding='7'>
                                            <tr>    
                                                <td style='width:5%'>Requestor Name</td>
                                                <td style='width:30%' class='text-green'>: {gp.created_by}</td>
                                                <td colspan='2' rowspan='6'>
                                                    <table style='width:100%;text-align:center;border:0.5px solid grey' cellpadding='4' cellspacing='0'>
                                                        <tr class='bg-green'>
                                                            <td>APPROVAL LEVEL</td>
                                                            <td>NAME</td>
                                                            <td>STATUS</td>
                                                        </tr>
                                                        {listApproval}
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style='width:5%'>Security/Guard</td>
                                                <td style='width:30%' class='text-green'>: {gp.security_guard}</td>
                                            </tr>
                                            <tr>
                                                <td style='width:5%'>Remark/Reason</td>
                                                <td style='width:30%' class='text-green'>: {gp.remark}</td>
                                            </tr>
                                        </table>
                                        <table style='width:100%; margin-top:10px'>
                                            <tr>
                                                <td colspan='3' class='text-green' style='font-size:11pt; border-bottom:1px solid #009E4C'></td>
                                            </tr>
                                            <tr>
                                                <td colspan='3' style='width:100%;text-align:left;padding-top:15px;'>
                                                    {imageOutput}
                                                </td>
                                            </tr>
                                        </table>
                                      </body>
                                    </html>";
            }
            else if (gp.category == "TRANSFER TO VENDOR")
            {
                htmlToPdf = $@"<html>
                                      <header>
                                        <style>
                                          body {{
                                            margin-left:20px;
                                            margin-right:20px;
                                            margin-top:0px;
                                            margin-bottom:10px;
                                            font-family:Helvetica ;
                                            font-size:9pt;
                                          }}
                                          .text-top-justify {{
                                            text-align: justify;
                                            text-justify: inter-word;
                                            vertical-align: top;
                                          }}
                                          .bg-green {{
                                            background-color: #009E4C;
                                            color: white;
                                          }}
                                          .text-green {{
                                            color: #009E4C;
                                          }}
                                          #table_jumbobag > tr > td {{
                                            border-bottom: 1px solid black;
                                          }}
                                        </style>
                                      </header>
                                      <body style=''>
                                        <table style='width:100%; text-align:center'>
                                            <tr>
                                                <td colspan='4' style='background-color:#f5f5f5;padding-top:5px;padding-bottom:5px'>
                                                    <img src='wwwroot/images/logo/aml_black.png' style='width: 20%; height:auto; text-align:center; padding-left:30%' />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='font-size:12pt'>GATE PASS</td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='font-size:12pt;color:green'>{gp.gatepass_no}</td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='font-size:12pt;color:green'>{gp.category}</td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='text-align:right'>Date : {gp.create_date}</td>
                                            </tr>
                                        </table>
                                        <table style='width:100%; text-align:center' table-bordered cellpadding='4'>
                                            <tr>
                                                <td colspan='4' style='' class='bg-green'>GATE PASS DETAIL</td>
                                            </tr>
                                            <tr style='text-align:left'>
                                                <td style='width:20%'>Deliver To</td>
                                                <td style='width:25%' class='text-green' style='font-size:10pt;'>: {gp.vendor_name}</td>
                                                <td style='width:25%'>Vehicle Police No.</td>
                                                <td style='width:25%' class='text-green'>: {gp.vehicle_no}</td>
                                            </tr>
                                            <tr style='text-align:left'>
                                                <td>Delivery Address</td>
                                                <td class='text-green'>: {gp.vendor_address}</td>
                                                <td>Driver's Name</td>
                                                <td class='text-green'>: {gp.driver_name}</td>
                                            </tr>
                                            <tr style='text-align:left'>
                                                <td>Recipient's Phone No.</td>
                                                <td class='text-green'>: {gp.recipient_phone}</td>
                                                <td></td>
                                                <td></td>
                                            </tr>
                                            <tr style='text-align:left'>
                                                <td>Recipient's Email</td>
                                                <td class='text-green'>: {gp.recipient_email}</td>
                                                <td></td>
                                                <td></td>
                                            </tr>
                                        </table>
                                        <table id='table_jumbobag' style='width:100%;margin-top:20px;text-align:left;border:0.5px solid grey' cellpadding='7' cellspacing='0'>
                                            <tr class='bg-green'>
                                                <td>ASSET NO</td>
                                                <td>SUBNUMBER</td>
                                                <td>DESCRIPTION</td>
                                            </tr>{listJumbobag}
                                        </table>
                                            <table style='width:100%;margin-top:20px' cellpadding='7'>
                                                <tr>    
                                                    <td style='width:5%'>Requestor Name</td>
                                                    <td style='width:30%' class='text-green'>: {gp.created_by}</td>
                                                    <td colspan='2' rowspan='6'>
                                                        <table style='width:100%;text-align:center;border:0.5px solid grey' cellpadding='4' cellspacing='0'>
                                                            <tr class='bg-green'>
                                                                <td>APPROVAL LEVEL</td>
                                                                <td>NAME</td>
                                                                <td>STATUS</td>
                                                            </tr>
                                                            {listApproval}
                                                        </table>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td style='width:5%'>Security/Guard</td>
                                                    <td style='width:30%' class='text-green'>: {gp.security_guard}</td>
                                                </tr>
                                                <tr>
                                                    <td style='width:5%'>Remark/Reason</td>
                                                    <td style='width:30%' class='text-green'>: {gp.remark}</td>
                                                </tr>
                                            </table>
                                        <table style='width:100%; margin-top:10px'>
                                            <tr>
                                                <td colspan='3' class='text-green' style='font-size:11pt; border-bottom:1px solid #009E4C'></td>
                                            </tr>
                                            <tr>
                                                <td colspan='3' style='width:100%;text-align:left;padding-top:15px;'>
                                                    {imageOutput}
                                                </td>
                                            </tr>
                                        </table>
                                      </body>
                                    </html>";
            }
            else if (gp.category == "TRANSFER PLANT TO PLANT")
            {
                htmlToPdf = $@"<html>
                                      <header>
                                        <style>
                                          body {{
                                            margin-left:20px;
                                            margin-right:20px;
                                            margin-top:0px;
                                            margin-bottom:10px;
                                            font-family:Helvetica ;
                                            font-size:9pt;
                                          }}
                                          .text-top-justify {{
                                            text-align: justify;
                                            text-justify: inter-word;
                                            vertical-align: top;
                                          }}
                                          .bg-green {{
                                            background-color: #009E4C;
                                            color: white;
                                          }}
                                          .text-green {{
                                            color: #009E4C;
                                          }}
                                          #table_jumbobag > tr > td {{
                                            border-bottom: 1px solid black;
                                          }}
                                        </style>
                                      </header>
                                      <body style=''>
                                        <table style='width:100%; text-align:center'>
                                            <tr>
                                                <td colspan='4' style='background-color:#f5f5f5;padding-top:5px;padding-bottom:5px'>
                                                    <img src='wwwroot/images/logo/aml_black.png' style='width: 20%; height:auto; text-align:center; padding-left:30%' />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='font-size:12pt'>GATE PASS</td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='font-size:12pt;color:green'>{gp.gatepass_no}</td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='font-size:12pt;color:green'>{gp.category}</td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='text-align:right'>Date : {gp.create_date}</td>
                                            </tr>
                                        </table>
                                        <table style='width:100%; text-align:center' table-bordered cellpadding='4'>
                                            <tr>
                                                <td colspan='4' style='' class='bg-green'>GATE PASS DETAIL</td>
                                            </tr>
                                            <tr style='text-align:left'>
                                                <td style='width:20%'>Deliver To</td>
                                                <td style='width:25%' class='text-green' style='font-size:10pt;'>: {gp.location}</td>
                                                <td style='width:25%'>Vehicle Police No.</td>
                                                <td style='width:25%' class='text-green'>: {gp.vehicle_no}</td>
                                            </tr>
                                            <tr style='text-align:left'>
                                                <td>Driver's Name</td>
                                                <td class='text-green'>: {gp.driver_name}</td>
                                                <td style='width:25%'>New Pic</td>
                                                <td style='width:25%' class='text-green'>: {gp.new_pic}</td>
                                            </tr>
                                        </table>
                                        <table id='table_jumbobag' style='width:100%;margin-top:20px;text-align:left;border:0.5px solid grey' cellpadding='7' cellspacing='0'>
                                            <tr class='bg-green'>
                                                <td style='width:15%'>ASSET NO</td>
                                                <td style='width:5%'>SUBNUMBER</td>
                                                <td style='width:30%'>DESCRIPTION</td>
                                                <td style='width:50%'>NEW NAMING</td>
                                            </tr>{listJumbobag}
                                        </table>
                                        <table style='width:100%;margin-top:20px' cellpadding='7'>
                                                <tr>    
                                                    <td style='width:5%'>Requestor Name</td>
                                                    <td style='width:30%' class='text-green'>: {gp.created_by}</td>
                                                    <td colspan='2' rowspan='6'>
                                                        <table style='width:100%;text-align:center;border:0.5px solid grey' cellpadding='4' cellspacing='0'>
                                                            <tr class='bg-green'>
                                                                <td>APPROVAL LEVEL</td>
                                                                <td>NAME</td>
                                                                <td>STATUS</td>
                                                            </tr>
                                                            {listApproval}
                                                        </table>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td style='width:5%'>Security/Guard</td>
                                                    <td style='width:30%' class='text-green'>: {gp.security_guard}</td>
                                                </tr>
                                                <tr>
                                                    <td style='width:5%'>Remark/Reason</td>
                                                    <td style='width:30%' class='text-green'>: {gp.remark}</td>
                                                </tr>
                                            </table>
                                        <table style='width:100%; margin-top:10px'>
                                            <tr>
                                                <td colspan='3' class='text-green' style='font-size:11pt; border-bottom:1px solid #009E4C'></td>
                                            </tr>
                                            <tr>
                                                <td colspan='3' style='width:100%;text-align:left;padding-top:15px;'>
                                                    {imageOutput}
                                                </td>
                                            </tr>
                                        </table>
                                      </body>
                                    </html>";
            }
            else if (gp.category == "FA SALES - WITHOUT DISMANTLING" && gp.type == "VENDOR")
            {
                htmlToPdf = $@"<html>
                                      <header>
                                        <style>
                                          body {{
                                            margin-left:20px;
                                            margin-right:20px;
                                            margin-top:0px;
                                            margin-bottom:10px;
                                            font-family:Helvetica ;
                                            font-size:9pt;
                                          }}
                                          .text-top-justify {{
                                            text-align: justify;
                                            text-justify: inter-word;
                                            vertical-align: top;
                                          }}
                                          .bg-green {{
                                            background-color: #009E4C;
                                            color: white;
                                          }}
                                          .text-green {{
                                            color: #009E4C;
                                          }}
                                          #table_jumbobag > tr > td {{
                                            border-bottom: 1px solid black;
                                          }}
                                        </style>
                                      </header>
                                      <body style=''>
                                        <table style='width:100%; text-align:center'>
                                            <tr>
                                                <td colspan='4' style='background-color:#f5f5f5;padding-top:5px;padding-bottom:5px'>
                                                    <img src='wwwroot/images/logo/aml_black.png' style='width: 20%; height:auto; text-align:center; padding-left:30%' />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='font-size:12pt'>GATE PASS</td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='font-size:12pt;color:green'>{gp.gatepass_no}</td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='font-size:12pt;color:green'>{gp.category}</td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='text-align:right'>Date : {gp.create_date}</td>
                                            </tr>
                                        </table>
                                        <table style='width:100%; text-align:center' table-bordered cellpadding='4'>
                                            <tr>
                                                <td colspan='4' style='' class='bg-green'>GATE PASS DETAIL</td>
                                            </tr>
                                            <tr style='text-align:left'>
                                                <td style='width:20%'>Deliver To</td>
                                                <td style='width:25%' class='text-green' style='font-size:10pt;'>: {gp.vendor_name}</td>
                                                <td style='width:20%'>Type</td>
                                                <td style='width:25%' class='text-green' style='font-size:10pt;'>: {gp.type}</td>
                                            </tr>
                                            <tr style='text-align:left'>
                                                <td>Delivery Address</td>
                                                <td class='text-green'>: {gp.vendor_address}</td>
                                                <td>Driver's Name</td>
                                                <td class='text-green'>: {gp.driver_name}</td>
                                            </tr>
                                            <tr style='text-align:left'>
                                                <td>Recipient's Phone No.</td>
                                                <td class='text-green'>: {gp.recipient_phone}</td>
                                                <td style='width:25%'>Vehicle Police No.</td>
                                                <td style='width:25%' class='text-green'>: {gp.vehicle_no}</td>
                                            </tr>
                                            <tr style='text-align:left'>
                                                <td>Recipient's Email</td>
                                                <td class='text-green'>: {gp.recipient_email}</td>
                                                <td></td>
                                                <td></td>
                                            </tr>
                                        </table>
                                        <table id='table_jumbobag' style='width:100%;margin-top:20px;text-align:left;border:0.5px solid grey' cellpadding='7' cellspacing='0'>
                                            <tr class='bg-green'>
                                                <td>ASSET NO</td>
                                                <td>SUBNUMBER</td>
                                                <td>DESCRIPTION</td>
                                            </tr>{listJumbobag}
                                        </table>
                                            <table style='width:100%;margin-top:20px' cellpadding='7'>
                                                <tr>    
                                                    <td style='width:5%'>Requestor Name</td>
                                                    <td style='width:30%' class='text-green'>: {gp.created_by}</td>
                                                    <td colspan='2' rowspan='6'>
                                                        <table style='width:100%;text-align:center;border:0.5px solid grey' cellpadding='4' cellspacing='0'>
                                                            <tr class='bg-green'>
                                                                <td>APPROVAL LEVEL</td>
                                                                <td>NAME</td>
                                                                <td>STATUS</td>
                                                            </tr>
                                                            {listApproval}
                                                        </table>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td style='width:5%'>Security/Guard</td>
                                                    <td style='width:30%' class='text-green'>: {gp.security_guard}</td>
                                                </tr>
                                                <tr>
                                                    <td style='width:5%'>Remark/Reason</td>
                                                    <td style='width:30%' class='text-green'>: {gp.remark}</td>
                                                </tr>
                                            </table>
                                        <table style='width:100%; margin-top:10px'>
                                            <tr>
                                                <td colspan='3' class='text-green' style='font-size:11pt; border-bottom:1px solid #009E4C'></td>
                                            </tr>
                                            <tr>
                                                <td colspan='3' style='width:100%;text-align:left;padding-top:15px;'>
                                                    {imageOutput}
                                                </td>
                                            </tr>
                                        </table>
                                      </body>
                                    </html>";
            }
            else if (gp.category == "FA SALES - WITHOUT DISMANTLING" && gp.type == "EMPLOYEE")
            {
                htmlToPdf = $@"<html>
                                      <header>
                                        <style>
                                          body {{
                                            margin-left:20px;
                                            margin-right:20px;
                                            margin-top:0px;
                                            margin-bottom:10px;
                                            font-family:Helvetica ;
                                            font-size:9pt;
                                          }}
                                          .text-top-justify {{
                                            text-align: justify;
                                            text-justify: inter-word;
                                            vertical-align: top;
                                          }}
                                          .bg-green {{
                                            background-color: #009E4C;
                                            color: white;
                                          }}
                                          .text-green {{
                                            color: #009E4C;
                                          }}
                                          #table_jumbobag > tr > td {{
                                            border-bottom: 1px solid black;
                                          }}
                                        </style>
                                      </header>
                                      <body style=''>
                                        <table style='width:100%; text-align:center'>
                                            <tr>
                                                <td colspan='4' style='background-color:#f5f5f5;padding-top:5px;padding-bottom:5px'>
                                                    <img src='wwwroot/images/logo/aml_black.png' style='width: 20%; height:auto; text-align:center; padding-left:30%' />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='font-size:12pt'>GATE PASS</td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='font-size:12pt;color:green'>{gp.gatepass_no}</td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='font-size:12pt;color:green'>{gp.category}</td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='text-align:right'>Date : {gp.create_date}</td>
                                            </tr>
                                        </table>
                                        <table style='width:100%; text-align:center' table-bordered cellpadding='4'>
                                            <tr>
                                                <td colspan='4' style='' class='bg-green'>GATE PASS DETAIL</td>
                                            </tr>
                                            <tr style='text-align:left'>
                                                <td style='width:20%'>Type</td>
                                                <td style='width:25%' class='text-green' style='font-size:10pt;'>: {gp.type}</td>
                                                <td style='width:20%'>Employee Name</td>
                                                <td style='width:25%' class='text-green' style='font-size:10pt;'>: {gp.employee_name}</td>
                                            </tr>
                                        </table>
                                        <table id='table_jumbobag' style='width:100%;margin-top:20px;text-align:left;border:0.5px solid grey' cellpadding='7' cellspacing='0'>
                                            <tr class='bg-green'>
                                                <td>ASSET NO</td>
                                                <td>SUBNUMBER</td>
                                                <td>DESCRIPTION</td>
                                            </tr>{listJumbobag}
                                        </table>
                                            <table style='width:100%;margin-top:20px' cellpadding='7'>
                                                <tr>    
                                                    <td style='width:5%'>Requestor Name</td>
                                                    <td style='width:30%' class='text-green'>: {gp.created_by}</td>
                                                    <td colspan='2' rowspan='6'>
                                                        <table style='width:100%;text-align:center;border:0.5px solid grey' cellpadding='4' cellspacing='0'>
                                                            <tr class='bg-green'>
                                                                <td>APPROVAL LEVEL</td>
                                                                <td>NAME</td>
                                                            </tr>
                                                            {listApproval}
                                                        </table>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td style='width:5%'>Security/Guard</td>
                                                    <td style='width:30%' class='text-green'>: {gp.security_guard}</td>
                                                </tr>
                                                <tr>
                                                    <td style='width:5%'>Remark/Reason</td>
                                                    <td style='width:30%' class='text-green'>: {gp.remark}</td>
                                                </tr>
                                            </table>
                                        <table style='width:100%; margin-top:10px'>
                                            <tr>
                                                <td colspan='3' class='text-green' style='font-size:11pt; border-bottom:1px solid #009E4C'></td>
                                            </tr>
                                            <tr>
                                                <td colspan='3' style='width:100%;text-align:left;padding-top:15px;'>
                                                    {imageOutput}
                                                </td>
                                            </tr>
                                        </table>
                                      </body>
                                    </html>";
            }
            else if (gp.category == "FA SALES - WITHOUT DISMANTLING" && gp.type == "CSR")
            {
                htmlToPdf = $@"<html>
                                      <header>
                                        <style>
                                          body {{
                                            margin-left:20px;
                                            margin-right:20px;
                                            margin-top:0px;
                                            margin-bottom:10px;
                                            font-family:Helvetica ;
                                            font-size:9pt;
                                          }}
                                          .text-top-justify {{
                                            text-align: justify;
                                            text-justify: inter-word;
                                            vertical-align: top;
                                          }}
                                          .bg-green {{
                                            background-color: #009E4C;
                                            color: white;
                                          }}
                                          .text-green {{
                                            color: #009E4C;
                                          }}
                                          #table_jumbobag > tr > td {{
                                            border-bottom: 1px solid black;
                                          }}
                                        </style>
                                      </header>
                                      <body style=''>
                                        <table style='width:100%; text-align:center'>
                                            <tr>
                                                <td colspan='4' style='background-color:#f5f5f5;padding-top:5px;padding-bottom:5px'>
                                                    <img src='wwwroot/images/logo/aml_black.png' style='width: 20%; height:auto; text-align:center; padding-left:30%' />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='font-size:12pt'>GATE PASS</td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='font-size:12pt;color:green'>{gp.gatepass_no}</td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='font-size:12pt;color:green'>{gp.category}</td>
                                            </tr>
                                            <tr>
                                                <td colspan='4' style='text-align:right'>Date : {gp.create_date}</td>
                                            </tr>
                                        </table>
                                        <table style='width:100%; text-align:center' table-bordered cellpadding='4'>
                                            <tr>
                                                <td colspan='4' style='' class='bg-green'>GATE PASS DETAIL</td>
                                            </tr>
                                            <tr style='text-align:left'>
                                                <td style='width:20%'>Type</td>
                                                <td style='width:25%' class='text-green' style='font-size:10pt;'>: {gp.type}</td>
                                                <td style='width:20%'>Donation To</td>
                                                <td style='width:25%' class='text-green' style='font-size:10pt;'>: {gp.csr_to}</td>
                                            </tr>
                                        </table>
                                        <table id='table_jumbobag' style='width:100%;margin-top:20px;text-align:left;border:0.5px solid grey' cellpadding='7' cellspacing='0'>
                                            <tr class='bg-green'>
                                                <td>ASSET NO</td>
                                                <td>SUBNUMBER</td>
                                                <td>DESCRIPTION</td>
                                            </tr>{listJumbobag}
                                        </table>
                                            <table style='width:100%;margin-top:20px' cellpadding='7'>
                                                <tr>    
                                                    <td style='width:5%'>Requestor Name</td>
                                                    <td style='width:30%' class='text-green'>: {gp.created_by}</td>
                                                    <td colspan='2' rowspan='6'>
                                                        <table style='width:100%;text-align:center;border:0.5px solid grey' cellpadding='4' cellspacing='0'>
                                                            <tr class='bg-green'>
                                                                <td>APPROVAL LEVEL</td>
                                                                <td>NAME</td>
                                                            </tr>
                                                            {listApproval}
                                                        </table>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td style='width:5%'>Security/Guard</td>
                                                    <td style='width:30%' class='text-green'>: {gp.security_guard}</td>
                                                </tr>
                                                <tr>
                                                    <td style='width:5%'>Remark/Reason</td>
                                                    <td style='width:30%' class='text-green'>: {gp.remark}</td>
                                                </tr>
                                            </table>
                                        <table style='width:100%; margin-top:10px'>
                                            <tr>
                                                <td colspan='3' class='text-green' style='font-size:11pt; border-bottom:1px solid #009E4C'></td>
                                            </tr>
                                            <tr>
                                                <td colspan='3' style='width:100%;text-align:left;padding-top:15px;'>
                                                    {imageOutput}
                                                </td>
                                            </tr>
                                        </table>
                                      </body>
                                    </html>";
            }


            MemoryStream stream = new MemoryStream();
            HtmlConverter.ConvertToPdf(htmlToPdf, stream);

            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            PdfWriter writer = new PdfWriter(byteArrayOutputStream);
            PdfDocument pdfDocument = new PdfDocument(writer);

            MemoryStream stream2 = new MemoryStream(stream.ToArray());
            using (var inputPdfDoc = new PdfDocument(new PdfReader(stream2)))
            {
                inputPdfDoc.CopyPagesTo(1, inputPdfDoc.GetNumberOfPages(), pdfDocument);
            }
            pdfDocument.Close();
            string filename = $"{gp.gatepass_no}.pdf";
            Response.Headers["Content-Disposition"] = $"inline; filename={filename}";
            return File(byteArrayOutputStream.ToArray(), "application/pdf");
        }

        public IActionResult PRINT_GATEPASS_NON_ASSET_PDF(string id_gatepass)
        {
            var db = new DatabaseAccessLayer();
            List<GatePassNonAssetModel> gpDet = db.GetGatepassNonAssetInfo(id_gatepass);
            var gp = gpDet.First();
            List<GatePassNonAssetModel> jumbobagList = db.GET_GATEPASS_NON_ASSET_DETAIL(int.Parse(id_gatepass));
            List<ApprovalModel> appList = db.GetGatepassNonAssetApproval(id_gatepass);

            string[] imageList = (gp.image_before ?? "").Split(';');
            StringBuilder htmlBuilder = new StringBuilder();

            foreach (string image in imageList)
            {
                if (!string.IsNullOrEmpty(image))
                {
                    htmlBuilder.AppendFormat("<img src='wwwroot/Upload/GatePass_Non_Asset/{0}' alt='Image' style='width: 50%; height:auto; text-align:center; margin-right:5px' />", image);
                }
            }

            string imageOutput = htmlBuilder.ToString();

            string listJumbobag = "";
            foreach (var row in jumbobagList)
            {
                listJumbobag += "<tr>" +
                                "<td style='color:black;border-bottom: 0.5px solid grey;'>" + row.po_or_asset_no + "</td>" +
                                "<td style='color:black;border-bottom: 0.5px solid grey;'>" + row.description + "</td>" +
                                "<td style='color:black;border-bottom: 0.5px solid grey;'>" + row.qty + "</td>" +
                                "<td style='color:black;border-bottom: 0.5px solid grey;'>" + row.uom + "</td>" +
                                "</tr>";
            }

            string listApproval = "";
            foreach (var row in appList)
            {
                string approvalLevelText = "";
                if (row.approval_level == "HOD")
                {
                    approvalLevelText = "HEAD DEPARTMENT";
                }
                else if (row.approval_level == "FBP")
                {
                    approvalLevelText = "FINANCE MANAGER";
                }
                else if (row.approval_level == "PH")
                {
                    approvalLevelText = "PLANT HEAD";
                }
                else
                {
                    approvalLevelText = row.approval_level.ToUpper();
                }
                listApproval += $@"<tr>
                            <td>{approvalLevelText}</td>
                            <td>{row.approval_name}</td>
                            <td>APPROVED</td>
                        </tr>";
            }

            string htmlToPdf = $@"<html>
                              <header>
                                <style>
                                  body {{
                                    margin-left:20px;
                                    margin-right:20px;
                                    margin-top:0px;
                                    margin-bottom:10px;
                                    font-family:Helvetica ;
                                    font-size:9pt;
                                  }}
                                  .text-top-justify {{
                                    text-align: justify;
                                    text-justify: inter-word;
                                    vertical-align: top;
                                  }}
                                  .bg-green {{
                                    background-color: #009E4C;
                                    color: white;
                                  }}
                                  .text-green {{
                                    color: #009E4C;
                                  }}
                                  #table_jumbobag > tr > td {{
                                    border-bottom: 1px solid black;
                                  }}
                                </style>
                              </header>
                              <body style=''>
                                <table style='width:100%; text-align:center'>
                                    <tr>
                                        <td colspan='4' style='background-color:#f5f5f5;padding-top:5px;padding-bottom:5px'>
                                            <img src='wwwroot/images/logo/aml_black.png' style='width: 20%; height:auto; text-align:center; padding-left:30%' />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td colspan='4' style='font-size:12pt'>GATE PASS NON-ASSET</td>
                                    </tr>
                                    <tr>
                                        <td colspan='4' style='font-size:12pt;color:green'>{gp.gatepass_no}</td>
                                    </tr>
                                    <tr>
                                        <td colspan='4' style='font-size:12pt;color:green'>{gp.category}</td>
                                    </tr>
                                    <tr>
                                        <td colspan='4' style='text-align:right'>Date : {gp.create_date}</td>
                                    </tr>
                                </table>
                                <table style='width:100%; text-align:center' table-bordered cellpadding='4'>
                                    <tr>
                                        <td colspan='4' style='' class='bg-green'>GATE PASS DETAIL</td>
                                    </tr>
                                    <tr style='text-align:left'>
                                        <td style='width:20%'>Deliver To</td>
                                        <td style='width:25%' class='text-green' style='font-size:10pt;'>: {gp.vendor_name}</td>
                                        <td style='width:25%'>Vehicle Police No.</td>
                                        <td style='width:25%' class='text-green'>: {gp.vehicle_no}</td>
                                    </tr>
                                    <tr style='text-align:left'>
                                        <td>Delivery Address</td>
                                        <td class='text-green'>: {gp.vendor_address}</td>
                                        <td>Driver's Name</td>
                                        <td class='text-green'>: {gp.driver_name}</td>
                                    </tr>
                                    <tr style='text-align:left'>
                                        <td>Recipient's Phone No.</td>
                                        <td class='text-green'>: {gp.recipient_phone}</td>
                                        <td>Return Date</td>
                                        <td class='text-green'>: {gp.return_date}</td>
                                    </tr>
                                    <tr style='text-align:left'>
                                        <td>Recipient's Email</td>
                                        <td class='text-green'>: {gp.recipient_email}</td>
                                        <td></td>
                                        <td></td>
                                    </tr>
                                </table>
                                <table id='table_jumbobag' style='width:100%;margin-top:20px;text-align:left;border:0.5px solid grey' cellpadding='7' cellspacing='0'>
                                    <tr class='bg-green'>
                                        <td>PO/ASSET NO</td>
                                        <td>DESCRIPTION</td>
                                        <td>QTY</td>
                                        <td>UOM</td>
                                    </tr>{listJumbobag}
                                </table>
                                <table style='width:100%;margin-top:20px' cellpadding='7'>
                                    <tr>    
                                        <td style='width:5%'>Requestor Name</td>
                                        <td style='width:30%' class='text-green'>: {gp.created_by}</td>
                                        <td colspan='2' rowspan='6'>
                                            <table style='width:100%;text-align:center;border:0.5px solid grey' cellpadding='4' cellspacing='0'>
                                                <tr class='bg-green'>
                                                    <td>APPROVAL LEVEL</td>
                                                    <td>NAME</td>
                                                    <td>STATUS</td>
                                                </tr>
                                                {listApproval}
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style='width:5%'>Security/Guard</td>
                                        <td style='width:30%' class='text-green'>: {gp.security_guard}</td>
                                    </tr>
                                    <tr>
                                        <td style='width:5%'>Remark/Reason</td>
                                        <td style='width:30%' class='text-green'>: {gp.remark}</td>
                                    </tr>
                                </table>
                                <table style='width:100%; margin-top:10px'>
                                    <tr>
                                        <td colspan='3' class='text-green' style='font-size:11pt; border-bottom:1px solid #009E4C'></td>
                                    </tr>
                                    <tr>
                                        <td colspan='3' style='width:100%;text-align:left;padding-top:15px;'>
                                            {imageOutput}
                                        </td>
                                    </tr>
                                </table>
                              </body>
                            </html>";

            MemoryStream stream = new MemoryStream();
            HtmlConverter.ConvertToPdf(htmlToPdf, stream);

            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            PdfWriter writer = new PdfWriter(byteArrayOutputStream);
            PdfDocument pdfDocument = new PdfDocument(writer);

            MemoryStream stream2 = new MemoryStream(stream.ToArray());
            using (var inputPdfDoc = new PdfDocument(new PdfReader(stream2)))
            {
                inputPdfDoc.CopyPagesTo(1, inputPdfDoc.GetNumberOfPages(), pdfDocument);
            }
            pdfDocument.Close();
            string filename = $"{gp.gatepass_no}.pdf";
            Response.Headers["Content-Disposition"] = $"inline; filename={filename}";
            return File(byteArrayOutputStream.ToArray(), "application/pdf");
        }

        public string CREATE_IMAGE_GP(IFormFile file_imgs, string filename)
        {
            try
            {
                if (file_imgs.Length > 0)
                {
                    using (var memoryStream = new MemoryStream())
                    {
                        file_imgs.CopyTo(memoryStream);
                        memoryStream.Position = 0;

                        using (var image = SixLabors.ImageSharp.Image.Load(memoryStream))
                        {
                            int maxDimension = 1000;
                            if (image.Width > maxDimension || image.Height > maxDimension)
                            {
                                image.Mutate(x => x.Resize(new ResizeOptions
                                {
                                    Size = new Size(maxDimension, maxDimension),
                                    Mode = ResizeMode.Max
                                }));
                            }

                            using (var finalStream = new MemoryStream())
                            {
                                image.SaveAsJpeg(finalStream, new JpegEncoder { Quality = 100 });
                                finalStream.Position = 0;
                                string filenameSave = filename + ".png";
                                string savePath = System.IO.Path.Combine("wwwroot", "Upload", "GatePass", filenameSave);
                                System.IO.File.WriteAllBytes(savePath, finalStream.ToArray());
                                return "OK;" + filenameSave;
                            }
                        }
                    }
                }
                return "No file uploaded";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string CREATE_IMAGE_GP_NA(IFormFile file_imgs, string filename)
        {
            try
            {
                if (file_imgs.Length > 0)
                {
                    using (var memoryStream = new MemoryStream())
                    {
                        file_imgs.CopyTo(memoryStream);
                        memoryStream.Position = 0;

                        using (var image = SixLabors.ImageSharp.Image.Load(memoryStream))
                        {
                            int maxDimension = 1000;
                            if (image.Width > maxDimension || image.Height > maxDimension)
                            {
                                image.Mutate(x => x.Resize(new ResizeOptions
                                {
                                    Size = new Size(maxDimension, maxDimension),
                                    Mode = ResizeMode.Max
                                }));
                            }

                            using (var finalStream = new MemoryStream())
                            {
                                image.SaveAsJpeg(finalStream, new JpegEncoder { Quality = 100 });
                                finalStream.Position = 0;
                                string baseFilename = filename + ".png";
                                string savePath = System.IO.Path.Combine("wwwroot", "Upload", "GatePass_Non_Asset", baseFilename);
                                int counter = 1;
                                while (System.IO.File.Exists(savePath))
                                {
                                    string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(baseFilename);
                                    string extension = System.IO.Path.GetExtension(baseFilename);
                                    baseFilename = $"{fileNameWithoutExt}({counter}){extension}";
                                    savePath = System.IO.Path.Combine("wwwroot", "Upload", "GatePass_Non_Asset", baseFilename);
                                    counter++;
                                }
                                System.IO.File.WriteAllBytes(savePath, finalStream.ToArray());
                                return "OK;" + baseFilename;
                            }
                        }
                    }
                }
                return "No file uploaded";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string RETURN_GATEPASS_NON_ASSET(IFormFile file_imgs, string filename, int id_gatepass)
        {
            try
            {
                if (file_imgs.Length > 0)
                {
                    using (var memoryStream = new MemoryStream())
                    {
                        file_imgs.CopyTo(memoryStream);
                        memoryStream.Position = 0;

                        using (var image = SixLabors.ImageSharp.Image.Load(memoryStream))
                        {
                            int maxDimension = 1000;
                            if (image.Width > maxDimension || image.Height > maxDimension)
                            {
                                image.Mutate(x => x.Resize(new ResizeOptions
                                {
                                    Size = new Size(maxDimension, maxDimension),
                                    Mode = ResizeMode.Max
                                }));
                            }

                            using (var finalStream = new MemoryStream())
                            {
                                image.SaveAsJpeg(finalStream, new JpegEncoder { Quality = 100 });
                                finalStream.Position = 0;
                                string filenameSave = filename + ".png";
                                string savePath = System.IO.Path.Combine("wwwroot", "Upload", "GatePass_Non_Asset", filenameSave);
                                System.IO.File.WriteAllBytes(savePath, finalStream.ToArray());
                                return filenameSave;
                            }
                        }
                    }
                }
                return "No file uploaded";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string UPLOAD_PROFORMA_FILES(IFormFile file_attach, IFormFileCollection file_support)
        {
            try
            {
                string uploadPath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Upload/Proforma");
                if (!Directory.Exists(uploadPath))
                {
                    Directory.CreateDirectory(uploadPath);
                }

                string file_attach_name = "";
                string file_support_names = "";

                if (file_attach != null && file_attach.Length > 0)
                {
                    string originalName = System.IO.Path.GetFileNameWithoutExtension(file_attach.FileName);
                    string extension = System.IO.Path.GetExtension(file_attach.FileName);
                    string fileName = file_attach.FileName;

                    int counter = 1;
                    string finalFileName = fileName;
                    while (System.IO.File.Exists(System.IO.Path.Combine(uploadPath, finalFileName)))
                    {
                        finalFileName = $"{originalName}({counter}){extension}";
                        counter++;
                    }

                    string filePath = System.IO.Path.Combine(uploadPath, finalFileName);
                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        file_attach.CopyTo(stream);
                    }
                    file_attach_name = finalFileName;
                }

                if (file_support != null && file_support.Count > 0)
                {
                    List<string> supportFileNames = new List<string>();
                    foreach (var supportFile in file_support.Take(2))
                    {
                        string originalName = System.IO.Path.GetFileNameWithoutExtension(supportFile.FileName);
                        string extension = System.IO.Path.GetExtension(supportFile.FileName);
                        string fileName = supportFile.FileName;

                        int counter = 1;
                        string finalFileName = fileName;
                        while (System.IO.File.Exists(System.IO.Path.Combine(uploadPath, finalFileName)))
                        {
                            finalFileName = $"{originalName}({counter}){extension}";
                            counter++;
                        }

                        string filePath = System.IO.Path.Combine(uploadPath, finalFileName);
                        using (var stream = new FileStream(filePath, FileMode.Create))
                        {
                            supportFile.CopyTo(stream);
                        }
                        supportFileNames.Add(finalFileName);
                    }
                    file_support_names = string.Join(";", supportFileNames);
                }

                return $"OK;{file_attach_name};{file_support_names}";
            }
            catch (Exception ex)
            {
                return $"ERROR;{ex.Message}";
            }
        }

        public string UPLOAD_SUPPORTING_DOCUMENTS(IFormFileCollection supporting_documents)
        {
            try
            {
                string uploadPath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Upload/Gatepass_Fa_Sales");
                if (!Directory.Exists(uploadPath))
                {
                    Directory.CreateDirectory(uploadPath);
                }

                if (supporting_documents == null || supporting_documents.Count == 0)
                {
                    return "ERROR;No files uploaded";
                }

                if (supporting_documents.Count < 1 || supporting_documents.Count > 3)
                {
                    return "ERROR;Must upload between 1 and 3 files";
                }

                List<string> uploadedFileNames = new List<string>();

                foreach (var file in supporting_documents)
                {
                    string originalName = System.IO.Path.GetFileNameWithoutExtension(file.FileName);
                    string extension = System.IO.Path.GetExtension(file.FileName);
                    string fileName = file.FileName;

                    int counter = 10;
                    string finalFileName = fileName;
                    while (System.IO.File.Exists(System.IO.Path.Combine(uploadPath, finalFileName)))
                    {
                        finalFileName = $"{originalName}({counter}){extension}";
                        counter++;
                    }

                    string filePath = System.IO.Path.Combine(uploadPath, finalFileName);
                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        file.CopyTo(stream);
                    }
                    uploadedFileNames.Add(finalFileName);
                }

                return $"OK;{string.Join(";", uploadedFileNames)}";
            }
            catch (Exception ex)
            {
                return $"ERROR;{ex.Message}";
            }
        }

        public string UPLOAD_PROFORMA_FIN_FILES(IFormFile file, string file_type)
        {
            try
            {
                string uploadPath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Upload", "Proforma_Finance");
                if (!Directory.Exists(uploadPath))
                {
                    Directory.CreateDirectory(uploadPath);
                }

                if (file != null && file.Length > 0)
                {
                    string originalName = System.IO.Path.GetFileNameWithoutExtension(file.FileName);
                    string extension = System.IO.Path.GetExtension(file.FileName);
                    string fileName = file.FileName;

                    int counter = 1;
                    string finalFileName = fileName;
                    while (System.IO.File.Exists(System.IO.Path.Combine(uploadPath, finalFileName)))
                    {
                        finalFileName = $"{originalName}({counter}){extension}";
                        counter++;
                    }

                    string filePath = System.IO.Path.Combine(uploadPath, finalFileName);
                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        file.CopyTo(stream);
                    }

                    return $"OK;{finalFileName}";
                }
                return "ERROR;No file selected";
            }
            catch (Exception ex)
            {
                return $"ERROR;{ex.Message}";
            }
        }
        public string UPLOAD_PEB_FILE(IFormFileCollection files)
        {
            try
            {
                if (files != null && files.Count > 0)
                {
                    string uploadsFolder = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Upload", "PEB");

                    if (!Directory.Exists(uploadsFolder))
                    {
                        Directory.CreateDirectory(uploadsFolder);
                    }

                    List<string> uploadedFiles = new List<string>();

                    foreach (var file in files)
                    {
                        if (file.Length > 0)
                        {
                            string originalFileName = System.IO.Path.GetFileNameWithoutExtension(file.FileName);
                            string fileExtension = System.IO.Path.GetExtension(file.FileName);
                            string fileName = originalFileName + fileExtension;
                            string filePath = System.IO.Path.Combine(uploadsFolder, fileName);

                            int counter = 1;
                            while (System.IO.File.Exists(filePath))
                            {
                                fileName = $"{originalFileName}({counter}){fileExtension}";
                                filePath = System.IO.Path.Combine(uploadsFolder, fileName);
                                counter++;
                            }

                            using (var fileStream = new FileStream(filePath, FileMode.Create))
                            {
                                file.CopyTo(fileStream);
                            }

                            uploadedFiles.Add(fileName);
                        }
                    }

                    if (uploadedFiles.Count > 0)
                    {
                        string combinedFilenames = string.Join(";", uploadedFiles);
                        return $"OK;{combinedFilenames}";
                    }
                    else
                    {
                        return "ERROR;No valid files provided";
                    }
                }
                return "ERROR;No files provided";
            }
            catch (Exception ex)
            {
                return $"ERROR;{ex.Message}";
            }
        }

        public string UPLOAD_PROFORMA_NON_ASSET_FIN_FILES(IFormFile file, string file_type)
        {
            try
            {
                string uploadPath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Upload", "Proforma_Finance_Non_Asset");
                if (!Directory.Exists(uploadPath))
                {
                    Directory.CreateDirectory(uploadPath);
                }

                if (file != null && file.Length > 0)
                {
                    string originalName = System.IO.Path.GetFileNameWithoutExtension(file.FileName);
                    string extension = System.IO.Path.GetExtension(file.FileName);
                    string fileName = file.FileName;

                    int counter = 1;
                    string finalFileName = fileName;
                    while (System.IO.File.Exists(System.IO.Path.Combine(uploadPath, finalFileName)))
                    {
                        finalFileName = $"{originalName}({counter}){extension}";
                        counter++;
                    }

                    string filePath = System.IO.Path.Combine(uploadPath, finalFileName);
                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        file.CopyTo(stream);
                    }

                    return $"OK;{finalFileName}";
                }
                return "ERROR;No file selected";
            }
            catch (Exception ex)
            {
                return $"ERROR;{ex.Message}";
            }
        }
        public string UPLOAD_PEB_NON_ASSET_FILE(IFormFileCollection files)
        {
            try
            {
                if (files != null && files.Count > 0)
                {
                    string uploadsFolder = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Upload", "PEB_Non_Asset");

                    if (!Directory.Exists(uploadsFolder))
                    {
                        Directory.CreateDirectory(uploadsFolder);
                    }

                    List<string> uploadedFiles = new List<string>();

                    foreach (var file in files)
                    {
                        if (file.Length > 0)
                        {
                            string originalFileName = System.IO.Path.GetFileNameWithoutExtension(file.FileName);
                            string fileExtension = System.IO.Path.GetExtension(file.FileName);
                            string fileName = originalFileName + fileExtension;
                            string filePath = System.IO.Path.Combine(uploadsFolder, fileName);

                            int counter = 1;
                            while (System.IO.File.Exists(filePath))
                            {
                                fileName = $"{originalFileName}({counter}){fileExtension}";
                                filePath = System.IO.Path.Combine(uploadsFolder, fileName);
                                counter++;
                            }

                            using (var fileStream = new FileStream(filePath, FileMode.Create))
                            {
                                file.CopyTo(fileStream);
                            }

                            uploadedFiles.Add(fileName);
                        }
                    }

                    if (uploadedFiles.Count > 0)
                    {
                        string combinedFilenames = string.Join(";", uploadedFiles);
                        return $"OK;{combinedFilenames}";
                    }
                    else
                    {
                        return "ERROR;No valid files provided";
                    }
                }
                return "ERROR;No files provided";
            }
            catch (Exception ex)
            {
                return $"ERROR;{ex.Message}";
            }
        }

        public string UPLOAD_PROFORMA_NON_ASSET_FILES(IFormFile file_attach, IFormFileCollection file_support)
        {
            try
            {
                string uploadPath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Upload/Proforma_Non_Asset");
                if (!Directory.Exists(uploadPath))
                {
                    Directory.CreateDirectory(uploadPath);
                }

                string file_attach_name = "";
                string file_support_names = "";

                if (file_attach != null && file_attach.Length > 0)
                {
                    string originalName = System.IO.Path.GetFileNameWithoutExtension(file_attach.FileName);
                    string extension = System.IO.Path.GetExtension(file_attach.FileName);
                    string fileName = file_attach.FileName;

                    int counter = 1;
                    string finalFileName = fileName;
                    while (System.IO.File.Exists(System.IO.Path.Combine(uploadPath, finalFileName)))
                    {
                        finalFileName = $"{originalName}({counter}){extension}";
                        counter++;
                    }

                    string filePath = System.IO.Path.Combine(uploadPath, finalFileName);
                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        file_attach.CopyTo(stream);
                    }
                    file_attach_name = finalFileName;
                }

                if (file_support != null && file_support.Count > 0)
                {
                    List<string> supportFileNames = new List<string>();
                    foreach (var supportFile in file_support.Take(2))
                    {
                        string originalName = System.IO.Path.GetFileNameWithoutExtension(supportFile.FileName);
                        string extension = System.IO.Path.GetExtension(supportFile.FileName);
                        string fileName = supportFile.FileName;

                        int counter = 1;
                        string finalFileName = fileName;
                        while (System.IO.File.Exists(System.IO.Path.Combine(uploadPath, finalFileName)))
                        {
                            finalFileName = $"{originalName}({counter}){extension}";
                            counter++;
                        }

                        string filePath = System.IO.Path.Combine(uploadPath, finalFileName);
                        using (var stream = new FileStream(filePath, FileMode.Create))
                        {
                            supportFile.CopyTo(stream);
                        }
                        supportFileNames.Add(finalFileName);
                    }
                    file_support_names = string.Join(";", supportFileNames);
                }

                return $"OK;{file_attach_name};{file_support_names}";
            }
            catch (Exception ex)
            {
                return $"ERROR;{ex.Message}";
            }
        }

        public string UPLOAD_INVOICE_FILE(IFormFile file, string fileType)
        {
            try
            {
                string uploadPath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Upload", "Gatepass_Invoice");

                if (!Directory.Exists(uploadPath))
                {
                    Directory.CreateDirectory(uploadPath);
                }

                string originalFileName = System.IO.Path.GetFileNameWithoutExtension(file.FileName);
                string extension = System.IO.Path.GetExtension(file.FileName);
                string finalFileName = file.FileName;
                string fullPath = System.IO.Path.Combine(uploadPath, finalFileName);

                int counter = 1;
                while (System.IO.File.Exists(fullPath))
                {
                    finalFileName = $"{originalFileName} ({counter}){extension}";
                    fullPath = System.IO.Path.Combine(uploadPath, finalFileName);
                    counter++;
                }

                using (var stream = new FileStream(fullPath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }

                return $"OK;{finalFileName}";
            }
            catch (Exception ex)
            {
                return $"ERROR;{ex.Message}";
            }
        }

        public IActionResult InvoicePDF(int id_invoice)
        {
            var db = new DatabaseAccessLayer();
            var invoiceData = db.GET_GATEPASS_INVOICE_EXPORT_DATA(id_invoice);

            if (invoiceData == null || invoiceData.Header == null)
            {
                return NotFound("Invoice not found");
            }

            var currencySymbols = new Dictionary<string, string>
            {
                { "EUR", "€" },
                { "USD", "$" },
                { "AUD", "A$" },
                { "CNY", "¥" },
                { "IDR", "Rp" },
                { "JPY", "¥" },
                { "SGD", "S$" },
                { "VND", "₫" },
                { "THB", "฿" }
            };

            string currency = invoiceData.Header.invoice_currency ?? "IDR";
            string currencySymbol = currencySymbols.ContainsKey(currency) ? currencySymbols[currency] : "";
            string bankName = currency switch
            {
                "USD" => "Bank America",
                "SGD" => "Bank Singapore",
                "IDR" => "Bank Indonesia",
                _ => throw new NotImplementedException()
            };
            string accountNumber = currency switch
            {
                "USD" => "665-4433-123 (USD)",
                "SGD" => "883-4443-144 (SGD)",
                "IDR" => "665-3332-131 (IDR)",
                _ => throw new NotImplementedException()
            };
            string swiftCode = currency == "SGD" ? "IDNSGSG" : "IDNIDJX";

            string commList = "";
            int no = 0;
            decimal totalAmount = 0;

            var groupedDetails = invoiceData.Details
                .GroupBy(d => d.id_invoice_detail)
                .ToList();

            foreach (var group in groupedDetails)
            {
                var detailsList = group.ToList();
                decimal groupAmount = detailsList.First().amount;
                totalAmount += groupAmount;
                int rowCount = detailsList.Count;

                for (int i = 0; i < detailsList.Count; i++)
                {
                    var row = detailsList[i];
                    no++;

                    commList += "<tr>";
                    commList += $"<td class='border_table' style='text-align:center; padding: 5px;'>{no}</td>";
                    commList += $"<td class='border_table' style='padding: 5px;'>{row.asset_no}-{row.asset_subnumber}</td>";
                    commList += $"<td class='border_table' style='padding: 5px;'>{row.asset_desc}</td>";
                    commList += $"<td class='border_table' style='padding: 5px;'>{row.asset_class}</td>";

                    if (i == 0)
                    {
                        commList += $"<td class='border_table' rowspan='{rowCount}'><div style='display: flex; justify-content: space-between; align-items: center; padding: 5px;'><span style='flex-shrink: 0;'>{currencySymbol}</span><span style='text-align: right; flex-grow: 1;'>{groupAmount.ToString("N2")}</span></div></td>";
                    }

                    commList += "</tr>";
                }
            }

            commList += $@"<tr style='background-color: #e8f5e9; font-weight: bold;'>
        <td colspan='4' class='border_table' style='text-align:right; padding: 8px;'>Total Amount ({currency})</td>
        <td class='border_table' style='padding: 5px; font-size: 11pt;'><div style='display: flex; justify-content: space-between; align-items: center;'><span style='flex-shrink: 0;'>{currencySymbol}</span><span style='text-align: right; flex-grow: 1;'>{totalAmount.ToString("N2")}</span></div></td>
        </tr>";

            string htmlToPdf = $@"<html>
          <header>
            <style>
              body {{
                margin-left: 20px;
                margin-right: 20px;
                margin-top: 10px;
                margin-bottom: 10px;
                font-family: Helvetica, Arial, sans-serif;
                font-size: 9pt;
              }}
              .border_table {{
                border: 0.5px solid #ababab;
              }}
              .bg-green {{
                background-color: #009E4C;
                color: white;
                font-weight: bold;
              }}
            </style>
          </header>
          <body>
            <table style='width:100%; text-align:left; line-height:1.5; margin-bottom: 20px;'>
                <tr>
                    <td style='font-size:11pt; width:50%; vertical-align:top;'>
                        <strong>Asset Management</strong><br>
                        Batam<br>
                        Jln. Semangka<br>
                        Batam Center<br>
                        Batam Island
                    </td>
                    <td style='padding-top:0px;padding-bottom:5px; text-align:right'>
                        <img src='wwwroot/images/logo/aml_black.png' style='width: 70%; height:auto; padding-left:30%' /><br><br>
                        Invoice No. {invoiceData.Header.invoice_no}<br>
                        Invoice Date : {invoiceData.Header.invoice_date:dd MMM yyyy}
                    </td>
                </tr>
            </table>
            
            <table style='width:100%; margin-bottom: 10px; line-height:1.5;'>
                <tr>
                    <td colspan=2 style='font-size:12pt;padding-top:20px'>
                        Debit to :<br>
                        <b>{invoiceData.Header.vendor_name}</b><br>
                        {invoiceData.Header.vendor_address}
                    </td>
                </tr>
            </table>
            
            <table style='width:100%; margin-bottom: 15px;'>
                <tr>
                    <td style='font-size:10pt; padding-top:10px; padding-bottom:10px; border-bottom: 2px solid #009E4C;'>
                        <strong>Attn: Finance/Accounts Department</strong>
                    </td>
                </tr>
            </table>
            
            <table style='width:100%; font-size:8pt; text-align:left; border-collapse: collapse;' cellpadding='0' cellspacing='0'>
                <tr class='bg-green'>
                    <th class='border_table' style='padding: 8px; text-align:center;'>NO.</th>
                    <th class='border_table' style='padding: 8px;'>ASSET NO</th>
                    <th class='border_table' style='padding: 8px;'>DESCRIPTION</th>
                    <th class='border_table' style='padding: 8px;'>ASSET CLASS</th>
                    <th class='border_table' style='padding: 8px; text-align:right;'>AMOUNT ({currency})</th>
                </tr>
                {commList}
            </table>
            
            <table style='margin-top:30px; width:100%;'>
                <tr>
                    <td style='width:65%; vertical-align:top; color:#575656; line-height:1.8; font-size:9pt;'>
                        <span style='color:black'>Please transfer the payment to :</span><br>
                        Beneficiary Name : Asset Management<br>
                        Beneficiary Bank Name : {bankName}<br>
                        Beneficiary Account Number : {accountNumber}<br>
                        Swift Code : {swiftCode}
                    </td>
                    <td style='width:35%; text-align:center; vertical-align:top; font-size:9pt;'>
                        <strong>Approved By</strong><br>
                        Accounting Manager<br><br><br><br><br>
                        <strong>{invoiceData.Header.approval_by_name}</strong><br>
                        {invoiceData.Header.approval_date:dd MMM yyyy}
                    </td>
                </tr>
            </table>
          </body>
        </html>";

            MemoryStream stream = new MemoryStream();
            HtmlConverter.ConvertToPdf(htmlToPdf, stream);

            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            PdfWriter writer = new PdfWriter(byteArrayOutputStream);
            PdfDocument pdfDocument = new PdfDocument(writer);

            MemoryStream stream2 = new MemoryStream(stream.ToArray());
            using (var inputPdfDoc = new PdfDocument(new PdfReader(stream2)))
            {
                inputPdfDoc.CopyPagesTo(1, inputPdfDoc.GetNumberOfPages(), pdfDocument);
            }

            pdfDocument.Close();
            Response.Headers["Content-Disposition"] = "inline; filename=" + invoiceData.Header.invoice_no + ".pdf";
            return File(byteArrayOutputStream.ToArray(), "application/pdf");
        }

        public byte[] GenerateInvoicePDFBytes(int id_invoice)
        {
            var db = new DatabaseAccessLayer();
            var invoiceData = db.GET_GATEPASS_INVOICE_EXPORT_DATA(id_invoice);

            if (invoiceData == null || invoiceData.Header == null)
            {
                throw new Exception("Invoice not found");
            }

            var currencySymbols = new Dictionary<string, string>
            {
                { "EUR", "€" },
                { "USD", "$" },
                { "AUD", "A$" },
                { "CNY", "¥" },
                { "IDR", "Rp" },
                { "JPY", "¥" },
                { "SGD", "S$" },
                { "VND", "₫" },
                { "THB", "฿" }
            };

            string currency = invoiceData.Header.invoice_currency ?? "IDR";
            string currencySymbol = currencySymbols.ContainsKey(currency) ? currencySymbols[currency] : "";
            string bankName = currency switch
            {
                "USD" => "Bank America",
                "SGD" => "Bank Singapore",
                "IDR" => "Bank Indonesia",
                _ => throw new NotImplementedException()
            };
            string accountNumber = currency switch
            {
                "USD" => "665-4433-123 (USD)",
                "SGD" => "883-4443-144 (SGD)",
                "IDR" => "665-3332-131 (IDR)",
                _ => throw new NotImplementedException()
            };
            string swiftCode = currency == "SGD" ? "IDNSGSG" : "IDNIDJX";

            string commList = "";
            int no = 0;
            decimal totalAmount = 0;

            var groupedDetails = invoiceData.Details
                .GroupBy(d => d.id_invoice_detail)
                .ToList();

            foreach (var group in groupedDetails)
            {
                var detailsList = group.ToList();
                decimal groupAmount = detailsList.First().amount;
                totalAmount += groupAmount;
                int rowCount = detailsList.Count;

                for (int i = 0; i < detailsList.Count; i++)
                {
                    var row = detailsList[i];
                    no++;

                    commList += "<tr>";
                    commList += $"<td class='border_table' style='text-align:center; padding: 5px;'>{no}</td>";
                    commList += $"<td class='border_table' style='padding: 5px;'>{row.asset_no}-{row.asset_subnumber}</td>";
                    commList += $"<td class='border_table' style='padding: 5px;'>{row.asset_desc}</td>";
                    commList += $"<td class='border_table' style='padding: 5px;'>{row.asset_class}</td>";

                    if (i == 0)
                    {
                        commList += $"<td class='border_table' rowspan='{rowCount}'><div style='display: flex; justify-content: space-between; align-items: center; padding: 5px;'><span style='flex-shrink: 0;'>{currencySymbol}</span><span style='text-align: right; flex-grow: 1;'>{groupAmount.ToString("N2")}</span></div></td>";
                    }

                    commList += "</tr>";
                }
            }

            commList += $@"<tr style='background-color: #e8f5e9; font-weight: bold;'>
        <td colspan='4' class='border_table' style='text-align:right; padding: 8px;'>Total Amount ({currency})</td>
        <td class='border_table' style='padding: 5px; font-size: 11pt;'><div style='display: flex; justify-content: space-between; align-items: center;'><span style='flex-shrink: 0;'>{currencySymbol}</span><span style='text-align: right; flex-grow: 1;'>{totalAmount.ToString("N2")}</span></div></td>
        </tr>";

            string htmlToPdf = $@"<html>
          <header>
            <style>
              body {{
                margin-left: 20px;
                margin-right: 20px;
                margin-top: 10px;
                margin-bottom: 10px;
                font-family: Helvetica, Arial, sans-serif;
                font-size: 9pt;
              }}
              .border_table {{
                border: 0.5px solid #ababab;
              }}
              .bg-green {{
                background-color: #009E4C;
                color: white;
                font-weight: bold;
              }}
            </style>
          </header>
          <body>
            <table style='width:100%; text-align:left; line-height:1.5; margin-bottom: 20px;'>
                <tr>
                    <td style='font-size:11pt; width:50%; vertical-align:top;'>
                        <strong>Asset Management</strong><br>
                        Batam<br>
                        Jln. Semangka<br>
                        Batam Center<br>
                        Batam Island
                    </td>
                    <td style='padding-top:0px;padding-bottom:5px; text-align:right'>
                        <img src='wwwroot/images/logo/aml_black.png' style='width: 70%; height:auto; padding-left:30%' /><br><br>
                        Invoice No. {invoiceData.Header.invoice_no}<br>
                        Invoice Date : {invoiceData.Header.invoice_date:dd MMM yyyy}
                    </td>
                </tr>
            </table>
            
            <table style='width:100%; margin-bottom: 10px; line-height:1.5;'>
                <tr>
                    <td colspan=2 style='font-size:12pt;padding-top:20px'>
                        Debit to :<br>
                        <b>{invoiceData.Header.vendor_name}</b><br>
                        {invoiceData.Header.vendor_address}
                    </td>
                </tr>
            </table>
            
            <table style='width:100%; margin-bottom: 15px;'>
                <tr>
                    <td style='font-size:10pt; padding-top:10px; padding-bottom:10px; border-bottom: 2px solid #009E4C;'>
                        <strong>Attn: Finance/Accounts Department</strong>
                    </td>
                </tr>
            </table>
            
            <table style='width:100%; font-size:8pt; text-align:left; border-collapse: collapse;' cellpadding='0' cellspacing='0'>
                <tr class='bg-green'>
                    <th class='border_table' style='padding: 8px; text-align:center;'>NO.</th>
                    <th class='border_table' style='padding: 8px;'>ASSET NO</th>
                    <th class='border_table' style='padding: 8px;'>DESCRIPTION</th>
                    <th class='border_table' style='padding: 8px;'>ASSET CLASS</th>
                    <th class='border_table' style='padding: 8px; text-align:right;'>AMOUNT ({currency})</th>
                </tr>
                {commList}
            </table>
            
            <table style='margin-top:30px; width:100%;'>
                <tr>
                    <td style='width:65%; vertical-align:top; color:#575656; line-height:1.8; font-size:9pt;'>
                        <span style='color:black'>Please transfer the payment to :</span><br>
                        Beneficiary Name : Asset Management<br>
                        Beneficiary Bank Name : {bankName}<br>
                        Beneficiary Account Number : {accountNumber}<br>
                        Swift Code : {swiftCode}
                    </td>
                    <td style='width:35%; text-align:center; vertical-align:top; font-size:9pt;'>
                        <strong>Approved By</strong><br>
                        Accounting Manager<br><br><br><br><br>
                        <strong>{invoiceData.Header.approval_by_name}</strong><br>
                        {invoiceData.Header.approval_date:dd MMM yyyy}
                    </td>
                </tr>
            </table>
          </body>
        </html>";

            MemoryStream stream = new MemoryStream();
            HtmlConverter.ConvertToPdf(htmlToPdf, stream);

            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            PdfWriter writer = new PdfWriter(byteArrayOutputStream);
            PdfDocument pdfDocument = new PdfDocument(writer);

            MemoryStream stream2 = new MemoryStream(stream.ToArray());
            using (var inputPdfDoc = new PdfDocument(new PdfReader(stream2)))
            {
                inputPdfDoc.CopyPagesTo(1, inputPdfDoc.GetNumberOfPages(), pdfDocument);
            }

            pdfDocument.Close();
            return byteArrayOutputStream.ToArray();
        }

        public string UPLOAD_INVOICE_PAYMENT_FILE(IFormFile file)
        {
            try
            {
                string uploadPath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Upload", "Gatepass_Invoice_Payment");

                if (!Directory.Exists(uploadPath))
                {
                    Directory.CreateDirectory(uploadPath);
                }

                string originalFileName = System.IO.Path.GetFileNameWithoutExtension(file.FileName);
                string extension = System.IO.Path.GetExtension(file.FileName);
                string finalFileName = originalFileName + extension;
                string fullPath = System.IO.Path.Combine(uploadPath, finalFileName);

                int counter = 1;
                while (System.IO.File.Exists(fullPath))
                {
                    finalFileName = $"{originalFileName} ({counter}){extension}";
                    fullPath = System.IO.Path.Combine(uploadPath, finalFileName);
                    counter++;
                }

                using (var stream = new FileStream(fullPath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }

                return $"OK;{finalFileName}";
            }
            catch (Exception ex)
            {
                return $"ERROR;{ex.Message}";
            }
        }

    }
}
