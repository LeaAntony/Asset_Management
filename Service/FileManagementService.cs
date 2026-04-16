using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.Net.Http.Headers;

namespace Asset_Management.Service
{
    public class FileManagementService
    {
        public string UploadFile(IFormFile file)
        {
            var fileName = ContentDispositionHeaderValue.Parse(file.ContentDisposition).FileName.Trim();

            var mainPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Upload");

            if (!Directory.Exists(mainPath))
            {
                Directory.CreateDirectory(mainPath);
            }

            var filePath = Path.Combine(mainPath, file.FileName);

            try
            {
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }

                return filePath;
            }
            catch (Exception e)
            {
                return "Upload Fail";
            }

        }

        public string UploadFileRename(IFormFile file, string filename, string subfolder)
        {
            var originalFileName = Path.GetFileName(ContentDispositionHeaderValue.Parse(file.ContentDisposition).FileName.Trim());
            var fileExtension = Path.GetExtension(originalFileName);

            var newFileName = filename+fileExtension.ToString();

            var mainPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Upload", subfolder);

            if (!Directory.Exists(mainPath))
            {
                Directory.CreateDirectory(mainPath);
            }

            var filePath = Path.Combine(mainPath, newFileName);

            try
            {
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }

                return "OK;"+newFileName;
            }
            catch (Exception e)
            {
                return "ERROR;Upload Fail: " + e.Message;
            }
        }

    }
}
