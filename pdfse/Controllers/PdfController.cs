﻿using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using PDFService.Business;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace PDFService.Controllers
{
    [Route("api/[controller]")]
    public class PdfController : Controller
    {
        private readonly FileManager manager;

        private readonly Encoding utf8 = Encoding.UTF8;

        public PdfController(FileManager manager)
        {
            this.manager = manager;
        }

        [HttpPost("convert")]
        public async Task<IActionResult> Convert(IFormFile file, bool sign = false, string flag = null)
        {
            try
            {
                var fileName = HttpUtility.UrlDecode(file.FileName, utf8);
                var extension = Path.GetExtension(fileName).ToLower();

                if (extension.EndsWith(FileManager.PDF))
                {
                    return File(file.OpenReadStream(), file.ContentType, HttpUtility.UrlEncode(fileName, utf8));
                }

                var filePath = manager.GetInputDocPath(fileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                filePath = manager.Convert(filePath);

                if (sign)
                {
                    filePath = manager.Sign(filePath, flag);
                }

                fileName = HttpUtility.UrlEncode(Path.ChangeExtension(fileName, FileManager.PDF), utf8);

                return PhysicalFile(filePath, MimeMapping.GetMimeMapping(filePath), fileName);
            }
            catch
            {
                return StatusCode(StatusCodes.Status500InternalServerError);
            }
        }

        [HttpPost("sign")]
        public async Task<IActionResult> Sign(IFormFile file, string flag = null)
        {
            try
            {
                var fileName = HttpUtility.UrlDecode(file.FileName, utf8);

                var filePath = manager.GetInputPdfPath(fileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                filePath = manager.Sign(filePath, flag);

                return PhysicalFile(filePath, MimeMapping.GetMimeMapping(filePath));
            }
            catch
            {
                return StatusCode(StatusCodes.Status500InternalServerError);
            }
        }
    }
}
