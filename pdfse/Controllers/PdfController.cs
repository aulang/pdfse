using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using PDFService.Business;
using System;
using System.IO;
using System.Web;
using System.Threading.Tasks;
using System.Text;

namespace PDFService.Controllers
{
    [Route("api/[controller]")]
    public class PdfController : Controller
    {
        private FileManager Manager;

        private Encoding utf8 = Encoding.UTF8;

        public PdfController(FileManager manager)
        {
            this.Manager = manager;
        }

        [HttpPost("convert")]
        public async Task<IActionResult> Convert(IFormFile file, bool sign = false, string flag = null)
        {
            try
            {
                var fileName = HttpUtility.UrlDecode(file.FileName, utf8);

                var filePath = Manager.GetInputDocPath(fileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                filePath = Manager.Convert(filePath);

                if (sign)
                {
                    filePath = Manager.Sign(filePath, flag);
                }

                return PhysicalFile(filePath, MimeMapping.GetMimeMapping(filePath));
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

                var filePath = Manager.GetInputPdfPath(fileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                filePath = Manager.Sign(filePath, flag);

                return PhysicalFile(filePath, MimeMapping.GetMimeMapping(filePath));
            }
            catch
            {
                return StatusCode(StatusCodes.Status500InternalServerError);
            }
        }
    }
}
