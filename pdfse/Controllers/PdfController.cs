using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Threading.Tasks;

namespace PDFService.Controllers
{
    [Route("api/[controller]")]
    public class PdfController : Controller
    {
        [HttpPost("convert")]
        public async Task<IActionResult> Post(IFormFile file, bool sign = false, string flag=null)
        {
            var fileName = file.FileName;
            
            var filePath = Path.GetTempFileName();

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            //return PhysicalFile("", "pdf");

            return Ok(new { count = file.Length, filePath });
        }
    }
}
