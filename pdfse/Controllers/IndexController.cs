using Microsoft.AspNetCore.Mvc;

namespace PDFService.Controllers
{
    [Route("api/index")]
    public class IndexController : Controller
    {
        [HttpGet]
        public string Get()
        {
            return "PDF Service is running!";
        }
    }
}
