using Microsoft.AspNetCore.Mvc;

namespace PDFService.Controllers
{
    [Route("api/index")]
    public class IndexController : Controller
    {
        [HttpGet]
        public string Get(int id)
        {
            return "PDF Service is running!";
        }
    }
}
