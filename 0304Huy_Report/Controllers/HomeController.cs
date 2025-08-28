using System.Diagnostics;
using _0304Huy_Report.Models;
using Microsoft.AspNetCore.Mvc;

namespace _0304Huy_Report.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        public IActionResult BangKeThu()
        {
            return RedirectToAction("Index", "C0304BangKeThu");
        }

        public IActionResult BCHoaDonDienTuDV()
        {
            return RedirectToAction("Index", "C0304BCHoaDonDienTuDV");
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
