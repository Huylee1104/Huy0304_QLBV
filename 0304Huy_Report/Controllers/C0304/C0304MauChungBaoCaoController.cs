using Microsoft.AspNetCore.Mvc;

namespace MauChungBaoCao.Controllers
{
    [Route("mau_bao_cao_chung")]
    public class C0304MauChungBaoCaoController : Controller
    {
        //private string _maChucNang = "/mau_bao_cao_chung";
        //private IMemoryCachingServices _memoryCache;

        private readonly ILogger<C0304MauChungBaoCaoController> _logger;

        public C0304MauChungBaoCaoController(ILogger<C0304MauChungBaoCaoController> logger /*, IMemoryCachingServices memoryCache*/)
        {
            _logger = logger;
            //_memoryCache = memoryCache;
        }
        public async Task<IActionResult> Index()
        {
            //var quyenVaiTro = await _memoryCache.getQuyenVaiTro(_maChucNang);
            //if (quyenVaiTro == null)
            //{
            //    return RedirectToAction("NotFound", "Home");
            //}
            //ViewBag.quyenVaiTro = quyenVaiTro;
            //ViewData["Title"] = CommonServices.toEmptyData(quyenVaiTro);
            ViewBag.quyenVaiTro = new
            {
                Them = true,
                Sua = true,
                Xoa = true,
                Xuat = true,
                CaNhan = true,
                Xem = true,
            };

            return View("~/Views/V0304/V0304MauChungBaoCao/Index.cshtml");
        }

        //[HttpPost("filter")]
        //public async Task<IActionResult> FilterByDay(string tuNgay, string denNgay, long IdChiNhanh, int page = 1, int pageSize = 20)
        //{
        //    try
        //    {
        //        var result = await _service.GetBCThuDichVu(tuNgay, denNgay, IdChiNhanh, page, pageSize);

        //        if (!result.BaoCaoThuDichVu.Success)
        //        {
        //            _logger.LogWarning("Service trả về lỗi: {Message}", result.BaoCaoThuDichVu.Message);
        //            return Json(new { success = false, message = result.BaoCaoThuDichVu.Message });
        //        }

        //        return Json(new
        //        {
        //            success = true,
        //            message = result.BaoCaoThuDichVu.Message,
        //            data = result.BaoCaoThuDichVu.Data,
        //            totalRecords = result.BaoCaoThuDichVu.TotalRecords,
        //            totalPages = result.BaoCaoThuDichVu.TotalPages,
        //            currentPage = result.BaoCaoThuDichVu.CurrentPage,
        //            doanhNghiep = result.DoanhNghiep
        //        });
        //    }
        //    catch (Exception ex)
        //    {
        //        _logger.LogError(ex, "Lỗi xảy ra trong FilterByDay");
        //        return StatusCode(500, "Lỗi server: " + ex.Message);
        //    }

        //}

        //[HttpPost("export/pdf")]
        //public async Task<IActionResult> ExportToPDF([FromBody] ExportRequest request)
        //{
        //    var pdfBytes = await _service.ExportBaoCaoThuDichVuPdfAsync(request, HttpContext.Session);

        //    string fileName = $"BaoCaoThuDichVu_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.pdf";
        //    return File(pdfBytes, "application/pdf", fileName);
        //}

        //[HttpPost("export/excel")]
        //public async Task<IActionResult> ExportToExcel([FromBody] ExportRequest request)
        //{
        //    var excelBytes = await _service.ExportBaoCaoThuDichVuExcelAsync(request, HttpContext.Session);

        //    string fileName = $"BaoCaoThuDichVu_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.xlsx";
        //    return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        //}
    }
}
