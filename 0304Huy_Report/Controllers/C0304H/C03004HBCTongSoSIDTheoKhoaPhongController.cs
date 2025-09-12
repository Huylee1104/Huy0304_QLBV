using C0304BCHoaDonDienTuDV.Controllers;
using M0304H.Models.BCTongSoSIDTheoKhoaPhong;
using Microsoft.AspNetCore.Mvc;

namespace C0304HBCTongSoSIDTheoKhoaPhong.Controllers
{
    [Route("bao_cao_tong_so_SID_theo_khoa_phong")]
    public class C0304HBCTongSoSIDTheoKhoaPhongController : Controller
    {
        //private string _maChucNang = "/bao_cao_tong_so_SID_theo_khoa_phong";
        //private IMemoryCachingServices _memoryCache;

        private readonly I0304HBCTongSoSIDTheoKhoaPhongService _service;
        private readonly ILogger<C0304HBCTongSoSIDTheoKhoaPhongController> _logger;

        public C0304HBCTongSoSIDTheoKhoaPhongController(ILogger<C0304HBCTongSoSIDTheoKhoaPhongController> logger, I0304HBCTongSoSIDTheoKhoaPhongService service /*, IMemoryCachingServices memoryCache*/)
        {
            _logger = logger;
            _service = service;
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
            return View("~/Views/V0304HBCTongSoSIDTheoKhoaPhong/Index.cshtml");
        }

        [HttpPost("filter")]
        public async Task<IActionResult> FilterByDay(string tuNgay, string denNgay, long IdChiNhanh, int page = 1, int pageSize = 20)
        {
            try
            {
                var result = await _service.GetBCTongSoSIDTheoKhoaPhong(tuNgay, denNgay, IdChiNhanh, page, pageSize);

                if (!result.BCTongSoSIDTheoKhoaPhong.Success)
                {
                    _logger.LogWarning("Service trả về lỗi: {Message}", result.BCTongSoSIDTheoKhoaPhong.Message);
                    return Json(new { success = false, message = result.BCTongSoSIDTheoKhoaPhong.Message });
                }

                return Json(new
                {
                    success = true,
                    message = result.BCTongSoSIDTheoKhoaPhong.Message,
                    data = result.BCTongSoSIDTheoKhoaPhong.Data,
                    totalRecords = result.BCTongSoSIDTheoKhoaPhong.TotalRecords,
                    totalPages = result.BCTongSoSIDTheoKhoaPhong.TotalPages,
                    currentPage = result.BCTongSoSIDTheoKhoaPhong.CurrentPage,
                    doanhNghiep = result.DoanhNghiep
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi xảy ra trong FilterByDay");
                return StatusCode(500, "Lỗi server: " + ex.Message);
            }

        }

        [HttpPost("export/pdf")]
        public async Task<IActionResult> ExportToPDF([FromBody] ExportRequest request)
        {
            var pdfBytes = await _service.ExportBCTongSoSIDTheoKhoaPhongPdfAsync(request, HttpContext.Session);

            string fileName = $"BaoCaoTongSoSIDTheoKhoaPhong_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.pdf";
            return File(pdfBytes, "application/pdf", fileName);
        }

        [HttpPost("export/excel")]
        public async Task<IActionResult> ExportToExcel([FromBody] ExportRequest request)
        {
            var excelBytes = await _service.ExportBCTongSoSIDTheoKhoaPhongExcelAsync(request, HttpContext.Session);

            string fileName = $"BaoCaoTongSoSIDTheoKhoaPhong_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.xlsx";
            return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
    }
}
