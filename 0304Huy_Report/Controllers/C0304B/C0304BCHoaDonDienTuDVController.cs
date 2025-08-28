using C0304BCHoaDonDienTuDV.Controllers;
using M0304B.Models.BCHoaDonDienTuDV;
using Microsoft.AspNetCore.Mvc;

namespace C0304BCHoaDonDienTuDV.Controllers
{
    [Route("bao_cao_hoa_don_dien_tu_dich_vu")]
    public class C0304BCHoaDonDienTuDVController : Controller
    {
        //private string _maChucNang = "/bang_ke_thu_ngoai_tru";
        //private IMemoryCachingServices _memoryCache;

        private readonly I0304BBCHoaDonDienTuDVService _service;
        private readonly ILogger<C0304BCHoaDonDienTuDVController> _logger;

        public C0304BCHoaDonDienTuDVController(ILogger<C0304BCHoaDonDienTuDVController> logger, I0304BBCHoaDonDienTuDVService service /*, IMemoryCachingServices memoryCache*/)
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
            return View("~/Views/V0304B/Index.cshtml");
        }

        [HttpPost("filter")]
        public async Task<IActionResult> FilterByDay(string tuNgay, string denNgay, long IdChiNhanh, int page = 1, int pageSize = 20)
        {
            try
            {
                var result = await _service.GetBCHoaDonDIenTuDV(tuNgay, denNgay, IdChiNhanh, page, pageSize);

                if (!result.BCHoaDonDienTuDV.Success)
                {
                    _logger.LogWarning("Service trả về lỗi: {Message}", result.BCHoaDonDienTuDV.Message);
                    return Json(new { success = false, message = result.BCHoaDonDienTuDV.Message });
                }

                return Json(new
                {
                    success = true,
                    message = result.BCHoaDonDienTuDV.Message,
                    data = result.BCHoaDonDienTuDV.Data,
                    totalRecords = result.BCHoaDonDienTuDV.TotalRecords,
                    totalPages = result.BCHoaDonDienTuDV.TotalPages,
                    currentPage = result.BCHoaDonDienTuDV.CurrentPage,
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
            var pdfBytes = await _service.ExportBaoCaoGoiKhamPdfAsync(request, HttpContext.Session);

            string fileName = $"BaoCaoHoaDonDienTuDichVu_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.pdf";
            return File(pdfBytes, "application/pdf", fileName);
        }

        [HttpPost("export/excel")]
        public async Task<IActionResult> ExportToExcel([FromBody] ExportRequest request)
        {
            var excelBytes = await _service.ExportBaoCaoGoiKhamExcelAsync(request, HttpContext.Session);

            string fileName = $"BaoCaoHoaDonDienTuDichVu_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.xlsx";
            return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
    }
}
