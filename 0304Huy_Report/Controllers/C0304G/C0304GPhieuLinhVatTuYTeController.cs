using _0304Huy_Report.Models;
using M0304.Models.ThongTinDoanhNghiep;
using S0304GPhieuLinhVatTuYTe.Services;
using M0304G.Models.PhieuLinhVatTuYTe;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Diagnostics;

namespace C0304GPhieuLinhVatTuYTe.Controllers
{
    [Route("phieu_linh_vat_tu_y_te")]
    public class C0304GPhieuLinhVatTuYTeController : Controller
    {
        //private string _maChucNang = "/phieu_linh_vat_tu_y_te";
        //private IMemoryCachingServices _memoryCache;

        private readonly I0304GPhieuLinhVatTuYTeService _service;
        private readonly I0304KhoHangService _khoHangService;
        private readonly ILogger<C0304GPhieuLinhVatTuYTeController> _logger;

        public C0304GPhieuLinhVatTuYTeController(ILogger<C0304GPhieuLinhVatTuYTeController> logger, I0304GPhieuLinhVatTuYTeService service,
            I0304KhoHangService khoHangService /*, IMemoryCachingServices memoryCache*/)
        {
            _logger = logger;
            _service = service;
            _khoHangService = khoHangService;
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

            var dsKhoHang = await _khoHangService.GetKhoHang();
            System.Diagnostics.Debug.WriteLine("DSKhoHang: " + Newtonsoft.Json.JsonConvert.SerializeObject(dsKhoHang));
            ViewBag.DSKhoHang = Newtonsoft.Json.JsonConvert.SerializeObject(dsKhoHang);


            return View("~/Views/V0304GPhieuLinhVatTuYTe/Index.cshtml");
        }

        [HttpPost("filter")]
        public async Task<IActionResult> FilterByDay(string tuNgay, string denNgay, long IdChiNhanh,
            long? idKhoHang = null, int page = 1, int pageSize = 20)
        {
            try
            {
                var result = await _service.GetPhieuLinhVatTuYTe(tuNgay, denNgay, IdChiNhanh, idKhoHang, page, pageSize);

                if (!result.PhieuLinhVatTuYTe.Success)
                {
                    _logger.LogWarning("Service trả về lỗi: {Message}", result.PhieuLinhVatTuYTe.Message);
                    return Json(new { success = false, message = result.PhieuLinhVatTuYTe.Message });
                }

                return Json(new
                {
                    success = true,
                    message = result.PhieuLinhVatTuYTe.Message,
                    data = result.PhieuLinhVatTuYTe.Data,
                    totalRecords = result.PhieuLinhVatTuYTe.TotalRecords,
                    totalPages = result.PhieuLinhVatTuYTe.TotalPages,
                    currentPage = result.PhieuLinhVatTuYTe.CurrentPage,
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
            var pdfBytes = await _service.ExportPhieuLinhVatTuYTePdfAsync(request, HttpContext.Session);

            string fileName = $"PhieuLinhVatTuYTe_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.pdf";
            return File(pdfBytes, "application/pdf", fileName);
        }

        [HttpPost("export/excel")]
        public async Task<IActionResult> ExportToExcel([FromBody] ExportRequest request)
        {
            var excelBytes = await _service.ExportPhieuLinhVatTuYTeExcelAsync(request, HttpContext.Session);

            string fileName = $"PhieuLinhVatTuYTe_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.xlsx";
            return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
    }

}