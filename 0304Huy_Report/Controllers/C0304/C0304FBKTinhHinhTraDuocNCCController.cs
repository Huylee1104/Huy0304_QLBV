using _0304Huy_Report.Models;
using M0304.Models.ThongTinDoanhNghiep;
using S0304FBKTinhHinhTraDuocNCC.Services;
using M0304F.Models.BKTinhHinhTraDuocNCC;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Diagnostics;

namespace C0304FBKTinhHinhTraDuocNCC.Controllers
{
    [Route("bang_ke_tinh_hinh_tra_duoc_NCC")]
    public class C0304FBKTinhHinhTraDuocNCCController : Controller
    {
        //private string _maChucNang = "/bang_ke_tinh_hinh_tra_duoc_NCC";
        //private IMemoryCachingServices _memoryCache;

        private readonly I0304FBKTinhHinhTraDuocNCCService _service;
        private readonly I0304KhoHangService _khoHangService;
        private readonly ILogger<C0304FBKTinhHinhTraDuocNCCController> _logger;

        public C0304FBKTinhHinhTraDuocNCCController(ILogger<C0304FBKTinhHinhTraDuocNCCController> logger, I0304FBKTinhHinhTraDuocNCCService service,
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


            return View("~/Views/V0304/V0304FBKTinhHinhTraDuocNCC/Index.cshtml");
        }

        [HttpPost("filter")]
        public async Task<IActionResult> FilterByDay(string tuNgay, string denNgay, long IdChiNhanh,
            long? idKhoHang = null, int page = 1, int pageSize = 20)
        {
            try
            {
                var result = await _service.GetBKTinhHinhTraDuocNCC(tuNgay, denNgay, IdChiNhanh, idKhoHang, page, pageSize);

                if (!result.BKTinhHinhTraDuocNCC.Success)
                {
                    _logger.LogWarning("Service trả về lỗi: {Message}", result.BKTinhHinhTraDuocNCC.Message);
                    return Json(new { success = false, message = result.BKTinhHinhTraDuocNCC.Message });
                }

                return Json(new
                {
                    success = true,
                    message = result.BKTinhHinhTraDuocNCC.Message,
                    data = result.BKTinhHinhTraDuocNCC.Data,
                    totalRecords = result.BKTinhHinhTraDuocNCC.TotalRecords,
                    totalPages = result.BKTinhHinhTraDuocNCC.TotalPages,
                    currentPage = result.BKTinhHinhTraDuocNCC.CurrentPage,
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
            var pdfBytes = await _service.ExportBKTinhHinhTraDuocNCCPdfAsync(request, HttpContext.Session);

            string fileName = $"BangKeTinhHinhTraDuocNCC_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.pdf";
            return File(pdfBytes, "application/pdf", fileName);
        }

        [HttpPost("export/excel")]
        public async Task<IActionResult> ExportToExcel([FromBody] ExportRequest request)
        {
            var excelBytes = await _service.ExportBKTinhHinhTraDuocNCCExcelAsync(request, HttpContext.Session);

            string fileName = $"BangKeTinhHinhTraDuocNCC_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.xlsx";
            return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
    }

}