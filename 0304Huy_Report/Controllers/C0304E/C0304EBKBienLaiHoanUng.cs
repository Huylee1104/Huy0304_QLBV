using _0304Huy_Report.Models;
using M0304.Models.ThongTinDoanhNghiep;
using S0304EBKBienLaiHoanUng.Services;
using M0304E.Models.BKBienLaiHoanUng;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Diagnostics;

namespace C0304EBKBienLaiHoanUng.Controllers
{
    [Route("bang_ke_bien_lai_hoan_ung")]
    public class C0304EBKBienLaiHoanUngController : Controller
    {
        //private string _maChucNang = "/bang_ke_bien_lai_hoan_ung";
        //private IMemoryCachingServices _memoryCache;

        private readonly I0304EBKBienLaiHoanUngService _service;
        private readonly I0304NhanVienService _nhanVienService;
        private readonly ILogger<C0304EBKBienLaiHoanUngController> _logger;

        public C0304EBKBienLaiHoanUngController(ILogger<C0304EBKBienLaiHoanUngController> logger, I0304EBKBienLaiHoanUngService service, 
            I0304NhanVienService nhanVienService /*, IMemoryCachingServices memoryCache*/)
        {
            _logger = logger;
            _service = service;
            _nhanVienService = nhanVienService;
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

            var dsNhanVien = await _nhanVienService.GetAllNhanVien(); // cũng giải Task
            System.Diagnostics.Debug.WriteLine("DSNhanVien: " + Newtonsoft.Json.JsonConvert.SerializeObject(dsNhanVien));
            ViewBag.DSNhanVien = Newtonsoft.Json.JsonConvert.SerializeObject(dsNhanVien);


            return View("~/Views/V0304E/Index.cshtml");
        }

        [HttpPost("filter")]
        public async Task<IActionResult> FilterByDay(string tuNgay, string denNgay, long IdChiNhanh,
            long? idNhanVien = null, int page = 1, int pageSize = 20)
        {
            try
            {
                var result = await _service.GetBKBienLaiHoanUng(tuNgay, denNgay, IdChiNhanh, idNhanVien, page, pageSize);

                if (!result.BangKeBienLaiHoanUng.Success)
                {
                    _logger.LogWarning("Service trả về lỗi: {Message}", result.BangKeBienLaiHoanUng.Message);
                    return Json(new { success = false, message = result.BangKeBienLaiHoanUng.Message });
                }

                return Json(new
                {
                    success = true,
                    message = result.BangKeBienLaiHoanUng.Message,
                    data = result.BangKeBienLaiHoanUng.Data,
                    totalRecords = result.BangKeBienLaiHoanUng.TotalRecords,
                    totalPages = result.BangKeBienLaiHoanUng.TotalPages,
                    currentPage = result.BangKeBienLaiHoanUng.CurrentPage,
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
            var pdfBytes = await _service.ExportBKBienLaiHoanUngPdfAsync(request, HttpContext.Session);

            string fileName = $"BangKeBienLaiHoanUngNgoaiTru_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.pdf";
            return File(pdfBytes, "application/pdf", fileName);
        }

        [HttpPost("export/excel")]
        public async Task<IActionResult> ExportToExcel([FromBody] ExportRequest request)
        {
            var excelBytes = await _service.ExportBKBienLaiHoanUngExcelAsync(request, HttpContext.Session);

            string fileName = $"BangKeBienLaiHoanUngNgoaiTru_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.xlsx";
            return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
    }

}