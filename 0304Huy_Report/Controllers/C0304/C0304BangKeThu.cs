using _0304Huy_Report.Models;
using M0304.Models.ThongTinDOanhNghiep;
using S0304BangKeThu.Services;
using M0304.Models.BangKeThu;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Diagnostics;

namespace C0304BangKeThu.Controllers
{
    [Route("bang_ke_thu_ngoai_tru")]
    public class C0304BangKeThuController : Controller
    {
        //private string _maChucNang = "/bang_ke_thu_ngoai_tru";
        //private IMemoryCachingServices _memoryCache;

        private readonly I0304BangKeThuService _service;
        private readonly I0304HTTTService _htttService;
        private readonly I0304NhanVienService _nhanVienService;
        private readonly ILogger<C0304BangKeThuController> _logger;

        public C0304BangKeThuController(ILogger<C0304BangKeThuController> logger, I0304BangKeThuService service, I0304HTTTService htttService, 
            I0304NhanVienService nhanVienService /*, IMemoryCachingServices memoryCache*/)
        {
            _logger = logger;
            _service = service;
            _htttService = htttService;
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

            var dsHTTT = await _htttService.GetAllHTTT();  // giải Task ra List
            System.Diagnostics.Debug.WriteLine("DSHTTT: " + Newtonsoft.Json.JsonConvert.SerializeObject(dsHTTT));
            ViewBag.DSHTTT = Newtonsoft.Json.JsonConvert.SerializeObject(dsHTTT); // đưa sang JSON string

            var dsNhanVien = await _nhanVienService.GetAllNhanVien(); // cũng giải Task
            System.Diagnostics.Debug.WriteLine("DSNhanVien: " + Newtonsoft.Json.JsonConvert.SerializeObject(dsNhanVien));
            ViewBag.DSNhanVien = Newtonsoft.Json.JsonConvert.SerializeObject(dsNhanVien);


            return View("~/Views/V0304/Index.cshtml");
        }

        [HttpPost("filter")]
        public async Task<IActionResult> FilterByDay(string tuNgay, string denNgay, long IdChiNhanh, long? idHTTT = null,
            long? idNhanVien = null, int page = 1, int pageSize = 20)
        {
            try
            {
                var result = await _service.GetBangKeThu(tuNgay, denNgay, IdChiNhanh, idHTTT, idNhanVien, page, pageSize);

                if (!result.BangKeThu.Success)
                {
                    _logger.LogWarning("Service trả về lỗi: {Message}", result.BangKeThu.Message);
                    return Json(new { success = false, message = result.BangKeThu.Message });
                }

                return Json(new
                {
                    success = true,
                    message = result.BangKeThu.Message,
                    data = result.BangKeThu.Data,
                    totalRecords = result.BangKeThu.TotalRecords,
                    totalPages = result.BangKeThu.TotalPages,
                    currentPage = result.BangKeThu.CurrentPage,
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

            string fileName = $"BangKeThuNgoaiTru_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.pdf";
            return File(pdfBytes, "application/pdf", fileName);
        }

        [HttpPost("export/excel")]
        public async Task<IActionResult> ExportToExcel([FromBody] ExportRequest request)
        {
            var excelBytes = await _service.ExportBaoCaoGoiKhamExcelAsync(request, HttpContext.Session);

            string fileName = $"BangKeThuNgoaiTru_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.xlsx";
            return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
    }

}