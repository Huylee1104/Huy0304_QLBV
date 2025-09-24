using _0304Huy_Report.Models;
using M0304.Models.ThongTinDoanhNghiep;
using M0304L.Models.PhieuTheoDoiTruyenDich;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Diagnostics;

namespace C0304LPhieuTheoDoiTruyenDich.Controllers
{
    [Route("phieu_theo_doi_truyen_dich")]
    public class C0304LPhieuTheoDoiTruyenDichController : Controller
    {
        //private string _maChucNang = "/phieu_theo_doi_truyen_dich";
        //private IMemoryCachingServices _memoryCache;

        private readonly I0304LPhieuTheoDoiTruyenDichService _service;
        private readonly I0304BenhNhanService _benhNhanService;
        private readonly ILogger<C0304LPhieuTheoDoiTruyenDichController> _logger;

        public C0304LPhieuTheoDoiTruyenDichController(ILogger<C0304LPhieuTheoDoiTruyenDichController> logger, I0304LPhieuTheoDoiTruyenDichService service,
            I0304BenhNhanService benhNhanService /*, IMemoryCachingServices memoryCache*/)
        {
            _logger = logger;
            _service = service;
            _benhNhanService = benhNhanService;
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

            var dsBenhNhan = await _benhNhanService.GetAllBenhNhan();
            System.Diagnostics.Debug.WriteLine("DSBenhNhan: " + Newtonsoft.Json.JsonConvert.SerializeObject(dsBenhNhan));
            ViewBag.DSBenhNhan = Newtonsoft.Json.JsonConvert.SerializeObject(dsBenhNhan);


            return View("~/Views/V0304LPhieuTheoDoiTruyenDich/Index.cshtml");
        }

        [HttpPost("filter")]
        public async Task<IActionResult> Filter(long IdChiNhanh, long? idBenhNhan = null, int page = 1, int pageSize = 20)
        {
            try
            {
                var result = await _service.GetPhieuTheoDoiTruyenDich(IdChiNhanh, idBenhNhan, page, pageSize);

                if (!result.PhieuTheoDoiTruyenDich.Success)
                {
                    _logger.LogWarning("Service trả về lỗi: {Message}", result.PhieuTheoDoiTruyenDich.Message);
                    return Json(new { success = false, message = result.PhieuTheoDoiTruyenDich.Message });
                }

                return Json(new
                {
                    success = true,
                    message = result.PhieuTheoDoiTruyenDich.Message,
                    data = result.PhieuTheoDoiTruyenDich.Data,
                    totalRecords = result.PhieuTheoDoiTruyenDich.TotalRecords,
                    totalPages = result.PhieuTheoDoiTruyenDich.TotalPages,
                    currentPage = result.PhieuTheoDoiTruyenDich.CurrentPage,
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
            var pdfBytes = await _service.ExportGetPhieuTheoDoiTruyenDichPdfAsync(request, HttpContext.Session);

            _logger.LogWarning("pdfBytes: " + pdfBytes);

            string fileName = $"PhieuTheoDoiTruyenDich.pdf";
            return File(pdfBytes, "application/pdf", fileName);
        }
    }

}