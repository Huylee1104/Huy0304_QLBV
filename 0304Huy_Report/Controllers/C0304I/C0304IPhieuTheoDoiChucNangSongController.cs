using _0304Huy_Report.Models;
using M0304.Models.ThongTinDoanhNghiep;
using M0304I.Models.PhieuTheoDoiChucNangSong;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Diagnostics;

namespace C0304IPhieuTheoDoiChucNangSong.Controllers
{
    [Route("phieu_theo_doi_chuc_nang_song")]
    public class C0304IPhieuTheoDoiChucNangSongController : Controller
    {
        //private string _maChucNang = "/phieu_theo_doi_chuc_nang_song";c
        //private IMemoryCachingServices _memoryCache;

        private readonly I0304CPhieuTheoDoiChucNangSongService _service;
        private readonly I0304BenhNhanService _benhNhanService;
        private readonly ILogger<C0304IPhieuTheoDoiChucNangSongController> _logger;

        public C0304IPhieuTheoDoiChucNangSongController(ILogger<C0304IPhieuTheoDoiChucNangSongController> logger, I0304CPhieuTheoDoiChucNangSongService service,
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


            return View("~/Views/V0304IPhieuTheoDoiChucNangSong/Index.cshtml");
        }

        [HttpPost("filter")]
        public async Task<IActionResult> Filter(long IdChiNhanh, long? idVaoVien = null, int page = 1, int pageSize = 20)
        {
            try
            {
                var result = await _service.GetPhieuTheoDoiChucNangSong(IdChiNhanh, idVaoVien, page, pageSize);

                if (!result.PhieuTheoDoiChucNangSong.Success)
                {
                    _logger.LogWarning("Service trả về lỗi: {Message}", result.PhieuTheoDoiChucNangSong.Message);
                    return Json(new { success = false, message = result.PhieuTheoDoiChucNangSong.Message });
                }

                return Json(new
                {
                    success = true,
                    message = result.PhieuTheoDoiChucNangSong.Message,
                    data = result.PhieuTheoDoiChucNangSong.Data,
                    totalRecords = result.PhieuTheoDoiChucNangSong.TotalRecords,
                    totalPages = result.PhieuTheoDoiChucNangSong.TotalPages,
                    currentPage = result.PhieuTheoDoiChucNangSong.CurrentPage,
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
            var pdfBytes = await _service.ExportGetPhieuTheoDoiChucNangSongPdfAsync(request, HttpContext.Session);

            _logger.LogWarning("pdfBytes: " + pdfBytes);

            string fileName = $"PhieuTheoDoiChucNangSong.pdf";
            return File(pdfBytes, "application/pdf", fileName);
        }
    }

}