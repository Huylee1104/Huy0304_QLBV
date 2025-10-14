using _0304Huy_Report.Models;
using C0304.Db.Models;
using M0304.Models.ThongTinDoanhNghiep;
using M0304.Models.BangKeBanLeHangHoaDichVu;
using P0304.PDFDocument.BangKeBanLeHangHoaDichVu;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using QuestPDF.Fluent;
using System.Diagnostics;

namespace C0304BangKeBanLeHangHoaDichVu.Controllers
{
    [Route("bang_ke_ban_le_hang_hoa_dich_vu")]
    public class C0304BangKeBanLeHangHoaDichVuController : Controller
    {
        //private string _maChucNang = "/bang_ke_ban_le_hang_hoa_dich_vu";
        //private IMemoryCachingServices _memoryCache;
        private readonly ILogger<C0304BangKeBanLeHangHoaDichVuController> _logger;
        private readonly M0304Context _context;
        private readonly I0304ThongTinDoanhNghiep _thongTinDoanhNghiepService;
        private readonly IWebHostEnvironment _env;

        public C0304BangKeBanLeHangHoaDichVuController(ILogger<C0304BangKeBanLeHangHoaDichVuController> logger,
            M0304Context context, I0304ThongTinDoanhNghiep thongTinDoanhNghiepService,
            IWebHostEnvironment env /*, IMemoryCachingServices memoryCache*/)
        {
            _logger = logger;
            _context = context;
            _thongTinDoanhNghiepService = thongTinDoanhNghiepService;
            _env = env;
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

            return View("~/Views/V0304/V0304BangKeBanLeHangHoaDichVu/Index.cshtml");
        }

        [HttpPost("filter")]
        public async Task<IActionResult> FilterByDay(string tuNgay, string denNgay, long IdChiNhanh, long idKhoHang, 
            long idNhanVien, int page = 1, int pageSize = 20)
        {
            try
            {
                var result = await GetBangKeBanLeHangHoaDichVu(tuNgay, denNgay, IdChiNhanh, idKhoHang, idNhanVien, page, pageSize);

                if (!result.BangKeBanLeHangHoaDichVu.Success)
                {
                    _logger.LogWarning("Service trả về lỗi: {Message}", result.BangKeBanLeHangHoaDichVu.Message);
                    return Json(new { success = false, message = result.BangKeBanLeHangHoaDichVu.Message });
                }

                return Json(new
                {
                    success = true,
                    message = result.BangKeBanLeHangHoaDichVu.Message,
                    data = result.BangKeBanLeHangHoaDichVu.Data,
                    totalRecords = result.BangKeBanLeHangHoaDichVu.TotalRecords,
                    totalPages = result.BangKeBanLeHangHoaDichVu.TotalPages,
                    currentPage = result.BangKeBanLeHangHoaDichVu.CurrentPage,
                    doanhNghiep = result.DoanhNghiep,
                    AllSoTien = result.TongCong,
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
            var pdfBytes = await ExportBangKeBanLeHangHoaDichVuPdfAsync(request, HttpContext.Session);

            string fileName = $"BangKeBanLeHangHoaDichVu_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.pdf";
            return File(pdfBytes, "application/pdf", fileName);
        }

        [HttpPost("export/excel")]
        public async Task<IActionResult> ExportToExcel([FromBody] ExportRequest request)
        {
            var excelBytes = await ExportBangKeBanLeHangHoaDichVuExcelAsync(request, HttpContext.Session);

            string fileName = $"BangKeBanLeHangHoaDichVu_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.xlsx";
            return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }


        private async Task<M0304BangKeBanLeHangHoaDichVuResponse> GetBangKeBanLeHangHoaDichVu(string ngayBatDau, string ngayKetThuc, long idCN,
            long idKhoHang, long idNhanVien, int page = 1, int pageSize = 20)
        {
            var doanhNghiep = await _thongTinDoanhNghiepService.GetThongTinDoanhNghiep(idCN);

            var session = HttpContext?.Session;

            if (doanhNghiep != null)
            {
                // Lưu thông tin doanh nghiệp vào session
                session?.SetString("DoanhNghiepInfo", JsonConvert.SerializeObject(doanhNghiep));
                _logger.LogInformation("Doanh Nghiep Info: {@DoanhNghiep}", doanhNghiep);
            }
            else
            {
                _logger.LogWarning("No doanh nghiep found for ChiNhanh ID: {IdChiNhanh}", idCN);
                return new M0304BangKeBanLeHangHoaDichVuResponse
                {
                    BangKeBanLeHangHoaDichVu = new M0304BangKeBanLeHangHoaDichVuPagedResult<M0304BangKeBanLeHangHoaDichVu>
                    {
                        Success = false,
                        Message = "Khong tim thay doanh nghiep.",
                        Data = null,
                        TotalRecords = 0,
                        TotalPages = 0,
                        CurrentPage = page
                    },
                    DoanhNghiep = null,
                    TongCong = 0,
                };
            }
            var allData = await _context.BangKeBanLeHangHoaDichVus
                .FromSqlRaw("EXEC dbo.[S0304_BangKeBanLeHangHoaDichVu] @TuNgay, @DenNgay, @IDCN, @IDKhoHang, @IDNhanVien",
                    new SqlParameter("@TuNgay", ngayBatDau),
                    new SqlParameter("@DenNgay", ngayKetThuc),
                    new SqlParameter("@IDCN", idCN),
                    new SqlParameter("@IDKhoHang", idKhoHang),
                    new SqlParameter("@IDNhanVien", idNhanVien))
                .AsNoTracking()
                .ToListAsync();

            double allTongSoTien = allData.Sum(x => x.ThanhTien) ?? 0f;

            var totalRecords = allData.Count;
            var totalPages = (int)Math.Ceiling((double)totalRecords / pageSize);
            var pagedData = allData.Skip((page - 1) * pageSize).Take(pageSize).ToList();

            string message = pagedData.Any()
                ? $"Tìm thấy {totalRecords} kết quả từ {ngayBatDau} đến {ngayKetThuc}."
                : $"Không tìm thấy kết quả nào từ {ngayBatDau} đến {ngayKetThuc}.";

            var sessionData = new
            {
                Data = allData,
                FromDate = ngayBatDau,
                ToDate = ngayKetThuc
            };
            session?.SetString("FilteredData", JsonConvert.SerializeObject(sessionData));

            return new M0304BangKeBanLeHangHoaDichVuResponse
            {
                BangKeBanLeHangHoaDichVu = new M0304BangKeBanLeHangHoaDichVuPagedResult<M0304BangKeBanLeHangHoaDichVu>
                {
                    Success = true,
                    Message = message,
                    Data = pagedData,
                    TotalRecords = totalRecords,
                    TotalPages = totalPages,
                    CurrentPage = page
                },
                DoanhNghiep = doanhNghiep,
                TongCong = allTongSoTien
            };
        }
        private M0304ThongTinDoanhNghiep GetDoanhNghiepFromRequestOrSession(ExportRequest request, ISession session)
        {
            M0304ThongTinDoanhNghiep doanhNghiepObj = null;
            try
            {
                if (request.DoanhNghiep != null)
                {
                    var json = JsonConvert.SerializeObject(request.DoanhNghiep);
                    doanhNghiepObj = JsonConvert.DeserializeObject<M0304ThongTinDoanhNghiep>(json);
                }

                if (doanhNghiepObj == null)
                {
                    var doanhNghiepJson = session.GetString("DoanhNghiepInfo");
                    if (!string.IsNullOrEmpty(doanhNghiepJson))
                    {
                        doanhNghiepObj = JsonConvert.DeserializeObject<M0304ThongTinDoanhNghiep>(doanhNghiepJson);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi parse doanh nghiep từ request hoặc session");
            }

            return doanhNghiepObj ?? new M0304ThongTinDoanhNghiep
            {
                TenCSKCB = "Tên đơn vị",
                DiaChi = "",
                DienThoai = ""
            };
        }

        private async Task<byte[]> ExportBangKeBanLeHangHoaDichVuExcelAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var logoPath = "";

            var data = request.Data ?? new List<M0304BangKeBanLeHangHoaDichVu>();
            var document = new P0304BangKeBanLeHangHoaDichVuExcelReportTemplate(data, request.FromDate, request.ToDate, doanhNghiepObj, logoPath);

            var excelBytes = document.GenerateExcel();
            return excelBytes;
        }

        private async Task<byte[]> ExportBangKeBanLeHangHoaDichVuPdfAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var logoPath ="";

            var data = request.Data ?? new List<M0304BangKeBanLeHangHoaDichVu>();
            var document = new P0304BangKeBanLeHangHoaDichVuReportTemplatePDF(data, request.FromDate, request.ToDate, doanhNghiepObj, logoPath);

            var pdfBytes = document.GeneratePdf();
            return pdfBytes;
        }

    }
}