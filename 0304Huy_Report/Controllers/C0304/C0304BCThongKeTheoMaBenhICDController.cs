using _0304Huy_Report.Models;
using C0304.Db.Models;
using DocumentFormat.OpenXml.Wordprocessing;
using M0304.Models.BCThongKeTheoMaBenhICD;
using M0304.Models.ThongTinDoanhNghiep;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using P0304.PDFDocument.BCThongKeTheoMaBenhICD;
using QuestPDF.Fluent;
using System.Diagnostics;

namespace C0304BCThongKeTheoMaBenhICD.Controllers
{
    [Route("bc_thong_ke_theo_ma_benh_icd")]
    public class C0304BCThongKeTheoMaBenhICDController : Controller
    {
        //private string _maChucNang = "/bc_thong_ke_theo_ma_benh_icd";
        //private IMemoryCachingServices _memoryCache;
        private readonly ILogger<C0304BCThongKeTheoMaBenhICDController> _logger;
        private readonly M0304Context _context;
        private readonly I0304ThongTinDoanhNghiep _thongTinDoanhNghiepService;
        private readonly IWebHostEnvironment _env;

        public C0304BCThongKeTheoMaBenhICDController(ILogger<C0304BCThongKeTheoMaBenhICDController> logger,
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

            return View("~/Views/V0304/V0304BCThongKeTheoMaBenhICD/Index.cshtml");
        }

        [HttpPost("filter")]
        public async Task<IActionResult> FilterByDay(string tuNgay, string denNgay, long IdChiNhanh, int page = 1, int pageSize = 20)
        {
            try
            {
                var result = await GetBCThongKeTheoMaBenhICD(tuNgay, denNgay, IdChiNhanh, page, pageSize);

                if (!result.BCThongKeTheoMaBenhICD.Success)
                {
                    _logger.LogWarning("Service trả về lỗi: {Message}", result.BCThongKeTheoMaBenhICD.Message);
                    return Json(new { success = false, message = result.BCThongKeTheoMaBenhICD.Message });
                }

                return Json(new
                {
                    success = true,
                    message = result.BCThongKeTheoMaBenhICD.Message,
                    data = result.BCThongKeTheoMaBenhICD.Data,
                    totalRecords = result.BCThongKeTheoMaBenhICD.TotalRecords,
                    totalPages = result.BCThongKeTheoMaBenhICD.TotalPages,
                    currentPage = result.BCThongKeTheoMaBenhICD.CurrentPage,
                    doanhNghiep = result.DoanhNghiep,
                    tongTiepNhan = result.TongLuotTiepNhan,
                    tongNam = result.TongNam,
                    tongNu = result.TongNu,
                    tongCoBHYT = result.TongCoBHYT,
                    tongKhongBHYT = result.TongKHongBHYT
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
            var pdfBytes = await ExportBCThongKeTheoMaBenhICDPdfAsync(request, HttpContext.Session);

            string fileName = $"BCThongKeTheoMaBenhICD_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.pdf";
            return File(pdfBytes, "application/pdf", fileName);
        }

        [HttpPost("export/excel")]
        public async Task<IActionResult> ExportToExcel([FromBody] ExportRequest request)
        {
            var excelBytes = await ExportBCThongKeTheoMaBenhICDExcelAsync(request, HttpContext.Session);

            string fileName = $"BCThongKeTheoMaBenhICD_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.xlsx";
            return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }


        private async Task<M0304BCThongKeTheoMaBenhICDResponse> GetBCThongKeTheoMaBenhICD(string ngayBatDau, string ngayKetThuc, long idCN,
             int page = 1, int pageSize = 20)
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
                return new M0304BCThongKeTheoMaBenhICDResponse
                {
                    BCThongKeTheoMaBenhICD = new M0304BCThongKeTheoMaBenhICDPagedResult<M0304BCThongKeTheoMaBenhICD>
                    {
                        Success = false,
                        Message = "Khong tim thay doanh nghiep.",
                        Data = null,
                        TotalRecords = 0,
                        TotalPages = 0,
                        CurrentPage = page
                    },
                    DoanhNghiep = null,
                    TongLuotTiepNhan = 0,
                    TongNam = 0,
                    TongNu = 0,
                    TongCoBHYT = 0,
                    TongKHongBHYT = 0
                };
            }
            var allData = await _context.BCThongKeTheoMaBenhICDs
                .FromSqlRaw("EXEC dbo.[S0304_BCThongKeTheoMaBenhICD] @TuNgay, @DenNgay, @IDCN",
                    new SqlParameter("@TuNgay", ngayBatDau),
                    new SqlParameter("@DenNgay", ngayKetThuc),
                    new SqlParameter("@IDCN", idCN))
                .AsNoTracking()
                .ToListAsync();

            var allTiepNhan = allData.Sum(x => x.SoLuotTiepNhan);
            var allNam = allData.Sum(x => x.SoLuongNam);
            var allNu = allData.Sum(x => x.SoLuongNu);
            var allCoBHYT = allData.Sum(x => x.CoBHYT);
            var allKhongBHYT = allData.Sum(x => x.KhongBHYT);

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

            return new M0304BCThongKeTheoMaBenhICDResponse
            {
                BCThongKeTheoMaBenhICD = new M0304BCThongKeTheoMaBenhICDPagedResult<M0304BCThongKeTheoMaBenhICD>
                {
                    Success = true,
                    Message = message,
                    Data = pagedData,
                    TotalRecords = totalRecords,
                    TotalPages = totalPages,
                    CurrentPage = page
                },
                DoanhNghiep = doanhNghiep,
                TongLuotTiepNhan = allTiepNhan,
                TongNam = allNam,
                TongNu = allNu,
                TongCoBHYT = allCoBHYT,
                TongKHongBHYT = allKhongBHYT,
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

        private async Task<byte[]> ExportBCThongKeTheoMaBenhICDExcelAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var logoPath = "";

            var data = request.Data ?? new List<M0304BCThongKeTheoMaBenhICD>();
            var document = new P0304BCThongKeTheoMaBenhICDExcelReportTemplate(data, request.FromDate, request.ToDate, request.TenNVDN, doanhNghiepObj, logoPath);

            var excelBytes = document.GenerateExcel();
            return excelBytes;
        }

        private async Task<byte[]> ExportBCThongKeTheoMaBenhICDPdfAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var logoPath ="";

            var data = request.Data ?? new List<M0304BCThongKeTheoMaBenhICD>();
            var document = new P0304BCThongKeTheoMaBenhICDReportTemplate(data, request.FromDate, request.ToDate, request.TenNVDN, doanhNghiepObj, logoPath);

            var pdfBytes = document.GeneratePdf();
            return pdfBytes;
        }

    }
}