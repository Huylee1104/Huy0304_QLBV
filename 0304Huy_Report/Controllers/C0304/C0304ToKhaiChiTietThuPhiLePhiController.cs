using _0304Huy_Report.Models;
using C0304.Db.Models;
using M0304.Models.ThongTinDoanhNghiep;
using M0304.Models.ToKhaiChiTietThuPhiLePhi;
using P0304.PDFDocument.ToKhaiChiTietThuPhiLePhi;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using QuestPDF.Fluent;
using System.Diagnostics;

namespace C0304ToKhaiChiTietThuPhiLePhi.Controllers
{
    [Route("to_khai_chi_tiet_thu_phi_le_phi")]
    public class C0304ToKhaiChiTietThuPhiLePhiController : Controller
    {
        //private string _maChucNang = "/to_khai_chi_tiet_thu_phi_le_phi";
        //private IMemoryCachingServices _memoryCache;
        private readonly ILogger<C0304ToKhaiChiTietThuPhiLePhiController> _logger;
        private readonly M0304Context _context;
        private readonly I0304ThongTinDoanhNghiep _thongTinDoanhNghiepService;
        private readonly IWebHostEnvironment _env;

        public C0304ToKhaiChiTietThuPhiLePhiController(ILogger<C0304ToKhaiChiTietThuPhiLePhiController> logger,
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

            return View("~/Views/V0304/V0304ToKhaiChiTietThuPhiLePhi/Index.cshtml");
        }

        [HttpPost("filter")]
        public async Task<IActionResult> FilterByDay(string tuNgay, string denNgay, long IdChiNhanh, long? idNhanVien = null,
            long? idHTTT = null, long? idLoai = null, int page = 1, int pageSize = 20)
        {
            try
            {
                var result = await GetToKhaiChiTietThuPhiLePhi(tuNgay, denNgay, IdChiNhanh, idNhanVien, idHTTT, idLoai, page, pageSize);

                if (!result.ToKhaiChiTietThuPhiLePhi.Success)
                {
                    _logger.LogWarning("Service trả về lỗi: {Message}", result.ToKhaiChiTietThuPhiLePhi.Message);
                    return Json(new { success = false, message = result.ToKhaiChiTietThuPhiLePhi.Message });
                }

                return Json(new
                {
                    success = true,
                    message = result.ToKhaiChiTietThuPhiLePhi.Message,
                    data = result.ToKhaiChiTietThuPhiLePhi.Data,
                    totalRecords = result.ToKhaiChiTietThuPhiLePhi.TotalRecords,
                    totalPages = result.ToKhaiChiTietThuPhiLePhi.TotalPages,
                    currentPage = result.ToKhaiChiTietThuPhiLePhi.CurrentPage,
                    doanhNghiep = result.DoanhNghiep,
                    AllSoTien = result.AllSoTien,
                    AllHoan_Huy = result.AllHoan_Huy,
                    AllTienThucThu = result.AllTienThucThu
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
            var pdfBytes = await ExportToKhaiChiTietThuPhiLePhiPdfAsync(request, HttpContext.Session);

            string fileName = $"ToKhaiChiTietThuPhiLePhi_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.pdf";
            return File(pdfBytes, "application/pdf", fileName);
        }

        [HttpPost("export/excel")]
        public async Task<IActionResult> ExportToExcel([FromBody] ExportRequest request)
        {
            var excelBytes = await ExportToKhaiChiTietThuPhiLePhiExcelAsync(request, HttpContext.Session);

            string fileName = $"ToKhaiChiTietThuPhiLePhi_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.xlsx";
            return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }


        private async Task<M0304ToKhaiChiTietThuPhiLePhiResponse> GetToKhaiChiTietThuPhiLePhi(string ngayBatDau, string ngayKetThuc, long idCN,
            long? idNhanVien = null, long? idHTTT = null, long? idLoai = null, int page = 1, int pageSize = 20)
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
                return new M0304ToKhaiChiTietThuPhiLePhiResponse
                {
                    ToKhaiChiTietThuPhiLePhi = new M0304ToKhaiChiTietThuPhiLePhiPagedResult<M0304ToKhaiChiTietThuPhiLePhi>
                    {
                        Success = false,
                        Message = "Khong tim thay doanh nghiep.",
                        Data = null,
                        TotalRecords = 0,
                        TotalPages = 0,
                        CurrentPage = page
                    },
                    DoanhNghiep = null,
                    AllSoTien = 0,
                    AllHoan_Huy = 0,
                    AllTienThucThu = 0,
                };
            }
            var allData = await _context.ToKhaiChiTietThuPhiLePhis
                .FromSqlRaw("EXEC dbo.[S0304_ToKhaiChiTietThuPhiLePhi] @TuNgay, @DenNgay, @IDCN, @IDNhanVien, @IDHTTT, @IDLoai",
                    new SqlParameter("@TuNgay", ngayBatDau),
                    new SqlParameter("@DenNgay", ngayKetThuc),
                    new SqlParameter("@IDCN", idCN),
                    new SqlParameter("@IDNhanVien", idNhanVien),
                    new SqlParameter("@IDNhanVien", idHTTT),
                    new SqlParameter("@IDNhanVien", idLoai))
                .AsNoTracking()
                .ToListAsync();

            var allTongSoTien = allData.Sum(x => x.TongSoTien);
            var allHuy_Hoan = allData.Sum(x => x.Huy_Hoan);
            var allSoTienThucThu = allData.Sum(x => x.SoTienThucThu);

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

            return new M0304ToKhaiChiTietThuPhiLePhiResponse
            {
                ToKhaiChiTietThuPhiLePhi = new M0304ToKhaiChiTietThuPhiLePhiPagedResult<M0304ToKhaiChiTietThuPhiLePhi>
                {
                    Success = true,
                    Message = message,
                    Data = pagedData,
                    TotalRecords = totalRecords,
                    TotalPages = totalPages,
                    CurrentPage = page
                },
                DoanhNghiep = doanhNghiep,
                AllSoTien = allTongSoTien,
                AllHoan_Huy = allHuy_Hoan,
                AllTienThucThu = allSoTienThucThu,
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

        private async Task<byte[]> ExportToKhaiChiTietThuPhiLePhiExcelAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var logoPath = "";

            var data = request.Data ?? new List<M0304ToKhaiChiTietThuPhiLePhi>();
            var document = new P0304ToKhaiChiTietThuPhiLePhiExcelReportTemplate(data, request.FromDate, request.ToDate, request.TenNVDN, doanhNghiepObj, logoPath);

            var excelBytes = document.GenerateExcel();
            return excelBytes;
        }

        private async Task<byte[]> ExportToKhaiChiTietThuPhiLePhiPdfAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var logoPath ="";

            var data = request.Data ?? new List<M0304ToKhaiChiTietThuPhiLePhi>();
            var document = new P0304ToKhaiChiTietThuPhiLePhiReportTemplate(data, request.FromDate, request.ToDate, request.TenNVDN, doanhNghiepObj, logoPath);

            var pdfBytes = document.GeneratePdf();
            return pdfBytes;
        }

    }
}