using _0304Huy_Report.Models;
using C0304.Db.Models;
using M0304.Models.ThongTinDoanhNghiep;
using M0304.Models.BaoCaoChiDinhCLS_Phong_BS;
using P0304.PDFDocument.BaoCaoChiDinhCLS_Phong_BS;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using QuestPDF.Fluent;
using System.Diagnostics;

namespace C0304BaoCaoChiDinhCLS_Phong_BS.Controllers
{
    [Route("bao_cao_chi_dinh_cls_theo_phong_bs")]
    public class C0304BaoCaoChiDinhCLS_Phong_BSController : Controller
    {
        //private string _maChucNang = "/bao_cao_chi_dinh_cls_theo_phong_bs";
        //private IMemoryCachingServices _memoryCache;
        private readonly ILogger<C0304BaoCaoChiDinhCLS_Phong_BSController> _logger;
        private readonly M0304Context _context;
        private readonly I0304ThongTinDoanhNghiep _thongTinDoanhNghiepService;
        private readonly IWebHostEnvironment _env;

        public C0304BaoCaoChiDinhCLS_Phong_BSController(ILogger<C0304BaoCaoChiDinhCLS_Phong_BSController> logger,
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

            return View("~/Views/V0304/V0304BaoCaoChiDinhCLS_Phong_BS/Index.cshtml");
        }

        [HttpPost("filter")]
        public async Task<IActionResult> FilterByDay(string tuNgay, string denNgay, long IdChiNhanh, long? idPhongBuong = null,
            long? idBacSi = null, long? idDVKT = null, int page = 1, int pageSize = 20)
        {
            try
            {
                var result = await GetBaoCaoChiDinhCLS_Phong_BS(tuNgay, denNgay, IdChiNhanh, idPhongBuong, idBacSi, idDVKT, page, pageSize);

                if (!result.BaoCaoChiDinhCLS_Phong_BS.Success)
                {
                    _logger.LogWarning("Service trả về lỗi: {Message}", result.BaoCaoChiDinhCLS_Phong_BS.Message);
                    return Json(new { success = false, message = result.BaoCaoChiDinhCLS_Phong_BS.Message });
                }

                return Json(new
                {
                    success = true,
                    message = result.BaoCaoChiDinhCLS_Phong_BS.Message,
                    data = result.BaoCaoChiDinhCLS_Phong_BS.Data,
                    totalRecords = result.BaoCaoChiDinhCLS_Phong_BS.TotalRecords,
                    totalPages = result.BaoCaoChiDinhCLS_Phong_BS.TotalPages,
                    currentPage = result.BaoCaoChiDinhCLS_Phong_BS.CurrentPage,
                    doanhNghiep = result.DoanhNghiep,
                    tongDonGia = result.TongDonGia,
                    tongLuot = result.TongLuot
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
            var pdfBytes = await ExportBaoCaoChiDinhCLS_Phong_BSPdfAsync(request, HttpContext.Session);

            string fileName = $"BaoCaoChiDinhCLS_Phong_BS_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.pdf";
            return File(pdfBytes, "application/pdf", fileName);
        }

        [HttpPost("export/excel")]
        public async Task<IActionResult> ExportToExcel([FromBody] ExportRequest request)
        {
            var excelBytes = await ExportBaoCaoChiDinhCLS_Phong_BSExcelAsync(request, HttpContext.Session);

            string fileName = $"BaoCaoChiDinhCLS_Phong_BS_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.xlsx";
            return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }


        private async Task<M0304BaoCaoChiDinhCLS_Phong_BSResponse> GetBaoCaoChiDinhCLS_Phong_BS(string ngayBatDau, string ngayKetThuc, long idCN,
            long? idPhongBuong = null, long? idBacSi = null, long? idDVKT = null, int page = 1, int pageSize = 20)
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
                return new M0304BaoCaoChiDinhCLS_Phong_BSResponse
                {
                    BaoCaoChiDinhCLS_Phong_BS = new M0304BaoCaoChiDinhCLS_Phong_BSPagedResult<M0304BaoCaoChiDinhCLS_Phong_BS>
                    {
                        Success = false,
                        Message = "Khong tim thay doanh nghiep.",
                        Data = null,
                        TotalRecords = 0,
                        TotalPages = 0,
                        CurrentPage = page
                    },
                    DoanhNghiep = null,
                    TongDonGia = 0,
                    TongLuot = 0
                };
            }
            var allData = await _context.BaoCaoChiDinhCLS_Phong_BSs
                .FromSqlRaw("EXEC dbo.[S0304_BaoCaoChiDinhCLS_Phong_BS] @TuNgay, @DenNgay, @IDCN, @IDPhongBuong, @IDBacSi, @IDDVKT",
                    new SqlParameter("@TuNgay", ngayBatDau),
                    new SqlParameter("@DenNgay", ngayKetThuc),
                    new SqlParameter("@IDCN", idCN),
                    new SqlParameter("@IDPhongBuong", idPhongBuong),
                    new SqlParameter("@IDBacSi", idBacSi),
                    new SqlParameter("@IDDVKT", idDVKT))
                .AsNoTracking()
                .ToListAsync();

            var allDonGia = allData.Sum(x => x.DonGia);
            var allSoLuot = allData.Sum(x => x.SoLuot);

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

            return new M0304BaoCaoChiDinhCLS_Phong_BSResponse
            {
                BaoCaoChiDinhCLS_Phong_BS = new M0304BaoCaoChiDinhCLS_Phong_BSPagedResult<M0304BaoCaoChiDinhCLS_Phong_BS>
                {
                    Success = true,
                    Message = message,
                    Data = pagedData,
                    TotalRecords = totalRecords,
                    TotalPages = totalPages,
                    CurrentPage = page
                },
                DoanhNghiep = doanhNghiep,
                TongDonGia = allDonGia,
                TongLuot = allSoLuot
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

        private async Task<byte[]> ExportBaoCaoChiDinhCLS_Phong_BSExcelAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var logoPath = "";

            var data = request.Data ?? new List<M0304BaoCaoChiDinhCLS_Phong_BS>();
            var document = new P0304BaoCaoChiDinhCLS_Phong_BSExcelReportTemplate(data, request.FromDate, request.ToDate, doanhNghiepObj, logoPath);

            var excelBytes = document.GenerateExcel();
            return excelBytes;
        }

        private async Task<byte[]> ExportBaoCaoChiDinhCLS_Phong_BSPdfAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var logoPath ="";

            var data = request.Data ?? new List<M0304BaoCaoChiDinhCLS_Phong_BS>();
            var document = new P0304BaoCaoChiDinhCLS_Phong_BSReportTemplate(data, request.FromDate, request.ToDate, doanhNghiepObj, logoPath);

            var pdfBytes = document.GeneratePdf();
            return pdfBytes;
        }

    }
}