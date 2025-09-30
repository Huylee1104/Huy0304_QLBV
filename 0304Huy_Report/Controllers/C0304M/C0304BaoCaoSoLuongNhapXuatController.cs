using _0304Huy_Report.Models;
using C0304.Db.Models;
using M0304.Models.ThongTinDoanhNghiep;
using M0304M.Models.BaoCaoHangHoa;
using M0304NhanVien.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using P0304M.PDFDocument;
using QuestPDF.Fluent;
using System.Diagnostics;

namespace C0304MBaoCaoSoLuongNhapXuat.Controllers
{
    [Route("bao_cao_so_luong_nhap_xuat")]
    public class C0304MBaoCaoSoLuongNhapXuatController : Controller
    {
        //private string _maChucNang = "/bao_cao_so_luong_nhap_xuat";
        //private IMemoryCachingServices _memoryCache;

        private readonly ILogger<C0304MBaoCaoSoLuongNhapXuatController> _logger;
        private readonly M0304Context _context;
        private readonly I0304ThongTinDoanhNghiep _thongTinDoanhNghiepService;

        public C0304MBaoCaoSoLuongNhapXuatController(M0304Context context, 
            ILogger<C0304MBaoCaoSoLuongNhapXuatController> logger,
            I0304ThongTinDoanhNghiep thongTinDoanhNghiepService /*, IMemoryCachingServices memoryCache*/)
        {
            _context = context;
            _logger = logger;
            _thongTinDoanhNghiepService = thongTinDoanhNghiepService;
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

            return View("~/Views/V0304MBaoCaoSoLuongNhapXuat/Index.cshtml");
        }

        [HttpPost("filterNhap")]
        public async Task<IActionResult> FilterByDayNhap(string tuNgay, string denNgay, long IdChiNhanh,
            long? idKhoHang, long? idNhomHang, long? idHangHoa, int page = 1, int pageSize = 20)
        {
            try
            {
                var result = await GetHangNhap(tuNgay, denNgay, IdChiNhanh, idKhoHang, idNhomHang, idHangHoa, page, pageSize);

                if (!result.HangNhap.Success)
                {
                    _logger.LogWarning("Service trả về lỗi: {Message}", result.HangNhap.Message);
                    return Json(new { success = false, message = result.HangNhap.Message });
                }

                return Json(new
                {
                    success = true,
                    message = result.HangNhap.Message,
                    data = result.HangNhap.Data,
                    totalRecords = result.HangNhap.TotalRecords,
                    totalPages = result.HangNhap.TotalPages,
                    currentPage = result.HangNhap.CurrentPage,
                    doanhNghiep = result.DoanhNghiep
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi xảy ra trong FilterByDay");
                return StatusCode(500, "Lỗi server: " + ex.Message);
            }

        }

        [HttpPost("filterXuat")]
        public async Task<IActionResult> FilterByDayXuat(string tuNgay, string denNgay, long IdChiNhanh,
            long? idKhoHang, long? idNhomHang, long? idHangHoa, int page = 1, int pageSize = 20)
        {
            try
            {
                var result = await GetHangXuat(tuNgay, denNgay, IdChiNhanh, idKhoHang, idNhomHang, idHangHoa, page, pageSize);

                if (!result.HangXuat.Success)
                {
                    _logger.LogWarning("Service trả về lỗi: {Message}", result.HangXuat.Message);
                    return Json(new { success = false, message = result.HangXuat.Message });
                }

                return Json(new
                {
                    success = true,
                    message = result.HangXuat.Message,
                    data = result.HangXuat.Data,
                    totalRecords = result.HangXuat.TotalRecords,
                    totalPages = result.HangXuat.TotalPages,
                    currentPage = result.HangXuat.CurrentPage,
                    doanhNghiep = result.DoanhNghiep
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi xảy ra trong FilterByDay");
                return StatusCode(500, "Lỗi server: " + ex.Message);
            }

        }

        [HttpPost("exportnhap/pdf")]
        public async Task<IActionResult> ExportNhapToPDF([FromBody] ExportRequest<M0304MHangNhap> request)
        {
            var pdfBytes = await ExportHangNhapPdfAsync(request, HttpContext.Session);

            string fileName = $"HangNhap_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.pdf";
            return File(pdfBytes, "application/pdf", fileName);
        }

        [HttpPost("exportnhap/excel")]
        public async Task<IActionResult> ExportNhapToExcel([FromBody] ExportRequest<M0304MHangNhap> request)
        {
            var excelBytes = await ExportHangNhapExcelAsync(request, HttpContext.Session);

            string fileName = $"HangNhap_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.xlsx";
            return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        //[HttpPost("exportxuat/pdf")]
        //public async Task<IActionResult> ExportXuatToPDF([FromBody] ExportRequest<M0304MHangXuat> request)
        //{
        //    var pdfBytes = await ExportHangXuatPdfAsync(request, HttpContext.Session);

        //    string fileName = $"HangNhap_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.pdf";
        //    return File(pdfBytes, "application/pdf", fileName);
        //}

        //[HttpPost("exportxuat/excel")]
        //public async Task<IActionResult> ExportXuatToXExcel([FromBody] ExportRequest<M0304MHangXuat> request)
        //{
        //    var excelBytes = await ExportHangXuatExcelAsync(request, HttpContext.Session);

        //    string fileName = $"HangNhap_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.xlsx";
        //    return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        //}

        // Các hàm xử lý nhập==================
        private async Task<M0304MHangNhapResponse> GetHangNhap(string ngayBatDau, string ngayKetThuc, long idCN, long? idKhoHang,
            long? idNhomHang = null, long? idHangHoa = null, int page = 1, int pageSize = 20)
        {
            var doanhNghiep = await _thongTinDoanhNghiepService.GetThongTinDoanhNghiep(idCN);

            var session = HttpContext.Session;

            if (doanhNghiep != null)
            {
                // Lưu thông tin doanh nghiệp vào session
                session?.SetString("DoanhNghiepInfo", JsonConvert.SerializeObject(doanhNghiep));
                _logger.LogInformation("Doanh Nghiep Info: {@DoanhNghiep}", doanhNghiep);
            }
            else
            {
                _logger.LogWarning("No doanh nghiep found for ChiNhanh ID: {IdChiNhanh}", idCN);
                return new M0304MHangNhapResponse
                {
                    HangNhap = new M0304MPagedResult<M0304MHangNhap>
                    {
                        Success = false,
                        Message = "Khong tim thay doanh nghiep.",
                        Data = null,         // không có dữ liệu
                        TotalRecords = 0,
                        TotalPages = 0,
                        CurrentPage = page
                    },
                    DoanhNghiep = null     // không có thông tin doanh nghiệp
                };
            }
            var allData = await _context.HangNhapReports
                .FromSqlRaw("EXEC dbo.[S0304_BaoCaoSoLuongHangNhap] @TuNgay, @DenNgay, @IDCN, @IDKhoHang, @IDNhomHang, @IDHangHoa",
                    new SqlParameter("@TuNgay", ngayBatDau),
                    new SqlParameter("@DenNgay", ngayKetThuc),
                    new SqlParameter("@IDCN", idCN),
                    new SqlParameter("@IDKhoHang", idKhoHang),
                    new SqlParameter("@IDNhomHang", idNhomHang),
                    new SqlParameter("@IDHangHoa", idHangHoa))
                .AsNoTracking()
                .ToListAsync();

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

            return new M0304MHangNhapResponse
            {
                HangNhap = new M0304MPagedResult<M0304MHangNhap>
                {
                    Success = true,
                    Message = message,
                    Data = pagedData,
                    TotalRecords = totalRecords,
                    TotalPages = totalPages,
                    CurrentPage = page
                },
                DoanhNghiep = doanhNghiep
            };
        }
        private M0304ThongTinDoanhNghiep GetDoanhNghiepFromRequestOrSession<T>(ExportRequest<T> request, ISession session)
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

        private async Task<byte[]> ExportHangNhapPdfAsync(ExportRequest<M0304MHangNhap> request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var logoPath = "null";

            var data = request.Data ?? new List<M0304MHangNhap>();
            var document = new P0304MReportNhapTemplatePDF(data, request.FromDate, request.ToDate, doanhNghiepObj);

            var pdfBytes = document.GeneratePdf();
            return pdfBytes;
        }
        private async Task<byte[]> ExportHangNhapExcelAsync(ExportRequest<M0304MHangNhap> request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var logoPath = "";

            var data = request.Data ?? new List<M0304MHangNhap>();
            var document = new P0304ExcelReportNhapTemplate(data, request.FromDate, request.ToDate, doanhNghiepObj);

            var excelBytes = document.GenerateExcel();
            return excelBytes;
        }

        // Các hàm xử lý xuất ==========
        private async Task<M0304GHangXuatResponse> GetHangXuat(string ngayBatDau, string ngayKetThuc, long idCN, long? idKhoHang,
            long? idNhomHang = null, long? idHangHoa = null, int page = 1, int pageSize = 20)
        {
            var doanhNghiep = await _thongTinDoanhNghiepService.GetThongTinDoanhNghiep(idCN);

            var session = HttpContext.Session;

            if (doanhNghiep != null)
            {
                // Lưu thông tin doanh nghiệp vào session
                session?.SetString("DoanhNghiepInfo", JsonConvert.SerializeObject(doanhNghiep));
                _logger.LogInformation("Doanh Nghiep Info: {@DoanhNghiep}", doanhNghiep);
            }
            else
            {
                _logger.LogWarning("No doanh nghiep found for ChiNhanh ID: {IdChiNhanh}", idCN);
                return new M0304GHangXuatResponse
                {
                    HangXuat = new M0304MPagedResult<M0304MHangXuat>
                    {
                        Success = false,
                        Message = "Khong tim thay doanh nghiep.",
                        Data = null,         // không có dữ liệu
                        TotalRecords = 0,
                        TotalPages = 0,
                        CurrentPage = page
                    },
                    DoanhNghiep = null        // không có thông tin doanh nghiệp
                };
            }
            var allData = await _context.HangXuatReports
                .FromSqlRaw("EXEC dbo.[S0304_BaoCaoSoLuongHangXuat] @TuNgay, @DenNgay, @IDCN, @IDKhoHang, @IDNhomHang, @IDHangHoa",
                    new SqlParameter("@TuNgay", ngayBatDau),
                    new SqlParameter("@DenNgay", ngayKetThuc),
                    new SqlParameter("@IDCN", idCN),
                    new SqlParameter("@IDKhoHang", idKhoHang),
                    new SqlParameter("@IDNhomHang", idNhomHang),
                    new SqlParameter("@IDHangHoa", idHangHoa))
                .AsNoTracking()
                .ToListAsync();

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

            return new M0304GHangXuatResponse
            {
                HangXuat = new M0304MPagedResult<M0304MHangXuat>
                {
                    Success = true,
                    Message = message,
                    Data = pagedData,
                    TotalRecords = totalRecords,
                    TotalPages = totalPages,
                    CurrentPage = page
                },
                DoanhNghiep = doanhNghiep
            };
        }

        //private async Task<byte[]> ExportHangXuatPdfAsync(ExportRequest<M0304MHangXuat> request, ISession session)
        //{
        //    var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
        //    var logoPath = "null";

        //    var data = request.Data ?? new List<M0304MHangXuat>();
        //    var document = new P0304ReportTemplatePDF(data, request.FromDate, request.ToDate,
        //         doanhNghiepObj, logoPath);

        //    var pdfBytes = document.GeneratePdf();
        //    return pdfBytes;
        //}
        //private async Task<byte[]> ExportHangXuatExcelAsync(ExportRequest<M0304MHangXuat> request, ISession session)
        //{
        //    var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
        //    var logoPath = "";

        //    var data = request.Data ?? new List<M0304MHangXuat>();
        //    var document = new P0304ExcelReportTemplate(data, request.FromDate, request.ToDate,
        //        doanhNghiepObj, logoPath);

        //    var excelBytes = document.GenerateExcel();
        //    return excelBytes;
        //}
    }

}