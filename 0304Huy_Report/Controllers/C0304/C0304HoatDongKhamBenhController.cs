using _0304Huy_Report.Models;
using C0304.Db.Models;
using M0304.Models.ThongTinDoanhNghiep;
using M0304.Models.HoatDongKhamBenh;
using P0304.PDFDocument.HoatDongKhamBenh;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using QuestPDF.Fluent;
using System.Diagnostics;

namespace C0304HoatDongKhamBenh.Controllers
{
    [Route("hoat_dong_kham_benh")]
    public class C0304HoatDongKhamBenhController : Controller
    {
        //private string _maChucNang = "/hoat_dong_kham_benh";
        //private IMemoryCachingServices _memoryCache;
        private readonly ILogger<C0304HoatDongKhamBenhController> _logger;
        private readonly M0304Context _context;
        private readonly I0304ThongTinDoanhNghiep _thongTinDoanhNghiepService;
        private readonly IWebHostEnvironment _env;

        public C0304HoatDongKhamBenhController(ILogger<C0304HoatDongKhamBenhController> logger,
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

            return View("~/Views/V0304/V0304HoatDongKhamBenh/Index.cshtml");
        }

        [HttpPost("filter")]
        public async Task<IActionResult> FilterByDay(string tuNgay, string denNgay, long IdChiNhanh, long? idNhanVien = null, int page = 1, int pageSize = 20)
        {
            try
            {
                var result = await GetHoatDongKhamBenh(tuNgay, denNgay, IdChiNhanh, idNhanVien, page, pageSize);

                if (!result.HoatDongKhamBenh.Success)
                {
                    _logger.LogWarning("Service trả về lỗi: {Message}", result.HoatDongKhamBenh.Message);
                    return Json(new { success = false, message = result.HoatDongKhamBenh.Message });
                }

                return Json(new
                {
                    success = true,
                    message = result.HoatDongKhamBenh.Message,
                    data = result.HoatDongKhamBenh.Data,
                    totalRecords = result.HoatDongKhamBenh.TotalRecords,
                    totalPages = result.HoatDongKhamBenh.TotalPages,
                    currentPage = result.HoatDongKhamBenh.CurrentPage,
                    doanhNghiep = result.DoanhNghiep,
                    AllTongSo = result.AllTongSo,
                    AllYHCT = result.AllYHocCoTruyen,
                    AllTreEmDuoi6 = result.AllTEDuoi6Tuoi,
                    AllBHYT = result.AllBHYT,
                    AllVienPhi = result.AllVienPhi,
                    AllKhongThuDuoc = result.AllKhongThuDuoc,
                    AllCapCuu = result.AllCapCuu,
                    AllSoNguoiVaoVien = result.AllSoNguoiVaoVien,
                    AllSoNguoiChuyenVien = result.AllSoNguoiChuyenVien,
                    AllSoNguoiBenh = result.AllNTSoNguoiBenh,
                    AllNTYHCT = result.AllNTYHocCoTruyen,
                    AllNTTreEmDuoi6 = result.AllNTTEDuoi6Tuoi,
                    AllSoNgay = result.AllNTSoNgay,

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
            var pdfBytes = await ExportHoatDongKhamBenhPdfAsync(request, HttpContext.Session);

            string fileName = $"HoatDongKhamBenh_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.pdf";
            return File(pdfBytes, "application/pdf", fileName);
        }

        //[HttpPost("export/excel")]
        //public async Task<IActionResult> ExportToExcel([FromBody] ExportRequest request)
        //{
        //    var excelBytes = await ExportHoatDongKhamBenhExcelAsync(request, HttpContext.Session);

        //    string fileName = $"HoatDongKhamBenh_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.xlsx";
        //    return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        //}


        private async Task<M0304HoatDongKhamBenhResponse> GetHoatDongKhamBenh(string ngayBatDau, string ngayKetThuc, long idCN,
            long? idNhanVien = null, int page = 1, int pageSize = 20)
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
                return new M0304HoatDongKhamBenhResponse
                {
                    HoatDongKhamBenh = new M0304HoatDongKhamBenhPagedResult<M0304HoatDongKhamBenh>
                    {
                        Success = false,
                        Message = "Khong tim thay doanh nghiep.",
                        Data = null,
                        TotalRecords = 0,
                        TotalPages = 0,
                        CurrentPage = page
                    },
                    DoanhNghiep = null,
                    AllTongSo = 0,
                    AllYHocCoTruyen = 0,
                    AllTEDuoi6Tuoi = 0,
                    AllBHYT = 0,
                    AllVienPhi = 0,
                    AllKhongThuDuoc = 0,
                    AllCapCuu = 0,
                    AllSoNguoiVaoVien = 0,
                    AllSoNguoiChuyenVien = 0,
                    AllNTSoNguoiBenh = 0,
                    AllNTYHocCoTruyen = 0,
                    AllNTTEDuoi6Tuoi = 0,
                    AllNTSoNgay = 0,
                };
            }
            var allData = await _context.HoatDongKhamBenhs
                .FromSqlRaw("EXEC dbo.[S0304_HoatDongKhamBenh] @TuNgay, @DenNgay, @IDCN",
                    new SqlParameter("@TuNgay", ngayBatDau),
                    new SqlParameter("@DenNgay", ngayKetThuc),
                    new SqlParameter("@IDCN", idCN))
                    //new SqlParameter("@IDNhanVien", idNhanVien))
                .AsNoTracking()
                .ToListAsync();

            var allTongSoTien = allData.Sum(x => x.TongSo);
            var allY_HOC_CO_TRUYEN = allData.Sum(x => x.YHocCoTruyen);
            var TRE_EM_DUOI_6 = allData.Sum(x => x.TreEmDuoi6Tuoi);
            var BHYT = allData.Sum(x => x.BHYT);
            var VIEN_PHI = allData.Sum(x => x.VienPhi);
            var KHONG_THU_DUOC = allData.Sum(x => x.KhongThuDuoc);
            var CAP_CUU = allData.Sum(x => x.CapCuu);
            var SO_NGUOI_VAO_VIEN = allData.Sum(x => x.SoNguoiVaoVien);
            var SO_NGUOI_CHUYEN_VIEN = allData.Sum(x => x.SoNguoiChuyenVien);
            var NT_SO_NGUOI_BENH = allData.Sum(x => x.NTSoNguoiBenh);
            var NT_YHCT = allData.Sum(x => x.NTYHocCoTruyen);
            var NT_TRE_EM_DUOI_6 = allData.Sum(x => x.NTTreEmDuoi6Tuoi);
            var NT_SO_NGAY = allData.Sum(x => x.NTSoNgay);

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

            return new M0304HoatDongKhamBenhResponse
            {
                HoatDongKhamBenh = new M0304HoatDongKhamBenhPagedResult<M0304HoatDongKhamBenh>
                {
                    Success = true,
                    Message = message,
                    Data = pagedData,
                    TotalRecords = totalRecords,
                    TotalPages = totalPages,
                    CurrentPage = page
                },
                DoanhNghiep = doanhNghiep,
                AllTongSo = allTongSoTien,
                AllYHocCoTruyen = allY_HOC_CO_TRUYEN,
                AllTEDuoi6Tuoi = TRE_EM_DUOI_6,
                AllBHYT = BHYT,
                AllVienPhi = VIEN_PHI,
                AllKhongThuDuoc = KHONG_THU_DUOC,
                AllCapCuu = CAP_CUU,
                AllSoNguoiVaoVien = SO_NGUOI_VAO_VIEN,
                AllSoNguoiChuyenVien = SO_NGUOI_CHUYEN_VIEN,
                AllNTSoNguoiBenh = NT_SO_NGUOI_BENH,
                AllNTYHocCoTruyen = NT_YHCT,
                AllNTTEDuoi6Tuoi = NT_TRE_EM_DUOI_6,
                AllNTSoNgay = NT_SO_NGAY,
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

        //private async Task<byte[]> ExportHoatDongKhamBenhExcelAsync(ExportRequest request, ISession session)
        //{
        //    var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
        //    var logoPath = "";

        //    var data = request.Data ?? new List<M0304HoatDongKhamBenh>();
        //    var document = new P0304HoatDongKhamBenhExcelReportTemplate(data, request.FromDate, request.ToDate, request.TenNVDN, doanhNghiepObj, logoPath);

        //    var excelBytes = document.GenerateExcel();
        //    return excelBytes;
        //}

        private async Task<byte[]> ExportHoatDongKhamBenhPdfAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var logoPath ="";

            var data = request.Data ?? new List<M0304HoatDongKhamBenh>();
            var document = new P0304HoatDongKhamBenhReportTemplate(data, request.FromDate, request.ToDate, request.TenNVDN, doanhNghiepObj, logoPath);

            var pdfBytes = document.GeneratePdf();
            return pdfBytes;
        }
    }
}