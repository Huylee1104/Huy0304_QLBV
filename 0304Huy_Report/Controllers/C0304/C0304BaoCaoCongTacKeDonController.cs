using _0304Huy_Report.Models;
using C0304.Db.Models;
using M0304.Models.BaoCaoCongTacKeDon;
using M0304.Models.ThongTinDoanhNghiep;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using System.Diagnostics;

namespace C0304BaoCaoCongTacKeDon.Controllers
{
    [Route("bao_cao_cong_tac_ke_don")]
    public class C0304BaoCaoCongTacKeDonController : Controller
    {
        //private string _maChucNang = "/bao_cao_cong_tac_ke_don";
        //private IMemoryCachingServices _memoryCache;
        private readonly ILogger<C0304BaoCaoCongTacKeDonController> _logger;
        private readonly M0304Context _context;
        private readonly I0304ThongTinDoanhNghiep _thongTinDoanhNghiepService;
        private readonly IWebHostEnvironment _env;

        public C0304BaoCaoCongTacKeDonController(ILogger<C0304BaoCaoCongTacKeDonController> logger,
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

            return View("~/Views/V0304/V0304BaoCaoCongTacKeDon/Index.cshtml");
        }

        [HttpPost("filter")]
        public async Task<IActionResult> FilterByDay(string tuNgay, string denNgay, long IdChiNhanh, long? idKhoHang, long? idNhomHang = null,
            long ? idHangHoa = null, int page = 1, int pageSize = 20)
        {
            try
            {
                var result = await GetBaoCaoCongTacKeDon(tuNgay, denNgay, IdChiNhanh, idKhoHang, idNhomHang, idHangHoa, page, pageSize);

                if (!result.BaoCaoCongTacKeDon.Success)
                {
                    _logger.LogWarning("Service trả về lỗi: {Message}", result.BaoCaoCongTacKeDon.Message);
                    return Json(new { success = false, message = result.BaoCaoCongTacKeDon.Message });
                }

                return Json(new
                {
                    success = true,
                    message = result.BaoCaoCongTacKeDon.Message,
                    data = result.BaoCaoCongTacKeDon.Data,
                    totalRecords = result.BaoCaoCongTacKeDon.TotalRecords,
                    totalPages = result.BaoCaoCongTacKeDon.TotalPages,
                    currentPage = result.BaoCaoCongTacKeDon.CurrentPage,
                    doanhNghiep = result.DoanhNghiep
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi xảy ra trong FilterByDay");
                return StatusCode(500, "Lỗi server: " + ex.Message);
            }

        }

        [HttpPost("export/excel")]
        public async Task<IActionResult> ExportToExcel([FromBody] ExportRequest request)
        {
            var excelBytes = await ExportBaoCaoGoiKhamExcelAsync(request, HttpContext.Session);

            string fileName = $"BaoCaoCongTacKeDonNgoaiTru_{request.FromDate ?? "all"}_den_{request.ToDate ?? "now"}.xlsx";
            return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }


        private async Task<M0304BaoCaoCongTacKeDonResponse> GetBaoCaoCongTacKeDon(string ngayBatDau, string ngayKetThuc, long idCN, long? idKhoHang,
            long? idNhomHang = null, long? idHangHoa = null, int page = 1, int pageSize = 20)
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
                return new M0304BaoCaoCongTacKeDonResponse
                {
                    BaoCaoCongTacKeDon = new M0304BaoCaoCongTacKeDonPagedResult<M0304BaoCaoCongTacKeDon>
                    {
                        Success = false,
                        Message = "Khong tim thay doanh nghiep.",
                        Data = null,
                        TotalRecords = 0,
                        TotalPages = 0,
                        CurrentPage = page
                    },
                    DoanhNghiep = null
                };
            }
            var allData = await _context.BaoCaoCongTacKeDons
                .FromSqlRaw("EXEC dbo.[S0304_BaoCaoCongTacKeDon] @TuNgay, @DenNgay, @IDCN, @IDKho, @IDNhomHang, @IDHangHoa",
                    new SqlParameter("@TuNgay", ngayBatDau),
                    new SqlParameter("@DenNgay", ngayKetThuc),
                    new SqlParameter("@IDCN", idCN),
                    new SqlParameter("@IDKho", idKhoHang),
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

            return new M0304BaoCaoCongTacKeDonResponse
            {
                BaoCaoCongTacKeDon = new M0304BaoCaoCongTacKeDonPagedResult<M0304BaoCaoCongTacKeDon>
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

        private async Task<byte[]> ExportBaoCaoGoiKhamExcelAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var logoPath = Path.Combine(_env.WebRootPath, "dist", "img", "logo.png");

            var data = request.Data ?? new List<M0304BaoCaoCongTacKeDon>();
            var document = new P0304BaoCaoCongTacKeDonExcelReportTemplate(data, request.FromDate, request.ToDate, doanhNghiepObj, logoPath);

            var excelBytes = document.GenerateExcel();
            return excelBytes;
        }

    }
}