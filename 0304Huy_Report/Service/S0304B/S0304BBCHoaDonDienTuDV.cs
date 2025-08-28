using C0304.Db.Models;
using M0304.Models.ThongTinDOanhNghiep;
using M0304B.Models.BCHoaDonDienTuDV;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using P0304.PDFDocument;
using QuestPDF.Fluent;

namespace S0304BBCHoaDonDienTuDV.Services
{
    public class S0304BBCHoaDonDienTuService : I0304BBCHoaDonDienTuDVService
    {
        private readonly M0304Context _context;
        private readonly ILogger<S0304BBCHoaDonDienTuService> _logger;
        private readonly IHttpContextAccessor _httpContextAccessor;
        private readonly I0304ThongTinDoanhNghiep _thongTinDoanhNghiepService;
        private readonly IWebHostEnvironment _env;

        public S0304BBCHoaDonDienTuService(M0304Context context, ILogger<S0304BBCHoaDonDienTuService> logger, IHttpContextAccessor httpContextAccessor,
            I0304ThongTinDoanhNghiep thongTinDoanhNghiepService, IWebHostEnvironment env)
        {
            _context = context;
            _logger = logger;
            _httpContextAccessor = httpContextAccessor;
            _thongTinDoanhNghiepService = thongTinDoanhNghiepService;
            _env = env;
        }

        public async Task<M0304BBCHoaDonDienTuDVResponse> GetBCHoaDonDIenTuDV(string ngayBatDau, string ngayKetThuc, long idCN, int page = 1, int pageSize = 0)
        {
            var doanhNghiep = await _thongTinDoanhNghiepService.GetThongTinDoanhNghiep(idCN);

            var session = _httpContextAccessor.HttpContext?.Session;

            if (doanhNghiep != null)
            {
                // Lưu thông tin doanh nghiệp vào session
                session?.SetString("DoanhNghiepInfo", JsonConvert.SerializeObject(doanhNghiep));
                _logger.LogInformation("Doanh Nghiep Info: {@DoanhNghiep}", doanhNghiep);
            }
            else
            {
                _logger.LogWarning("No doanh nghiep found for ChiNhanh ID: {IdChiNhanh}", idCN);
                return new M0304BBCHoaDonDienTuDVResponse
                {
                    BCHoaDonDienTuDV = new M0304BPagedResult<M0304BBCHoaDonDienTuDV>
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
            var allData = await _context.M0304BBCHoaDonDienTuDVs
                .FromSqlRaw("EXEC dbo.S0304_BCHoaDonDienTuDV @TuNgay, @DenNgay, @IDCN",
                    new SqlParameter("@TuNgay", ngayBatDau),
                    new SqlParameter("@DenNgay", ngayKetThuc),
                    new SqlParameter("@IDCN", idCN))
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

            return new M0304BBCHoaDonDienTuDVResponse
            {
                BCHoaDonDienTuDV = new M0304BPagedResult<M0304BBCHoaDonDienTuDV>
                {
                    Success = pagedData.Any(),
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
        public async Task<byte[]> ExportBaoCaoGoiKhamPdfAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var logoPath = Path.Combine(_env.WebRootPath, "dist", "img", "logo.png");

            var data = request.Data ?? new List<M0304BBCHoaDonDienTuDV>();
            var document = new P0304BReportTemplatePDF(data, request.FromDate, request.ToDate, doanhNghiepObj, logoPath);

            var pdfBytes = document.GeneratePdf();
            return pdfBytes;
        }
        public async Task<byte[]> ExportBaoCaoGoiKhamExcelAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var logoPath = Path.Combine(_env.WebRootPath, "dist", "img", "logo.png");

            var data = request.Data ?? new List<M0304BBCHoaDonDienTuDV>();
            var document = new P0304BExcelReportTemplate(data, request.FromDate, request.ToDate, doanhNghiepObj, logoPath);

            var excelBytes = document.GenerateExcel();
            return excelBytes;
        }
    }
}