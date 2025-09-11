using C0304.Db.Models;
using DocumentFormat.OpenXml.Office2010.Excel;
using M0304.Models.ThongTinDoanhNghiep;
using M0304F.Models.BKTinhHinhTraDuocNCC;
using M0304NhanVien.Models;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using P0304F.PDFDocument;
using QuestPDF.Fluent;

namespace S0304FBKTinhHinhTraDuocNCC.Services
{
    public class S0304FBKTinhHinhTraDuocNCCService : I0304FBKTinhHinhTraDuocNCCService
    {
        private readonly M0304Context _context;
        private readonly ILogger<S0304FBKTinhHinhTraDuocNCCService> _logger;
        private readonly IHttpContextAccessor _httpContextAccessor;
        private readonly I0304ThongTinDoanhNghiep _thongTinDoanhNghiepService;
        private readonly I0304KhoHangService _khoHangService;
        private readonly IWebHostEnvironment _env;

        public S0304FBKTinhHinhTraDuocNCCService(M0304Context context, ILogger<S0304FBKTinhHinhTraDuocNCCService> logger, IHttpContextAccessor httpContextAccessor,
            I0304ThongTinDoanhNghiep thongTinDoanhNghiepService, I0304KhoHangService khoHangService, IWebHostEnvironment env)
        {
            _context = context;
            _logger = logger;
            _httpContextAccessor = httpContextAccessor;
            _thongTinDoanhNghiepService = thongTinDoanhNghiepService;
            _khoHangService = khoHangService;
            _env = env;
        }

        public async Task<M0304FBKTinhHinhTraDuocNCCResponse> GetBKTinhHinhTraDuocNCC(string ngayBatDau, string ngayKetThuc, long idCN,
            long? idKhoHang = null, int page = 1, int pageSize = 20)
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
                return new M0304FBKTinhHinhTraDuocNCCResponse
                {
                    BKTinhHinhTraDuocNCC = new M0304FPagedResult<M0304FBKTinhHinhTraDuocNCC>
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
            var allData = await _context.M0304FBKTinhHinhTraDuocNCCs
                .FromSqlRaw("EXEC dbo.S0304_BKTinhHinhTraDuocNCC @TuNgay, @DenNgay, @IDCN, @IDKhoHang",
                    new SqlParameter("@TuNgay", ngayBatDau),
                    new SqlParameter("@DenNgay", ngayKetThuc),
                    new SqlParameter("@IDCN", idCN),
                    new SqlParameter("@IDKhoHang", idKhoHang))
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

            return new M0304FBKTinhHinhTraDuocNCCResponse
            {
                BKTinhHinhTraDuocNCC = new M0304FPagedResult<M0304FBKTinhHinhTraDuocNCC>
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

        private async Task<(List<CongTyDto> dsCongTy, string tenKho)> GetCongTy_Kho(List<M0304FBKTinhHinhTraDuocNCC> allData, long idKhoHang)
        {
            var dsCongTy = allData
                .Where(x => !string.IsNullOrEmpty(x.CongTy) && x.IDCongTy != null)
                .GroupBy(x => x.IDCongTy.Value)
                .Select(g => new CongTyDto
                {
                    ID = g.Key,
                    Ten = g.First().CongTy!
                })
                .OrderBy(x => x.Ten)
                .ToList();

            var khoHang = await _context.M0304KhoHangs
                .AsNoTracking()
                .FirstOrDefaultAsync(x => x.ID == idKhoHang);

            var tenKho = khoHang?.TenKhoHang;
            return (dsCongTy, tenKho);
        }

        public async Task<byte[]> ExportBKTinhHinhTraDuocNCCPdfAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var allData = request.Data ?? new List<M0304FBKTinhHinhTraDuocNCC>();
            var info = await GetCongTy_Kho(allData, request.IdKhoHang ?? 0);
            var logoPath = Path.Combine(_env.WebRootPath, "dist", "img", "logo.png");
            
            var data = request.Data ?? new List<M0304FBKTinhHinhTraDuocNCC>();
            var document = new P0304FReportTemplatePDF(data, doanhNghiepObj, info.dsCongTy, info.tenKho, request.FromDate, request.ToDate, logoPath);

            var pdfBytes = document.GeneratePdf();
            return pdfBytes;
        }
        public async Task<byte[]> ExportBKTinhHinhTraDuocNCCExcelAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var allData = request.Data ?? new List<M0304FBKTinhHinhTraDuocNCC>();
            var info = await GetCongTy_Kho(allData, request.IdKhoHang ?? 0);
            var logoPath = Path.Combine(_env.WebRootPath, "dist", "img", "logo.png");

            var data = request.Data ?? new List<M0304FBKTinhHinhTraDuocNCC>();
            var document = new P0304FExcelReportTemplate(data, doanhNghiepObj, info.dsCongTy, info.tenKho, request.FromDate, request.ToDate, logoPath);

            var excelBytes = document.GenerateExcel();
            return excelBytes;
        }
    }
}