using C0304.Db.Models;
using M0304.Models.ThongTinDoanhNghiep;
using M0304E.Models.BKBienLaiHoanUng;
using M0304NhanVien.Models;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using P0304E.PDFDocument;
using QuestPDF.Fluent;

namespace S0304EBKBienLaiHoanUng.Services
{
    public class S0304EBKBienLaiHoanUngService : I0304EBKBienLaiHoanUngService
    {
        private readonly M0304Context _context;
        private readonly ILogger<S0304EBKBienLaiHoanUngService> _logger;
        private readonly IHttpContextAccessor _httpContextAccessor;
        private readonly I0304ThongTinDoanhNghiep _thongTinDoanhNghiepService;
        private readonly I0304NhanVienService _nhanVienService;
        private readonly IWebHostEnvironment _env;

        public S0304EBKBienLaiHoanUngService(M0304Context context, ILogger<S0304EBKBienLaiHoanUngService> logger, IHttpContextAccessor httpContextAccessor,
            I0304ThongTinDoanhNghiep thongTinDoanhNghiepService, I0304NhanVienService nhanVienService, IWebHostEnvironment env)
        {
            _context = context;
            _logger = logger;
            _httpContextAccessor = httpContextAccessor;
            _thongTinDoanhNghiepService = thongTinDoanhNghiepService;
            _nhanVienService = nhanVienService;
            _env = env;
        }

        public async Task<M0304EBKBienLaiHoanUngResponse> GetBKBienLaiHoanUng(string ngayBatDau, string ngayKetThuc, long idCN, long? idNhanVien = null, int page = 1, int pageSize = 20)
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
                return new M0304EBKBienLaiHoanUngResponse
                {
                    BangKeBienLaiHoanUng = new M0304EPagedResult<M0304EBKBienLaiHoanUng>
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
            var allData = await _context.M0304EBKBienLaiHoanUngs
                .FromSqlRaw("EXEC dbo.S0304_BangKeBienLaiHoanUng @TuNgay, @DenNgay, @IDCN, @IDNhanVien",
                    new SqlParameter("@TuNgay", ngayBatDau),
                    new SqlParameter("@DenNgay", ngayKetThuc),
                    new SqlParameter("@IDCN", idCN),
                    new SqlParameter("@IDNhanVien", idNhanVien))
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

            return new M0304EBKBienLaiHoanUngResponse
            {
                BangKeBienLaiHoanUng = new M0304EPagedResult<M0304EBKBienLaiHoanUng>
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

        private async Task<(List<M0304NhanVienModel> danhSachNhanVien, List<M0304TongTheoNhanVien> tongTheoNhanVien)>
            Get_NV(long idNhanVien, List<M0304EBKBienLaiHoanUng> data)
        {

            List<M0304NhanVienModel> danhSachNhanVien = null;
            List<M0304TongTheoNhanVien> tongTheoNhanVien = null;
            string tenNhanVien = "Tất cả nhân viên";
            var allNhanVien = await _nhanVienService.GetAllNhanVien();

            var ids = data.Select(d => d.IDNhanVien).Distinct().ToList();
            danhSachNhanVien = allNhanVien.Where(nv => ids.Contains(nv.Id)).ToList();

            // 2. Tính tổng theo nhân viên
            tongTheoNhanVien = data
            .GroupBy(r => r.IDNhanVien)
            .Select(g => new M0304TongTheoNhanVien
            {
                IDNhanVien = g.Key ?? 0,
                TongHuy = g.Sum(x => x.Huy ?? 0m),
                TongHoan = g.Sum(x => x.HoanTra ?? 0m),
                TongSoTien = g.Sum(x => x.GiaTriHoanUng ?? 0m),
            })
            .ToList();

            return (danhSachNhanVien, tongTheoNhanVien);
        }

        public async Task<byte[]> ExportBKBienLaiHoanUngPdfAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var NV = await Get_NV(request.IdNhanVien ?? 0, request.Data ?? new List<M0304EBKBienLaiHoanUng>());
            var logoPath = Path.Combine(_env.WebRootPath, "dist", "img", "logo.png");

            var data = request.Data ?? new List<M0304EBKBienLaiHoanUng>();
            var document = new P0304EReportTemplatePDF(data, request.FromDate, request.ToDate, NV.danhSachNhanVien, NV.tongTheoNhanVien, doanhNghiepObj, logoPath);

            var pdfBytes = document.GeneratePdf();
            return pdfBytes;
        }
        public async Task<byte[]> ExportBKBienLaiHoanUngExcelAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var NV = await Get_NV(request.IdNhanVien ?? 0, request.Data ?? new List<M0304EBKBienLaiHoanUng>());
            var logoPath = Path.Combine(_env.WebRootPath, "dist", "img", "logo.png");

            var data = request.Data ?? new List<M0304EBKBienLaiHoanUng>();
            var document = new P0304EExcelReportTemplate(data, request.FromDate, request.ToDate, NV.danhSachNhanVien, NV.tongTheoNhanVien, doanhNghiepObj, logoPath);

            var excelBytes = document.GenerateExcel();
            return excelBytes;
        }
    }
}