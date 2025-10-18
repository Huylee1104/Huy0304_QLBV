using C0304.Db.Models;
using M0304.Models.BangKeThu;
using M0304.Models.ThongTinDoanhNghiep;
using M0304NhanVien.Models;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using P0304.PDFDocument;
using QuestPDF.Fluent;
using System.Linq;

namespace S0304BangKeThu.Services
{
    public class S0304BangKeThuService : I0304BangKeThuService
    {
        private readonly M0304Context _context;
        private readonly ILogger<S0304BangKeThuService> _logger;
        private readonly IHttpContextAccessor _httpContextAccessor;
        private readonly I0304ThongTinDoanhNghiep _thongTinDoanhNghiepService;
        private readonly I0304HTTTService _htttService;
        private readonly I0304NhanVienService _nhanVienService;
        private readonly IWebHostEnvironment _env;

        public S0304BangKeThuService(M0304Context context, ILogger<S0304BangKeThuService> logger, IHttpContextAccessor httpContextAccessor,
            I0304ThongTinDoanhNghiep thongTinDoanhNghiepService, I0304HTTTService htttService, I0304NhanVienService nhanVienService, IWebHostEnvironment env)
        {
            _context = context;
            _logger = logger;
            _httpContextAccessor = httpContextAccessor;
            _thongTinDoanhNghiepService = thongTinDoanhNghiepService;
            _htttService = htttService;
            _nhanVienService = nhanVienService;
            _env = env;
        }

        public async Task<M0304BangKeThuResponse> GetBangKeThu(string ngayBatDau, string ngayKetThuc, long idCN, string tenNVDN,  long? idHTTT = null,
            long? idNhanVien = null, int? idLoai = null,  int page = 1, int pageSize = 20)
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
                return new M0304BangKeThuResponse
                {
                    BangKeThu = new M0304PagedResult<M0304BangKeThu>
                    {
                        Success = false,
                        Message = "Khong tim thay doanh nghiep.",
                        Data = null,         // không có dữ liệu
                        TotalRecords = 0,
                        TotalPages = 0,
                        CurrentPage = page
                    },
                    DoanhNghiep = null,
                    TongHoan = 0,
                    TongHuy = 0,
                    TongSoTien = 0,
                    TenNVDN = tenNVDN,
                };
            }
            var allData = await _context.M0304BangKeThus
                .FromSqlRaw("EXEC dbo.[S0304_BangKeThuNgoaiTru] @TuNgay, @DenNgay, @IDCN, @IDHTTT, @IDNhanVien",
                    new SqlParameter("@TuNgay", ngayBatDau),
                    new SqlParameter("@DenNgay", ngayKetThuc),
                    new SqlParameter("@IDCN", idCN),
                    new SqlParameter("@IDHTTT", idHTTT),
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

            // Tính tổng
            var allHoan = allData.Sum(x => x.Hoan);
            var allHuy = allData.Sum(x => x.Huy);
            var allSoTien = allData.Sum(x => x.SoTien);

            return new M0304BangKeThuResponse
            {
                BangKeThu = new M0304PagedResult<M0304BangKeThu>
                {
                    Success = true,
                    Message = message,
                    Data = pagedData,
                    TotalRecords = totalRecords,
                    TotalPages = totalPages,
                    CurrentPage = page
                },
                DoanhNghiep = doanhNghiep,
                TongHoan = allHoan,
                TongHuy = allHuy,
                TongSoTien = allSoTien,
                TenNVDN = tenNVDN,
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

        private async Task<(string tenHTTT, string tenNhanVien, List<M0304NhanVienModel> danhSachNhanVien, List<M0304TongTheoQuyenSo> tongTheoQuyenSo)>
    Get_HTTT_NV(long idHTTT, long idNhanVien, List<M0304BangKeThu> data)
        {
            var allHTTT = await _htttService.GetAllHTTT();
            var tenHTTT = allHTTT.FirstOrDefault(ht => ht.id == idHTTT)?.ten ?? "Tất cả";

            List<M0304NhanVienModel> danhSachNhanVien = null;
            List<M0304TongTheoQuyenSo> tongTheoQuyenSo = null;
            string tenNhanVien = "Tất cả nhân viên";

            var allNhanVien = await _nhanVienService.GetAllNhanVien();

            if (idNhanVien != 0)
            {

                danhSachNhanVien = allNhanVien.Where(nv => nv.ID == idNhanVien).ToList();
            }
            else
            {
                var ids = data.Select(d => d.IDNhanVien).Distinct().ToList();
                danhSachNhanVien = allNhanVien.Where(nv => ids.Contains(nv.ID)).ToList();
            }

            tongTheoQuyenSo = data
                .GroupBy(r => r.QuyenSo) // gom nhóm theo QuyenSo
                .Select(g => new M0304TongTheoQuyenSo
                {
                    QuyenSo = g.Key, // g.Key chính là QuyenSo
                    Seri = g.First().Seri,
                    TongHuy = g.Sum(x => x.Huy ?? 0m),
                    TongHoan = g.Sum(x => x.Hoan ?? 0m),
                    TongSoTien = g.Sum(x => x.SoTien ?? 0m),
                    TongChenhLech = g.Sum(x => (x.SoTien ?? 0m) - ((x.Huy ?? 0m) + (x.Hoan ?? 0m)))
                })
                .ToList();

            return (tenHTTT, tenNhanVien, danhSachNhanVien, tongTheoQuyenSo);
        }

        public async Task<byte[]> ExportBaoCaoGoiKhamPdfAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var HTTT_NV = await Get_HTTT_NV(request.IdHTTT ?? 0, request.IdNhanVien ?? 0, request.Data ?? new List<M0304BangKeThu>());
            var logoPath = Path.Combine(_env.WebRootPath, "dist", "img", "logo.png");

            var data = request.Data ?? new List<M0304BangKeThu>();
            var document = new P0304ReportTemplatePDF(data, request.FromDate, request.ToDate, request.TenNVDN,
                HTTT_NV.tenHTTT, HTTT_NV.danhSachNhanVien, HTTT_NV.tongTheoQuyenSo, doanhNghiepObj, logoPath);

            var pdfBytes = document.GeneratePdf();
            return pdfBytes;
        }
        public async Task<byte[]> ExportBaoCaoGoiKhamExcelAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var HTTT_NV = await Get_HTTT_NV(request.IdHTTT ?? 0, request.IdNhanVien ?? 0, request.Data ?? new List<M0304BangKeThu>());
            var logoPath = Path.Combine(_env.WebRootPath, "dist", "img", "logo.png");

            var data = request.Data ?? new List<M0304BangKeThu>();
            var document = new P0304ExcelReportTemplate(data, request.FromDate, request.ToDate, request.TenNVDN,
                HTTT_NV.tenHTTT, HTTT_NV.danhSachNhanVien, HTTT_NV.tongTheoQuyenSo, doanhNghiepObj, logoPath);

            var excelBytes = document.GenerateExcel();
            return excelBytes;
        }
    }
}