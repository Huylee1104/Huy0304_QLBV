using C0304.Db.Models;
using M0304.Models.ThongTinDoanhNghiep;
using M0304I.Models.PhieuTheoDoiChucNangSong;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using P0304I.PDFDocument;
using QuestPDF.Fluent;
using System.Data;

namespace S0304CPhieuTheoDoiChucNangSong.Services
{
    public class S0304ITheoDoiChucNangSongService : I0304CPhieuTheoDoiChucNangSongService
    {
        private readonly M0304Context _context;
        private readonly ILogger<S0304ITheoDoiChucNangSongService> _logger;
        private readonly IHttpContextAccessor _httpContextAccessor;
        private readonly I0304ThongTinDoanhNghiep _thongTinDoanhNghiepService;
        private readonly IWebHostEnvironment _env;

        public S0304ITheoDoiChucNangSongService(M0304Context context, ILogger<S0304ITheoDoiChucNangSongService> logger, IHttpContextAccessor httpContextAccessor,
            I0304ThongTinDoanhNghiep thongTinDoanhNghiepService, IWebHostEnvironment env)
        {
            _context = context;
            _logger = logger;
            _httpContextAccessor = httpContextAccessor;
            _thongTinDoanhNghiepService = thongTinDoanhNghiepService;
            _env = env;
        }

        public async Task<M0304IPhieuTheoDoiChucNangSongResponse> GetPhieuTheoDoiChucNangSong(long idCN, long? idBenhNhan, int page = 1, int pageSize = 20)
        {
            var doanhNghiep = await _thongTinDoanhNghiepService.GetThongTinDoanhNghiep(idCN);

            var session = _httpContextAccessor.HttpContext?.Session;

            if (doanhNghiep != null)
            {
                session?.SetString("DoanhNghiepInfo", JsonConvert.SerializeObject(doanhNghiep));
                _logger.LogInformation("Doanh Nghiep Info: {@DoanhNghiep}", doanhNghiep);
            }
            else
            {
                _logger.LogWarning("No doanh nghiep found for ChiNhanh ID: {IdChiNhanh}", idCN);
                return new M0304IPhieuTheoDoiChucNangSongResponse
                {
                    PhieuTheoDoiChucNangSong = new M0304IPagedResult<HoSoBenhAnModel>
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

            // Gọi service bạn đã viết
            var hoSo = await GetHoSoBenhAnAsync(idBenhNhan ?? 0, idCN);

            // Phân trang chỉ áp dụng cho SinhHieus
            var totalRecords = hoSo.SinhHieus?.Count ?? 0;
            var totalPages = (int)Math.Ceiling((double)totalRecords / pageSize);
            var pagedSinhHieus = hoSo.SinhHieus?
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .ToList();

            hoSo.SinhHieus = pagedSinhHieus; // gán lại list sau phân trang

            string message = totalRecords > 0
                ? "Tìm thấy kết quả"
                : "Không tìm thấy kết quả nào";

            var sessionData = new { Data = hoSo };
            session?.SetString("FilteredData", JsonConvert.SerializeObject(sessionData));

            return new M0304IPhieuTheoDoiChucNangSongResponse
            {
                PhieuTheoDoiChucNangSong = new M0304IPagedResult<HoSoBenhAnModel>
                {
                    Success = true,
                    Message = message,
                    Data = hoSo, // bọc vào list để giữ kiểu cũ
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
        public async Task<byte[]> ExportGetPhieuTheoDoiChucNangSongPdfAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var logoPath = Path.Combine(_env.WebRootPath, "dist", "img", "logo.png");

            var data = request.Data ?? new List<HoSoBenhAnModel>();
            var document = new P0304IReportTemplatePDF(data, doanhNghiepObj, logoPath);

            var pdfBytes = document.GeneratePdf();
            return pdfBytes;
        }

        private async Task<HoSoBenhAnModel> GetHoSoBenhAnAsync(long idBenhNhan, long idCN)
        {
            var result = new HoSoBenhAnModel();

            using (var conn = _context.Database.GetDbConnection())
            {
                await conn.OpenAsync();

                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "dbo.S0304_PhieuTheoDoiChucNangSong";
                    cmd.CommandType = CommandType.StoredProcedure;

                    var p1 = cmd.CreateParameter();
                    p1.ParameterName = "@IdBenhNhan";
                    p1.Value = idBenhNhan;
                    cmd.Parameters.Add(p1);

                    var p2 = cmd.CreateParameter();
                    p2.ParameterName = "@IDCN";
                    p2.Value = idCN;
                    cmd.Parameters.Add(p2);

                    using (var reader = await cmd.ExecuteReaderAsync())
                    {
                        // Lấy result set 1
                        if (await reader.ReadAsync())
                        {
                            result.ThongTinBenhNhan = new BenhNhanThongTinModel
                            {
                                MaVaoVien = reader.GetString(reader.GetOrdinal("MaVaoVien")),
                                TenBenhNhan = reader.GetString(reader.GetOrdinal("TenBenhNhan")),
                                Tuoi = reader.GetInt32(reader.GetOrdinal("Tuoi")),
                                GioiTinh = reader.GetString(reader.GetOrdinal("GioiTinh")),
                                ChanDoan = reader.GetString(reader.GetOrdinal("ChanDoan"))
                            };
                        }

                        // Sang result set 2
                        if (await reader.NextResultAsync())
                        {
                            result.SinhHieus = new List<SinhHieuModel>();
                            while (await reader.ReadAsync())
                            {
                                result.SinhHieus.Add(new SinhHieuModel
                                {
                                    NgayKhaoSat = reader.GetDateTime(reader.GetOrdinal("NgayKhaoSat")),
                                    Mach = reader["Mach"]?.ToString(),
                                    NhietDo = reader["NhietDo"]?.ToString(),
                                    HuyetAp = reader["HuyetAp"]?.ToString(),
                                    CanNang = reader["CanNang"]?.ToString(),
                                    NhipTho = reader["NhipTho"]?.ToString()
                                });
                            }
                        }
                    }
                }
            }

            return result;
        }
    }
}