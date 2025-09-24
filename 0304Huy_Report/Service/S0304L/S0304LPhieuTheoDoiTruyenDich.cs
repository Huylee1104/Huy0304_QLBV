using C0304.Db.Models;
using M0304.Models.ThongTinDoanhNghiep;
using M0304L.Models.PhieuTheoDoiTruyenDich;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using P0304L.PDFDocument;
using QuestPDF.Fluent;
using System.Data;
using System.Data.Common;

namespace S0304LPhieuTheoDoiTruyenDich.Services
{
    public class S0304LPhieuTheoDoiTruyenDichService : I0304LPhieuTheoDoiTruyenDichService
    {
        private readonly M0304Context _context;
        private readonly ILogger<S0304LPhieuTheoDoiTruyenDichService> _logger;
        private readonly IHttpContextAccessor _httpContextAccessor;
        private readonly I0304ThongTinDoanhNghiep _thongTinDoanhNghiepService;
        private readonly IWebHostEnvironment _env;

        public S0304LPhieuTheoDoiTruyenDichService(M0304Context context, ILogger<S0304LPhieuTheoDoiTruyenDichService> logger, IHttpContextAccessor httpContextAccessor,
            I0304ThongTinDoanhNghiep thongTinDoanhNghiepService, IWebHostEnvironment env)
        {
            _context = context;
            _logger = logger;
            _httpContextAccessor = httpContextAccessor;
            _thongTinDoanhNghiepService = thongTinDoanhNghiepService;
            _env = env;
        }

        public async Task<M0304LPhieuTheoDoiTruyenDichResponse> GetPhieuTheoDoiTruyenDich(long idCN, long? idVaoVien, int page = 1, int pageSize = 20)
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
                return new M0304LPhieuTheoDoiTruyenDichResponse
                {
                    PhieuTheoDoiTruyenDich = new M0304LPagedResult<HoSoBenhAnModel>
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
            var hoSo = await GetHoSoBenhAnAsync(idVaoVien ?? 0, idCN);

            // Phân trang chỉ áp dụng cho SinhHieus
            var totalRecords = hoSo.TruyenDich?.Count ?? 0;
            var totalPages = (int)Math.Ceiling((double)totalRecords / pageSize);
            var pagedTruyenDichs = hoSo.TruyenDich?
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .ToList();

            hoSo.TruyenDich = pagedTruyenDichs; // gán lại list sau phân trang

            string message = totalRecords > 0
                ? "Tìm thấy kết quả"
                : "Không tìm thấy kết quả nào";

            var sessionData = new { Data = hoSo };
            session?.SetString("FilteredData", JsonConvert.SerializeObject(sessionData));

            return new M0304LPhieuTheoDoiTruyenDichResponse
            {
                PhieuTheoDoiTruyenDich = new M0304LPagedResult<HoSoBenhAnModel>
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
        public async Task<byte[]> ExportGetPhieuTheoDoiTruyenDichPdfAsync(ExportRequest request, ISession session)
        {
            var doanhNghiepObj = GetDoanhNghiepFromRequestOrSession(request, session);
            var logoPath = Path.Combine(_env.WebRootPath, "dist", "img", "logo.png");

            var data = request.Data ?? new HoSoBenhAnModel();
            var document = new P0304LReportTemplatePDF(data, doanhNghiepObj, logoPath);

            var pdfBytes = document.GeneratePdf();
            return pdfBytes;
        }

        private async Task<HoSoBenhAnModel> GetHoSoBenhAnAsync(long idVaoVien, long idCN)
        {
            var result = new HoSoBenhAnModel();

            using (var conn = _context.Database.GetDbConnection())
            {
                await conn.OpenAsync();

                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "dbo.S0304_PhieuTheoDoiTruyenDich";
                    cmd.CommandType = CommandType.StoredProcedure;

                    var p1 = cmd.CreateParameter();
                    p1.ParameterName = "@IdVaoVien";
                    p1.Value = idVaoVien;
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
                            result.ThongTinBN = new ThongTinBNModel()
                            {
                                MaVaoVien = GetSafeString(reader, "MaVaoVien"),
                                TenBenhNhan = GetSafeString(reader, "TenBenhNhan"),
                                TenKhoa = GetSafeString(reader, "TenKhoa"),
                                TenPhong = GetSafeString(reader, "TenPhong"),
                                TenGiuong = GetSafeString(reader, "TenGiuong"),
                                NgaySinh = GetSafeDateTime(reader, "NgaySinh"),
                                GioiTinh = GetSafeString(reader, "GioiTinh"),
                                ChanDoan = GetSafeString(reader, "ChanDoan")
                            };
                        }

                        // Sang result set 2
                        if (await reader.NextResultAsync())
                        {
                            result.TruyenDich = new List<TruyenDich>();
                            while (await reader.ReadAsync())
                            {
                                result.TruyenDich.Add(new TruyenDich
                                {
                                    NgayThang = reader.IsDBNull(reader.GetOrdinal("NgayThang")) ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("NgayThang")),
                                    TenDichTruyen = reader.IsDBNull(reader.GetOrdinal("TenDichTruyen")) ? null : reader["TenDichTruyen"].ToString(),
                                    SoLuong = reader.IsDBNull(reader.GetOrdinal("SoLuong")) ? (int?)null : reader.GetInt32(reader.GetOrdinal("SoLuong")),
                                    SoLo = reader.IsDBNull(reader.GetOrdinal("SoLo")) ? null : reader["SoLo"].ToString(),
                                    BatDau = reader.IsDBNull(reader.GetOrdinal("BatDau")) ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("BatDau")),
                                    KetThuc = reader.IsDBNull(reader.GetOrdinal("KetThuc")) ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("KetThuc")),
                                    BSChiDinh = reader.IsDBNull(reader.GetOrdinal("BSChiDinh")) ? null : reader["BSChiDinh"].ToString(),
                                    NguoiThucHien = reader.IsDBNull(reader.GetOrdinal("NguoiThucHien")) ? null : reader["NguoiThucHien"].ToString(),

                                });
                            }
                        }
                    }
                }
            }

            return result;
        }
        private static string GetSafeString(DbDataReader reader, string columnName)
        {
            int ordinal = reader.GetOrdinal(columnName);
            return reader.IsDBNull(ordinal) ? null : reader.GetString(ordinal);
        }

        private static DateTime? GetSafeDateTime(DbDataReader reader, string columnName)
        {
            int ordinal = reader.GetOrdinal(columnName);
            return reader.IsDBNull(ordinal) ? (DateTime?)null : reader.GetDateTime(ordinal);
        }
    }
}