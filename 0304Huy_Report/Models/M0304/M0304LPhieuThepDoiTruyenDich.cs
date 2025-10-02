using M0304.Models.ThongTinDoanhNghiep;
using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations;

namespace M0304L.Models.PhieuTheoDoiTruyenDich
{
    [Keyless]
    public class ThongTinBNModel
    {
        public string? MaVaoVien { get; set; }
        public string? TenBenhNhan { get; set; }
        public string? TenKhoa { get; set; }
        public string? TenPhong { get; set; }
        public string? TenGiuong { get; set; }
        public DateTime? NgaySinh { get; set; }
        public string? GioiTinh { get; set; }
        public string? ChanDoan { get; set; } 
    }

    [Keyless]
    public class TruyenDich
    {
        public DateTime? NgayThang { get; set; }
        public string? TenDichTruyen { get; set; }
        public int? SoLuong { get; set; }
        public string? SoLo { get; set; } 
        public string? TocDo { get; set; } 
        public DateTime? BatDau { get; set; }
        public DateTime? KetThuc { get; set; }
        public string? BSChiDinh { get; set; }
        public string? NguoiThucHien { get; set; }
    }

    public class HoSoBenhAnModel
    {
        public ThongTinBNModel ThongTinBN { get; set; }
        public List<TruyenDich> TruyenDich { get; set; }
    }

    public class M0304LPagedResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public T Data { get; set; }
        public int TotalRecords { get; set; }
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class M0304LPhieuTheoDoiTruyenDichResponse
    {
        public M0304LPagedResult<HoSoBenhAnModel> PhieuTheoDoiTruyenDich { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }

    public class ExportRequest
    {
        public HoSoBenhAnModel Data { get; set; }
        public long? IdBenhNhan { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }
}