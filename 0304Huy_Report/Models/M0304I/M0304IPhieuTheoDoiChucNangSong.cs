using M0304.Models.ThongTinDoanhNghiep;
using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations;

namespace M0304I.Models.PhieuTheoDoiChucNangSong
{
    [Keyless]
    public class BenhNhanThongTinModel
    {
        public string? MaVaoVien { get; set; }
        public string? TenBenhNhan { get; set; }
        public int? Tuoi { get; set; } 
        public string? GioiTinh { get; set; }
        public string? ChanDoan { get; set; } 
    }

    [Keyless]
    public class SinhHieuModel
    {
        public DateTime? NgayKhaoSat { get; set; }
        public string? Mach { get; set; }  
        public string? NhietDo { get; set; }
        public string? HuyetAp { get; set; } 
        public string? CanNang { get; set; } 
        public string? NhipTho { get; set; }
    }

    public class HoSoBenhAnModel
    {
        public BenhNhanThongTinModel ThongTinBenhNhan { get; set; }
        public List<SinhHieuModel> SinhHieus { get; set; }
    }

    public class M0304IPagedResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public T Data { get; set; }
        public int TotalRecords { get; set; }
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class M0304IPhieuTheoDoiChucNangSongResponse
    {
        public M0304IPagedResult<HoSoBenhAnModel> PhieuTheoDoiChucNangSong { get; set; }   // danh sách bảng kê thu phân trang
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }    // thông tin doanh nghiệp
    }

    public class ExportRequest
    {
        public HoSoBenhAnModel Data { get; set; }
        public long? IdBenhNhan { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }
}