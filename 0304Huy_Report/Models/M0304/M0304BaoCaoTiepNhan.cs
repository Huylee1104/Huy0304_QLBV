using M0304.Models.ThongTinDoanhNghiep;
using M0304NhanVien.Models;
using System.ComponentModel.DataAnnotations;

namespace M0304.Models.BaoCaoTiepNhan
{
    public class M0304BaoCaoTiepNhan
    {
        public string? TenPhongBenh { get; set; }
        public int? SoLuotTiepNhan { get; set; }
        public int? SoLuongNam { get; set; }
        public int? SoLuongNu { get; set; }
        public int? CoBHYT { get; set; }
        public int? KhongBHYT { get; set; }
        public string? TenKhoa { get; set; }
        public long? IdKhoa { get; set; }
    }

    public class M0304BaoCaoTiepNhanPagedResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public IEnumerable<T> Data { get; set; } 
        public int TotalRecords { get; set; } 
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class M0304BaoCaoTiepNhanResponse
    {
        public M0304BaoCaoTiepNhanPagedResult<M0304BaoCaoTiepNhan> BaoCaoTiepNhan { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }

        public long? TongLuotTiepNhan { get; set; }
        public long? TongNam { get; set; }
        public long? TongNu { get; set; }
        public long? TongCoBHYT { get; set; }
        public long? TongKHongBHYT { get; set; }
    }

    public class ExportRequest
    {
        public List<M0304BaoCaoTiepNhan> Data { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public long? idKhoa { get; set; }
        public long? idPhongBuong { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }
}