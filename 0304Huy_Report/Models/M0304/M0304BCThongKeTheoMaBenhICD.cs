using M0304.Models.ThongTinDoanhNghiep;
using M0304NhanVien.Models;
using System.ComponentModel.DataAnnotations;

namespace M0304.Models.BCThongKeTheoMaBenhICD
{
    public class M0304BCThongKeTheoMaBenhICD
    {
        public string? TenICD { get; set; }
        public string? TenBenh { get; set; }
        public int? SoLuotTiepNhan { get; set; }
        public int? SoLuongNam { get; set; }
        public int? SoLuongNu { get; set; }
        public int? CoBHYT { get; set; }
        public int? KhongBHYT { get; set; }
    }

    public class M0304BCThongKeTheoMaBenhICDPagedResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public IEnumerable<T> Data { get; set; } 
        public int TotalRecords { get; set; } 
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class M0304BCThongKeTheoMaBenhICDResponse
    {
        public M0304BCThongKeTheoMaBenhICDPagedResult<M0304BCThongKeTheoMaBenhICD> BCThongKeTheoMaBenhICD { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }

        public long? TongLuotTiepNhan { get; set; }
        public long? TongNam { get; set; }
        public long? TongNu { get; set; }
        public long? TongCoBHYT { get; set; }
        public long? TongKHongBHYT { get; set; }
    }

    public class ExportRequest
    {
        public List<M0304BCThongKeTheoMaBenhICD> Data { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public string? TenNVDN { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }
}