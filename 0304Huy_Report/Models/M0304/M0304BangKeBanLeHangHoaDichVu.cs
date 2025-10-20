using M0304.Models.ThongTinDoanhNghiep;
using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations;

namespace M0304.Models.BangKeBanLeHangHoaDichVu
{

    public class M0304BangKeBanLeHangHoaDichVu
    {
        public string? TenKhoHang { get; set; }
        public string? MauSo { get; set; }
        public string? MaSo { get; set; }
        public string? TenNhanVien { get; set; }
        public string? TenHangHoa { get; set; }
        public string? DVT { get; set; }
        public double? SoLuong { get; set; }
        public double? DonGiaBan { get; set; }
        public double? ThanhTien { get; set; }
    }

    public class M0304BangKeBanLeHangHoaDichVuPagedResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public IEnumerable<T> Data { get; set; }
        public int TotalRecords { get; set; }
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class M0304BangKeBanLeHangHoaDichVuResponse
    {
        public M0304BangKeBanLeHangHoaDichVuPagedResult<M0304BangKeBanLeHangHoaDichVu> BangKeBanLeHangHoaDichVu { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
        public double TongCong { get; set; }
    }

    public class ExportRequest
    {
        public List<M0304BangKeBanLeHangHoaDichVu> Data { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public long? idNhanVien { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }
}