using M0304.Models.ThongTinDoanhNghiep;
using M0304NhanVien.Models;
using System.ComponentModel.DataAnnotations;

namespace M0304.Models.ToKhaiChiTietThuPhiLePhi
{
    public class M0304ToKhaiChiTietThuPhiLePhi
    {
        public string? QuyenSo { get; set; }
        public string? LoaiHoaDon { get; set; }
        public string? SoLan_soBLHDthu { get; set; }
        public int? SoLuongHDSuDung { get; set; }
        public double? TongSoTien { get; set; }
        public double? Huy_Hoan { get; set; }
        public double? SoTienThucThu { get; set; }
        public string? GhiChu { get; set; }
        public long? IDNhanVien { get; set; }
        public string? TenNhanVien { get; set; }
    }

    public class M0304ToKhaiChiTietThuPhiLePhiPagedResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public IEnumerable<T> Data { get; set; } 
        public int TotalRecords { get; set; } 
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class M0304ToKhaiChiTietThuPhiLePhiResponse
    {
        public M0304ToKhaiChiTietThuPhiLePhiPagedResult<M0304ToKhaiChiTietThuPhiLePhi> ToKhaiChiTietThuPhiLePhi { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }

        public double? AllSoTien { get; set; }
        public double? AllHoan_Huy { get; set; }
        public double? AllTienThucThu { get; set; }
    }

    public class ExportRequest
    {
        public List<M0304ToKhaiChiTietThuPhiLePhi> Data { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }
}