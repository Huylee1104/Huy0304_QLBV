using M0304.Models.ThongTinDoanhNghiep;
using System.ComponentModel.DataAnnotations;

namespace M0304.Models.BaoCaoSoLieuThuThuat
{
    public class M0304BaoCaoSoLieuThuThuat
    {
        public string? SoPhieu { get; set; }
        public string? ThietBi { get; set; }
        public string? MaYTe { get; set; }
        public string? MaDot { get; set; }
        public string? TenBenhNhan { get; set; }
        public int? NamSinh { get; set; }
        public string? GioiTinh { get; set; }
        public string? DiaChi { get; set; }
        public string? TenDichVu { get; set; }
        public string? DoiTuong { get; set; }
        public string? PhuongPhapVoCam { get; set; }
        public string? LoaiThuThuat { get; set; }
        public string? BacSiThucHien { get; set; }
        public string? DieuDuong { get; set; }
        public DateTime? NgayChiDinh { get; set; }
        public DateTime? NgayThucHien { get; set; }
        public string? NoiYeuCau { get; set; }
        public string? BacSiChiDinh { get; set; }
        public string? NoiThucHien { get; set; }
        public string? LoaiGia { get; set; }
        public string? MaHoaDon { get; set; }
    }

    public class M0304BaoCaoSoLieuThuThuatPagedResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public IEnumerable<T> Data { get; set; } 
        public int TotalRecords { get; set; } 
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class M0304BaoCaoSoLieuThuThuatResponse
    {
        public M0304BaoCaoSoLieuThuThuatPagedResult<M0304BaoCaoSoLieuThuThuat> BaoCaoSoLieuThuThuat { get; set; } 
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }

    public class ExportRequest
    {
        public List<M0304BaoCaoSoLieuThuThuat> Data { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }
}