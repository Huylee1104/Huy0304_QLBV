using M0304.Models.ThongTinDoanhNghiep;
using M0304NhanVien.Models;
using System.ComponentModel.DataAnnotations;

namespace M0304.Models.HoatDongKhamBenh
{
    public class M0304HoatDongKhamBenh
    {
        public string? DichVu { get; set; }
        public double? TongSo { get; set; }
        public double? YHocCoTruyen { get; set; }
        public double? TreEmDuoi6Tuoi { get; set; }
        public double? BHYT { get; set; }
        public double? VienPhi { get; set; }
        public double? KhongThuDuoc { get; set; }
        public double? CapCuu { get; set; }
        public double? SoNguoiVaoVien { get; set; }
        public double? SoNguoiChuyenVien { get; set; }
        public double? NTSoNguoiBenh { get; set; }
        public double? NTYHocCoTruyen { get; set; }
        public double? NTTreEmDuoi6Tuoi { get; set; }
        public double? NTSoNgay { get; set; }
    }

    public class M0304HoatDongKhamBenhPagedResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public IEnumerable<T> Data { get; set; } 
        public int TotalRecords { get; set; } 
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class M0304HoatDongKhamBenhResponse
    {
        public M0304HoatDongKhamBenhPagedResult<M0304HoatDongKhamBenh> HoatDongKhamBenh { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }

        public double? AllTongSo { get; set; }
        public double? AllYHocCoTruyen { get; set; }
        public double? AllTEDuoi6Tuoi { get; set; }
        public double? AllBHYT { get; set; }
        public double? AllVienPhi { get; set; }
        public double? AllKhongThuDuoc { get; set; }
        public double? AllCapCuu { get; set; }
        public double? AllSoNguoiVaoVien { get; set; }
        public double? AllSoNguoiChuyenVien { get; set; }
        public double? AllNTSoNguoiBenh { get; set; }
        public double? AllNTYHocCoTruyen { get; set; }
        public double? AllNTTEDuoi6Tuoi { get; set; }
        public double? AllNTSoNgay { get; set; }
    }

    public class ExportRequest
    {
        public List<M0304HoatDongKhamBenh> Data { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public string? TenNVDN { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }
}