using M0304.Models.ThongTinDoanhNghiep;
using M0304NhanVien.Models;
using System.ComponentModel.DataAnnotations;

namespace M0304.Models.BaoCaoChiDinhCLS_Phong_BS
{
    public class M0304BaoCaoChiDinhCLS_Phong_BS
    {
        public string? NoiGui { get; set; }
        public string? BacSi { get; set; }
        public string? YeuCau { get; set; }
        public double? DonGia { get; set; }
        public int? SoLuot { get; set; }
    }

    public class M0304BaoCaoChiDinhCLS_Phong_BSPagedResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public IEnumerable<T> Data { get; set; } 
        public int TotalRecords { get; set; } 
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class M0304BaoCaoChiDinhCLS_Phong_BSResponse
    {
        public M0304BaoCaoChiDinhCLS_Phong_BSPagedResult<M0304BaoCaoChiDinhCLS_Phong_BS> BaoCaoChiDinhCLS_Phong_BS { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }

        public double? TongDonGia { get; set; }
        public long? TongLuot { get; set; }
    }

    public class ExportRequest
    {
        public List<M0304BaoCaoChiDinhCLS_Phong_BS> Data { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public long? idPhongBuong { get; set; }
        public long? idBacSi { get; set; }
        public long? idDVKT { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }
}