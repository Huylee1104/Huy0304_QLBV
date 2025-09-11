using M0304.Models.ThongTinDoanhNghiep;
using M0304NhanVien.Models;
using System.ComponentModel.DataAnnotations;

namespace M0304G.Models.PhieuLinhVatTuYTe
{
    public class M0304GPhieuLinhVatTuYTe
    {
        //[Key]
        //public int Id { get; set; }
        //public DateTime? NgayLinhVatTu { get; set; }
        public string? MaVatTu { get; set; }
        public string? TenVatTu { get; set; }
        public string? DonViTinh { get; set; }
        public int? SoLuong { get; set; }
        //public long? IDKhoTra { get; set; }
    }

    public class M0304GPagedResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public IEnumerable<T> Data { get; set; } 
        public int TotalRecords { get; set; } 
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class M0304GPhieuLinhVatTuYTeResponse
    {
        public M0304GPagedResult<M0304GPhieuLinhVatTuYTe> PhieuLinhVatTuYTe { get; set; }   // danh sách bảng kê thu phân trang
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }    // thông tin doanh nghiệp
    }

    public class ExportRequest
    {
        public List<M0304GPhieuLinhVatTuYTe> Data { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public long? IdKhoHang { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }
}