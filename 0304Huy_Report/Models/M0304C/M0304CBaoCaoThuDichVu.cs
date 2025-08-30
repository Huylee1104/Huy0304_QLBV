using M0304.Models.ThongTinDoanhNghiep;
using System.ComponentModel.DataAnnotations;

namespace M0304C.Models.BaoCaoThuDichVu
{
    public class M0304CBaoCaoThuDichVu
    {
        [Key]
        public long ID { get; set; }
        public string? NhomDichVu { get; set; }
        public string? DichVu { get; set; }
        public int? SoLuong { get; set; }
        public decimal? TongHoaDon { get; set; }
    }

    public class M0304CPagedResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public IEnumerable<T> Data { get; set; } 
        public int TotalRecords { get; set; } 
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class M0304CBaoCaoThuDichVuResponse
    {
        public M0304CPagedResult<M0304CBaoCaoThuDichVu> BaoCaoThuDichVu { get; set; }   // danh sách bảng kê thu phân trang
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }    // thông tin doanh nghiệp
    }

    public class ExportRequest
    {
        public List<M0304CBaoCaoThuDichVu> Data { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }
}