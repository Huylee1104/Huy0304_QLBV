using M0304.Models.ThongTinDoanhNghiep;
using M0304NhanVien.Models;
using System.ComponentModel.DataAnnotations;

namespace M0304.Models.BangKeThu
{
    public class M0304BangKeThu
    {
        [Key]
        public int Id { get; set; }
        public string? MaYTe { get; set; }
        public string? HoVaTen { get; set; }
        public string? QuyenSo { get; set; }
        public string? SoBienLai { get; set; }
        public string? Loai { get; set; }
        public DateTime? NgayThu { get; set; }
        public decimal? Huy { get; set; }
        public decimal? Hoan { get; set; }
        public decimal? SoTien { get; set; }
        public long? IDCN { get; set; }
        public long? IDHTTT { get; set; }
        public long? IDNhanVien { get; set; }
    }

    public class M0304PagedResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public IEnumerable<T> Data { get; set; } 
        public int TotalRecords { get; set; } 
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class M0304BangKeThuResponse
    {
        public M0304PagedResult<M0304BangKeThu> BangKeThu { get; set; }   // danh sách bảng kê thu phân trang
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }    // thông tin doanh nghiệp
    }

    public class ExportRequest
    {
        public List<M0304BangKeThu> Data { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public long? IdHTTT { get; set; }
        public long? IdNhanVien { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }

    public class M0304TongHopBangKeThu
    {
        public decimal TongHuy { get; set; }
        public decimal TongHoan { get; set; }
        public decimal TongSoTien { get; set; }
        public decimal TongChenhLech { get; set; }
    }

    public class ReportSummary
    {
        public M0304TongHopBangKeThu TongChung { get; set; }
        public List<M0304TongTheoNhanVien> TongTheoNhanVien { get; set; }
    }
}