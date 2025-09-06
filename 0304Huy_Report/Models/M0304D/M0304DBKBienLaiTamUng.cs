using M0304.Models.ThongTinDoanhNghiep;
using System.ComponentModel.DataAnnotations;

namespace M0304D.Models.BKBienLaiTamUng
{
    public class M0304DBKBienLaiTamUng
    {
        [Key]
        public long ID { get; set; }
        public DateTime? NgayThu { get; set; }
        public string? MaYTe { get; set; }
        public string? SoBA { get; set; }
        public string? MaDot { get; set; }
        public string? HoTenBenhNhan { get; set; }
        public string? SoBL { get; set; }
        public string ? SoQuyen { get; set; }
        public decimal? ThuPhi { get; set; }
        public decimal? Huy { get; set; }
        public decimal? HoanTra { get; set; }
        public string? HTTT { get; set; }
        public long? IDCN { get; set; }
        public long? IDNhanVien { get; set; }
    }

    public class M0304DPagedResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public IEnumerable<T> Data { get; set; } 
        public int TotalRecords { get; set; } 
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class M0304DBKBienLaiTamUngResponse
    {
        public M0304DPagedResult<M0304DBKBienLaiTamUng> BangKeBienLaiTamUng { get; set; }   // danh sách bảng kê thu phân trang
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }    // thông tin doanh nghiệp
    }

    public class ExportRequest
    {
        public List<M0304DBKBienLaiTamUng> Data { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public long? IdNhanVien { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }
}