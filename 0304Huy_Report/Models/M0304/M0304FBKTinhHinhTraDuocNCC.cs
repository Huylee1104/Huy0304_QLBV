using M0304.Models.ThongTinDoanhNghiep;
using M0304NhanVien.Models;
using System.ComponentModel.DataAnnotations;

namespace M0304F.Models.BKTinhHinhTraDuocNCC
{
    public class M0304FBKTinhHinhTraDuocNCC
    {
        //[Key]
        //public int Id { get; set; }
        public DateTime? NgayHoaDon { get; set; }
        public string? SoHoaDon { get; set; }
        public DateTime? NgayTra { get; set; }
        public string? PhieuTra { get; set; }
        public string? CongTy { get; set; }
        public long? IDCongTy { get; set; }
        public string? MaID { get; set; }
        public string? TenThuoc { get; set; }
        public string? QuyCach { get; set; }
        public string? SoLo { get; set; }
        public long? SLDongGoi { get; set; }
        public long? SLLe { get; set; }

        public decimal? DonGiaDongGoi { get; set; }
        public decimal? DonGiaLe { get; set; }
        public decimal? ThanhTien { get; set; }
    }

    public class M0304FPagedResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public IEnumerable<T> Data { get; set; } 
        public int TotalRecords { get; set; } 
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class M0304FBKTinhHinhTraDuocNCCResponse
    {
        public M0304FPagedResult<M0304FBKTinhHinhTraDuocNCC> BKTinhHinhTraDuocNCC { get; set; }   // danh sách bảng kê thu phân trang
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }    // thông tin doanh nghiệp
    }

    public class ExportRequest
    {
        public List<M0304FBKTinhHinhTraDuocNCC> Data { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public long? IdKhoHang { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }

    public class CongTyDto
    {
        public long ID { get; set; }
        public string Ten { get; set; } = string.Empty;
    }
}