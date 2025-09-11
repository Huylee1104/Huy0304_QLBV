using M0304.Models.ThongTinDoanhNghiep;
using System.ComponentModel.DataAnnotations;

namespace M0304H.Models.BCTongSoSIDTheoKhoaPhong
{
    public class M0304HBCTongSoSIDTheoKhoaPhong
    {
        //[Key]
        //public long ID { get; set; }
        public string? TenKhoaPhong { get; set; }
        public int? VienPhi { get; set; }
        public int? QL01 { get; set; }
        public int? QL02 { get; set; }
        public int? QL03 { get; set; }
        public int? QL04 { get; set; }
        public int? QL05 { get; set; }
        public int? DichVu { get; set; }
        public int? KhamChuyenGia { get; set; }
        public int? Tong { get; set; }
    }

    public class M0304HPagedResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public IEnumerable<T> Data { get; set; } 
        public int TotalRecords { get; set; } 
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class M0304HBCTongSoSIDTheoKhoaPhongResponse
    {
        public M0304HPagedResult<M0304HBCTongSoSIDTheoKhoaPhong> BCTongSoSIDTheoKhoaPhong { get; set; }   // danh sách bảng kê thu phân trang
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }    // thông tin doanh nghiệp
    }

    public class ExportRequest
    {
        public List<M0304HBCTongSoSIDTheoKhoaPhong> Data { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }
}