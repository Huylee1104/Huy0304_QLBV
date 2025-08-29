using M0304.Models.ThongTinDOanhNghiep;
using System.ComponentModel.DataAnnotations;

namespace M0304B.Models.BCHoaDonDienTuDV
{
    public class M0304BBCHoaDonDienTuDV
    {
        [Key]
        public long ID { get; set; }
        public string SoChungTu { get; set; }
        public DateTime? NgayThu { get; set; }
        public decimal? GiaTri { get; set; }
        public string MaBenhNhan { get; set; }
        public string TenBenhNhan { get; set; }
        public int? NamSinh { get; set; }
        public string DiaChi { get; set; }
        public DateTime? NgayTaoHDDT { get; set; }
        public string E_InvoiceNo { get; set; }
        public decimal? GiaTriHDDT { get; set; }
        public string MaTraCuu { get; set; }
    }

    public class M0304BPagedResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public IEnumerable<T> Data { get; set; } 
        public int TotalRecords { get; set; } 
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class M0304BBCHoaDonDienTuDVResponse
    {
        public M0304BPagedResult<M0304BBCHoaDonDienTuDV> BCHoaDonDienTuDV { get; set; }   // danh sách bảng kê thu phân trang
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }    // thông tin doanh nghiệp
    }

    public class ExportRequest
    {
        public List<M0304BBCHoaDonDienTuDV> Data { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }
}