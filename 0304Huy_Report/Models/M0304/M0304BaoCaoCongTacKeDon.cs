using M0304.Models.ThongTinDoanhNghiep;
using System.ComponentModel.DataAnnotations;

namespace M0304.Models.BaoCaoCongTacKeDon
{
    public class M0304BaoCaoCongTacKeDon
    {
        public string? MaYTe { get; set; }
        public string? TenBenhNhan { get; set; }
        public int? NamSinh { get; set; }
        public string? GioiTinh { get; set; }
        public string? DoiTuong { get; set; }
        public string? SoLuuTru { get; set; }
        public string? SoBenhAn { get; set; }
        public string? KhoaDieuTri { get; set; }
        public DateTime? NgayKham { get; set; }
        public string? TenPhongKham { get; set; }
        public string? BacSiKeToa { get; set; }
        public string? TenThuoc { get; set; }
        public string? TenHoatChat { get; set; }
        public int? SoNgay { get; set; }
        public double? SoLuong { get; set; }
        public double? SoLuongPhat { get; set; }
        public decimal? DonGia { get; set; }
        public string? ChanDoan { get; set; }
        public string? MucDichXuat { get; set; }
    }

    public class M0304BaoCaoCongTacKeDonPagedResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public IEnumerable<T> Data { get; set; } 
        public int TotalRecords { get; set; } 
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class M0304BaoCaoCongTacKeDonResponse
    {
        public M0304BaoCaoCongTacKeDonPagedResult<M0304BaoCaoCongTacKeDon> BaoCaoCongTacKeDon { get; set; } 
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }

    public class ExportRequest
    {
        public List<M0304BaoCaoCongTacKeDon> Data { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public long? IdNhomHang { get; set; }
        public long? IdHangHoa { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }
}