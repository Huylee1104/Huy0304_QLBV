using M0304.Models.ThongTinDoanhNghiep;
using System.ComponentModel.DataAnnotations;

namespace M0304.Models.BCTTCapPhatThuocKS_KVirut
{
    public class M0304BCTTCapPhatThuocKS_KVirut
    {
        public string? MaYTe { get; set; }
        public string? SoBenhAn { get; set; }
        public string? TenBenhNhan { get; set; }
        public int? NamSinh { get; set; }
        public string? DiaChi { get; set; }
        public string? KhoaDieuTri { get; set; }
        public string? TenPhongKham { get; set; }
        public string? BacSiKeDon { get; set; }
        public string? TenThuoc { get; set; }
        public string? TenHoatChat { get; set; }
        public string? HamLuong { get; set; }
        public string? DVT { get; set; }
        public string? DuongDung { get; set; }
        public string? LieuDung { get; set; }
        public int? SoNgay { get; set; }
        public double? SoLuongKeDon { get; set; }
        public double? SoLuongXuat { get; set; }
        public string? ChanDoan { get; set; }
    }

    public class M0304BCTTCapPhatThuocKS_KVirutPagedResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public IEnumerable<T> Data { get; set; } 
        public int TotalRecords { get; set; } 
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class M0304BCTTCapPhatThuocKS_KVirutResponse
    {
        public M0304BCTTCapPhatThuocKS_KVirutPagedResult<M0304BCTTCapPhatThuocKS_KVirut> BCTTCapPhatThuocKS_KVirut { get; set; } 
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }

    public class ExportRequest
    {
        public List<M0304BCTTCapPhatThuocKS_KVirut> Data { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public long? IdNhomHang { get; set; }
        public long? IdHangHoa { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }
}