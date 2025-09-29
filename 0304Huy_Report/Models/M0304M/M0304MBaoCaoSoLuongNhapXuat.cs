using M0304.Models.ThongTinDoanhNghiep;
using System.ComponentModel.DataAnnotations;

namespace M0304M.Models.BaoCaoHangHoa
{
    // ===== MODEL NHẬP =====
    public class M0304MHangNhap
    {
        public string? TenThuoc { get; set; }
        public string? TenNhomHang { get; set; }
        public long? IDNhomHang { get; set; }

        public double? Thang1 { get; set; }
        public double? Thang2 { get; set; }
        public double? Thang3 { get; set; }
        public double? Thang4 { get; set; }
        public double? Thang5 { get; set; }
        public double? Thang6 { get; set; }
        public double? Thang7 { get; set; }
        public double? Thang8 { get; set; }
        public double? Thang9 { get; set; }
        public double? Thang10 { get; set; }
        public double? Thang11 { get; set; }
        public double? Thang12 { get; set; }

        public double? TongCong { get; set; }
    }

    public class M0304MHangNhapResponse
    {
        public M0304MPagedResult<M0304MHangNhap> HangNhap { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }

    // ===== MODEL XUẤT =====
    public class M0304MHangXuat
    {
        public string? TenThuoc { get; set; }
        public string? TenNhomHang { get; set; }
        public long? IDNhomHang { get; set; }

        public double? Thang1 { get; set; }
        public double? Thang2 { get; set; }
        public double? Thang3 { get; set; }
        public double? Thang4 { get; set; }
        public double? Thang5 { get; set; }
        public double? Thang6 { get; set; }
        public double? Thang7 { get; set; }
        public double? Thang8 { get; set; }
        public double? Thang9 { get; set; }
        public double? Thang10 { get; set; }
        public double? Thang11 { get; set; }
        public double? Thang12 { get; set; }

        public double? TongCong { get; set; }
    }

    public class M0304GHangXuatResponse
    {
        public M0304MPagedResult<M0304MHangXuat> HangXuat { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }

    // ===== DÙNG CHUNG =====
    public class M0304MPagedResult<T>
    {
        public bool Success { get; set; }
        public string? Message { get; set; }
        public IEnumerable<T>? Data { get; set; }
        public int TotalRecords { get; set; }
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
    }

    public class ExportRequest<T>
    {
        public List<T> Data { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public long? IdKhoHang { get; set; }
        public long? IdNhomHang { get; set; }
        public long? IdHangHoa { get; set; }
        public M0304ThongTinDoanhNghiep DoanhNghiep { get; set; }
    }
}
