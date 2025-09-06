using System.ComponentModel.DataAnnotations;

namespace M0304NhanVien.Models
{
    public class M0304ChucVu
    {
        public int id { get; set; }
        public string ten { get; set; }
    }
    public class M0304NhanVienModel
    {
        [Key]
        public long Id { get; set; }
        public string Ma { get; set; }
        public string Ten { get; set; }
        public int? IdKhoa { get; set; }
        public string TenKhoa { get; set; }
        public bool? KhoaCanLamSang { get; set; }
        public string Viettat { get; set; }
        public int? Slbn { get; set; }
        public bool Active { get; set; }
        public int? idChuyenMon { get; set; }
        public string HocVi { get; set; }
        public string MoTa { get; set; }
        public M0304ChucVu ChucVu { get; set; }
    }

    public class M0304TongTheoNhanVien
    {
        public long IDNhanVien { get; set; }
        public decimal TongHuy { get; set; }
        public decimal TongHoan { get; set; }
        public decimal TongSoTien { get; set; }
        public decimal TongChenhLech { get; set; }
    }
}