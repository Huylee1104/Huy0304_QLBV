using System.ComponentModel.DataAnnotations;

namespace M0304NhanVien.Models
{
    public class M0304NhanVienModel
    {
        [Key]
        public long ID { get; set; }
        public string MaNhanVien { get; set; }
        public string TenNhanVien { get; set; }
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