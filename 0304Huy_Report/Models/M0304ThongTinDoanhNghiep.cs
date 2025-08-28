using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace M0304.Models.ThongTinDOanhNghiep
{

    public class M0304ThongTinDoanhNghiep
    {
        [Key]
        public long ID { get; set;}
        public string TenCSKCB { get; set; }
        public string TenCoQuanChuyenMon { get; set; }
        public string DiaChi { get; set; }
        public string DienThoai { get; set; }
    }
}