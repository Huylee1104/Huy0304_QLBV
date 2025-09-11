using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace M0304.Models.KhoHang
{

    public class M0304KhoHang
    {
        [Key]
        public long ID { get; set;}
        public string MaKhoHang { get; set; }
        public string TenKhoHang { get; set; }
    }
}