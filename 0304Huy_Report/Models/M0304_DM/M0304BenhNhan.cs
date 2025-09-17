using System.ComponentModel.DataAnnotations;

namespace M0304BenhNhan.Models
{
    public class M0304BenhNhanModel
    {
        [Key]
        public long ID { get; set; }
        public string MaBN { get; set; }
        public string TenBN { get; set; }
    }
}