using System.ComponentModel.DataAnnotations;

namespace M0304HTTT.Models
{
    public class M0304HTTTModel
    {
        [Key]
        public int id { get; set; }
        public string ten { get; set; }
    }
}