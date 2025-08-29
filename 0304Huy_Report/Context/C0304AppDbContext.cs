using M0304.Models.BangKeThu;
using M0304.Models.ThongTinDoanhNghiep;
using M0304B.Models.BCHoaDonDienTuDV;
using Microsoft.EntityFrameworkCore;

namespace C0304.Db.Models
{
    public class M0304Context : DbContext
    {
        public M0304Context(DbContextOptions<M0304Context> options)
            : base(options)
        {
        }

        public DbSet<M0304BangKeThu> M0304BangKeThus { get; set; }
        public DbSet<M0304BBCHoaDonDienTuDV> M0304BBCHoaDonDienTuDVs { get; set; }
        public DbSet<M0304ThongTinDoanhNghiep> M0304ThongTinDoanhNghieps { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<M0304BBCHoaDonDienTuDV>().ToTable("T0304_BCHoaDonDienTuDV");
            modelBuilder.Entity<M0304BangKeThu>().ToTable("T0304_BangKeThuNgoaiTru");
            modelBuilder.Entity<M0304ThongTinDoanhNghiep>().ToTable("ThongTinDoanhNghiep");
        }
    }
}