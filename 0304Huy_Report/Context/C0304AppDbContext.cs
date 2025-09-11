using M0304.Models.BangKeThu;
using M0304.Models.KhoHang;
using M0304.Models.ThongTinDoanhNghiep;
using M0304B.Models.BCHoaDonDienTuDV;
using M0304C.Models.BaoCaoThuDichVu;
using M0304D.Models.BKBienLaiTamUng;
using M0304E.Models.BKBienLaiHoanUng;
using M0304F.Models.BKTinhHinhTraDuocNCC;
using M0304G.Models.PhieuLinhVatTuYTe;
using M0304H.Models.BCTongSoSIDTheoKhoaPhong;
using M0304NhanVien.Models;
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
        public DbSet<M0304KhoHang> M0304KhoHangs { get; set; }
        public DbSet<M0304NhanVienModel> M0304NhanViens { get; set; }
        public DbSet<M0304CBaoCaoThuDichVu> M0304CBaoCaoThuDichVus { get; set; }
        public DbSet<M0304DBKBienLaiTamUng> M0304DBKBienLaiTamUngs { get; set; }
        public DbSet<M0304EBKBienLaiHoanUng> M0304EBKBienLaiHoanUngs { get; set; }
        public DbSet<M0304FBKTinhHinhTraDuocNCC> M0304FBKTinhHinhTraDuocNCCs { get; set; }
        public DbSet<M0304GPhieuLinhVatTuYTe> M0304GPhieuLinhVatTuYTes { get; set; }
        public DbSet<M0304HBCTongSoSIDTheoKhoaPhong> M0304HBCTongSoSIDTheoKhoaPhongs { get; set; }


        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<M0304BBCHoaDonDienTuDV>().ToTable("T0304_BCHoaDonDienTuDV");
            modelBuilder.Entity<M0304BangKeThu>().ToTable("T0304_BangKeThuNgoaiTru");
            modelBuilder.Entity<M0304DBKBienLaiTamUng>().HasNoKey();
            modelBuilder.Entity<M0304EBKBienLaiHoanUng>().ToTable("T0304_BKBienLaiHoanUng");
            modelBuilder.Entity<M0304ThongTinDoanhNghiep>().ToTable("ThongTinDoanhNghiep");
            modelBuilder.Entity<M0304CBaoCaoThuDichVu>().HasNoKey();
            modelBuilder.Entity<M0304KhoHang>().ToTable("HH_DM_KhoHang");
            modelBuilder.Entity<M0304NhanVienModel>().ToTable("DM_NhanVien");
            modelBuilder.Entity<M0304FBKTinhHinhTraDuocNCC>().HasNoKey();
            modelBuilder.Entity<M0304GPhieuLinhVatTuYTe>().HasNoKey();
            modelBuilder.Entity<M0304HBCTongSoSIDTheoKhoaPhong>().HasNoKey();
        }
    }
}