using M0304.Models.BangKeThu;
using M0304.Models.KhoHang;
using M0304.Models.ThongTinDoanhNghiep;
using M0304BenhNhan.Models;
using M0304B.Models.BCHoaDonDienTuDV;
using M0304C.Models.BaoCaoThuDichVu;
using M0304D.Models.BKBienLaiTamUng;
using M0304E.Models.BKBienLaiHoanUng;
using M0304F.Models.BKTinhHinhTraDuocNCC;
using M0304G.Models.PhieuLinhVatTuYTe;
using M0304H.Models.BCTongSoSIDTheoKhoaPhong;
using M0304NhanVien.Models;
using M0304I.Models.PhieuTheoDoiChucNangSong;
using M0304L.Models.PhieuTheoDoiTruyenDich;
using M0304M.Models.BaoCaoHangHoa;
using M0304.Models.BCTTCapPhatThuocKS_KVirut;
using M0304.Models.BaoCaoCongTacKeDon;
using M0304.Models.ToKhaiChiTietThuPhiLePhi;
using M0304.Models.HoatDongKhamBenh;
using M0304.Models.BangKeBanLeHangHoaDichVu;
using M0304.Models.BaoCaoSoLieuThuThuat;
using M0304.Models.BaoCaoChiDinhCLS_Phong_BS;
using M0304.Models.BaoCaoTiepNhan;
using M0304.Models.BCThongKeTheoMaBenhICD;
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
        public DbSet<M0304BenhNhanModel> M0304BenhNhans { get; set; }
        public DbSet<M0304NhanVienModel> M0304NhanViens { get; set; }
        public DbSet<M0304CBaoCaoThuDichVu> M0304CBaoCaoThuDichVus { get; set; }
        public DbSet<M0304DBKBienLaiTamUng> M0304DBKBienLaiTamUngs { get; set; }
        public DbSet<M0304EBKBienLaiHoanUng> M0304EBKBienLaiHoanUngs { get; set; }
        public DbSet<M0304FBKTinhHinhTraDuocNCC> M0304FBKTinhHinhTraDuocNCCs { get; set; }
        public DbSet<M0304GPhieuLinhVatTuYTe> M0304GPhieuLinhVatTuYTes { get; set; }
        public DbSet<M0304HBCTongSoSIDTheoKhoaPhong> M0304HBCTongSoSIDTheoKhoaPhongs { get; set; }
        public DbSet<BenhNhanThongTinModel> BenhNhanThongTins { get; set; }
        public DbSet<SinhHieuModel> SinhHieus { get; set; }
        public DbSet<ThongTinBNModel> ThongTinBNs { get; set; }
        public DbSet<TruyenDich> TruyenDichs { get; set; }
        public DbSet<M0304MHangNhap> HangNhapReports { get; set; }
        public DbSet<M0304MHangXuat> HangXuatReports { get; set; }
        public DbSet<M0304BCTTCapPhatThuocKS_KVirut> BCTTCapPhatThuocKS_KViruts { get; set; }
        public DbSet<M0304BaoCaoCongTacKeDon> BaoCaoCongTacKeDons { get; set; }
        public DbSet<M0304ToKhaiChiTietThuPhiLePhi> ToKhaiChiTietThuPhiLePhis { get; set; }
        public DbSet<M0304HoatDongKhamBenh> HoatDongKhamBenhs { get; set; }
        public DbSet<M0304BangKeBanLeHangHoaDichVu> BangKeBanLeHangHoaDichVus { get; set; }
        public DbSet<M0304BaoCaoSoLieuThuThuat> BaoCaoSoLieuThuThuats { get; set; }
        public DbSet<M0304BaoCaoTiepNhan> BaoCaoTiepNhans { get; set; }
        public DbSet<M0304BaoCaoChiDinhCLS_Phong_BS> BaoCaoChiDinhCLS_Phong_BSs { get; set; }
        public DbSet<M0304BCThongKeTheoMaBenhICD> BCThongKeTheoMaBenhICDs { get; set; }



        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<M0304BBCHoaDonDienTuDV>().ToTable("T0304_BCHoaDonDienTuDV");
            modelBuilder.Entity<M0304BangKeThu>().HasNoKey();
            modelBuilder.Entity<M0304DBKBienLaiTamUng>().HasNoKey();
            modelBuilder.Entity<M0304EBKBienLaiHoanUng>().ToTable("T0304_BKBienLaiHoanUng");
            modelBuilder.Entity<M0304ThongTinDoanhNghiep>().ToTable("ThongTinDoanhNghiep");
            modelBuilder.Entity<M0304CBaoCaoThuDichVu>().HasNoKey();
            modelBuilder.Entity<M0304KhoHang>().ToTable("HH_DM_KhoHang");
            modelBuilder.Entity<M0304BenhNhanModel>().ToTable("DM_BenhNhan");
            modelBuilder.Entity<M0304NhanVienModel>().ToTable("DM_NhanVien");
            modelBuilder.Entity<M0304FBKTinhHinhTraDuocNCC>().HasNoKey();
            modelBuilder.Entity<M0304GPhieuLinhVatTuYTe>().HasNoKey();
            modelBuilder.Entity<M0304HBCTongSoSIDTheoKhoaPhong>().HasNoKey();
            modelBuilder.Entity<BenhNhanThongTinModel>().HasNoKey();
            modelBuilder.Entity<SinhHieuModel>().HasNoKey();
            modelBuilder.Entity<ThongTinBNModel>().HasNoKey();
            modelBuilder.Entity<TruyenDich>().HasNoKey();
            modelBuilder.Entity<M0304MHangNhap>().HasNoKey();
            modelBuilder.Entity<M0304MHangXuat>().HasNoKey();
            modelBuilder.Entity<M0304BCTTCapPhatThuocKS_KVirut>().HasNoKey();
            modelBuilder.Entity<M0304BaoCaoCongTacKeDon>().HasNoKey();
            modelBuilder.Entity<M0304ToKhaiChiTietThuPhiLePhi>().HasNoKey();
            modelBuilder.Entity<M0304HoatDongKhamBenh>().HasNoKey();
            modelBuilder.Entity<M0304BangKeBanLeHangHoaDichVu>().HasNoKey();
            modelBuilder.Entity<M0304BaoCaoSoLieuThuThuat>().HasNoKey();
            modelBuilder.Entity<M0304BaoCaoTiepNhan>().HasNoKey();
            modelBuilder.Entity<M0304BaoCaoChiDinhCLS_Phong_BS>().HasNoKey();
            modelBuilder.Entity<M0304BaoCaoSoLieuThuThuat>().HasNoKey();
            modelBuilder.Entity<M0304BCThongKeTheoMaBenhICD>().HasNoKey();
        }
    }
}