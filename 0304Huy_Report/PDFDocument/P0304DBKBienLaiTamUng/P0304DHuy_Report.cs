using H0304.NumberToText.Helpers;
using M0304NhanVien.Models;
using M0304D.Models.BKBienLaiTamUng;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.IO;
using M0304.Models.ThongTinDoanhNghiep;

namespace P0304D.PDFDocument
{
    public class P0304DReportTemplatePDF : IDocument
    {
        private readonly List<M0304DBKBienLaiTamUng> _data;
        private readonly M0304ThongTinDoanhNghiep _dataDN;
        private string _ngayBatDau;
        private string _ngayKetThuc;
        private readonly string _logoPath;

        private List<M0304TongTheoNhanVien> _tongTheoNhanVien;
        private List<M0304NhanVienModel> _danhSachNhanVien;

        public P0304DReportTemplatePDF(
            List<M0304DBKBienLaiTamUng> data,
            string ngayBatDau,
            string ngayKetThuc,
            List<M0304NhanVienModel> danhSachNhanVien,
            List<M0304TongTheoNhanVien> tongTheoNhanVien,
            M0304ThongTinDoanhNghiep dataDN,
            string logoPath = null
        )
        {
            _data = data ?? new List<M0304DBKBienLaiTamUng>();
            _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
            _ngayBatDau = ngayBatDau;
            _ngayKetThuc = ngayKetThuc;
            _danhSachNhanVien = danhSachNhanVien ?? new List<M0304NhanVienModel>();
            _tongTheoNhanVien = tongTheoNhanVien ?? new List<M0304TongTheoNhanVien>();
            _logoPath = logoPath;
        }

        public DocumentMetadata GetMetadata() => DocumentMetadata.Default;
        public DocumentSettings GetSettings() => DocumentSettings.Default;

        public void Compose(IDocumentContainer container)
        {
            container.Page(page =>
            {
                page.Size(PageSizes.A4);
                page.Margin(15);
                page.DefaultTextStyle(x => x.FontSize(8));

                page.Header().ShowOnce().Column(col =>
                {
                    col.Item().Row(row =>
                    {
                        row.RelativeItem(3).Row(left =>
                        {
                            if (!string.IsNullOrEmpty(_logoPath) && File.Exists(_logoPath))
                            {
                                left.ConstantItem(40).AlignMiddle().Image(_logoPath);
                            }
                            else
                            {
                                left.ConstantItem(40).AlignMiddle().Text("No Logo");
                            }

                            left.RelativeItem().Column(info =>
                            {
                                info.Item().Text(_dataDN.TenCSKCB ?? "").FontSize(8);
                                info.Item().Text(_dataDN.TenCoQuanChuyenMon ?? "").FontSize(8);
                                info.Item().Text(_dataDN.DiaChi ?? "").FontSize(8);
                                info.Item().Text(_dataDN.DienThoai ?? "").FontSize(8);
                            });
                        });
                    });

                    col.Item().AlignCenter().Column(center =>
                    {
                        center.Item()
                            .AlignCenter()
                            .Text("BẢNG KÊ TẠM ỨNG")
                            .Bold()
                            .FontSize(12);

                        center.Item()
                            .Width(250)
                            .AlignCenter()
                            .Text($"Từ ngày {_ngayBatDau} đến ngày {_ngayKetThuc}")
                            .FontSize(8);
                    });
                });

                page.Content().PaddingVertical(6).Column(col =>
                {
                    col.Item().Table(table =>
                    {
                        table.ColumnsDefinition(columns =>
                        {
                            columns.ConstantColumn(25);
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(45);
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(50);
                            columns.RelativeColumn(2);
                            columns.RelativeColumn(1);
                            columns.RelativeColumn(1);
                            columns.RelativeColumn(1);
                            columns.ConstantColumn(40);
                            columns.ConstantColumn(40);
                            columns.ConstantColumn(35);
                        });


                        table.Header(header =>
                        {
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("STT");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Ngày thu");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Mã y tế");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Số BA");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Mã đợt");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Họ và tên");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Sổ quyển");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Số biên lai");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Thu phí");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Hủy");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Hoàn");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("HTTT");

                            table.Cell().Element(CellStyle).AlignCenter().Text("1");
                            table.Cell().Element(CellStyle).AlignCenter().Text("2");
                            table.Cell().Element(CellStyle).AlignCenter().Text("3");
                            table.Cell().Element(CellStyle).AlignCenter().Text("4");
                            table.Cell().Element(CellStyle).AlignCenter().Text("5");
                            table.Cell().Element(CellStyle).AlignCenter().Text("6");
                            table.Cell().Element(CellStyle).AlignCenter().Text("7");
                            table.Cell().Element(CellStyle).AlignCenter().Text("8");
                            table.Cell().Element(CellStyle).AlignCenter().Text("9");
                            table.Cell().Element(CellStyle).AlignCenter().Text("10");
                            table.Cell().Element(CellStyle).AlignCenter().Text("11");
                            table.Cell().Element(CellStyle).AlignCenter().Text("12");
                        });

                        decimal tongThuPhi = _data.Sum(x => x.ThuPhi ?? 0);
                        decimal tongHuy = _data.Sum(x => x.Huy ?? 0);
                        decimal tongHoanTra = _data.Sum(x => x.HoanTra ?? 0);
                        decimal tongChenhLech = tongThuPhi - (tongHuy + tongHoanTra);

                        // Render tất cả nhân viên trừ nhân viên cuối
                        var nhanVienTruCuoi = _danhSachNhanVien.Take(_danhSachNhanVien.Count() - 1);
                        foreach (var nv in nhanVienTruCuoi)
                        {
                            int stt = 1;
                            var tongNV = _tongTheoNhanVien.FirstOrDefault(t => t.IDNhanVien == nv.ID);
                            var dataNV = _data.Where(d => d.IDNhanVien == nv.ID).ToList();

                            table.Cell().ColumnSpan(12).Border(1).Padding(3).Element(container =>
                            {
                                container.Row(row =>
                                {
                                    row.ConstantItem(25 + 50 + 45 + 50 + 50).ExtendHorizontal().AlignLeft().Text($"Nhân viên: {nv.TenNhanVien}").FontSize(7).Bold();
                                    row.RelativeItem(2).Text("");
                                    row.RelativeItem(1).Text("");
                                    row.RelativeItem(1).Text("");
                                    row.RelativeItem(1).AlignRight().Text(tongNV?.TongSoTien.ToString("N0") ?? "0").FontSize(7).Bold();
                                    row.ConstantItem(40).AlignRight().Text(tongNV?.TongHuy.ToString("N0") ?? "0").FontSize(7).Bold();
                                    row.ConstantItem(40).AlignRight().Text(tongNV?.TongHoan.ToString("N0") ?? "0").FontSize(7).Bold();
                                    row.ConstantItem(35).Text("");
                                });
                            });
                            foreach (var item in dataNV)
                            {
                                table.Cell().Element(CellStyle).AlignCenter().Text(stt++.ToString());
                                table.Cell().Element(CellStyle).AlignCenter().Text(item.NgayThu?.ToString("dd-MM-yyyy HH:mm:ss") ?? "");
                                table.Cell().Element(CellStyle).AlignCenter().Text(item.MaYTe ?? "");
                                table.Cell().Element(CellStyle).AlignCenter().Text(item.SoBA ?? "");
                                table.Cell().Element(CellStyle).AlignCenter().Text(item.MaDot ?? "");
                                table.Cell().Element(CellStyle).AlignLeft().Text(item.HoTenBenhNhan ?? "");
                                table.Cell().Element(CellStyle).AlignCenter().Text(item.SoBL ?? "");
                                table.Cell().Element(CellStyle).AlignCenter().Text(item.SoQuyen ?? "");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.ThuPhi?.ToString("N0") ?? "");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.Huy?.ToString("N0") ?? "");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.HoanTra?.ToString("N0") ?? "");
                                table.Cell().Element(CellStyle).AlignCenter().Text(item.HTTT ?? "");
                            }
                        }

                        var nvCuoi = _danhSachNhanVien.Last();
                        int sttCuoi = 1;
                        var tongNVCuoi = _tongTheoNhanVien.FirstOrDefault(t => t.IDNhanVien == nvCuoi.ID);
                        var dataNVCuoi = _data.Where(d => d.IDNhanVien == nvCuoi.ID).ToList();

                        table.Cell().ColumnSpan(12).Border(1).Padding(3).Element(container =>
                        {
                            container.Row(row =>
                            {
                                row.ConstantItem(25 + 50 + 45 + 50 + 50).ExtendHorizontal().AlignLeft().Text($"Nhân viên: {nvCuoi.TenNhanVien}").FontSize(7).Bold();
                                row.RelativeItem(2).Text("");
                                row.RelativeItem(1).Text("");
                                row.RelativeItem(1).Text("");
                                row.RelativeItem(1).AlignRight().Text(tongNVCuoi?.TongSoTien.ToString("N0") ?? "0").FontSize(7).Bold();
                                row.ConstantItem(40).AlignRight().Text(tongNVCuoi?.TongHuy.ToString("N0") ?? "0").FontSize(7).Bold();
                                row.ConstantItem(40).AlignRight().Text(tongNVCuoi?.TongHoan.ToString("N0") ?? "0").FontSize(7).Bold();
                                row.ConstantItem(35).Text("");
                            });
                        });

                        if (dataNVCuoi.Count > 1)
                        {
                            foreach (var item in dataNVCuoi.Take(dataNVCuoi.Count - 1))
                            {
                                table.Cell().Element(CellStyle).AlignCenter().Text(sttCuoi++.ToString());
                                table.Cell().Element(CellStyle).AlignCenter().Text(item.NgayThu?.ToString("dd-MM-yyyy HH:mm:ss") ?? "");
                                table.Cell().Element(CellStyle).AlignCenter().Text(item.MaYTe ?? "");
                                table.Cell().Element(CellStyle).AlignCenter().Text(item.SoBA ?? "");
                                table.Cell().Element(CellStyle).AlignCenter().Text(item.MaDot ?? "");
                                table.Cell().Element(CellStyle).AlignLeft().Text(item.HoTenBenhNhan ?? "");
                                table.Cell().Element(CellStyle).AlignCenter().Text(item.SoBL ?? "");
                                table.Cell().Element(CellStyle).AlignCenter().Text(item.SoQuyen ?? "");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.ThuPhi?.ToString("N0") ?? "");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.Huy?.ToString("N0") ?? "");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.HoanTra?.ToString("N0") ?? "");
                                table.Cell().Element(CellStyle).AlignCenter().Text(item.HTTT ?? "");
                            }
                        }

                        col.Item().EnsureSpace(86).Column(group =>
                        {
                            group.Item().Table(lastRowTable =>
                            {
                                lastRowTable.ColumnsDefinition(columns =>
                                {
                                    columns.ConstantColumn(25);
                                    columns.ConstantColumn(50);
                                    columns.ConstantColumn(45);
                                    columns.ConstantColumn(50);
                                    columns.ConstantColumn(50);
                                    columns.RelativeColumn(2);
                                    columns.RelativeColumn(1);
                                    columns.RelativeColumn(1);
                                    columns.RelativeColumn(1);
                                    columns.ConstantColumn(40);
                                    columns.ConstantColumn(40);
                                    columns.ConstantColumn(35);
                                });

                                var lastItem = dataNVCuoi.Last();
                                lastRowTable.Cell().Element(CellStyle).AlignCenter().Text(sttCuoi.ToString());
                                lastRowTable.Cell().Element(CellStyle).AlignCenter().Text(lastItem.NgayThu?.ToString("dd-MM-yyyy HH:mm:ss") ?? "");
                                lastRowTable.Cell().Element(CellStyle).AlignCenter().Text(lastItem.MaYTe ?? "");
                                lastRowTable.Cell().Element(CellStyle).AlignCenter().Text(lastItem.SoBA ?? "");
                                lastRowTable.Cell().Element(CellStyle).AlignCenter().Text(lastItem.MaDot ?? "");
                                lastRowTable.Cell().Element(CellStyle).AlignLeft().Text(lastItem.HoTenBenhNhan ?? "");
                                lastRowTable.Cell().Element(CellStyle).AlignCenter().Text(lastItem.SoBL ?? "");
                                lastRowTable.Cell().Element(CellStyle).AlignCenter().Text(lastItem.SoQuyen ?? "");
                                lastRowTable.Cell().Element(CellStyle).AlignRight().Text(lastItem.ThuPhi?.ToString("N0") ?? "");
                                lastRowTable.Cell().Element(CellStyle).AlignRight().Text(lastItem.Huy?.ToString("N0") ?? "");
                                lastRowTable.Cell().Element(CellStyle).AlignRight().Text(lastItem.HoanTra?.ToString("N0") ?? "");
                                lastRowTable.Cell().Element(CellStyle).AlignCenter().Text(lastItem.HTTT ?? "");
                            });

                            group.Item().Table(footerTable =>
                            {
                                footerTable.ColumnsDefinition(columns =>
                                {
                                    columns.ConstantColumn(25);
                                    columns.ConstantColumn(50);
                                    columns.ConstantColumn(45);
                                    columns.ConstantColumn(50);
                                    columns.ConstantColumn(50);
                                    columns.RelativeColumn(2);
                                    columns.RelativeColumn(1);
                                    columns.RelativeColumn(1);
                                    columns.RelativeColumn(1);
                                    columns.ConstantColumn(40);
                                    columns.ConstantColumn(40);
                                    columns.ConstantColumn(35);
                                });

                                footerTable.Cell().Row(1).Column(1).ColumnSpan(8).Element(CellStyle).AlignCenter().Text("Tổng cộng").Bold();
                                footerTable.Cell().Row(1).Column(9).Element(CellStyle).AlignRight().Text(tongThuPhi.ToString("N0")).Bold();
                                footerTable.Cell().Row(1).Column(10).Element(CellStyle).AlignRight().Text(tongHuy.ToString("N0")).Bold();
                                footerTable.Cell().Row(1).Column(11).Element(CellStyle).AlignRight().Text(tongHoanTra.ToString("N0")).Bold();
                                footerTable.Cell().Row(1).Column(12).Element(CellStyle).Text("");
                            });

                            group.Item().Column(cuoi =>
                            {
                                cuoi.Item().Height(5);

                                cuoi.Spacing(5);
                                cuoi.Item().Text($"Số tiền phải nộp: {tongChenhLech:N0}").Bold();
                                cuoi.Item().Text($"Bằng chữ: {H0304NumberToTextHelper.ConvertSoThanhChu(tongChenhLech)}").Italic();

                                cuoi.Item().Row(row =>
                                {
                                    row.RelativeItem().Text("");
                                    row.ConstantItem(200).Column(right =>
                                    {
                                        right.Item().AlignCenter().Text($"Ngày {DateTime.Now:dd} Tháng {DateTime.Now:MM} Năm {DateTime.Now:yyyy}");
                                        right.Item().AlignCenter().Text("Người lập bảng");
                                        right.Item().Height(40);
                                        right.Item().AlignCenter().Text("Trần Thị Hồng Châu");
                                    });
                                });
                            });
                        });
                    });
                });

                page.Footer()
                    .AlignRight()
                    .Text(txt =>
                    {
                        txt.CurrentPageNumber();
                        txt.Span(" / ");
                        txt.TotalPages();
                    });
            });
        }

        static IContainer CellStyleHeader(IContainer container) =>
            container
                .Border(1)
                .Background(Colors.Grey.Lighten3)
                .Padding(3)
                .AlignMiddle()
                .DefaultTextStyle(x => x.SemiBold().FontSize(8));

        static IContainer CellStyle(IContainer container) =>
            container
                .Border(1)
                .Padding(3)
                .AlignMiddle()
                .DefaultTextStyle(x => x.FontSize(7));
    }
}
