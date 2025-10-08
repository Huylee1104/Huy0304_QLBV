using H0304.NumberToText.Helpers;
using M0304NhanVien.Models;
using M0304.Models.ToKhaiChiTietThuPhiLePhi;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.IO;
using M0304.Models.ThongTinDoanhNghiep;

namespace P0304.PDFDocument.ToKhaiChiTietThuPhiLePhi
{
    public class P0304ToKhaiChiTietThuPhiLePhiReportTemplate : IDocument
    {
        private readonly List<M0304ToKhaiChiTietThuPhiLePhi> _data;
        private readonly M0304ThongTinDoanhNghiep _dataDN;
        private string _ngayBatDau;
        private string _ngayKetThuc;
        private readonly string _logoPath;

        public P0304ToKhaiChiTietThuPhiLePhiReportTemplate(
            List<M0304ToKhaiChiTietThuPhiLePhi> data,
            string ngayBatDau,
            string ngayKetThuc,
            M0304ThongTinDoanhNghiep dataDN,
            string logoPath = null
        )
        {
            _data = data ?? new List<M0304ToKhaiChiTietThuPhiLePhi>();
            _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
            _ngayBatDau = ngayBatDau;
            _ngayKetThuc = ngayKetThuc;
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
                page.DefaultTextStyle(x => x.FontSize(9));

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
                                info.Item().Text(_dataDN.TenCoQuanChuyenMon ?? "").FontSize(8).Bold();
                                info.Item().Text(_dataDN.TenCSKCB ?? "").FontSize(8).Bold();
                            });
                        });
                    });

                    col.Item().AlignCenter().Column(center =>
                    {
                        center.Item()
                            .AlignCenter() // Căn giữa dòng chữ
                            .Text("TỜ KHAI CHI TIẾT THU PHÍ - LỆ PHÍ")
                            .Bold()
                            .FontSize(12);

                        center.Item()
                            .AlignCenter() // Đảm bảo căn giữa tuyệt đối
                            .Text($"Từ ngày {_ngayBatDau} đến ngày {_ngayKetThuc}")
                            .FontSize(9)
                            .Italic()
                            .Bold();
                    });
                });

                var groupedData = _data
                    .GroupBy(nv => new { nv.IDNhanVien, nv.TenNhanVien })
                    .Select(nvGroup => new
                    {
                        NhanVien = nvGroup.Key,
                        LoaiHoaDons = nvGroup
                            .GroupBy(hd => hd.LoaiHoaDon)
                            .Select(hdGroup => new
                            {
                                LoaiHoaDon = hdGroup.Key,
                                ChiTiet = hdGroup.ToList()
                            })
                            .ToList()
                    })
                    .ToList();

                page.Content().PaddingVertical(6).Column(col =>
                {
                    col.Item().Table(table =>
                    {
                        table.ColumnsDefinition(columns =>
                        {
                            columns.ConstantColumn(35);
                            columns.ConstantColumn(80);
                            columns.ConstantColumn(80);
                            columns.ConstantColumn(60);
                            columns.ConstantColumn(80);
                            columns.ConstantColumn(80);
                            columns.ConstantColumn(80);
                            columns.RelativeColumn();
                        });

                        table.Header(header =>
                        {
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("STT");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Đơn vị tính quyển");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Số lần hoặc số BL/HĐ thu");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Số HĐ sử dụng");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Tổng số tiền thu");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Hủy/Hoàn trả thu phí cho");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Số tiền thực thu");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Ghi chú");
                        });

                        if (groupedData != null && groupedData.Any())
                        {
                            foreach (var nvGroup in groupedData)
                            {
                                // ===== CẤP 1: NHÂN VIÊN =====
                                table.Cell().ColumnSpan(8)
                                    .Element(CellStyle)
                                    .AlignLeft()
                                    .Text($"{nvGroup.NhanVien.TenNhanVien?.ToUpper() ?? "KHÔNG RÕ NHÂN VIÊN"}")
                                    .FontSize(9)
                                    .Bold();

                                foreach (var loaiGroup in nvGroup.LoaiHoaDons)
                                {
                                    // ===== CẤP 2: LOẠI HÓA ĐƠN =====
                                    table.Cell().ColumnSpan(8)
                                        .Element(CellStyle)
                                        .AlignLeft()
                                        .Text($"    Loại hóa đơn: {loaiGroup.LoaiHoaDon}")
                                        .FontSize(9)
                                        .Bold();
                                    int stt = 1;
                                    // ===== CẤP 3: CHI TIẾT =====
                                    foreach (var item in loaiGroup.ChiTiet)
                                    {
                                        table.Cell().Element(CellStyle).AlignCenter().Text((stt++).ToString()).Bold();
                                        table.Cell().Element(CellStyle).AlignLeft().Text(item.QuyenSo ?? "");
                                        table.Cell().Element(CellStyle).AlignCenter().Text(item.SoLan_soBLHDthu ?? "");
                                        table.Cell().Element(CellStyle).AlignCenter().Text(item.SoLuongHDSuDung?.ToString("N0") ?? "").Bold();
                                        table.Cell().Element(CellStyle).AlignCenter().Text(item.TongSoTien?.ToString("N0") ?? "");
                                        table.Cell().Element(CellStyle).AlignCenter().Text(item.Huy_Hoan?.ToString("N0") ?? "");
                                        table.Cell().Element(CellStyle).AlignRight().Text(item.SoTienThucThu?.ToString("N0") ?? "");
                                        table.Cell().Element(CellStyle).AlignLeft().Text(item.GhiChu ?? "").Italic();
                                    }
                                }
                            }
                        }
                        var tongTien = _data.Sum(x => x.TongSoTien ?? 0);
                        var tongHuy = _data.Sum(x => x.Huy_Hoan ?? 0);
                        var tongThucThu = _data.Sum(x => x.SoTienThucThu ?? 0);

                        table.Cell().ColumnSpan(4).PaddingRight(3).AlignRight().Text("Tổng cộng:").Bold();
                        table.Cell().AlignCenter().PaddingRight(3).Text(tongTien.ToString("N0")).Bold();
                        table.Cell().AlignCenter().PaddingRight(3).Text(tongHuy.ToString("N0")).Bold();
                        table.Cell().AlignRight().PaddingRight(3).Text(tongThucThu.ToString("N0")).Bold();
                        table.Cell().AlignLeft().Text("");

                        table.Cell().ColumnSpan(8)
                            .AlignRight()
                            .PaddingTop(10)
                            .Element(container =>
                            {
                                container.Column(column =>
                                {
                                    column.Item()
                                        .AlignCenter()
                                        .Text($"Ngày {DateTime.Now:dd} Tháng {DateTime.Now:MM} Năm {DateTime.Now:yyyy}")
                                        .FontSize(9)
                                        .Italic()
                                        .Bold();

                                    column.Item().AlignCenter().PaddingTop(5)
                                        .Text("Người lập")
                                        .Bold()
                                        .FontSize(9);
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
                .Padding(3)
                .AlignMiddle()
                .DefaultTextStyle(x => x.SemiBold().FontSize(10));

        static IContainer CellStyle(IContainer container) =>
            container
                .Border(1)
                .Padding(3)
                .AlignMiddle()
                .DefaultTextStyle(x => x.FontSize(8));
    }
}
