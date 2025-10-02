using H0304.NumberToText.Helpers;
using M0304NhanVien.Models;
using M0304G.Models.PhieuLinhVatTuYTe;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.IO;
using M0304.Models.ThongTinDoanhNghiep;

namespace P0304F.PDFDocument
{
    public class P0304GReportTemplatePDF : IDocument
    {
        private readonly List<M0304GPhieuLinhVatTuYTe> _data;
        private readonly M0304ThongTinDoanhNghiep _dataDN;
        private readonly string _tenKho;
        private string _ngayBatDau;
        private string _ngayKetThuc;
        private readonly string _logoPath;

        public P0304GReportTemplatePDF(
            List<M0304GPhieuLinhVatTuYTe> data,
            M0304ThongTinDoanhNghiep dataDN,
            string tenKho,
            string ngayBatDau,
            string ngayKetThuc,
            string logoPath = null
        )
        {
            _data = data ?? new List<M0304GPhieuLinhVatTuYTe>();
            _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
            _tenKho = tenKho;
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
                                info.Item().Text(_dataDN.TenCoQuanChuyenMon ?? "").FontSize(8);
                                info.Item().Text(_dataDN.TenCSKCB ?? "").FontSize(8);
                                info.Item().Text(_dataDN.DiaChi ?? "").FontSize(8);
                                info.Item().Text(_dataDN.DienThoai ?? "").FontSize(8);
                            });
                        });
                    });

                    col.Item().AlignCenter().PaddingVertical(10).Column(center =>
                    {
                        center.Item()
                            .AlignCenter()
                            .Text("PHIẾU LĨNH VẬT TƯ Y TẾ")
                            .Bold()
                            .FontSize(12);

                        center.Item()
                            .Width(250)
                            .AlignCenter()
                            .Text($"Từ ngày {_ngayBatDau} đến ngày {_ngayKetThuc}")
                            .FontSize(8)
                            .Italic();
                    });
                    col.Item().AlignLeft().Column(left =>
                    {
                        left.Item()
                        .AlignLeft()
                        .Text(text =>
                        {
                            text.Span("Kho phát: ").FontSize(8);
                            text.Span($"{_tenKho}").Bold().FontSize(8);
                        });

                        left.Item()
                            .AlignLeft()
                            .Text("Diễn giải: nhu cầu sử dụng cho bệnh nhân.")
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
                            columns.ConstantColumn(70);
                            columns.RelativeColumn(4); 
                            columns.ConstantColumn(45);
                            columns.ConstantColumn(60);
                            columns.ConstantColumn(60);
                            columns.RelativeColumn(2); 
                        });

                        table.Cell().RowSpan(2).Element(CellStyleHeader).AlignCenter().Text("STT").Bold();
                        table.Cell().RowSpan(2).Element(CellStyleHeader).AlignCenter().Text("Mã").Bold();
                        table.Cell().RowSpan(2).Element(CellStyleHeader).AlignCenter().Text("Tên thuốc/VTYT/Hóa chất").Bold();
                        table.Cell().RowSpan(2).Element(CellStyleHeader).AlignCenter().Text("Đơn vị tính").Bold();
                        table.Cell().ColumnSpan(2).Element(CellStyleHeader).AlignCenter().Text("Số lượng").Bold();
                        table.Cell().RowSpan(2).Element(CellStyleHeader).AlignCenter().Text("Ghi chú").Bold();

                        // ===== Header dòng 2 =====
                        table.Cell().Element(CellStyleHeader).AlignCenter().Text("Yêu cầu").Bold();
                        table.Cell().Element(CellStyleHeader).AlignCenter().Text("Phát").Bold();

                        // ===== Dữ liệu =====
                        int stt = 1;
                        foreach (var item in _data)
                        {
                            table.Cell().Element(CellStyle).AlignCenter().Text(stt++.ToString());
                            table.Cell().Element(CellStyle).AlignCenter().Text(item.MaVatTu ?? "");
                            table.Cell().Element(CellStyle).AlignLeft().Text(item.TenVatTu ?? "");
                            table.Cell().Element(CellStyle).AlignCenter().Text(item.DonViTinh ?? "");
                            table.Cell().Element(CellStyle).AlignRight().Text(item.SoLuong?.ToString("N0") ?? "");
                            table.Cell().Element(CellStyle).AlignCenter().Text("");
                            table.Cell().Element(CellStyle).AlignCenter().Text("");
                        }
                    });

                    col.Item().EnsureSpace()
                    .Column(cuoi =>
                    {
                        cuoi.Item().PaddingTop(10).Row(row =>
                        {
                            row.RelativeItem().AlignLeft().Text($"Cộng khoản: {_data.Count} khoản").Bold();
                            row.RelativeItem().AlignRight().Text($"Ngày {DateTime.Now:dd} Tháng {DateTime.Now:MM} Năm {DateTime.Now:yyyy}").Italic();
                        });

                        cuoi.Item().PaddingTop(10).Table(table =>
                        {
                            table.ColumnsDefinition(columns =>
                            {
                                columns.RelativeColumn();
                                columns.RelativeColumn();
                                columns.RelativeColumn();
                                columns.RelativeColumn();
                                columns.RelativeColumn();
                            });

                            table.Cell().AlignCenter().Text("Người lập bảng").Bold();
                            table.Cell().AlignCenter().Text("Trưởng khoa Dược/VTYT\nngười được uỷ quyền").Bold();
                            table.Cell().AlignCenter().Text("Trưởng khoa/phòng").Bold();
                            table.Cell().AlignCenter().Text("Người giao").Bold();
                            table.Cell().AlignCenter().Text("Người nhận").Bold();

                            table.Cell().AlignCenter().Text("(Ký, ghi rõ họ tên)").Italic();
                            table.Cell().AlignCenter().Text("(Ký, ghi rõ họ tên)").Italic();
                            table.Cell().AlignCenter().Text("(Ký, ghi rõ họ tên)").Italic();
                            table.Cell().AlignCenter().Text("(Ký, ghi rõ họ tên)").Italic();
                            table.Cell().AlignCenter().Text("(Ký, ghi rõ họ tên)").Italic();
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
                .DefaultTextStyle(x => x.SemiBold().FontSize(8));

        static IContainer CellStyle(IContainer container) =>
            container
                .Border(1)
                .Padding(3)
                .AlignMiddle()
                .DefaultTextStyle(x => x.FontSize(7));
    }
}
