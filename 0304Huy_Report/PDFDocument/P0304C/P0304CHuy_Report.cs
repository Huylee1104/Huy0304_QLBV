using H0304.NumberToText.Helpers;
using M0304.Models.ThongTinDoanhNghiep;
using M0304C.Models.BaoCaoThuDichVu;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace P0304.PDFDocument
{
    public class P0304CReportTemplatePDF : IDocument
    {
        private readonly List<M0304CBaoCaoThuDichVu> _data;
        private readonly M0304ThongTinDoanhNghiep _dataDN;
        private string _ngayBatDau;
        private string _ngayKetThuc;
        private readonly string _logoPath;

        public P0304CReportTemplatePDF(
            List<M0304CBaoCaoThuDichVu> data,
            string ngayBatDau,
            string ngayKetThuc,
            M0304ThongTinDoanhNghiep dataDN,
            string logoPath = null
        )
        {
            _data = data ?? new List<M0304CBaoCaoThuDichVu>();
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
                page.DefaultTextStyle(x => x.FontSize(10));

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
                                info.Item().Text(_dataDN.TenCSKCB ?? "").FontSize(9);
                                info.Item().Text(_dataDN.TenCoQuanChuyenMon ?? "").FontSize(9);
                                info.Item().Text(_dataDN.DiaChi ?? "").FontSize(9);
                                info.Item().Text(_dataDN.DienThoai ?? "").FontSize(9);
                            });
                        });
                    });

                    col.Item().AlignCenter().Column(center =>
                    {
                        center.Item()
                            .AlignCenter()
                            .Text("BÁO CÁO THU DỊCH VỤ")
                            .Bold()
                            .FontSize(12);

                        center.Item()
                            .Width(250)
                            .AlignCenter()
                            .Text($"Từ ngày {_ngayBatDau} đến ngày {_ngayKetThuc}")
                            .FontSize(9);
                    });
                });

                page.Content().PaddingVertical(6).Column(col =>
                {
                    col.Item().Table(table =>
                    {
                        table.ColumnsDefinition(columns =>
                        {
                            columns.ConstantColumn(30);
                            columns.ConstantColumn(150);
                            columns.RelativeColumn();
                            columns.ConstantColumn(60);
                            columns.ConstantColumn(80);
                        });

                        table.Header(header =>
                        {
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("STT");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Nhóm dịch vụ");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Dịch vụ");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Số lượng");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Tổng hóa đơn");
                        });

                        int stt = 1;
                        foreach (var item in _data)
                        {
                            table.Cell().Element(CellStyle).AlignCenter().Text(stt.ToString()); 
                            table.Cell().Element(CellStyle).AlignLeft().Text(item.NhomDichVu ?? string.Empty);
                            table.Cell().Element(CellStyle).AlignLeft().Text(item.DichVu ?? string.Empty);
                            table.Cell().Element(CellStyle).AlignCenter().Text(item.SoLuong?.ToString("N0") ?? "0");
                            table.Cell().Element(CellStyle).AlignRight().Text(item.TongHoaDon?.ToString("N0") ?? "0");  
                            stt++;
                        }

                        var tongHoaDon = _data.Sum(x => x.TongHoaDon ?? 0);

                        table.Cell().ColumnSpan(3).Border(1).Element(CellTong).AlignCenter().Text("Tổng cộng").Bold();

                        table.Cell().Border(1).Element(CellTong).Text("");

                        table.Cell().Border(1).Element(CellTong).AlignRight().Text($"{tongHoaDon:N0}").Bold(); 
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
                .Padding(4)
                .AlignMiddle()
                .DefaultTextStyle(x => x.SemiBold().FontSize(10));

        static IContainer CellStyle(IContainer container) =>
            container
                .Border(1)
                .Padding(4)
                .AlignMiddle()
                .DefaultTextStyle(x => x.FontSize(9));
        static IContainer CellTong(IContainer container) =>
            container
                .Padding(4)
                .AlignMiddle()
                .DefaultTextStyle(x => x.FontSize(9));
    }
}
