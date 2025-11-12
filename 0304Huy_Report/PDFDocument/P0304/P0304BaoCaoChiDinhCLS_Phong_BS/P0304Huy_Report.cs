using H0304.NumberToText.Helpers;
using M0304.Models.ThongTinDoanhNghiep;
using M0304.Models.BaoCaoChiDinhCLS_Phong_BS;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace P0304.PDFDocument.BaoCaoChiDinhCLS_Phong_BS
{
    public class P0304BaoCaoChiDinhCLS_Phong_BSReportTemplate : IDocument
    {
        private readonly List<M0304BaoCaoChiDinhCLS_Phong_BS> _data;
        private readonly M0304ThongTinDoanhNghiep _dataDN;
        private string _ngayBatDau;
        private string _ngayKetThuc;
        private readonly string _logoPath;

        public P0304BaoCaoChiDinhCLS_Phong_BSReportTemplate(
            List<M0304BaoCaoChiDinhCLS_Phong_BS> data,
            string ngayBatDau,
            string ngayKetThuc,
            M0304ThongTinDoanhNghiep dataDN,
            string logoPath = null
        )
        {
            _data = data ?? new List<M0304BaoCaoChiDinhCLS_Phong_BS>();
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
                                info.Item().Text(_dataDN.TenCoQuanChuyenMon ?? "").FontSize(8).Bold();
                                info.Item().Text(_dataDN.TenCSKCB ?? "").FontSize(8).Bold();
                                info.Item().Text(_dataDN.DiaChi ?? "").FontSize(8).Bold();
                                info.Item().Text(_dataDN.DienThoai ?? "").FontSize(8).Bold();
                            });
                        });
                    });

                    col.Item().AlignCenter().Column(center =>
                    {
                        center.Item()
                            .AlignCenter()
                            .Text("BÁO CÁO TỔNG HỢP CẬN LÂM SÀN")
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
                            columns.ConstantColumn(120);
                            columns.ConstantColumn(120);
                            columns.RelativeColumn();
                            columns.ConstantColumn(100);
                            columns.ConstantColumn(80);
                        });

                        table.Header(header =>
                        {
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("STT");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Nơi gửi (phòng khám)");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Bác sĩ");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Yêu cầu");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Đơn giá");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Tổng lượt");

                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("1");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("2");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("3");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("4");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("5");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("6");
                        });

                        int stt = 1;
                        foreach (var item in _data)
                        {
                            table.Cell().Element(CellStyle).AlignCenter().Text(stt.ToString()); 
                            table.Cell().Element(CellStyle).AlignLeft().Text(item.NoiGui ?? string.Empty);
                            table.Cell().Element(CellStyle).AlignLeft().Text(item.BacSi ?? string.Empty);
                            table.Cell().Element(CellStyle).AlignLeft().Text(item.YeuCau ?? string.Empty);
                            table.Cell().Element(CellStyle).AlignRight().Text(item.DonGia?.ToString("N2") ?? "0");  
                            table.Cell().Element(CellStyle).AlignRight().Text(item.SoLuot?.ToString("N0") ?? "0");  
                            stt++;
                        }

                        var TongLuot = _data.Sum(x => x.SoLuot ?? 0);
                        var TongDonGia = _data.Sum(x => x.DonGia ?? 0);

                        table.Cell().ColumnSpan(4).Border(1).Element(CellTong).AlignCenter().Text("Tổng cộng").Bold();
                         
                        table.Cell().Border(1).Element(CellTong).AlignRight().Text($"{TongDonGia:N2}").Bold(); 
                        table.Cell().Border(1).Element(CellTong).AlignRight().Text($"{TongLuot:N0}").Bold(); 
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
