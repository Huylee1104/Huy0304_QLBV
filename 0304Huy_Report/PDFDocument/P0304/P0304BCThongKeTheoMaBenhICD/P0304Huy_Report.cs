using H0304.NumberToText.Helpers;
using M0304.Models.ThongTinDoanhNghiep;
using M0304.Models.BCThongKeTheoMaBenhICD;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace P0304.PDFDocument.BCThongKeTheoMaBenhICD
{
    public class P0304BCThongKeTheoMaBenhICDReportTemplate : IDocument
    {
        private readonly List<M0304BCThongKeTheoMaBenhICD> _data;
        private readonly M0304ThongTinDoanhNghiep _dataDN;
        private string _ngayBatDau;
        private string _ngayKetThuc;
        private readonly string _logoPath;

        public P0304BCThongKeTheoMaBenhICDReportTemplate(
            List<M0304BCThongKeTheoMaBenhICD> data,
            string ngayBatDau,
            string ngayKetThuc,
            M0304ThongTinDoanhNghiep dataDN,
            string logoPath = null
        )
        {
            _data = data ?? new List<M0304BCThongKeTheoMaBenhICD>();
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
                            .Text("BÁO CÁO TỔNG HỢP SỐ LIỆU KHÁM BỆNH THEO NHIỀU TIÊU CHÍ")
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
                            columns.ConstantColumn(50);
                            columns.RelativeColumn();
                            columns.ConstantColumn(80);
                            columns.ConstantColumn(60);
                            columns.ConstantColumn(60);
                            columns.ConstantColumn(80);
                            columns.ConstantColumn(80);
                        });

                        table.Header(header =>
                        {
                            header.Cell().Row(1).Column(1).Element(CellStyleHeader).AlignCenter().Text("STT");
                            header.Cell().Row(1).Column(2).Element(CellStyleHeader).AlignCenter().Text("ICD");
                            header.Cell().Row(1).Column(3).Element(CellStyleHeader).AlignCenter().Text("Tên bệnh");
                            header.Cell().Row(1).Column(4).Element(CellStyleHeader).AlignCenter().Text("Tổng số");

                            header.Cell().Row(1).Column(5).ColumnSpan(2).Element(CellStyleHeader).AlignCenter().Text("Giới tính");

                            header.Cell().Row(1).Column(7).Element(CellStyleHeader).AlignCenter().Text("Có BHYT");
                            header.Cell().Row(1).Column(8).Element(CellStyleHeader).AlignCenter().Text("Không BHYT");

                            header.Cell().Row(2).Column(5).Element(CellStyleHeader).AlignCenter().Text("Nam");
                            header.Cell().Row(2).Column(6).Element(CellStyleHeader).AlignCenter().Text("Nữ");
                        });

                        int stt = 1;
                        foreach (var item in _data)
                        {
                            table.Cell().Element(CellStyle).AlignCenter().Text(stt.ToString()); 
                            table.Cell().Element(CellStyle).AlignLeft().Text(item.TenICD ?? string.Empty);
                            table.Cell().Element(CellStyle).AlignLeft().Text(item.TenBenh ?? string.Empty);
                            table.Cell().Element(CellStyle).AlignCenter().Text(item.SoLuotTiepNhan?.ToString("N0") ?? "0");
                            table.Cell().Element(CellStyle).AlignRight().Text(item.SoLuongNam?.ToString("N0") ?? "0");  
                            table.Cell().Element(CellStyle).AlignRight().Text(item.SoLuongNu?.ToString("N0") ?? "0");  
                            table.Cell().Element(CellStyle).AlignRight().Text(item.CoBHYT?.ToString("N0") ?? "0");  
                            table.Cell().Element(CellStyle).AlignRight().Text(item.KhongBHYT?.ToString("N0") ?? "0");  
                            stt++;
                        }

                        var TongLuot = _data.Sum(x => x.SoLuotTiepNhan ?? 0);
                        var TongNam = _data.Sum(x => x.SoLuongNam ?? 0);
                        var TongNu = _data.Sum(x => x.SoLuongNu ?? 0);
                        var TongBHYT = _data.Sum(x => x.CoBHYT ?? 0);
                        var TongKhongBHYT = _data.Sum(x => x.KhongBHYT ?? 0);

                        table.Cell().ColumnSpan(3).Border(1).Element(CellTong).AlignCenter().Text("Tổng cộng").Bold();
                         
                        table.Cell().Border(1).Element(CellTong).AlignRight().Text($"{TongLuot:N0}").Bold(); 
                        table.Cell().Border(1).Element(CellTong).AlignRight().Text($"{TongNam:N0}").Bold(); 
                        table.Cell().Border(1).Element(CellTong).AlignRight().Text($"{TongNu:N0}").Bold(); 
                        table.Cell().Border(1).Element(CellTong).AlignRight().Text($"{TongBHYT:N0}").Bold(); 
                        table.Cell().Border(1).Element(CellTong).AlignRight().Text($"{TongKhongBHYT:N0}").Bold(); 
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
