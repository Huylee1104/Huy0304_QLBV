using H0304.NumberToText.Helpers;
using M0304NhanVien.Models;
using M0304.Models.BaoCaoTiepNhan;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.IO;
using M0304.Models.ThongTinDoanhNghiep;

namespace P0304.PDFDocument.BaoCaoTiepNhan
{
    public class P0304BaoCaoTiepNhanReportTemplate : IDocument
    {
        private readonly List<M0304BaoCaoTiepNhan> _data;
        private readonly M0304ThongTinDoanhNghiep _dataDN;
        private string _ngayBatDau;
        private string _ngayKetThuc;
        private readonly string _logoPath;

        public P0304BaoCaoTiepNhanReportTemplate(
            List<M0304BaoCaoTiepNhan> data,
            M0304ThongTinDoanhNghiep dataDN,
            string ngayBatDau,
            string ngayKetThuc,
            string logoPath = null
        )
        {
            _data = data ?? new List<M0304BaoCaoTiepNhan>();
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
                                left.ConstantItem(40).AlignMiddle().Text("");
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
                            .AlignCenter()
                            .Text("BÁO CÁO TIẾP NHẬN")
                            .Bold()
                            .FontSize(12);

                        center.Item()
                            .Width(250)
                            .AlignCenter()
                            .Text($"Từ ngày {_ngayBatDau} đến ngày {_ngayKetThuc}")
                            .FontSize(8).Bold().Italic();
                    });
                });

                page.Content().PaddingVertical(6).Column(col =>
                {
                    col.Item().Table(table =>
                    {
                        table.ColumnsDefinition(columns =>
                        {
                            columns.ConstantColumn(40);
                            columns.RelativeColumn();
                            columns.ConstantColumn(100);
                            columns.ConstantColumn(80);
                            columns.ConstantColumn(80);
                            columns.ConstantColumn(80);
                            columns.ConstantColumn(80);
                        });


                        table.Header(header =>
                        {
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("STT");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Tên phòng bệnh");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Số lượt tiếp nhận");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Nam");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Nữ");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Có bảo hiểm");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Không BHYT");
                        });
                        var group = _data.GroupBy(x => new {x.IdKhoa, x.TenKhoa}).Select(g => new
                        {
                            TenKhoa = g.Key.TenKhoa,
                            Items = g.ToList()
                        }).ToList();

                        int stt = 1;

                        var tongLuot = _data.Sum(x => x.SoLuotTiepNhan ?? 0);
                        var tongNam = _data.Sum(x => x.SoLuongNam ?? 0);
                        var tongNu = _data.Sum(x => x.SoLuongNu ?? 0);
                        var tongBHYT = _data.Sum(x => x.CoBHYT ?? 0);
                        var tongKhongBHYT = _data.Sum(x => x.KhongBHYT ?? 0);

                        foreach (var khoa in group)
                        {
                            var tongLuotKhoa = 0;
                            var tongNamKhoa = 0;
                            var tongNuKhoa = 0;
                            var tongBHYTKhoa = 0;
                            var tongKhongBHYTKhoa = 0;
                            table.Cell().ColumnSpan(7).Border(1).Element(CellStyle).AlignLeft().Text(khoa.TenKhoa).Bold();

                            foreach (var item in khoa.Items)
                            {
                                table.Cell().Element(CellStyle).AlignCenter().Text(stt.ToString());
                                table.Cell().Element(CellStyle).AlignLeft().Text(item.TenPhongBenh ?? string.Empty);
                                table.Cell().Element(CellStyle).AlignRight().Text(item.SoLuotTiepNhan?.ToString("N0") ?? "0"); tongLuotKhoa+= item.SoLuotTiepNhan ?? 0;
                                table.Cell().Element(CellStyle).AlignRight().Text(item.SoLuongNam?.ToString("N0") ?? "0"); tongNamKhoa+= item.SoLuongNam ?? 0;
                                table.Cell().Element(CellStyle).AlignRight().Text(item.SoLuongNu?.ToString("N0") ?? "0"); tongNuKhoa+= item.SoLuongNu ?? 0;
                                table.Cell().Element(CellStyle).AlignRight().Text(item.CoBHYT?.ToString("N0") ?? "0"); tongBHYTKhoa+= item.CoBHYT ?? 0;
                                table.Cell().Element(CellStyle).AlignRight().Text(item.KhongBHYT?.ToString("N0") ?? "0"); tongKhongBHYTKhoa+= item.KhongBHYT ?? 0;
                                stt++;
                            }

                            table.Cell().ColumnSpan(2).Border(1).Element(CellStyle).AlignLeft().Text("").Bold();
                            table.Cell().Border(1).Element(CellStyle).AlignRight().Text(tongLuotKhoa.ToString("N0")).Bold();
                            table.Cell().Border(1).Element(CellStyle).AlignRight().Text(tongNamKhoa.ToString("N0")).Bold();
                            table.Cell().Border(1).Element(CellStyle).AlignRight().Text(tongNuKhoa.ToString("N0")).Bold();
                            table.Cell().Border(1).Element(CellStyle).AlignRight().Text(tongBHYTKhoa.ToString("N0")).Bold();
                            table.Cell().Border(1).Element(CellStyle).AlignRight().Text(tongKhongBHYTKhoa.ToString("N0")).Bold();
                        }

                        table.Cell().ColumnSpan(2).Border(1).Element(CellStyle).AlignRight().Text("").Bold();
                        table.Cell().Border(1).Element(CellStyle).AlignRight().Text($"{tongLuot:N0}").Bold();
                        table.Cell().Border(1).Element(CellStyle).AlignRight().Text($"{tongNam:N0}").Bold();
                        table.Cell().Border(1).Element(CellStyle).AlignRight().Text($"{tongNu:N0}").Bold();
                        table.Cell().Border(1).Element(CellStyle).AlignRight().Text($"{tongBHYT:N0}").Bold();
                        table.Cell().Border(1).Element(CellStyle).AlignRight().Text($"{tongKhongBHYT:N0}").Bold();
                    });
                });

                page.Footer()
                    .AlignRight()
                    .Text(txt =>
                    {
                        txt.Span("Trang ");
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
