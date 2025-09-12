using H0304.NumberToText.Helpers;
using M0304.Models.ThongTinDoanhNghiep;
using M0304H.Models.BCTongSoSIDTheoKhoaPhong;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace P0304H.PDFDocument
{
    public class P0304HReportTemplatePDF : IDocument
    {
        private readonly List<M0304HBCTongSoSIDTheoKhoaPhong> _data;
        private readonly M0304ThongTinDoanhNghiep _dataDN;
        private string _ngayBatDau;
        private string _ngayKetThuc;
        private readonly string _logoPath;

        public P0304HReportTemplatePDF(
            List<M0304HBCTongSoSIDTheoKhoaPhong> data,
            string ngayBatDau,
            string ngayKetThuc,
            M0304ThongTinDoanhNghiep dataDN,
            string logoPath = null
        )
        {
            _data = data ?? new List<M0304HBCTongSoSIDTheoKhoaPhong>();
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

                    col.Item().AlignCenter().PaddingVertical(10).Column(center =>
                    {
                        center.Item()
                            .AlignCenter()
                            .Text("BẢNG TỔNG KẾT XÉT NGHIỆM BỆNH NHÂN")
                            .Bold()
                            .FontSize(12);

                        center.Item()
                            .Width(250)
                            .AlignCenter()
                            .Text($"Từ ngày {_ngayBatDau:dd-MM-yyyy HH:mm:ss} đến ngày {_ngayKetThuc:dd-MM-yyyy HH:mm:ss}")
                            .FontSize(9);
                    });
                });

                page.Content().Column(col =>
                {
                    col.Item().Table(table =>
                    {
                        table.ColumnsDefinition(columns =>
                        {
                            columns.ConstantColumn(30);
                            columns.RelativeColumn();
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(40);
                            columns.ConstantColumn(40);
                            columns.ConstantColumn(40);
                            columns.ConstantColumn(40);
                            columns.ConstantColumn(40);
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(60);
                            columns.ConstantColumn(60);
                        });

                        table.Header(header =>
                        {
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("STT");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Tên khoa phòng");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Viện phí");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Bảo hiểm 100% (QL01)");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Bảo hiểm 100% (QL02)");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Bảo hiểm 95% (QL03)");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Bảo hiểm 80% (QL04)");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Bảo hiểm 100% (QL05)");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Dịch vụ");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Khám chuyên gia");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Tổng");
                        });

                        int stt = 1;
                        foreach (var item in _data)
                        {
                            table.Cell().Element(CellStyle).AlignCenter().Text(stt.ToString());
                            table.Cell().Element(CellStyle).AlignLeft().Text(item.TenKhoaPhong.Trim() ?? string.Empty);
                            table.Cell().Element(CellStyle).AlignRight().Text(item.VienPhi?.ToString("N0") ?? "0");
                            table.Cell().Element(CellStyle).AlignRight().Text(item.QL01?.ToString("N0") ?? "0");
                            table.Cell().Element(CellStyle).AlignRight().Text(item.QL02?.ToString("N0") ?? "0");
                            table.Cell().Element(CellStyle).AlignRight().Text(item.QL03?.ToString("N0") ?? "0");
                            table.Cell().Element(CellStyle).AlignRight().Text(item.QL04?.ToString("N0") ?? "0");
                            table.Cell().Element(CellStyle).AlignRight().Text(item.QL05?.ToString("N0") ?? "0");
                            table.Cell().Element(CellStyle).AlignRight().Text(item.DichVu?.ToString("N0") ?? "0");
                            table.Cell().Element(CellStyle).AlignRight().Text(item.KhamChuyenGia?.ToString("N0") ?? "0");
                            table.Cell().Element(CellStyle).AlignRight().Text(item.Tong?.ToString("N0") ?? "0");
                            stt++;
                        }

                        var tongVienPhi = _data.Sum(x => x.VienPhi ?? 0);
                        var tongQL01 = _data.Sum(x => x.QL01 ?? 0);
                        var tongQL02 = _data.Sum(x => x.QL02 ?? 0);
                        var tongQL03 = _data.Sum(x => x.QL03 ?? 0);
                        var tongQL04 = _data.Sum(x => x.QL04 ?? 0);
                        var tongQL05 = _data.Sum(x => x.QL05 ?? 0);
                        var tongDichVu = _data.Sum(x => x.DichVu ?? 0);
                        var tongKhamChuyenGia = _data.Sum(x => x.KhamChuyenGia ?? 0);
                        var tongToanBo = _data.Sum(x => x.Tong ?? 0);

                        table.Cell().ColumnSpan(2).Border(1).Element(CellTong).AlignCenter().Text("Tổng cộng").Bold();
                        table.Cell().Border(1).Element(CellTong).AlignRight().Text($"{tongVienPhi:N0}").Bold();
                        table.Cell().Border(1).Element(CellTong).AlignRight().Text($"{tongQL01:N0}").Bold();
                        table.Cell().Border(1).Element(CellTong).AlignRight().Text($"{tongQL02:N0}").Bold();
                        table.Cell().Border(1).Element(CellTong).AlignRight().Text($"{tongQL03:N0}").Bold();
                        table.Cell().Border(1).Element(CellTong).AlignRight().Text($"{tongQL04:N0}").Bold();
                        table.Cell().Border(1).Element(CellTong).AlignRight().Text($"{tongQL05:N0}").Bold();
                        table.Cell().Border(1).Element(CellTong).AlignRight().Text($"{tongDichVu:N0}").Bold();
                        table.Cell().Border(1).Element(CellTong).AlignRight().Text($"{tongKhamChuyenGia:N0}").Bold();
                        table.Cell().Border(1).Element(CellTong).AlignRight().Text($"{tongToanBo:N0}").Bold();
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
