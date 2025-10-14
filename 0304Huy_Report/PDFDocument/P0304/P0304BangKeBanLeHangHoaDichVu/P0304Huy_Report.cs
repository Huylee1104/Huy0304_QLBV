using DocumentFormat.OpenXml.EMMA;
using H0304.NumberToText.Helpers;
using M0304.Models.BangKeBanLeHangHoaDichVu;
using M0304.Models.ThongTinDoanhNghiep;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace P0304.PDFDocument.BangKeBanLeHangHoaDichVu
{
    public class P0304BangKeBanLeHangHoaDichVuReportTemplatePDF : IDocument
    {
        private readonly List<M0304BangKeBanLeHangHoaDichVu> _data;
        private readonly M0304ThongTinDoanhNghiep _dataDN;
        private string _ngayBatDau;
        private string _ngayKetThuc;
        private readonly string _logoPath;

        public P0304BangKeBanLeHangHoaDichVuReportTemplatePDF(
            List<M0304BangKeBanLeHangHoaDichVu> data,
            string ngayBatDau,
            string ngayKetThuc,
            M0304ThongTinDoanhNghiep dataDN,
            string logoPath = null
        )
        {
            _data = data ?? new List<M0304BangKeBanLeHangHoaDichVu>();
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

                            left.RelativeItem().Column(info =>
                            {
                                info.Item().Row(row =>
                                {
                                    row.RelativeItem().AlignLeft()
                                        .Text($"{_dataDN.TenCSKCB ?? ""}").FontSize(10);

                                    row.RelativeItem().AlignRight()
                                        .Text(text =>
                                        {
                                            text.Span("Mẫu số: ").FontSize(10);
                                            text.Span($"{_data[0].MauSo ?? ""}").FontSize(10);
                                        });
                                });
                                info.Item().Text(_data[0].TenKhoHang ?? "").FontSize(10).Bold();
                            });
                        });
                    });

                    col.Item().AlignCenter().Column(center =>
                    {
                        center.Item()
                            .AlignCenter()
                            .Text("BẢNG KÊ BÁN LẺ HÀNG HÓA, DỊCH VỤ")
                            .Bold()
                            .FontSize(12);

                        center.Item()
                            .Width(250)
                            .AlignCenter()
                            .Text($"Từ ngày {_ngayBatDau} đến ngày {_ngayKetThuc}")
                            .FontSize(9)
                            .Italic();
                    });
                    col.Item().AlignCenter().PaddingTop(5).Column(center =>
                    {
                        center.Item().Row(row =>
                        {
                            row.RelativeItem().AlignLeft()
                                .Text(text =>
                                {
                                    text.Span("Tên cơ sở kinh soanh: ").FontSize(10);
                                    text.Span($"{_dataDN.TenCSKCB ?? ""}").FontSize(10).Bold();
                                });
                            row.RelativeItem().AlignRight().Border(1).Padding(3)
                                .Text($"Mã số: {_data[0].MaSo ?? ""}").FontSize(10);

                        });
                    });
                    col.Item().AlignLeft().Column(center =>
                    {
                        center.Item()
                            .Text($"Địa chỉ: {_dataDN.DiaChi}")
                            .FontSize(10);
                    });
                    col.Item().AlignLeft().Column(center =>
                    {
                        center.Item()
                            .Text(text =>
                            {
                                text.Span("Họ tên người bán: ").FontSize(10);
                                text.Span($"{_data[0].TenNhanVien ?? ""}").FontSize(10).Bold();
                            });
                    });
                    col.Item().AlignLeft().Column(center =>
                    {
                        center.Item()
                            .Text($"Địa chỉ nơi bán: {_data[0].DiaChi}" ?? "")
                            .FontSize(10);
                    });
                });

                page.Content().PaddingVertical(6).Column(col =>
                {
                    col.Item().Table(table =>
                    {
                        table.ColumnsDefinition(columns =>
                        {
                            columns.ConstantColumn(30);
                            columns.RelativeColumn();
                            columns.ConstantColumn(60);
                            columns.ConstantColumn(80);
                            columns.ConstantColumn(100);
                            columns.ConstantColumn(120);
                        });

                        table.Header(header =>
                        {
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("STT");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Tên hàng hóa dịch vụ");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("ĐVT");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Số lượng");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Đơn vị bán");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Thành tiền bán");
                        });

                        int stt = 1;
                        foreach (var item in _data)
                        {
                            table.Cell().Element(CellStyle).AlignCenter().Text(stt.ToString()); 
                            table.Cell().Element(CellStyle).AlignLeft().Text(item.TenHangHoa ?? string.Empty);
                            table.Cell().Element(CellStyle).AlignCenter().Text(item.DVT ?? string.Empty);
                            table.Cell().Element(CellStyle).AlignCenter().Text(item.SoLuong?.ToString("N0") ?? "0");
                            table.Cell().Element(CellStyle).AlignRight().Text(item.DonGiaBan?.ToString("N0") ?? "0");  
                            table.Cell().Element(CellStyle).AlignRight().Text(item.ThanhTien?.ToString("N0") ?? "0");  
                            stt++;
                        }

                        double tongHoaDon = _data.Sum(x => x.ThanhTien ?? 0);

                        table.Cell().ColumnSpan(5).Border(1).Element(CellTong).AlignRight().Text("Tổng cộng").Bold();
                        table.Cell().Border(1).Element(CellTong).AlignRight().Text($"{tongHoaDon:N0}").Bold();
                        table.Cell().ColumnSpan(6).Border(1).Element(CellTong).Text(text =>
                        {
                            text.Span("Số tiền bằng chữ: ").Bold();
                            text.Span($"{H0304NumberToTextHelper.ConvertSoThanhChu((decimal)tongHoaDon)}").Bold();
                        });
                    });
                    col.Item().PaddingTop(10).Row(row =>
                    {
                        row.RelativeItem().AlignRight().Text(text =>
                        {
                            text.Span($"Ngày {DateTime.Now:dd} tháng {DateTime.Now:MM} năm {DateTime.Now:yyyy}")
                                .Italic().Bold();
                        });
                    });

                    col.Item().PaddingTop(10).Row(row =>
                    {
                        row.RelativeItem().AlignCenter().Text("BAN ĐIỀU HÀNH").Bold();
                        row.RelativeItem().AlignCenter().Text("THỦ QUỸ").Bold();
                        row.RelativeItem().AlignCenter().Text("NGƯỜI BÁN").Bold();
                        row.RelativeItem().AlignCenter().Text("KẾ TOÁN").Bold();
                    });

                });

                page.Footer()
                    .Row(row =>
                    {
                        // Bên trái: ngày giờ hiện tại
                        row.RelativeColumn()
                           .AlignLeft()
                           .Text($"Ngày in: {DateTime.Now:dd/MM/yyyy HH:mm}").Italic();

                        // Bên phải: số trang
                        row.RelativeColumn()
                           .AlignRight()
                           .Text(txt =>
                           {
                               txt.Span("Trang ");
                               txt.CurrentPageNumber();
                               txt.Span(" / ");
                               txt.TotalPages();
                           });
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
