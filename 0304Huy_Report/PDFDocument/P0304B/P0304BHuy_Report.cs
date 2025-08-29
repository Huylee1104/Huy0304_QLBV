using H0304.NumberToText.Helpers;
using M0304.Models.ThongTinDOanhNghiep;
using M0304B.Models.BCHoaDonDienTuDV;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.IO;

namespace P0304.PDFDocument
{
    public class P0304BReportTemplatePDF : IDocument
    {
        private readonly List<M0304BBCHoaDonDienTuDV> _data;
        private readonly M0304ThongTinDoanhNghiep _dataDN;
        private string _ngayBatDau;
        private string _ngayKetThuc;
        private readonly string _logoPath;

        public P0304BReportTemplatePDF(
            List<M0304BBCHoaDonDienTuDV> data,
            string ngayBatDau,
            string ngayKetThuc,
            M0304ThongTinDoanhNghiep dataDN,
            string logoPath = null
        )
        {
            _data = data ?? new List<M0304BBCHoaDonDienTuDV>();
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
                page.Size(PageSizes.A4.Landscape());
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
                            .Text("BẢNG KÊ THU TIỀN NGOẠI TRÚ THEO BL/HĐ")
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
                            columns.ConstantColumn(55);
                            columns.ConstantColumn(60);
                            columns.ConstantColumn(65);
                            columns.ConstantColumn(45);
                            columns.RelativeColumn(1);
                            columns.ConstantColumn(30);
                            columns.RelativeColumn(2);
                            columns.ConstantColumn(60);
                            columns.ConstantColumn(45);
                            columns.ConstantColumn(65);
                            columns.ConstantColumn(45);
                        });

                        table.Header(header =>
                        {
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("STT");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Số chứng từ");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Ngày thu");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Giá trị");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Mã bệnh nhân");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Tên bệnh nhân");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Năm sinh");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Địa chỉ");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Ngày tạo HDDT");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("E_InvoiceNo");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Giá trị HDDT");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Mã truy cứu");
                        });

                        // Dòng số thứ tự cột
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

                        int stt = 1;
                        foreach (var item in _data)
                        {
                            table.Cell().Element(CellStyle).AlignCenter().Text(stt.ToString()); 
                            table.Cell().Element(CellStyle).AlignCenter().Text(item.SoChungTu ?? string.Empty);      
                            table.Cell().Element(CellStyle).AlignCenter().Text(item.NgayThu?.ToString("dd-MM-yyyy"));
                            table.Cell().Element(CellStyle).AlignRight().Text(item.GiaTri?.ToString("N0") ?? "0");  
                            table.Cell().Element(CellStyle).AlignCenter().Text(item.MaBenhNhan ?? string.Empty);    
                            table.Cell().Element(CellStyle).AlignLeft().Text(item.TenBenhNhan ?? string.Empty);    
                            table.Cell().Element(CellStyle).AlignCenter().Text(item.NamSinh?.ToString() ?? "");    
                            table.Cell().Element(CellStyle).AlignLeft().Text(item.DiaChi ?? string.Empty);           
                            table.Cell().Element(CellStyle).AlignCenter().Text(item.NgayTaoHDDT?.ToString("dd-MM-yyyy")); 
                            table.Cell().Element(CellStyle).AlignCenter().Text(item.E_InvoiceNo?.ToString() ?? "");  
                            table.Cell().Element(CellStyle).AlignRight().Text(item.GiaTriHDDT?.ToString("N0") ?? "0"); 
                            table.Cell().Element(CellStyle).AlignCenter().Text(item.MaTraCuu?.ToString() ?? "");     

                            stt++;
                        }

                        // Tính tổng
                        var tongGiaTri = _data.Sum(x => x.GiaTri ?? 0);
                        var tongGiaTriHDDT = _data.Sum(x => x.GiaTriHDDT ?? 0);

                        table.Cell().ColumnSpan(12).Border(1).Element(cell =>
                        {
                            cell.Row(row =>
                            {
                                CellTong(row.ConstantItem(30)).AlignCenter().Text("");
                                CellTong(row.ConstantItem(55)).AlignCenter().Text("");
                                CellTong(row.ConstantItem(60)).AlignCenter().Text("");
                                CellTong(row.ConstantItem(65)).AlignRight().Text($"{tongGiaTri:N0}");
                                CellTong(row.ConstantItem(45)).AlignCenter().Text("");
                                CellTong(row.RelativeItem(1)).AlignCenter().Text("");
                                CellTong(row.ConstantItem(30)).AlignCenter().Text("");
                                CellTong(row.RelativeItem(2)).AlignCenter().Text("");
                                CellTong(row.ConstantItem(60)).AlignCenter().Text("");
                                CellTong(row.ConstantItem(45)).AlignCenter().Text("");
                                CellTong(row.ConstantItem(65)).AlignRight().Text($"{tongGiaTriHDDT:N0}");
                                CellTong(row.ConstantItem(45)).AlignCenter().Text("");
                            });
                        });

                        col.Item().Height(10);
                        col.Item().EnsureSpace()
                            .Column(cuoi =>
                            {
                                cuoi.Item().Row(row =>
                                {
                                    row.RelativeItem().Text("");
                                    row.ConstantItem(200).Column(right =>
                                    {
                                        right.Item().AlignCenter().Text($"Ngày {DateTime.Now:dd} Tháng {DateTime.Now:MM} Năm {DateTime.Now:yyyy}");
                                        right.Item().AlignCenter().Text("Người lập bảng").Bold();
                                        right.Item().Height(40);
                                        right.Item().AlignCenter().Text("Trần Thị Hồng Châu");
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
