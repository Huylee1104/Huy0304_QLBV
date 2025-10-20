using H0304.NumberToText.Helpers;
using M0304NhanVien.Models;
using M0304.Models.HoatDongKhamBenh;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.IO;
using M0304.Models.ThongTinDoanhNghiep;

namespace P0304.PDFDocument.HoatDongKhamBenh
{
    public class P0304HoatDongKhamBenhReportTemplate : IDocument
    {
        private readonly List<M0304HoatDongKhamBenh> _data;
        private readonly M0304ThongTinDoanhNghiep _dataDN;
        private string _ngayBatDau;
        private string _ngayKetThuc;
        private string _tenNVDN;
        private readonly string _logoPath;

        public P0304HoatDongKhamBenhReportTemplate(
            List<M0304HoatDongKhamBenh> data,
            string ngayBatDau,
            string ngayKetThuc,
            string tenNVDN,
            M0304ThongTinDoanhNghiep dataDN,
            string logoPath = null
        )
        {
            _data = data ?? new List<M0304HoatDongKhamBenh>();
            _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
            _ngayBatDau = ngayBatDau;
            _ngayKetThuc = ngayKetThuc;
            _tenNVDN = tenNVDN;
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
                            .Text("HOẠT ĐỘNG KHÁM BỆNH")
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

                        table.Cell().AlignLeft().Text("");

                        table.Cell().ColumnSpan(8)
                        .AlignRight()
                        .PaddingTop(10)
                        .Element(container =>
                        {
                            container.AlignRight().Column(column =>
                            {
                                column.Item()
                                    .AlignCenter()
                                    .Text($"Ngày {DateTime.Now:dd} Tháng {DateTime.Now:MM} Năm {DateTime.Now:yyyy}")
                                    .FontSize(9)
                                    .Italic()
                                    .Bold();

                                column.Item()
                                    .AlignCenter()
                                    .Text("Người lập bảng")
                                    .Bold();

                                column.Item()
                                    .Height(40);

                                column.Item()
                                    .AlignCenter()
                                    .Text($"{_tenNVDN}")
                                    .Bold();
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
