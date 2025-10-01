using H0304.NumberToText.Helpers;
using M0304M.Models.BaoCaoHangHoa;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.IO;
using M0304.Models.ThongTinDoanhNghiep;
using DocumentFormat.OpenXml.Bibliography;

namespace P0304M.PDFDocument
{
    public class P0304MReportNhapTemplatePDF : IDocument
    {
        private readonly List<M0304MHangNhap> _data;
        private readonly M0304ThongTinDoanhNghiep _dataDN;
        private int _nam;

        public P0304MReportNhapTemplatePDF(
            List<M0304MHangNhap> data,
            int nam,
            M0304ThongTinDoanhNghiep dataDN
        )
        {
            _data = data ?? new List<M0304MHangNhap>();
            _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
            _nam = nam;
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
                            .Text($"BÁO CÁO SỐ LƯỢNG HÀNG NHẬP NĂM {_nam}")
                            .Bold()
                            .FontSize(12);
                    });
                });

                page.Content().PaddingVertical(6).Column(col =>
                {
                    col.Item().Table(table =>
                    {
                        table.ColumnsDefinition(columns =>
                        {
                            columns.ConstantColumn(30);
                            columns.RelativeColumn(3);
                            for (int i = 0; i < 13; i++)
                                columns.RelativeColumn();
                        });

                        table.Header(header =>
                        {
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("STT");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Tên thuốc");
                            for (int i = 1; i <= 12; i++)
                                header.Cell().Element(CellStyleHeader).AlignCenter().Text($"Tháng {i}");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Tổng cộng");
                        });

                        int stt = 1;
                        var _dsNhomHang = _data
                        .GroupBy(x => new { x.IDNhomHang, x.TenNhomHang })
                        .Select(g => new
                        {
                            IDNH = g.Key.IDNhomHang,
                            TenNH = g.Key.TenNhomHang
                        }).ToList();

                        foreach (var nhomHang in _dsNhomHang)
                        {
                            var data = _data.Where(d => d.IDNhomHang == nhomHang.IDNH).ToList();
                            if (!data.Any()) continue;

                            int tongThang1 = (int)data.Sum(x => x.Thang1 ?? 0);
                            int tongThang2 = (int)data.Sum(x => x.Thang2 ?? 0);
                            int tongThang3 = (int)data.Sum(x => x.Thang3 ?? 0);
                            int tongThang4 = (int)data.Sum(x => x.Thang4 ?? 0);
                            int tongThang5 = (int)data.Sum(x => x.Thang5 ?? 0);
                            int tongThang6 = (int)data.Sum(x => x.Thang6 ?? 0);
                            int tongThang7 = (int)data.Sum(x => x.Thang7 ?? 0);
                            int tongThang8 = (int)data.Sum(x => x.Thang8 ?? 0);
                            int tongThang9 = (int)data.Sum(x => x.Thang9 ?? 0);
                            int tongThang10 = (int)data.Sum(x => x.Thang10 ?? 0);
                            int tongThang11 = (int)data.Sum(x => x.Thang11 ?? 0);
                            int tongThang12 = (int)data.Sum(x => x.Thang12 ?? 0);
                            int tongCong = (int)data.Sum(x => x.TongCong ?? 0);

                            // Dòng nhóm: merge 15 cột
                            table.Cell().ColumnSpan(2).Element(cell =>
                                CellStyle(cell)
                                    .AlignLeft()
                                    .AlignMiddle()
                                    .Text($"{nhomHang.TenNH}")
                                    .SemiBold()
                            );

                            // Điền tổng vào các cột Tháng 1 → 12 + Tổng cộng
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang1.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang2.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang3.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang4.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang5.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang6.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang7.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang8.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang9.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang10.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang11.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang12.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongCong.ToString("N0")).SemiBold();

                            foreach (var item in data)
                            {
                                table.Cell().Element(CellStyle).AlignCenter().Text(stt++.ToString());
                                table.Cell().Element(CellStyle).AlignLeft().Text(item.TenThuoc ?? "");
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang1 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang2 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang3 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang4 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang5 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang6 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang7 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang8 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang9 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang10 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang11 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang12 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.TongCong ?? 0)).ToString("N0"));
                            }
                        }
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
                .Padding(2)
                .AlignMiddle()
                .DefaultTextStyle(x => x.SemiBold().FontSize(9));

        static IContainer CellStyle(IContainer container) =>
            container
                .Border(1)
                .Padding(3)
                .AlignMiddle()
                .DefaultTextStyle(x => x.FontSize(8));
    }

    public class P0304MReportXuatTemplatePDF : IDocument
    {
        private readonly List<M0304MHangXuat> _data;
        private readonly M0304ThongTinDoanhNghiep _dataDN;
        private int _nam;

        public P0304MReportXuatTemplatePDF(
            List<M0304MHangXuat> data,
            int nam,
            M0304ThongTinDoanhNghiep dataDN
        )
        {
            _data = data ?? new List<M0304MHangXuat>();
            _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
            _nam = nam;
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
                            .Text($"BÁO CÁO SỐ LƯỢNG HÀNG XUẤT NĂM {_nam}")
                            .Bold()
                            .FontSize(12);
                    });
                });

                page.Content().PaddingVertical(6).Column(col =>
                {
                    col.Item().Table(table =>
                    {
                        table.ColumnsDefinition(columns =>
                        {
                            columns.ConstantColumn(30);
                            columns.RelativeColumn(3);
                            for (int i = 0; i < 13; i++)
                                columns.RelativeColumn();
                        });

                        table.Header(header =>
                        {
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("STT");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Tên thuốc");
                            for (int i = 1; i <= 12; i++)
                                header.Cell().Element(CellStyleHeader).AlignCenter().Text($"Tháng {i}");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Tổng cộng");
                        });

                        int stt = 1;
                        var _dsNhomHang = _data
                        .GroupBy(x => new { x.IDNhomHang, x.TenNhomHang })
                        .Select(g => new
                        {
                            IDNH = g.Key.IDNhomHang,
                            TenNH = g.Key.TenNhomHang
                        }).ToList();

                        foreach (var nhomHang in _dsNhomHang)
                        {
                            var data = _data.Where(d => d.IDNhomHang == nhomHang.IDNH).ToList();
                            if (!data.Any()) continue;

                            int tongThang1 = (int)data.Sum(x => x.Thang1 ?? 0);
                            int tongThang2 = (int)data.Sum(x => x.Thang2 ?? 0);
                            int tongThang3 = (int)data.Sum(x => x.Thang3 ?? 0);
                            int tongThang4 = (int)data.Sum(x => x.Thang4 ?? 0);
                            int tongThang5 = (int)data.Sum(x => x.Thang5 ?? 0);
                            int tongThang6 = (int)data.Sum(x => x.Thang6 ?? 0);
                            int tongThang7 = (int)data.Sum(x => x.Thang7 ?? 0);
                            int tongThang8 = (int)data.Sum(x => x.Thang8 ?? 0);
                            int tongThang9 = (int)data.Sum(x => x.Thang9 ?? 0);
                            int tongThang10 = (int)data.Sum(x => x.Thang10 ?? 0);
                            int tongThang11 = (int)data.Sum(x => x.Thang11 ?? 0);
                            int tongThang12 = (int)data.Sum(x => x.Thang12 ?? 0);
                            int tongCong = (int)data.Sum(x => x.TongCong ?? 0);

                            // Dòng nhóm: merge 15 cột
                            table.Cell().ColumnSpan(2).Element(cell =>
                                CellStyle(cell)
                                    .AlignLeft()
                                    .AlignMiddle()
                                    .Text($"{nhomHang.TenNH}")
                                    .SemiBold()
                            );

                            // Điền tổng vào các cột Tháng 1 → 12 + Tổng cộng
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang1.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang2.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang3.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang4.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang5.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang6.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang7.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang8.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang9.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang10.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang11.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongThang12.ToString("N0")).SemiBold();
                            table.Cell().Element(CellStyle).AlignRight().Text(tongCong.ToString("N0")).SemiBold();

                            foreach (var item in data)
                            {
                                table.Cell().Element(CellStyle).AlignCenter().Text(stt++.ToString());
                                table.Cell().Element(CellStyle).AlignLeft().Text(item.TenThuoc ?? "");
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang1 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang2 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang3 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang4 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang5 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang6 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang7 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang8 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang9 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang10 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang11 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.Thang12 ?? 0)).ToString("N0"));
                                table.Cell().Element(CellStyle).AlignRight().Text(((item.TongCong ?? 0)).ToString("N0"));
                            }
                        }
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
                .Padding(2)
                .AlignMiddle()
                .DefaultTextStyle(x => x.SemiBold().FontSize(9));

        static IContainer CellStyle(IContainer container) =>
            container
                .Border(1)
                .Padding(3)
                .AlignMiddle()
                .DefaultTextStyle(x => x.FontSize(8));
    }
}
