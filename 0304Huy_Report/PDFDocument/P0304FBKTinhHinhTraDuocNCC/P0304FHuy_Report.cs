using H0304.NumberToText.Helpers;
using M0304NhanVien.Models;
using M0304F.Models.BKTinhHinhTraDuocNCC;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.IO;
using M0304.Models.ThongTinDoanhNghiep;

namespace P0304F.PDFDocument
{
    public class P0304FReportTemplatePDF : IDocument
    {
        private readonly List<M0304FBKTinhHinhTraDuocNCC> _data;
        private readonly M0304ThongTinDoanhNghiep _dataDN;
        private readonly List<CongTyDto> _dsCongTy;
        private readonly string _tenKho;
        private string _ngayBatDau;
        private string _ngayKetThuc;
        private readonly string _logoPath;

        public P0304FReportTemplatePDF(
            List<M0304FBKTinhHinhTraDuocNCC> data,
            M0304ThongTinDoanhNghiep dataDN,
            List<CongTyDto> dsCongTy,
            string tenKho,
            string ngayBatDau,
            string ngayKetThuc,
            string logoPath = null
        )
        {
            _data = data ?? new List<M0304FBKTinhHinhTraDuocNCC>();
            _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
            _dsCongTy = dsCongTy ?? new List<CongTyDto>();
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
                page.Size(PageSizes.A4.Landscape());
                page.Margin(15);
                page.DefaultTextStyle(x => x.FontSize(8).FontFamily("Times New Roman", "Arial"));

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
                            .Text("BẢNG KÊ TÌNH HÌNH TRẢ DƯỢC NCC")
                            .Bold()
                            .FontSize(12);

                        center.Item()
                            .Width(250)
                            .AlignCenter()
                            .Text($"Từ ngày {_ngayBatDau} đến ngày {_ngayKetThuc}")
                            .Italic()
                            .FontSize(8);
                        center.Item()
                            .AlignCenter()
                            .Text("Nguồn dược: Mua")
                            .Italic()
                            .FontSize(8);

                        center.Item()
                            .Width(250)
                            .AlignCenter()
                            .Text($"Kho trả: {_tenKho}")
                            .Italic()
                            .FontSize(8);
                        center.Item()
                            .Width(250)
                            .AlignCenter()
                            .Text("")
                            .Italic()
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
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(45);
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(45);
                            columns.RelativeColumn(1);
                            columns.ConstantColumn(45);
                            columns.RelativeColumn(1);
                            columns.ConstantColumn(30);
                            columns.ConstantColumn(40);
                            columns.ConstantColumn(40);
                            columns.ConstantColumn(40);
                            columns.ConstantColumn(40);
                            columns.ConstantColumn(40);
                            columns.ConstantColumn(50);
                        });


                        table.Header(header =>
                        {
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("STT");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Ngày hoa đơn");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Số hóa đơn");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Ngày trả");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Phiếu trả");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Công ty");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Mã ID");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Tên thuốc, hàm lượng");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Quy cách");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Số lô");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("SL đóng gói");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("SLlẻ");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Đơn giá đóng gói");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Dơn giá lẻ");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Thành tiền");

                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("1");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("2");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("3");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("4");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("5");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("6");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("7");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("8");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("9");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("10");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("11");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("12");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("13");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("14");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("15 = 12*14");
                        });

                        foreach (var cty in _dsCongTy)
                        {
                            int stt = 1;
                            var data = _data.Where(d => d.IDCongTy == cty.ID).ToList();
                            table.Cell().ColumnSpan(15).Border(1).BorderColor(Colors.Black).MinHeight(13).AlignCenter().AlignMiddle().Text(cty.Ten).SemiBold();
                            foreach (var item in data)
                            {
                                table.Cell().Element(CellStyle).AlignCenter().Text(stt++.ToString());
                                table.Cell().Element(CellStyle).AlignCenter().Text(item.NgayHoaDon?.ToString("dd-MM-yyyy") ?? "");
                                table.Cell().Element(CellStyle).AlignLeft().Text(item.SoHoaDon ?? "");
                                table.Cell().Element(CellStyle).AlignCenter().Text(item.NgayTra?.ToString("dd-MM-yyyy") ?? "");
                                table.Cell().Element(CellStyle).AlignLeft().Text(item.PhieuTra ?? "");
                                table.Cell().Element(CellStyle).AlignLeft().Text(item.CongTy ?? "");
                                table.Cell().Element(CellStyle).AlignCenter().Text(item.MaID ?? "");
                                table.Cell().Element(CellStyle).AlignLeft().Text(item.TenThuoc ?? "");
                                table.Cell().Element(CellStyle).AlignLeft().Text(item.QuyCach ?? "");
                                table.Cell().Element(CellStyle).AlignCenter().Text(item.SoLo ?? "");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.SLDongGoi?.ToString("N0") ?? "");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.SLLe?.ToString("N0") ?? "");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.DonGiaDongGoi?.ToString("N0") ?? "");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.DonGiaLe?.ToString("N0") ?? "");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.ThanhTien?.ToString("N0") ?? "");
                            }
                        }

                        var tongCong = _data.Sum(x => x.ThanhTien ?? 0);
                        var vat = 0;
                        var tongTien = tongCong + vat;

                        col.Item().EnsureSpace()
                        .Column(cuoi =>
                        {
                            cuoi.Item().Height(5);

                            cuoi.Item().Row(row =>
                            {
                                row.RelativeItem(10);
                                row.RelativeItem(5).Column(col =>
                                {
                                    col.Item().Row(r =>
                                    {
                                        r.RelativeItem().AlignLeft().Text("Tổng cộng:").Bold();
                                        r.RelativeItem().AlignRight().Text($"{tongCong:N0}").Bold();
                                    });

                                    col.Item().Row(r =>
                                    {
                                        r.RelativeItem().AlignLeft().Text("Tiền VAT:").Bold();
                                        r.RelativeItem().AlignRight().Text($"{vat:N0}").Bold();
                                    });

                                    col.Item().Row(r =>
                                    {
                                        r.RelativeItem().AlignLeft().Text("Tổng tiền:").Bold();
                                        r.RelativeItem().AlignRight().Text($"{tongTien:N0}").Bold();
                                    });
                                });
                            });

                            cuoi.Item().PaddingTop(10).Row(row =>
                            {
                                row.RelativeItem().AlignRight().
                                Text($"Ngày {DateTime.Now:dd} Tháng {DateTime.Now:MM} Năm {DateTime.Now:yyyy}").Bold().Italic();
                            });

                            cuoi.Item().PaddingTop(10).Row(row =>
                            {
                                row.RelativeItem().AlignCenter().Text("Thủ kho").Bold();
                                row.RelativeItem().AlignCenter().Text("Kế toán").Bold();
                                row.RelativeItem().AlignCenter().Text("Người lập").Bold();
                                row.RelativeItem().AlignCenter().Text("Trưởng khoa").Bold();
                            });
                        });
                    });

                });

                page.Footer()
                    .Row(row =>
                    {
                        row.RelativeItem()
                            .AlignLeft()
                            .Text(txt =>
                            {
                                txt.Span("(In ngày: ").Italic();
                                txt.Span(DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")).Italic();
                                txt.Span(")").Italic();
                            });

                        row.RelativeItem()
                            .AlignLeft()
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
