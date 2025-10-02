using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using M0304.Models.ThongTinDoanhNghiep;
using M0304L.Models.PhieuTheoDoiTruyenDich;
using QuestPDF.Drawing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using SixLabors.Fonts;
using SkiaSharp;
using System.Globalization;
using System.IO;
using System.Linq;

namespace P0304L.PDFDocument
{
    public class P0304LReportTemplatePDF : IDocument
    {
        private readonly HoSoBenhAnModel _data;
        private readonly M0304ThongTinDoanhNghiep _dataDN;
        private readonly string _logoPath;

        public P0304LReportTemplatePDF(
            HoSoBenhAnModel data,
            M0304ThongTinDoanhNghiep dataDN,
            string logoPath = null
        )
        {
            _data = data ?? new HoSoBenhAnModel();
            _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
            _logoPath = logoPath;
        }

        public DocumentMetadata GetMetadata() => DocumentMetadata.Default;
        public DocumentSettings GetSettings() => DocumentSettings.Default;

        public void Compose(IDocumentContainer container)
        {
            var truyenDich = (_data.TruyenDich ?? Enumerable.Empty<TruyenDich>())
                .OrderBy(x => x.NgayThang ?? DateTime.MinValue)
                .ToList();

            DateTime? ngaySinh = _data?.ThongTinBN?.NgaySinh;
            DateTime? firstNgayThang = (_data?.TruyenDich ?? Enumerable.Empty<TruyenDich>())
                .Where(x => x.NgayThang.HasValue)
                .OrderBy(x => x.NgayThang.Value)
                .Select(x => x.NgayThang)
                .FirstOrDefault();

            int? tuoi = null;
            if (ngaySinh.HasValue && firstNgayThang.HasValue)
            {
                var refDate = firstNgayThang.Value.Date;
                int t = refDate.Year - ngaySinh.Value.Year;
                if (refDate < ngaySinh.Value.AddYears(t)) t--;
                tuoi = t;
            }

            string ngaySinhDisplay = ngaySinh?.ToString("dd-MM-yyyy") ?? "";
            string tuoiDisplay = tuoi.HasValue ? $" ({tuoi.Value} tuổi)" : "";
            string ngaySinhFull = $"{ngaySinhDisplay}{tuoiDisplay}";

            container.Page(page =>
            {
                page.Size(PageSizes.A4);
                page.Margin(20);
                page.DefaultTextStyle(x => x.FontSize(10));

                page.Header().ShowOnce().Column(col =>
                {
                    col.Item().Row(row =>
                    {
                        row.RelativeItem(3).Row(left =>
                        {
                            left.RelativeItem().Column(info =>
                            {
                                info.Item().Row(row =>
                                {
                                    row.RelativeItem().AlignLeft().Text($"{_dataDN.TenCoQuanChuyenMon ?? ""}").FontSize(9);

                                    row.RelativeItem().AlignCenter().Text("PHIẾU THEO DÕI").FontSize(12).Bold();

                                    row.RelativeItem().AlignLeft().PaddingLeft(50).Text("MS:07/BV-02");
                                });

                                info.Item().Row(row =>
                                {
                                    row.RelativeItem().AlignLeft().Text($"{_dataDN.TenCSKCB ?? ""}").FontSize(9);

                                    row.RelativeItem().AlignCenter().Text("TRUYỀN DỊCH").FontSize(12).Bold();

                                    row.RelativeItem().AlignLeft().PaddingLeft(50)
                                        .Text(text =>
                                        {
                                            text.Span("Số vào viện: ").FontSize(10);
                                            text.Span($"{_data?.ThongTinBN?.MaVaoVien ?? ""}").FontSize(10);
                                        });
                                });
                                info.Item().Text(text =>
                                {
                                    text.Span("Khoa: ").FontSize(10);
                                    text.Span($"{_data?.ThongTinBN?.TenKhoa ?? ""}").FontSize(10);
                                });
                            });
                        });
                    });

                    col.Item().Padding(5).Row(row =>
                    {
                        row.RelativeItem().AlignLeft()
                        .Text(text =>
                        {
                            text.Span("*     Họ tên người bệnh: ").FontSize(10);
                            text.Span($"{_data?.ThongTinBN?.TenBenhNhan ?? ""}").FontSize(10);

                            text.Span("        Ngày sinh: ").FontSize(10);
                            text.Span(ngaySinhFull).FontSize(10);

                            text.Span("       Giới tính: ").FontSize(10);
                            text.Span($"{_data?.ThongTinBN?.GioiTinh ?? ""}").FontSize(10);
                        });
                    });

                    col.Item().PaddingLeft(5).Row(row =>
                    {

                        row.RelativeItem().AlignLeft()
                            .Text(text =>
                            {
                                text.Span("*     Số giường: ").FontSize(10);
                                text.Span($"{_data?.ThongTinBN?.TenGiuong}").FontSize(10);
                            });

                        row.RelativeItem().AlignLeft()
                            .Text(text =>
                            {
                                text.Span("Buồng: ").FontSize(10);
                                text.Span($"{_data?.ThongTinBN?.TenPhong ?? ""}").FontSize(10);
                            });
                    });

                    col.Item().AlignLeft().Padding(5)
                        .Text(text =>
                        {
                            text.Span("*     Chẩn đoán: ").FontSize(10);
                            text.Span($"{_data?.ThongTinBN?.ChanDoan ?? ""}").FontSize(10);
                        });
                });

                page.Content().PaddingVertical(6).Column(col =>
                {
                    col.Item().Table(table =>
                    {
                        table.ColumnsDefinition(columns =>
                        {
                            columns.ConstantColumn(54);
                            columns.RelativeColumn();
                            columns.ConstantColumn(40);
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(44);
                            columns.ConstantColumn(48);
                            columns.ConstantColumn(48);
                            columns.ConstantColumn(68);
                            columns.ConstantColumn(68);
                        });

                        table.Cell().RowSpan(2).Element(CellStyleHeader).AlignCenter().Text("Ngày tháng").Bold();
                        table.Cell().RowSpan(2).Element(CellStyleHeader).Element(e =>
                        {
                            e.Column(column =>
                            {
                                column.Item().AlignCenter().Text("Tên dịch truyền/").Bold();
                                column.Item().AlignCenter().Text("Hàm lượng").Bold();
                            });
                        });
                        table.Cell().RowSpan(2).Element(CellStyleHeader).AlignCenter().Text("Số lượng").Bold();
                        table.Cell().RowSpan(2).Element(CellStyleHeader).AlignCenter().Text("Lô/ Số sản xuất").Bold();
                        table.Cell().RowSpan(2).Element(CellStyleHeader).AlignCenter().Text("Tốc độ giọt/ph").Bold();
                        table.Cell().ColumnSpan(2).Element(CellStyleHeader).AlignCenter().Text("Thời gian").Bold();
                        table.Cell().RowSpan(2).Element(CellStyleHeader).AlignCenter().Text("BS chỉ định").Bold();
                        table.Cell().RowSpan(2).Element(CellStyleHeader).AlignCenter().Text("YT (ĐD) thực hiện").Bold();

                        // ===== Header dòng 2 =====
                        table.Cell().Element(CellStyleHeader).AlignCenter().Text("bắt đầu").Bold();
                        table.Cell().Element(CellStyleHeader).AlignCenter().Text("kết thúc").Bold();

                        // ===== Dữ liệu =====
                        int stt = 1;
                        foreach (var item in truyenDich)
                        {
                            table.Cell().Element(CellStyle).AlignCenter().Text(item.NgayThang?.ToString("dd-MM-yyyy") ?? "");
                            table.Cell().Element(CellStyle).AlignLeft().Text(item.TenDichTruyen ?? "");
                            table.Cell().Element(CellStyle).AlignRight().Text(item.SoLuong?.ToString() ?? "");
                            table.Cell().Element(CellStyle).AlignCenter().Text(item.SoLo ?? "");
                            table.Cell().Element(CellStyle).AlignRight().Text(item.TocDo?.ToString() ?? "");
                            table.Cell().Element(CellStyle).AlignCenter().Text(item.BatDau.ToString() ?? "");
                            table.Cell().Element(CellStyle).AlignCenter().Text(item.KetThuc.ToString() ?? "");
                            table.Cell().Element(CellStyle).AlignLeft().Text(item.BSChiDinh ?? "");
                            table.Cell().Element(CellStyle).AlignLeft().Text(item.NguoiThucHien ?? "");
                        }
                    });
                    
                });
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
                        .DefaultTextStyle(x => x.FontSize(9));
            });
        }
    }
}

