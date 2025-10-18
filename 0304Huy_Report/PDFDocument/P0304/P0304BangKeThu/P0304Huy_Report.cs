using H0304.NumberToText.Helpers;
using M0304NhanVien.Models;
using M0304.Models.BangKeThu;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.IO;
using M0304.Models.ThongTinDoanhNghiep;

namespace P0304.PDFDocument
{
    public class P0304ReportTemplatePDF : IDocument
    {
        private readonly List<M0304BangKeThu> _data;
        private readonly M0304ThongTinDoanhNghiep _dataDN;
        private string _ngayBatDau;
        private string _ngayKetThuc;
        private string _tenNVDN;
        private string _tenHTTT;
        private readonly string _logoPath;

        private List<M0304TongTheoQuyenSo> _tongTheoQuyenSo;
        private List<M0304NhanVienModel> _danhSachNhanVien;

        public P0304ReportTemplatePDF(
            List<M0304BangKeThu> data,
            string ngayBatDau,
            string ngayKetThuc,
            string tenNVDN,
            string tenHTTT,
            List<M0304NhanVienModel> danhSachNhanVien,
            List<M0304TongTheoQuyenSo> tongTheoQuyenSo,
            M0304ThongTinDoanhNghiep dataDN,
            string logoPath = null
        )
        {
            _data = data ?? new List<M0304BangKeThu>();
            _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
            _ngayBatDau = ngayBatDau;
            _ngayKetThuc = ngayKetThuc;
            _tenNVDN = tenNVDN;
            _tenHTTT = tenHTTT;
            _danhSachNhanVien = danhSachNhanVien ?? new List<M0304NhanVienModel>();
            _tongTheoQuyenSo = tongTheoQuyenSo ?? new List<M0304TongTheoQuyenSo>();
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
                            .Text("BẢNG KÊ THU TIỀN NGOẠI TRÚ THEO BL/HĐ")
                            .Bold()
                            .FontSize(12);

                        center.Item()
                            .Width(250)
                            .AlignCenter()
                            .Text($"Từ ngày {_ngayBatDau} đến ngày {_ngayKetThuc}")
                            .FontSize(9).Italic();

                        center.Item()
                            .Width(250)
                            .AlignCenter()
                            .Text($"{_tenHTTT}")
                            .Bold()
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
                            columns.ConstantColumn(70);
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(45);
                            columns.ConstantColumn(60);
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(50);
                        });

                        table.Header(header =>
                        {
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("STT");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Mã y tế");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Họ và tên");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Quyển sổ");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Số biên lai");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Loại");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Ngày thu");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Hủy");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Hoàn");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Số tiền");
                        });

                        int stt = 1;

                        if (_danhSachNhanVien != null && _danhSachNhanVien.Any())
                        {
                            foreach (var nv in _danhSachNhanVien)
                            {
                                table.Cell().ColumnSpan(10)
                                    .Element(CellStyle)
                                    .AlignLeft()
                                    .Text(nv.TenNhanVien?.ToUpper() ?? "")
                                    .FontSize(8)
                                    .Bold();

                                var quyenSoList = _tongTheoQuyenSo
                                    .Where(q => _data.Any(d => d.IDNhanVien == nv.ID && d.QuyenSo == q.QuyenSo))
                                    .ToList();

                                foreach (var qs in quyenSoList)
                                {
                                    // Lấy tất cả bản ghi của NV hiện tại trong quyển sổ này
                                    var chiTietNvQs = _data
                                        .Where(d => d.IDNhanVien == nv.ID && d.QuyenSo == qs.QuyenSo)
                                        .ToList();

                                    // Seri thực tế theo NV (nếu có, fallback về qs.Seri)
                                    var seriForNv = chiTietNvQs.Select(d => d.Seri).FirstOrDefault() ?? qs.Seri ?? "";

                                    // Tổng theo NV cho quyển sổ
                                    var tongTheoNVvaQS = chiTietNvQs
                                        .GroupBy(x => x.QuyenSo)
                                        .Select(g => new
                                        {
                                            TongHuy = g.Sum(x => x.Huy ?? 0m),
                                            TongHoan = g.Sum(x => x.Hoan ?? 0m),
                                            TongSoTien = g.Sum(x => x.SoTien ?? 0m)
                                        })
                                        .FirstOrDefault() ?? new { TongHuy = 0m, TongHoan = 0m, TongSoTien = 0m };

                                    // Ngày thu lớn nhất trong quyển sổ của NV
                                    DateTime? ngayThu = chiTietNvQs.Max(d => d.NgayThu);

                                    // Header nhóm: Mã NV - Seri.QuyểnSố
                                    table.Cell().ColumnSpan(6)
                                        .Element(CellStyleLeft)
                                        .AlignLeft()
                                        .Text($"      {nv.MaNhanVien} - {seriForNv}.{qs.QuyenSo}")
                                        .FontSize(9)
                                        .Bold();

                                    table.Cell()
                                        .Element(CellStyleNoBorder)
                                        .AlignCenter()
                                        .Text(ngayThu.HasValue ? ngayThu.Value.ToString("dd-MM-yyyy") : "")
                                        .FontSize(8)
                                        .Bold();

                                    table.Cell()
                                        .Element(CellStyleNoBorder)
                                        .AlignRight()
                                        .Text(tongTheoNVvaQS.TongHuy.ToString("N0"))
                                        .FontSize(8)
                                        .Bold();

                                    table.Cell()
                                        .Element(CellStyleNoBorder)
                                        .AlignRight()
                                        .Text(tongTheoNVvaQS.TongHoan.ToString("N0"))
                                        .FontSize(8)
                                        .Bold();

                                    table.Cell()
                                        .Element(CellStyleRight)
                                        .AlignRight()
                                        .Text(tongTheoNVvaQS.TongSoTien.ToString("N0"))
                                        .FontSize(8)
                                        .Bold();

                                    // Chi tiết từng dòng
                                    foreach (var item in chiTietNvQs)
                                    {
                                        table.Cell()
                                            .Element(CellStyle)
                                            .AlignCenter()
                                            .Text((stt++).ToString());

                                        table.Cell()
                                            .Element(CellStyle)
                                            .AlignCenter()
                                            .Text(item.MaYTe ?? "");

                                        table.Cell()
                                            .Element(CellStyle)
                                            .Text(item.HoVaTen ?? "");

                                        table.Cell()
                                            .Element(CellStyle)
                                            .AlignCenter()
                                            .Text(item.QuyenSo ?? "");

                                        table.Cell()
                                            .Element(CellStyle)
                                            .AlignCenter()
                                            .Text(item.SoBienLai ?? "");

                                        table.Cell()
                                            .Element(CellStyle)
                                            .AlignCenter()
                                            .Text(item.Loai ?? "");

                                        table.Cell()
                                            .Element(CellStyle)
                                            .AlignCenter()
                                            .Text(item.NgayThu.HasValue ? item.NgayThu.Value.ToString("dd-MM-yyyy") : "");

                                        table.Cell()
                                            .Element(CellStyle)
                                            .AlignRight()
                                            .Text(item.Huy.HasValue ? item.Huy.Value.ToString("N0") : "");

                                        table.Cell()
                                            .Element(CellStyle)
                                            .AlignRight()
                                            .Text(item.Hoan.HasValue ? item.Hoan.Value.ToString("N0") : "");

                                        table.Cell()
                                            .Element(CellStyle)
                                            .AlignRight()
                                            .Text(item.SoTien.HasValue ? item.SoTien.Value.ToString("N0") : "");
                                    }
                                }
                            }
                        }


                        var tongHuyAll = _data.Sum(x => x.Huy ?? 0m);
                        var tongHoanAll = _data.Sum(x => x.Hoan ?? 0m);
                        var tongSoTienAll = _data.Sum(x => x.SoTien ?? 0m);

                        var phaiNop = tongSoTienAll- tongHuyAll - tongHoanAll;

                        col.Item().EnsureSpace()
                        .Column(cuoi =>
                        {
                            cuoi.Spacing(5);

                            // Phần tổng cộng trong bảng
                            cuoi.Item().Table(table =>
                            {
                                table.ColumnsDefinition(columns =>
                                {
                                    columns.ConstantColumn(30);
                                    columns.ConstantColumn(50);
                                    columns.RelativeColumn();
                                    columns.ConstantColumn(70);
                                    columns.ConstantColumn(50);
                                    columns.ConstantColumn(45);
                                    columns.ConstantColumn(60);
                                    columns.ConstantColumn(50);
                                    columns.ConstantColumn(50);
                                    columns.ConstantColumn(50);
                                });

                                table.Cell().ColumnSpan(7)
                                    .Element(CellStyle)
                                    .AlignCenter()
                                    .Text("Tổng cộng")
                                    .Bold();

                                table.Cell().Element(CellStyle).AlignRight().Text(tongHuyAll.ToString("N0")).Bold();
                                table.Cell().Element(CellStyle).AlignRight().Text(tongHoanAll.ToString("N0")).Bold();
                                table.Cell().Element(CellStyle).AlignRight().Text(tongSoTienAll.ToString("N0")).Bold();
                            });

                            cuoi.Item().Height(1);

                            cuoi.Item().Text(text =>
                            {
                                text.Span("Số tiền phải nộp: ").NormalWeight();
                                text.Span($"{phaiNop:N0}").Bold();
                            });

                            cuoi.Item().Text(text =>
                            {
                                text.Span("Bằng chữ: ").NormalWeight();
                                text.Span($"{H0304NumberToTextHelper.ConvertSoThanhChu(phaiNop)}").Bold().Italic();
                            });

                            cuoi.Item().Row(row =>
                            {
                                row.RelativeItem().Text("");
                                row.ConstantItem(200).Column(right =>
                                {
                                    right.Item().AlignCenter().Text($"Ngày {DateTime.Now:dd} Tháng {DateTime.Now:MM} Năm {DateTime.Now:yyyy}");
                                    right.Item().AlignCenter().Text("Người lập bảng").Bold();
                                    right.Item().Height(40);
                                    right.Item().AlignCenter().Text($"{_tenNVDN}").Bold();
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
                .Padding(3)
                .AlignMiddle()
                .DefaultTextStyle(x => x.SemiBold().FontSize(10));

        static IContainer CellStyle(IContainer container) =>
            container
                .Border(1)
                .Padding(3)
                .AlignMiddle()
                .DefaultTextStyle(x => x.FontSize(8));
        static IContainer CellStyleNoBorder(IContainer container) =>
            container
                .Padding(3)
                .AlignMiddle()
                .DefaultTextStyle(x => x.FontSize(8));
        static IContainer CellStyleLeft(IContainer container) =>
        container
        .BorderLeft(1)
        .Padding(3)
        .AlignMiddle()
        .DefaultTextStyle(x => x.FontSize(8));
        static IContainer CellStyleRight(IContainer container) =>
        container
        .BorderRight(1)
        .Padding(3)
        .AlignMiddle()
        .DefaultTextStyle(x => x.FontSize(8));
    }
}
