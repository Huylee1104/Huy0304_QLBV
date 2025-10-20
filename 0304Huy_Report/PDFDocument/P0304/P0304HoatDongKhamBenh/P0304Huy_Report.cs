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
                page.Size(PageSizes.A4.Landscape());
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
                                left.ConstantItem(40).AlignMiddle().Text("");
                            }

                            left.RelativeItem().Column(info =>
                            {
                                info.Item().Text(_dataDN.TenCoQuanChuyenMon ?? "").FontSize(8);
                                info.Item().Text(_dataDN.TenCSKCB ?? "").FontSize(8).Bold();
                            });
                        });
                    });

                    col.Item().AlignCenter().Column(center =>
                    {
                        center.Item()
                            .AlignCenter()
                            .Text("HOẠT ĐỘNG KHÁM BỆNH")
                            .Bold()
                            .FontSize(12);

                        center.Item()
                            .AlignCenter()
                            .Text($"Từ ngày {_ngayBatDau} đến ngày {_ngayKetThuc}")
                            .FontSize(9)
                            .Italic();
                    });
                });

                var allTongSoTien = _data.Sum(x => x.TongSo);
                var allY_HOC_CO_TRUYEN = _data.Sum(x => x.YHocCoTruyen);
                var TRE_EM_DUOI_6 = _data.Sum(x => x.TreEmDuoi6Tuoi);
                var BHYT = _data.Sum(x => x.BHYT);
                var VIEN_PHI = _data.Sum(x => x.VienPhi);
                var KHONG_THU_DUOC = _data.Sum(x => x.KhongThuDuoc);
                var CAP_CUU = _data.Sum(x => x.CapCuu);
                var SO_NGUOI_VAO_VIEN = _data.Sum(x => x.SoNguoiVaoVien);
                var SO_NGUOI_CHUYEN_VIEN = _data.Sum(x => x.SoNguoiChuyenVien);
                var NT_SO_NGUOI_BENH = _data.Sum(x => x.NTSoNguoiBenh);
                var NT_YHCT = _data.Sum(x => x.NTYHocCoTruyen);
                var NT_TRE_EM_DUOI_6 = _data.Sum(x => x.NTTreEmDuoi6Tuoi);
                var NT_SO_NGAY = _data.Sum(x => x.NTSoNgay);

                page.Content().PaddingVertical(6).Column(col =>
                {
                    col.Item().Table(table =>
                    {
                        table.ColumnsDefinition(columns =>
                        {
                            columns.ConstantColumn(35);
                            columns.RelativeColumn();
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(80);
                            columns.ConstantColumn(80);
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(50);
                            columns.ConstantColumn(50);
                        });

                        table.Header(header =>
                        {
                            // First row
                            header.Cell().RowSpan(3).Element(CellStyleHeader).AlignCenter().Text("STT");
                            header.Cell().RowSpan(3).Element(CellStyleHeader).AlignCenter().Text("Dịch Vụ");
                            header.Cell().ColumnSpan(7).Element(CellStyleHeader).AlignCenter().Text("Số lần khám");
                            header.Cell().RowSpan(3).Element(CellStyleHeader).AlignCenter().Text("Số người bệnh vào viện");
                            header.Cell().RowSpan(3).Element(CellStyleHeader).AlignCenter().Text("Số người bệnh chuyển viện");
                            header.Cell().ColumnSpan(4).Element(CellStyleHeader).AlignCenter().Text("Điều trị ngoại trú");

                            // Second row
                            header.Cell().RowSpan(2).Element(CellStyleHeader).AlignCenter().Text("Tổng số");
                            header.Cell().ColumnSpan(6).Element(CellStyleHeader).AlignCenter().Text("Trong đó");
                            header.Cell().RowSpan(2).Element(CellStyleHeader).AlignCenter().Text("Số người bệnh");
                            header.Cell().ColumnSpan(2).Element(CellStyleHeader).AlignCenter().Text("Trong đó");
                            header.Cell().RowSpan(2).Element(CellStyleHeader).AlignCenter().Text("Số ngày");

                            // Third row
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("YHCT");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("TE<6 tuổi");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Bảo hiểm y tế");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Viện phí");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Không thu được");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("Cấp cứu");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("YHCT");
                            header.Cell().Element(CellStyleHeader).AlignCenter().Text("TE<6 tuổi");

                            table.Header(header =>
                            {
                                header.Cell().ColumnSpan(2).Element(CellStyleHeader).AlignCenter().Text("Tổng số:");
                                header.Cell().Element(CellStyleHeader).AlignRight().Text(allTongSoTien?.ToString("N0" ?? "0"));
                                header.Cell().Element(CellStyleHeader).AlignRight().Text(allY_HOC_CO_TRUYEN?.ToString("N0" ?? "0"));
                                header.Cell().Element(CellStyleHeader).AlignRight().Text(TRE_EM_DUOI_6?.ToString("N0" ?? "0"));
                                header.Cell().Element(CellStyleHeader).AlignRight().Text(BHYT?.ToString("N0" ?? "0"));
                                header.Cell().Element(CellStyleHeader).AlignRight().Text(VIEN_PHI?.ToString("N0" ?? "0"));
                                header.Cell().Element(CellStyleHeader).AlignRight().Text(KHONG_THU_DUOC?.ToString("N0" ?? "0"));
                                header.Cell().Element(CellStyleHeader).AlignRight().Text(CAP_CUU?.ToString("N0" ?? "0"));
                                header.Cell().Element(CellStyleHeader).AlignRight().Text(SO_NGUOI_VAO_VIEN?.ToString("N0" ?? "0"));
                                header.Cell().Element(CellStyleHeader).AlignRight().Text(SO_NGUOI_CHUYEN_VIEN?.ToString("N0" ?? "0"));
                                header.Cell().Element(CellStyleHeader).AlignRight().Text(NT_SO_NGUOI_BENH?.ToString("N0" ?? "0"));
                                header.Cell().Element(CellStyleHeader).AlignRight().Text(NT_YHCT?.ToString("N0" ?? "0"));
                                header.Cell().Element(CellStyleHeader).AlignRight().Text(NT_TRE_EM_DUOI_6?.ToString("N0" ?? "0"));
                                header.Cell().Element(CellStyleHeader).AlignRight().Text(NT_SO_NGAY?.ToString("N0" ?? "0"));
                            });
                            int stt = 1;
                            foreach (var item in _data)
                            {
                                table.Cell().Element(CellStyle).AlignCenter().Text(stt.ToString());
                                table.Cell().Element(CellStyle).AlignLeft().Text(item.DichVu ?? string.Empty);
                                table.Cell().Element(CellStyle).AlignRight().Text(item.TongSo?.ToString("N0") ?? "0");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.YHocCoTruyen?.ToString("N0") ?? "0");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.TreEmDuoi6Tuoi?.ToString("N0") ?? "0");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.BHYT?.ToString("N0") ?? "0");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.VienPhi?.ToString("N0") ?? "0");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.KhongThuDuoc?.ToString("N0") ?? "0");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.CapCuu?.ToString("N0") ?? "0");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.SoNguoiVaoVien?.ToString("N0") ?? "0");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.SoNguoiChuyenVien?.ToString("N0") ?? "0");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.NTSoNguoiBenh?.ToString("N0") ?? "0");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.NTYHocCoTruyen?.ToString("N0") ?? "0");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.NTTreEmDuoi6Tuoi?.ToString("N0") ?? "0");
                                table.Cell().Element(CellStyle).AlignRight().Text(item.NTSoNgay?.ToString("N0") ?? "0");
                                stt++;
                            }
                        });

                        col.Item().PaddingTop(10).Row(row =>
                        {
                            row.RelativeItem().AlignRight().PaddingRight(50).Text(text =>
                            {
                                text.Span($"Ngày {DateTime.Now:dd} tháng {DateTime.Now:MM} năm {DateTime.Now:yyyy}")
                                    .Italic().Bold();
                            });
                        });

                        col.Item().PaddingTop(10).Row(row =>
                        {
                            row.RelativeItem().AlignCenter().Text("NGƯỜI LẬP BIỂU").Bold();
                            row.RelativeItem().AlignCenter().Text("TRƯỞNG PHÒNG KẾ HOẠCH TỔNG HỢP").Bold();
                            row.RelativeItem().AlignCenter().Text("GIÁM ĐỐC").Bold();
                        });
                    });

                    page.Footer()
                    .Row(row =>
                    {
                        // Bên trái: ngày giờ hiện tại
                        row.RelativeColumn()
                           .AlignLeft()
                           .Text("Biểu 02 - KB");

                        // Bên phải: số trang
                        row.RelativeColumn()
                           .AlignRight()
                           .Text(txt =>
                           {
                               txt.CurrentPageNumber();
                           });
                    });
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
                    .DefaultTextStyle(x => x.FontSize(8));
        }
    }
}
