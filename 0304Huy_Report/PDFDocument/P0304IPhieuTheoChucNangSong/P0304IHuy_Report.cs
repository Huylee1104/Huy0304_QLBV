using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using M0304.Models.ThongTinDoanhNghiep;
using M0304I.Models.PhieuTheoDoiChucNangSong;
using QuestPDF.Drawing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.Globalization;
using System.IO;
using System.Linq;

namespace P0304I.PDFDocument
{
    public class P0304IReportTemplatePDF : IDocument
    {
        private readonly HoSoBenhAnModel _data;
        private readonly M0304ThongTinDoanhNghiep _dataDN;
        private readonly string _logoPath;

        public P0304IReportTemplatePDF(
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
            var sinhHoa = (_data.SinhHieus ?? Enumerable.Empty<SinhHieuModel>())
                .OrderBy(x => x.NgayKhaoSat ?? DateTime.MinValue) 
                .ToList();

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
                                info.Item().Text(_dataDN.TenCSKCB ?? "").FontSize(9);
                                info.Item().Text(_dataDN.TenCoQuanChuyenMon ?? "").FontSize(9);
                                info.Item().Row(row =>
                                {
                                    row.RelativeItem().AlignLeft()
                                        .Text($"{_dataDN.DiaChi ?? ""}").FontSize(9);

                                    row.RelativeItem().AlignLeft().PaddingLeft(145)
                                        .Text(text =>
                                        {
                                            text.Span("MS: ").SemiBold().FontSize(10);
                                            text.Span($"{_data?.ThongTinBenhNhan?.MaBenhNhan ?? ""}").FontSize(10);
                                        });
                                });

                                info.Item().Row(row =>
                                {
                                    row.RelativeItem().AlignLeft()
                                        .Text($"{_dataDN.DienThoai ?? ""}").FontSize(9);

                                    row.RelativeItem().AlignLeft().PaddingLeft(145)
                                        .Text(text =>
                                        {
                                            text.Span("Mã vào viện: ").SemiBold().FontSize(10);
                                            text.Span($"{_data?.ThongTinBenhNhan?.MaVaoVien ?? ""}").FontSize(10);
                                        });
                                });
                                info.Item().Text(_data?.ThongTinBenhNhan?.TenKhoa ?? "").FontSize(9);
                            });
                        });
                    });

                    col.Item().AlignCenter()
                        .Text("PHIẾU THEO DÕI CHỨC NĂNG SỐNG")
                        .Bold().FontSize(12);

                    col.Item().Padding(5).Row(row =>
                    {
                        row.RelativeItem().AlignLeft()
                            .Text(text =>
                            {
                                text.Span("Họ tên người bệnh: ").SemiBold().FontSize(10);
                                text.Span($"{_data?.ThongTinBenhNhan?.TenBenhNhan ?? ""}").FontSize(10);
                            });

                        row.RelativeItem().AlignCenter()
                            .Text(text =>
                            {
                                text.Span("Tuổi: ").SemiBold().FontSize(10);
                                text.Span($"{_data?.ThongTinBenhNhan?.NgaySinh}").FontSize(10);
                            });

                        row.RelativeItem().AlignRight()
                            .Text(text =>
                            {
                                text.Span("Giới tính: ").SemiBold().FontSize(10);
                                text.Span($"{_data?.ThongTinBenhNhan?.GioiTinh ?? ""}").FontSize(10);
                            });
                    });

                    col.Item().PaddingBottom(5).PaddingLeft(5).AlignLeft()
                        .Text(text =>
                        {
                            text.Span("Địa chỉ: ").SemiBold().FontSize(10);
                            text.Span($"{_data?.ThongTinBenhNhan?.DiaChi ?? ""}").FontSize(10);
                        });

                    col.Item().AlignLeft().PaddingLeft(5)
                        .Text(text =>
                        {
                            text.Span("Chẩn đoán: ").SemiBold().FontSize(10);
                            text.Span($"{_data?.ThongTinBenhNhan?.ChanDoan ?? ""}").FontSize(10);
                        });
                    col.Item().Padding(5).Row(row =>
                    {
                        row.RelativeItem().AlignLeft()
                            .Text(text =>
                            {
                                text.Span("Tên phòng: ").SemiBold().FontSize(10);
                                text.Span($"{_data?.ThongTinBenhNhan?.TenPhong ?? ""}").FontSize(10);
                            });

                        row.RelativeItem().AlignLeft()
                            .Text(text =>
                            {
                                text.Span("Tên giường: ").SemiBold().FontSize(10);
                                text.Span($"{_data?.ThongTinBenhNhan?.TenGiuong}").FontSize(10);
                            });

                    });
                });

                var listNgay = _data.SinhHieus?.Select(sh => sh.NgayKhaoSat?.ToString("dd-MM")).ToList();
                var listGio = _data.SinhHieus?.Select(sh => sh.NgayKhaoSat?.ToString("HH:mm")).ToList();
                var listNhietDo = _data.SinhHieus?.Select(sh => sh.NhietDo).ToList();
                var listMach = _data.SinhHieus?.Select(sh => sh.Mach).ToList();
                var listHuyetAp = _data.SinhHieus?.Select(sh => sh.HuyetAp).ToList();
                var listCanNang = _data.SinhHieus?.Select(sh => sh.CanNang).ToList();
                var listNhipTho = _data.SinhHieus?.Select(sh => sh.NhipTho).ToList();

                (int paddingTop, int paddingBottom) TinhPadding010_Auto(double value, double step, bool roundToInt = true)
                {
                    double cellValueTop = Math.Ceiling(value / step) * step;

                    double offsetRatio = (cellValueTop - value) / step;
                    offsetRatio = Math.Clamp(offsetRatio, 0, 1);

                    double pTop = offsetRatio * 11.0;
                    double pBot = 11.0 - pTop;

                    if (roundToInt)
                    {
                        pTop = Math.Round(pTop);
                        pBot = 11 - pTop;
                    }

                    return ((int)pTop, (int)pBot);
                }

                page.Content().PaddingVertical(6).Element(e =>
                {
                    e.Column(column =>
                    {
                        column.Item().Table(table =>
                        {
                            table.ColumnsDefinition(columns =>
                            {
                                columns.ConstantColumn(52);
                                columns.ConstantColumn(53);
                                for (int i = 0; i < _data.SinhHieus.Count(); i++)
                                {
                                    columns.RelativeColumn();
                                }
                            });
                            table.Cell().ColumnSpan(2).Element(CellStyle).AlignCenter().Text("Ngày tháng");
                            int j = 0;
                            while (j < listNgay.Count)
                            {
                                var currentNgay = listNgay[j];

                                if (j + 1 < listNgay.Count && listNgay[j + 1] == currentNgay)
                                {
                                    table.Cell().ColumnSpan(2).Element(CellStyle).AlignCenter().Text($"{currentNgay}");
                                    j += 2;
                                }
                                else
                                {
                                    table.Cell().Element(CellStyle).AlignCenter().Text($"{currentNgay}");
                                    j++;
                                }
                            }
                            table.Cell().ColumnSpan(2).Element(CellStyle).AlignCenter().Text("Giờ");
                            foreach (var gio in listGio)
                            {
                                table.Cell().Element(CellStyle).AlignCenter().Text($"{gio}");
                            }
                            table.Cell().Element(CellStyle).AlignCenter().Element(e =>
                            {
                                e.Column(column =>
                                {
                                    column.Item().Text("Mạch");
                                    column.Item().Text("(L/ph)");
                                });
                            });
                            table.Cell().Element(CellStyle).AlignCenter().Text("Nhiệt độ (C)");
                            for (int i = 0; i < _data.SinhHieus.Count(); i++)
                            {
                                table.Cell().Element(CellStyle).AlignCenter().Text("");
                            }

                            table.Cell().Element(CellStyleTop).Height(25).AlignCenter().Text("160");
                            table.Cell().Element(CellStyleTop).Height(25).AlignCenter().Text("41");
                            for (int i = 0; i < listMach.Count && i < listNhietDo.Count; i++) // tác hàm ra
                            {
                                double mach = double.TryParse(listMach[i], out var mVal) ? mVal : 0;
                                double nhietDo = double.TryParse(listNhietDo[i], out var nVal) ? nVal : 0;

                                var (ptMach, pbMach) = TinhPadding010_Auto(mach, step: 20);
                                var (ptNhiet, pbNhiet) = TinhPadding010_Auto(nhietDo, step: 1);

                                if ((mach > 140 && mach <= 160) || (nhietDo > 40 && nhietDo <= 41))
                                {
                                    if ((mach > 140 && mach <= 160) && !(nhietDo > 40 && nhietDo <= 41))
                                    {
                                        table.Cell().AlignCenter().Border(1).Row(row =>
                                        {
                                            row.RelativeItem(1)
                                               .Element(c => c.CellStyleCham(paddingTop: ptMach, paddingBottom: pbMach))
                                               .AlignCenter()
                                               .Text(text =>
                                               {
                                                   text.Span($"{mach} ").FontColor("#FF0000").FontSize(8);
                                                   text.Span("●").FontColor("#FF0000");
                                               });
                                        });
                                    }
                                    else if (!(mach > 140 && mach <= 160) && (nhietDo > 40 && nhietDo <= 41))
                                    {
                                        table.Cell().AlignCenter().Border(1).Row(row =>
                                        {
                                            row.RelativeItem(1)
                                               .Element(c => c.CellStyleCham(paddingTop: ptNhiet, paddingBottom: pbNhiet)).AlignCenter().Text(text =>
                                               {
                                                   text.Span($"{nhietDo}").FontColor("#0000FF").FontSize(8);
                                                   text.Span("●").FontColor("#0000FF");
                                               });
                                        });
                                    }
                                    else if ((mach > 140 && mach <= 160) && (nhietDo > 40 && nhietDo <= 41))
                                    {
                                        table.Cell().AlignCenter().Border(1).Element(container =>
                                        {
                                            container.Row(row =>
                                            {
                                                row.RelativeItem(1)
                                                   .Element(c => c.CellStyleCham(paddingTop: ptMach, paddingBottom: pbMach)).AlignCenter().Text(text =>
                                                   {
                                                       text.Span($"{mach} ").FontColor("#FF0000").FontSize(8);
                                                       text.Span("●").FontColor("#FF0000");
                                                   });

                                                row.RelativeItem(1)
                                                   .Element(c => c.CellStyleCham(paddingTop: ptNhiet, paddingBottom: pbNhiet)).AlignCenter().Text(text =>
                                                   {
                                                       text.Span($"{nhietDo}").FontColor("#0000FF").FontSize(8);
                                                       text.Span("●").FontColor("#0000FF");
                                                   });
                                            });
                                        });
                                    }
                                }
                                else
                                {
                                    table.Cell().Element(CellStyleBigSize).AlignCenter().Text("");
                                }
                            }

                            table.Cell().Element(CellStyleTop).Height(25).AlignCenter().Text("140");
                            table.Cell().Element(CellStyleTop).Height(25).AlignCenter().Text("40");
                            for (int i = 0; i < listMach.Count && i < listNhietDo.Count; i++) // tác hàm ra
                            {
                                double mach = double.TryParse(listMach[i], out var mVal) ? mVal : 0;
                                double nhietDo = double.TryParse(listNhietDo[i], out var nVal) ? nVal : 0;

                                var (ptMach, pbMach) = TinhPadding010_Auto(mach, step: 20);
                                var (ptNhiet, pbNhiet) = TinhPadding010_Auto(nhietDo, step: 1);

                                if ((mach > 120 && mach <= 140) || (nhietDo > 39 && nhietDo <= 40))
                                {
                                    if ((mach > 120 && mach <= 140) && !(nhietDo > 39 && nhietDo <= 40))
                                    {
                                        table.Cell().AlignCenter().Border(1).Row(row =>
                                        {
                                            row.RelativeItem(1)
                                               .Element(c => c.CellStyleCham(paddingTop: ptMach, paddingBottom: pbMach))
                                               .AlignCenter()
                                               .Text(text =>
                                               {
                                                   text.Span($"{mach} ").FontColor("#FF0000").FontSize(8);
                                                   text.Span("●").FontColor("#FF0000");
                                               });
                                        });
                                    }
                                    else if (!(mach > 120 && mach <= 140) && (nhietDo > 39 && nhietDo <= 40))
                                    {
                                        table.Cell().AlignCenter().Border(1).Row(row =>
                                        {
                                            row.RelativeItem(1)
                                               .Element(c => c.CellStyleCham(paddingTop: ptNhiet, paddingBottom: pbNhiet)).AlignCenter().Text(text =>
                                               {
                                                   text.Span($"{nhietDo}").FontColor("#0000FF").FontSize(8);
                                                   text.Span("●").FontColor("#0000FF");
                                               });
                                        });
                                    }
                                    else if ((mach > 120 && mach <= 140) && (nhietDo > 39 && nhietDo <= 40))
                                    {
                                        table.Cell().AlignCenter().Border(1).Element(container =>
                                        {
                                            container.Row(row =>
                                            {
                                                row.RelativeItem(1)
                                                   .Element(c => c.CellStyleCham(paddingTop: ptMach, paddingBottom: pbMach)).AlignCenter().Text(text =>
                                                   {
                                                       text.Span($"{mach} ").FontColor("#FF0000").FontSize(8);
                                                       text.Span("●").FontColor("#FF0000");
                                                   });

                                                row.RelativeItem(1)
                                                   .Element(c => c.CellStyleCham(paddingTop: ptNhiet, paddingBottom: pbNhiet)).AlignCenter().Text(text =>
                                                   {
                                                       text.Span($"{nhietDo}").FontColor("#0000FF").FontSize(8);
                                                       text.Span("●").FontColor("#0000FF");
                                                   });
                                            });
                                        });
                                    }
                                }
                                else
                                {
                                    table.Cell().Element(CellStyleBigSize).AlignCenter().Text("");
                                }
                            }

                            table.Cell().Element(CellStyleTop).Height(25).AlignCenter().Text("120");
                            table.Cell().Element(CellStyleTop).Height(25).AlignCenter().Text("39");
                            for (int i = 0; i < listMach.Count && i < listNhietDo.Count; i++) // tác hàm ra
                            {
                                double mach = double.TryParse(listMach[i], out var mVal) ? mVal : 0;
                                double nhietDo = double.TryParse(listNhietDo[i], out var nVal) ? nVal : 0;

                                var (ptMach, pbMach) = TinhPadding010_Auto(mach, step: 20);
                                var (ptNhiet, pbNhiet) = TinhPadding010_Auto(nhietDo, step: 1);

                                if ((mach > 100 && mach <= 120) || (nhietDo > 38 && nhietDo <= 39))
                                {
                                    if ((mach > 100 && mach <= 120) && !(nhietDo > 38 && nhietDo <= 39))
                                    {
                                        table.Cell().AlignCenter().Border(1).Row(row =>
                                        {
                                            row.RelativeItem(1)
                                               .Element(c => c.CellStyleCham(paddingTop: ptMach, paddingBottom: pbMach))
                                               .AlignCenter()
                                               .Text(text =>
                                               {
                                                   text.Span($"{mach} ").FontColor("#FF0000").FontSize(8);
                                                   text.Span("●").FontColor("#FF0000");
                                               });
                                        });
                                    }
                                    else if (!(mach > 100 && mach <= 120) && (nhietDo > 38 && nhietDo <= 39))
                                    {
                                        table.Cell().AlignCenter().Border(1).Row(row =>
                                        {
                                            row.RelativeItem(1)
                                               .Element(c => c.CellStyleCham(paddingTop: ptNhiet, paddingBottom: pbNhiet)).AlignCenter().Text(text =>
                                               {
                                                   text.Span($"{nhietDo}").FontColor("#0000FF").FontSize(8);
                                                   text.Span("●").FontColor("#0000FF");
                                               });
                                        });
                                    }
                                    else if ((mach > 100 && mach <= 120) && (nhietDo > 38 && nhietDo <= 39))
                                    {
                                        table.Cell().AlignCenter().Border(1).Element(container =>
                                        {
                                            container.Row(row =>
                                            {
                                                row.RelativeItem(1)
                                                   .Element(c => c.CellStyleCham(paddingTop: ptMach, paddingBottom: pbMach)).AlignCenter().Text(text =>
                                                   {
                                                       text.Span($"{mach} ").FontColor("#FF0000").FontSize(8);
                                                       text.Span("●").FontColor("#FF0000");
                                                   });

                                                row.RelativeItem(1)
                                                   .Element(c => c.CellStyleCham(paddingTop: ptNhiet, paddingBottom: pbNhiet)).AlignCenter().Text(text =>
                                                   {
                                                       text.Span($"{nhietDo}").FontColor("#0000FF").FontSize(8);
                                                       text.Span("●").FontColor("#0000FF");
                                                   });
                                            });
                                        });
                                    }
                                }
                                else
                                {
                                    table.Cell().Element(CellStyleBigSize).AlignCenter().Text("");
                                }
                            }

                            table.Cell().Element(CellStyleTop).Height(25).AlignCenter().Text("100");
                            table.Cell().Element(CellStyleTop).Height(25).AlignCenter().Text("38");
                            for (int i = 0; i < listMach.Count && i < listNhietDo.Count; i++) // tác hàm ra
                            {
                                double mach = double.TryParse(listMach[i], out var mVal) ? mVal : 0;
                                double nhietDo = double.TryParse(listNhietDo[i], out var nVal) ? nVal : 0;

                                var (ptMach, pbMach) = TinhPadding010_Auto(mach, step: 20);
                                var (ptNhiet, pbNhiet) = TinhPadding010_Auto(nhietDo, step: 1);

                                if ((mach > 80 && mach <= 100) || (nhietDo > 37 && nhietDo <= 38))
                                {
                                    if ((mach > 80 && mach <= 100) && !(nhietDo > 37 && nhietDo <= 38))
                                    {
                                        table.Cell().AlignCenter().Border(1).Row(row =>
                                        {
                                            row.RelativeItem(1)
                                               .Element(c => c.CellStyleCham(paddingTop: ptMach, paddingBottom: pbMach))
                                               .AlignCenter()
                                               .Text(text =>
                                               {
                                                   text.Span($"{mach} ").FontColor("#FF0000").FontSize(8);
                                                   text.Span("●").FontColor("#FF0000");
                                               });
                                        });
                                    }
                                    else if (!(mach > 80 && mach <= 100) && (nhietDo > 37 && nhietDo <= 38))
                                    {
                                        table.Cell().AlignCenter().Border(1).Row(row =>
                                        {
                                            row.RelativeItem(1)
                                               .Element(c => c.CellStyleCham(paddingTop: ptNhiet, paddingBottom: pbNhiet)).AlignCenter().Text(text =>
                                               {
                                                   text.Span($"{nhietDo}").FontColor("#0000FF").FontSize(8);
                                                   text.Span("●").FontColor("#0000FF");
                                               });
                                        });
                                    }
                                    else if ((mach > 80 && mach <= 100) && (nhietDo > 37 && nhietDo <= 38))
                                    {
                                        table.Cell().AlignCenter().Border(1).Element(container =>
                                        {
                                            container.Row(row =>
                                            {
                                                row.RelativeItem(1)
                                                   .Element(c => c.CellStyleCham(paddingTop: ptMach, paddingBottom: pbMach)).AlignCenter().Text(text =>
                                                   {
                                                       text.Span($"{mach} ").FontColor("#FF0000").FontSize(8);
                                                       text.Span("●").FontColor("#FF0000");
                                                   });

                                                row.RelativeItem(1)
                                                   .Element(c => c.CellStyleCham(paddingTop: ptNhiet, paddingBottom: pbNhiet)).AlignCenter().Text(text =>
                                                   {
                                                       text.Span($"{nhietDo}").FontColor("#0000FF").FontSize(8);
                                                       text.Span("●").FontColor("#0000FF");
                                                   });
                                            });
                                        });
                                    }
                                }
                                else
                                {
                                    table.Cell().Element(CellStyleBigSize).AlignCenter().Text("");
                                }
                            }

                            table.Cell().Element(CellStyleTop).Height(25).AlignCenter().Text("80");
                            table.Cell().Element(CellStyleTop).Height(25).AlignCenter().Text("37");
                            for (int i = 0; i < listMach.Count && i < listNhietDo.Count; i++) // tác hàm ra
                            {
                                double mach = double.TryParse(listMach[i], out var mVal) ? mVal : 0;
                                double nhietDo = double.TryParse(listNhietDo[i], out var nVal) ? nVal : 0;

                                var (ptMach, pbMach) = TinhPadding010_Auto(mach, step: 20);
                                var (ptNhiet, pbNhiet) = TinhPadding010_Auto(nhietDo, step: 1);

                                if ((mach > 60 && mach <= 80) || (nhietDo > 36 && nhietDo <= 37))
                                {
                                    if ((mach > 60 && mach <= 80) && !(nhietDo > 36 && nhietDo <= 37))
                                    {
                                        table.Cell().AlignCenter().Border(1).Row(row =>
                                        {
                                            row.RelativeItem(1)
                                               .Element(c => c.CellStyleCham(paddingTop: ptMach, paddingBottom: pbMach))
                                               .AlignCenter()
                                               .Text(text =>
                                               {
                                                   text.Span($"{mach} ").FontColor("#FF0000").FontSize(8);
                                                   text.Span("●").FontColor("#FF0000");
                                               });
                                        });
                                    }
                                    else if (!(mach > 60 && mach <= 80) && (nhietDo > 36 && nhietDo <= 37))
                                    {
                                        table.Cell().AlignCenter().Border(1).Row(row =>
                                        {
                                            row.RelativeItem(1)
                                               .Element(c => c.CellStyleCham(paddingTop: ptNhiet, paddingBottom: pbNhiet)).AlignCenter().Text(text =>
                                               {
                                                   text.Span($"{nhietDo}").FontColor("#0000FF").FontSize(8);
                                                   text.Span("●").FontColor("#0000FF");
                                               });
                                        });
                                    }
                                    else if ((mach > 60 && mach <= 80) && (nhietDo > 36 && nhietDo <= 37))
                                    {
                                        table.Cell().AlignCenter().Border(1).Element(container =>
                                        {
                                            container.Row(row =>
                                            {
                                                row.RelativeItem(1)
                                                   .Element(c => c.CellStyleCham(paddingTop: ptMach, paddingBottom: pbMach)).AlignCenter().Text(text =>
                                                   {
                                                       text.Span($"{mach} ").FontColor("#FF0000").FontSize(8);
                                                       text.Span("●").FontColor("#FF0000");
                                                   });

                                                row.RelativeItem(1)
                                                   .Element(c => c.CellStyleCham(paddingTop: ptNhiet, paddingBottom: pbNhiet)).AlignCenter().Text(text =>
                                                   {
                                                       text.Span($"{nhietDo}").FontColor("#0000FF").FontSize(8);
                                                       text.Span("●").FontColor("#0000FF");
                                                   });
                                            });
                                        });
                                    }
                                }
                                else
                                {
                                    table.Cell().Element(CellStyleBigSize).AlignCenter().Text("");
                                }
                            }

                            table.Cell().Element(CellStyleTop).Height(25).AlignCenter().Text("60");
                            table.Cell().Element(CellStyleTop).Height(25).AlignCenter().Text("36");
                            for (int i = 0; i < listMach.Count && i < listNhietDo.Count; i++) // tác hàm ra
                            {
                                double mach = double.TryParse(listMach[i], out var mVal) ? mVal : 0;
                                double nhietDo = double.TryParse(listNhietDo[i], out var nVal) ? nVal : 0;

                                var (ptMach, pbMach) = TinhPadding010_Auto(mach, step: 20);
                                var (ptNhiet, pbNhiet) = TinhPadding010_Auto(nhietDo, step: 1);

                                if ((mach > 40 && mach <= 60) || (nhietDo > 35 && nhietDo <= 36))
                                {
                                    if ((mach > 40 && mach <= 60) && !(nhietDo > 35 && nhietDo <= 36))
                                    {
                                        table.Cell().AlignCenter().Border(1).Row(row =>
                                        {
                                            row.RelativeItem(1)
                                               .Element(c => c.CellStyleCham(paddingTop: ptMach, paddingBottom: pbMach))
                                               .AlignCenter()
                                               .Text(text =>
                                               {
                                                   text.Span($"{mach} ").FontColor("#FF0000").FontSize(8);
                                                   text.Span("●").FontColor("#FF0000");
                                               });
                                        });
                                    }
                                    else if (!(mach > 40 && mach <= 60) && (nhietDo > 35 && nhietDo <= 36))
                                    {
                                        table.Cell().AlignCenter().Border(1).Row(row =>
                                        {
                                            row.RelativeItem(1)
                                               .Element(c => c.CellStyleCham(paddingTop: ptNhiet, paddingBottom: pbNhiet)).AlignCenter().Text(text =>
                                               {
                                                   text.Span($"{nhietDo}").FontColor("#0000FF").FontSize(8);
                                                   text.Span("●").FontColor("#0000FF");
                                               });
                                        });
                                    }
                                    else if ((mach > 40 && mach <= 60) && (nhietDo > 35 && nhietDo <= 36))
                                    {
                                        table.Cell().AlignCenter().Border(1).Element(container =>
                                        {
                                            container.Row(row =>
                                            {
                                                row.RelativeItem(1)
                                                   .Element(c => c.CellStyleCham(paddingTop: ptMach, paddingBottom: pbMach)).AlignCenter().Text(text =>
                                                   {
                                                       text.Span($"{mach} ").FontColor("#FF0000").FontSize(8);
                                                       text.Span("●").FontColor("#FF0000");
                                                   });

                                                row.RelativeItem(1)
                                                   .Element(c => c.CellStyleCham(paddingTop: ptNhiet, paddingBottom: pbNhiet)).AlignCenter().Text(text =>
                                                   {
                                                       text.Span($"{nhietDo}").FontColor("#0000FF").FontSize(8);
                                                       text.Span("●").FontColor("#0000FF");
                                                   });
                                            });
                                        });
                                    }
                                }
                                else
                                {
                                    table.Cell().Element(CellStyleBigSize).AlignCenter().Text("");
                                }
                            }

                            table.Cell().Element(CellStyleTop).Height(25).AlignCenter().Text("40");
                            table.Cell().Element(CellStyleTop).Height(25).AlignCenter().Text("35");
                            for (int i = 0; i < listMach.Count && i < listNhietDo.Count; i++) // tác hàm ra
                            {
                                double mach = double.TryParse(listMach[i], out var mVal) ? mVal : 0;
                                double nhietDo = double.TryParse(listNhietDo[i], out var nVal) ? nVal : 0;

                                var (ptMach, pbMach) = TinhPadding010_Auto(mach, step: 20);
                                var (ptNhiet, pbNhiet) = TinhPadding010_Auto(nhietDo, step: 1);

                                if ((mach > 20 && mach <= 40) || (nhietDo > 34 && nhietDo <= 35))
                                {
                                    if ((mach > 20 && mach <= 40) && !(nhietDo > 34 && nhietDo <= 35))
                                    {
                                        table.Cell().AlignCenter().Border(1).Row(row =>
                                        {
                                            row.RelativeItem(1)
                                               .Element(c => c.CellStyleCham(paddingTop: ptMach, paddingBottom: pbMach))
                                               .AlignCenter()
                                               .Text(text =>
                                               {
                                                   text.Span($"{mach} ").FontColor("#FF0000").FontSize(8);
                                                   text.Span("●").FontColor("#FF0000");
                                               });
                                        });
                                    }
                                    else if (!(mach > 20 && mach <= 40) && (nhietDo > 34 && nhietDo <= 35))
                                    {
                                        table.Cell().AlignCenter().Border(1).Row(row =>
                                        {
                                            row.RelativeItem(1)
                                               .Element(c => c.CellStyleCham(paddingTop: ptNhiet, paddingBottom: pbNhiet)).AlignCenter().Text(text =>
                                               {
                                                   text.Span($"{nhietDo}").FontColor("#0000FF").FontSize(8);
                                                   text.Span("●").FontColor("#0000FF");
                                               });
                                        });
                                    }
                                    else if ((mach > 20 && mach <= 40) && (nhietDo > 34 && nhietDo <= 35))
                                    {
                                        table.Cell().AlignCenter().Border(1).Element(container =>
                                        {
                                            container.Row(row =>
                                            {
                                                row.RelativeItem(1)
                                                   .Element(c => c.CellStyleCham(paddingTop: ptMach, paddingBottom: pbMach)).AlignCenter().Text(text =>
                                                   {
                                                       text.Span($"{mach} ").FontColor("#FF0000").FontSize(8);
                                                       text.Span("●").FontColor("#FF0000");
                                                   });

                                                row.RelativeItem(1)
                                                   .Element(c => c.CellStyleCham(paddingTop: ptNhiet, paddingBottom: pbNhiet)).AlignCenter().Text(text =>
                                                   {
                                                       text.Span($"{nhietDo}").FontColor("#0000FF").FontSize(8);
                                                       text.Span("●").FontColor("#0000FF");
                                                   });
                                            });
                                        });
                                    }
                                }
                                else
                                {
                                    table.Cell().Element(CellStyleBigSize).AlignCenter().Text("");
                                }
                            }

                            table.Cell().ColumnSpan(2).Element(CellStyle).AlignLeft().Text("1. Huyết áp (mmHg)");
                            foreach (var huyetAp in listHuyetAp)
                            {
                                table.Cell().Element(CellStyle).AlignCenter().Text($"{huyetAp}");
                            }
                            table.Cell().ColumnSpan(2).Element(CellStyle).AlignLeft().Text("2. Cân nặng (Kg)");
                            foreach (var canNang in listCanNang)
                            {
                                table.Cell().Element(CellStyle).AlignCenter().Text($"{canNang}");
                            }
                            table.Cell().ColumnSpan(2).Element(CellStyle).AlignLeft().Text("3. Nhịp thở (lần/ph)");
                            foreach (var nhipTho in listNhipTho)
                            {
                                table.Cell().Element(CellStyle).AlignCenter().Text($"{nhipTho}");
                            }
                            table.Cell().ColumnSpan(2).Element(CellStyle).AlignLeft().Text("4. ");
                            for (int i = 0; i < _data.SinhHieus.Count(); i++)
                            {
                                table.Cell().Element(CellStyle).AlignCenter().Text("");
                            }
                            table.Cell().ColumnSpan(2).Element(CellStyle).AlignLeft().Text("5. ");
                            for (int i = 0; i < _data.SinhHieus.Count(); i++)
                            {
                                table.Cell().Element(CellStyle).AlignCenter().Text("");
                            }
                            table.Cell().ColumnSpan(2).Element(CellStyle).AlignLeft().Text("6. ");
                            for (int i = 0; i < _data.SinhHieus.Count(); i++)
                            {
                                table.Cell().Element(CellStyle).AlignCenter().Text("");
                            }
                            table.Cell().ColumnSpan(2).Element(CellStyle).AlignLeft().Element(e =>
                            {
                                e.Column(column =>
                                {
                                    column.Item().Text("Y tá ĐD");
                                    column.Item().Text("Ký tên tại đây");
                                    column.Item().Height(40).Text("");
                                });
                            });
                            for (int i = 0; i < _data.SinhHieus.Count(); i++)
                            {
                                table.Cell().Element(CellStyle).AlignCenter().Text("");
                            }
                        });
                    });
                });
            });

            static IContainer CellStyle(IContainer container) =>
            container
            .Border(1)
            .Padding(5)
            .AlignMiddle()
            .DefaultTextStyle(x => x.FontSize(10));

            static IContainer CellStyleTop(IContainer container) =>
            container
            .Border(1)
            .PaddingLeft(5)
            .PaddingRight(5)
            .PaddingBottom(10)
            .AlignTop()
            .DefaultTextStyle(x => x.FontSize(10));

            static IContainer CellStyleBigSize(IContainer container) =>
            container
            .Border(1)
            .Padding(5)
            .AlignMiddle()
            .DefaultTextStyle(x => x.FontSize(10));
        }
    }

    public static class PdfExtensions
    {
        // Đổi thành extension method
        public static IContainer CellStyleCham(
            this IContainer container,
            int paddingTop,
            int paddingBottom)
        {
            return container
                .PaddingLeft(5)
                .PaddingRight(5)
                .TranslateY(5 - paddingBottom)
                .PaddingTop(paddingTop)
                .PaddingBottom(paddingBottom)
                .DefaultTextStyle(x => x.FontSize(15));
        }
    }
}


