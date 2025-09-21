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
                                info.Item().Text(text =>
                                {
                                    text.Span("Khoa: ").SemiBold().FontSize(10);
                                    text.Span($"{_data?.ThongTinBenhNhan?.TenKhoa ?? ""}").FontSize(10);
                                });
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
                                var ngaySinh = _data?.ThongTinBenhNhan?.NgaySinh?.ToString("dd-MM-yyyy");
                                text.Span("Ngày sinh: ").SemiBold().FontSize(10);
                                text.Span($"{ngaySinh}").FontSize(10);
                            });

                        row.RelativeItem().AlignRight()
                            .Text(text =>
                            {
                                text.Span("Giới tính: ").SemiBold().FontSize(10);
                                text.Span($"{_data?.ThongTinBenhNhan?.GioiTinh ?? ""}").FontSize(10);
                            });
                    });

                    col.Item().PaddingLeft(5).AlignLeft()
                        .Text(text =>
                        {
                            text.Span("Địa chỉ: ").SemiBold().FontSize(10);
                            text.Span($"{_data?.ThongTinBenhNhan?.DiaChi ?? ""}").FontSize(10);
                        });

                    col.Item().Padding(5).Row(row =>
                    {

                        row.RelativeItem().AlignLeft()
                            .Text(text =>
                            {
                                text.Span("Số giường: ").SemiBold().FontSize(10);
                                text.Span($"{_data?.ThongTinBenhNhan?.TenGiuong}").FontSize(10);
                            });

                        row.RelativeItem().AlignLeft()
                            .Text(text =>
                            {
                                text.Span("Buồng: ").SemiBold().FontSize(10);
                                text.Span($"{_data?.ThongTinBenhNhan?.TenPhong ?? ""}").FontSize(10);
                            });
                    });

                    col.Item().AlignLeft().PaddingLeft(5)
                        .Text(text =>
                        {
                            text.Span("Chẩn đoán: ").SemiBold().FontSize(10);
                            text.Span($"{_data?.ThongTinBenhNhan?.ChanDoan ?? ""}").FontSize(10);
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

                    double pTop = offsetRatio * 9.0;
                    double pBot = 9.0 - pTop;

                    if (roundToInt)
                    {
                        pTop = Math.Round(pTop);
                        pBot = 9 - pTop;
                    }

                    return ((int)pTop, (int)pBot);
                }

                page.Content().PaddingVertical(6).Element(e =>
                {
                    e.Column(column =>
                    {

                        int maxColumnsPerPage = 20;
                        int totalColumns = _data.SinhHieus.Count();
                        int totalPages = (int)Math.Ceiling((double)totalColumns / maxColumnsPerPage);

                        for (int page = 0; page < totalPages; page++)
                        {
                            int startIndex = page*maxColumnsPerPage;
                            int endIndex = Math.Min(startIndex + maxColumnsPerPage, totalColumns);
                            column.Item().Table(table =>
                            {
                                table.ColumnsDefinition(columns =>
                                {
                                    columns.ConstantColumn(26);
                                    columns.ConstantColumn(28);
                                    for (int i = startIndex; i < endIndex; i++)
                                    {
                                        columns.RelativeColumn();
                                    }
                                });

                                table.Cell().ColumnSpan(2).Element(CellStyle).AlignCenter().Text("Ngày tháng");
                                int j = startIndex;
                                while (j < endIndex && j < listNgay.Count)
                                {
                                    var currentNgay = listNgay[j];

                                    if (j + 1 < endIndex && j + 1 < listNgay.Count && listNgay[j + 1] == currentNgay)
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
                                foreach (var gio in listGio.Skip(startIndex).Take(endIndex - startIndex))
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
                                for (int i = startIndex; i < endIndex; i++)
                                {
                                    table.Cell().Element(CellStyle).AlignCenter().Text("");
                                }

                                table.Cell().Element(CellStyleTop).Height(23).AlignCenter().Text("160");
                                table.Cell().Element(CellStyleTop).Height(23).AlignCenter().Text("41");
                                for (int i = startIndex; i < endIndex; i++) // tác hàm ra
                                {
                                    double mach = double.TryParse(listMach[i], out var mVal) ? mVal : 0;
                                    double nhietDo = double.TryParse(listNhietDo[i], out var nVal) ? nVal : 0;

                                    var (ptMach, pbMach) = TinhPadding010_Auto(mach, step: 20);
                                    var (ptNhiet, pbNhiet) = TinhPadding010_Auto(nhietDo, step: 1);

                                    if ((mach > 140 && mach <= 160) || (nhietDo > 40 && nhietDo <= 41))
                                    {
                                        if ((mach > 140 && mach <= 160) && !(nhietDo > 40 && nhietDo <= 41))
                                        {
                                            RenderMachCell(table, mach, ptMach, pbMach);
                                        }
                                        else if (!(mach > 140 && mach <= 160) && (nhietDo > 40 && nhietDo <= 41))
                                        {
                                            RenderNhietDoCell(table, nhietDo, ptNhiet, pbNhiet);
                                        }
                                        else if ((mach > 140 && mach <= 160) && (nhietDo > 40 && nhietDo <= 41))
                                        {
                                            RenderMachNhietDoCell(table, mach, nhietDo, ptMach, pbMach, ptNhiet, pbNhiet);
                                        }
                                    }
                                    else
                                    {
                                        table.Cell().Element(CellStyleBigSize).AlignCenter().Text("");
                                    }
                                }

                                table.Cell().Element(CellStyleTop).Height(23).AlignCenter().Text("140");
                                table.Cell().Element(CellStyleTop).Height(23).AlignCenter().Text("40");
                                for (int i = startIndex; i < endIndex; i++) // tác hàm ra
                                {
                                    double mach = double.TryParse(listMach[i], out var mVal) ? mVal : 0;
                                    double nhietDo = double.TryParse(listNhietDo[i], out var nVal) ? nVal : 0;

                                    var (ptMach, pbMach) = TinhPadding010_Auto(mach, step: 20);
                                    var (ptNhiet, pbNhiet) = TinhPadding010_Auto(nhietDo, step: 1);

                                    if ((mach > 120 && mach <= 140) || (nhietDo > 39 && nhietDo <= 40))
                                    {
                                        if ((mach > 120 && mach <= 140) && !(nhietDo > 39 && nhietDo <= 40))
                                        {
                                            RenderMachCell(table, mach, ptMach, pbMach);
                                        }
                                        else if (!(mach > 120 && mach <= 140) && (nhietDo > 39 && nhietDo <= 40))
                                        {
                                            RenderNhietDoCell(table, nhietDo, ptNhiet, pbNhiet);
                                        }
                                        else if ((mach > 120 && mach <= 140) && (nhietDo > 39 && nhietDo <= 40))
                                        {
                                            RenderMachNhietDoCell(table, mach, nhietDo, ptMach, pbMach, ptNhiet, pbNhiet);
                                        }
                                    }
                                    else
                                    {
                                        table.Cell().Element(CellStyleBigSize).AlignCenter().Text("");
                                    }
                                }

                                table.Cell().Element(CellStyleTop).Height(23).AlignCenter().Text("120");
                                table.Cell().Element(CellStyleTop).Height(23).AlignCenter().Text("39");
                                for (int i = startIndex; i < endIndex; i++) // tác hàm ra
                                {
                                    double mach = double.TryParse(listMach[i], out var mVal) ? mVal : 0;
                                    double nhietDo = double.TryParse(listNhietDo[i], out var nVal) ? nVal : 0;

                                    var (ptMach, pbMach) = TinhPadding010_Auto(mach, step: 20);
                                    var (ptNhiet, pbNhiet) = TinhPadding010_Auto(nhietDo, step: 1);

                                    if ((mach > 100 && mach <= 120) || (nhietDo > 38 && nhietDo <= 39))
                                    {
                                        if ((mach > 100 && mach <= 120) && !(nhietDo > 38 && nhietDo <= 39))
                                        {
                                            RenderMachCell(table, mach, ptMach, pbMach);
                                        }
                                        else if (!(mach > 100 && mach <= 120) && (nhietDo > 38 && nhietDo <= 39))
                                        {
                                            RenderNhietDoCell(table, nhietDo, ptNhiet, pbNhiet);
                                        }
                                        else if ((mach > 100 && mach <= 120) && (nhietDo > 38 && nhietDo <= 39))
                                        {
                                            RenderMachNhietDoCell(table, mach, nhietDo, ptMach, pbMach, ptNhiet, pbNhiet);
                                        }
                                    }
                                    else
                                    {
                                        table.Cell().Element(CellStyleBigSize).AlignCenter().Text("");
                                    }
                                }

                                table.Cell().Element(CellStyleTop).Height(23).AlignCenter().Text("100");
                                table.Cell().Element(CellStyleTop).Height(23).AlignCenter().Text("38");
                                for (int i = startIndex; i < endIndex; i++) // tác hàm ra
                                {
                                    double mach = double.TryParse(listMach[i], out var mVal) ? mVal : 0;
                                    double nhietDo = double.TryParse(listNhietDo[i], out var nVal) ? nVal : 0;

                                    var (ptMach, pbMach) = TinhPadding010_Auto(mach, step: 20);
                                    var (ptNhiet, pbNhiet) = TinhPadding010_Auto(nhietDo, step: 1);

                                    if ((mach > 80 && mach <= 100) || (nhietDo > 37 && nhietDo <= 38))
                                    {
                                        if ((mach > 80 && mach <= 100) && !(nhietDo > 37 && nhietDo <= 38))
                                        {
                                            RenderMachCell(table, mach, ptMach, pbMach);
                                        }
                                        else if (!(mach > 80 && mach <= 100) && (nhietDo > 37 && nhietDo <= 38))
                                        {
                                            RenderNhietDoCell(table, nhietDo, ptNhiet, pbNhiet);
                                        }
                                        else if ((mach > 80 && mach <= 100) && (nhietDo > 37 && nhietDo <= 38))
                                        {
                                            RenderMachNhietDoCell(table, mach, nhietDo, ptMach, pbMach, ptNhiet, pbNhiet);
                                        }
                                    }
                                    else
                                    {
                                        table.Cell().Element(CellStyleBigSize).AlignCenter().Text("");
                                    }
                                }

                                table.Cell().Element(CellStyleTop).Height(23).AlignCenter().Text("80");
                                table.Cell().Element(CellStyleTop).Height(23).AlignCenter().Text("37");
                                for (int i = startIndex; i < endIndex; i++) // tác hàm ra
                                {
                                    double mach = double.TryParse(listMach[i], out var mVal) ? mVal : 0;
                                    double nhietDo = double.TryParse(listNhietDo[i], out var nVal) ? nVal : 0;

                                    var (ptMach, pbMach) = TinhPadding010_Auto(mach, step: 20);
                                    var (ptNhiet, pbNhiet) = TinhPadding010_Auto(nhietDo, step: 1);

                                    if ((mach > 60 && mach <= 80) || (nhietDo > 36 && nhietDo <= 37))
                                    {
                                        if ((mach > 60 && mach <= 80) && !(nhietDo > 36 && nhietDo <= 37))
                                        {
                                            RenderMachCell(table, mach, ptMach, pbMach);
                                        }
                                        else if (!(mach > 60 && mach <= 80) && (nhietDo > 36 && nhietDo <= 37))
                                        {
                                            RenderNhietDoCell(table, nhietDo, ptNhiet, pbNhiet);
                                        }
                                        else if ((mach > 60 && mach <= 80) && (nhietDo > 36 && nhietDo <= 37))
                                        {
                                            RenderMachNhietDoCell(table, mach, nhietDo, ptMach, pbMach, ptNhiet, pbNhiet);
                                        }
                                    }
                                    else
                                    {
                                        table.Cell().Element(CellStyleBigSize).AlignCenter().Text("");
                                    }
                                }

                                table.Cell().Element(CellStyleTop).Height(23).AlignCenter().Text("60");
                                table.Cell().Element(CellStyleTop).Height(23).AlignCenter().Text("36");
                                for (int i = startIndex; i < endIndex; i++) // tác hàm ra
                                {
                                    double mach = double.TryParse(listMach[i], out var mVal) ? mVal : 0;
                                    double nhietDo = double.TryParse(listNhietDo[i], out var nVal) ? nVal : 0;

                                    var (ptMach, pbMach) = TinhPadding010_Auto(mach, step: 20);
                                    var (ptNhiet, pbNhiet) = TinhPadding010_Auto(nhietDo, step: 1);

                                    if ((mach > 40 && mach <= 60) || (nhietDo > 35 && nhietDo <= 36))
                                    {
                                        if ((mach > 40 && mach <= 60) && !(nhietDo > 35 && nhietDo <= 36))
                                        {
                                            RenderMachCell(table, mach, ptMach, pbMach);
                                        }
                                        else if (!(mach > 40 && mach <= 60) && (nhietDo > 35 && nhietDo <= 36))
                                        {
                                            RenderNhietDoCell(table, nhietDo, ptNhiet, pbNhiet);
                                        }
                                        else if ((mach > 40 && mach <= 60) && (nhietDo > 35 && nhietDo <= 36))
                                        {
                                            RenderMachNhietDoCell(table, mach, nhietDo, ptMach, pbMach, ptNhiet, pbNhiet);
                                        }
                                    }
                                    else
                                    {
                                        table.Cell().Element(CellStyleBigSize).AlignCenter().Text("");
                                    }
                                }

                                table.Cell().Element(CellStyleTop).Height(23).AlignCenter().Text("40");
                                table.Cell().Element(CellStyleTop).Height(23).AlignCenter().Text("35");
                                for (int i = startIndex; i < endIndex; i++) // tác hàm ra
                                {
                                    double mach = double.TryParse(listMach[i], out var mVal) ? mVal : 0;
                                    double nhietDo = double.TryParse(listNhietDo[i], out var nVal) ? nVal : 0;

                                    var (ptMach, pbMach) = TinhPadding010_Auto(mach, step: 20);
                                    var (ptNhiet, pbNhiet) = TinhPadding010_Auto(nhietDo, step: 1);

                                    if ((mach > 20 && mach <= 40) || (nhietDo > 34 && nhietDo <= 35))
                                    {
                                        if ((mach > 20 && mach <= 40) && !(nhietDo > 34 && nhietDo <= 35))
                                        {
                                            RenderMachCell(table, mach, ptMach, pbMach);
                                        }
                                        else if (!(mach > 20 && mach <= 40) && (nhietDo > 34 && nhietDo <= 35))
                                        {
                                            RenderNhietDoCell(table, nhietDo, ptNhiet, pbNhiet);
                                        }
                                        else if ((mach > 20 && mach <= 40) && (nhietDo > 34 && nhietDo <= 35))
                                        {
                                            RenderMachNhietDoCell(table, mach, nhietDo, ptMach, pbMach, ptNhiet, pbNhiet);
                                        }
                                    }
                                    else
                                    {
                                        table.Cell().Element(CellStyleBigSize).AlignCenter().Text("");
                                    }
                                }

                                table.Cell().ColumnSpan(2).Element(CellStyle).AlignLeft().Element(e =>
                                {
                                    e.Column(column =>
                                    {
                                        column.Item().Text("1. Huyết áp");
                                        column.Item().Text("(mmHg)");
                                    });
                                });
                                foreach (var huyetAp in listHuyetAp.Skip(startIndex).Take(endIndex - startIndex))
                                {
                                    table.Cell().Element(CellStyle).AlignCenter().Text($"{huyetAp}");
                                }

                                table.Cell().ColumnSpan(2).Element(CellStyle).AlignLeft().Element(e =>
                                {
                                    e.Column(column =>
                                    {
                                        column.Item().Text("Cân nặng");
                                        column.Item().Text("(Kg)");
                                    });
                                });
                                foreach (var canNang in listCanNang.Skip(startIndex).Take(endIndex - startIndex))
                                {
                                    table.Cell().Element(CellStyle).AlignCenter().Text($"{canNang}");
                                }

                                table.Cell().ColumnSpan(2).Element(CellStyle).AlignLeft().Element(e =>
                                {
                                    e.Column(column =>
                                    {
                                        column.Item().Text("Nhịp thở");
                                        column.Item().Text("(Lần/ph)");
                                    });
                                });
                                foreach (var nhipTho in listNhipTho.Skip(startIndex).Take(endIndex - startIndex))
                                {
                                    table.Cell().Element(CellStyle).AlignCenter().Text($"{nhipTho}");
                                }

                                table.Cell().ColumnSpan(2).Element(CellStyle).AlignLeft().Text("4. ");
                                for (int i = startIndex; i < endIndex; i++)
                                {
                                    table.Cell().Element(CellStyle).AlignCenter().Text("");
                                }

                                table.Cell().ColumnSpan(2).Element(CellStyle).AlignLeft().Text("5. ");
                                for (int i = startIndex; i < endIndex; i++)
                                {
                                    table.Cell().Element(CellStyle).AlignCenter().Text("");
                                }

                                table.Cell().ColumnSpan(2).Element(CellStyle).AlignLeft().Text("6. ");
                                for (int i = startIndex; i < endIndex; i++)
                                {
                                    table.Cell().Element(CellStyle).AlignCenter().Text("");
                                }
                                table.Cell().ColumnSpan(2).Element(CellStyle).AlignLeft().Element(e =>
                                {
                                    e.Column(column =>
                                    {
                                        column.Item().Text("Y tá - ĐD");
                                        column.Item().Text("Ký và ghi tên");
                                        column.Item().Height(40).Text("");
                                    });
                                });
                                for (int i = startIndex; i < endIndex; i++)
                                {
                                    table.Cell().Element(CellStyle).AlignCenter().Text("");
                                }
                            });
                            if (page < totalPages - 1)
                            {
                                column.Item().PageBreak(); // ✅ CHÍNH XÁC Ở ĐÂY
                            }
                        }
                    }); 
                });
            });

            static IContainer CellStyle(IContainer container) =>
            container
            .Border(1)
            .Padding(3)
            .AlignMiddle()
            .DefaultTextStyle(x => x.FontSize(8));

            static IContainer CellStyleTop(IContainer container) =>
            container
            .Border(1)
            .PaddingLeft(5)
            .PaddingRight(5)
            .PaddingBottom(10)
            .AlignTop()
            .DefaultTextStyle(x => x.FontSize(8));

            static IContainer CellStyleBigSize(IContainer container) =>
            container
            .Border(1)
            .Padding(5)
            .AlignMiddle()
            .DefaultTextStyle(x => x.FontSize(8));

            static void RenderMachCell(TableDescriptor table, double mach, int ptMach, int pbMach)
            {
                table.Cell().Border(1)
                .Row(row => {
                    row.RelativeItem(1)
                       .AlignCenter()
                       .PaddingRight(4)
                       .Element(c => c.CellStyleCham(paddingTop: ptMach, paddingBottom: pbMach))
                       .Column(col =>
                       {
                           col.Spacing(0);
                           col.Item().Unconstrained().Text($"{mach}")
                               .FontColor("#FF0000")
                               .FontSize(6)
                               .LineHeight(0.3f);
                           col.Item().Unconstrained().Text("●")
                               .FontColor("#FF0000")
                               .FontSize(11)
                               .LineHeight(0.9f);
                       });
                });
            }

            static void RenderNhietDoCell(TableDescriptor table, double nhietDo, int ptNhiet, int pbNhiet)
            {
                table.Cell().Border(1)
                .Row(row => {
                    row.RelativeItem(1)
                       .AlignCenter()
                       .PaddingRight(4)
                       .Element(c => c.CellStyleCham(paddingTop: ptNhiet, paddingBottom: pbNhiet))
                       .Column(col =>
                       {
                           col.Spacing(0);
                           col.Item().Unconstrained().Text($"{nhietDo}")
                               .FontColor("#0000FF")
                               .FontSize(6)
                               .LineHeight(2.4f);
                           col.Item().Unconstrained().Text("●")
                               .FontColor("#0000FF")
                               .FontSize(11)
                               .LineHeight(0.8f);
                       });
                });
            }

            static void RenderMachNhietDoCell(TableDescriptor table, double mach, double nhietDo, int ptMach, int pbMach, int ptNhiet, int pbNhiet)
            {
                table.Cell().Border(1)
                .Row(row =>
                {
                    row.RelativeItem(1).Element(c =>
                    {
                        c.AlignCenter().Row(inner =>
                        {
                            inner.AutoItem().Layers(layers =>
                            {
                                layers.PrimaryLayer()
                                    .TranslateX(-12)
                                    .Unconstrained()
                                    .Element(x => x.CellStyleCham(paddingTop: ptMach, paddingBottom: pbMach))
                                    .Text(text =>
                                    {
                                        text.Span($"{mach}").FontColor("#FF0000").FontSize(6).LineHeight(0.5f);
                                        text.Span("●").FontColor("#FF0000").FontSize(11).LineHeight(1f);
                                    });

                                layers.Layer()
                                    .TranslateX(-4)
                                    .Unconstrained()
                                    .Element(x => x.CellStyleCham(paddingTop: ptNhiet, paddingBottom: pbNhiet))
                                    .Text(text =>
                                    {
                                        text.Span("●").FontColor("#0000FF").FontSize(11).LineHeight(0.8f);
                                        text.Span($"{nhietDo}").FontColor("#0000FF").FontSize(6).LineHeight(0.5f);
                                    });
                            });
                        });
                    });
                });
            }
        }
    }

    public static class PdfExtensions
    {
        public static IContainer CellStyleCham(
            this IContainer container,
            int paddingTop,
            int paddingBottom)
        {
            float translateY;

            if (5 - paddingBottom < 0)
            {
                translateY = 5 - paddingBottom;
            }
            else if (5 - paddingBottom == 0)
            {
                translateY = paddingTop + paddingBottom;
            }
            else
            {
                translateY = paddingTop + paddingBottom + 5;
            }

                return container
                    .TranslateY(translateY)
                    .PaddingTop(paddingTop)
                    .PaddingBottom(paddingBottom)
                    .PaddingLeft(0)
                    .PaddingRight(0)
                    .DefaultTextStyle(x => x.FontSize(13));
        }
    }
}


