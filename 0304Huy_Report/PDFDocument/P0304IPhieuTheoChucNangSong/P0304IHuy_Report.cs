using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using M0304.Models.ThongTinDoanhNghiep;
using M0304I.Models.PhieuTheoDoiChucNangSong;
using QuestPDF.Drawing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using SixLabors.Fonts;
using SkiaSharp;
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

                page.Content().PaddingVertical(6).Element(e =>
                {
                    e.Column(column =>
                    {

                        int maxColumnsPerPage = 20;
                        int totalColumns = _data.SinhHieus.Count();
                        int totalPages = (int)Math.Ceiling((double)totalColumns / maxColumnsPerPage);

                        for (int page = 0; page < totalPages; page++)
                        {
                            int startIndex = page * maxColumnsPerPage;
                            int endIndex = Math.Min(startIndex + maxColumnsPerPage, totalColumns);
                            column.Item().Element(x =>
                            {
                                x.Table(table =>
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
                                });
                            });

                            column.Item().Element(x =>
                            {
                                x.Table(table =>
                                {
                                    table.ColumnsDefinition(columns =>
                                    {
                                        columns.ConstantColumn(26);
                                        columns.ConstantColumn(28);
                                        columns.RelativeColumn();
                                    });

                                    var leftMachLabels = new[] { "160", "140", "120", "100", "80", "60", "40"};
                                    var leftNhietLabels = new[] { "41", "40", "39", "38", "37", "36", "35"};
                                    int rowsCount = leftMachLabels.Length;
                                    float rowHeightPt = 30f;

                                    table.Cell().Element(CellStyleTop).Element(c =>
                                    {
                                        c.Column(col =>
                                        {
                                            for (int r = 0; r < rowsCount; r++)
                                                col.Item().Height(rowHeightPt).AlignCenter().Text(leftMachLabels[r]);
                                        });
                                    });

                                    table.Cell().Element(CellStyleTop).Element(c =>
                                    {
                                        c.Column(col =>
                                        {
                                            for (int r = 0; r < rowsCount; r++)
                                                col.Item().Height(rowHeightPt).AlignCenter().Text(leftNhietLabels[r]);
                                        });
                                    });

                                    int columnsCount = endIndex - startIndex;
                                    byte[] pngBytes = MatrixRenderer.RenderMatrixPng(
                                        columns: columnsCount,
                                        rows: rowsCount,
                                        rowHeightPt: rowHeightPt,
                                        listMach: listMach ?? new List<string>(),
                                        listNhiet: listNhietDo ?? new List<string>(),
                                        startIndex: startIndex,
                                        endIndex: endIndex,
                                        dpi: 450
                                    );

                                    table.Cell().Element(c =>
                                    {
                                        using var msImg = new MemoryStream(pngBytes);
                                        c.Background("#FFFFFF")
                                         .Border(1)
                                         .BorderColor("#000000")
                                         .Image(msImg)
                                         .FitWidth();
                                    });
                                });
                            });

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
                                        column.Item().Text("2. Cân nặng");
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
                                        column.Item().Text("3. Nhịp thở");
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
                                column.Item().PageBreak();
                            }
                        };
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
            });
        }
        public static class MatrixRenderer
        {
            public static byte[] RenderMatrixPng(
                int columns,
                int rows,
                float rowHeightPt,
                List<string> listMach,
                List<string> listNhiet,
                int startIndex,
                int endIndex,
                int dpi = 450)
            {
                if (columns <= 0 || rows <= 0) return Array.Empty<byte>();

                float totalHeightPt = rows * rowHeightPt;
                float pxPerPt = dpi / 72f;
                int heightPx = (int)Math.Ceiling(totalHeightPt * pxPerPt/2f);

                int basePxPerColumn = (int)Math.Ceiling(12f * pxPerPt);
                int widthPx = Math.Max((int)(columns * basePxPerColumn), 300);

                using var surface = SKSurface.Create(new SKImageInfo(widthPx, heightPx, SKColorType.Rgba8888, SKAlphaType.Premul));
                var canvas = surface.Canvas;
                canvas.Clear(SKColors.Transparent);

                var thin = Math.Max(1f, pxPerPt * 0.25f);
                var gridPaint = new SKPaint { Color = new SKColor(0xE0, 0xE0, 0xE0), StrokeWidth = thin, IsAntialias = false, Style = SKPaintStyle.Stroke };
                var axisPaint = new SKPaint { Color = new SKColor(0xB0, 0xB0, 0xB0), StrokeWidth = Math.Max(1f, pxPerPt * 0.6f), IsAntialias = false, Style = SKPaintStyle.Stroke };
                var redLine = new SKPaint { Color = new SKColor(0xFF, 0x00, 0x00), StrokeWidth = Math.Max(1f, pxPerPt * 0.5f), IsAntialias = true, Style = SKPaintStyle.Stroke };
                var blueLine = new SKPaint { Color = new SKColor(0x00, 0x00, 0xFF), StrokeWidth = Math.Max(1f, pxPerPt * 0.5f), IsAntialias = true, Style = SKPaintStyle.Stroke };
                var redDot = new SKPaint { Color = new SKColor(0xFF, 0x00, 0x00), IsAntialias = true, Style = SKPaintStyle.Fill };
                var blueDot = new SKPaint { Color = new SKColor(0x00, 0x00, 0xFF), IsAntialias = true, Style = SKPaintStyle.Fill };
                var textRed = new SKPaint { Color = SKColors.Red, TextSize = 5 * pxPerPt, IsAntialias = true, Typeface = SKTypeface.FromFamilyName("Arial", SKFontStyle.Bold) };
                var textBlue = new SKPaint { Color = SKColors.Blue, TextSize = 5 * pxPerPt, IsAntialias = true, Typeface = SKTypeface.FromFamilyName("Arial", SKFontStyle.Bold) };

                for (int c = 0; c <= columns; c++)
                {
                    float x = c * (widthPx / (float)columns);
                    canvas.DrawLine(x, 0, x, heightPx, gridPaint);
                }

                for (int r = 0; r <= rows; r++)
                {
                    float y = r * (heightPx / (float)rows);
                    canvas.DrawLine(0, y, widthPx, y, gridPaint);
                }

                float machMin = 40f, machMax = 160f;
                float nhietMin = 35f, nhietMax = 41f;

                float cellHpx = heightPx / (float)rows;

                float MapY(float v, float vMin, float vMax)
                {
                    float stepValue = (vMax - vMin) / Math.Max(1f, (rows - 1));
                    if (stepValue <= 0f) stepValue = 1f;

                    float rowIndex = (vMax - v) / stepValue;
                    rowIndex = Math.Clamp(rowIndex, 0f, rows - 1f);

                    float y = (rowIndex + 0.0f) * cellHpx;

                    return Math.Clamp(y, 0f, heightPx - 1f);
                }

                float ColumnCenterX(int colIndex)
                {
                    return (colIndex + 0.5f) * (widthPx / (float)columns);
                }

                var machPts = new List<(int col, float value)>();
                var nhietPts = new List<(int col, float value)>();

                for (int i = startIndex; i < endIndex && i < listMach.Count && i < listNhiet.Count; i++)
                {
                    int colIdx = i - startIndex;
                    if (double.TryParse(listMach[i], NumberStyles.Any, CultureInfo.InvariantCulture, out var mVal))
                    {
                        if (!double.IsNaN(mVal) && mVal > 0) machPts.Add((colIdx, (float)mVal));
                    }
                    if (double.TryParse(listNhiet[i], NumberStyles.Any, CultureInfo.InvariantCulture, out var nVal))
                    {
                        if (!double.IsNaN(nVal) && nVal > 0) nhietPts.Add((colIdx, (float)nVal));
                    }
                }

                void DrawSeries(
                    List<(int col, float value)> pts,
                    SKPaint paintLine,
                    SKPaint paintDot,
                    float vMin,
                    float vMax,
                    SKPaint paintText,
                    bool textAbove
                )
                {
                    if (pts.Count == 0) return;
                    var sorted = pts.OrderBy(p => p.col).ToList();

                    paintText.IsAntialias = true;
                    paintText.SubpixelText = true;
                    paintText.LcdRenderText = true;
                    paintText.HintingLevel = SKPaintHinting.Full;
                    paintText.IsStroke = false;
                    paintText.FilterQuality = SKFilterQuality.High;

                    for (int i = 0; i < sorted.Count - 1; i++)
                    {
                        var a = sorted[i];
                        var b = sorted[i + 1];
                        var ax = ColumnCenterX(a.col);
                        var ay = MapY(a.value, vMin, vMax);
                        var bx = ColumnCenterX(b.col);
                        var by = MapY(b.value, vMin, vMax);
                        canvas.DrawLine(ax, ay, bx, by, paintLine);
                    }

                    float dotR = Math.Max(1.6f, pxPerPt * 0.8f);
                    foreach (var p in sorted)
                    {
                        var x = ColumnCenterX(p.col);
                        var y = MapY(p.value, vMin, vMax);

                        canvas.DrawCircle(x, y, dotR, paintDot);

                        var text = p.value.ToString("0.#", CultureInfo.InvariantCulture);
                        float textW = paintText.MeasureText(text);
                        float textX = x - textW / 2f;

                        float textBaselineY;
                        if (textAbove)
                            textBaselineY = y - dotR - (paintText.TextSize * 0.4f) - 2f;
                        else
                            textBaselineY = y + dotR + (paintText.TextSize) + 2f;

                        canvas.DrawText(text, textX, textBaselineY, paintText);
                    }
                }

                DrawSeries(machPts, redLine, redDot, machMin, machMax, textRed, true);
                DrawSeries(nhietPts, blueLine, blueDot, nhietMin, nhietMax, textBlue, false);

                using var img = surface.Snapshot();
                using var data = img.Encode(SKEncodedImageFormat.Png, 100);

                return data.ToArray();
            }
        }
    }
}

