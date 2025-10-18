using ClosedXML.Excel;
using H0304.NumberToText.Helpers;
using M0304NhanVien.Models;
using M0304.Models.BangKeThu;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using M0304.Models.ThongTinDoanhNghiep;

public class P0304ExcelReportTemplate
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

    public P0304ExcelReportTemplate(
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

    public byte[] GenerateExcel()
    {
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Báo cáo");

            int currentRow = 1;

            if (!string.IsNullOrEmpty(_logoPath) && File.Exists(_logoPath))
            {
                var img = ws.AddPicture(_logoPath)
                    .MoveTo(ws.Cell(1, 2))
                    .Scale(0.2);
                ws.Row(1).AdjustToContents();
            }

            ws.Range(1, 3, 1, 11).Merge();
            ws.Cell(1, 3).Value = _dataDN.TenCoQuanChuyenMon ?? "";
            ws.Cell(1, 3).Style.Font.FontSize = 9;
            ws.Cell(1, 3).Style.Font.Bold = true;

            ws.Range(2, 3, 2, 11).Merge();
            ws.Cell(2, 3).Value = _dataDN.TenCSKCB ?? "";
            ws.Cell(2, 3).Style.Font.FontSize = 9;
            ws.Cell(2, 3).Style.Font.Bold = true;

            currentRow += 4;

            ws.Range(currentRow, 2, currentRow, 11).Merge();
            ws.Cell(currentRow, 2).Value = "BẢNG KÊ THU TIỀN NGOẠI TRÚ THEO BL/HĐ";
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Font.FontSize = 14;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 11).Merge();
            DateTime dtStart, dtEnd;
            if (DateTime.TryParse(_ngayBatDau, out dtStart) && DateTime.TryParse(_ngayKetThuc, out dtEnd))
            {
                ws.Cell(currentRow, 2).Value = $"Từ ngày {dtStart:dd-MM-yyyy} đến ngày {dtEnd:dd-MM-yyyy}";
            }
            else
            {
                ws.Cell(currentRow, 2).Value = $"Từ ngày {_ngayBatDau} đến ngày {_ngayKetThuc}";
            }
            ws.Cell(currentRow, 2).Style.Font.FontSize = 10;
            ws.Cell(currentRow, 2).Style.Font.Italic = true;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 11).Merge();
            ws.Cell(currentRow, 2).Value = _tenHTTT;
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Font.FontSize = 10;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow += 2;

            string[] headers = new string[]
            {
            "STT", "Mã y tế", "Họ và tên", "Quyển sổ", "Số biên lai",
            "Loại", "Ngày thu", "Hủy", "Hoàn", "Số tiền"
            };
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(currentRow, i + 2).Value = headers[i];
                ws.Cell(currentRow, i + 2).Style.Font.Bold = true;
                ws.Cell(currentRow, i + 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(currentRow, i + 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            int stt = 1;

            currentRow++;
            if (_danhSachNhanVien != null && _danhSachNhanVien.Any())
            {
                foreach (var nv in _danhSachNhanVien)
                {
                    // Dòng tiêu đề nhân viên
                    ws.Range(currentRow, 2, currentRow, 11).Merge();
                    ws.Cell(currentRow, 2).Value = $"{nv.TenNhanVien}".ToUpper();
                    ws.Cell(currentRow, 2).Style.Font.Bold = true;
                    ws.Cell(currentRow, 2).Style.Font.FontSize = 11;
                    ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    ws.Range(currentRow, 2, currentRow, 11).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    currentRow++;

                    // Danh sách quyển sổ mà nhân viên có tham gia
                    var quyenSoList = _tongTheoQuyenSo
                        .Where(q => _data.Any(d => d.IDNhanVien == nv.ID && d.QuyenSo == q.QuyenSo))
                        .ToList();

                    foreach (var qs in quyenSoList)
                    {
                        // Lấy danh sách chi tiết của nhân viên theo quyển sổ
                        var chiTietNvQs = _data
                            .Where(d => d.IDNhanVien == nv.ID && d.QuyenSo == qs.QuyenSo)
                            .ToList();

                        // Seri lấy theo dữ liệu nhân viên (nếu null thì fallback về qs.Seri)
                        var seriForNv = chiTietNvQs.Select(d => d.Seri).FirstOrDefault() ?? qs.Seri ?? "";

                        // Tính tổng riêng cho nhân viên trong quyển sổ
                        var tongTheoNVvaQS = chiTietNvQs
                            .GroupBy(x => x.QuyenSo)
                            .Select(g => new
                            {
                                TongHuy = g.Sum(x => x.Huy ?? 0m),
                                TongHoan = g.Sum(x => x.Hoan ?? 0m),
                                TongSoTien = g.Sum(x => x.SoTien ?? 0m)
                            })
                            .FirstOrDefault() ?? new { TongHuy = 0m, TongHoan = 0m, TongSoTien = 0m };

                        // Ngày thu gần nhất của nhân viên trong quyển sổ này
                        var ngayThu = chiTietNvQs.Max(d => d.NgayThu);

                        // Dòng tiêu đề quyển sổ cho nhân viên
                        ws.Range(currentRow, 2, currentRow, 7).Merge();
                        ws.Cell(currentRow, 2).Value = $"      {nv.MaNhanVien} - {seriForNv}.{qs.QuyenSo}";
                        ws.Cell(currentRow, 2).Style.Font.Bold = true;
                        ws.Cell(currentRow, 2).Style.Font.FontSize = 10;
                        ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                        ws.Cell(currentRow, 8).Value = $"{ngayThu:dd-MM-yyyy}";
                        ws.Cell(currentRow, 8).Style.Font.Bold = true;
                        ws.Cell(currentRow, 8).Style.Font.FontSize = 10;
                        AlignCellCenter(ws.Cell(currentRow, 8));

                        ws.Cell(currentRow, 9).Value = tongTheoNVvaQS.TongHuy;
                        ws.Cell(currentRow, 9).Style.Font.Bold = true;
                        ws.Cell(currentRow, 9).Style.Font.FontSize = 10;
                        AlignCellRight(ws.Cell(currentRow, 9));

                        ws.Cell(currentRow, 10).Value = tongTheoNVvaQS.TongHoan;
                        ws.Cell(currentRow, 10).Style.Font.Bold = true;
                        ws.Cell(currentRow, 10).Style.Font.FontSize = 10;
                        AlignCellRight(ws.Cell(currentRow, 10));

                        ws.Cell(currentRow, 11).Value = tongTheoNVvaQS.TongSoTien;
                        ws.Cell(currentRow, 11).Style.Font.Bold = true;
                        ws.Cell(currentRow, 11).Style.Font.FontSize = 10;
                        AlignCellRight(ws.Cell(currentRow, 11));

                        ws.Cell(currentRow, 9).Style.NumberFormat.Format = "#,##0";
                        ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";
                        ws.Cell(currentRow, 11).Style.NumberFormat.Format = "#,##0";
                        ws.Range(currentRow, 2, currentRow, 11).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        currentRow++;

                        // Ghi chi tiết từng dòng của nhân viên trong quyển sổ này
                        foreach (var item in chiTietNvQs)
                        {
                            ws.Cell(currentRow, 2).Value = stt++; AlignCellCenter(ws.Cell(currentRow, 2));
                            ws.Cell(currentRow, 3).Value = item.MaYTe ?? ""; AlignCellCenter(ws.Cell(currentRow, 3));
                            ws.Cell(currentRow, 4).Value = item.HoVaTen ?? "";
                            ws.Cell(currentRow, 5).Value = item.QuyenSo ?? ""; AlignCellCenter(ws.Cell(currentRow, 5));
                            ws.Cell(currentRow, 6).Value = item.SoBienLai ?? ""; AlignCellCenter(ws.Cell(currentRow, 6));
                            ws.Cell(currentRow, 7).Value = item.Loai ?? ""; AlignCellCenter(ws.Cell(currentRow, 7));
                            ws.Cell(currentRow, 8).Value = item.NgayThu?.ToString("dd-MM-yyyy") ?? ""; AlignCellCenter(ws.Cell(currentRow, 8));
                            ws.Cell(currentRow, 9).Value = item.Huy ?? (decimal?)null;
                            ws.Cell(currentRow, 10).Value = item.Hoan ?? (decimal?)null;
                            ws.Cell(currentRow, 11).Value = item.SoTien ?? (decimal?)null;

                            ws.Cell(currentRow, 9).Style.NumberFormat.Format = "#,##0";
                            ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";
                            ws.Cell(currentRow, 11).Style.NumberFormat.Format = "#,##0";

                            for (int col = 2; col <= headers.Length + 1; col++)
                                ws.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                            currentRow++;
                        }
                    }
                }
            }


            var tongHuyAll = _data.Sum(x => x.Huy ?? 0m);
            var tongHoanAll = _data.Sum(x => x.Hoan ?? 0m);
            var tongSoTienAll = _data.Sum(x => x.SoTien ?? 0m);

            var phaiNop = tongSoTienAll - tongHuyAll - tongHoanAll;

            var totalRange = ws.Range(currentRow, 2, currentRow, 8);
            totalRange.Merge();
            totalRange.Value = "Tổng cộng";
            totalRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            totalRange.Style.Font.Bold = true;
            totalRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            totalRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            ws.Cell(currentRow, 9).Value = tongHuyAll;
            ws.Cell(currentRow, 10).Value = tongHoanAll;
            ws.Cell(currentRow, 11).Value = tongSoTienAll;

            ws.Cell(currentRow, 9).Style.NumberFormat.Format = "#,##0";
            ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";
            ws.Cell(currentRow, 11).Style.NumberFormat.Format = "#,##0";

            for (int col = 9; col <= 11; col++)
            {
                ws.Cell(currentRow, col).Style.Font.Bold = true;
                ws.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }
            currentRow += 2;

            ws.Range(currentRow, 2, currentRow, 7).Merge();
            var tongSo = ws.Cell(currentRow, 2).GetRichText();
            tongSo.AddText("Số tiền phải nộp: ");
            tongSo.AddText($"{phaiNop:N0}").SetBold();
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 7).Merge();
            var tongChu = ws.Cell(currentRow, 2).GetRichText();
            tongChu.AddText("Bằng chữ: ");
            tongChu.AddText(H0304NumberToTextHelper.ConvertSoThanhChu(phaiNop)).SetItalic().SetBold();
            currentRow += 2;

            ws.Range(currentRow, 8, currentRow + 5, 11).Merge();

            var cell = ws.Cell(currentRow, 8);
            cell.Value = "";

            var rt = cell.GetRichText();

            rt.AddText($"Ngày {DateTime.Now:dd} Tháng {DateTime.Now:MM} Năm {DateTime.Now:yyyy}\n");
            rt.AddText("Người lập bảng\n\n\n").SetBold();
            rt.AddText($"{_tenNVDN}").SetBold();

            ws.Cell(currentRow, 8).Style.Alignment.WrapText = true;
            ws.Cell(currentRow, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(currentRow, 8).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            ws.Columns().AdjustToContents();
            ws.Column(1).Width = 2;
            ws.Column(2).Width = 12;


            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms);
                return ms.ToArray();
            }

            void AlignCellCenter(IXLCell cell)
            {
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }

            void AlignCellRight(IXLCell cell)
            {
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }
        }
    }
}
