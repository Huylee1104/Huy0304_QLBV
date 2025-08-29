using ClosedXML.Excel;
using H0304.NumberToText.Helpers;
using M0304NhanVien.Models;
using M0304.Models.BangKeThu;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using M0304.Models.ThongTinDOanhNghiep;

public class P0304ExcelReportTemplate
{
    private readonly List<M0304BangKeThu> _data;
    private readonly M0304ThongTinDoanhNghiep _dataDN;
    private string _ngayBatDau;
    private string _ngayKetThuc;
    private string _tenNhanVien;
    private string _tenHTTT;
    private readonly string _logoPath;

    private M0304TongHopBangKeThu _tongChung;
    private List<M0304TongTheoNhanVien> _tongTheoNhanVien;
    private List<M0304NhanVienModel> _danhSachNhanVien;

    public P0304ExcelReportTemplate(
        List<M0304BangKeThu> data,
        string ngayBatDau,
        string ngayKetThuc,
        string tenHTTT,
        string tenNhanVien,
        List<M0304NhanVienModel> danhSachNhanVien,
        M0304TongHopBangKeThu tongChung,
        List<M0304TongTheoNhanVien> tongTheoNhanVien,
        M0304ThongTinDoanhNghiep dataDN,
        string logoPath = null
    )
    {
        _data = data ?? new List<M0304BangKeThu>();
        _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
        _ngayBatDau = ngayBatDau;
        _ngayKetThuc = ngayKetThuc;
        _tenHTTT = tenHTTT;
        _tenNhanVien = tenNhanVien;
        _danhSachNhanVien = danhSachNhanVien ?? new List<M0304NhanVienModel>();
        _tongChung = tongChung ?? new M0304TongHopBangKeThu();
        _tongTheoNhanVien = tongTheoNhanVien ?? new List<M0304TongTheoNhanVien>();
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
                ws.Range(1, 1, 2, 2).Merge();
                var img = ws.AddPicture(_logoPath)
                    .MoveTo(ws.Cell(1, 1))
                    .Scale(0.2);
                ws.Row(1).Height = 0;
            }

            ws.Range(1, 3, 1, 10).Merge();
            ws.Cell(1, 3).Value = _dataDN.TenCSKCB ?? "";
            ws.Cell(1, 3).Style.Font.FontSize = 9;

            ws.Range(2, 3, 2, 10).Merge();
            ws.Cell(2, 3).Value = _dataDN.TenCoQuanChuyenMon ?? "";
            ws.Cell(2, 3).Style.Font.FontSize = 9;

            ws.Range(3, 3, 3, 10).Merge();
            ws.Cell(3, 3).Value = _dataDN.DiaChi ?? "";
            ws.Cell(3, 3).Style.Font.FontSize = 9;

            ws.Range(4, 3, 4, 10).Merge();
            ws.Cell(4, 3).Value = _dataDN.DienThoai ?? "";
            ws.Cell(4, 3).Style.Font.FontSize = 9;

            currentRow += 5;

            ws.Range(currentRow, 1, currentRow, 10).Merge();
            ws.Cell(currentRow, 1).Value = "BẢNG KÊ THU TIỀN NGOẠI TRÚ THEO BL/HĐ";
            ws.Cell(currentRow, 1).Style.Font.Bold = true;
            ws.Cell(currentRow, 1).Style.Font.FontSize = 14;
            ws.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            ws.Range(currentRow, 1, currentRow, 10).Merge();
            DateTime dtStart, dtEnd;
            if (DateTime.TryParse(_ngayBatDau, out dtStart) && DateTime.TryParse(_ngayKetThuc, out dtEnd))
            {
                ws.Cell(currentRow, 1).Value = $"Từ ngày {dtStart:dd/MM/yyyy} đến ngày {dtEnd:dd/MM/yyyy}";
            }
            else
            {
                ws.Cell(currentRow, 1).Value = $"Từ ngày {_ngayBatDau} đến ngày {_ngayKetThuc}";
            }
            ws.Cell(currentRow, 1).Style.Font.FontSize = 10;
            ws.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            ws.Range(currentRow, 1, currentRow, 10).Merge();
            ws.Cell(currentRow, 1).Value = _tenHTTT;
            ws.Cell(currentRow, 1).Style.Font.Bold = true;
            ws.Cell(currentRow, 1).Style.Font.FontSize = 10;
            ws.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow += 2;

            string[] headers = new string[]
            {
            "STT", "Mã y tế", "Họ và tên", "Quyển sổ", "Số biên lai",
            "Loại", "Ngày thu", "Hủy", "Hoàn", "Số tiền"
            };
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(currentRow, i + 1).Value = headers[i];
                ws.Cell(currentRow, i + 1).Style.Font.Bold = true;
                ws.Cell(currentRow, i + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                ws.Cell(currentRow, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(currentRow, i + 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            int stt = 1;

            currentRow++;
            if (_danhSachNhanVien != null && _danhSachNhanVien.Any())
            {
                foreach (var nv in _danhSachNhanVien)
                {
                    var dataNV = _data.Where(d => d.IDNhanVien == nv.Id).ToList();
                    if (!dataNV.Any()) continue;
                    var tongNV = _tongTheoNhanVien.FirstOrDefault(t => t.IDNhanVien == nv.Id);
                    if (tongNV.TongSoTien > 0)
                    {
                        ws.Range(currentRow, 1, currentRow, 7).Merge();
                        ws.Cell(currentRow, 1).Value = $"Tên nhân viên: {nv.Ten}";
                        ws.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        ws.Cell(currentRow, 1).Style.Font.Bold = true;
                        ws.Cell(currentRow, 1).Style.Font.FontSize = 10;

                        ws.Cell(currentRow, 8).Value = 0;
                        ws.Cell(currentRow, 9).Value = 0;
                        ws.Cell(currentRow, 10).Value = tongNV?.TongSoTien;
                        ws.Cell(currentRow, 8).Style.NumberFormat.Format = "#,##0";
                        ws.Cell(currentRow, 9).Style.NumberFormat.Format = "#,##0";
                        ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";

                        ws.Range(currentRow, 1, currentRow, 10).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Range(currentRow, 1, currentRow, 10).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                        currentRow++;
                    }
                    foreach (var item in dataNV)
                    {
                        if (item.SoTien > 0)
                        {
                            ws.Cell(currentRow, 1).Value = stt++; AlignCellCenter(ws.Cell(currentRow, 1));
                            ws.Cell(currentRow, 2).Value = item.MaYTe ?? ""; AlignCellCenter(ws.Cell(currentRow, 2));
                            ws.Cell(currentRow, 3).Value = item.HoVaTen ?? "";
                            ws.Cell(currentRow, 4).Value = item.QuyenSo ?? ""; AlignCellCenter(ws.Cell(currentRow, 4));
                            ws.Cell(currentRow, 5).Value = item.SoBienLai ?? ""; AlignCellCenter(ws.Cell(currentRow, 5));
                            ws.Cell(currentRow, 6).Value = item.Loai ?? ""; AlignCellCenter(ws.Cell(currentRow, 6));
                            ws.Cell(currentRow, 7).Value = item.NgayThu?.ToString("dd/MM/yyyy") ?? ""; AlignCellCenter(ws.Cell(currentRow, 7));
                            ws.Cell(currentRow, 8).Value = item.Huy ?? (decimal?)null;
                            ws.Cell(currentRow, 9).Value = item.Hoan ?? (decimal?)null;
                            ws.Cell(currentRow, 10).Value = item.SoTien ?? (decimal?)null;

                            ws.Cell(currentRow, 8).Style.NumberFormat.Format = "#,##0";
                            ws.Cell(currentRow, 9).Style.NumberFormat.Format = "#,##0";
                            ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";

                            for (int col = 1; col <= headers.Length; col++)
                                ws.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                            currentRow++;
                        }
                    }
                    if (tongNV.TongHuy > 0)
                    {
                        ws.Range(currentRow, 1, currentRow, 7).Merge();
                        ws.Cell(currentRow, 1).Value = $"Tên nhân viên: {nv.Ten}";
                        ws.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        ws.Cell(currentRow, 1).Style.Font.Bold = true;
                        ws.Cell(currentRow, 1).Style.Font.FontSize = 10;

                        ws.Cell(currentRow, 8).Value = tongNV?.TongHuy;
                        ws.Cell(currentRow, 9).Value = 0;
                        ws.Cell(currentRow, 10).Value = 0;
                        ws.Cell(currentRow, 8).Style.NumberFormat.Format = "#,##0";
                        ws.Cell(currentRow, 9).Style.NumberFormat.Format = "#,##0";
                        ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";

                        ws.Range(currentRow, 1, currentRow, 10).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Range(currentRow, 1, currentRow, 10).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                        currentRow++;
                    }
                    foreach (var item in dataNV)
                    {
                        if (item.Huy > 0)
                        {
                            ws.Cell(currentRow, 1).Value = stt++; AlignCellCenter(ws.Cell(currentRow, 1));
                            ws.Cell(currentRow, 2).Value = item.MaYTe ?? ""; AlignCellCenter(ws.Cell(currentRow, 2));
                            ws.Cell(currentRow, 3).Value = item.HoVaTen ?? ""; 
                            ws.Cell(currentRow, 4).Value = item.QuyenSo ?? ""; AlignCellCenter(ws.Cell(currentRow, 4));
                            ws.Cell(currentRow, 5).Value = item.SoBienLai ?? ""; AlignCellCenter(ws.Cell(currentRow, 5));
                            ws.Cell(currentRow, 6).Value = item.Loai ?? ""; AlignCellCenter(ws.Cell(currentRow, 6));
                            ws.Cell(currentRow, 7).Value = item.NgayThu?.ToString("dd/MM/yyyy") ?? ""; AlignCellCenter(ws.Cell(currentRow, 7));
                            ws.Cell(currentRow, 8).Value = item.Huy ?? (decimal?)null;
                            ws.Cell(currentRow, 9).Value = item.Hoan ?? (decimal?)null;
                            ws.Cell(currentRow, 10).Value = item.SoTien ?? (decimal?)null;

                            ws.Cell(currentRow, 8).Style.NumberFormat.Format = "#,##0";
                            ws.Cell(currentRow, 9).Style.NumberFormat.Format = "#,##0";
                            ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";

                            for (int col = 1; col <= headers.Length; col++)
                                ws.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                            currentRow++;
                        }
                    }
                    if (tongNV.TongHoan > 0)
                    {
                        ws.Range(currentRow, 1, currentRow, 7).Merge();
                        ws.Cell(currentRow, 1).Value = $"Tên nhân viên: {nv.Ten}";
                        ws.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        ws.Cell(currentRow, 1).Style.Font.Bold = true;
                        ws.Cell(currentRow, 1).Style.Font.FontSize = 10;

                        ws.Cell(currentRow, 8).Value = 0;
                        ws.Cell(currentRow, 9).Value = tongNV?.TongHoan;
                        ws.Cell(currentRow, 10).Value = 0;
                        ws.Cell(currentRow, 8).Style.NumberFormat.Format = "#,##0";
                        ws.Cell(currentRow, 9).Style.NumberFormat.Format = "#,##0";
                        ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";

                        ws.Range(currentRow, 1, currentRow, 10).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Range(currentRow, 1, currentRow, 10).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                        currentRow++;
                    }
                    foreach (var item in dataNV)
                    {
                        if (item.Hoan > 0)
                        {
                            ws.Cell(currentRow, 1).Value = stt++; AlignCellCenter(ws.Cell(currentRow, 1));
                            ws.Cell(currentRow, 2).Value = item.MaYTe ?? ""; AlignCellCenter(ws.Cell(currentRow, 2));
                            ws.Cell(currentRow, 3).Value = item.HoVaTen ?? "";
                            ws.Cell(currentRow, 4).Value = item.QuyenSo ?? ""; AlignCellCenter(ws.Cell(currentRow, 4));
                            ws.Cell(currentRow, 5).Value = item.SoBienLai ?? ""; AlignCellCenter(ws.Cell(currentRow, 5));
                            ws.Cell(currentRow, 6).Value = item.Loai ?? ""; AlignCellCenter(ws.Cell(currentRow, 6));
                            ws.Cell(currentRow, 7).Value = item.NgayThu?.ToString("dd/MM/yyyy") ?? ""; AlignCellCenter(ws.Cell(currentRow, 7));
                            ws.Cell(currentRow, 8).Value = item.Huy ?? (decimal?)null;
                            ws.Cell(currentRow, 9).Value = item.Hoan ?? (decimal?)null;
                            ws.Cell(currentRow, 10).Value = item.SoTien ?? (decimal?)null;

                            ws.Cell(currentRow, 8).Style.NumberFormat.Format = "#,##0";
                            ws.Cell(currentRow, 9).Style.NumberFormat.Format = "#,##0";
                            ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";

                            for (int col = 1; col <= headers.Length; col++)
                                ws.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                            currentRow++;
                        }
                    }
                }
            }
            else
            {
                if (_tongChung.TongSoTien > 0)
                {
                    ws.Range(currentRow, 1, currentRow, 7).Merge();
                    ws.Cell(currentRow, 1).Value = $"Tên nhân viên: {_tenNhanVien}";
                    ws.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    ws.Cell(currentRow, 1).Style.Font.Bold = true;
                    ws.Cell(currentRow, 1).Style.Font.FontSize = 10;

                    ws.Cell(currentRow, 8).Value = 0;
                    ws.Cell(currentRow, 9).Value = 0;
                    ws.Cell(currentRow, 10).Value = _tongChung.TongSoTien;
                    ws.Cell(currentRow, 8).Style.NumberFormat.Format = "#,##0";
                    ws.Cell(currentRow, 9).Style.NumberFormat.Format = "#,##0";
                    ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";

                    ws.Range(currentRow, 1, currentRow, 10).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Range(currentRow, 1, currentRow, 10).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    currentRow++;
                }
                foreach (var item in _data)
                {
                    if (item.SoTien > 0)
                    {
                        ws.Cell(currentRow, 1).Value = stt++; AlignCellCenter(ws.Cell(currentRow, 1));
                        ws.Cell(currentRow, 2).Value = item.MaYTe ?? ""; AlignCellCenter(ws.Cell(currentRow, 2));
                        ws.Cell(currentRow, 3).Value = item.HoVaTen ?? "";
                        ws.Cell(currentRow, 4).Value = item.QuyenSo ?? ""; AlignCellCenter(ws.Cell(currentRow, 4));
                        ws.Cell(currentRow, 5).Value = item.SoBienLai ?? ""; AlignCellCenter(ws.Cell(currentRow, 5));
                        ws.Cell(currentRow, 6).Value = item.Loai ?? ""; AlignCellCenter(ws.Cell(currentRow, 6));
                        ws.Cell(currentRow, 7).Value = item.NgayThu?.ToString("dd/MM/yyyy") ?? "";  AlignCellCenter(ws.Cell(currentRow, 7));
                        ws.Cell(currentRow, 8).Value = item.Huy ?? (decimal?)null;
                        ws.Cell(currentRow, 9).Value = item.Hoan ?? (decimal?)null;
                        ws.Cell(currentRow, 10).Value = item.SoTien ?? (decimal?)null;

                        ws.Cell(currentRow, 8).Style.NumberFormat.Format = "#,##0";
                        ws.Cell(currentRow, 9).Style.NumberFormat.Format = "#,##0";
                        ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";

                        for (int col = 1; col <= headers.Length; col++)
                            ws.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        currentRow++;
                    }
                }
                if (_tongChung.TongHuy > 0)
                {
                    ws.Range(currentRow, 1, currentRow, 7).Merge();
                    ws.Cell(currentRow, 1).Value = $"Tên nhân viên: {_tenNhanVien}";
                    ws.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    ws.Cell(currentRow, 1).Style.Font.Bold = true;
                    ws.Cell(currentRow, 1).Style.Font.FontSize = 10;

                    ws.Cell(currentRow, 8).Value = _tongChung.TongHuy;
                    ws.Cell(currentRow, 9).Value = 0;
                    ws.Cell(currentRow, 10).Value = 0;
                    ws.Cell(currentRow, 8).Style.NumberFormat.Format = "#,##0";
                    ws.Cell(currentRow, 9).Style.NumberFormat.Format = "#,##0";
                    ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";

                    ws.Range(currentRow, 1, currentRow, 10).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Range(currentRow, 1, currentRow, 10).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    currentRow++;
                }
                foreach (var item in _data)
                {
                    if (item.Huy > 0)
                    {
                        ws.Cell(currentRow, 1).Value = stt++; AlignCellCenter(ws.Cell(currentRow, 1));
                        ws.Cell(currentRow, 2).Value = item.MaYTe ?? ""; AlignCellCenter(ws.Cell(currentRow, 2));
                        ws.Cell(currentRow, 3).Value = item.HoVaTen ?? "";
                        ws.Cell(currentRow, 4).Value = item.QuyenSo ?? ""; AlignCellCenter(ws.Cell(currentRow, 4));
                        ws.Cell(currentRow, 5).Value = item.SoBienLai ?? ""; AlignCellCenter(ws.Cell(currentRow, 5));
                        ws.Cell(currentRow, 6).Value = item.Loai ?? ""; AlignCellCenter(ws.Cell(currentRow, 6));
                        ws.Cell(currentRow, 7).Value = item.NgayThu?.ToString("dd/MM/yyyy") ?? ""; AlignCellCenter(ws.Cell(currentRow, 7));
                        ws.Cell(currentRow, 8).Value = item.Huy ?? (decimal?)null;
                        ws.Cell(currentRow, 9).Value = item.Hoan ?? (decimal?)null;
                        ws.Cell(currentRow, 10).Value = item.SoTien ?? (decimal?)null;

                        ws.Cell(currentRow, 8).Style.NumberFormat.Format = "#,##0";
                        ws.Cell(currentRow, 9).Style.NumberFormat.Format = "#,##0";
                        ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";

                        for (int col = 1; col <= headers.Length; col++)
                            ws.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        currentRow++;
                    }
                }
                if (_tongChung.TongHoan > 0)
                {
                    ws.Range(currentRow, 1, currentRow, 7).Merge();
                    ws.Cell(currentRow, 1).Value = $"Tên nhân viên: {_tenNhanVien}";
                    ws.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    ws.Cell(currentRow, 1).Style.Font.Bold = true;
                    ws.Cell(currentRow, 1).Style.Font.FontSize = 10;

                    ws.Cell(currentRow, 8).Value = 0;
                    ws.Cell(currentRow, 9).Value = _tongChung.TongHoan;
                    ws.Cell(currentRow, 10).Value = 0;
                    ws.Cell(currentRow, 8).Style.NumberFormat.Format = "#,##0";
                    ws.Cell(currentRow, 9).Style.NumberFormat.Format = "#,##0";
                    ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";

                    ws.Range(currentRow, 1, currentRow, 10).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Range(currentRow, 1, currentRow, 10).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    currentRow++;
                }
                foreach (var item in _data)
                {
                    if (item.Hoan > 0)
                    {
                        ws.Cell(currentRow, 1).Value = stt++; AlignCellCenter(ws.Cell(currentRow, 1));
                        ws.Cell(currentRow, 2).Value = item.MaYTe ?? ""; AlignCellCenter(ws.Cell(currentRow, 2));
                        ws.Cell(currentRow, 3).Value = item.HoVaTen ?? "";
                        ws.Cell(currentRow, 4).Value = item.QuyenSo ?? "";  AlignCellCenter(ws.Cell(currentRow, 4));
                        ws.Cell(currentRow, 5).Value = item.SoBienLai ?? ""; AlignCellCenter(ws.Cell(currentRow, 5));
                        ws.Cell(currentRow, 6).Value = item.Loai ?? ""; AlignCellCenter(ws.Cell(currentRow, 6));
                        ws.Cell(currentRow, 7).Value = item.NgayThu?.ToString("dd/MM/yyyy") ?? ""; AlignCellCenter(ws.Cell(currentRow, 7));
                        ws.Cell(currentRow, 8).Value = item.Huy ?? (decimal?)null;
                        ws.Cell(currentRow, 9).Value = item.Hoan ?? (decimal?)null;
                        ws.Cell(currentRow, 10).Value = item.SoTien ?? (decimal?)null;

                        ws.Cell(currentRow, 8).Style.NumberFormat.Format = "#,##0";
                        ws.Cell(currentRow, 9).Style.NumberFormat.Format = "#,##0";
                        ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";

                        for (int col = 1; col <= headers.Length; col++)
                            ws.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        currentRow++;
                    }
                }
            }
            var totalRange = ws.Range(currentRow, 1, currentRow, 7);
            totalRange.Merge();
            totalRange.Value = "Tổng cộng";
            totalRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            totalRange.Style.Font.Bold = true;
            totalRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            totalRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            ws.Cell(currentRow, 8).Value = _tongChung.TongHuy;
            ws.Cell(currentRow, 9).Value = _tongChung.TongHoan;
            ws.Cell(currentRow, 10).Value = _tongChung.TongSoTien;

            ws.Cell(currentRow, 8).Style.NumberFormat.Format = "#,##0";
            ws.Cell(currentRow, 9).Style.NumberFormat.Format = "#,##0";
            ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";

            for (int col = 8; col <= 10; col++)
            {
                ws.Cell(currentRow, col).Style.Font.Bold = true;
                ws.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }
            currentRow += 2;

            ws.Range(currentRow, 1, currentRow, 6).Merge(); 
            ws.Cell(currentRow, 1).Value = $"Số tiền phải nộp: {_tongChung.TongChenhLech:N0}";
            ws.Cell(currentRow, 1).Style.Font.Bold = true;
            currentRow++;

            ws.Range(currentRow, 1, currentRow, 6).Merge(); 
            ws.Cell(currentRow, 1).Value = $"Bằng chữ: {H0304NumberToTextHelper.ConvertSoThanhChu(_tongChung.TongChenhLech)}";
            ws.Cell(currentRow, 1).Style.Font.Italic = true;
            currentRow += 2;

            ws.Range(currentRow, 7, currentRow + 5, 10).Merge();
            ws.Cell(currentRow, 7).Value =
                $"Ngày {DateTime.Now:dd} Tháng {DateTime.Now:MM} Năm {DateTime.Now:yyyy}\n" +
                "Người lập bảng\n\n\n" +
                "Trần Thanh Thảo";
            ws.Cell(currentRow, 7).Style.Font.Bold = true;
            ws.Cell(currentRow, 7).Style.Alignment.WrapText = true;
            ws.Cell(currentRow, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(currentRow, 7).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            ws.Columns().AdjustToContents();


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
        }
    }
}
