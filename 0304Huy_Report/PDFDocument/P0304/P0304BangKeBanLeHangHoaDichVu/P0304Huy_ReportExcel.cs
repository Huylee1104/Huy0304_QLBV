using ClosedXML.Excel;
using H0304.NumberToText.Helpers;
using M0304.Models.BangKeBanLeHangHoaDichVu;
using M0304.Models.ThongTinDoanhNghiep;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

public class P0304BangKeBanLeHangHoaDichVuExcelReportTemplate
{
    private readonly List<M0304BangKeBanLeHangHoaDichVu> _data;
    private string _ngayBatDau;
    private string _ngayKetThuc;
    private string _tenNhanVien;
    private readonly M0304ThongTinDoanhNghiep _dataDN;
    private readonly string _logoPath;

    public P0304BangKeBanLeHangHoaDichVuExcelReportTemplate(
        List<M0304BangKeBanLeHangHoaDichVu> data,
        string ngayBatDau,
        string ngayKetThuc,
        string tenNhanVien,
        M0304ThongTinDoanhNghiep dataDN,
        string logoPath = null
    )
    {
        _data = data ?? new List<M0304BangKeBanLeHangHoaDichVu>();
        _ngayBatDau = ngayBatDau;
        _ngayKetThuc = ngayKetThuc;
        _tenNhanVien = tenNhanVien;
        _dataDN = dataDN;
        _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
        _logoPath = logoPath;
    }

    public byte[] GenerateExcel()
    {
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Báo cáo");

            int currentRow = 1;

            // ===== HEADER =====
            if (!string.IsNullOrEmpty(_logoPath) && File.Exists(_logoPath))
            {
                var img = ws.AddPicture(_logoPath)
                    .MoveTo(ws.Cell(currentRow, 2))
                    .Scale(0.3);
            }

            // Dòng 1: Tên cơ sở + Mẫu số
            ws.Range(currentRow, 2, currentRow, 6).Merge();
            ws.Cell(currentRow, 2).Value = _dataDN.TenCSKCB ?? "";
            ws.Cell(currentRow, 2).Style.Font.FontSize = 10;
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

            ws.Range(currentRow, 7, currentRow, 7).Merge();
            ws.Cell(currentRow, 7).Value = $"Mẫu số: {_data.FirstOrDefault()?.MauSo ?? ""}";
            ws.Cell(currentRow, 7).Style.Font.FontSize = 10;
            ws.Cell(currentRow, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            currentRow++;
            ws.Range(currentRow, 2, currentRow, 7).Merge();
            ws.Cell(currentRow, 2).Value = _data[0].TenKhoHang ?? "";
            ws.Cell(currentRow, 2).Style.Font.FontSize = 10;
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            currentRow += 2;

            // ===== TIÊU ĐỀ CHÍNH =====
            ws.Range(currentRow, 2, currentRow, 8).Merge();
            ws.Cell(currentRow, 2).Value = "BẢNG KÊ BÁN LẺ HÀNG HÓA, DỊCH VỤ";
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Font.FontSize = 13;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 8).Merge();
            if (DateTime.TryParse(_ngayBatDau, out var dtStart) && DateTime.TryParse(_ngayKetThuc, out var dtEnd))
                ws.Cell(currentRow, 2).Value = $"Từ ngày {dtStart:dd-MM-yyyy} đến ngày {dtEnd:dd-MM-yyyy}";
            else
                ws.Cell(currentRow, 2).Value = $"Từ ngày {_ngayBatDau} đến ngày {_ngayKetThuc}";
            ws.Cell(currentRow, 2).Style.Font.Italic = true;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow += 2;

            // ===== THÔNG TIN CƠ SỞ =====
            ws.Range(currentRow, 2, currentRow, 6).Merge();
            ws.Cell(currentRow, 2).Value = $"Tên cơ sở kinh doanh: {_dataDN.TenCSKCB ?? ""}";
            ws.Cell(currentRow, 2).Style.Font.FontSize = 10;

            ws.Cell(currentRow, 7).Value = $"Mã số: {_data.FirstOrDefault()?.MaSo ?? ""}";
            ws.Cell(currentRow, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            ws.Cell(currentRow, 7).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 7).Merge();
            ws.Cell(currentRow, 2).Value = $"Địa chỉ: {_dataDN.DiaChi ?? ""}";
            ws.Cell(currentRow, 2).Style.Font.FontSize = 10;
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 7).Merge();
            ws.Cell(currentRow, 2).Value = $"Họ tên người bán: {_tenNhanVien}";
            ws.Cell(currentRow, 2).Style.Font.FontSize = 10;
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 7).Merge();
            ws.Cell(currentRow, 2).Value = $"Địa chỉ nơi bán: {_dataDN.DiaChi ?? ""}";
            ws.Cell(currentRow, 2).Style.Font.FontSize = 10;
            currentRow += 2;

            // ===== BẢNG DỮ LIỆU =====
            string[] headers = {
            "STT", "Tên hàng hóa dịch vụ", "ĐVT", "Số lượng", "Đơn giá bán", "Thành tiền bán"
        };

            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(currentRow, i + 2).Value = headers[i];
                ws.Cell(currentRow, i + 2).Style.Font.Bold = true;
                ws.Cell(currentRow, i + 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(currentRow, i + 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            currentRow++;

            int stt = 1;
            foreach (var item in _data)
            {
                int col = 2;
                SetMiddle(ws.Cell(currentRow, col++), stt.ToString());
                ws.Cell(currentRow, col++).Value = item.TenHangHoa ?? "";
                SetMiddle(ws.Cell(currentRow, col++), item.DVT ?? "");
                SetMiddle(ws.Cell(currentRow, col++), item.SoLuong?.ToString("N2") ?? "0");
                ws.Cell(currentRow, col++).SetValue(item.DonGiaBan?.ToString("N2") ?? "0").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                ws.Cell(currentRow, col++).SetValue(item.ThanhTien?.ToString("N2") ?? "0").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                var rangeRow = ws.Range(currentRow, 2, currentRow, col - 1);
                rangeRow.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                rangeRow.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                rangeRow.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                currentRow++;
                stt++;
            }

            // ===== TỔNG CỘNG =====
            double tongHoaDon = _data.Sum(x => x.ThanhTien ?? 0);
            ws.Range(currentRow, 2, currentRow, 6).Merge();
            ws.Cell(currentRow, 2).Value = "Tổng cộng";
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            ws.Cell(currentRow, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            ws.Cell(currentRow, 7).SetValue(tongHoaDon.ToString("N2"));
            ws.Cell(currentRow, 7).Style.Font.Bold = true;
            ws.Cell(currentRow, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            ws.Cell(currentRow, 7).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 7).Merge();
            var tongChu = ws.Cell(currentRow, 2).GetRichText();
            tongChu.AddText("Số tiền bằng chữ: ").SetBold();
            tongChu.AddText(H0304NumberToTextHelper.chuyenDoiSoTienThanhChu2(tongHoaDon.ToString())).SetBold();
            ws.Range(currentRow, 2, currentRow, 7).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            currentRow += 2;

            ws.Range(currentRow, 6, currentRow, 7).Merge();
            ws.Cell(currentRow, 6).Value =
                $"Ngày {DateTime.Now:dd} tháng {DateTime.Now:MM} năm {DateTime.Now:yyyy}";
            ws.Cell(currentRow, 6).Style.Font.Bold = true;
            ws.Cell(currentRow, 6).Style.Font.Italic = true;
            ws.Cell(currentRow, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;
            // Chữ ký 

            ws.Cell(currentRow, 3).Value = "BAN ĐIỀU HÀNH";
            ws.Cell(currentRow, 3).Style.Font.Bold = true;
            ws.Cell(currentRow, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ws.Range(currentRow, 4, currentRow, 5).Merge();
            ws.Cell(currentRow, 4).Value = "THỦ QUỸ";
            ws.Cell(currentRow, 4).Style.Font.Bold = true;
            ws.Cell(currentRow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ws.Cell(currentRow, 6).Value = "NGƯỜI BÁN";
            ws.Cell(currentRow, 6).Style.Font.Bold = true;
            ws.Cell(currentRow, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ws.Cell(currentRow, 7).Value = "KẾ TOÁN";
            ws.Cell(currentRow, 7).Style.Font.Bold = true;
            ws.Cell(currentRow, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


            ws.Columns().AdjustToContents();
            ws.Column(1).Width = 2;

            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms);
                return ms.ToArray();
            }
        }
    }

    void SetMiddle(IXLCell cell, string? value)
    {
        cell.Value = value ?? "";
        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    }

}
