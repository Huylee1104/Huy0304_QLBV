using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using H0304.NumberToText.Helpers;
using M0304.Models.BaoCaoSoLieuThuThuat;
using M0304.Models.ThongTinDoanhNghiep;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

public class P0304BaoCaoSoLieuThuThuatExcelReportTemplate
{
    private readonly List<M0304BaoCaoSoLieuThuThuat> _data;
    private string _ngayBatDau;
    private string _ngayKetThuc;
    private readonly M0304ThongTinDoanhNghiep _dataDN;
    private readonly string _logoPath;

    public P0304BaoCaoSoLieuThuThuatExcelReportTemplate(
        List<M0304BaoCaoSoLieuThuThuat> data,
        string ngayBatDau,
        string ngayKetThuc,
        M0304ThongTinDoanhNghiep dataDN,
        string logoPath = null
    )
    {
        _data = data ?? new List<M0304BaoCaoSoLieuThuThuat>();
        _ngayBatDau = ngayBatDau;
        _ngayKetThuc = ngayKetThuc;
        _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
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

            ws.Range(1, 2, 1, 23).Merge();
            ws.Cell(1, 2).Value = _dataDN.TenCoQuanChuyenMon ?? "";
            ws.Cell(1, 2).Style.Font.FontSize = 9;
            ws.Cell(1, 2).Style.Font.Bold = true;

            ws.Range(2, 2, 2, 23).Merge();
            ws.Cell(2, 2).Value = _dataDN.TenCSKCB ?? "";
            ws.Cell(2, 2).Style.Font.FontSize = 9;
            ws.Cell(2, 2).Style.Font.Bold = true;

            var dsKhoa = _data.Where(d => !string.IsNullOrWhiteSpace(d.TenKhoa))
                  .Select(d => d.TenKhoa)
                  .Distinct()
                  .ToList();

            string tenKhoa = dsKhoa.Count == 0 ? "" : dsKhoa.Count == 1 ? dsKhoa.First() : "Tất cả khoa";

            ws.Range(3, 2, 3, 23).Merge();
            ws.Cell(3, 2).Value = tenKhoa;
            ws.Cell(3, 2).Style.Font.FontSize = 9;
            ws.Cell(3, 2).Style.Font.Bold = true;

            currentRow += 4;

            ws.Range(currentRow, 2, currentRow, 15).Merge();
            ws.Cell(currentRow, 2).Value = "BÁO CÁO SỐ LIỆU THỦ THUẬT";
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Font.FontSize = 14;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 15).Merge();
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

            string[] headers = new string[]
            {
            "STT", "Số phiếu", "Thiết bị", "Mã Y Tế", "Mã đợt", "Tên bệnh nhân", "Năm sinh", "Giới tính",
            "Địa chỉ", "Tên dịch vụ", "Đối tượng", "Phương pháp vô cảm", "Loại thủ thuật", "Bác sĩ thực hiện",
            "ĐD/KTV", "Ngày chỉ định", "Ngày thực hiện", "Nơi yêu cầu", "BS chỉ định", "Nợi thực hiện", "Loại giá", "Mã hóa đơn"
            };
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(currentRow, i + 2).Value = headers[i];
                ws.Cell(currentRow, i + 2).Style.Font.Bold = true;
                ws.Cell(currentRow, i + 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(currentRow, i + 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }
            currentRow++;
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(currentRow, i + 2).Value = i + 1;
                ws.Cell(currentRow, i + 2).Style.Font.Bold = true;
                ws.Cell(currentRow, i + 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(currentRow, i + 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            int stt = 1;

            currentRow++;

            foreach (var item in _data)
            {
                int col = 2;
                ws.Cell(currentRow, col++).Value= stt++;
                ws.Cell(currentRow, col++).Value = item.SoPhieu ?? "";
                var thietBi = item.ThietBi ?? "";
                var maxLength = 50;
                var sb = new StringBuilder();
                var lastBreak = 0;

                while (lastBreak < thietBi.Length)
                {
                    if (lastBreak + maxLength >= thietBi.Length)
                    {
                        sb.Append(thietBi.Substring(lastBreak));
                        break;
                    }

                    int breakIndex = thietBi.LastIndexOf(' ', lastBreak + maxLength, maxLength);
                    if (breakIndex <= lastBreak)
                        breakIndex = lastBreak + maxLength;

                    sb.Append(thietBi.Substring(lastBreak, breakIndex - lastBreak));
                    sb.Append("\n");
                    lastBreak = breakIndex + 1;
                }

                var cell2 = ws.Cell(currentRow, col++);
                cell2.Value = sb.ToString();
                cell2.Style.Alignment.WrapText = true;
                cell2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                ws.Cell(currentRow, col++).Value = item.MaYTe ?? "";
                ws.Cell(currentRow, col++).Value = item.MaDot ?? "";
                ws.Cell(currentRow, col++).Value = item.TenBenhNhan ?? "";
                SetMiddle(ws.Cell(currentRow, col++), item.NamSinh.ToString() ?? "");
                SetMiddle(ws.Cell(currentRow, col++), item.GioiTinh ?? "");
                var diaChi = item.DiaChi ?? "";
                maxLength = 50;
                sb = new StringBuilder();
                lastBreak = 0;

                while (lastBreak < diaChi.Length)
                {
                    if (lastBreak + maxLength >= diaChi.Length)
                    {
                        sb.Append(diaChi.Substring(lastBreak));
                        break;
                    }

                    int breakIndex = diaChi.LastIndexOf(' ', lastBreak + maxLength, maxLength);
                    if (breakIndex <= lastBreak)
                        breakIndex = lastBreak + maxLength;

                    sb.Append(diaChi.Substring(lastBreak, breakIndex - lastBreak));
                    sb.Append("\n");
                    lastBreak = breakIndex + 1;
                }

                var cell3 = ws.Cell(currentRow, col++);
                cell3.Value = sb.ToString();
                cell3.Style.Alignment.WrapText = true;
                cell3.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                var tenDichVu = item.TenDichVu ?? "";
                maxLength = 50;
                sb = new StringBuilder();
                lastBreak = 0;

                while (lastBreak < tenDichVu.Length)
                {
                    if (lastBreak + maxLength >= tenDichVu.Length)
                    {
                        sb.Append(tenDichVu.Substring(lastBreak));
                        break;
                    }

                    int breakIndex = tenDichVu.LastIndexOf(' ', lastBreak + maxLength, maxLength);
                    if (breakIndex <= lastBreak)
                        breakIndex = lastBreak + maxLength;

                    sb.Append(tenDichVu.Substring(lastBreak, breakIndex - lastBreak));
                    sb.Append("\n");
                    lastBreak = breakIndex + 1;
                }

                var cell4 = ws.Cell(currentRow, col++);
                cell4.Value = sb.ToString();
                cell4.Style.Alignment.WrapText = true;
                cell4.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                ws.Cell(currentRow, col++).Value = item.DoiTuong ?? "";
                ws.Cell(currentRow, col++).Value = item.PhuongPhapVoCam ?? "";
                ws.Cell(currentRow, col++).Value = item.LoaiThuThuat ?? "";
                ws.Cell(currentRow, col++).Value = item.BacSiThucHien ?? "";
                ws.Cell(currentRow, col++).Value = item.DieuDuong ?? "";
                SetMiddle(ws.Cell(currentRow, col++), item.NgayChiDinh?.ToString("dd-MM-yyyy") ?? "");
                SetMiddle(ws.Cell(currentRow, col++), item.NgayThucHien?.ToString("dd-MM-yyyy") ?? "");
                ws.Cell(currentRow, col++).Value = item.NoiYeuCau ?? "";
                ws.Cell(currentRow, col++).Value = item.BacSiChiDinh ?? "";
                ws.Cell(currentRow, col++).Value = item.NoiThucHien?? "";
                ws.Cell(currentRow, col++).Value = item.LoaiGia ?? "";
                ws.Cell(currentRow, col++).Value = item.MaHoaDon ?? "";

                var rangeRow = ws.Range(currentRow, 2, currentRow, col - 1);
                rangeRow.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                rangeRow.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                rangeRow.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                currentRow++;
            }

            ws.Columns().AdjustToContents();
            ws.Column(1).Width = 2;
            ws.Column(2).Width = 12;


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
