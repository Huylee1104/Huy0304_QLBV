using ClosedXML.Excel;
using H0304.NumberToText.Helpers;
using M0304.Models.BCTTCapPhatThuocKS_KVirut;
using M0304.Models.ThongTinDoanhNghiep;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

public class P0304BCTTCapPhatThuocKS_KVirutExcelReportTemplate
{
    private readonly List<M0304BCTTCapPhatThuocKS_KVirut> _data;
    private string _ngayBatDau;
    private string _ngayKetThuc;
    private readonly M0304ThongTinDoanhNghiep _dataDN;
    private readonly string _logoPath;

    public P0304BCTTCapPhatThuocKS_KVirutExcelReportTemplate(
        List<M0304BCTTCapPhatThuocKS_KVirut> data,
        string ngayBatDau,
        string ngayKetThuc,
        M0304ThongTinDoanhNghiep dataDN,
        string logoPath = null
    )
    {
        _data = data ?? new List<M0304BCTTCapPhatThuocKS_KVirut>();
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

            ws.Range(1, 3, 1, 20).Merge();
            ws.Cell(1, 3).Value = _dataDN.TenCoQuanChuyenMon ?? "";
            ws.Cell(1, 3).Style.Font.FontSize = 9;
            ws.Cell(1, 3).Style.Font.Bold = true;

            ws.Range(2, 3, 2, 20).Merge();
            ws.Cell(2, 3).Value = _dataDN.TenCSKCB ?? "";
            ws.Cell(2, 3).Style.Font.FontSize = 9;
            ws.Cell(2, 3).Style.Font.Bold = true;

            currentRow += 4;

            ws.Range(currentRow, 2, currentRow, 20).Merge();
            ws.Cell(currentRow, 2).Value = "BÁO CÁO THÔNG TIN CẤP PHÁT THUỐC KHÁNG SINH, THUỐC KHÁNG VI RÚT";
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Font.FontSize = 14;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 20).Merge();
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
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            string[] headers = new string[]
            {
            "STT", "Mã y tế", "Số bệnh án", "Tên bệnh nhân", "Năm sinh", "Địa chỉ", "Khoa điều trị", "Tên phòng khám",
            "Bác sĩ kê đơn", "Tên thuốc", "Tên hoạt chất", "Hàm lượng", "ĐVT", "Đường dùng", "Liều dùng",
            "Số ngày", "Số lượng kê đơn", "Số lượng phát", "Chẩn đoán"
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

            foreach (var item in _data)
            {
                int col = 2;
                ws.Cell(currentRow, col++).Value =  stt++;
                SetMiddle(ws.Cell(currentRow, col++), item.MaYTe ?? "");
                SetMiddle(ws.Cell(currentRow, col++), item.SoBenhAn ?? "");
                ws.Cell(currentRow, col++).Value = item.TenBenhNhan ?? "";
                SetMiddle(ws.Cell(currentRow, col++), item.NamSinh.ToString() ?? "");
                var diaChi = item.DiaChi ?? "";
                int maxLength = 60;
                var sb = new StringBuilder();
                int lastBreak = 0;

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
                ws.Cell(currentRow, col++).Value = item.KhoaDieuTri ?? "";
                ws.Cell(currentRow, col++).Value = item.TenPhongKham ?? "";
                ws.Cell(currentRow, col++).Value = item.BacSiKeDon ?? "";
                ws.Cell(currentRow, col++).Value = item.TenThuoc ?? "";
                var hoatChat = item.TenHoatChat ?? "";
                maxLength = 60;
                sb = new StringBuilder();
                lastBreak = 0;

                while (lastBreak < hoatChat.Length)
                {
                    if (lastBreak + maxLength >= hoatChat.Length)
                    {
                        sb.Append(hoatChat.Substring(lastBreak));
                        break;
                    }

                    int breakIndex = hoatChat.LastIndexOf(' ', lastBreak + maxLength, maxLength);
                    if (breakIndex <= lastBreak)
                        breakIndex = lastBreak + maxLength;

                    sb.Append(hoatChat.Substring(lastBreak, breakIndex - lastBreak));
                    sb.Append("\n");
                    lastBreak = breakIndex + 1;
                }

                var cell4 = ws.Cell(currentRow, col++);
                cell4.Value = sb.ToString();
                cell4.Style.Alignment.WrapText = true;
                cell4.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                ws.Cell(currentRow, col++).Value = item.HamLuong ?? "";
                ws.Cell(currentRow, col++).Value = item.DVT ?? "";
                ws.Cell(currentRow, col++).Value = item.DuongDung ?? "";
                ws.Cell(currentRow, col++).Value = item.LieuDung ?? "";
                ws.Cell(currentRow, col++).Value = item.SoNgay ?? 0;
                ws.Cell(currentRow, col++).Value = item.SoLuongKeDon ?? 0;
                ws.Cell(currentRow, col++).Value = item.SoLuongXuat ?? 0;
                var chanDoan = item.ChanDoan ?? "";
                maxLength = 60;
                sb = new StringBuilder();
                lastBreak = 0;

                while (lastBreak < chanDoan.Length)
                {
                    if (lastBreak + maxLength >= chanDoan.Length)
                    {
                        sb.Append(chanDoan.Substring(lastBreak));
                        break;
                    }

                    int breakIndex = chanDoan.LastIndexOf(' ', lastBreak + maxLength, maxLength);
                    if (breakIndex <= lastBreak)
                        breakIndex = lastBreak + maxLength;

                    sb.Append(chanDoan.Substring(lastBreak, breakIndex - lastBreak));
                    sb.Append("\n");
                    lastBreak = breakIndex + 1;
                }

                var cell2 = ws.Cell(currentRow, col++);
                cell2.Value = sb.ToString();
                cell2.Style.Alignment.WrapText = true;
                cell2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                var rangeRow = ws.Range(currentRow, 2, currentRow, col - 1);
                rangeRow.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                rangeRow.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                rangeRow.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                currentRow++;
            }

            var rangeNgay = ws.Range(currentRow, 16, currentRow, 20);
            rangeNgay.Merge();

            var today = DateTime.Now;
            string ngayViet = $"Ngày {today.Day} tháng {today.Month} năm {today.Year}";


            rangeNgay.Value = ngayViet;

            rangeNgay.Style.Font.Italic = true;
            rangeNgay.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rangeNgay.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

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
