using ClosedXML.Excel;
using H0304.NumberToText.Helpers;
using M0304.Models.BaoCaoCongTacKeDon;
using M0304.Models.ThongTinDoanhNghiep;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

public class P0304BaoCaoCongTacKeDonExcelReportTemplate
{
    private readonly List<M0304BaoCaoCongTacKeDon> _data;
    private string _ngayBatDau;
    private string _ngayKetThuc;
    private readonly M0304ThongTinDoanhNghiep _dataDN;
    private readonly string _logoPath;

    public P0304BaoCaoCongTacKeDonExcelReportTemplate(
        List<M0304BaoCaoCongTacKeDon> data,
        string ngayBatDau,
        string ngayKetThuc,
        M0304ThongTinDoanhNghiep dataDN,
        string logoPath = null
    )
    {
        _data = data ?? new List<M0304BaoCaoCongTacKeDon>();
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
            ws.Cell(currentRow, 2).Value = "BÁO CÁO CÔNG TÁC KÊ ĐƠN";
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
            ws.Cell(currentRow, 2).Style.Font.Italic = true;
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            string[] headers = new string[]
            {
            "Mã y tế", "Tên bệnh nhân", "Năm sinh", "Giới tính", "Đối tương", "Số lưu trữ", "Số bệnh án", "Khoa điều trị",
            "Ngày khám", "Tên phòng khám", "Bác sĩ kê toa", "Tên dược đầy đủ", "Tên hoạt chất", "Số ngày", "Số lượng", "Số lượng phát",
            "Đơn giá", "Chuẩn đoán", "Mục đích xuất"
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
                SetMiddle(ws.Cell(currentRow, col++), item.MaYTe ?? "");
                ws.Cell(currentRow, col++).Value = item.TenBenhNhan ?? "";
                SetMiddle(ws.Cell(currentRow, col++), item.NamSinh.ToString() ?? "");
                ws.Cell(currentRow, col++).Value = item.GioiTinh?? "";
                ws.Cell(currentRow, col++).Value = item.DoiTuong?? "";
                ws.Cell(currentRow, col++).Value = item.SoLuuTru?? "";
                ws.Cell(currentRow, col++).Value = item.SoBenhAn?? "";
                ws.Cell(currentRow, col++).Value = item.KhoaDieuTri?? "";
                SetMiddle(ws.Cell(currentRow, col++), item.NgayKham ?? "");
                ws.Cell(currentRow, col++).Value = item.TenPhongKham ?? "";
                ws.Cell(currentRow, col++).Value = item.BacSiKeToa ?? "";
                ws.Cell(currentRow, col++).Value = item.TenThuoc ?? "";

                var hoatChat = item.TenHoatChat ?? "";
                int maxLength = 60;
                var sb = new StringBuilder();
                int lastBreak = 0;

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

                var cell3 = ws.Cell(currentRow, col++);
                cell3.Value = sb.ToString();
                cell3.Style.Alignment.WrapText = true;
                cell3.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                ws.Cell(currentRow, col++).Value = item.SoNgay ?? 0;
                ws.Cell(currentRow, col++).Value = item.SoLuong ?? 0;
                ws.Cell(currentRow, col++).Value = item.SoLuongPhat ?? 0;
                var cell = ws.Cell(currentRow, col++);cell.Value = item.DonGia ?? 0;cell.Style.NumberFormat.Format = "#,##0";

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

                ws.Cell(currentRow, col++).Value = item.MucDichXuat ?? "";

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
