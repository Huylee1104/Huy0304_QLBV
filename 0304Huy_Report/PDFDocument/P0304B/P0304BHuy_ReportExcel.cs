using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using H0304.NumberToText.Helpers;
using M0304.Models.ThongTinDoanhNghiep;
using M0304B.Models.BCHoaDonDienTuDV;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

public class P0304BExcelReportTemplate
{
    private readonly List<M0304BBCHoaDonDienTuDV> _data;
    private readonly M0304ThongTinDoanhNghiep _dataDN;
    private string _ngayBatDau;
    private string _ngayKetThuc;
    private readonly string _logoPath;

    public P0304BExcelReportTemplate(
        List<M0304BBCHoaDonDienTuDV> data,
        string ngayBatDau,
        string ngayKetThuc,
        M0304ThongTinDoanhNghiep dataDN,
        string logoPath = null
    )
    {
        _data = data ?? new List<M0304BBCHoaDonDienTuDV>();
        _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
        _ngayBatDau = ngayBatDau;
        _ngayKetThuc = ngayKetThuc;
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
                //ws.Range(1, 1, 2, 2).Merge(); // dòng dầu, cột đầu, dòng cuối, cột cuối
                var img = ws.AddPicture(_logoPath)
                    .MoveTo(ws.Cell(1, 2))
                    .Scale(0.2);
                ws.Row(1).AdjustToContents();
            }

            ws.Range(1, 3, 1, 13).Merge();
            ws.Cell(1, 3).Value = _dataDN.TenCSKCB ?? "";
            ws.Cell(1, 2).Style.Font.FontSize = 9;

            ws.Range(2, 3, 2, 13).Merge();
            ws.Cell(2, 3).Value = _dataDN.TenCoQuanChuyenMon ?? "";
            ws.Cell(2, 3).Style.Font.FontSize = 9;

            ws.Range(3, 3, 3, 13).Merge();
            ws.Cell(3, 3).Value = _dataDN.DiaChi ?? "";
            ws.Cell(3, 3).Style.Font.FontSize = 9;

            ws.Range(4, 3, 4, 13).Merge();
            ws.Cell(4, 3).Value = _dataDN.DienThoai ?? "";
            ws.Cell(4, 3).Style.Font.FontSize = 9;

            currentRow += 5;

            ws.Range(currentRow, 2, currentRow, 13).Merge();
            ws.Cell(currentRow, 2).Value = "BÁO CÁO HÓA ĐƠN ĐIỆN TỬ DỊCH VỤ";
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Font.FontSize = 14;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 13).Merge();
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
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            string[] headers = new string[]
            {
            "STT", "Số chứng từ", "Ngày thu", "Giá trị", "Mã bệnh nhân",
            "Tên bệnh nhân", "Năm sinh", "Địa chỉ", "Ngày tạo HDDT", "E_InvoiceNo", "Giá trị HDDT", "Mã truy cứu"
            };
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(currentRow, i + 2).Value = headers[i];
                ws.Cell(currentRow, i + 2).Style.Font.Bold = true;
                ws.Cell(currentRow, i + 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            currentRow++;

            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(currentRow, i + 2).Value = i+1;
                ws.Cell(currentRow, i + 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            int stt = 1;

            currentRow++;

            foreach (var item in _data)
            {
                int col = 2;

                SetCenterCellNumber(ws.Cell(currentRow, col++), stt++);
                SetMiddle(ws.Cell(currentRow, col++), item.SoChungTu ?? "");
                SetDateCell(ws.Cell(currentRow, col++), item.NgayThu);
                SetNumberCell(ws.Cell(currentRow, col++), item.GiaTri);
                SetMiddle(ws.Cell(currentRow, col++), item.MaBenhNhan ?? "");
                SetMiddle(ws.Cell(currentRow, col++), item.TenBenhNhan ?? "");
                SetCenterCell(ws.Cell(currentRow, col++), item.NamSinh?.ToString() ?? "");

                var diaChi = item.DiaChi ?? "";
                int maxLength = 50;
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

                var cell = ws.Cell(currentRow, col++);
                cell.Value = sb.ToString();
                cell.Style.Alignment.WrapText = true;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                SetDateCell(ws.Cell(currentRow, col++), item.NgayTaoHDDT);
                SetCenterCell(ws.Cell(currentRow, col++), item.E_InvoiceNo);
                SetNumberCell(ws.Cell(currentRow, col++), item.GiaTriHDDT);
                SetCenterCell(ws.Cell(currentRow, col++), item.MaTraCuu);
                currentRow++;
            }

            var tongGiaTri = _data.Sum(x => x.GiaTri ?? 0);
            var tongGiaTriHDDT = _data.Sum(x => x.GiaTriHDDT ?? 0);

            int colIndex = 4;
            var cellTongGT = ws.Cell(currentRow, colIndex++);
            SetNumberCell(cellTongGT, tongGiaTri);
            cellTongGT.Style.Font.Bold = true;

            colIndex = 11;
            var cellTongGTHDDT = ws.Cell(currentRow, colIndex++);
            SetNumberCell(cellTongGTHDDT, tongGiaTriHDDT);
            cellTongGTHDDT.Style.Font.Bold = true;

            int firstRow = 8;
            int lastRow = currentRow;
            int firstCol = 2;
            int lastCol = 13;

            var fullRange = ws.Range(firstRow, firstCol, lastRow, lastCol);
            fullRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            currentRow++;
            ws.Range(currentRow, 10, currentRow + 5, 13).Merge();
            ws.Cell(currentRow, 10).Value =
                $"Ngày {DateTime.Now:dd} Tháng {DateTime.Now:MM} Năm {DateTime.Now:yyyy}\n" +
                "Người lập bảng\n\n\n" +
                "Trần Thanh Thảo";
            ws.Cell(currentRow, 10).Style.Alignment.WrapText = true;
            ws.Cell(currentRow, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(currentRow, 10).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            ws.Columns().AdjustToContents();
            ws.Column(1).Width = 2;
            ws.Column(2).Width = 12;

            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms);
                return ms.ToArray();
            }

            void SetDateCell(IXLCell cell, DateTime? date)
            {
                if (date.HasValue)
                {
                    cell.Value = date.Value;
                    cell.Style.DateFormat.Format = "dd-MM-yyyy";
                }
                else
                {
                    cell.Value = "";
                }
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }

            void SetNumberCell(IXLCell cell, decimal? number)
            {
                cell.Value = number ?? 0;
                cell.Style.NumberFormat.Format = "#,##0";
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }

            void SetCenterCell(IXLCell cell, string? value)
            {
                cell.Value = value ?? "";
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }

            void SetCenterCellNumber(IXLCell cell, int? value)
            {
                cell.Value = value ?? 0;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }

            void SetMiddle(IXLCell cell, string? value)
            {
                cell.Value = value ?? "";
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }
        }
    }
}
