using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using M0304.Models.ThongTinDoanhNghiep;
using M0304.Models.BCThongKeTheoMaBenhICD;
using System;
using System.Text;

public class P0304BCThongKeTheoMaBenhICDExcelReportTemplate
{
    private readonly List<M0304BCThongKeTheoMaBenhICD> _data;
    private readonly M0304ThongTinDoanhNghiep _dataDN;
    private string _ngayBatDau;
    private string _ngayKetThuc;
    private readonly string _logoPath;

    public P0304BCThongKeTheoMaBenhICDExcelReportTemplate(
        List<M0304BCThongKeTheoMaBenhICD> data,
        string ngayBatDau,
        string ngayKetThuc,
        M0304ThongTinDoanhNghiep dataDN,
        string logoPath = null
    )
    {
        _data = data ?? new List<M0304BCThongKeTheoMaBenhICD>();
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

            ws.Range(1, 2, 1, 9).Merge();
            ws.Cell(1, 2).Value = _dataDN.TenCSKCB ?? "";
            ws.Cell(1, 2).Style.Font.FontSize = 9;

            ws.Range(2, 2, 2, 9).Merge();
            ws.Cell(2, 2).Value = _dataDN.TenCoQuanChuyenMon ?? "";
            ws.Cell(2, 2).Style.Font.FontSize = 9;

            ws.Range(3, 2, 3, 9).Merge();
            ws.Cell(3, 2).Value = _dataDN.DiaChi ?? "";
            ws.Cell(3, 2).Style.Font.FontSize = 9;

            ws.Range(4, 2, 4, 9).Merge();
            ws.Cell(4, 2).Value = _dataDN.DienThoai ?? "";
            ws.Cell(4, 2).Style.Font.FontSize = 9;

            currentRow += 5;

            ws.Range(currentRow, 2, currentRow, 9).Merge();
            ws.Cell(currentRow, 2).Value = "BÁO CÁO TỔNG HỢP SỐ LIỆU KHÁM BỆNH THEO NHIỀU TIÊU CHÍ";
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Font.FontSize = 14;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 9).Merge();
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

            ws.Cell(currentRow, 2).Value = "STT";
            ws.Cell(currentRow, 3).Value = "ICD";
            ws.Cell(currentRow, 4).Value = "Tên bệnh";
            ws.Cell(currentRow, 5).Value = "Tổng số";
            ws.Cell(currentRow, 6).Value = "Giới tính";
            ws.Cell(currentRow, 8).Value = "Có BHYT";
            ws.Cell(currentRow, 9).Value = "Không BHYT";

            ws.Range(currentRow, 2, currentRow + 1, 2).Merge();
            ws.Range(currentRow, 3, currentRow + 1, 3).Merge();
            ws.Range(currentRow, 4, currentRow + 1, 4).Merge();
            ws.Range(currentRow, 5, currentRow + 1, 5).Merge();
            ws.Range(currentRow, 6, currentRow, 7).Merge();
            ws.Range(currentRow, 8, currentRow + 1, 8).Merge();
            ws.Range(currentRow, 9, currentRow + 1, 9).Merge();

            ws.Cell(currentRow + 1, 6).Value = "Nam";
            ws.Cell(currentRow + 1, 7).Value = "Nữ";

            var headerRange = ws.Range(currentRow, 2, currentRow + 1, 9);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
            headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            long stt = 1;

            currentRow++;

            foreach (var item in _data)
            {
                int col = 2;

                SetCenterCellNumber(ws.Cell(currentRow, col++), stt++);
                SetMiddle(ws.Cell(currentRow, col++), item.TenICD ?? "");

                var tenBenh = item.TenBenh ?? "";
                int maxLength = 60;
                var sb = new StringBuilder();
                int lastBreak = 0;

                while (lastBreak < tenBenh.Length)
                {
                    if (lastBreak + maxLength >= tenBenh.Length)
                    {
                        sb.Append(tenBenh.Substring(lastBreak));
                        break;
                    }

                    int breakIndex = tenBenh.LastIndexOf(' ', lastBreak + maxLength, maxLength);
                    if (breakIndex <= lastBreak)
                        breakIndex = lastBreak + maxLength;

                    sb.Append(tenBenh.Substring(lastBreak, breakIndex - lastBreak));
                    sb.Append("\n");
                    lastBreak = breakIndex + 1;
                }

                var cell = ws.Cell(currentRow, col++);
                cell.Value = sb.ToString();
                cell.Style.Alignment.WrapText = true;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                SetCenterCellNumber(ws.Cell(currentRow, col++), item.SoLuotTiepNhan);
                SetCenterCellNumber(ws.Cell(currentRow, col++), item.SoLuongNam);
                SetCenterCellNumber(ws.Cell(currentRow, col++), item.SoLuongNu);
                SetCenterCellNumber(ws.Cell(currentRow, col++), item.CoBHYT);
                SetCenterCellNumber(ws.Cell(currentRow, col++), item.KhongBHYT);
                currentRow++;
            }

            currentRow++;
            var TongLuot = _data.Sum(x => x.SoLuotTiepNhan);
            var TongNam = _data.Sum(x => x.SoLuongNam);
            var TongNu = _data.Sum(x => x.SoLuongNu);
            var TongBHYT = _data.Sum(x => x.CoBHYT);
            var TongKhongBHYT = _data.Sum(x => x.KhongBHYT);

            int colFirst = 2;

            ws.Range(currentRow, colFirst, currentRow, colFirst + 2).Merge();

            var cellLabel = ws.Cell(currentRow, colFirst);
            cellLabel.Value = "Tổng Cộng";
            cellLabel.Style.Font.Bold = true;
            cellLabel.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cellLabel.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            // Cột 7: tổng giá trị
            SetCenterCellNumber(ws.Cell(currentRow, colFirst + 3), TongLuot);
            SetCenterCellNumber(ws.Cell(currentRow, colFirst + 4), TongNam);
            SetCenterCellNumber(ws.Cell(currentRow, colFirst + 5), TongNu);
            SetCenterCellNumber(ws.Cell(currentRow, colFirst + 6), TongBHYT);
            SetCenterCellNumber(ws.Cell(currentRow, colFirst + 7), TongKhongBHYT);

            ws.Range(2, 2, currentRow, 9).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(2, 2, currentRow, 9).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            ws.Range(2, 2, currentRow, 9).Style.Font.Bold = true;

            ws.Columns().AdjustToContents();
            ws.Column(1).Width = 2;
            ws.Column(2).Width = 12;

            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms);
                return ms.ToArray();
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

            void SetCenterCellNumber(IXLCell cell, long? value)
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
