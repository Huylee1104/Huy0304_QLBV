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
    private string _tenNVDN;
    private readonly string _logoPath;

    public P0304BCThongKeTheoMaBenhICDExcelReportTemplate(
        List<M0304BCThongKeTheoMaBenhICD> data,
        string ngayBatDau,
        string ngayKetThuc,
        string tenNVDN,
        M0304ThongTinDoanhNghiep dataDN,
        string logoPath = null
    )
    {
        _data = data ?? new List<M0304BCThongKeTheoMaBenhICD>();
        _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
        _ngayBatDau = ngayBatDau;
        _ngayKetThuc = ngayKetThuc;
        _tenNVDN = tenNVDN;
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
            ws.Cell(1, 2).Style.Font.Bold = true;

            ws.Range(2, 2, 2, 9).Merge();
            ws.Cell(2, 2).Value = _dataDN.TenCoQuanChuyenMon ?? "";
            ws.Cell(2, 2).Style.Font.FontSize = 9;
            ws.Cell(2, 2).Style.Font.Bold = true;

            currentRow += 3;

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
                ws.Cell(currentRow, 2).Value = $"TỪ NGÀY {dtStart:dd-MM-yyyy} ĐẾN NGÀY {dtEnd:dd-MM-yyyy}";
            }
            else
            {
                ws.Cell(currentRow, 2).Value = $"TỪ NGÀY {_ngayBatDau} ĐẾN NGÀY {_ngayKetThuc}";
            }
            ws.Cell(currentRow, 2).Style.Font.FontSize = 10;
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            // Header
            ws.Cell(currentRow, 2).Value = "STT";
            ws.Cell(currentRow, 3).Value = "ICD";
            ws.Cell(currentRow, 4).Value = "Tên bệnh";
            ws.Cell(currentRow, 5).Value = "Tổng số";
            ws.Cell(currentRow, 6).Value = "Giới tính";
            ws.Cell(currentRow, 8).Value = "Có BHYT";
            ws.Cell(currentRow, 9).Value = "Không BHYT";

            // Merge header
            ws.Range(currentRow, 2, currentRow + 1, 2).Merge();
            ws.Range(currentRow, 3, currentRow + 1, 3).Merge();
            ws.Range(currentRow, 4, currentRow + 1, 4).Merge();
            ws.Range(currentRow, 5, currentRow + 1, 5).Merge();
            ws.Range(currentRow, 6, currentRow, 7).Merge();
            ws.Range(currentRow, 8, currentRow + 1, 8).Merge();
            ws.Range(currentRow, 9, currentRow + 1, 9).Merge();

            ws.Cell(currentRow + 1, 6).Value = "Nam";
            ws.Cell(currentRow + 1, 7).Value = "Nữ";

            // Header style
            var headerRange = ws.Range(currentRow, 2, currentRow + 1, 9);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            headerRange.Style.Alignment.WrapText = true;
            headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            long stt = 1;

            currentRow+=2;

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
            currentRow--;
            //var TongLuot = _data.Sum(x => x.SoLuotTiepNhan);
            //var TongNam = _data.Sum(x => x.SoLuongNam);
            //var TongNu = _data.Sum(x => x.SoLuongNu);
            //var TongBHYT = _data.Sum(x => x.CoBHYT);
            //var TongKhongBHYT = _data.Sum(x => x.KhongBHYT);

            //int colFirst = 2;

            //ws.Range(currentRow, colFirst, currentRow, colFirst + 2).Merge();

            //var cellLabel = ws.Cell(currentRow, colFirst);
            //cellLabel.Value = "Tổng Cộng";
            //cellLabel.Style.Font.Bold = true;
            //cellLabel.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //cellLabel.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            //cellLabel.Style.Font.Bold = true;

            //// Cột 7: tổng giá trị
            //SetCenterCellNumber(ws.Cell(currentRow, colFirst + 3), TongLuot);
            //SetCenterCellNumber(ws.Cell(currentRow, colFirst + 4), TongNam);
            //SetCenterCellNumber(ws.Cell(currentRow, colFirst + 5), TongNu);
            //SetCenterCellNumber(ws.Cell(currentRow, colFirst + 6), TongBHYT);
            //SetCenterCellNumber(ws.Cell(currentRow, colFirst + 7), TongKhongBHYT);

            ws.Range(2, 2, currentRow, 9).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(2, 2, currentRow, 9).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            currentRow+=2;

            var cell2 = ws.Range(currentRow, 6, currentRow + 5, 9);
            cell2.Merge();
            var rt = cell2.FirstCell().GetRichText();
            rt.AddText($"Ngày {DateTime.Now:dd} Tháng {DateTime.Now:MM} Năm {DateTime.Now:yyyy}\n");
            rt.AddText("Người lập bảng\n\n\n");
            rt.AddText($"{_tenNVDN}");
            cell2.Style.Alignment.WrapText = true;
            cell2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            ws.Columns().AdjustToContents();
            ws.Column(1).Width = 2;
            ws.Column(2).Width = 12;
            ws.Column(5).Width = 12;
            ws.Column(6).Width = 12;
            ws.Column(7).Width = 12;
            ws.Column(8).Width = 12;
            ws.Column(9).Width = 12;

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
