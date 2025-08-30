using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using M0304.Models.ThongTinDoanhNghiep;
using M0304C.Models.BaoCaoThuDichVu;
using System;
using System.Text;

public class P0304CExcelReportTemplate
{
    private readonly List<M0304CBaoCaoThuDichVu> _data;
    private readonly M0304ThongTinDoanhNghiep _dataDN;
    private string _ngayBatDau;
    private string _ngayKetThuc;
    private readonly string _logoPath;

    public P0304CExcelReportTemplate(
        List<M0304CBaoCaoThuDichVu> data,
        string ngayBatDau,
        string ngayKetThuc,
        M0304ThongTinDoanhNghiep dataDN,
        string logoPath = null
    )
    {
        _data = data ?? new List<M0304CBaoCaoThuDichVu>();
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

            ws.Range(1, 3, 1, 6).Merge();
            ws.Cell(1, 3).Value = _dataDN.TenCSKCB ?? "";
            ws.Cell(1, 2).Style.Font.FontSize = 9;

            ws.Range(2, 3, 2, 6).Merge();
            ws.Cell(2, 3).Value = _dataDN.TenCoQuanChuyenMon ?? "";
            ws.Cell(2, 3).Style.Font.FontSize = 9;

            ws.Range(3, 3, 3, 6).Merge();
            ws.Cell(3, 3).Value = _dataDN.DiaChi ?? "";
            ws.Cell(3, 3).Style.Font.FontSize = 9;

            ws.Range(4, 3, 4, 6).Merge();
            ws.Cell(4, 3).Value = _dataDN.DienThoai ?? "";
            ws.Cell(4, 3).Style.Font.FontSize = 9;

            currentRow += 5;

            ws.Range(currentRow, 2, currentRow, 6).Merge();
            ws.Cell(currentRow, 2).Value = "BÁO CÁO THU DỊCH VỤ";
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Font.FontSize = 14;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 6).Merge();
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
            "STT", "Nhóm dịch vụ", "Dịch vụ", "Số lượng", "Tổng hóa đơn"
            };
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(currentRow, i + 2).Value = headers[i];
                ws.Cell(currentRow, i + 2).Style.Font.Bold = true;
                ws.Cell(currentRow, i + 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            int stt = 1;

            currentRow++;

            foreach (var item in _data)
            {
                int col = 2;

                SetCenterCellNumber(ws.Cell(currentRow, col++), stt++);
                SetMiddle(ws.Cell(currentRow, col++), item.NhomDichVu ?? "");

                var dichVu = item.DichVu ?? "";
                int maxLength = 60;
                var sb = new StringBuilder();
                int lastBreak = 0;

                while (lastBreak < dichVu.Length)
                {
                    if (lastBreak + maxLength >= dichVu.Length)
                    {
                        sb.Append(dichVu.Substring(lastBreak));
                        break;
                    }

                    int breakIndex = dichVu.LastIndexOf(' ', lastBreak + maxLength, maxLength);
                    if (breakIndex <= lastBreak)
                        breakIndex = lastBreak + maxLength;

                    sb.Append(dichVu.Substring(lastBreak, breakIndex - lastBreak));
                    sb.Append("\n");
                    lastBreak = breakIndex + 1;
                }

                var cell = ws.Cell(currentRow, col++);
                cell.Value = sb.ToString();
                cell.Style.Alignment.WrapText = true;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                SetCenterCellNumber(ws.Cell(currentRow, col++), item.SoLuong);
                SetNumberCell(ws.Cell(currentRow, col++), item.TongHoaDon);
                currentRow++;
            }

            currentRow++;
            var tongTatCaHoaDon = _data.Sum(x => x.TongHoaDon);

            int colFirst = 2;
            int colLast = 6; 

            ws.Range(currentRow, colFirst, currentRow, colFirst + 2).Merge();

            var cellLabel = ws.Cell(currentRow, colFirst);
            cellLabel.Value = "Tổng Cộng";
            cellLabel.Style.Font.Bold = true;
            cellLabel.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cellLabel.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            ws.Cell(currentRow, colFirst + 3).Value = "";

            // Cột 7: tổng giá trị
            SetNumberCell(ws.Cell(currentRow, colLast), tongTatCaHoaDon);

            // Kẻ border cho cả dòng
            ws.Range(8, 2, currentRow, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(8, 2, currentRow, 6).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

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
