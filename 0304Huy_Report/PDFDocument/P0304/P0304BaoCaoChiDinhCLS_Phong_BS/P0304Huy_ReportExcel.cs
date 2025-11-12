using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using M0304.Models.ThongTinDoanhNghiep;
using M0304.Models.BaoCaoChiDinhCLS_Phong_BS;
using System;
using System.Text;

public class P0304BaoCaoChiDinhCLS_Phong_BSExcelReportTemplate
{
    private readonly List<M0304BaoCaoChiDinhCLS_Phong_BS> _data;
    private readonly M0304ThongTinDoanhNghiep _dataDN;
    private string _ngayBatDau;
    private string _ngayKetThuc;
    private readonly string _logoPath;

    public P0304BaoCaoChiDinhCLS_Phong_BSExcelReportTemplate(
        List<M0304BaoCaoChiDinhCLS_Phong_BS> data,
        string ngayBatDau,
        string ngayKetThuc,
        M0304ThongTinDoanhNghiep dataDN,
        string logoPath = null
    )
    {
        _data = data ?? new List<M0304BaoCaoChiDinhCLS_Phong_BS>();
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

            ws.Range(1, 2, 1, 7).Merge();
            ws.Cell(1, 2).Value = _dataDN.TenCSKCB ?? "";
            ws.Cell(1, 2).Style.Font.FontSize = 9;

            ws.Range(2, 2, 2, 7).Merge();
            ws.Cell(2, 2).Value = _dataDN.TenCoQuanChuyenMon ?? "";
            ws.Cell(2, 2).Style.Font.FontSize = 9;

            ws.Range(3, 2, 3, 7).Merge();
            ws.Cell(3, 2).Value = _dataDN.DiaChi ?? "";
            ws.Cell(3, 2).Style.Font.FontSize = 9;

            ws.Range(4, 2, 4, 7).Merge();
            ws.Cell(4, 2).Value = _dataDN.DienThoai ?? "";
            ws.Cell(4, 2).Style.Font.FontSize = 9;

            currentRow += 5;

            ws.Range(currentRow, 2, currentRow, 7).Merge();
            ws.Cell(currentRow, 2).Value = "BÁO CÁO TỔNG HỢP CẬN LÂM SÀN";
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Font.FontSize = 14;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 7).Merge();
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
                "STT", "Nơi gửi (phòng khám)", "Bác sĩ", "Yêu cầu", "Đơn giá", "Tổng lượt"
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
                ws.Cell(currentRow, i + 2).Value = ++i;
                ws.Cell(currentRow, i + 2).Style.Font.Bold = true;
                ws.Cell(currentRow, i + 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(currentRow, i + 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            long stt = 1;

            currentRow++;

            foreach (var item in _data)
            {
                int col = 2;

                SetCenterCellNumber(ws.Cell(currentRow, col++), stt++);
                SetMiddle(ws.Cell(currentRow, col++), item.NoiGui ?? "");
                SetMiddle(ws.Cell(currentRow, col++), item.BacSi ?? "");
                SetMiddle(ws.Cell(currentRow, col++), item.YeuCau ?? "");

                SetNumberCell(ws.Cell(currentRow, col++), item.DonGia);
                SetCenterCellNumber(ws.Cell(currentRow, col++), item.SoLuot);
                currentRow++;
            }

            currentRow++;
            var TongLuot = _data.Sum(x => x.SoLuot);
            var TongDonGia = _data.Sum(x => x.DonGia);

            int colFirst = 2;

            ws.Range(currentRow, colFirst, currentRow, colFirst + 3).Merge();

            var cellLabel = ws.Cell(currentRow, colFirst);
            cellLabel.Value = "Tổng Cộng";
            cellLabel.Style.Font.Bold = true;
            cellLabel.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cellLabel.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            // Cột 7: tổng giá trị
            SetNumberCell(ws.Cell(currentRow, colFirst + 4), TongDonGia);
            SetCenterCellNumber(ws.Cell(currentRow, colFirst + 5), TongLuot);

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

            void SetNumberCell(IXLCell cell, double? number)
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
