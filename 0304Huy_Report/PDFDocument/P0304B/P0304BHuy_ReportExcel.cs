using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using H0304.NumberToText.Helpers;
using M0304.Models.ThongTinDOanhNghiep;
using M0304B.Models.BCHoaDonDienTuDV;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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

            string[] headers = new string[]
            {
            "STT", "Số chứng từ", "Ngày thu", "Giá trị", "Mã bệnh nhân",
            "Tên bệnh nhân", "Năm sinh", "Địa chỉ", "Ngày tạo HDDT", "E_InvoiceNo", "Giá trị HDDT", "Mã truy cứu"
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

            foreach (var item in _data)
            {
                int col = 1;

                ws.Cell(currentRow, col++).Value = stt++; // STT
                ws.Cell(currentRow, col++).Value = item.SoChungTu ?? "";
                ws.Cell(currentRow, col++).Value = item.NgayThu?.ToString("dd/MM/yyyy") ?? "";
                ws.Cell(currentRow, col++).Value = item.GiaTri ?? 0;
                ws.Cell(currentRow, col++).Value = item.MaBenhNhan ?? "";
                ws.Cell(currentRow, col++).Value = item.TenBenhNhan ?? "";
                ws.Cell(currentRow, col++).Value = item.NamSinh?.ToString() ?? "";
                ws.Cell(currentRow, col++).Value = item.DiaChi ?? "";
                ws.Cell(currentRow, col++).Value = item.NgayTaoHDDT?.ToString("dd/MM/yyyy") ?? "";
                ws.Cell(currentRow, col++).Value = item.E_InvoiceNo ?? 0;
                ws.Cell(currentRow, col++).Value = item.GiaTriHDDT ?? 0;
                ws.Cell(currentRow, col++).Value = item.MaTraCuu ?? 0;

                // Style từng dòng
                for (int i = 1; i <= headers.Length; i++)
                {
                    ws.Cell(currentRow, i).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Cell(currentRow, i).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    if (i == 1 || i == 4 || i == 11) // STT, Giá trị, Giá trị HDDT -> canh phải/giữa
                        ws.Cell(currentRow, i).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                currentRow++;
            }

            var tongGiaTri = _data.Sum(x => x.GiaTri ?? 0);
            var tongGiaTriHDDT = _data.Sum(x => x.GiaTriHDDT ?? 0);

            int colIndex = 4;
            ws.Cell(currentRow, colIndex++).Value = tongGiaTri;

            colIndex = 11;
            ws.Cell(currentRow, colIndex).Value = tongGiaTriHDDT;

            currentRow++;
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
        }
    }
}
