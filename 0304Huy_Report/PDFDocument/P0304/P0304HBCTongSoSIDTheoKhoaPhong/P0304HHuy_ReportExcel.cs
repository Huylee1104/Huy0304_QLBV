using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using M0304.Models.ThongTinDoanhNghiep;
using M0304H.Models.BCTongSoSIDTheoKhoaPhong;
using System;
using System.Text;

public class P0304HExcelReportTemplate
{
    private readonly List<M0304HBCTongSoSIDTheoKhoaPhong> _data;
    private readonly M0304ThongTinDoanhNghiep _dataDN;
    private string _ngayBatDau;
    private string _ngayKetThuc;
    private readonly string _logoPath;

    public P0304HExcelReportTemplate(
        List<M0304HBCTongSoSIDTheoKhoaPhong> data,
        string ngayBatDau,
        string ngayKetThuc,
        M0304ThongTinDoanhNghiep dataDN,
        string logoPath = null
    )
    {
        _data = data ?? new List<M0304HBCTongSoSIDTheoKhoaPhong>();
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

            ws.Range(1, 3, 1, 12).Merge();
            ws.Cell(1, 3).Value = _dataDN.TenCSKCB ?? "";
            ws.Cell(1, 2).Style.Font.FontSize = 9;

            ws.Range(2, 3, 2, 12).Merge();
            ws.Cell(2, 3).Value = _dataDN.TenCoQuanChuyenMon ?? "";
            ws.Cell(2, 3).Style.Font.FontSize = 9;

            ws.Range(3, 3, 3, 12).Merge();
            ws.Cell(3, 3).Value = _dataDN.DiaChi ?? "";
            ws.Cell(3, 3).Style.Font.FontSize = 9;

            ws.Range(4, 3, 4, 12).Merge();
            ws.Cell(4, 3).Value = _dataDN.DienThoai ?? "";
            ws.Cell(4, 3).Style.Font.FontSize = 9;

            currentRow += 5;

            ws.Range(currentRow, 2, currentRow, 12).Merge();
            ws.Cell(currentRow, 2).Value = "BẢNG TỔNG KẾT XÉT NGHIỆM BỆNH NHÂN";
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Font.FontSize = 14;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 12).Merge();
            DateTime dtStart, dtEnd;
            if (DateTime.TryParse(_ngayBatDau, out dtStart) && DateTime.TryParse(_ngayKetThuc, out dtEnd))
            {
                ws.Cell(currentRow, 2).Value = $"Từ ngày {dtStart:dd-MM-yyyy HH:mm:ss} đến ngày {dtEnd:dd-MM-yyyy HH:mm:ss}";
            }
            else
            {
                ws.Cell(currentRow, 2).Value = $"Từ ngày {_ngayBatDau} đến ngày {_ngayKetThuc}";
            }
            ws.Cell(currentRow, 2).Style.Font.FontSize = 10;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;
            currentRow++;

            string[] headers = new string[]
{
                "STT",
                "Tên Khoa Phòng",
                "Viện phí",
                "Bảo hiểm 100% \n(QL01)",
                "Bảo hiểm 100% \n(QL02)",
                "Bảo hiểm 95% \n(QL03)",
                "Bảo hiểm 80% \n(QL04)",
                "Bảo hiểm 100% \n(QL05)",
                "Dịch vụ",
                "Khám chuyên gia",
                "Tổng"
            };
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(currentRow, i + 2).Value = headers[i];

                // Bật wrap text để Excel cho xuống dòng
                var cell = ws.Cell(currentRow, i + 2);
                cell.Style.Font.Bold = true;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                cell.Style.Alignment.WrapText = true;
            }

            long stt = 1;

            currentRow++;

            foreach (var item in _data)
            {
                int col = 2;

                SetCenterCellNumber(ws.Cell(currentRow, col++), stt++);
                ws.Cell(currentRow, col++).Value = item.TenKhoaPhong ?? "";
                SetNumberCell(ws.Cell(currentRow, col++), item.VienPhi);
                SetNumberCell(ws.Cell(currentRow, col++), item.QL01);
                SetNumberCell(ws.Cell(currentRow, col++), item.QL02);
                SetNumberCell(ws.Cell(currentRow, col++), item.QL03);
                SetNumberCell(ws.Cell(currentRow, col++), item.QL04);
                SetNumberCell(ws.Cell(currentRow, col++), item.QL05);
                SetNumberCell(ws.Cell(currentRow, col++), item.DichVu);
                SetNumberCell(ws.Cell(currentRow, col++), item.KhamChuyenGia);
                SetNumberCell(ws.Cell(currentRow, col++), item.Tong);
                currentRow++;
            }

            var tongVienPhi = _data.Sum(x => x.VienPhi);
            var tongQL01 = _data.Sum(x => x.QL01);
            var tongQL02 = _data.Sum(x => x.QL02);
            var tongQL03 = _data.Sum(x => x.QL03);
            var tongQL04 = _data.Sum(x => x.QL04);
            var tongQL05 = _data.Sum(x => x.QL05);
            var tongDichVu = _data.Sum(x => x.DichVu);
            var tongKhamChuyenGia = _data.Sum(x => x.KhamChuyenGia);
            var tongTatCa = _data.Sum(x => x.Tong);

            int colFirst = 2;
            int colLast = 3;

            ws.Range(currentRow, colFirst, currentRow, colLast).Merge();

            var cellLabel = ws.Cell(currentRow, colFirst);
            cellLabel.Value = "Tổng Cộng"; 
            cellLabel.Style.Font.Bold = true;
            cellLabel.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cellLabel.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            ws.Cell(currentRow, 4).Value = tongVienPhi;
            ws.Cell(currentRow, 5).Value = tongQL01;
            ws.Cell(currentRow, 6).Value = tongQL02;
            ws.Cell(currentRow, 7).Value = tongQL03;
            ws.Cell(currentRow, 8).Value = tongQL04;
            ws.Cell(currentRow, 9).Value = tongQL05;
            ws.Cell(currentRow, 10).Value = tongDichVu;
            ws.Cell(currentRow, 11).Value = tongKhamChuyenGia;
            ws.Cell(currentRow, 12).Value = tongTatCa;

            // In đậm các ô số liệu
            for (int col = 4; col <= 12; col++)
            {
                var cell = ws.Cell(currentRow, col);
                cell.Style.Font.Bold = true;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            }

            ws.Range(9, 2, currentRow, 12).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(9, 2, currentRow, 12).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

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
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }

            void SetCenterCellNumber(IXLCell cell, long? value)
            {
                cell.Value = value ?? 0;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center; 
            }

        }
    }
}
