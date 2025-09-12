using ClosedXML.Excel;
using H0304.NumberToText.Helpers;
using M0304NhanVien.Models;
using M0304E.Models.BKBienLaiHoanUng;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using M0304.Models.ThongTinDoanhNghiep;

public class P0304EExcelReportTemplate
{
    private readonly List<M0304EBKBienLaiHoanUng> _data;
    private readonly M0304ThongTinDoanhNghiep _dataDN;
    private string _ngayBatDau;
    private string _ngayKetThuc;
    private readonly string _logoPath;

    private List<M0304TongTheoNhanVien> _tongTheoNhanVien;
    private List<M0304NhanVienModel> _danhSachNhanVien;

    public P0304EExcelReportTemplate(
        List<M0304EBKBienLaiHoanUng> data,
        string ngayBatDau,
        string ngayKetThuc,
        List<M0304NhanVienModel> danhSachNhanVien,
        List<M0304TongTheoNhanVien> tongTheoNhanVien,
        M0304ThongTinDoanhNghiep dataDN,
        string logoPath = null
    )
    {
        _data = data ?? new List<M0304EBKBienLaiHoanUng>();
        _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
        _ngayBatDau = ngayBatDau;
        _ngayKetThuc = ngayKetThuc;
        _danhSachNhanVien = danhSachNhanVien ?? new List<M0304NhanVienModel>();
        _tongTheoNhanVien = tongTheoNhanVien ?? new List<M0304TongTheoNhanVien>();
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

            ws.Range(1, 3, 1, 13).Merge();
            ws.Cell(1, 3).Value = _dataDN.TenCSKCB ?? "";
            ws.Cell(1, 3).Style.Font.FontSize = 9;

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
            ws.Cell(currentRow, 2).Value = "BẢNG KÊ HOÀN ỨNG THEO SỐ BIÊN LAI";
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

            currentRow += 2;

            string[] headers = new string[]
            {
            "STT", "Ngày thu", "Mã y tế", "Số BA", "Mã đợt",
            "Họ và tên", "Sổ BL hoàn ứng", "Số BL tạm ứng", "Giá trị hoàn ứng", "Hủy", "Hoàn trả", "HTTT"
            };
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(currentRow, i + 2).Value = headers[i];
                ws.Cell(currentRow, i + 2).Style.Font.Bold = true;
                ws.Cell(currentRow, i + 2).Style.Fill.BackgroundColor = XLColor.LightGray;
                ws.Cell(currentRow, i + 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(currentRow, i + 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            currentRow++;

            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(currentRow, i + 2).Value = i + 1;
                ws.Cell(currentRow, i + 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(currentRow, i + 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(currentRow, i + 2).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            }

            decimal tongThuPhi = _data.Sum(x => x.GiaTriHoanUng ?? 0);
            decimal tongHuy = _data.Sum(x => x.Huy ?? 0);
            decimal tongHoanTra = _data.Sum(x => x.HoanTra ?? 0);
            decimal tongChenhLech = tongThuPhi - (tongHuy + tongHoanTra);

            currentRow++;
            foreach (var nv in _danhSachNhanVien)
            {
                int stt = 1;
                var dataNV = _data.Where(d => d.IDNhanVien == nv.ID).ToList();
                if (!dataNV.Any()) continue;
                var tongNV = _tongTheoNhanVien.FirstOrDefault(t => t.IDNhanVien == nv.ID);

                ws.Range(currentRow, 2, currentRow, 9).Merge();
                ws.Cell(currentRow, 2).Value = $"Nhân viên: {nv.TenNhanVien}";
                ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                ws.Cell(currentRow, 2).Style.Font.Bold = true;
                ws.Cell(currentRow, 2).Style.Font.FontSize = 10;

                ws.Cell(currentRow, 10).Value = tongNV?.TongSoTien;
                ws.Cell(currentRow, 11).Value = tongNV?.TongHuy;
                ws.Cell(currentRow, 12).Value = tongNV?.TongHoan;
                ws.Cell(currentRow, 13).Value = "";
                ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";
                ws.Cell(currentRow, 11).Style.NumberFormat.Format = "#,##0";
                ws.Cell(currentRow, 12).Style.NumberFormat.Format = "#,##0";

                ws.Range(currentRow, 2, currentRow, 13).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                for (int col = 10; col <= 13; col++)
                {
                    ws.Cell(currentRow, col).Style.Font.Bold = true;
                    ws.Cell(currentRow, col).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }
                currentRow++;

                foreach (var item in dataNV)
                {
                    ws.Cell(currentRow, 2).Value = stt++; AlignCellCenter(ws.Cell(currentRow, 2));
                    ws.Cell(currentRow, 3).Value = item.NgayThu?.ToString("dd-MM-yyyy hh:mm tt") ?? ""; AlignCellCenter(ws.Cell(currentRow, 3));
                    ws.Cell(currentRow, 4).Value = item.MaYTe ?? ""; AlignCellCenter(ws.Cell(currentRow, 4));
                    ws.Cell(currentRow, 5).Value = item.SoBA ?? ""; AlignCellCenter(ws.Cell(currentRow, 5));
                    ws.Cell(currentRow, 6).Value = item.MaDot ?? ""; AlignCellCenter(ws.Cell(currentRow, 6));
                    ws.Cell(currentRow, 7).Value = item.HoTenBenhNhan ?? "";
                    ws.Cell(currentRow, 8).Value = item.SoBLHoanUng ?? ""; AlignCellCenter(ws.Cell(currentRow, 8));
                    ws.Cell(currentRow, 9).Value = item.SoBLTamUng ?? ""; AlignCellCenter(ws.Cell(currentRow, 9));
                    ws.Cell(currentRow, 10).Value = item.GiaTriHoanUng ?? (decimal?)null;
                    ws.Cell(currentRow, 11).Value = item.Huy ?? (decimal?)null;
                    ws.Cell(currentRow, 12).Value = item.HoanTra ?? (decimal?)null;
                    ws.Cell(currentRow, 13).Value = item.HTTT ?? ""; AlignCellCenter(ws.Cell(currentRow, 13));

                    ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";
                    ws.Cell(currentRow, 11).Style.NumberFormat.Format = "#,##0";
                    ws.Cell(currentRow, 12).Style.NumberFormat.Format = "#,##0";

                    for (int col = 2; col <= headers.Length + 1; col++)
                    {
                        ws.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(currentRow, col).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    }
                    currentRow++;
                }
            }
            var totalRange = ws.Range(currentRow, 2, currentRow, 9);
            totalRange.Merge();
            totalRange.Value = "Tổng cộng";
            totalRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            totalRange.Style.Font.Bold = true;
            totalRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            ws.Cell(currentRow, 10).Value = tongThuPhi;
            ws.Cell(currentRow, 11).Value = tongHuy;
            ws.Cell(currentRow, 12).Value = tongHoanTra;
            ws.Cell(currentRow, 13).Value = "";

            ws.Cell(currentRow, 10).Style.NumberFormat.Format = "#,##0";
            ws.Cell(currentRow, 11).Style.NumberFormat.Format = "#,##0";
            ws.Cell(currentRow, 12).Style.NumberFormat.Format = "#,##0";

            for (int col = 10; col <= 13; col++)
            {
                ws.Cell(currentRow, col).Style.Font.Bold = true;
                ws.Cell(currentRow, col).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }
            currentRow += 2;

            ws.Range(currentRow, 2, currentRow, 9).Merge();
            ws.Cell(currentRow, 2).Value = $"Số tiền hoàn ứng: {tongChenhLech:N0}";
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 9).Merge();
            ws.Cell(currentRow, 2).Value = $"Bằng chữ: {H0304NumberToTextHelper.ConvertSoThanhChu(tongChenhLech)}";
            ws.Cell(currentRow, 2).Style.Font.Italic = true;
            currentRow += 2;

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

            void AlignCellCenter(IXLCell cell)
            {
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }
        }
    }
}
