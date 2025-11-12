using ClosedXML.Excel;
using H0304.NumberToText.Helpers;
using M0304NhanVien.Models;
using M0304.Models.BaoCaoTiepNhan;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using M0304.Models.ThongTinDoanhNghiep;

public class P0304BaoCaoTiepNhanExcelReportTemplate
{
    private readonly List<M0304BaoCaoTiepNhan> _data;
    private readonly M0304ThongTinDoanhNghiep _dataDN;
    private string _ngayBatDau;
    private string _ngayKetThuc;
    private readonly string _logoPath;

    public P0304BaoCaoTiepNhanExcelReportTemplate(
        List<M0304BaoCaoTiepNhan> data,
        string ngayBatDau,
        string ngayKetThuc,
        M0304ThongTinDoanhNghiep dataDN,
        string logoPath = null
    )
    {
        _data = data ?? new List<M0304BaoCaoTiepNhan>();
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

            ws.Range(1, 2, 1, 8).Merge();
            ws.Cell(1, 2).Value = _dataDN.TenCSKCB ?? "";
            ws.Cell(1, 2).Style.Font.FontSize = 9;

            ws.Range(2, 2, 2, 8).Merge();
            ws.Cell(2, 2).Value = _dataDN.TenCoQuanChuyenMon ?? "";
            ws.Cell(2, 2).Style.Font.FontSize = 9;

            currentRow += 3;

            // Tiêu đề chính
            ws.Cell(currentRow, 2).Value = "BÁO CÁO THEO KHOA PHÒNG";
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Font.FontSize = 14;
            ws.Range(currentRow, 2, currentRow, 8).Merge();
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow += 2;

            // Header
            string[] headers = { "STT", "Tên phòng bệnh", "Số lượt tiếp nhận", "Nam", "Nữ", "Có bảo hiểm", "Không BHYT" };

            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(currentRow, i + 2).Value = headers[i];
                ws.Cell(currentRow, i + 2).Style.Font.Bold = true;
                ws.Cell(currentRow, i + 2).Style.Fill.BackgroundColor = XLColor.LightGray;
                ws.Cell(currentRow, i + 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(currentRow, i + 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            currentRow++;

            // Nhóm dữ liệu theo IdKhoa, TenKhoa
            var group = _data.GroupBy(x => new { x.IdKhoa, x.TenKhoa }).Select(g => new
            {
                TenKhoa = g.Key.TenKhoa,
                Items = g.ToList()
            }).ToList();

            int stt = 1;

            var tongLuot = _data.Sum(x => x.SoLuotTiepNhan ?? 0);
            var tongNam = _data.Sum(x => x.SoLuongNam ?? 0);
            var tongNu = _data.Sum(x => x.SoLuongNu ?? 0);
            var tongBHYT = _data.Sum(x => x.CoBHYT ?? 0);
            var tongKhongBHYT = _data.Sum(x => x.KhongBHYT ?? 0);

            foreach (var khoa in group)
            {
                // Dòng tiêu đề cho từng khoa
                ws.Range(currentRow, 2, currentRow, 8).Merge();
                ws.Cell(currentRow, 2).Value = khoa.TenKhoa;
                ws.Cell(currentRow, 2).Style.Font.Bold = true;
                ws.Cell(currentRow, 2).Style.Font.FontSize = 11;
                ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                ws.Range(currentRow, 2, currentRow, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                currentRow++;

                int tongLuotKhoa = 0, tongNamKhoa = 0, tongNuKhoa = 0, tongBHYTKhoa = 0, tongKhongBHYTKhoa = 0;

                // Dữ liệu trong từng khoa
                foreach (var item in khoa.Items)
                {
                    ws.Cell(currentRow, 2).Value = stt++;
                    ws.Cell(currentRow, 3).Value = item.TenPhongBenh ?? "";
                    ws.Cell(currentRow, 4).Value = item.SoLuotTiepNhan ?? 0;
                    ws.Cell(currentRow, 5).Value = item.SoLuongNam ?? 0;
                    ws.Cell(currentRow, 6).Value = item.SoLuongNu ?? 0;
                    ws.Cell(currentRow, 7).Value = item.CoBHYT ?? 0;
                    ws.Cell(currentRow, 8).Value = item.KhongBHYT ?? 0;

                    tongLuotKhoa += item.SoLuotTiepNhan ?? 0;
                    tongNamKhoa += item.SoLuongNam ?? 0;
                    tongNuKhoa += item.SoLuongNu ?? 0;
                    tongBHYTKhoa += item.CoBHYT ?? 0;
                    tongKhongBHYTKhoa += item.KhongBHYT ?? 0;

                    for (int col = 2; col <= 8; col++)
                    {
                        ws.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(currentRow, col).Style.Alignment.Horizontal =
                            col == 3 ? XLAlignmentHorizontalValues.Left : XLAlignmentHorizontalValues.Center;
                    }

                    ws.Cell(currentRow, 4).Style.NumberFormat.Format = "#,##0";
                    ws.Cell(currentRow, 5).Style.NumberFormat.Format = "#,##0";
                    ws.Cell(currentRow, 6).Style.NumberFormat.Format = "#,##0";
                    ws.Cell(currentRow, 7).Style.NumberFormat.Format = "#,##0";
                    ws.Cell(currentRow, 8).Style.NumberFormat.Format = "#,##0";

                    currentRow++;
                }

                // Tổng từng khoa
                ws.Range(currentRow, 2, currentRow, 3).Merge();
                ws.Cell(currentRow, 2).Value = "Tổng " + khoa.TenKhoa;
                ws.Cell(currentRow, 2).Style.Font.Bold = true;
                ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                ws.Cell(currentRow, 4).Value = tongLuotKhoa;
                ws.Cell(currentRow, 5).Value = tongNamKhoa;
                ws.Cell(currentRow, 6).Value = tongNuKhoa;
                ws.Cell(currentRow, 7).Value = tongBHYTKhoa;
                ws.Cell(currentRow, 8).Value = tongKhongBHYTKhoa;

                for (int col = 2; col <= 8; col++)
                {
                    ws.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Cell(currentRow, col).Style.Font.Bold = true;
                    ws.Cell(currentRow, col).Style.NumberFormat.Format = "#,##0";
                    ws.Cell(currentRow, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                currentRow++;
            }

            // Tổng cuối bảng
            ws.Range(currentRow, 2, currentRow, 3).Merge();
            ws.Cell(currentRow, 2).Value = "Tổng cộng";
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            ws.Cell(currentRow, 4).Value = tongLuot;
            ws.Cell(currentRow, 5).Value = tongNam;
            ws.Cell(currentRow, 6).Value = tongNu;
            ws.Cell(currentRow, 7).Value = tongBHYT;
            ws.Cell(currentRow, 8).Value = tongKhongBHYT;

            for (int col = 2; col <= 8; col++)
            {
                ws.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(currentRow, col).Style.Font.Bold = true;
                ws.Cell(currentRow, col).Style.NumberFormat.Format = "#,##0";
                ws.Cell(currentRow, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            ws.Columns().AdjustToContents();
            ws.Column(1).Width = 2;
            ws.Column(2).Width = 6; // STT
            ws.Column(3).Width = 25; // Tên phòng bệnh
            ws.Column(4).Width = 14;
            ws.Column(5).Width = 10;
            ws.Column(6).Width = 10;
            ws.Column(7).Width = 14;
            ws.Column(8).Width = 14;

            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms);
                return ms.ToArray();
            }
        }
    }
}
