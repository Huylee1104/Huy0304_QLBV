using ClosedXML.Excel;
using H0304.NumberToText.Helpers;
using M0304NhanVien.Models;
using M0304.Models.ToKhaiChiTietThuPhiLePhi;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using M0304.Models.ThongTinDoanhNghiep;

public class P0304ToKhaiChiTietThuPhiLePhiExcelReportTemplate
{
    private readonly List<M0304ToKhaiChiTietThuPhiLePhi> _data;
    private readonly M0304ThongTinDoanhNghiep _dataDN;
    private string _ngayBatDau;
    private string _ngayKetThuc;
    private readonly string _logoPath;

    public P0304ToKhaiChiTietThuPhiLePhiExcelReportTemplate(
        List<M0304ToKhaiChiTietThuPhiLePhi> data,
        string ngayBatDau,
        string ngayKetThuc,
        M0304ThongTinDoanhNghiep dataDN,
        string logoPath = null
    )
    {
        _data = data ?? new List<M0304ToKhaiChiTietThuPhiLePhi>();
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
                var img = ws.AddPicture(_logoPath)
                    .MoveTo(ws.Cell(1, 2))
                    .Scale(0.2);
                ws.Row(1).AdjustToContents();
            }

            ws.Range(1, 2, 1, 9).Merge();
            ws.Cell(1, 2).Value = _dataDN.TenCoQuanChuyenMon ?? "";
            ws.Cell(1, 2).Style.Font.FontSize = 9;
            ws.Cell(1, 2).Style.Font.Bold = true;

            ws.Range(2, 2, 2, 9).Merge();
            ws.Cell(2, 2).Value = _dataDN.TenCSKCB ?? "";
            ws.Cell(2, 2).Style.Font.FontSize = 9;
            ws.Cell(2, 2).Style.Font.Bold = true;

            currentRow += 3;

            ws.Range(currentRow, 2, currentRow, 9).Merge();
            ws.Cell(currentRow, 2).Value = "TỜ KHAI CHI TIẾT THU PHÍ - LỆ PHÍ";
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
            ws.Cell(currentRow, 2).Style.Font.Italic = true;
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow += 2;

            // ======== Header =========
            string[] headers = new string[]
            {
            "STT", "Đơn vị tính quyển", "Số lần hoặc \nsố BL/HĐ thu", "Số HĐ sử dụng",
            "Tổng số tiền thu", "Hoàn/Hủy trả thu phí cho", "Số tiền thực thu", "Ghi chú"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = ws.Cell(currentRow, i + 2);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Alignment.WrapText = true;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            ws.Row(currentRow).AdjustToContents();

            currentRow++;

            var groupedData = _data
                .GroupBy(x => new { x.IDNhanVien, x.TenNhanVien })
                .Select(nvGroup => new
                {
                    nvGroup.Key.TenNhanVien,
                    LoaiHoaDons = nvGroup
                        .GroupBy(x => x.LoaiHoaDon)
                        .Select(loaiGroup => new
                        {
                            LoaiHoaDon = loaiGroup.Key,
                            ChiTiet = loaiGroup.ToList()
                        })
                        .ToList()
                })
                .ToList();

            foreach (var nv in groupedData)
            {
                ws.Range(currentRow, 2, currentRow, 9).Merge();
                ws.Cell(currentRow, 2).Value = $"NHÂN VIÊN: {nv.TenNhanVien}".ToUpper();
                ws.Cell(currentRow, 2).Style.Font.Bold = true;
                ws.Cell(currentRow, 2).Style.Font.FontSize = 11;
                ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                ws.Range(currentRow, 2, currentRow, 9).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                currentRow++;

                foreach (var loai in nv.LoaiHoaDons)
                {
                    ws.Range(currentRow, 2, currentRow, 9).Merge();
                    ws.Cell(currentRow, 2).Value = $"Loại HĐ: {loai.LoaiHoaDon}";
                    ws.Cell(currentRow, 2).Style.Font.Italic = true;
                    ws.Cell(currentRow, 2).Style.Font.Bold = true;
                    ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    ws.Range(currentRow, 2, currentRow, 9).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    currentRow++;
                    int stt = 1;
                    foreach (var item in loai.ChiTiet)
                    {
                        ws.Cell(currentRow, 2).Value = stt++; AlignCellCenter(ws.Cell(currentRow, 2));
                        ws.Cell(currentRow, 2).Style.Font.Bold = true;
                        ws.Cell(currentRow, 3).Value = item.QuyenSo ?? "";
                        ws.Cell(currentRow, 4).Value = item.SoLan_soBLHDthu ?? ""; AlignCellCenter(ws.Cell(currentRow, 4));
                        ws.Cell(currentRow, 5).Value = item.SoLuongHDSuDung ?? 0; AlignCellCenter(ws.Cell(currentRow, 5));
                        ws.Cell(currentRow, 5).Style.Font.Bold = true;
                        ws.Cell(currentRow, 6).Value = item.TongSoTien ?? 0; AlignCellCenter(ws.Cell(currentRow, 6));
                        ws.Cell(currentRow, 7).Value = item.Huy_Hoan ?? 0; AlignCellCenter(ws.Cell(currentRow, 7));
                        ws.Cell(currentRow, 8).Value = item.SoTienThucThu ?? 0;
                        ws.Cell(currentRow, 9).Value = item.GhiChu ?? "";
                        ws.Cell(currentRow, 9).Style.Font.Italic = true;

                        // Định dạng số
                        ws.Cell(currentRow, 6).Style.NumberFormat.Format = "#,##0";
                        ws.Cell(currentRow, 7).Style.NumberFormat.Format = "#,##0";
                        ws.Cell(currentRow, 8).Style.NumberFormat.Format = "#,##0";

                        // Viền cho từng dòng chi tiết
                        for (int col = 2; col <= 9; col++)
                            ws.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        currentRow++;
                    }
                }
            }

            ws.Cell(currentRow, 9).Style.Alignment.WrapText = true;
            ws.Cell(currentRow, 9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(currentRow, 9).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            var tongTien = _data.Sum(x => x.TongSoTien ?? 0);
            var tongHuyHoan = _data.Sum(x => x.Huy_Hoan ?? 0);
            var tongThucThu = _data.Sum(x => x.SoTienThucThu ?? 0);

            ws.Cell(currentRow, 5).Value = "TỔNG CỘNG:";
            ws.Cell(currentRow, 5).Style.Font.Bold = true;
            ws.Cell(currentRow, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            ws.Cell(currentRow, 6).Value = tongTien;
            ws.Cell(currentRow, 7).Value = tongHuyHoan;
            ws.Cell(currentRow, 8).Value = tongThucThu;

            ws.Cell(currentRow, 6).Style.Font.Bold = true; AlignCellCenter(ws.Cell(currentRow, 6));
            ws.Cell(currentRow, 7).Style.Font.Bold = true; AlignCellCenter(ws.Cell(currentRow, 7));
            ws.Cell(currentRow, 8).Style.Font.Bold = true;

            ws.Cell(currentRow, 6).Style.NumberFormat.Format = "#,##0";
            ws.Cell(currentRow, 7).Style.NumberFormat.Format = "#,##0";
            ws.Cell(currentRow, 8).Style.NumberFormat.Format = "#,##0";

            ws.Range(currentRow, 2, currentRow, 9).Style.Border.OutsideBorder = XLBorderStyleValues.None;

            currentRow++;
            string ngayThangNam = $"Ngày {DateTime.Now:dd} Tháng {DateTime.Now:MM} Năm {DateTime.Now:yyyy}";
            ws.Range(currentRow, 7, currentRow, 9).Merge();
            ws.Cell(currentRow, 7).Value = ngayThangNam;
            ws.Cell(currentRow, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(currentRow, 7).Style.Font.Italic = true;
            ws.Cell(currentRow, 7).Style.Font.Bold = true;
            ws.Cell(currentRow, 7).Style.Font.FontSize = 10;

            currentRow++;
            ws.Range(currentRow, 7, currentRow, 9).Merge();
            ws.Cell(currentRow, 7).Value = "Người lập";
            ws.Cell(currentRow, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(currentRow, 7).Style.Font.Bold = true;
            ws.Cell(currentRow, 7).Style.Font.FontSize = 10;

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

            void AlignCellRight(IXLCell cell)
            {
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }

        }
        
    }
}
