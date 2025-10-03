using ClosedXML.Excel;
using H0304.NumberToText.Helpers;
using M0304NhanVien.Models;
using M0304M.Models.BaoCaoHangHoa;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using M0304.Models.ThongTinDoanhNghiep;

public class P0304ExcelReportNhapTemplate
{
    private readonly List<M0304MHangNhap> _data;
    private readonly M0304ThongTinDoanhNghiep _dataDN;
    private int _nam;
    private readonly string _logoPath;
    public P0304ExcelReportNhapTemplate(
        List<M0304MHangNhap> data,
        int nam,
        M0304ThongTinDoanhNghiep dataDN,
        string logoPath
    )
    {
        _data = data ?? new List<M0304MHangNhap>();
        _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
        _nam = nam;
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

            ws.Range(1, 3, 1, 17).Merge();
            ws.Cell(1, 3).Value = _dataDN.TenCoQuanChuyenMon ?? "";
            ws.Cell(1, 3).Style.Font.FontSize = 9;
            ws.Cell(1, 3).Style.Font.Bold = true;

            ws.Range(2, 3, 2, 17).Merge();
            ws.Cell(2, 3).Value = _dataDN.TenCSKCB ?? "";
            ws.Cell(2, 3).Style.Font.FontSize = 9;
            ws.Cell(2, 3).Style.Font.Bold = true;

            currentRow += 3;

            ws.Range(currentRow, 2, currentRow, 17).Merge();
            ws.Cell(currentRow, 2).Value = $"BÁO CÁO SỐ LƯỢNG HÀNG NHẬP {_nam}";
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Font.FontSize = 14;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            string[] headers = new string[]
            {
            "STT", "Mã thuốc", "Tên thuốc", "Tháng 1", "Tháng 2", "Tháng 3", "Tháng 4", "Tháng 5", "Tháng 6",
            "Tháng 7", "Tháng 8", "Tháng 9", "Tháng 10", "Tháng 11", "Tháng 12", "Tổng cộng"
            };
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(currentRow, i + 2).Value = headers[i];
                ws.Cell(currentRow, i + 2).Style.Font.Bold = true;
                ws.Cell(currentRow, i + 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(currentRow, i + 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            int stt = 1;

            var _dsNhomHang = _data
                .GroupBy(x => new { x.IDNhomHang, x.TenNhomHang })
                .Select(g => new { IDNH = g.Key.IDNhomHang, TenNH = g.Key.TenNhomHang })
                .ToList();

            foreach (var nhomHang in _dsNhomHang)
            {
                var data = _data.Where(d => d.IDNhomHang == nhomHang.IDNH).ToList();
                if (!data.Any()) continue;

                int tongThang1 = (int)data.Sum(x => x.Thang1 ?? 0);
                int tongThang2 = (int)data.Sum(x => x.Thang2 ?? 0);
                int tongThang3 = (int)data.Sum(x => x.Thang3 ?? 0);
                int tongThang4 = (int)data.Sum(x => x.Thang4 ?? 0);
                int tongThang5 = (int)data.Sum(x => x.Thang5 ?? 0);
                int tongThang6 = (int)data.Sum(x => x.Thang6 ?? 0);
                int tongThang7 = (int)data.Sum(x => x.Thang7 ?? 0);
                int tongThang8 = (int)data.Sum(x => x.Thang8 ?? 0);
                int tongThang9 = (int)data.Sum(x => x.Thang9 ?? 0);
                int tongThang10 = (int)data.Sum(x => x.Thang10 ?? 0);
                int tongThang11 = (int)data.Sum(x => x.Thang11 ?? 0);
                int tongThang12 = (int)data.Sum(x => x.Thang12 ?? 0);
                int tongCong = (int)data.Sum(x => x.TongCong ?? 0);

                currentRow++;
                ws.Range(currentRow, 2, currentRow, 4).Merge().Value = nhomHang.TenNH;
                ws.Range(currentRow, 2, currentRow, 4).Style
                    .Font.SetBold()
                    .Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                    .Border.OutsideBorder = XLBorderStyleValues.Thin;

                ws.Cell(currentRow, 5).Value = tongThang1;
                ws.Cell(currentRow, 6).Value = tongThang2;
                ws.Cell(currentRow, 7).Value = tongThang3;
                ws.Cell(currentRow, 8).Value = tongThang4;
                ws.Cell(currentRow, 9).Value = tongThang5;
                ws.Cell(currentRow, 10).Value = tongThang6;
                ws.Cell(currentRow, 11).Value = tongThang7;
                ws.Cell(currentRow, 12).Value = tongThang8;
                ws.Cell(currentRow, 13).Value = tongThang9;
                ws.Cell(currentRow, 14).Value = tongThang10;
                ws.Cell(currentRow, 15).Value = tongThang11;
                ws.Cell(currentRow, 16).Value = tongThang12;
                ws.Cell(currentRow, 17).Value = tongCong;

                ws.Range(currentRow, 5, currentRow, 17).Style.NumberFormat.SetFormat("#,##0");
                ws.Range(currentRow, 5, currentRow, 17).Style.Font.SetBold();
                // Thêm viền cho dòng tổng
                ws.Range(currentRow, 2, currentRow, 17).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Range(currentRow, 2, currentRow, 17).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                // Dữ liệu từng thuốc
                foreach (var item in data)
                {
                    currentRow++;
                    ws.Cell(currentRow, 2).Value = stt++;
                    ws.Cell(currentRow, 3).Value = item.MaThuoc ?? "";
                    ws.Cell(currentRow, 4).Value = item.TenThuoc ?? "";
                    ws.Cell(currentRow, 5).Value = item.Thang1;
                    ws.Cell(currentRow, 6).Value = item.Thang2;
                    ws.Cell(currentRow, 7).Value = item.Thang3;
                    ws.Cell(currentRow, 8).Value = item.Thang4;
                    ws.Cell(currentRow, 9).Value = item.Thang5;
                    ws.Cell(currentRow, 10).Value = item.Thang6;
                    ws.Cell(currentRow, 11).Value = item.Thang7;
                    ws.Cell(currentRow, 12).Value = item.Thang8;
                    ws.Cell(currentRow, 13).Value = item.Thang9;
                    ws.Cell(currentRow, 14).Value = item.Thang10;
                    ws.Cell(currentRow, 15).Value = item.Thang11;
                    ws.Cell(currentRow, 16).Value = item.Thang12;
                    ws.Cell(currentRow, 17).Value = item.TongCong;

                    ws.Range(currentRow, 5, currentRow, 17).Style.NumberFormat.SetFormat("#,##0");
                    ws.Range(currentRow, 2, currentRow, 17).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Range(currentRow, 2, currentRow, 17).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }
            }

            ws.Columns().AdjustToContents();
            ws.Column(1).Width = 2;
            ws.Column(2).Width = 12;


            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms);
                return ms.ToArray();
            }
        }
    }
}

public class P0304ExcelReportXuatTemplate
{
    private readonly List<M0304MHangXuat> _data;
    private readonly M0304ThongTinDoanhNghiep _dataDN;
    private int _nam;
    private readonly string _logoPath;

    public P0304ExcelReportXuatTemplate(
        List<M0304MHangXuat> data,
        int nam,
        M0304ThongTinDoanhNghiep dataDN,
        string logoPath
    )
    {
        _data = data ?? new List<M0304MHangXuat>();
        _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
        _nam = nam;
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

            ws.Range(1, 3, 1, 17).Merge();
            ws.Cell(1, 3).Value = _dataDN.TenCoQuanChuyenMon ?? "";
            ws.Cell(1, 3).Style.Font.FontSize = 9;
            ws.Cell(1, 3).Style.Font.Bold = true;

            ws.Range(2, 3, 2, 17).Merge();
            ws.Cell(2, 3).Value = _dataDN.TenCSKCB ?? "";
            ws.Cell(2, 3).Style.Font.FontSize = 9;
            ws.Cell(2, 3).Style.Font.Bold = true;

            currentRow += 3;

            ws.Range(currentRow, 2, currentRow, 17).Merge();
            ws.Cell(currentRow, 2).Value = $"BÁO CÁO SỐ LƯỢNG HÀNG XUẤT {_nam}";
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Font.FontSize = 14;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            string[] headers = new string[]
            {
            "STT","Mã thuốc", "Tên thuốc", "Tháng 1", "Tháng 2", "Tháng 3", "Tháng 4", "Tháng 5", "Tháng 6",
            "Tháng 7", "Tháng 8", "Tháng 9", "Tháng 10", "Tháng 11", "Tháng 12", "Tổng cộng"
            };
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(currentRow, i + 2).Value = headers[i];
                ws.Cell(currentRow, i + 2).Style.Font.Bold = true;
                ws.Cell(currentRow, i + 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(currentRow, i + 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            int stt = 1;

            var _dsNhomHang = _data
                .GroupBy(x => new { x.IDNhomHang, x.TenNhomHang })
                .Select(g => new { IDNH = g.Key.IDNhomHang, TenNH = g.Key.TenNhomHang })
                .ToList();

            foreach (var nhomHang in _dsNhomHang)
            {
                var data = _data.Where(d => d.IDNhomHang == nhomHang.IDNH).ToList();
                if (!data.Any()) continue;

                int tongThang1 = (int)data.Sum(x => x.Thang1 ?? 0);
                int tongThang2 = (int)data.Sum(x => x.Thang2 ?? 0);
                int tongThang3 = (int)data.Sum(x => x.Thang3 ?? 0);
                int tongThang4 = (int)data.Sum(x => x.Thang4 ?? 0);
                int tongThang5 = (int)data.Sum(x => x.Thang5 ?? 0);
                int tongThang6 = (int)data.Sum(x => x.Thang6 ?? 0);
                int tongThang7 = (int)data.Sum(x => x.Thang7 ?? 0);
                int tongThang8 = (int)data.Sum(x => x.Thang8 ?? 0);
                int tongThang9 = (int)data.Sum(x => x.Thang9 ?? 0);
                int tongThang10 = (int)data.Sum(x => x.Thang10 ?? 0);
                int tongThang11 = (int)data.Sum(x => x.Thang11 ?? 0);
                int tongThang12 = (int)data.Sum(x => x.Thang12 ?? 0);
                int tongCong = (int)data.Sum(x => x.TongCong ?? 0);

                currentRow++;
                ws.Range(currentRow, 2, currentRow, 4).Merge().Value = nhomHang.TenNH;
                ws.Range(currentRow, 2, currentRow, 4).Style
                    .Font.SetBold()
                    .Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                    .Border.OutsideBorder = XLBorderStyleValues.Thin;

                ws.Cell(currentRow, 5).Value = tongThang1;
                ws.Cell(currentRow, 6).Value = tongThang2;
                ws.Cell(currentRow, 7).Value = tongThang3;
                ws.Cell(currentRow, 8).Value = tongThang4;
                ws.Cell(currentRow, 9).Value = tongThang5;
                ws.Cell(currentRow, 10).Value = tongThang6;
                ws.Cell(currentRow, 11).Value = tongThang7;
                ws.Cell(currentRow, 12).Value = tongThang8;
                ws.Cell(currentRow, 13).Value = tongThang9;
                ws.Cell(currentRow, 14).Value = tongThang10;
                ws.Cell(currentRow, 15).Value = tongThang11;
                ws.Cell(currentRow, 16).Value = tongThang12;
                ws.Cell(currentRow, 17).Value = tongCong;

                ws.Range(currentRow, 5, currentRow, 17).Style.NumberFormat.SetFormat("#,##0");
                ws.Range(currentRow, 5, currentRow, 17).Style.Font.SetBold();
                // Thêm viền cho dòng tổng
                ws.Range(currentRow, 2, currentRow, 17).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Range(currentRow, 2, currentRow, 17).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                // Dữ liệu từng thuốc
                foreach (var item in data)
                {
                    currentRow++;
                    ws.Cell(currentRow, 2).Value = stt++;
                    ws.Cell(currentRow, 3).Value = item.MaThuoc ?? "";
                    ws.Cell(currentRow, 4).Value = item.TenThuoc ?? "";
                    ws.Cell(currentRow, 5).Value = item.Thang1;
                    ws.Cell(currentRow, 6).Value = item.Thang2;
                    ws.Cell(currentRow, 7).Value = item.Thang3;
                    ws.Cell(currentRow, 8).Value = item.Thang4;
                    ws.Cell(currentRow, 9).Value = item.Thang5;
                    ws.Cell(currentRow, 10).Value = item.Thang6;
                    ws.Cell(currentRow, 11).Value = item.Thang7;
                    ws.Cell(currentRow, 12).Value = item.Thang8;
                    ws.Cell(currentRow, 13).Value = item.Thang9;
                    ws.Cell(currentRow, 14).Value = item.Thang10;
                    ws.Cell(currentRow, 15).Value = item.Thang11;
                    ws.Cell(currentRow, 16).Value = item.Thang12;
                    ws.Cell(currentRow, 17).Value = item.TongCong;

                    ws.Range(currentRow, 5, currentRow, 17).Style.NumberFormat.SetFormat("#,##0");
                    ws.Range(currentRow, 2, currentRow, 17).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Range(currentRow, 2, currentRow, 17).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }
            }

            ws.Columns().AdjustToContents();
            ws.Column(1).Width = 2;
            ws.Column(2).Width = 12;


            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms);
                return ms.ToArray();
            }
        }
    }
}
