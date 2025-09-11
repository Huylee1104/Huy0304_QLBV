using ClosedXML.Excel;
using H0304.NumberToText.Helpers;
using M0304.Models.ThongTinDoanhNghiep;
using M0304G.Models.PhieuLinhVatTuYTe;
using M0304NhanVien.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

public class P0304GExcelReportTemplate
{
    private readonly List<M0304GPhieuLinhVatTuYTe> _data;
    private readonly M0304ThongTinDoanhNghiep _dataDN;
    private readonly string _tenKho;
    private string _ngayBatDau;
    private string _ngayKetThuc;
    private readonly string _logoPath;

    public P0304GExcelReportTemplate(
        List<M0304GPhieuLinhVatTuYTe> data,
        M0304ThongTinDoanhNghiep dataDN,
        string tenKho,
        string ngayBatDau,
        string ngayKetThuc,
        string logoPath = null
    )
    {
        _data = data ?? new List<M0304GPhieuLinhVatTuYTe>();
        _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
        _tenKho = tenKho;
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

            ws.Range(1, 3, 1, 8).Merge();
            ws.Cell(1, 3).Value = _dataDN.TenCSKCB ?? "";
            ws.Cell(1, 3).Style.Font.FontSize = 9;

            ws.Range(2, 3, 2, 8).Merge();
            ws.Cell(2, 3).Value = _dataDN.TenCoQuanChuyenMon ?? "";
            ws.Cell(2, 3).Style.Font.FontSize = 9;

            ws.Range(3, 3, 3, 8).Merge();
            ws.Cell(3, 3).Value = _dataDN.DiaChi ?? "";
            ws.Cell(3, 3).Style.Font.FontSize = 9;

            ws.Range(4, 3, 4, 8).Merge();
            ws.Cell(4, 3).Value = _dataDN.DienThoai ?? "";
            ws.Cell(4, 3).Style.Font.FontSize = 9;

            currentRow += 5;

            ws.Range(currentRow, 2, currentRow, 8).Merge();
            ws.Cell(currentRow, 2).Value = "PHIẾU LĨNH VẬT TƯ";
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Font.FontSize = 14;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 8).Merge();
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
            var cell = ws.Cell(currentRow, 2);
            ws.Range(currentRow, 2, currentRow, 8).Merge();
            var rt = cell.GetRichText();
            rt.AddText("Kho phát: ");
            rt.AddText("Kho chẵn vật tư").SetBold();
            cell.Style.Font.FontSize = 9;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

            currentRow++;
            ws.Range(currentRow, 2, currentRow, 8).Merge();
            ws.Cell(currentRow, 2).Value = "Diễn giải: nhu cầu sử dụng cho bệnh nhân.";
            ws.Cell(currentRow, 2).Style.Font.FontSize = 9;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            
            currentRow ++;
            ws.Cell(currentRow, 2).Value = "S T T";
            ws.Cell(currentRow, 3).Value = "Mã";
            ws.Cell(currentRow, 4).Value = "Tên thuốc/VTYT/Hóa chất";
            ws.Cell(currentRow, 5).Value = "Đơn vị tính";
            ws.Cell(currentRow, 6).Value = "Số lượng";
            ws.Cell(currentRow, 8).Value = "Ghi chú";
            ws.Range(currentRow, 6, currentRow, 7).Merge();

            ws.Range(currentRow, 2, currentRow, 8)
                .Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                .Font.SetBold();

            currentRow++;
            ws.Cell(currentRow, 6).Value = "Yêu cầu";
            ws.Cell(currentRow, 7).Value = "Phát";

            ws.Range(currentRow, 2, currentRow, 8)
                .Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                .Font.SetBold();

            int stt = 1;
            foreach (var item in _data)
            {
                currentRow++;
                var tableRange = ws.Range(10, 2, currentRow, 8);
                tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(currentRow, 2).Value = stt++;AlignCellCenter(ws.Cell(currentRow, 2));
                ws.Cell(currentRow, 3).Value = item.MaVatTu; AlignCellCenter(ws.Cell(currentRow, 3));
                var tenVatTu = item.TenVatTu ?? "";
                int maxLength = 60;
                var sb = new StringBuilder();
                int lastBreak = 0;

                while (lastBreak < tenVatTu.Length)
                {
                    if (lastBreak + maxLength >= tenVatTu.Length)
                    {
                        sb.Append(tenVatTu.Substring(lastBreak));
                        break;
                    }

                    int breakIndex = tenVatTu.LastIndexOf(' ', lastBreak + maxLength, maxLength);
                    if (breakIndex <= lastBreak)
                        breakIndex = lastBreak + maxLength;

                    sb.Append(tenVatTu.Substring(lastBreak, breakIndex - lastBreak));
                    sb.Append("\n");
                    lastBreak = breakIndex + 1;
                }

                var cellten = ws.Cell(currentRow, 4);
                cellten.Value = sb.ToString();
                cellten.Style.Alignment.WrapText = true;
                cellten.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                ws.Cell(currentRow, 5).Value = item.DonViTinh; AlignCellCenter(ws.Cell(currentRow, 5));
                ws.Cell(currentRow, 6).Value = item.SoLuong ?? 0; AlignCellRight(ws.Cell(currentRow, 6));
                ws.Cell(currentRow, 7).Value = "";
                ws.Cell(currentRow, 8).Value = "";
            }

            currentRow++;
            ws.Range(currentRow, 2, currentRow, 5).Merge();
            ws.Cell(currentRow, 2).Value = $"Cộng khoản: {_data.Count} khoản";
            ws.Cell(currentRow, 2).Style.Font.SetBold();
            ws.Cell(currentRow, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);

            ws.Range(currentRow, 6, currentRow, 8).Merge();
            ws.Cell(currentRow, 6).Value = $"Ngày {DateTime.Now:dd} Tháng {DateTime.Now:MM} Năm {DateTime.Now:yyyy}";
            ws.Cell(currentRow, 6).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);

            currentRow += 2;
            ws.Range(currentRow, 2, currentRow, 3).Merge().Value = "Người lập bảng";
            ws.Range(currentRow, 4, currentRow, 4).Merge().Value = "Trưởng khoa Dược/VTYT\nngười được uỷ quyền";
            ws.Range(currentRow, 5, currentRow, 6).Merge().Value = "Trưởng khoa/phòng";
            ws.Range(currentRow, 7, currentRow, 7).Merge().Value = "Người giao";
            ws.Range(currentRow, 8, currentRow, 8).Merge().Value = "Người nhận";

            // Định dạng chung
            ws.Range(currentRow, 2, currentRow, 8).Style.Alignment
                .SetHorizontal(XLAlignmentHorizontalValues.Center)
                .Font.SetBold();

            currentRow++;

            // Dòng thứ 2 (ký tên)
            ws.Range(currentRow, 2, currentRow, 3).Merge().Value = "(Ký, ghi rõ họ tên)";
            ws.Range(currentRow, 4, currentRow, 4).Merge().Value = "(Ký, ghi rõ họ tên)";
            ws.Range(currentRow, 5, currentRow, 6).Merge().Value = "(Ký, ghi rõ họ tên)";
            ws.Range(currentRow, 7, currentRow, 7).Merge().Value = "(Ký, ghi rõ họ tên)";
            ws.Range(currentRow, 8, currentRow, 8).Merge().Value = "(Ký, ghi rõ họ tên)";
            ws.Range(currentRow, 2, currentRow, 8).Style.Alignment
                .SetHorizontal(XLAlignmentHorizontalValues.Center);

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

            void CellBold(IXLCell cell)
            {
                cell.Style.Font.Bold = true;
            }
        }
    }
}
