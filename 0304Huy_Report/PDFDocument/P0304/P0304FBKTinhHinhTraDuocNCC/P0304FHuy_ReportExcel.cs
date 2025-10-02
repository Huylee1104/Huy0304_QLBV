using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using H0304.NumberToText.Helpers;
using M0304.Models.ThongTinDoanhNghiep;
using M0304F.Models.BKTinhHinhTraDuocNCC;
using M0304NhanVien.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

public class P0304FExcelReportTemplate
{
    private readonly List<M0304FBKTinhHinhTraDuocNCC> _data;
    private readonly M0304ThongTinDoanhNghiep _dataDN;
    private readonly List<CongTyDto> _dsCongTy;
    private readonly string _tenKho;
    private string _ngayBatDau;
    private string _ngayKetThuc;
    private readonly string _logoPath;

    public P0304FExcelReportTemplate(
        List<M0304FBKTinhHinhTraDuocNCC> data,
        M0304ThongTinDoanhNghiep dataDN,
        List<CongTyDto> dsCongTy,
        string tenKho,
        string ngayBatDau,
        string ngayKetThuc,
        string logoPath = null
    )
    {
        _data = data ?? new List<M0304FBKTinhHinhTraDuocNCC>();
        _dataDN = dataDN ?? new M0304ThongTinDoanhNghiep();
        _dsCongTy = dsCongTy ?? new List<CongTyDto>();
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

            ws.Range(1, 3, 1, 16).Merge();
            ws.Cell(1, 3).Value = _dataDN.TenCSKCB ?? "";
            ws.Cell(1, 3).Style.Font.FontSize = 9;

            ws.Range(2, 3, 2, 16).Merge();
            ws.Cell(2, 3).Value = _dataDN.TenCoQuanChuyenMon ?? "";
            ws.Cell(2, 3).Style.Font.FontSize = 9;

            ws.Range(3, 3, 3, 16).Merge();
            ws.Cell(3, 3).Value = _dataDN.DiaChi ?? "";
            ws.Cell(3, 3).Style.Font.FontSize = 9;

            ws.Range(4, 3, 4, 16).Merge();
            ws.Cell(4, 3).Value = _dataDN.DienThoai ?? "";
            ws.Cell(4, 3).Style.Font.FontSize = 9;

            currentRow += 5;

            ws.Range(currentRow, 2, currentRow, 16).Merge();
            ws.Cell(currentRow, 2).Value = "BẢNG KÊ TÌNH HÌNH TRẢ DƯỢC NCC";
            ws.Cell(currentRow, 2).Style.Font.Bold = true;
            ws.Cell(currentRow, 2).Style.Font.FontSize = 14;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 16).Merge();
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
            ws.Range(currentRow, 2, currentRow, 16).Merge();
            ws.Cell(currentRow, 2).Value = "Nguồn dược: Mua";
            ws.Cell(currentRow, 2).Style.Font.FontSize = 9;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            currentRow++;
            ws.Range(currentRow, 2, currentRow, 16).Merge();
            ws.Cell(currentRow, 2).Value = $"Kho trả: {_tenKho}";
            ws.Cell(currentRow, 2).Style.Font.FontSize = 9;
            ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            currentRow += 2;

            string[] headers = new string[]
            {
            "STT", "Ngày hóa đơn", "Số hóa đơn", "Ngày trả", "Phiếu trả", "Công ty", "Mã ID",
            "Tên thuốc hàm lượng", "Quy cách", "Số lô", "SL đóng gói", "SL lẻ", "Đơn giá đóng gói", "Đơn giá lẻ", "Thành tiền"
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

            for (int i = 0; i < headers.Length - 1; i++)
            {
                ws.Cell(currentRow, i + 2).Value = i + 1;
                ws.Cell(currentRow, i + 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(currentRow, i + 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(currentRow, i + 2).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            }
            ws.Cell(currentRow, 16).Value = "15 = 12*14";
            ws.Cell(currentRow, 16).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(currentRow, 16).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Cell(currentRow, 16).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            currentRow++;

            foreach (var cty in _dsCongTy)
            {
                int stt = 1;
                var data = _data.Where(d => d.IDCongTy == cty.ID).ToList();

                ws.Cell(currentRow, 2).Value = cty.Ten;
                ws.Range(currentRow, 2, currentRow, 16).Merge();
                ws.Range(currentRow, 2, currentRow, 16).Style.Font.Bold = true;
                ws.Range(currentRow, 2, currentRow, 16).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Range(currentRow, 2, currentRow, 16).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                currentRow++;

                foreach (var item in data)
                {
                    int col = 2;
                    ws.Cell(currentRow, col++).Value = stt++; AlignCellCenter(ws.Cell(currentRow, col));
                    ws.Cell(currentRow, col).Value = item.NgayHoaDon?.ToString("dd-MM-yyyy"); AlignCellCenter(ws.Cell(currentRow, col++));
                    ws.Cell(currentRow, col++).Value = item.SoHoaDon;
                    ws.Cell(currentRow, col).Value = item.NgayTra?.ToString("dd-MM-yyyy"); AlignCellCenter(ws.Cell(currentRow, col++));
                    ws.Cell(currentRow, col++).Value = item.PhieuTra;
                    ws.Cell(currentRow, col++).Value = item.CongTy;
                    ws.Cell(currentRow, col).Value = item.MaID; AlignCellCenter(ws.Cell(currentRow, col++));

                    var tenVatTu = item.TenThuoc ?? "";
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

                    var cellten = ws.Cell(currentRow, col++);
                    cellten.Value = sb.ToString();
                    cellten.Style.Alignment.WrapText = true;
                    cellten.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    ws.Cell(currentRow, col++).Value = item.QuyCach;
                    ws.Cell(currentRow, col++).Value = item.SoLo;
                    ws.Cell(currentRow, col).Value = item.SLDongGoi; AlignCellRight(ws.Cell(currentRow, col++));
                    ws.Cell(currentRow, col).Value = item.SLLe; AlignCellRight(ws.Cell(currentRow, col++));
                    ws.Cell(currentRow, col).Value = item.DonGiaDongGoi;
                    ws.Cell(currentRow, col).Style.NumberFormat.Format = "#,##0";
                    AlignCellRight(ws.Cell(currentRow, col++));

                    ws.Cell(currentRow, col).Value = item.DonGiaLe;
                    ws.Cell(currentRow, col).Style.NumberFormat.Format = "#,##0";
                    AlignCellRight(ws.Cell(currentRow, col++));

                    ws.Cell(currentRow, col).Value = item.ThanhTien;
                    ws.Cell(currentRow, col).Style.NumberFormat.Format = "#,##0";
                    AlignCellRight(ws.Cell(currentRow, col++));

                    ws.Range(currentRow, 2, currentRow, 16).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Range(currentRow, 2, currentRow, 16).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    currentRow++;
                }
            }

            var tongCong = _data.Sum(x => x.ThanhTien ?? 0);
            var vat = 0;
            var tongTien = tongCong + vat;

            ws.Cell(currentRow, 12).Value = "Tổng cộng:";
            ws.Cell(currentRow, 12).Style.Font.Bold = true;
            ws.Cell(currentRow, 16).Value = tongCong;
            ws.Cell(currentRow, 16).Style.Font.Bold = true;
            ws.Cell(currentRow, 16).Style.NumberFormat.Format = "#,##0";
            currentRow++;

            ws.Cell(currentRow, 12).Value = "Tiền VAT:";
            ws.Cell(currentRow, 12).Style.Font.Bold = true;
            ws.Cell(currentRow, 16).Value = vat;
            ws.Cell(currentRow, 16).Style.Font.Bold = true;
            ws.Cell(currentRow, 16).Style.NumberFormat.Format = "#,##0";
            currentRow++;

            ws.Cell(currentRow, 12).Value = "Tổng tiền:";
            ws.Cell(currentRow, 12).Style.Font.Bold = true;
            ws.Cell(currentRow, 16).Value = tongTien;
            ws.Cell(currentRow, 16).Style.Font.Bold = true;
            ws.Cell(currentRow, 16).Style.NumberFormat.Format = "#,##0";
            currentRow++;

            ws.Range(currentRow, 2, currentRow, 16).Merge();
            ws.Cell(currentRow, 2).Value = $"Ngày {DateTime.Now:dd} Tháng {DateTime.Now:MM} Năm {DateTime.Now:yyyy}";
            ws.Cell(currentRow, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            currentRow += 2;

            ws.Cell(currentRow, 3).Value = "Thủ kho";
            ws.Cell(currentRow, 7).Value = "Kế toán";
            ws.Cell(currentRow, 10).Value = "Người lập";
            ws.Cell(currentRow, 14).Value = "Trưởng khoa";

            ws.Range(currentRow, 2, currentRow, 16).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Range(currentRow, 2, currentRow, 16).Style.Font.Bold = true;

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
