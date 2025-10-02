using M0304E.Models.BKBienLaiHoanUng;
using Microsoft.AspNetCore.Mvc.RazorPages;
using static System.Runtime.InteropServices.JavaScript.JSType;
public interface I0304EBKBienLaiHoanUngService
{
    Task<M0304EBKBienLaiHoanUngResponse> GetBKBienLaiHoanUng(string ngayBatDau, string ngayKetThuc, long idCN, long? idNhanVien = null, int page = 1, int pageSize = 20);
    Task<byte[]> ExportBKBienLaiHoanUngPdfAsync(ExportRequest request, ISession session);
    Task<byte[]> ExportBKBienLaiHoanUngExcelAsync(ExportRequest request, ISession session);
}