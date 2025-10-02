using M0304D.Models.BKBienLaiTamUng;
using Microsoft.AspNetCore.Mvc.RazorPages;
using static System.Runtime.InteropServices.JavaScript.JSType;
public interface I0304DBKBienLaiTamUngService
{
    Task<M0304DBKBienLaiTamUngResponse> GetBKBienLaiTamUng(string ngayBatDau, string ngayKetThuc, long idCN, long? IdNhanVien = null, int page = 1, int pageSize = 20);
    Task<byte[]> ExportBKBienLaiTamUngPdfAsync(ExportRequest request, ISession session);
    Task<byte[]> ExportBKBienLaiTamUngExcelAsync(ExportRequest request, ISession session);
}