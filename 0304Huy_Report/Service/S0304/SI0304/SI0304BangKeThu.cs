using M0304.Models.BangKeThu;
using Microsoft.AspNetCore.Mvc.RazorPages;
using static System.Runtime.InteropServices.JavaScript.JSType;
public interface I0304BangKeThuService
{
    Task<M0304BangKeThuResponse> GetBangKeThu(string ngayBatDau, string ngayKetThuc, long idCN, long? idHTTT = null, 
        long? idNhanVien = null, int page = 1, int pageSize = 20);
    Task<byte[]> ExportBaoCaoGoiKhamPdfAsync(ExportRequest request, ISession session);
    Task<byte[]> ExportBaoCaoGoiKhamExcelAsync(ExportRequest request, ISession session);
}