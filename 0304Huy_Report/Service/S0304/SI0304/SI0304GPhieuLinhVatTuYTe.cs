using M0304G.Models.PhieuLinhVatTuYTe;
using Microsoft.AspNetCore.Mvc.RazorPages;
using static System.Runtime.InteropServices.JavaScript.JSType;
public interface I0304GPhieuLinhVatTuYTeService
{
    Task<M0304GPhieuLinhVatTuYTeResponse> GetPhieuLinhVatTuYTe(string ngayBatDau, string ngayKetThuc, long idCN,
        long? idKhoHang = null, int page = 1, int pageSize = 20);
    Task<byte[]> ExportPhieuLinhVatTuYTePdfAsync(ExportRequest request, ISession session);
    Task<byte[]> ExportPhieuLinhVatTuYTeExcelAsync(ExportRequest request, ISession session);
}