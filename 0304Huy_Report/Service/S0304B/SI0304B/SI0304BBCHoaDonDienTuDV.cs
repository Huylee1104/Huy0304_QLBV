using M0304B.Models.BCHoaDonDienTuDV;
using Microsoft.AspNetCore.Mvc.RazorPages;
using static System.Runtime.InteropServices.JavaScript.JSType;
public interface I0304BBCHoaDonDienTuDVService
{
    Task<M0304BBCHoaDonDienTuDVResponse> GetBCHoaDonDIenTuDV(string ngayBatDau, string ngayKetThuc, long idCN, int page = 1, int pageSize = 20);
    Task<byte[]> ExportBaoCaoGoiKhamPdfAsync(ExportRequest request, ISession session);
    Task<byte[]> ExportBaoCaoGoiKhamExcelAsync(ExportRequest request, ISession session);
}