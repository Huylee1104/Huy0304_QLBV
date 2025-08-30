using M0304C.Models.BaoCaoThuDichVu;
using Microsoft.AspNetCore.Mvc.RazorPages;
using static System.Runtime.InteropServices.JavaScript.JSType;
public interface I0304CBaoCaoThuDichVuService
{
    Task<M0304CBaoCaoThuDichVuResponse> GetBCThuDichVu(string ngayBatDau, string ngayKetThuc, long idCN, int page = 1, int pageSize = 20);
    Task<byte[]> ExportBaoCaoThuDichVuPdfAsync(ExportRequest request, ISession session);
    Task<byte[]> ExportBaoCaoThuDichVuExcelAsync(ExportRequest request, ISession session);
}