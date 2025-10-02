using M0304F.Models.BKTinhHinhTraDuocNCC;
using Microsoft.AspNetCore.Mvc.RazorPages;
using static System.Runtime.InteropServices.JavaScript.JSType;
public interface I0304FBKTinhHinhTraDuocNCCService
{
    Task<M0304FBKTinhHinhTraDuocNCCResponse> GetBKTinhHinhTraDuocNCC(string ngayBatDau, string ngayKetThuc, long idCN,
        long? idKhoHang = null, int page = 1, int pageSize = 20);
    Task<byte[]> ExportBKTinhHinhTraDuocNCCPdfAsync(ExportRequest request, ISession session);
    Task<byte[]> ExportBKTinhHinhTraDuocNCCExcelAsync(ExportRequest request, ISession session);
}