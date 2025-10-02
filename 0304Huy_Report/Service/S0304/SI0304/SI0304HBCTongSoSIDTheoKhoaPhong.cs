using M0304H.Models.BCTongSoSIDTheoKhoaPhong;
using Microsoft.AspNetCore.Mvc.RazorPages;
using static System.Runtime.InteropServices.JavaScript.JSType;
public interface I0304HBCTongSoSIDTheoKhoaPhongService
{
    Task<M0304HBCTongSoSIDTheoKhoaPhongResponse> GetBCTongSoSIDTheoKhoaPhong(string ngayBatDau, string ngayKetThuc, long idCN, int page = 1, int pageSize = 20);
    Task<byte[]> ExportBCTongSoSIDTheoKhoaPhongPdfAsync(ExportRequest request, ISession session);
    Task<byte[]> ExportBCTongSoSIDTheoKhoaPhongExcelAsync(ExportRequest request, ISession session);
}