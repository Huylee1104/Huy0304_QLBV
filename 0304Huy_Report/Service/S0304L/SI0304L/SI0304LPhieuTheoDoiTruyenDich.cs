using M0304L.Models.PhieuTheoDoiTruyenDich;
using Microsoft.AspNetCore.Mvc.RazorPages;
using static System.Runtime.InteropServices.JavaScript.JSType;
public interface I0304LPhieuTheoDoiTruyenDichService
{
    Task<M0304LPhieuTheoDoiTruyenDichResponse> GetPhieuTheoDoiTruyenDich(long idCN, long? idBenhNhan, int page = 1, int pageSize = 20);
    Task<byte[]> ExportGetPhieuTheoDoiTruyenDichPdfAsync(ExportRequest request, ISession session);
}