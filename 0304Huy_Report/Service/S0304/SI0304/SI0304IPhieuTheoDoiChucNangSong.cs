using M0304I.Models.PhieuTheoDoiChucNangSong;
using Microsoft.AspNetCore.Mvc.RazorPages;
using static System.Runtime.InteropServices.JavaScript.JSType;
public interface I0304CPhieuTheoDoiChucNangSongService
{
    Task<M0304IPhieuTheoDoiChucNangSongResponse> GetPhieuTheoDoiChucNangSong(long idCN, long? idVaoVien, int page = 1, int pageSize = 20);
    Task<byte[]> ExportGetPhieuTheoDoiChucNangSongPdfAsync(ExportRequest request, ISession session);
}