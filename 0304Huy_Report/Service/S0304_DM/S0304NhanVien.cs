using C0304.Db.Models;
using M0304.Models.KhoHang;
using M0304NhanVien.Models;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;

namespace S0304NhanVien.Services
{
    public class S0304NhanVienService : I0304NhanVienService
    {
        private readonly M0304Context _context;

        public S0304NhanVienService(M0304Context context)
        {
            _context = context;
        }

        public async Task<List<M0304NhanVienModel>> GetAllNhanVien()
        {
            var khoHangs = await _context.M0304NhanViens
                .AsNoTracking()
                .OrderBy(x => x.TenNhanVien)
                .ToListAsync();

            return khoHangs;
        }
    }
}
