using M0304.Models.KhoHang;
using C0304.Db.Models;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using System.Data;

namespace S0304KhoHang.Services
{
    public class S0304KhohangService : I0304KhoHangService
    {
        private readonly M0304Context _context;

        public S0304KhohangService(M0304Context context)
        {
            _context = context;
        }

        public async Task<List<M0304KhoHang>> GetKhoHang()
        {
            var khoHangs = await _context.M0304KhoHangs
                .AsNoTracking()
                .OrderBy(x => x.TenKhoHang)
                .ToListAsync();

            return khoHangs;
        }

    }
}
