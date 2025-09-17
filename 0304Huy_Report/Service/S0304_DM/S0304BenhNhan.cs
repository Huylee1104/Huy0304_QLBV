using C0304.Db.Models;
using M0304BenhNhan.Models;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;

namespace S0304BenhNhan.Services
{
    public class S0304BenhNhanService : I0304BenhNhanService
    {
        private readonly M0304Context _context;

        public S0304BenhNhanService(M0304Context context)
        {
            _context = context;
        }

        public async Task<List<M0304BenhNhanModel>> GetAllBenhNhan()
        {
            var benhNhans = await _context.M0304BenhNhans
                .AsNoTracking()
                .OrderBy(x => x.TenBN)
                .ToListAsync();

            return benhNhans;
        }
    }
}
