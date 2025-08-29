using M0304.Models.ThongTinDoanhNghiep;
using C0304.Db.Models;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using System.Data;

namespace S0304ThongTinDoanhNghiep.Services
{
    public class S0304ThongTinDoanhNghiepService : I0304ThongTinDoanhNghiep
    {
        private readonly M0304Context _context;

        public S0304ThongTinDoanhNghiepService(M0304Context context)
        {
            _context = context;
        }

        public async Task<M0304ThongTinDoanhNghiep> GetThongTinDoanhNghiep(long idChiNhanh)
        {
            var result = await _context.M0304ThongTinDoanhNghieps
                .FromSqlRaw("EXEC dbo.S0304_ThongTinDoanhNghiep @id", new SqlParameter("@id", idChiNhanh))
                .AsNoTracking()
                .ToListAsync();

            return result.FirstOrDefault();
        }

    }
}
