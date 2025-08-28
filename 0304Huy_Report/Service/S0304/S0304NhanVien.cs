using M0304NhanVien.Models;
using Newtonsoft.Json;

namespace S0304NhanVien.Services
{
    public class S0304NhanVienService : I0304NhanVienService
    {
        private readonly string _filePath;

        public S0304NhanVienService(IWebHostEnvironment env)
        {
            _filePath = Path.Combine(env.WebRootPath, "dist", "data", "json", "Dm_NhanVien.json");
        }

        public string GetFilePath()
        {
            return _filePath;
        }

        public async Task<List<M0304NhanVienModel>> GetAllNhanVien()
        {
            if (!File.Exists(_filePath))
                return new List<M0304NhanVienModel>();

            var json = File.ReadAllText(_filePath);
            var htttList = JsonConvert.DeserializeObject<List<M0304NhanVienModel>>(json);

            return htttList.OrderBy(httt => httt.Ten).ToList();
        }
    }
}
