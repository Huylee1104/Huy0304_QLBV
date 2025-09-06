using M0304HTTT.Models;
using Newtonsoft.Json;

namespace S0304HTTT.Services
{
    public class S0304HTTTService : I0304HTTTService
    {
        private readonly string _filePath;

        public S0304HTTTService(IWebHostEnvironment env)
        {
            _filePath = Path.Combine(env.WebRootPath, "dist", "data", "json", "DM_HTTT.json");
        }

        public string GetFilePath()
        {
            return _filePath;
        }

        public async Task<List<M0304HTTTModel>> GetAllHTTT()
        {
            if (!File.Exists(_filePath))
                return new List<M0304HTTTModel>();

            var json = File.ReadAllText(_filePath);
            var htttList = JsonConvert.DeserializeObject<List<M0304HTTTModel>>(json);

            return htttList.OrderBy(httt => httt.ten).ToList();
        }
    }
}
