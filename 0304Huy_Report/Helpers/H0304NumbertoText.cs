namespace H0304.NumberToText.Helpers
{
    public static class H0304NumberToTextHelper
    {
        public static string ConvertSoThanhChu(decimal number)
        {
            string[] dv = { "", "nghìn", "triệu", "tỷ" };
            string[] cs = { "không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín" };

            if (number == 0) return "Không đồng";

            string s = number.ToString("N0");
            string[] parts = s.Split(',');
            string result = "";
            int group = parts.Length;

            foreach (var p in parts)
            {
                int num = int.Parse(p);
                if (num != 0)
                {
                    result += ReadBlock(num, cs) + " " + dv[group - 1] + " ";
                }
                group--;
            }

            return char.ToUpper(result.Trim()[0]) + result.Trim().Substring(1) + " đồng";
        }

        private static string ReadBlock(int num, string[] cs)
        {
            int tram = num / 100;
            int chuc = (num % 100) / 10;
            int donvi = num % 10;
            string r = "";

            if (tram > 0) { r += cs[tram] + " trăm "; if (chuc == 0 && donvi > 0) r += "linh "; }
            if (chuc > 1) { r += cs[chuc] + " mươi "; if (donvi == 1) r += "mốt "; else if (donvi == 5) r += "lăm "; else if (donvi > 0) r += cs[donvi] + " "; }
            else if (chuc == 1) { r += "mười "; if (donvi == 5) r += "lăm "; else if (donvi > 0) r += cs[donvi] + " "; }
            else if (chuc == 0 && donvi > 0) r += cs[donvi] + " ";

            return r.Trim();
        }
    }
}