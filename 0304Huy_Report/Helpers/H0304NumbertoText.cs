using System.Globalization;
using System.Text.RegularExpressions;

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

        public static string chuyenDoiSoTienThanhChu2(string soTien)
        {
            var chuoiChu = hamChuyenDoiChuoiSoThanhChuoiChu(soTien);

            if (chuoiChu.Trim() == "")
            {
                chuoiChu += " Không đồng";
            }
            else
            {
                if (chuoiChu.Trim().ToLower() == "không")
                {
                    chuoiChu += " đồng";
                }
                else
                {
                    if (chuoiChu.Contains("lẻ"))
                    {
                        chuoiChu += " đồng";
                    }
                    else
                    {
                        chuoiChu += " đồng chẵn";
                    }
                }
            }

            chuoiChu = Regex.Replace(chuoiChu, @"\s+", " ").Trim();
            var chuoiChuSo = vietHoaChuCaiDauTien(chuoiChu);

            return chuoiChuSo;
        }
        public static string vietHoaChuCaiDauTien(string chuoi)
        {
            if (string.IsNullOrWhiteSpace(chuoi))
            {
                return chuoi;
            }

            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            return string.Join(". ", Regex.Split(chuoi, @"\.\s+").Select(c => char.ToUpper(c[0]) + c.Substring(1).ToLower()));
        }
        public static string hamChuyenDoiChuoiSoThanhChuoiChu(string soTien)
        {
            string chuSoTien = "";
            string[] donVi = { "", "ngàn", "triệu", "tỷ", "nghìn tỷ", "nghìn triệu tỷ" };
            string[] so = { "", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín" };

            string HamDocHaiChuSo(int soHienTai, string chuoiHienTai)
            {
                string ketQua = "";
                int hangChuc = soHienTai / 10;
                int hangDonVi = soHienTai % 10;

                if (hangChuc > 0)
                {
                    if (hangChuc == 1)
                    {
                        ketQua += "mười";
                    }
                    else
                    {
                        ketQua += so[hangChuc] + " mươi";
                    }
                    if (hangDonVi > 0)
                    {
                        switch (hangDonVi)
                        {
                            case 1:
                                if (hangChuc == 1)
                                {
                                    ketQua += " " + so[hangDonVi];
                                }
                                else
                                {
                                    ketQua += " mốt";
                                }
                                break;
                            case 5:
                                ketQua += " lăm";
                                break;
                            default:
                                ketQua += " " + so[hangDonVi];
                                break;
                        }
                    }
                }
                else
                {
                    if (hangDonVi > 0)
                    {
                        if (chuoiHienTai.Length > 0)
                        {
                            if (!ketQua.EndsWith(" linh ") && chuoiHienTai != "Âm")
                            {
                                ketQua += " linh ";
                            }
                        }
                        ketQua += so[hangDonVi];
                    }
                }
                return ketQua;
            }


            string HamDocBaChuSo(int soHienTai, string chuoiHienTai)
            {
                string ketQua = "";
                int hangTram = soHienTai / 100;
                int hangChucDonVi = soHienTai % 100;
                int hangDonVi = soHienTai % 10;

                if (hangTram > 0)
                {
                    ketQua += so[hangTram] + " trăm ";
                    if (hangChucDonVi > 0)
                    {
                        ketQua += " " + HamDocHaiChuSo(hangChucDonVi, ketQua);
                    }
                }
                else
                {
                    if (hangChucDonVi > 0)
                    {
                        if (chuoiHienTai.Length > 0 && chuoiHienTai != "Âm")
                        {
                            ketQua += " không trăm ";
                        }
                        ketQua += HamDocHaiChuSo(hangChucDonVi, ketQua);
                    }
                }
                return ketQua;
            }


            string ChuyenSoTienThanhChu(string soTien)
            {
                long phanNguyen = 0;
                long phanThapPhan = 0;
                string chuoiPhanThapPhan = "";
                string chuoiKetQua = "";

                if (soTien.Length > 0)
                {
                    // Một tỷ hai trăm bốn triệu năm trăm năm mươi mốt nghìn hai trăm đồng lẻ năm trăm.
                    // Tách phần nguyên và phần thập phân
                    if (soTien.TrimStart().StartsWith("-"))
                    {
                        chuoiKetQua += "Âm";
                        soTien = soTien.TrimStart('-', ' ');
                    }

                    if (soTien.Contains("."))
                    {
                        int viTriDauCham = soTien.IndexOf('.');
                        phanNguyen = long.Parse(soTien.Substring(0, viTriDauCham));
                        chuoiPhanThapPhan = soTien.Substring(viTriDauCham + 1);
                        phanThapPhan = long.Parse(soTien.Substring(viTriDauCham + 1));
                    }
                    else
                    {
                        phanNguyen = long.Parse(soTien);
                    }
                    if (phanNguyen == 0)
                    {
                        chuoiKetQua = "Không";
                    }
                    else
                    {
                        for (int i = donVi.Length - 1; i >= 0; i--)
                        {
                            int soHienTai = (int)(phanNguyen / Math.Pow(10, i * 3));
                            phanNguyen -= soHienTai * (long)Math.Pow(10, i * 3);

                            if (soHienTai > 0)
                            {
                                if (soHienTai >= 10)
                                {
                                    chuoiKetQua += " " + HamDocBaChuSo(soHienTai, chuoiKetQua);
                                }
                                else
                                {
                                    if (chuoiKetQua.Length > 0 && chuoiKetQua != "Âm")
                                    {
                                        chuoiKetQua += " không trăm ";
                                    }
                                    chuoiKetQua += " " + HamDocHaiChuSo(soHienTai, chuoiKetQua);
                                }
                                chuoiKetQua += " " + donVi[i];
                            }
                        }
                    }
                    if (phanThapPhan > 0)
                    {
                        //làm tròn 3 chữ số thập phân
                        chuoiKetQua += " lẻ ";
                        if (phanThapPhan.ToString().Length < chuoiPhanThapPhan.Length)
                        {
                            chuoiKetQua += " không trăm ";
                        }
                        chuoiKetQua += HamDocBaChuSo((int)phanThapPhan, "");
                    }
                }
                return chuoiKetQua;
            }
            chuSoTien = ChuyenSoTienThanhChu(soTien.Replace(",", ""));
            chuSoTien = Regex.Replace(chuSoTien, @"\s+", " ").Trim();
            var chuoiChuSo = vietHoaChuCaiDauTien(chuSoTien);
            return chuoiChuSo;
        }
    }
}