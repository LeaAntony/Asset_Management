using System.Security.Cryptography;
using System.Text;

namespace Asset_Management.Function
{
    public class Authentication
    {
        public string GenerateCodeVerifier(int length = 43)
        {
            var randomBytes = new byte[length];
            using (var rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(randomBytes);
            }
            return Base64UrlEncode(randomBytes);
        }

        public string GenerateCodeChallenge(string codeVerifier)
        {
            using (var hasher = SHA256.Create())
            {
                var hashed = hasher.ComputeHash(Encoding.UTF8.GetBytes(codeVerifier));
                return Base64UrlEncode(hashed);
            }
        }

        private string Base64UrlEncode(byte[] input)
        {
            var base64 = Convert.ToBase64String(input)
                                .TrimEnd('=')
                                .Replace('+', '-')
                                .Replace('/', '_');
            return base64;
        }

        public string HashPassword(string password)
        {
            return BCrypt.Net.BCrypt.HashPassword(password, workFactor: 12);
        }

        public bool VerifyPassword(string password, string storedHash)
        {
            if (string.IsNullOrEmpty(storedHash))
                return false;

            if (storedHash.StartsWith("$2", StringComparison.Ordinal))
            {
                return BCrypt.Net.BCrypt.Verify(password, storedHash);
            }

            return storedHash.Equals(LegacyMD5Hash(password), StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsBcryptHash(string storedHash)
        {
            return !string.IsNullOrEmpty(storedHash) && storedHash.StartsWith("$2", StringComparison.Ordinal);
        }

        private static string LegacyMD5Hash(string text)
        {
            using (var md5 = MD5.Create())
            {
                byte[] inputBytes = Encoding.UTF8.GetBytes(text);
                byte[] hashBytes = md5.ComputeHash(inputBytes);

                var sb = new StringBuilder();
                for (int i = 0; i < hashBytes.Length; i++)
                {
                    sb.Append(hashBytes[i].ToString("x2"));
                }
                return sb.ToString();
            }
        }
    }
}
