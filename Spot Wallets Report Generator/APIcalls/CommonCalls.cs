using System;
using System.Security.Cryptography;
using System.Text;

namespace Spot_Wallets_Report_Generator.APIcalls {
    internal class CommonCalls {
        /// <summary>
        /// Generate a HmacSha256 signature
        /// </summary>
        /// <param name="secret"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public static string CreateSignature(string secret, string message) {
            StringBuilder hex = null;
            try {
                byte[] keyByte = Encoding.UTF8.GetBytes(secret);
                byte[] messageBytes = Encoding.UTF8.GetBytes(message);

                HMACSHA256 hash = new HMACSHA256(keyByte);
                byte[] hashBytes = hash.ComputeHash(messageBytes);
                hex = new StringBuilder(hashBytes.Length * 2);

                foreach (var b in hashBytes) {
                    hex.AppendFormat("{0:x2}", b);
                }

            }
            catch (Exception e) {
                Program.WriteLog("Exception CreateSignature()", e.Message + e.StackTrace);
            }
            return hex.ToString();
        }

        /// <summary>
        /// Get timestamp in millis
        /// </summary>
        /// <returns></returns>
        public static string GetTimeStamp() => DateTimeOffset.Now.ToUnixTimeMilliseconds().ToString();
    }
}
