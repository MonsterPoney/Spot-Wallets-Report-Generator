using Spot_Wallets_Report_Generator.Models;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Web.Helpers;

/*
https://docs.kucoin.com/#list-accounts
*/

namespace Spot_Wallets_Report_Generator.APIcalls {
    internal class KucoinCalls {
        private static readonly string baseUrl = "https://api.kucoin.com";
        private static readonly string key = Program.ini.ReadKey("API", "KeyKucoin");
        private static readonly string secret = Program.ini.ReadKey("API", "SecretKucoin");
        private static readonly string passPhrase = Program.ini.ReadKey("API", "PassPhraseKucoin");

        internal static float btcUsdtPrice = 0;
        public static float BtcUsdtPrice {
            get
            {
                if (btcUsdtPrice == 0) {
                    btcUsdtPrice = GetAveragePrice("BTC-USDT");
                }
                return btcUsdtPrice;
            }
        }

        private static float ethBtcPrice = 0;
        public static float EthBtcPrice {
            get
            {
                if (ethBtcPrice == 0) {
                    ethBtcPrice = GetAveragePrice("ETH-BTC");
                }
                return ethBtcPrice;
            }
        }

        private static float kcsBtcPrice = 0;
        public static float KcsBtcPrice {
            get
            {
                if (kcsBtcPrice == 0) {
                    kcsBtcPrice = GetAveragePrice("KCS-BTC");
                }
                return kcsBtcPrice;
            }
        }

        private static HttpClient clientHttp;
        private static HttpClient ClientHttp {
            get
            {
                if (clientHttp == null) {
                    clientHttp = new HttpClient();
                    clientHttp.DefaultRequestHeaders.Add("KC-API-KEY", key);
                    clientHttp.DefaultRequestHeaders.Add("KC-API-KEY-VERSION", "2");
                    clientHttp.DefaultRequestHeaders.Add("KC-API-PASSPHRASE", CreateSignatureB64(secret, passPhrase));

                }
                return clientHttp;
            }
        }

        public static List<Balance> GetWallet() {
            List<Balance> balances = new List<Balance>();
            try {
                string timeStamp = CommonCalls.GetTimeStamp();
                string requestUrl = baseUrl + "/api/v1/accounts";
                ClientHttp.DefaultRequestHeaders.Add("KC-API-SIGN", CreateSignatureB64(secret, timeStamp + "GET" + "/api/v1/accounts"));
                ClientHttp.DefaultRequestHeaders.Add("KC-API-TIMESTAMP", timeStamp);

                HttpResponseMessage response = ClientHttp.GetAsync(requestUrl).Result;
#if DEBUG
                string ress = response.Content.ReadAsStringAsync().Result;
#endif
                dynamic json = Json.Decode(response.Content.ReadAsStringAsync().Result);
                if ((int)response.StatusCode == 200) {
                    if (int.Parse(json.code) != 400002) {
                        foreach (dynamic balance in json.data) {
                            if (balance.type == "main") {
                                float free = float.Parse(balance.available.Replace('.', ','));
                                float locked = float.Parse(balance.holds.Replace('.', ','));
                                string asset = balance.currency;
                                float valInBTC = 0;
                                float avgPrice = 0;
                                string assetAvg = "";
                                if (asset == "USDT") {
                                    avgPrice = 1;
                                    valInBTC = (free + locked) / BtcUsdtPrice;
                                    assetAvg = "USDT";
                                } else {
                                    avgPrice = GetAveragePrice($"{asset}-USDT");
                                    if (avgPrice == 0) {
                                        avgPrice = GetAveragePrice($"{asset}-BTC");
                                        if (avgPrice == 0) {
                                            avgPrice = GetAveragePrice($"{asset}-ETH");
                                            if (avgPrice == 0) {
                                                avgPrice = GetAveragePrice($"{asset}-KCS");
                                                if (avgPrice == 0) {
                                                    assetAvg = "N/A";
                                                    valInBTC = 0;
                                                } else {
                                                    assetAvg = "KCS";
                                                    valInBTC = avgPrice * (free + locked) * KcsBtcPrice;
                                                }
                                            } else {
                                                assetAvg = "ETH";
                                                valInBTC = avgPrice * (free + locked) * EthBtcPrice;
                                            }
                                        } else {
                                            assetAvg = "BTC";
                                            valInBTC = avgPrice * (free + locked);
                                        }
                                    } else {
                                        assetAvg = "USDT";
                                        valInBTC = avgPrice * (free + locked) / BtcUsdtPrice;
                                    }
                                }

                                balances.Add(new Balance
                                {
                                    Asset = asset,
                                    Free = free,
                                    Locked = locked,
                                    AvgInBTC = valInBTC,
                                    AvgInUSDT = valInBTC * BtcUsdtPrice,
                                    AssetAvg = assetAvg,
                                    AvgPrice = avgPrice,
                                    Site = "KuCoin"
                                });
                            }
                        }
                    } else {
                        Program.WriteLog($"KuCoin GetWallet() : 400002 Timestamp invalid");
                    }
                } else {
                    Program.WriteLog($"KuCoin GetWallet() : API return code {json.code}");
                }

            }
            catch (Exception e) {
                Program.WriteLog("Exception KuCoinCalls.GetWallet()", e.Message + e.StackTrace);
            }
            return balances;
        }

        public static float GetAveragePrice(string symbol) {
            float price = 0;
            try {
                string requestUrl = baseUrl + "/api/v1/market/orderbook/level1?symbol=" + symbol;
                string timeStamp = CommonCalls.GetTimeStamp();
                ClientHttp.DefaultRequestHeaders.Remove("KC-API-SIGN");
                ClientHttp.DefaultRequestHeaders.Remove("KC-API-TIMESTAMP");
                ClientHttp.DefaultRequestHeaders.Add("KC-API-SIGN", CreateSignatureB64(secret, timeStamp + "GET" + "/api/v1/accounts"));
                ClientHttp.DefaultRequestHeaders.Add("KC-API-TIMESTAMP", timeStamp);
                HttpResponseMessage response = ClientHttp.GetAsync(requestUrl).Result;
#if DEBUG
                string ress = response.Content.ReadAsStringAsync().Result;
#endif
                if ((int)response.StatusCode == 200) {
                    dynamic json = Json.Decode(response.Content.ReadAsStringAsync().Result);
                    if (json.data != null) {
                        price = float.Parse(json.data.price.Replace('.', ','));
                    }
                } else {
                    dynamic json = Json.Decode(response.Content.ReadAsStringAsync().Result);
                    Program.WriteLog($"Error KucoinCalls.GetAveragePrice({symbol}) => code {json.code} : {json.data}");
                    Program.error = true;
                }
            }
            catch (Exception e) {
                Program.WriteLog($"Exception KuCoinCalls.GetAveragePrice({symbol})", e.Message + e.StackTrace);
            }
            return price;
        }

        /// <summary>
        /// Generate a HmacSha256 signature
        /// </summary>
        /// <param name="key"></param>
        /// <param name="query"></param>
        /// <returns></returns>
        public static string CreateSignatureB64(string key, string query) {
            byte[] bytes = null;
            try {
                byte[] keyBytes = Encoding.UTF8.GetBytes(key);
                byte[] queryStringBytes = Encoding.UTF8.GetBytes(query);
                HMACSHA256 hmacsha256 = new HMACSHA256(keyBytes);

                bytes = hmacsha256.ComputeHash(queryStringBytes);
            }
            catch (Exception e) {
                Program.WriteLog("Exeption CreateSignatureB64(*,*)", e.Message + e.StackTrace);
            }
            return Convert.ToBase64String(bytes);
        }
    }
}
