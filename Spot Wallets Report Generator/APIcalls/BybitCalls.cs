using Spot_Wallets_Report_Generator.Models;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Web.Helpers;

namespace Spot_Wallets_Report_Generator.APIcalls {
    //https://bybit-exchange.github.io/docs/inverse/#t-introduction
    internal class BybitCalls {
        private static readonly string baseUrl = "https://api.bybit.com";
        private static readonly string key = Program.ini.ReadKey("API", "KeyBybit");
        private static readonly string secret = Program.ini.ReadKey("API", "SecretBybit");

        internal static float btcUsdtPrice = 0;
        public static float BtcUsdtPrice {
            get
            {
                if (btcUsdtPrice == 0) {
                    btcUsdtPrice = GetAveragePrice("BTCUSDT");
                }
                return btcUsdtPrice;
            }
        }

        internal static float usdcBtcPrice = 0;
        public static float UsdcBtcPrice {
            get
            {
                if (usdcBtcPrice == 0) {
                    usdcBtcPrice = GetAveragePrice("USDCBTC");
                }
                return usdcBtcPrice;
            }
        }

        private static HttpClient clientHttp;
        private static HttpClient ClientHttp {
            get
            {
                if (clientHttp == null) {
                    clientHttp = new HttpClient();
                }
                return clientHttp;
            }
        }

        public static List<Balance> GetWallet() {
            List<Balance> balances = new List<Balance>();
            try {
                string timeStamp = CommonCalls.GetTimeStamp();
                string requestUrl = baseUrl + $"/spot/v1/account?api_key={key}&timestamp={timeStamp}&sign={CommonCalls.CreateSignature(secret, $"api_key={key}&timestamp={timeStamp}")}";

                HttpResponseMessage response = ClientHttp.GetAsync(requestUrl).Result;
#if DEBUG
                string ress = response.Content.ReadAsStringAsync().Result;
#endif
                if ((int)response.StatusCode == 200) {
                    dynamic json = Json.Decode(response.Content.ReadAsStringAsync().Result);
                    if (json.ret_code == 0) {
                        foreach (dynamic balance in json.result.balances) {
                            float free = float.Parse(balance.free.Replace('.', ','));
                            float locked = float.Parse(balance.locked.Replace('.', ','));
                            string asset = balance.coin;
                            float valInBTC = 0;
                            float avgPrice = 0;
                            string assetAvg = "";
                            if (asset == "USDT") {
                                avgPrice = 1;
                                valInBTC = (free + locked) * btcUsdtPrice;
                                break;
                            }
                            avgPrice = GetAveragePrice($"{asset}USDT");
                            if (avgPrice == -100010) {
                                avgPrice = GetAveragePrice($"{asset}BTC");
                                if (avgPrice == -100010) {
                                    avgPrice = GetAveragePrice($"{asset}USDC");
                                    if (avgPrice == -100010) {
                                        avgPrice = 0;
                                        assetAvg = "N/A";
                                        valInBTC = 0;
                                    } else {
                                        assetAvg = "USDC";
                                        valInBTC = avgPrice * (free + locked) * UsdcBtcPrice;
                                    }
                                } else {
                                    assetAvg = "BTC";
                                    valInBTC = avgPrice * (free + locked);
                                }
                            } else {
                                assetAvg = "USDT";
                                valInBTC = avgPrice * (free + locked) / BtcUsdtPrice;
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
                                Site = "Bybit"
                            });
                        }
                    } else {
                        Program.WriteLog($"Error BybitCalls.GetWallet(), code {json.ret_code} : {json.ret_msg}");
                        Program.error = true;
                    }
                } else {
                    Program.WriteLog($"Error contacting Bybit API, response code : {response.StatusCode}.");
                    Program.error = true;
                }

            }
            catch (Exception e) {
                Program.WriteLog("Exception BybitCalls.GetWallet()", e.Message + e.StackTrace);
            }
            return balances;
        }

        public static float GetAveragePrice(string symbol) {
            float price = 0;
            try {
                string requestUrl = baseUrl + $"/spot/quote/v1/ticker/24hr?symbol={symbol}";

                HttpResponseMessage response = ClientHttp.GetAsync(requestUrl).Result;
#if DEBUG
                string ress = response.Content.ReadAsStringAsync().Result;
#endif
                if ((int)response.StatusCode == 200) {
                    dynamic json = Json.Decode(response.Content.ReadAsStringAsync().Result);
                    price = float.Parse(json.result.lastPrice.Replace('.', ','));
                } else {
                    dynamic json = Json.Decode(response.Content.ReadAsStringAsync().Result);
                    if (json.ret_code == -100010)
                        return -100010;
                    Program.WriteLog($"Error BybitCalls.GetAveragePrice({symbol}) -> {json.ret_code} : {json.ret_msg}");
                    Program.error = true;
                }
            }
            catch (Exception e) {
                Program.WriteLog($"Exception BybitCalls.GetAveragePrice({symbol})", e.Message + e.StackTrace);
            }
            return price;
        }
    }
}
