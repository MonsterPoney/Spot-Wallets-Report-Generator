using Spot_Wallets_Report_Generator.Models;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Web.Helpers;

/*
https://binance-docs.github.io/apidocs/spot/en/#system-status-sapi-system
*/

namespace Spot_Wallets_Report_Generator.APIcalls {
    internal class BinanceCalls {
        private static readonly string baseUrl = "https://api.binance.com";
        private static readonly string key = Program.ini.ReadKey("API", "KeyBinance");
        private static readonly string secret = Program.ini.ReadKey("API", "SecretBinance");

        private static float btcUsdtPrice = 0;
        public static float BtcUsdtPrice {
            get
            {
                if (btcUsdtPrice == 0) {
                    btcUsdtPrice = GetAveragePrice("BTCUSDT");
                }
                return btcUsdtPrice;
            }
        }

        private static float ethBtcPrice = 0;
        public static float EthBtcPrice {
            get
            {
                if (ethBtcPrice == 0) {
                    ethBtcPrice = GetAveragePrice("ETHBTC");
                }
                return ethBtcPrice;
            }
        }

        internal static HttpClient clientHttp;
        private static HttpClient ClientHttp {
            get
            {
                if (clientHttp == null) {
                    clientHttp = new HttpClient();
                    clientHttp.DefaultRequestHeaders.Add("X-MBX-APIKEY", key);
                }
                return clientHttp;
            }
        }

        public static long GetServerTime() {
            long time = 0;
            try {
                string requestUrl = baseUrl + "/api/v3/time";
                HttpResponseMessage response = ClientHttp.GetAsync(requestUrl).Result;
#if DEBUG
                string ress = response.Content.ReadAsStringAsync().Result;
#endif
                if ((int)response.StatusCode == 200) {
                    dynamic json = Json.Decode(response.Content.ReadAsStringAsync().Result);
                    time = json.serverTime;

                }
            }
            catch (Exception e) {
                Program.WriteLog("Exception BinanceCalls.GetServerTime()", e.Message + e.StackTrace);
            }
            return time;
        }

        public static List<Balance> GetWallet() {
            List<Balance> balances = new List<Balance>();
            try {
                string timeStamp = CommonCalls.GetTimeStamp();
                string requestUrl = baseUrl + $"/sapi/v1/capital/config/getall?timestamp={timeStamp}&signature={CommonCalls.CreateSignature(secret, $"timestamp={timeStamp}")}";

                HttpResponseMessage response = ClientHttp.GetAsync(requestUrl).Result;
#if DEBUG
                string ress = response.Content.ReadAsStringAsync().Result;
#endif
                if ((int)response.StatusCode == 200) {
                    dynamic json = Json.Decode(response.Content.ReadAsStringAsync().Result);
                    foreach (dynamic capital in json) {
                        float free = float.Parse(capital.free.Replace('.', ','));
                        float locked = float.Parse(capital.locked.Replace('.', ','));
                        float freeze = float.Parse(capital.freeze.Replace('.', ','));
                        if (free != 0 || locked != 0 || freeze != 0) {
                            string asset = capital.coin;
                            float valInBTC = 0;
                            float avgPrice = 0;
                            string assetAvg = "";
                            if (asset.EndsWith("UP") || asset.EndsWith("DOWN")) {
                                avgPrice = GetAveragePrice($"{asset}USDT");
                                valInBTC = avgPrice * (free + locked + freeze) / BtcUsdtPrice;
                                assetAvg = "USDT";
                            } else if (asset == "USDT") {
                                valInBTC = (free + locked + freeze) / BtcUsdtPrice;
                                avgPrice = 1;
                                assetAvg = "USDT";
                            } else
                                foreach (dynamic network in capital.networkList) {
                                    if (network.network == "ETF") {
                                        avgPrice = GetAveragePrice($"{asset}USDT");
                                        valInBTC = avgPrice * (free + locked + freeze) / BtcUsdtPrice;
                                        assetAvg = "USDT";
                                    } else if (network.network == "BSC") {
                                        if (asset == "BTC") {
                                            valInBTC = free + locked + freeze;
                                            avgPrice = 1;
                                            assetAvg = "BTC";
                                        } else {
                                            avgPrice = GetAveragePrice($"{asset}BTC");
                                            valInBTC = avgPrice * (free + locked + freeze);
                                            assetAvg = "BTC";
                                        }
                                        break;
                                    } else if (network.network == "ETH") {
                                        if (asset == "ETH") {
                                            valInBTC = (free + locked + freeze) * EthBtcPrice;
                                            avgPrice = 1;
                                            assetAvg = "ETH";
                                        } else {
                                            avgPrice = GetAveragePrice($"{asset}ETH");
                                            valInBTC = avgPrice * (free + locked + freeze) * EthBtcPrice;
                                            assetAvg = "ETH";
                                        }
                                        break;
                                    }
                                }
                            // No "valid" network
                            if (valInBTC == 0) {
                                avgPrice = GetAveragePrice($"{asset}USDT");
                                valInBTC = avgPrice * (free + locked + freeze) / BtcUsdtPrice;
                                assetAvg = "USDT";
                            }

                            balances.Add(new Balance
                            {
                                Asset = asset,
                                Free = free,
                                Locked = locked,
                                Freeze = freeze,
                                AvgInBTC = valInBTC,
                                AvgInUSDT = valInBTC * BtcUsdtPrice,
                                AssetAvg = assetAvg,
                                AvgPrice = avgPrice,
                                Site = "Binance"
                            });
                        }
                    }
                } else {
                    dynamic json = Json.Decode(response.Content.ReadAsStringAsync().Result);
                    Program.WriteLog($"Error BinanceCalls.GetWallet() => error code : {json.code}\r\nMessage : {json.msg}\r\n");
                }

            }
            catch (Exception e) {
                Program.WriteLog("Exception BinanceCalls.GetWallet()", e.Message + e.StackTrace);
            }
            return balances;
        }

        public static float GetAveragePrice(string symbol) {
            float price = 0;
            try {
                using (HttpClient client = new HttpClient()) {
                    string requestUrl = baseUrl + "/api/v3/avgPrice?symbol=" + symbol;
                    client.Timeout = TimeSpan.FromSeconds(30);
                    HttpResponseMessage response = client.GetAsync(requestUrl).Result;
                    if ((int)response.StatusCode == 200) {
                        dynamic json = Json.Decode(response.Content.ReadAsStringAsync().Result);
                        price = float.Parse(json.price.Replace('.', ','));
                    } else if ((int)response.StatusCode == 429) {
                        Program.WriteLog("Breaking Binance API rate limit.");
                        Environment.Exit(1);
                    } else {
                        dynamic json = Json.Decode(response.Content.ReadAsStringAsync().Result);
                        if (json.code != -1121)
                            Program.WriteLog($"Error Binance.GetAveragePrice({symbol}) => error code : {json.code}\r\nMessage : {json.msg}");
                    }
                }

            }
            catch (Exception e) {
                Program.WriteLog($"Exception Binance.GetAveragePrice({symbol})", e.Message + e.StackTrace);
            }
            return price;
        }
    }
}
