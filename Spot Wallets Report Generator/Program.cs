using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using ProjectEnv;
using Spot_Wallets_Report_Generator.APIcalls;
using Spot_Wallets_Report_Generator.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Security.Principal;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace Spot_Wallets_Report_Generator {
    // https://learn.microsoft.com/en-us/windows/win32/debug/system-error-codes
    internal class Program {
        public static Logger log;

        public static bool initialInsert = false, error = false, stopIfBadTimeStamp;
        public static readonly ConfigFile ini = new ConfigFile($"{Environment.CurrentDirectory}/config.ini");
        private static string reportFolder, reportPrefix, sortBy, reportExtension;
        private static float ignoreUnder;
        private static List<string> ignoredAsset;
        public static string dbPath;
        private static bool useDB, useBTCEvol, useUSDTEvol, openLog = false, openReport = false, autoTimeSync;
        private static List<Balance> dailyDatas;
        private static ExcelPackage package;

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("Kernel32")]
        private static extern IntPtr GetConsoleWindow();

        private const int SW_HIDE = 0;
        static void Main(string[] args) {
            try {
                log = new Logger("./", "", true);
                SetConsoleCtrlHandler(new HandlerRoutine(ConsoleCtrlCheck), true);

                // Get arguments
                try {
                    string Arguments = ini.ReadKey("Options", "Arguments").ToLower();
                    // Add arguments from config file if not specified from Main args
                    if (!string.IsNullOrWhiteSpace(Arguments)) {
                        foreach (string arg in Arguments.Split(' ')) {
                            if (!args.Contains(arg)) {
                                Array.Resize(ref args, args.Length + 1);
                                args[args.Length - 1] = arg;
                            }
                        }
                    }
                    // Apply arguments
                    for (int i = 0; i < args.Length; i++) {
                        WriteLog($"Argument {args[i]}");
                        if (args[i] == "-nc" || args[i] == "--noconsole") {
                            var hwnd = GetConsoleWindow();
                            ShowWindow(hwnd, SW_HIDE);
                        } else if (args[i] == "-ol" || args[i] == "--openlog")
                            openLog = true;
                        else if (args[i] == "-or" || args[i] == "--openreport")
                            openReport = true;
                    }
                }
                catch (Exception e) {
                    WriteLog("Exception when retrieving arguments.", e.Message + e.StackTrace);
                }

                // Get config
                reportFolder = ini.ReadKey("Path", "ReportFolder");
                reportPrefix = ini.ReadKey("Path", "ReportPrefix");
                useDB = bool.Parse(ini.ReadKey("Options", "UseDatabase"));
                useBTCEvol = bool.Parse(ini.ReadKey("Options", "UseBTCEvolution").ToLower());
                useUSDTEvol = bool.Parse(ini.ReadKey("Options", "UseUSDTEvolution").ToLower());
                autoTimeSync = bool.Parse(ini.ReadKey("Options", "AutoTimeSync").ToLower());
                dbPath = ini.ReadKey("Options", "DatabasePath");
                sortBy = ini.ReadKey("Options", "SortBy").ToLower();
                reportExtension = ini.ReadKey("Options", "ReportExtension");
                stopIfBadTimeStamp = bool.Parse(ini.ReadKey("Options", "StopIfBadTimestamp").ToLower());
                if (!reportExtension.StartsWith("."))
                    reportExtension = "." + reportExtension;
                if (!string.IsNullOrWhiteSpace(ini.ReadKey("Options", "IgnoreUnder")) && !float.TryParse(ini.ReadKey("Options", "IgnoreUnder").Replace('.', ','), out ignoreUnder))
                    WriteLog($"Synthax error with parameter 'IngoreUnder', value: { ini.ReadKey("Options", "IgnoreUnder")}\r\nMake sure to use only digits with an optional decimal separator like '.' or ','\r\n Ignored parameter.");
                ignoredAsset = ini.ReadKey("Options", "IgnoreAsset").ToUpper().Split(',').ToList();
                if (ignoredAsset.Count > 0)
                    ignoredAsset.ForEach(it => WriteLog($"Asset {it} ignored in report."));
                if (string.IsNullOrWhiteSpace(reportFolder) || string.IsNullOrWhiteSpace(reportPrefix) || (string.IsNullOrWhiteSpace(dbPath) && useDB == true)) {
                    WriteLog("Incomplete configuration file.");
                    if (!args.Contains("-nc") || !args.Contains("--noconsole")) {
                        Console.WriteLine("Press Any key to exit.");
                        Console.ReadKey();
                    }
                    Environment.Exit(1610);
                }
            }

            // Exception if null for booleans
            catch (Exception e) {
                WriteLog("Error with the configuration file : ", e?.Message + e?.StackTrace);
                if (!args.Contains("-nc") || !args.Contains("--noconsole")) {
                    Console.WriteLine("Press Any key to exit.");
                    Console.ReadKey();
                }
                Environment.Exit(1610);
            }
            try {
                // If user wants to use the local database
                if (useDB)
                    if (DB.VerifDB() == false) {
                        Console.WriteLine("Verification DB false");
                        Console.ReadKey();
                        Environment.Exit(1065);
                    }

                // Get wallets from APIs
                dailyDatas = new List<Balance>();
                if (bool.TryParse(ini.ReadKey("API", "UseBinance").ToLower(), out bool useBinance) == true) {
                    if (useBinance) {

                        WriteLog("Get Binance spot wallet.");

                        if (!CheckTimeStamp()) {
                            if (autoTimeSync) {
                                if (!IsRunAsAdmin()) {
                                    var exeName = Process.GetCurrentProcess().MainModule.FileName;
                                    var startInfo = new ProcessStartInfo(exeName);
                                    startInfo.Verb = "runas";
                                    Process.Start(startInfo);
                                    Environment.Exit(1398);
                                } else {
                                    WriteLog("Sync System time");
                                    SyncSystemTime();
                                }
                            } else {
                                WriteLog("System time is 1000ms+ different from the Binance server time.");
                                if (stopIfBadTimeStamp)
                                    Environment.Exit(1398);
                            }                       
                        }

                        dailyDatas.AddRange(BinanceCalls.GetWallet());
                    }
                } else {
                    WriteLog("Error when reading the 'UseBinance' option");
                    Environment.Exit(1610);
                }

                if (bool.TryParse(ini.ReadKey("API", "UseBybit").ToLower(), out bool useBybit) == true) {
                    if (useBybit) {
                        WriteLog("Get Bybit spot wallet.");
                        dailyDatas.AddRange(BybitCalls.GetWallet());
                    }
                } else {
                    WriteLog("Error when reading the 'UseBybit' option");
                    Environment.Exit(1610);
                }

                if (bool.TryParse(ini.ReadKey("API", "UseKucoin").ToLower(), out bool useKucoin) == true) {
                    if (useKucoin) {
                        WriteLog("Get Kucoin spot wallet.");
                        dailyDatas.AddRange(KucoinCalls.GetWallet());
                    }
                } else {
                    WriteLog("Error when reading the 'UseKucoin' option");
                    Environment.Exit(1610);
                }

                dailyDatas.RemoveAll(d => ignoredAsset.Contains(d.Asset));

                if (ignoreUnder != 0)
                    dailyDatas.RemoveAll(d => d.AvgInUSDT < ignoreUnder);

                if (dailyDatas.Count == 0) {
                    WriteLog("0 asset recovered, exit.");
                    if (!args.Contains("-nc") || !args.Contains("--noconsole")) {
                        Console.WriteLine("Press Any key to exit.");
                        Console.ReadKey();
                    }
                    Environment.Exit(259);
                }

                if (sortBy == "site") {
                    // Sort by site
                    dailyDatas.Sort(delegate (Balance a, Balance b)
                    {
                        if (a.Site == null || b.Site == null) return 0;
                        else if (a.Site == null) return 1;
                        else if (b.Site == null) return -1;
                        else return a.Site.CompareTo(b.Site);
                    });
                } else if (sortBy == "asset") {
                    // Sort by asset asc
                    dailyDatas.Sort(delegate (Balance a, Balance b)
                    {
                        if (a.Asset == null || b.Asset == null) return 0;
                        else if (a.Asset == null) return 1;
                        else if (b.Asset == null) return -1;
                        else return a.Asset.CompareTo(b.Asset);
                    });
                }

                if (useDB)
                    dailyDatas.ForEach(balance => { DB.InsertAsset(balance); });

                GenerateReport();
                WriteLog("----End of execution----");
                Environment.ExitCode=0;
            }
            catch (Exception e) {
                WriteLog("Global exception : ", e.Message + e.StackTrace);
                Environment.ExitCode=574;
            }
        }

        private static void GenerateReport() {
            try {
                WriteLog("Generate report.");
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                bool newFile = File.Exists($"{reportFolder}{reportPrefix}_{DateTime.Now:Y}{reportExtension}") == false;
                FileInfo file = new FileInfo($"{reportFolder}{reportPrefix}_{DateTime.Now:Y}{reportExtension}");
                using (package = new ExcelPackage(file)) {

                    List<ExcelWorksheet> worksheets = package.Workbook.Worksheets.OrderBy(w => w.Name).ToList();
                    List<KeyValuePair<string, string>> previousTotal = new List<KeyValuePair<string, string>>();

                    // Get last day asset and values for PNL
                    List<Balance> lastDayBalances = new List<Balance>();
                    var lastDay = worksheets.LastOrDefault();
          
                    if (newFile) {
                        // Search for last month excel
                        if (File.Exists($"{reportFolder}{reportPrefix}_{DateTime.Now.AddMonths(-1):Y}{reportExtension}")) {
                            using (var packageLastMonth = new ExcelPackage($"{reportFolder}{reportPrefix}_{DateTime.Now.AddMonths(-1):Y}{reportExtension}")) {
                                int y = 2;
                                worksheets = packageLastMonth.Workbook.Worksheets.OrderBy(w => w.Name).ToList();
                                // Get last month last day value in BTC
                                string lastValue = "";
                                while (worksheets.LastOrDefault().Cells["H" + y].Value != null) {
                                    // lastValue format -> "totBTC|totUSDT"
                                    lastValue = $"{worksheets.LastOrDefault().Cells["H" + y].Value}|{worksheets.LastOrDefault().Cells["I" + y].Value}";
                                    y++;
                                }
                                previousTotal.Add(new KeyValuePair<string, string>(worksheets.LastOrDefault().Name, lastValue));

                                // Get last balances
                                lastDay = worksheets.LastOrDefault();
                                if (lastDay != null) {
                                    y = 2;
                                    // Use B column to avoid "Total" cell
                                    while (lastDay.Cells["B" + y].Value != null) {
                                        lastDayBalances.Add(new Balance
                                        {
                                            Asset = lastDay.Cells["A" + y].Value.ToString(),
                                            AvgInBTC = float.Parse(lastDay.Cells["H" + y].Value.ToString()),
                                            AvgInUSDT = float.Parse(lastDay.Cells["I" + y].Value.ToString()),
                                            Site = lastDay.Cells["B" + y].Value.ToString(),
                                            Notes = lastDay.Cells["K" + y].Value?.ToString()
                                        });
                                        y++;
                                    }
                                }
                            }
                        }
                    } else {
                        int y = 2;
                        // Get all last sheets total values in BTC
                        foreach (var sheeet in worksheets) {
                            y = 2;
                            string lastValue = "";
                            while (sheeet.Cells["H" + y].Value != null) {
                                //totBTC|totUSDT
                                lastValue = $"{sheeet.Cells["H" + y].Value}|{sheeet.Cells["I" + y].Value}";
                                y++;
                            }
                            previousTotal.Add(new KeyValuePair<string, string>(sheeet.Name, lastValue));
                        }

                        // Get last balances
                        if (lastDay != null) {
                            y = 2;
                            // Use B column to avoid "Total" line
                            while (lastDay.Cells["B" + y].Value != null) {
                                lastDayBalances.Add(new Balance
                                {
                                    Asset = lastDay.Cells["A" + y].Value.ToString(),
                                    AvgInBTC = float.Parse(lastDay.Cells["H" + y].Value.ToString()),
                                    AvgInUSDT = float.Parse(lastDay.Cells["I" + y].Value.ToString()),
                                    Site = lastDay.Cells["B" + y].Value.ToString(),
                                    Notes = lastDay.Cells["K" + y].Value?.ToString()
                                });
                                y++;
                            }
                        }
                    }

                    if (package.Workbook.Worksheets.Any(w => w.Name == DateTime.Now.ToString("yyyyMMdd"))) {
                        WriteLog("Rewriting previous sheet.");
                        package.Workbook.Worksheets.Delete(DateTime.Now.ToString("yyyyMMdd"));
                    }
                    ExcelWorksheet sheet = package.Workbook.Worksheets.Add(DateTime.Now.ToString("yyyyMMdd"));
                    sheet.Cells["A1"].Value = "Asset";
                    sheet.Cells["B1"].Value = "Plateform";
                    sheet.Cells["C1"].Value = "Free";
                    sheet.Cells["D1"].Value = "Locked";
                    sheet.Cells["E1"].Value = "Freezed";
                    sheet.Cells["F1"].Value = "Price";
                    sheet.Cells["G1"].Value = "/Asset";
                    sheet.Cells["H1"].Value = "Average BTC";
                    sheet.Cells["I1"].Value = "Average USDT";
                    sheet.Cells["J1"].Value = "Daily PNL (USDT)";
                    sheet.Cells["K1"].Value = "Notes";

                    int i = 2;
                    foreach (Balance balance in dailyDatas) {
                        sheet.Cells["A" + i].Value = balance.Asset;
                        sheet.Cells["B" + i].Value = balance.Site;
                        sheet.Cells["C" + i].Value = balance.Free;
                        sheet.Cells["D" + i].Value = balance.Locked;
                        sheet.Cells["E" + i].Value = balance.Freeze;
                        sheet.Cells["F" + i].Value = balance.AvgPrice;
                        sheet.Cells["G" + i].Value = "/" + balance.AssetAvg;

                        float lockb = balance.Locked;
                        float freeb = balance.Free;
                        float freezb = balance.Freeze;
                        sheet.Cells["H" + i].Value = balance.AvgInBTC;
                        sheet.Cells["I" + i].Value = balance.AvgInBTC * BinanceCalls.BtcUsdtPrice;
                        // TODO : config => PNL on x days
                        if (lastDayBalances != null) {

                            // USDT PNL
                            sheet.Cells["J" + i].Value = ((balance.AvgInUSDT - lastDayBalances.FirstOrDefault(b => b.Asset == balance.Asset && b.Site == balance.Site)?.AvgInUSDT) / lastDayBalances.FirstOrDefault(b => b.Asset == balance.Asset && b.Site == balance.Site)?.AvgInUSDT);
                            if (sheet.Cells["J" + i].Value != null) {
                                sheet.Cells["J" + i].Style.Font.Color.SetColor((float.Parse(sheet.Cells["J" + i].Value.ToString()) < 0) ? Color.Red : Color.Green);
                            }

                            sheet.Cells["K" + i].Value = lastDayBalances.FirstOrDefault(b => b.Asset == balance.Asset)?.Notes;
                        }

                        i++;
                    }
                    i--;

                    sheet.Cells["C2:E" + i].Style.Numberformat.Format = "#,##0.000";
                    sheet.Cells["F2:F" + i].Style.Numberformat.Format = "#,##0.0000000";
                    sheet.Cells["H2:H" + i].Style.Numberformat.Format = "#,##0.000000";
                    sheet.Cells["J2:J" + i].Style.Numberformat.Format = "#,##0.00%";
                    sheet.Cells["A" + (i + 1)].Value = "Total";
                    sheet.Cells["H" + (i + 1)].Formula = $"SUM(H2:H{i})";
                    sheet.Cells["I" + (i + 1)].Formula = $"SUM(I2:I{i})";
                    sheet.Cells["A1:K" + (i + 1)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    sheet.Cells["A1:K" + (i + 1)].AutoFitColumns();
                    sheet.Calculate();
                    sheet.ClearFormulas();

                    var table = sheet.Tables.Add(sheet.Cells["A1:K" + (i + 1)], $"DayTable_{DateTime.Now:yyyyMMdd}");
                    table.TableStyle = OfficeOpenXml.Table.TableStyles.Light1;

                    try {
                        // TODO :  config => choose charts style
                        // Ok : Pie, Sunburst, OfPie, line, funnel, doughnut, bar
                        ExcelPieChart pieChart = sheet.Drawings.AddPieChart($"Repartition of assets {DateTime.Now:yyyyMMdd}", ePieChartType.Pie);
                        pieChart.SetPosition(1, 0, 12, 0);
                        pieChart.SetSize(400, 400);
                        pieChart.Series.Add(sheet.Cells["H2:H" + (i)], sheet.Cells["A2:A" + (i)]);
                        pieChart.StyleManager.SetChartStyle(ePresetChartStyle.PieChartStyle11);
                        pieChart.RoundedCorners = true;
                    }
                    catch (Exception e) {
                        WriteLog("Exception when creating Repartion char", e.Message + e.StackTrace);
                    }

                    // Wallet evolution
                    if (lastDay != null) {
                        if (useBTCEvol) {
                            try {
                                ExcelLineChart lineChart = sheet.Drawings.AddLineChart($"Evolution in BTC {DateTime.Now:yyyyMMdd}", eLineChartType.Line);
                                lineChart.Title.Text = "Evolution in BTC";
                                lineChart.Legend.Position = eLegendPosition.Right;
                                lineChart.SetPosition(22, 0, 12, 0);
                                lineChart.SetSize(400, 400);
                                for (int x = 0; x < previousTotal.Count; x++) {
                                    sheet.Cells["Q" + (x + 1)].Value = DateTime.ParseExact(previousTotal[x].Key, "yyyyMMdd", null).ToString("yyyy/MM/dd");
                                    sheet.Cells["R" + (x + 1)].Value = float.Parse(previousTotal[x].Value.Split('|')[0]);
                                }
                                // Date format
                                sheet.Cells["Q" + (previousTotal.Count + 1)].Value = DateTime.ParseExact(sheet.Name, "yyyyMMdd", null).ToString("yyyy/MM/dd");
                                sheet.Cells["Q1:Q" + (previousTotal.Count + 1)].Style.Numberformat.Format = "0";
                                sheet.Cells["R" + (previousTotal.Count + 1)].Value = sheet.Cells["H" + (i + 1)].Value;
                                sheet.Cells["R1:R" + (previousTotal.Count + 1)].Style.Numberformat.Format = "#,##0.000000";
                                lineChart.Series.Add(sheet.Cells["R1:R" + (previousTotal.Count + 1)], sheet.Cells["Q1:Q" + (previousTotal.Count + 1)]);
                                lineChart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle1);
                            }
                            catch (Exception e) {
                                WriteLog("Exception when creating BTC evolution chart", e.Message + e.StackTrace);
                            }
                        }

                        if (useUSDTEvol) {
                            try {
                                ExcelLineChart lineChartUSDT = sheet.Drawings.AddLineChart($"Evolution in USDT {DateTime.Now:yyyyMMdd}", eLineChartType.Line);
                                lineChartUSDT.Title.Text = "Evolution in USDT";
                                lineChartUSDT.SetPosition(22, 0, 18, 0);
                                lineChartUSDT.SetSize(400, 400);
                                for (int x = 0; x < previousTotal.Count; x++) {
                                    sheet.Cells["S" + (x + 1)].Value = DateTime.ParseExact(previousTotal[x].Key, "yyyyMMdd", null).ToString("yyyy/MM/dd");
                                    sheet.Cells["T" + (x + 1)].Value = float.Parse(previousTotal[x].Value.Split('|')[1]);
                                }
                                sheet.Cells["S" + (previousTotal.Count + 1)].Value = DateTime.ParseExact(sheet.Name, "yyyyMMdd", null).ToString("yyyy/MM/dd");
                                sheet.Cells["S1:S" + (previousTotal.Count + 1)].Style.Numberformat.Format = "0";
                                sheet.Cells["T" + (previousTotal.Count + 1)].Value = sheet.Cells["I" + (i + 1)].Value;
                                sheet.Cells["T1:T" + (previousTotal.Count + 1)].Style.Numberformat.Format = "#,##0.00";
                                lineChartUSDT.Series.Add(sheet.Cells["T1:T" + (previousTotal.Count + 1)], sheet.Cells["S1:S" + (previousTotal.Count + 1)]);
                                lineChartUSDT.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle1);
                            }
                            catch (Exception e) {
                                WriteLog("Exception when creating USDT evolution chart", e.Message + e.StackTrace);
                            }
                        }
                    }

                    // Save to file
                    package.Save();
                }
            }
            catch (Exception e) {
                WriteLog("Exception when generating report", e.Message + e.StackTrace);
            }
        }

        private static bool IsRunAsAdmin() {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private static void SyncSystemTime() {
            try {
                //default Windows time server
                const string ntpServer = "time.windows.com";

                // NTP message size - 16 bytes of the digest (RFC 2030)
                var ntpData = new byte[48];

                //Setting the Leap Indicator, Version Number and Mode values
                ntpData[0] = 0x1B; //LI = 0 (no warning), VN = 3 (IPv4 only), Mode = 3 (Client Mode)

                var addresses = Dns.GetHostEntry(ntpServer).AddressList;

                //The UDP port number assigned to NTP is 123
                var ipEndPoint = new IPEndPoint(addresses[0], 123);
                //NTP uses UDP

                using (var socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp)) {
                    socket.Connect(ipEndPoint);

                    //Stops code hang if NTP is blocked
                    socket.ReceiveTimeout = 3000;

                    socket.Send(ntpData);
                    socket.Receive(ntpData);
                    socket.Close();
                }

                //Offset to get to the "Transmit Timestamp" field (time at which the reply 
                //departed the server for the client, in 64-bit timestamp format."
                const byte serverReplyTime = 40;

                //Get the seconds part
                ulong intPart = BitConverter.ToUInt32(ntpData, serverReplyTime);

                //Get the seconds fraction
                ulong fractPart = BitConverter.ToUInt32(ntpData, serverReplyTime + 4);

                //Convert From big-endian to little-endian
                intPart = SwapEndianness(intPart);
                fractPart = SwapEndianness(fractPart);

                var milliseconds = (intPart * 1000) + ((fractPart * 1000) / 0x100000000L);


                var dtDateTime = new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                dtDateTime = dtDateTime.AddMilliseconds((long)milliseconds).ToLocalTime();

                NativeMethods.SYSTEMTIME time = new NativeMethods.SYSTEMTIME
                {
                    wDay = (ushort)dtDateTime.Day,
                    wMonth = (ushort)dtDateTime.Month,
                    wYear = (ushort)dtDateTime.Year,
                    wHour = (ushort)dtDateTime.Hour,
                    wMinute = (ushort)dtDateTime.Minute,
                    wSecond = (ushort)dtDateTime.Second,
                    wMilliseconds = (ushort)dtDateTime.Millisecond
                };

                if (!NativeMethods.SetLocalTime(ref time)) {
                    // The native function call failed, so throw an exception
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }
            }
            catch (Exception ex) {
                Console.WriteLine("Error : " + ex.Message);
            }
        }

        static uint SwapEndianness(ulong x) {
            return (uint)(((x & 0x000000ff) << 24) +
                           ((x & 0x0000ff00) << 8) +
                           ((x & 0x00ff0000) >> 8) +
                           ((x & 0xff000000) >> 24));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns>True if delta < 1000ms </returns>
        private static bool CheckTimeStamp() => Math.Abs(BinanceCalls.GetServerTime() - long.Parse(CommonCalls.GetTimeStamp())) < 1000;
       
        public static void WriteLog(string message, string exception = null) {
            Console.WriteLine(message);
            log.WriteLog(message??"", exception);
        }

        // Declare the SetConsoleCtrlHandler function
        // as external and receiving a delegate.
        [DllImport("Kernel32")]
        public static extern bool SetConsoleCtrlHandler(HandlerRoutine Handler, bool Add);

        // A delegate type to be used as the handler routine
        // for SetConsoleCtrlHandler.
        public delegate bool HandlerRoutine(CtrlTypes CtrlType);

        // An enumerated type for the control messages
        // sent to the handler routine.
        public enum CtrlTypes {
            CTRL_C_EVENT = 0,
            CTRL_BREAK_EVENT,
            CTRL_CLOSE_EVENT,
            CTRL_LOGOFF_EVENT = 5,
            CTRL_SHUTDOWN_EVENT
        }

        public static bool ConsoleCtrlCheck(CtrlTypes ctrlType) {
            if (package != null) {
                package.Save();
                package.Dispose();
            }
            if (openLog && (log.HasException || error)) {
                try {
                    Process.Start(new FileInfo(log.LogName).FullName);
                }
                catch (Exception e) {
                    WriteLog($"Exception when trying to open log.", e.Message + e.StackTrace);
                }
            }
            if (openReport) {
                try {
                    Process.Start($"{reportFolder}{reportPrefix}_{DateTime.Now:Y}{reportExtension}");
                }
                catch (Exception e) {
                    WriteLog($"Exception when trying to open report.", e.Message + e.StackTrace);
                }
            }
            return true;
        }
    }
    static class NativeMethods {
        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
        internal static extern bool SetLocalTime(ref System.DateTime lpSystemTime);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern bool SetLocalTime(ref SYSTEMTIME lpSystemTime);

        [StructLayout(LayoutKind.Sequential)]
        internal struct SYSTEMTIME {
            public ushort wYear;
            public ushort wMonth;
            public ushort wDayOfWeek;    // ignored for the SetLocalTime function
            public ushort wDay;
            public ushort wHour;
            public ushort wMinute;
            public ushort wSecond;
            public ushort wMilliseconds;
        }
    }
}
