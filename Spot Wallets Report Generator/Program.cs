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
using System.Management.Automation;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace Spot_Wallets_Report_Generator {

    internal class Program {
        public static Logger log;

        public static bool initialInsert = false;
        public static readonly ConfigFile ini = new ConfigFile($"{Environment.CurrentDirectory}/config.ini");
        private static string ReportFolder, ReportPrefix, SortBy, ReportExtension;
        private static float IgnoreUnder;
        public static string DbPath;
        private static bool UseDB, UseBTCEvol, UseUSDTEvol, AutoTimeSync;
        private static List<Balance> dailyDatas;

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
                        if (args[i] == "-nc" || args[i] == "--noconsole") {
                            var hwnd = GetConsoleWindow();
                            ShowWindow(hwnd, SW_HIDE);
                        }
                    }
                }
                catch (Exception e) {
                    WriteLog("Exception when retrieving arguments.", e.Message + e.StackTrace);
                }

                // Get config
                ReportFolder = ini.ReadKey("Path", "ReportFolder");
                ReportPrefix = ini.ReadKey("Path", "ReportPrefix");
                UseDB = bool.Parse(ini.ReadKey("Options", "UseDatabase"));
                UseBTCEvol = bool.Parse(ini.ReadKey("Options", "UseBTCEvolution").ToLower());
                UseUSDTEvol = bool.Parse(ini.ReadKey("Options", "UseUSDTEvolution").ToLower());
                AutoTimeSync = bool.Parse(ini.ReadKey("Options", "AutoTimeSync").ToLower());
                DbPath = ini.ReadKey("Options", "DatabasePath");
                SortBy = ini.ReadKey("Options", "SortBy").ToLower();
                ReportExtension = ini.ReadKey("Options", "ReportExtension");
                if (!ReportExtension.StartsWith("."))
                    ReportExtension = "." + ReportExtension;
                if (!string.IsNullOrWhiteSpace(ini.ReadKey("Options", "IgnoreUnder")) && !float.TryParse(ini.ReadKey("Options", "IgnoreUnder").Replace('.', ','), out IgnoreUnder))
                    WriteLog($"Synthax error with parameter 'IngoreUnder', value: { ini.ReadKey("Options", "IgnoreUnder")}\r\nMake sure to use only digits with an optional decimal separator like '.' or ','\r\n Ignored parameter.");

                if (ReportFolder == null || ReportPrefix == null || (DbPath == null && UseDB == true)) {
                    WriteLog("Incomplete configuration file.");
                    if (!args.Contains("-nc") || !args.Contains("--noconsole")) {
                        Console.WriteLine("Press Any key to exit.");
                        Console.ReadKey();
                    }
                    Environment.Exit(1);
                }
            }

            // Exception if null for booleans
            catch (Exception e) {
                WriteLog("Error with the configuration file : ", e.Message + e.StackTrace);
                if (!args.Contains("-nc") || !args.Contains("--noconsole")) {
                    Console.WriteLine("Press Any key to exit.");
                    Console.ReadKey();
                }
                Environment.Exit(1);
            }
            try {
                // If user wants to use the local database
                if (UseDB)
                    if (DB.VerifDB() == false) {
                        Console.WriteLine("Verification DB false");
                        Console.ReadKey();
                        Environment.Exit(1);
                    }

                // Get wallets from APIs
                dailyDatas = new List<Balance>();
                if (bool.TryParse(ini.ReadKey("API", "UseBinance").ToLower(), out bool useBinance) == true) {
                    if (useBinance) {
                        
                        WriteLog("Get Binance spot wallet.");
                        dailyDatas.AddRange(BinanceCalls.GetWallet());
                        /*
                        if (CheckTimeStamp())
                            dailyDatas.AddRange(BinanceCalls.GetWallet());
                        else if (AutoTimeSync)
                            SyncSystemTime();
                        else
                            WriteLog("System time is 1000ms+ different from the Binance server time.");
                        */
                    }
                } else {
                    WriteLog("Error when reading the 'UseBinance' option");
                    Environment.Exit(1);
                }

                if (bool.TryParse(ini.ReadKey("API", "UseBybit").ToLower(), out bool useBybit) == true) {
                    if (useBybit) {
                        WriteLog("Get Bybit spot wallet.");
                        dailyDatas.AddRange(BybitCalls.GetWallet());
                    }
                } else {
                    WriteLog("Error when reading the 'UseBybit' option");
                    Environment.Exit(1);
                }

                if (bool.TryParse(ini.ReadKey("API", "UseKucoin").ToLower(), out bool useKucoin) == true) {
                    if (useKucoin) {
                        WriteLog("Get Kucoin spot wallet.");
                        dailyDatas.AddRange(KucoinCalls.GetWallet());
                    }
                } else {
                    WriteLog("Error when reading the 'UseKucoin' option");
                    Environment.Exit(1);
                }


                if (IgnoreUnder != 0)
                    dailyDatas.RemoveAll(d => d.AvgInUSDT < IgnoreUnder);

                if (dailyDatas.Count == 0) {
                    WriteLog("0 asset recovered, exit.");
                    if (!args.Contains("-nc") || !args.Contains("--noconsole")) {
                        Console.WriteLine("Press Any key to exit.");
                        Console.ReadKey();
                    }
                    Environment.Exit(0);
                }

                if (SortBy == "site") {
                    // Sort by site
                    dailyDatas.Sort(delegate (Balance a, Balance b)
                    {
                        if (a.Site == null || b.Site == null) return 0;
                        else if (a.Site == null) return 1;
                        else if (b.Site == null) return -1;
                        else return a.Site.CompareTo(b.Site);
                    });
                } else if (SortBy == "asset") {
                    // Sort by asset asc
                    dailyDatas.Sort(delegate (Balance a, Balance b)
                    {
                        if (a.Asset == null || b.Asset == null) return 0;
                        else if (a.Asset == null) return 1;
                        else if (b.Asset == null) return -1;
                        else return a.Asset.CompareTo(b.Asset);
                    });
                }

                if (UseDB)
                    dailyDatas.ForEach(balance => { DB.InsertAsset(balance); });

                GenerateReport();
                WriteLog("----End of execution----");
            }
            catch (Exception e) {
                WriteLog("Global exception : ", e.Message + e.StackTrace);
            }
        }

        private static void GenerateReport() {
            try {
                WriteLog("Generate report.");
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                bool newFile = File.Exists($"{ReportFolder}{ReportPrefix}_{DateTime.Now:Y}{ReportExtension}") == false;
                FileInfo file = new FileInfo($"{ReportFolder}{ReportPrefix}_{DateTime.Now:Y}{ReportExtension}");
                using (var package = new ExcelPackage(file)) {

                    List<ExcelWorksheet> worksheets = package.Workbook.Worksheets.OrderBy(w => w.Name).ToList();
                    List<KeyValuePair<string, string>> previousTotal = new List<KeyValuePair<string, string>>();

                    // Get last day asset and values for PNL
                    List<Balance> lastDayBalances = new List<Balance>();
                    var lastDay = worksheets.LastOrDefault();

                    if (newFile) {
                        // Search for last month excel
                        if (File.Exists($"{ReportFolder}{ReportPrefix}_{DateTime.Now.AddMonths(-1):Y}{ReportExtension}")) {
                            using (var packageLastMonth = new ExcelPackage($"{ReportFolder}{ReportPrefix}_{DateTime.Now.AddMonths(-1):Y}{ReportExtension}")) {
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
                    sheet.Cells["A1:K" + (i+1)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    sheet.Cells["A1:K" + (i+1)].AutoFitColumns();
                    sheet.Calculate();
                    sheet.ClearFormulas();

                    var table = sheet.Tables.Add(sheet.Cells["A1:K" + (i+1)], $"DayTable_{DateTime.Now:yyyyMMdd}");
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
                        if (UseBTCEvol) {
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
                                sheet.Cells["Q" + (previousTotal.Count + 1)].Value = DateTime.ParseExact(sheet.Name,"yyyyMMdd",null).ToString("yyyy/MM/dd");
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

                        if (UseUSDTEvol) {
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

        private static void SyncSystemTime() {
            try {
                string status = ExecPS("w32tm /query /status");
                WindowsPrincipal principal = new WindowsPrincipal(WindowsIdentity.GetCurrent());
                // If service not started
                if (status.Contains("0x80070426")) {
                    if (!principal.IsInRole(WindowsBuiltInRole.Administrator)) {
                        // Ask to escalate admin privileges
                        ProcessStartInfo processInfo = new ProcessStartInfo
                        {
                            Verb = "runas",
                            FileName = Assembly.GetExecutingAssembly().Location
                        };

                        try {
                            // Relaunch the application with admin rights
                            Process.Start(processInfo);
                            Environment.Exit(0);
                        }
                        // Thrown if the user cancels the prompt
                        catch (Win32Exception) {
                            WriteLog("Please resync the system time and retry.");
                            Environment.Exit(1);
                        }
                    } else {
                        status = ExecPS("restart-service w32time");
                    }
                } else {
                    // Indicateur de dérive : 0(Aucun avertissement) Indicateur de dérive : 3(Non synchronisé)
                    if (!status.Split('\n')[0].Contains("0")) {
                        if (principal.IsInRole(WindowsBuiltInRole.Administrator)) {
                            status = ExecPS("w32tm /resync /force");
                        } else {
                            WriteLog("Please resync the system time and retry.");
                        }
                    }
                }
            }
            catch (Exception e) {
                WriteLog("Exception SyncSystemTime()", e.Message + e.StackTrace);
            }
        }

        private static string ExecPS(string script) {
            string response = "";
            try {
                PowerShell powerShell = PowerShell.Create();
                powerShell.AddScript(script);
                foreach (var className in powerShell.Invoke()) {
                    response += className;
                }
            }
            catch (Exception e) {
                Program.WriteLog($"Exception ExecPS(*)", e.Message + e.StackTrace);
            }
            return response;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns>True if delta < 1000ms </returns>
        private static bool CheckTimeStamp() => Math.Abs(BinanceCalls.GetServerTime() - long.Parse(CommonCalls.GetTimeStamp())) < 1000;

        public static void WriteLog(string message, string exception = null) {
            Console.WriteLine(message);
            log.WriteLog(message, exception);
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
            if (ctrlType == CtrlTypes.CTRL_CLOSE_EVENT && !log.IsWritten)
                File.Delete(log.LogName);
            return true;
        }
    }
}
