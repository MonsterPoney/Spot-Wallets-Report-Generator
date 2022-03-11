using Spot_Wallets_Report_Generator.Models;
using System;
using System.Data.SQLite;

namespace Spot_Wallets_Report_Generator {
    internal class DB {
        private static readonly string cnx = $"URI=file:{Program.dbPath}";

        /// <summary>
        /// Create db if not exists
        /// </summary>
        public static bool VerifDB() {
            try {
                using (SQLiteConnection con = new SQLiteConnection(cnx)) {
                    con.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(con)) {
                        cmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name='wallet';";
                        if (cmd.ExecuteScalar() == null) {
                            cmd.CommandText = "CREATE TABLE wallet(w_id TEXT NOT NULL PRIMARY KEY, w_asset TEXT NOT NULL, w_free FLOAT NOT NULL, w_locked FLOAT NOT NULL, w_avgInBTC FLOAT NOT NULL, w_date NUMERIC NOT NULL, w_site TEXT NOT NULL)";
                            cmd.ExecuteNonQuery();
                            return true;
                        } else {
                            cmd.CommandText = "SELECT w_date FROM wallet ORDER BY w_date DESC LIMIT 1";
                            SQLiteDataReader reader = cmd.ExecuteReader();
                            if (reader.Read()) {
                                if (reader.GetString(0) == DateTime.Today.ToString("yyyy-MM-dd")) {
                                    Program.WriteLog("Overwrite today's data");
                                    reader.Close();
                                    DeleteTodayLines();
                                    return true;
                                }
                            }
                            reader.Close();
                            return true;
                        }
                    }
                }
            }
            catch (Exception e) {
                Program.WriteLog("Exception VerifDB()", e.Message + e.StackTrace);
                return false;
            }
        }

        public static void InsertAsset(Balance balance) {
            try {
                using (SQLiteConnection con = new SQLiteConnection(cnx)) {
                    con.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(con)) {
                        cmd.CommandText = "INSERT INTO wallet(w_id, w_asset, w_free, w_locked, w_avgInBTC, w_date, w_site) VALUES (@id, @asset, @free, @locked, @avgInBTC, @date, @site)";
                        cmd.Parameters.AddWithValue("@id", $"{balance.Asset}_{balance.Site}_{DateTime.Today:yyyyMMdd}");
                        cmd.Parameters.AddWithValue("@asset", balance.Asset);
                        cmd.Parameters.AddWithValue("@free", balance.Free);
                        cmd.Parameters.AddWithValue("@locked", balance.Locked);
                        cmd.Parameters.AddWithValue("@avgInBTC", balance.AvgInBTC);
                        cmd.Parameters.AddWithValue("@date", DateTime.Today.ToString("yyyy-MM-dd"));
                        cmd.Parameters.AddWithValue("@site", balance.Site);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception e) {
                Program.WriteLog("Exception InsertAsset()", e.Message + e.StackTrace);
                Console.ReadKey();
            }
        }

        private static void DeleteTodayLines() {
            try {
                using (SQLiteConnection con = new SQLiteConnection(cnx)) {
                    con.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(con)) {
                        cmd.CommandText = $"DELETE FROM wallet WHERE w_date ='{DateTime.Today:yyyy-MM-dd}'";
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception e) {
                Program.WriteLog("Exception DeleteTodayLines()", e.Message + e.StackTrace);
                Console.ReadKey();
            }
        }
    }
}
