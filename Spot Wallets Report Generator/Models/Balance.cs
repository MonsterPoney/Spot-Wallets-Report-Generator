namespace Spot_Wallets_Report_Generator.Models {
    internal class Balance {
        public string Asset { get; set; }
        public float Free { get; set; }
        public float Locked { get; set; }
        public float Freeze { get; set; }
        public float AvgInBTC { get; set; }
        public float AvgInUSDT { get; set; }
        public float AvgPrice { get; set; }
        public string AssetAvg { get; set; }
        public string Notes { get; set; }
        public string Site { get; set; }
    }
}
