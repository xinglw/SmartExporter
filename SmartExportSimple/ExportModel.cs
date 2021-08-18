using System;
using SmartExporter;

namespace SmartExportSimple
{
    public class ExportModel
    {
        [ExcelDisplayInfo("订单号", 1)]
        public string OrderId { get; set; }
        [ExcelDisplayInfo("物料", 2)]
        public string MaterialName { get; set; }
        [ExcelDisplayInfo("批号", 3)]
        public string PiCi { get; set; }
        [ExcelDisplayInfo("锅号", 4)]
        public string GuoHao { get; set; }
        [ExcelDisplayInfo("班号", 5)]
        public string BanCi { get; set; }
        [ExcelDisplayInfo("总重量", 6)]
        public double TotalWeight { get; set; }
        [ExcelDisplayInfo("垛数", 7)]
        public int DuoShu { get; set; }
        [ExcelDisplayInfo("块数", 8)]
        public int KuaiShu { get; set; }
        [ExcelDisplayInfo("生产日期", 9)]
        public string Date { get; set; }
    }
}
