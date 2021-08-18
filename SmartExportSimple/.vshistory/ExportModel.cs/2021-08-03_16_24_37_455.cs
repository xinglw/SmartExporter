using System;
using SmartExporter;

namespace SmartExportSimple
{
    public class ExportModel
    {
        [ExcelDisplayInfo("订单号", 1)]
        public string OrderId { get; set; }
        [ExcelDisplayInfo("物料", 1)]
        public string MaterialName { get; set; }
        [ExcelDisplayInfo("批号", 1)]
        public string PiCi { get; set; }
        [ExcelDisplayInfo("锅号", 1)]
        public string GuoHao { get; set; }
        [ExcelDisplayInfo("班号", 1)]
        public string BanCi { get; set; }
        [ExcelDisplayInfo("总重量", 1)]
        public string TotalWeight { get; set; }
        [ExcelDisplayInfo("垛数", 1)]
        public string DuoShu { get; set; }
        [ExcelDisplayInfo("块数", 1)]
        public string KuaiShu { get; set; }
        [ExcelDisplayInfo("生产日期", 1)]
        public string Date { get; set; }
    }
}
