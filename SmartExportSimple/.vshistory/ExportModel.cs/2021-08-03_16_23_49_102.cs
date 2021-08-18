using System;
using SmartExporter;

namespace SmartExportSimple
{
    public class ExportModel
    {
        [ExcelDisplayInfo("序号", 1)]
        public string OrderId { get; set; }
        [ExcelDisplayInfo("序号", 1)]
        public string MaterialName { get; set; }
        [ExcelDisplayInfo("序号", 1)]
        public string PiCi { get; set; }
        [ExcelDisplayInfo("序号", 1)]
        public string GuoHao { get; set; }
        [ExcelDisplayInfo("序号", 1)]
        public string BanCi { get; set; }
        [ExcelDisplayInfo("序号", 1)]
        public string TotalWeight { get; set; }
        [ExcelDisplayInfo("序号", 1)]
        public string DuoShu { get; set; }
        [ExcelDisplayInfo("序号", 1)]
        public string KuaiShu { get; set; }
        [ExcelDisplayInfo("序号", 1)]
        public string Date { get; set; }
    }
}
