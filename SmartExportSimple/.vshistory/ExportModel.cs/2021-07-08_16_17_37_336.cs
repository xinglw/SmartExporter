using System;
using SmartExporter;

namespace SmartExportSimple
{
    public class ExportModel
    {
        [ExcelDisplayInfo("序号",1)]
        public string XuHao { get; set; }
        [ExcelDisplayInfo("物料",2)]
        public string WuLiao { get; set; }
        [ExcelDisplayInfo("皮重",3)]
        public double PiZhong { get; set; }
        [ExcelDisplayInfo("毛重",4)]
        public double MaoZhong { get; set; }
        [ExcelDisplayInfo("净重",5)]
        public double JingZhong { get; set; }
        [ExcelDisplayInfo("单位",6)]
        public string DanWei { get; set; }
        [ExcelDisplayInfo("过磅时间",9)]
        public DateTime WeightTime { get; set; } = DateTime.Now;
        [ExcelDisplayInfo("备注",8)]
        public string BeiZhu { get; set; }
        [ExcelDisplayInfo("司磅员",7)]
        public string SiBangYuan { get; set; } = "admin";
    }
}
