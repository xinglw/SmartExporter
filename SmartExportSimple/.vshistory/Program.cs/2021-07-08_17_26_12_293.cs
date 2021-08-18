using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SmartExporter;

namespace SmartExportSimple
{
    class Program
    {
        static void Main(string[] args)
        {
            List<ExportModel> list = new List<ExportModel>();
            for (int i = 0; i < 10; i++)
            {
                list.Add(new ExportModel()
                {
                    XuHao = "22222",
                    MaoZhong = 223.22,
                    PiZhong = 0.22,
                    JingZhong = 223,
                    WuLiao = "废铁",
                    DanWei = "中国",
                    BeiZhu = "没有"
                });
            }
           
            ExportUtil exportUtil = new ExportUtil();
            exportUtil.ProgressCompleted += (a, b) =>
            {
                if (!a)
                {
                    Console.WriteLine(b);
                }
            };
            exportUtil.ExportAsync(@"D:\test.xls", list);
            Console.ReadLine();
        }
    }
}
