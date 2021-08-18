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
                    OrderId = "123123123",
                    MaterialName = "废料",
                    PiCi="afaidfngdasiffjng",
                    GuoHao="3",
                    BanCi = "2",
                    TotalWeight=22222,
                    DuoShu="60",
                    KuaiShu="1223",
                    Date="2021-08-03"
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
            
            exportUtil.ExportAsync<ExportModel>(@"D:\test.xls", list,list.Sum(x=>x.TotalWeight));
            Console.ReadLine();
        }
    }
}
