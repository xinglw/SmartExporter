using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartExporter
{
    /// <summary>
    /// 用于表示Model中属性在Excel中显示的列名,exportIndex属性用于对导出的列进行排序
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelDisplayInfo:Attribute
    {
        public ExcelDisplayInfo(string name,int exportIndex)
        {
            this.name = name;
            this.exportIndex = exportIndex;
        }
        private string name;
        private int exportIndex;

        public string Name
        {
            get => name;
        }
        public int ExportIndex => exportIndex;
    }
}
