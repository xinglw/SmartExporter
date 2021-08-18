using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartExporter
{
    public class FileHelper
    {
        public static void CheckPath(string filePath)
        {
            var dic = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(dic))
            {
                Directory.CreateDirectory(dic);
            }
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
    }
}
