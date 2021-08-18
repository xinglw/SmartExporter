using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartExporter
{
    public class TypeCastErrorException:Exception
    {
        public TypeCastErrorException(string message):base(message)
        {
           
        }
    }
}
