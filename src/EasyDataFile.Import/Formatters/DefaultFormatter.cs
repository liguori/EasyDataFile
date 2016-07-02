using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDataFile.Import.Formatters
{
    public class DefaultFormatter: IFormatter
    {
        public string Format(string val)
        {
            if (val == null) return val;
            return val.ToString().Trim();
        }
    }
}
