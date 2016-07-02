using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDataFile.Import.Formatters
{
    public class DateDD_MM_YYYYFormatter : IFormatter
    {
        public string Format(string val)
        {
            return val.Replace("/", string.Empty);
        }
    }
}
