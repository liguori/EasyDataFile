using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDataFile.Import.Formatters
{
    public class DateWithoutSlashFormatter : IFormatter
    {
        public string Format(string val)
        {
            return val.Replace("/", string.Empty);
        }
    }
}
