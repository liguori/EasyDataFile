using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDataFile.Import.Formatters
{
    public class DateWithSlashDDMMYYYFormatter : IFormatter
    {
        string IFormatter.Format(string val)
        {
            return new DateTime(int.Parse(val.Substring(6, 4)), int.Parse(val.Substring(3, 2)), int.Parse(val.Substring(0, 2))).ToShortDateString();
        }
    }
}
