using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDataFile.Import.Formatters
{
    public interface IFormatter
    {
        string Format(string val);
    }
}
