using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SmartETL.Formatters
{
    public interface IFormatter
    {
        string Format(string val);
    }
}
