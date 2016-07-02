using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDataFile.Import.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class FieldConverterAttribute : Attribute
    {
    }
}
