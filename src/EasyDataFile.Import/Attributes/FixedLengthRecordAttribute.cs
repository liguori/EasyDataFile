using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDataFile.Import.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public class FixedLengthRecordAttribute : Attribute
    {

        public FixedLengthRecordAttribute()
        {

        }
    }
}
