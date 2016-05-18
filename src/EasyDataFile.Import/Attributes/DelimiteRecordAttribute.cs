using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SmartETL.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public class DelimiteRecordAttribute : Attribute
    {

        public DelimiteRecordAttribute(string delimiter)
        {

        }
    }
}
