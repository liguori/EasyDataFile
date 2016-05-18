using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SmartETL.Attributes
{
     [AttributeUsage(AttributeTargets.Class)]
    public class MultipleInlineRecordAttribute : Attribute
    {
         public MultipleInlineRecordAttribute( params int[] offsets)
         {
             this.Offsets = offsets;
         }

         public int[] Offsets { get; set; }
    }
}
