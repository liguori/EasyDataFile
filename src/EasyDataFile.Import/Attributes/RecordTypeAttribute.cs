using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDataFile.Import.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public class RecordTypeAttribute : Attribute
    {
        public RecordTypeAttribute(int start, int lenght, string value,int recordLeght=0)
        {
            this.Start = start;
            this.Lenght = lenght;
            this.Value = value;
            this.RecordLenght = recordLeght;
        }

        public int Start { get; set; }

        public int Lenght { get; set; }

        public string Value { get; set; }

        public int RecordLenght { get; set; }
    }
}
