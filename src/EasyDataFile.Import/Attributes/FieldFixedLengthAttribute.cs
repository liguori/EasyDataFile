using SmartETL.Formatters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SmartETL.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class FieldFixedLengthAttribute : Attribute
    {

        public FieldFixedLengthAttribute(int lenght,Type formatter, bool useoffset = true)
        {
            this.Lenght = lenght;
            this.UseOffset = useoffset;
            this.Formatter = (IFormatter)Activator.CreateInstance(formatter);
        }

        public FieldFixedLengthAttribute(int lenght, bool useoffset = true)
        {
            this.Lenght = lenght;
            this.UseOffset = useoffset;
            this.Formatter = new DefaultFormatter();
        }

        public FieldFixedLengthAttribute(int start, int lenght, Type formatter, bool useoffset = true)
        {
            this.Start = start;
            this.Lenght = lenght;
            this.UseOffset = useoffset;
            this.Formatter = (IFormatter)Activator.CreateInstance(formatter);
        }

        public FieldFixedLengthAttribute(int start, int lenght, bool useoffset = true)
        {
            this.Start = start;
            this.Lenght = lenght;
            this.UseOffset = useoffset;
            this.Formatter = new DefaultFormatter();
        }

        public int Start { get; set; }

        public int Lenght { get; set; }

        public bool UseOffset { get; set; }

        public IFormatter Formatter { get; set; }
    }
}
