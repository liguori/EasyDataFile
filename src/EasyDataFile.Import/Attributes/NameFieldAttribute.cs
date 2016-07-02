using EasyDataFile.Import.Formatters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDataFile.Import.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class NameFieldAttribute : Attribute
    {
        public NameFieldAttribute()
        {
            this.Formatter = new DefaultFormatter();
            this.Name = string.Empty;
        }

        public NameFieldAttribute(string name)
        {
            this.Formatter = new DefaultFormatter();
            this.Name = name;
        }

        public NameFieldAttribute(string name,Type formatter)
        {
            this.Formatter = (IFormatter)Activator.CreateInstance(formatter);
            this.Name = name;
        }

        public string Name { get; set; }
        public IFormatter Formatter { get; set; }
    }
}
