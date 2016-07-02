using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NExcel;
using EasyDataFile.Import.Attributes;

namespace EasyDataFile.Import.ETLProvider
{
    public class ExcelXlsETL : ISingleModelProviderETL
    {
        public Type Model { get; set; }


        public void Import(string fileName)
        {
            var wk = Workbook.getWorkbook(fileName);
            int numRow = wk.Sheets[0].Rows;
            int currentRow = 0;
            for (int i = 1; i < wk.Sheets[0].Rows; i++)
            {
                object row = null;
                Exception ex = null;
                currentRow++;
                try
                {
                    row = Process(wk.Sheets[0], i);
                }
                catch (Exception e)
                {
                    ex = e;
                }
                RecordReadyMethod(currentRow, numRow, row, ex);
            }
        }


        private object Process(Sheet sh, int i)
        {
            var item = Activator.CreateInstance(Model);
            foreach (var p in Model.GetProperties())
            {
                var flr = (NameFieldAttribute)p.GetCustomAttributes(typeof(NameFieldAttribute), true).FirstOrDefault();
                var val = (ValidateAttribute)p.GetCustomAttributes(typeof(ValidateAttribute), true).FirstOrDefault();
                if (flr != null)
                {
                    var propName = string.IsNullOrEmpty(flr.Name) ? p.Name : flr.Name;
                    var value = GetFileValue(sh, i, propName);
                    if (val != null) val.Validate(propName, value);
                    InitializeProperty(item, p, flr, value);
                }
            }
            return item;
        }


        private string GetFileValue(Sheet sh, int i, string name)
        {
            for (int j = 0; j < sh.Columns; j++)
            {
                if (!string.IsNullOrWhiteSpace(sh.getCell(j, 0).Contents))
                {
                    if (sh.getCell(j, 0).Contents.Trim().ToUpper() == name.Trim().ToUpper())
                    {
                        if (sh.getCell(j, i).Value != null && sh.getCell(j, i).Value is DateTime)
                        {
                            return DateTime.Parse(sh.getCell(j, i).Value.ToString()).ToShortDateString();
                        }
                        else
                        {
                            return sh.getCell(j, i).Contents;
                        }
                    }
                }
            }
            return null;
        }

        private void InitializeProperty(object item, System.Reflection.PropertyInfo p, NameFieldAttribute flr, string val)
        {
            string value = flr.Formatter.Format(val);
            p.SetValue(item, DefaultValueConverter()[p.PropertyType](value), null);
        }

        private Dictionary<Type, Func<string, object>> DefaultValueConverter()
        {
            return new Dictionary<Type, Func<string, object>> {
            { typeof(Int16),  (v) => {return Int16.Parse(v);}},
            { typeof(Int32),  (v) => {return Int32.Parse(v);}},
            { typeof(Int64),  (v) => {return Int64.Parse(v);}},
            { typeof(Double),  (v) => {return Double.Parse(v);}},
            { typeof( Nullable<Int16>),  (v) => {return (string.IsNullOrWhiteSpace(v) ? null: new Nullable<Int16>(Int16.Parse(v)));}},
            { typeof( Nullable<Int32>),  (v) => {return (string.IsNullOrWhiteSpace(v)? null: new Nullable<Int32>(Int32.Parse(v)));}},
            { typeof( Nullable<Int64>),  (v) => {return (string.IsNullOrWhiteSpace(v)? null: new Nullable<Int64>(Int64.Parse(v)));}},
            { typeof( Nullable<Double>),  (v) => {return (string.IsNullOrWhiteSpace(v)? null: new Nullable<Double>(Double.Parse(v)));}},
            { typeof( Nullable<DateTime>),  (v) => {return (string.IsNullOrWhiteSpace(v) || v.Trim(' ','0')==string.Empty? null: new Nullable<DateTime>(DateTime.ParseExact(v, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture)));}},
            { typeof(string),  (v) => {return v;}}
            };
        }

        public Action<int, int, Object, Exception> RecordReadyMethod { get; set; }
    }
}
