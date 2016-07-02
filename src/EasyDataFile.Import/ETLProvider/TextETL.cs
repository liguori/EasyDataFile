using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EasyDataFile.Import.Attributes;
using System.IO;

namespace EasyDataFile.Import.ETLProvider
{
    public class TextETL : IMultipleModelProviderETL
    {
        public List<Type> Models { get; set; }
        
        public void AddModel(Type t)
        {
            if (Models == null) Models = new List<Type>();
            Models.Add(t);
        }

        public void Import(string fileName)
        {
            long numRow = GetRowNumber(fileName);
            long currentRow = 0;
            using (var sr = new System.IO.StreamReader(fileName))
            {
                while (!sr.EndOfStream)
                {
                    string record = sr.ReadLine();
                    currentRow++;
                    IEnumerable<object> rows = null;
                    Exception ex = null;
                    try
                    {
                        rows= ProcessRow(record);
                        RecordReadyMethod(currentRow, numRow, rows, ex);
                    }
                    catch (Exception e)
                    {
                        ex = e;
                    }
                }
            }
        }


        long GetRowNumber(string fileName)
        {
            //Ottiene il numero di righe totali del file
            StreamReader txtRd = new StreamReader(fileName);
            int numeroRighe = 0;
            while (txtRd.EndOfStream == false)
            {
                txtRd.ReadLine();
                numeroRighe += 1;
            }
            txtRd.Close();
            return numeroRighe;
        }

        private IEnumerable<Object> ProcessRow(string record)
        {
            var items = new List<object>();
            foreach (var m in Models)
            {
                var rt = (RecordTypeAttribute)m.GetCustomAttributes(typeof(Attributes.RecordTypeAttribute), true).FirstOrDefault();
                var mult = (MultipleInlineRecordAttribute)m.GetCustomAttributes(typeof(Attributes.MultipleInlineRecordAttribute), true).FirstOrDefault();
                if (rt==null ||( record.Substring(rt.Start, rt.Lenght) == rt.Value && rt.RecordLenght == record.Length))
                {
                    int offsetNumber = mult == null ? 1 : mult.Offsets.Length;
                    for (int iOffset = 0; iOffset < offsetNumber; iOffset++)
                    {
                        var item = CreateInstance(m);
                        foreach (var p in m.GetProperties())
                        {
                            try
                            {
                                var flr = (FieldFixedLengthAttribute)p.GetCustomAttributes(typeof(Attributes.FieldFixedLengthAttribute), true).FirstOrDefault();
                                if (flr != null) InitializeProperty(item, p, flr, record, mult == null ? 0 : mult.Offsets[iOffset]);
                            }
                            catch (Exception)
                            {
                                throw;
                            }
                        }
                        items.Add(item);
                    }
                }
            }
            return items;
        }

        private object CreateInstance(Type m)
        {
            var item = Activator.CreateInstance(m);
            if (ObecjtCreated != null) ObecjtCreated(item);
            return item;
        }


        private void InitializeProperty(object item, System.Reflection.PropertyInfo p, FieldFixedLengthAttribute flr, string record,int offset)
        {
            string value = flr.Formatter.Format(record.Substring(flr.Start + (flr.UseOffset ? offset : 0), flr.Lenght));
            p.SetValue(item, DefaultValueConverter()[p.PropertyType](value), null);
        }


        private Dictionary<Type, Func<string, object>> DefaultValueConverter()
        {
            return new Dictionary<Type, Func<string, object>> {
            { typeof(Int16),  (v) => {return Int16.Parse(v);}},
            { typeof(Int32),  (v) => {return Int16.Parse(v);}},
            { typeof(Int64),  (v) => {return Int16.Parse(v);}},
            { typeof( Nullable<Int16>),  (v) => {return (v.Trim()==string.Empty? null: new Nullable<Int16>(Int16.Parse(v)));}},
            { typeof( Nullable<Int32>),  (v) => {return (v.Trim()==string.Empty? null: new Nullable<Int32>(Int32.Parse(v)));}},
            { typeof( Nullable<Int64>),  (v) => {return (v.Trim()==string.Empty? null: new Nullable<Int64>(Int64.Parse(v)));}},
            { typeof( Nullable<DateTime>),  (v) => {return (v.Trim(' ','0')==string.Empty? null: new Nullable<DateTime>(DateTime.ParseExact(v, "ddMMyyyy", System.Globalization.CultureInfo.InvariantCulture)));}},
            { typeof(string),  (v) => {return v;}}
            };
        }

        public Action<long, long, IEnumerable<Object>,Exception> RecordReadyMethod { get; set; }

        public Action<object> ObecjtCreated { get; set; }

    }
}
