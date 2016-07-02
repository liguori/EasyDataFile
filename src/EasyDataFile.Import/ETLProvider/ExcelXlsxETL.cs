using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EasyDataFile.Import.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDataFile.Import.ETLProvider
{
    public class ExcelXlsxETL : ISingleModelProviderETL
    {
        private class CellInformation
        {
            public string Name { get; set; }
            public string Value { get; set; }
            public int ColumnIndex { get; set; }
        }

        private List<CellInformation> CellsInformation { get; set; } = new List<CellInformation>();

        private int RowCount { get; set; }
        private int CurrentRowCount { get; set; }
        private int CurrentColumnCount { get; set; }

        public Type Model { get; set; }

        private string GetCellValue(WorkbookPart wbPart, Cell cell)
        {
            if (cell.DataType != null && cell.DataType == CellValues.SharedString && !string.IsNullOrWhiteSpace(cell.CellValue.InnerText))
            {
                var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (stringTable != null)
                {
                    return stringTable.SharedStringTable.ElementAt(int.Parse(cell.CellValue.InnerText)).InnerText;
                }
                else
                {
                    return cell.CellValue.InnerText;
                }
            }
            else
            {
                return cell?.CellValue?.InnerText;
            }
        }

        private int GetFileRowCount(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                bool openRowsElement = false;
                int fileRowCount = 0;
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(Row))
                    {
                        openRowsElement = !openRowsElement;
                        if (openRowsElement) fileRowCount++;
                    }
                }
                return fileRowCount;
            }
        }


        public void Import(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                this.CurrentRowCount = 0;
                this.CurrentColumnCount = 0;
                this.RowCount = GetFileRowCount(fileName);
                bool openRowsElement = false;
                while (reader.Read())
                {
                    try
                    {
                        if (reader.ElementType == typeof(Cell))
                        {
                            Cell cell = reader.LoadCurrentElement() as Cell;
                            if (this.CurrentRowCount == 1)
                            {
                                CellsInformation.Add(new CellInformation { Name = GetCellValue(workbookPart, cell), ColumnIndex = ExcelColumnNameToIndex(cell.CellReference) });
                            }
                            else
                            {
                                var c = CellsInformation.Where(x => x.ColumnIndex == ExcelColumnNameToIndex(cell.CellReference)).FirstOrDefault();
                                if (c != null) c.Value = GetCellValue(workbookPart, cell);
                            }
                            this.CurrentColumnCount++;
                        }

                        if (reader.ElementType == typeof(Row))
                        {
                            openRowsElement = !openRowsElement;

                            if (openRowsElement)
                            {
                                this.CurrentColumnCount = 0;
                                this.CurrentRowCount++;
                                this.CellsInformation.ForEach(x => x.Value = string.Empty);
                            }
                            else if (this.CurrentRowCount > 1)
                            {
                                object row = null;
                                Exception ex = null;
                                try
                                {
                                    row = Process();
                                }
                                catch (Exception e)
                                {
                                    ex = e;
                                }
                                RecordReadyMethod(this.CurrentRowCount, this.RowCount, row, ex);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        RecordReadyMethod(this.CurrentRowCount, this.RowCount,null, ex);
                    }
                }
            }
        }


        private static int ExcelColumnNameToIndex(string cellReference)
        {
            string columnName = string.Empty;
            for (int i = 0; i < cellReference.Length; i++)
            {
                if (char.IsLetter(cellReference[i]))
                    columnName += cellReference[i];
                else
                    break;
            }
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");
            columnName = columnName.ToUpperInvariant();
            int sum = 0;
            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }
            return sum;
        }



        private object Process()
        {
            var item = Activator.CreateInstance(Model);
            foreach (var p in Model.GetProperties())
            {
                var flr = (NameFieldAttribute)p.GetCustomAttributes(typeof(NameFieldAttribute), true).FirstOrDefault();
                var val = (ValidateAttribute)p.GetCustomAttributes(typeof(ValidateAttribute), true).FirstOrDefault();
                if (flr != null)
                {
                    var propName = string.IsNullOrEmpty(flr.Name) ? p.Name : flr.Name;
                    var cell = this.CellsInformation.Where(x => x.Name == propName).FirstOrDefault();

                    var value = cell?.Value;
                    if (val != null) val.Validate(propName, value);
                    InitializeProperty(item, p, flr, value);
                }
            }
            return item;
        }

        private void InitializeProperty(object item, System.Reflection.PropertyInfo p, NameFieldAttribute flr, string val)
        {
            string value = flr.Formatter.Format(val);
            p.SetValue(item, DefaultValueConverter()[p.PropertyType](value), null);
        }

        private Dictionary<Type, Func<string, object>> DefaultValueConverter()
        {
            return new Dictionary<Type, Func<string, object>> {
            { typeof(short),  (v) => {return Int16.Parse(v);}},
            { typeof(int),  (v) => {return Int32.Parse(v);}},
            { typeof(long),  (v) => {return Int64.Parse(v);}},
            { typeof(double),  (v) => {return Double.Parse(v);}},
            { typeof(DateTime),  (v) => {return  (
                                                  v.Contains("/")?
                                                  DateTime.ParseExact(v, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture) :
                                                  DateTime.FromOADate(double.Parse(v))
                                                  );}},
            { typeof( short?),  (v) => {return (string.IsNullOrWhiteSpace(v) ? null: new Nullable<Int16>(Int16.Parse(v)));}},
            { typeof( int?),  (v) => {return (string.IsNullOrWhiteSpace(v)? null: new Nullable<Int32>(Int32.Parse(v)));}},
            { typeof( long?),  (v) => {return (string.IsNullOrWhiteSpace(v)? null: new Nullable<Int64>(Int64.Parse(v)));}},
            { typeof( double?),  (v) => {return (string.IsNullOrWhiteSpace(v)? null: new Nullable<Double>(Double.Parse(v)));}},
            { typeof( DateTime?),  (v) => {return (string.IsNullOrWhiteSpace(v) || string.IsNullOrWhiteSpace(v.Trim(' ','0')) ?
                                                                null:
                                                                (
                                                                v.Contains("/")?
                                                                new DateTime?(DateTime.ParseExact(v, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture)) :
                                                                new DateTime?(DateTime.FromOADate(double.Parse(v)))
                                                                ));
                                                    }
            },
            { typeof(string),  (v) => {return v;}}
            };
        }

        public Action<int, int, Object, Exception> RecordReadyMethod { get; set; }
    }
}
