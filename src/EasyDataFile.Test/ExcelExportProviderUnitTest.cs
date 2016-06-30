using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EasyDataFile.Export.Engine;

namespace EasyDataFile.Test
{
    [TestClass]
    public class ExcelExportProviderUnitTest
    {
        [TestMethod]
        public void TestSimpleObjectExport()
        {
            var f = new ExcelExportProvider("test.xlsx");
            f.WriteRecord(new Person { Name = "Gianluigi", Surname = "Liguori", BirthDate = new DateTime(1990, 2, 1) });
            f.Close();
        }

        class Person
        {
            [ExportDefinition(10,"Nome",false,true)]
            public string Name { get; set; }

            [ExportDefinition(10,"Cognome",false,true)]
            public string Surname { get; set; }

            [ExportDefinition(10, "DataNascita", false, true)]
            public DateTime BirthDate { get; set; }

            [ExportDefinition(10, "Eta", false, true)]
            public int Age { get
                {
                    return (new DateTime(1, 1, 1) + (DateTime.Now - BirthDate)).Year;
                }
            }
        }
    }
}
