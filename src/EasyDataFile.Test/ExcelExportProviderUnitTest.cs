using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EasyDataFile.Export.Engine;
using EasyDataFile.Import.ETLProvider;
using EasyDataFile.Import.Attributes;
using System.IO;
using EasyDataFile.Import.Formatters;

namespace EasyDataFile.Test
{
    [TestClass]
    public class ExcelExportProviderUnitTest
    {
        [TestMethod]
        public void TestSimpleObjectExport()
        {
            const string FileNotCreatedMessage = "File was not created";
            const string ExpectedValuesNotFindedMessage = "Unable to find the same exported values";
            const string fileName = "testSimpleObjectExport.xlsx";
            const string nameToWrite = "Gianluigi";
            const string surnameToWrite = "Liguori";
            DateTime birthDateToWrtite = new DateTime(1990, 2, 1);

            File.Delete(fileName);
            var f = new ExcelExportProvider(fileName);
            f.WriteRecord(new Person { Name = nameToWrite, Surname = surnameToWrite, BirthDate = birthDateToWrtite });
            f.Close();

            Assert.IsTrue(File.Exists(fileName), FileNotCreatedMessage);

            var c = new ExcelXlsxETL();
            c.Model = typeof(Person);
            c.RecordReadyMethod = (currentRow, totalRows, record, ex) =>
              {
                  if (ex != null) throw ex;
                  Assert.AreEqual(nameToWrite, ((Person)record).Name, ExpectedValuesNotFindedMessage);
                  Assert.AreEqual(surnameToWrite, ((Person)record).Surname, ExpectedValuesNotFindedMessage);
                  Assert.AreEqual(birthDateToWrtite, ((Person)record).BirthDate, ExpectedValuesNotFindedMessage);
              };
            c.Import(fileName);
        }

        class Person
        {
            [NameField("Nome")]
            [ExportDefinition(10,"Nome",false,true)]
            public string Name { get; set; }

            [NameField("Cognome")]
            [ExportDefinition(10,"Cognome",false,true)]
            public string Surname { get; set; }

            [NameField("DataNascita", typeof(DateWithSlashDDMMYYYFormatter))]
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
