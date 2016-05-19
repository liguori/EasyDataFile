# EasyDataFile
Easy to use .NET library to import/export data from fixed length, delimited records or Excel files

Export data
--------
You have just to decorate your class's properties with attributes like the following example:
```csharp
public class Person
{
    [ExportDefinition(1 , "Name of the person", false, false)]
    public string Name { get; set; }
    
    [ExportDefinition(2 , "Surname of the person", false, false)]
    public string Surname { get; set; }
    
    [ExportDefinition(3 , "Age of the person", false, false)]
    public int Age { get; set; }
    
    [ExportDefinition(4 , "Birthdate of the person", false, false)]
    public DateTime BirthDate { get; set; }

}   
```
Parameters of **ExportDefinition** attribute are in order:

- Order: Output order on a file
- Header: Title of the field in output file (ex. in Excel or text first row)
- ShowValueInHeader: Indicate if header (first row of the file or )  must contain the first value that appears in the relative column of the file
- ShowHeaderOnce: Indicate if header must be only on the first record of file and not every data record or list of the same record's data type

By default are implemented 2 type of export that are:

- Excel Open Document Format (*.xlsx)
- Comma Separated Value (*.csv)

You can implement your own just inheriting and implementing **ExportProvider** abstract class.

Finally you can export records on file in this way:
```csharp
var p1 = new Person { Name="Gianluigi", Surname="Liguori", Age=3, BirthDate = new DateTime(1990, 2, 1); };
var p2 = new Person { Name="Vincenzo", Surname="Chianese", Age=3, BirthDate = new DateTime(1990, 2, 1); };

var exp = new CsvExportProvider(@"C:\test.csv");
exp.WriteRecordv
exp.WriteRecord(p2);
exp.Close();
```

Import data
--------

TO DO
--------
1. Add Unit Tests
2. Optimize Code Interfaces and Implementations
3. Add output by convention or by configuration avoiding attributes usage
4. Add final release on NuGet


Feel free to contribute and join the project
