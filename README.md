# EasyDataFile
Easy to use .NET library to import/export data from fixed length, delimited records or Excel files in typed objects

Export data
--------
You have just to decorate your properties of the class with attributes like the following example:
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

- *Order*: Output order on a file
- *Header*: Title of the field in output file (ex. in Excel or text first row)
- *ShowValueInHeader*: Indicate if header (first row of the file or )  must contain the first value that appears in the relative column of the file
- *ShowHeaderOnce*: Indicate if header must be only on the first record of file and not every data record or list of the same record's data type

By default are implemented 2 type of export that are:

- Excel Open Document Format (*.xlsx)
- Comma Separated Value (*.csv)

You can implement your own just inheriting and implementing **ExportProvider** abstract class.

Finally you can export records on file in this way:
```csharp
var p1 = new Person { Name="Gianluigi", Surname="Liguori", Age=3, BirthDate = new DateTime(1990, 2, 1); };
var p2 = new Person { Name="Vincenzo", Surname="Chianese", Age=3, BirthDate = new DateTime(1990, 2, 1); };

var exp = new CsvExportProvider(@"C:\test.csv"); //Or use for Excel = new ExcelExportProvider(@"C:\test.xlsx");
exp.WriteRecord(p1);
exp.WriteRecord(p2);
```

you can also write in output of the same file, collection of object of the same data type:
```csharp
var lis = new List<Person>();
//List initialization
exp.WriteRecord(lis);
```
or record of differents data type (whose properties must be decorated with **ExportDefinition** attribute like in Person class)

```csharp
var otherObjectOfDifferentType = new DifferentType();
exp.WriteRecord(otherObjectOfDifferentType);
```
and remember to close the file stream that is open automatically with the **ExportProvider** constructor:

```csharp
exp.Close();
```

Import data
--------
**IMPORT FIXED LENGHT FILE RECCORD**
Here is a complete example that show you how to import a file with fixed lenght records 
```csharp
    [FixedLengthRecord()]
    [MultipleInlineRecord(1, 25, 48)]
    [RecordType(0, 1, "C", 71)]
    public class Car
    {
        [FieldFixedLength(1,10,false)]
        public string Menufacturer { get; set; }

        [FieldFixedLength(11, 10, false)]
        public string Name { get; set; }

        [FieldFixedLength(21, 4, false)]
        public int Age { get; set; }
    }

```
Used attributes are:

- **FixedLengthRecord**: Used for mark calss as container of the fixed lenght recod
- **MultipleInlineRecord**: (optional) Used for indicate that the fixed record is repeated multiple times per every file line, and you should specify as argument an array of index that indicates the beginning of every single record ripetition into the same file line
- **RecordType**: (optional) Sometimes files are composed of multiple definition of record per line that depends on a single value that record line assume, this attribute accept following parameters:
    - *Start* index of the single value position into the file
    - *Lenght* numbero of characters of the single value into the file
    - *Value* value that should assume for activate the record definition
    - *RecordLenght* the record structure lenght that this record type activate
- **FieldFixedLength** must be indicated for every properties that should be filled with values from file. Arguments are:
    - *Start*: index of the start position of the value into the file 
    - *Lenght*: numbero of characters of the value from the start into the file
    - *UseOffset*: indicate whether properties is inflfuenced by **MultipleInlineRecord** attribute

Using attributes with previous class declaration we sayd that this class must contains records that:
- Are Fixed Lenght
- There are 3 occurrences per every file line at position 1, 105 and 209
- Due that this hypothetical file is composed of different structures this class contain record definition just for line of the file that contains 'C' as first character (because RecordType attribute)

We can start import with these lines:
```csharp
var c = new TextETL();
c.AddModel(typeof(Car));
c.RecordReadyMethod = (currentRow, totalRows, record, ex) => { 
    //Here your logic code to manage record object filled with values from file
};
c.Import(@"C:\fileToImport.txt");
```

and of course you can specify also different Model for the file (that should be activated making properly usage of **RecordType** attribute (as explained previously), for example:
```csharp
    [FixedLengthRecord()]
    [RecordType(10, 5, "TABLE", 200)]
    public class Table
    {
        //Properties definition
    }
    
    [FixedLengthRecord()]
    [RecordType(0, 3, "SKY", 100)]
    public class Sky
    {
        //Properties definition
    }


    //And adding them to the import
    c.AddModel(typeof(Table));
    c.AddModel(typeof(Sky));
    
```

Finally a file like this will be importable:
![Full fixed lenght file example](https://raw.github.com/liguori/EasyDataFile/master/docs/fixedLenghtFullExample.png)


TO DO
--------
1. Add Unit Tests
2. Optimize Code Interfaces and Implementations
3. Add output by convention or by configuration avoiding attributes usage
4. Add final release on NuGet


Feel free to contribute and join the project
