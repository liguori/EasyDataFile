Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Office2010.ExcelAc
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace Engine
    Public Class ExcelExportProvider
        Inherits ExportProvider

        Private Property WBPath As String
        Private Property document As SpreadsheetDocument
        Private Property writer As OpenXmlWriter
        Private Property attributes As List(Of OpenXmlAttribute)
        Private Property workSheetPart As WorksheetPart
        Private Property CurrentRow As Integer = 1
        Private Property CurrentColumn As Integer = 1


        Public Sub New(path As String)
            Me.WBPath = path
            Me.document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook)
            document.AddWorkbookPart()
            Me.workSheetPart = document.WorkbookPart.AddNewPart(Of WorksheetPart)
            Me.writer = OpenXmlWriter.Create(workSheetPart)
            writer.WriteStartElement(New Worksheet())
            writer.WriteStartElement(New SheetData())
        End Sub

        Public Overrides Sub Close()
            ' write the end SheetData element
            writer.WriteEndElement()
            ' write the end Worksheet element
            writer.WriteEndElement()
            writer.Close()

            writer = OpenXmlWriter.Create(document.WorkbookPart)
            writer.WriteStartElement(New Workbook())
            writer.WriteStartElement(New Sheets())

            writer.WriteElement(New Sheet() With
            {
                .Name = "Large Sheet",
                .SheetId = 1,
                .Id = document.WorkbookPart.GetIdOfPart(workSheetPart)
            })

            'End Sheets
            writer.WriteEndElement()
            'End Workbook
            writer.WriteEndElement()

            writer.Close()
            document.Close()
            document.Dispose()
        End Sub

        Public Overrides Sub RecordOpen()
            'create a New list of attributes
            attributes = New List(Of OpenXmlAttribute)()
            'add the row index attribute to the list
            attributes.Add(New OpenXmlAttribute("r", Nothing, CurrentRow.ToString()))
            'write the row start element with the row index attribute
            writer.WriteStartElement(New Row(), attributes)
        End Sub

        Public Overrides Sub RecordClose()
            ' write the end row element
            writer.WriteEndElement()
            CurrentRow += 1
            CurrentColumn = 1
        End Sub

        Public Overrides Sub WriteField(str As Object)
            ExcelRow(str)
        End Sub


        Function GetDataType(ByVal ele As Object) As String
            If TypeOf ele Is Integer Or TypeOf ele Is Double Or TypeOf ele Is Long Then
                Return "n"
            Else
                Return "str"
            End If
        End Function

        Public Sub ExcelRow(ParamArray elencovalori() As Object)
            Dim columnNum As Integer
            For columnNum = CurrentColumn To CurrentColumn + elencovalori.Count - 1
                Dim value = elencovalori(columnNum - CurrentColumn)

                'reset the list of attributes
                attributes = New List(Of OpenXmlAttribute)

                ' add data type attribute - in this case inline string (you might want to look at the shared strings table)
                attributes.Add(New OpenXmlAttribute("t", Nothing, GetDataType(value)))

                'add the cell reference attribute
                attributes.Add(New OpenXmlAttribute("r", "", String.Format("{0}{1}", GetColumnName(columnNum), CurrentRow)))

                'write the cell start element with the type And reference attributes
                writer.WriteStartElement(New Cell(), attributes)
                'write the cell value
                writer.WriteElement(New CellValue(Regex.Replace(value, "[^\u0020-\u007E]", String.Empty)))

                'write the end cell element
                writer.WriteEndElement()
            Next
            CurrentColumn = columnNum
        End Sub



        'A simple helper to get the column name from the column index. This is not well tested!
        Private Shared Function GetColumnName(columnIndex As Integer) As String
            Dim dividend As Integer = columnIndex
            Dim columnName As String = [String].Empty
            Dim modifier As Integer

            While dividend > 0
                modifier = (dividend - 1) Mod 26
                columnName = Convert.ToChar(65 + modifier).ToString() & columnName
                dividend = CInt((dividend - modifier) / 26)
            End While

            Return columnName
        End Function


    End Class
End Namespace
