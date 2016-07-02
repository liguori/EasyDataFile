Namespace Engine
    Public Class CsvExportProvider
        Inherits ExportProvider
        Private Property sw As IO.StreamWriter

        Public Sub New(path As String)
            sw = New IO.StreamWriter(path, False, System.Text.Encoding.Default)
        End Sub

        Public Overrides Sub Close()
            sw.Close()
        End Sub

        Public Shared Function CreateFile(path As String) As ExportProvider
            Return New CsvExportProvider(path)
        End Function

        Public Overrides Sub RecordOpen()
            'do nothing
        End Sub

        Public Overrides Sub RecordClose()
            sw.Write(vbNewLine)
            sw.Flush()
        End Sub

        Public Overrides Sub WriteField(str As Object)
            Csv(sw, str)
        End Sub

        Public Shared Sub Csv(ByRef sw As IO.StreamWriter, ParamArray elencovalori() As Object)
            If elencovalori Is Nothing Then
                sw.Write(FormattaValoreCsv("") & ";")
                Exit Sub
            End If
            For Each v In elencovalori
                sw.Write(FormattaValoreCsv(v) & ";")
            Next
        End Sub

        ''' <summary>Formatta un valore per inserirlo in un file CSV</summary>
        ''' <param name="valore">Il valore da formattare</param>
        Public Shared Function FormattaValoreCsv(ByVal valore As Object) As String
            If valore Is Nothing Then
                Return ""
            Else
                Dim val = valore.ToString()
                Dim cifre = val.All(AddressOf Char.IsDigit)
                Dim ris = String.Empty
                If cifre And Not String.IsNullOrWhiteSpace(val) Then ris &= "="
                ris &= """" & valore.ToString().Replace("""", """""").Trim() & """" 'Aggiunge " ad inizio e fine stringa e sostituisce i caratteri " all'interno della stringa con un doppio ""
                Return ris
            End If
        End Function
    End Class
End Namespace
