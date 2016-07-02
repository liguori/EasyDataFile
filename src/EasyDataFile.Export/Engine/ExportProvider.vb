Imports System.Linq
Imports System.Reflection

Namespace Engine
    ''' <summary>Classe base che rappresenta un provider di report</summary>
    Public MustInherit Class ExportProvider
        ''' <summary>Indica se l'intestazione deve essere stampata automaticamente quando cambia l'oggetto che viene stampatao ogni nuovo record</summary>
        Public Property AutoPrintHeader As Boolean = True
        ''' <summary>Contiene l'elenco delle proprietà stampate per il record precedente</summary>
        Private Property PreviousRecord As New List(Of ExportField)
        ''' <summary>Buffer di appoggio per memorizzare i dati prima che vengano effettivamente renderizzati sul report</summary>
        Private Property DataBuffer As New List(Of List(Of ExportField))
        ''' <summary>Indica se il primo header per il report è stato stampato</summary>
        Private Property FistHeaderRendered As Boolean = False

        ''' <summary>Implementa l'effettivo render su report dell'informazione passata</summary>
        ''' <param name="str">L'informazione da visualizzare</param>
        Public MustOverride Sub WriteField(str As Object)
        ''' <summary>Chiude e finalizza il report</summary>
        Public MustOverride Sub Close()
        Public MustOverride Sub RecordOpen()
        Public MustOverride Sub RecordClose()

        ''' <summary>Verifica se le liste di proprietà fanno riferimento allo stesso ordine e stessa struttura </summary>
        ''' <param name="lst1">La prima lista</param>
        ''' <param name="lst2">La seconda lista</param>
        Private Function SameFieldList(lst1 As List(Of ExportField), lst2 As List(Of ExportField)) As Boolean
            If lst1.Count <> lst2.Count Then Return False
            For j As Integer = 0 To lst1.Count - 1
                If Not (lst1(j).Name = lst2(j).Name And lst1(j).DefinitionAttribute.Header = lst2(j).DefinitionAttribute.Header And lst1(j).DefinitionAttribute.Order = lst2(j).DefinitionAttribute.Order) Then Return False
            Next
            Return True
        End Function

        ''' <summary>Finalizza un record scrivendo il buffer e la testata sul file</summary>
        Sub FinalizeRecord()
            If AutoPrintHeader And Not SameFieldList(DataBuffer(0), PreviousRecord) Then
                RecordOpen()
                For Each field In DataBuffer(0)
                    WriteField(If(field.DefinitionAttribute.ShowValueInHeader, field.Value, field.DefinitionAttribute.Header))
                Next
                RecordClose()
                PreviousRecord = DataBuffer(0)
                FistHeaderRendered = True
            End If

            For Each ele In DataBuffer
                RecordOpen()
                For Each field In ele
                    WriteField(field.Value)
                Next
                RecordClose()
            Next
            DataBuffer = New List(Of List(Of ExportField))
        End Sub


        ''' <summary>Scrive il record sul report, non finalizzandolo</summary>
        ''' <param name="rec">L'oggetto da scrivere nel report</param>
        Public Sub WriteInlineRecord(rec As Object)
            WriteInlineRecord(GetFields(rec))
        End Sub

        ''' <summary>Scrive il record sul report, non finalizzandolo</summary>
        ''' <param name="lstField">L'oggetto da scrivere nel report</param>
        Public Sub WriteInlineRecord(lstField As List(Of ExportField))
            For Each ele In lstField.OrderBy(Function(x) x.DefinitionAttribute.Order)
                If ele.Value IsNot Nothing AndAlso GetType(IList).IsAssignableFrom(ele.Value.GetType()) Then
                    For Each child In ele.Value
                        WriteInlineRecord(child)
                    Next
                Else
                    If DataBuffer.Count = 0 Then DataBuffer.Add(New List(Of ExportField))
                    DataBuffer.Last().Add(ele)
                End If
            Next
        End Sub

        ''' <summary>Scrive il record sul report, finalizzandolo</summary>
        ''' <param name="rec">L'oggetto da scrivere nel report</param>
        Public Sub WriteRecord(rec As List(Of ExportField))
            DataBuffer.Add(New List(Of ExportField))
            WriteInlineRecord(rec)
            FinalizeRecord()
        End Sub


        ''' <summary>Scrive il record sul report, finalizzandolo</summary>
        ''' <param name="rec">L'oggetto da scrivere nel report</param>
        Public Sub WriteRecord(rec As Object)
            DataBuffer.Add(New List(Of ExportField))
            WriteInlineRecord(rec)
            FinalizeRecord()
        End Sub

        ''' <summary>Scrive una lista di record sul report, finalizzandoli</summary>
        ''' <param name="rec">L'oggetto da scrivere nel report</param>
        Public Sub WriteRecord(Of T)(rec As IList(Of T))
            For Each ele In rec
                WriteRecord(ele)
            Next
        End Sub

        ''' <summary>Ottiene l'elenco delle proprietà stampabili su un report ordinate con i metadati</summary>
        ''' <param name="obj"></param>
        Private Function GetFields(obj As Object) As List(Of ExportField)
            Dim ris As New List(Of ExportField)
            Dim p = obj.GetType().GetProperties().Where(Function(x) x.GetCustomAttributes(GetType(ExportDefinitionAttribute), True).Length > 0)
            For Each e In p
                Dim pi As New ExportField()
                pi.DefinitionAttribute = e.GetCustomAttributes(GetType(ExportDefinitionAttribute), True).FirstOrDefault()
                pi.Name = e.Name
                pi.Value = e.GetValue(obj, Nothing)
                ris.Add(pi)
            Next
            Return ris
        End Function


        ''' <summary>Contiene i dettagli di una proprietà stampabile</summary>
        Class ExportField
            Sub New(iValue As Object, iDefinitionAttribute As ExportDefinitionAttribute)
                Me.Value = iValue
                Me.DefinitionAttribute = iDefinitionAttribute
            End Sub

            Sub New()

            End Sub

            Public Property DefinitionAttribute As ExportDefinitionAttribute
            Public Property Name As String
            Public Property Value As Object
        End Class

        ''' <summary>Copia un record in un altro, copiando i valori delle proprietà che hanno lo stesso nome</summary>
        ''' <param name="source">Il record sorgente</param>
        ''' <param name="desta">Il record destinazione</param>
        Public Shared Sub RecordCopy(source As Object, desta As Object)
            For Each p In source.GetType().GetProperties()
                Dim pd = desta.GetType().GetProperties.Where(Function(x) x.Name = p.Name).FirstOrDefault()
                If pd IsNot Nothing Then pd.SetValue(desta, p.GetValue(source, Nothing), Nothing)
            Next
        End Sub
    End Class
End Namespace