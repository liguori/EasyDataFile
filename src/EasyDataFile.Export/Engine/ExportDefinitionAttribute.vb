Namespace Engine
    ''' <summary>Definisce un attributo applicabile a livello di proprietà per memorizzare informazioni aggiuntive riguardanti l'esportazione del campo</summary>
    <AttributeUsage(AttributeTargets.Property)>
    Public Class ExportDefinitionAttribute
        Inherits Attribute
        ''' <summary>L'ordine di stampa</summary>
        Public Property Order As Integer
        ''' <summary>L'intestazione per il campo</summary>
        Public Property Header As String
        ''' <summary>Indica se al posto dell'intestazione deve essere stampato il valore</summary>
        Public Property ShowValueInHeader As Boolean
        ''' <summary>Indica se visualizzare l'intestazione solo alla prima intestazione utile per il report, poi mostrerà sempre il valore</summary>
        Public Property ShowHeaderOnce As Boolean

        Sub New(Order As Integer, Header As String, ShowValueInHeader As Boolean, ShowHeaderOnce As Boolean)
            Me.Order = Order
            Me.Header = Header
            Me.ShowValueInHeader = ShowValueInHeader
            Me.ShowHeaderOnce = ShowHeaderOnce
        End Sub

    End Class
End Namespace
