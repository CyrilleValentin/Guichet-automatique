Attribute VB_Name = "Module2"
Option Explicit

Public cnnn As ADODB.Connection
Public Sub getconnected()
Set cnnn = New ADODB.Connection
cnnn.CursorLocation = adUseClient
cnnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BASE_DE_DONNEES_VB611.mdb" & ";Persist Security Info= False;"
cnnn.Open
End Sub


