Attribute VB_Name = "Module4"
Option Explicit

Public cnnnnn As ADODB.Connection
Public Sub getconnected()
Set cnnnnn = New ADODB.Connection
cnnnnn.CursorLocation = adUseClient
cnnnnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BASE_DE_DONNEES_VB611.mdb" & ";Persist Security Info= False;"
cnnnnn.Open
End Sub
