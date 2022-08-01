Attribute VB_Name = "Module3"
Option Explicit

Public cnnnn As ADODB.Connection
Public Sub getconnected()
Set cnnnn = New ADODB.Connection
cnnnn.CursorLocation = adUseClient
cnnnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BASE_DE_DONNEES_VB611.mdb" & ";Persist Security Info= False;"
cnnnn.Open
End Sub
