Attribute VB_Name = "Module1"
Option Explicit

Public cnn As ADODB.Connection
Public Sub getconnected()
Set cnn = New ADODB.Connection
cnn.CursorLocation = adUseClient
cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BASE_DE_DONNEES_VB611.mdb" & ";Persist Security Info= False;"
cnn.Open
End Sub

