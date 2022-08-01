VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRetrait 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Retrait"
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10590
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNumCompte 
      Height          =   615
      Left            =   9240
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdRetour 
      Caption         =   "Retour"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8640
      TabIndex        =   8
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txtMontant_Client 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   525
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2370
      Width           =   4515
   End
   Begin VB.CommandButton cmdEnregistrer 
      Caption         =   "Retirer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8640
      TabIndex        =   5
      Top             =   1620
      Width           =   1575
   End
   Begin VB.TextBox txtMontantRestant 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "MontantRestant"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3120
      Width           =   4500
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7680
      Top             =   4920
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmRetrait.frx":0000
      OLEDBString     =   $"frmRetrait.frx":009E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TABLE_CLIENT"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8640
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtMontantRetirer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   525
      Left            =   3480
      TabIndex        =   1
      Top             =   1680
      Width           =   4500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RETRAIT CLIENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   39
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1020
      Left            =   2040
      TabIndex        =   9
      Top             =   360
      Width           =   6825
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Montant_Client:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   2730
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "MontantRestant:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   2820
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Montant à Retirer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   3120
   End
End
Attribute VB_Name = "frmRetrait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnnuler_Click()
txtMontantRestant.Text = ""
txtMontant_Client.Text = ""
txtMontantRetirer.Text = ""
txtMontantRetirer.SetFocus
End Sub

Private Sub cmdEnregistrer_Click()
strChaine = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BASE_DE_DONNEES_VB611.mdb;Persist Security Info=False"
Dim cn As New ADODB.Connection
cn.Open strChaine
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "select * from table_client where Numero_de_Compte_Client = " & Val(txtNumCompte.Text), cn, adOpenStatic, adLockBatchOptimistic
'rs.Open "select * from table_client where Nom_Client = " & Val(txt_Client), cn, adOpenStatic, adLockBatchOptimistic
'rs.Open "select * from table_client where Montant_Client = " & Val(txtMontant_Client), cn, adOpenStatic, adLockBatchOptimistic
'If MsgBox("voulez vous vraiment?", vbCritical, "avertissement") = vbYes Then
If MsgBox("Voulez-Vous vraiment faire ce Retrait", vbCritical + vbYesNo, "avertissement") = vbYes Then
    rs!Montant_Client = Val(txtMontantRestant.Text)
        rs.UpdateBatch
        MsgBox " Retrait effectué avec succès", vbInformation
        txtMontantRetirer.Text = ""
        txtMontantRestant.Text = ""
        txtMontantRetirer.SetFocus
        txtMontant_Client.Text = rs!Montant_Client
 End If
End Sub
Private Sub cmdRetour_Click()
Unload Me
frmMenuClient.Show
End Sub
Private Sub txtMontantRetirer_Change()
If Val(txtMontantRetirer) > Val(txtMontant_Client) Then
  MsgBox "Votre solde ne vous permet pas d'effectuer ce retrait", vbCritical
  txtMontantRetirer.Text = ""
  Else
   txtMontantRestant = Val(txtMontant_Client) - Val(txtMontantRetirer)
  End If
End Sub

Private Sub txtMontantRetirer_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
 KeyAscii = KeyAscii
 Else
  KeyAscii = 0
 End If
End Sub
