VERSION 5.00
Begin VB.Form frmClientSolde 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "ClientSolde"
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   ForeColor       =   &H0080FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdValider2 
      Caption         =   " Valider"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Interval        =   750
      Left            =   7920
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdS 
      Caption         =   "stop"
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2640
      Width           =   4095
   End
   Begin VB.TextBox txtId 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   540
      Left            =   2520
      TabIndex        =   2
      Top             =   1905
      Width           =   4095
   End
   Begin VB.CommandButton cmdValider 
      Caption         =   " Valider"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4920
      TabIndex        =   1
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton CmdRetour 
      Caption         =   "Retour"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   """Veuillez saisir votre mot de passe avant 60 secondes"""
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   1320
      Width           =   6495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "System"
         Size            =   39
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   960
      Left            =   720
      TabIndex        =   10
      Top             =   3240
      Width           =   960
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   """Veuillez saisir votre mot de passe avant 60 secondes"""
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   -4080
      TabIndex        =   9
      Top             =   1320
      Width           =   6495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CONNEXION CLIENT"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   1320
      TabIndex        =   6
      Top             =   240
      Width           =   5310
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Code:"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom Client:"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
End
Attribute VB_Name = "frmClientSolde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public montant As Double
Public nom As String
Private Sub cmdAnnuler_Click()
txtNom.Text = ""
txtPrénoms.Text = ""
txtMotdePasse.Text = ""
txtNom.SetFocus
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
txtPassword.PasswordChar = ""
Else
If Check1.Value = 0 Then
txtPassword.PasswordChar = "*"
End If
End If
End Sub

Private Sub cmdRetour_Click()
Unload Me
frmMenuClient.Show
End Sub

Private Sub cmdS_Click()
Timer2.Enabled = False
End Sub
Private Sub cmdValider2_Click()
If txtId.Text = "" Or txtPassword.Text = "" Then
    MsgBox "Veuillez entrer vos identifiants"
    txtId.SetFocus
Else
    login2

End If
    

End Sub
Private Sub cmdValider_Click()
If txtId.Text = "" Then
MsgBox "Veuillez entrer le Nom d'Utilisateur!! "
txtId.SetFocus
Exit Sub
Else
If txtPassword.Text = "" Then
MsgBox "Veuillez entrer le mot de passe!!"
txtPassword.SetFocus
Exit Sub
Else
Call Login
End If
End If
End Sub
Private Sub login2()
'Module1.getconnected
'Dim rs As New ADODB.Recordset
'rs.Open " Select * from TABLE_CLIENT Where Nom_Client='" & txtId.Text & "' and Password='" & txtPassword.Text & "'", cnn, adOpenStatic, adLockReadOnly
'rs.Open " Select * from TABLE_CLIENT Where Nom_Client='" & txtId.Text & "'", cnn, adOpenStatic, adLockReadOnly
strChaine = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BASE_DE_DONNEES_VB611.mdb;Persist Security Info=False"
Dim cn As New ADODB.Connection
cn.Open strChaine
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open " Select * from TABLE_CLIENT Where Nom_Client='" & txtId.Text & "' and Password='" & txtPassword.Text & "'", cn, adOpenStatic, adLockBatchOptimistic

'rs.Open "select * from table_client where Numero_de_Compte_Client = " & Val(txtNumCompte.Text), cn, adOpenStatic, adLockBatchOptimistic
MsgBox rs.RecordCount
End Sub

Private Sub Login()
Module1.getconnected
Dim rs As New ADODB.Recordset
rs.Open " Select * from TABLE_CLIENT Where Nom_Client='" & txtId.Text & "'", cnn, adOpenStatic, adLockReadOnly
'MsgBox rs.RecordCount
If rs.RecordCount < 1 Then
MsgBox "Nom invalide", vbCritical, "Connexion"
txtId.SetFocus
Exit Sub
Else
If txtPassword = rs!Password Then
'Unload Me
MsgBox "Cher(e) client le solde de votre compte est " & rs!Montant_Client, vbInformation, "Affichage du solde"
txtId.Text = ""
txtPassword.Text = ""
txtId.SetFocus
'If MsgBox("Cliquer sur Ok pour revenir au MenuClient", vbInformation, vbOKOnly) Then
'frmMenuClient.Show
'End If
montant = rs!Montant_Client
Exit Sub
Else
MsgBox " Password invalide", vbCritical, "Connexion"
txtPassword.SetFocus
Exit Sub
End If
End If
Set rs = Nothing
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
Label4.Left = Label4.Left - 5
If Label4.Left <= -Label4.Width Then
Label4.Left = Me.Width
End If
Label6.Left = Label6.Left - 5
If Label6.Left <= -Label6.Width Then
Label6.Left = Me.Width
End If
End Sub

Private Sub Timer2_Timer()
Label5.Caption = Label5.Caption - 1
If Label5.Caption = "0" Then
Unload Me
frmMenuClient.Show
cmdS_Click
txtId.Text = ""
txtPassword.Text = ""
End If
End Sub
