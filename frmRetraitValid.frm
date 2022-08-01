VERSION 5.00
Begin VB.Form frmRetraitValid 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "&H00FFFF00&"
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSt 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Interval        =   750
      Left            =   120
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   240
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8640
      Picture         =   "frmRetraitValid.frx":0000
      TabIndex        =   4
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txtPasswor1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   540
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3120
      Width           =   5535
   End
   Begin VB.TextBox txtIde1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   540
      Left            =   3000
      TabIndex        =   2
      Top             =   2280
      Width           =   5535
   End
   Begin VB.CommandButton cmdValider 
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
      Height          =   540
      Left            =   6960
      TabIndex        =   1
      Top             =   3840
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
      Height          =   540
      Left            =   3000
      TabIndex        =   0
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   """Veuillez saisir votre mot de passe avant 60 secondes"""
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   6120
      TabIndex        =   11
      Top             =   1680
      Width           =   7695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   """Veuillez saisir votre mot de passe avant 60 secondes"""
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   -3000
      TabIndex        =   9
      Top             =   1680
      Width           =   7695
   End
   Begin VB.Label Label5 
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
      Height          =   975
      Left            =   1320
      TabIndex        =   8
      Top             =   3720
      Width           =   1095
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
      Left            =   720
      TabIndex        =   7
      Top             =   3120
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
      Left            =   720
      TabIndex        =   6
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CONNEXION POUR RETRAIT"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1020
      Left            =   600
      TabIndex        =   5
      Top             =   240
      Width           =   8955
   End
End
Attribute VB_Name = "frmRetraitValid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
txtPasswor1.PasswordChar = ""
Else
If Check1.Value = 0 Then
txtPasswor1.PasswordChar = "*"
End If
End If
End Sub

Private Sub cmdRetour_Click()
Unload Me
frmMenuClient.Show
End Sub

Private Sub cmdSt_Click()
Timer2.Enabled = False
End Sub

Private Sub cmdValider_Click()
If txtIde1.Text = "" Then
MsgBox "Veuillez entrer le Nom d'Utilisateur!! "
txtIde1.SetFocus
Exit Sub
Else
If txtPasswor1.Text = "" Then
MsgBox "Veuillez entrer le mot de passe!!"
txtPasswor1.SetFocus
Exit Sub
Else
Call Login1
End If
End If
End Sub

Private Sub Login1()
Module3.getconnected
Dim rs As New ADODB.Recordset
rs.Open " Select * from TABLE_CLIENT Where Nom_Client='" & txtIde1.Text & "'", cnnnn, adOpenStatic, adLockReadOnly
If rs.RecordCount < 1 Then
MsgBox "Nom invalide", vbCritical, "Connexion"
txtIde1.SetFocus
Exit Sub
Else
If txtPasswor1 = rs!Password Then
Unload Me
frmRetrait.Show


  montant = rs!Montant_Client
frmRetrait.txtMontant_Client.Text = montant
frmRetrait.txtNumCompte.Text = rs!Numero_de_Compte_Client
Exit Sub
Else
MsgBox " Password invalide", vbCritical, "Connexion"
txtPasswor1.SetFocus
Exit Sub
End If
End If
Set rs = Nothing
End Sub
Private Sub Form_Load()
Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
Label4.Left = Label4.Left - 5
If Label4.Left <= -Label4.Width Then
Label4.Left = Me.Width
End If
Label3.Left = Label3.Left - 5
If Label3.Left <= -Label3.Width Then
Label3.Left = Me.Width
End If
End Sub

Private Sub Timer2_Timer()
Label5.Caption = Label5.Caption - 1
If Label5.Caption = "0" Then
Unload Me
frmMenuClient.Show
cmdSt_Click
txtIde1.Text = ""
txtPasswor1.Text = ""
End If
End Sub

