VERSION 5.00
Begin VB.Form frmAdminConnex2 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "AdminConnex2"
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStp 
      Caption         =   "stop"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   240
      Top             =   3720
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   240
      Top             =   240
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
      Left            =   2280
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
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
      Height          =   555
      Left            =   5640
      TabIndex        =   3
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtIde 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   525
      Left            =   2520
      TabIndex        =   2
      Top             =   2160
      Width           =   4935
   End
   Begin VB.TextBox txtPasswor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2880
      Width           =   4935
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   """Veuillez saisir votre mot de passe avant 60 secondes"""
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   6360
      TabIndex        =   11
      Top             =   1680
      Width           =   7125
   End
   Begin VB.Label Label4 
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
      Height          =   855
      Left            =   480
      TabIndex        =   9
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   """Veuillez saisir votre mot de passe avant 60 secondes"""
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   -2520
      TabIndex        =   8
      Top             =   1680
      Width           =   7125
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ADMINISTRATEUR SOLDE "
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   33
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1110
      Left            =   600
      TabIndex        =   7
      Top             =   360
      Width           =   8865
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
      TabIndex        =   6
      Top             =   2160
      Width           =   2175
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
      Top             =   2880
      Width           =   1035
   End
End
Attribute VB_Name = "frmAdminConnex2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = 1 Then
txtPasswor.PasswordChar = ""
Else
If Check1.Value = 0 Then
txtPasswor.PasswordChar = "*"
End If
End If
End Sub

Private Sub cmdRetour_Click()
Unload Me
frmMenuAdmin.Show
frmAdminConnex2.Timer2.Enabled = True
frmAdminConnex2.Label4.Caption = 20
End Sub

Private Sub cmdStp_Click()
Timer2.Enabled = False
End Sub

Private Sub cmdValider_Click()
If txtIde.Text = "" Then
MsgBox "Veuillez entrer le Nom d'Utilisateur!! "
txtIde.SetFocus
Exit Sub
Else
If txtPasswor.Text = "" Then
MsgBox "Veuillez entrer le mot de passe!!"
txtPasswor.SetFocus
Exit Sub
Else
Call Login1
End If
End If
End Sub

Private Sub Login1()
Module2.getconnected
Dim rs As New ADODB.Recordset
rs.Open " Select * from TABLE_CLIENT Where Nom_Client='" & txtIde.Text & "'", cnnn, adOpenStatic, adLockReadOnly
If rs.RecordCount < 1 Then
MsgBox "Nom invalide", vbCritical, "Connexion"
txtIde.SetFocus
Exit Sub
Else
If txtPasswor = rs!Password Then
Unload Me
MsgBox rs!Montant_Client
If MsgBox("Cliquer sur Ok pour revenir au MenuAdmin", vbOKOnly) Then
frmMenuAdmin.Show
End If
montant = rs!Montant_Client
Exit Sub
Else
MsgBox " Password invalide", vbCritical, "Connexion"
txtPasswor.SetFocus
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
Label3.Left = Label3.Left - 5
If Label3.Left <= -Label3.Width Then
Label3.Left = Me.Width
End If
Label5.Left = Label5.Left - 5
If Label5.Left <= -Label5.Width Then
Label5.Left = Me.Width
End If
End Sub

Private Sub Timer2_Timer()
Label4.Caption = Label4.Caption - 1
If Label4.Caption = "0" Then
Unload Me
frmMenuAdmin.Show
cmdStp_Click
txtIde.Text = ""
txtPasswor.Text = ""
End If
End Sub
