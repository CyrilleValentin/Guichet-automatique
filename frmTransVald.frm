VERSION 5.00
Begin VB.Form frmTransVald 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "transVal"
   ClientHeight    =   4920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   750
      Left            =   10320
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   10320
      Top             =   840
   End
   Begin VB.CommandButton cmd 
      Caption         =   "stop"
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
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
      Height          =   510
      Left            =   2760
      TabIndex        =   4
      Top             =   4080
      Width           =   1815
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
      Height          =   510
      Left            =   6480
      TabIndex        =   3
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox txtI 
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
      Height          =   555
      Left            =   2880
      TabIndex        =   2
      Top             =   2400
      Width           =   5415
   End
   Begin VB.TextBox txtPas 
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
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3240
      Width           =   5415
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8400
      Picture         =   "frmTransVald.frx":0000
      TabIndex        =   0
      Top             =   3360
      Width           =   255
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
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   6240
      TabIndex        =   11
      Top             =   1800
      Width           =   7485
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CONNEXION POUR TRASNFERT"
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
      Left            =   480
      TabIndex        =   9
      Top             =   360
      Width           =   9930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom Client:"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   480
      TabIndex        =   8
      Top             =   2400
      Width           =   2400
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Code:"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   480
      TabIndex        =   7
      Top             =   3120
      Width           =   1140
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
      Height          =   855
      Left            =   960
      TabIndex        =   6
      Top             =   3840
      Width           =   975
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
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   -2400
      TabIndex        =   5
      Top             =   1800
      Width           =   7485
   End
End
Attribute VB_Name = "frmTransVald"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
txtPas.PasswordChar = ""
Else
If Check1.Value = 0 Then
txtPas.PasswordChar = "*"
End If
End If
End Sub

Private Sub cmd_Click()
Timer1.Enabled = False
End Sub

Private Sub cmdRetour_Click()
Unload Me
frmMenuClient.Show
End Sub

Private Sub cmdValider_Click()
If txtI.Text = "" Then
MsgBox "Veuillez entrer le Nom d'Utilisateur!! "
txtI.SetFocus
Exit Sub
Else
If txtPas.Text = "" Then
MsgBox "Veuillez entrer le mot de passe!!"
txtPas.SetFocus
Exit Sub
Else
Call Login1
End If
End If
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
Timer3.Enabled = True
End Sub
Private Sub Login1()
Module4.getconnected
Dim rs As New ADODB.Recordset
rs.Open " Select * from TABLE_CLIENT Where Nom_Client='" & txtI.Text & "'", cnnnnn, adOpenStatic, adLockReadOnly
If rs.RecordCount < 1 Then
MsgBox "Nom invalide", vbCritical, "Connexion"
txtI.SetFocus
Exit Sub
Else
If txtPas = rs!Password Then
Unload Me
frmTransfert.Show
  'montant = rs!Montant_Client
  'nom = rs!Nom_Client
frmTransfert.txtNom.Text = rs!Nom_Client
frmTransfert.txtMontant_Client.Text = rs!Montant_Client
frmTransfert.txtNumCompte.Text = rs!Numero_de_Compte_Client


Exit Sub
Else
MsgBox " Password invalide", vbCritical, "Connexion"
txtPas.SetFocus
Exit Sub
End If
End If
Set rs = Nothing
End Sub

Private Sub Form1_Load()
Timer1.Enabled = True

End Sub



Private Sub Time_Timer()
Label4.Left = Label4.Left - 5
If Label4.Left <= -Label4.Width Then
Label4.Left = Me.Width
End If
End Sub

Private Sub Timer1_Timer()
Label5.Caption = Label5.Caption - 1
If Label5.Caption = "0" Then
Unload Me
frmMenuClient.Show
cmd_Click
txtI.Text = ""
txtPas.Text = ""
End If
End Sub

Private Sub Timer3_Timer()
Label4.Left = Label4.Left - 5
If Label4.Left <= -Label4.Width Then
Label4.Left = Me.Width
End If
Label3.Left = Label3.Left - 5
If Label3.Left <= -Label3.Width Then
Label3.Left = Me.Width
End If
End Sub
