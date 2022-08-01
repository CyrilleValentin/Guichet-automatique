VERSION 5.00
Begin VB.Form frmAdminConnex 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "AdminConnex"
   ClientHeight    =   3945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStop 
      Caption         =   "stop"
      Height          =   300
      Left            =   3480
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   750
      Left            =   1200
      Top             =   0
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   8640
      MaskColor       =   &H80000002&
      Picture         =   "frmAdminConnex.frx":0000
      TabIndex        =   4
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox txtMotdePasse 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataMember      =   "Command1"
      DataSource      =   "BaseCompte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2280
      Width           =   5655
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
      Left            =   2880
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdValider 
      Caption         =   "Valider"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   0
      Top             =   3000
      Width           =   1815
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
      Left            =   5400
      TabIndex        =   9
      Top             =   1560
      Width           =   7485
   End
   Begin VB.Label Label2 
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
      Left            =   -3360
      TabIndex        =   7
      Top             =   1560
      Width           =   7485
   End
   Begin VB.Label Label1 
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
      Left            =   480
      TabIndex        =   6
      Top             =   2880
      Width           =   960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CONNEXION ADMINISTRATEUR"
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
      Left            =   720
      TabIndex        =   5
      Top             =   480
      Width           =   8100
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
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
      Left            =   1755
      TabIndex        =   2
      Top             =   2280
      Width           =   1035
   End
End
Attribute VB_Name = "frmAdminConnex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = 1 Then
txtMotdePasse.PasswordChar = ""
Else
If Check1.Value = 0 Then
txtMotdePasse.PasswordChar = "*"
End If
End If
End Sub

Private Sub cmdRetour_Click()
txtMotdePasse.Text = ""
Unload Me
frmAccueil.Show
End Sub

Private Sub cmdStop_Click()
Timer1.Enabled = False
End Sub

Private Sub cmdValider_Click()
Dim Pass As String
Dim i As Integer
Pass = "admin120"
If (txtMotdePasse = Pass) Then
frmMenuAdmin.Show
cmdStop_Click
frmAdminConnex.Hide
txtMotdePasse.Text = ""
Else
txtMotdePasse.Text = ""
txtMotdePasse.SetFocus
MsgBox ("Erreur veuillez resaisir le Mot de passe")
End If
End Sub

Private Sub Form_Load()
Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
Label1.Caption = Label1.Caption - 1
If Label1.Caption = "0" Then
Unload Me
frmAccueil.Show

txtMotdePasse.Text = ""
End If
End Sub

Private Sub Timer2_Timer()
Label2.Left = Label2.Left - 5
If Label2.Left <= -Label2.Width Then
Label2.Left = Me.Width
End If
Label3.Left = Label3.Left - 5
If Label3.Left <= -Label3.Width Then
Label3.Left = Me.Width
End If
End Sub

