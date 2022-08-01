VERSION 5.00
Begin VB.Form frmAccueil 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Accueil"
   ClientHeight    =   5565
   ClientLeft      =   3600
   ClientTop       =   5340
   ClientWidth     =   13140
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   13140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton cmdQuitter 
      Caption         =   "Quitter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9840
      TabIndex        =   2
      Top             =   3960
      Width           =   2655
   End
   Begin VB.CommandButton cmdClient 
      Caption         =   "Client"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5520
      Picture         =   "frmAccueil.frx":0000
      TabIndex        =   1
      Top             =   3960
      Width           =   3375
   End
   Begin VB.CommandButton cmdAdmin 
      Caption         =   "Administrateur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   600
      TabIndex        =   0
      Top             =   3960
      Width           =   3975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   """Veuillez cliquer sur une commande pour continuer"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   7920
      TabIndex        =   6
      Top             =   2520
      Width           =   10110
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CHOIX DE L'UTILISATEUR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   41.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1035
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   10380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   """Veuillez cliquer sur une commande pour continuer"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   -3960
      TabIndex        =   4
      Top             =   2520
      Width           =   10110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3120
      TabIndex        =   3
      Top             =   4440
      Width           =   180
   End
End
Attribute VB_Name = "frmAccueil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdmin_Click()
Unload Me
frmAdminConnex.Show
End Sub

Private Sub cmdClient_Click()
Unload Me
frmMenuClient.Show
End Sub

Private Sub cmdQuitter_Click()
If MsgBox("Voulez vous vraiment quitter?", vbQuestions + vbYesNo, Quitter) = vbYes Then
End
End If
End Sub

Private Sub Timer1_Timer()
Label1.Left = Label1.Left - 5
If Label1.Left <= -Label1.Width Then
Label1.Left = Me.Width
End If
Label2.Left = Label2.Left - 5
If Label2.Left <= -Label2.Width Then
Label2.Left = Me.Width
End If
End Sub
