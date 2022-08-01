VERSION 5.00
Begin VB.Form frmMenuClient 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "MenuClient"
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTrans 
      Caption         =   "Transfert"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4800
      TabIndex        =   4
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton cmdSolde 
      Caption         =   "Solde"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2760
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdRetour 
      Caption         =   "Retour"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7560
      TabIndex        =   1
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdRetrait 
      BackColor       =   &H0000FFFF&
      Caption         =   "Retrait"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   480
      TabIndex        =   0
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MENU CLIENT"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   39
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1305
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   5760
   End
End
Attribute VB_Name = "frmMenuClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRetour_Click()
Unload Me
frmAccueil.Show
End Sub

Private Sub cmdRetrait_Click()
Unload Me
frmRetraitValid.Show
frmRetraitValid.Timer2.Enabled = True
End Sub

Private Sub cmdSolde_Click()
Unload Me
frmClientSolde.Show
frmClientSolde.Timer2.Enabled = True
End Sub

Private Sub cmdTrans_Click()
Unload Me
frmTransVald.Show
frmTransVald.Timer1.Enabled = True
End Sub


