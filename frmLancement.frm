VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLancement 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12495
   ForeColor       =   &H00400000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   3960
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1931
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
      Max             =   105
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   12480
      Top             =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   765
      Left            =   7320
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.1"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   945
      Left            =   8280
      TabIndex        =   4
      Top             =   2280
      Width           =   2985
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TRANSFERT D'ARGENT"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   41.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1380
      Left            =   1320
      TabIndex        =   3
      Top             =   480
      Width           =   9735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   765
      Left            =   4200
      TabIndex        =   2
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   765
      Left            =   480
      TabIndex        =   0
      Top             =   3240
      Width           =   975
   End
End
Attribute VB_Name = "frmLancement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
Label1.Caption = "Chargement....."
Label2.Caption = ProgressBar1.Value & "%"
Label5.Caption = "Veuillez patienter....."
If ProgressBar1.Value = ProgressBar1.Max Then
Timer1.Enabled = False
Unload Me
frmAccueil.Show
End If
End Sub
