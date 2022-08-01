VERSION 5.00
Begin VB.Form frmSolde 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Solde"
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   FillColor       =   &H0000FFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSolde.frx":0000
   ScaleHeight     =   2445
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMontant_Client 
      BackColor       =   &H00C0FFC0&
      DataField       =   "Montant_Client"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   2700
   End
   Begin VB.CommandButton cmdRetour 
      Caption         =   "Retour Au Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Montant_Client:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   390
      TabIndex        =   2
      Top             =   1080
      Width           =   1905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "SOLDE CLIENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   435
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   2790
   End
End
Attribute VB_Name = "frmSolde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRetour_Click()
frmClientSolde.Show
frmSolde.Hide
End Sub

Private Sub Form_Load()
txtMontant_Client.Text = frmClientSolde.montant


End Sub

