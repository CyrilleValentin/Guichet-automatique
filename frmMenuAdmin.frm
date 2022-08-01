VERSION 5.00
Begin VB.Form frmMenuAdmin 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "MenuAdmin"
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13095
   FillColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   13095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdENRE 
      Caption         =   "Comptes"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   5400
      TabIndex        =   5
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton cmdRetour 
      Caption         =   "Retour"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   10920
      TabIndex        =   3
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdCreer 
      Caption         =   "Créer"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton cmdDepot 
      Caption         =   "Dépot"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   8520
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdSolde 
      Caption         =   "Solde"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2760
      TabIndex        =   0
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MENU ADMINISTRATEUR"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   33
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1110
      Left            =   2040
      TabIndex        =   4
      Top             =   600
      Width           =   8760
   End
End
Attribute VB_Name = "frmMenuAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCreer_Click()
frmCompte.Show
frmMenuAdmin.Hide
End Sub

Private Sub cmdDepot_Click()
frmDepot.Show
frmMenuAdmin.Hide
End Sub

Private Sub cmdENRE_Click()
Unload Me
frmEnregistrement.Show
End Sub

Private Sub cmdRetour_Click()
frmAccueil.Show
frmMenuAdmin.Hide
frmAdminConnex.Timer1.Enabled = True
frmAdminConnex.Label1.Caption = 20
End Sub

Private Sub cmdSolde_Click()
frmAdminConnex2.Show
frmMenuAdmin.Hide
End Sub
