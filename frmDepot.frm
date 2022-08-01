VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDepot 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Dépot"
   ClientHeight    =   6450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10275
   FillColor       =   &H0000FFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmDepot.frx":0000
      Height          =   1815
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   14737632
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   24
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "LISTE DES COMPTES"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Nom_Client"
         Caption         =   "Nom_Client"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Prenoms_Client"
         Caption         =   "Prenoms_Client"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Montant_Client"
         Caption         =   "Montant_Client"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Numero_de_Compte_Client"
         Caption         =   "Numero_de_Compte_Client"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         Size            =   0
         BeginProperty Column00 
            ColumnWidth     =   1844,787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2294,929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2174,74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2970,142
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdVisualiser 
      Caption         =   "Visualiser"
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
      Left            =   7680
      TabIndex        =   11
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
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
      Left            =   11040
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtMont 
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
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3600
      Width           =   3975
   End
   Begin VB.CommandButton cmdRetour 
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
      Left            =   7680
      TabIndex        =   7
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton cmdAnnuler 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Annuler"
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
      Left            =   7680
      TabIndex        =   6
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdDeposer 
      Caption         =   "Déposer"
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
      Left            =   7680
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   10440
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmDepot.frx":0015
      OLEDBString     =   $"frmDepot.frx":00B3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TABLE_CLIENT"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtMontant 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2880
      Width           =   3945
   End
   Begin VB.TextBox txtMontantDepot 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   3480
      TabIndex        =   2
      Top             =   2160
      Width           =   3975
   End
   Begin VB.TextBox txtNom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1455
      Width           =   3975
   End
   Begin VB.Label Label6 
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
      Left            =   600
      TabIndex        =   14
      Top             =   240
      Width           =   8760
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom_Client"
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
      TabIndex        =   13
      Top             =   1440
      Width           =   2115
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Montant_Client"
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
      TabIndex        =   12
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "MontantActu"
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
      TabIndex        =   8
      Top             =   3600
      Width           =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "MontantaDéposer"
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
      TabIndex        =   1
      Top             =   2160
      Width           =   3195
   End
End
Attribute VB_Name = "frmDepot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAnnuler_Click()
txtNom.Text = ""
txtMontantDepot = ""
txtMontant = ""
txtMont = ""
'txtNom.SetFocus
End Sub

Private Sub cmdDeposer_Click()
   If MsgBox("Voulez-Vous vraiment faire ce Dépot", vbYesNo + vbQuestion, avertissement) = vbYes Then
   cmdOk_Click
   Adodc1.Recordset!Montant_Client = txtMontant.Text
    Adodc1.Recordset.UpdateBatch
     txtMontantDepot = ""
     txtMontant = ""
     End If
End Sub

Private Sub cmdOk_Click()
txtMont = Val(txtMontantDepot) + Val(txtMontant)
txtMontant = txtMont
End Sub

Private Sub CmdRetour_Click()
frmMenuAdmin.Show
frmDepot.Hide
End Sub

Private Sub cmdVisualiser_Click()
txtNom.Text = Adodc1.Recordset!Nom_Client
txtMontant.Text = Adodc1.Recordset!Montant_Client
End Sub

Private Sub txtMontantDepot_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
 KeyAscii = KeyAscii
 Else
  KeyAscii = 0
 End If
End Sub

