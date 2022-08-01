VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTranfert 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "saveDesti"
      Height          =   255
      Left            =   5040
      TabIndex        =   27
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Saveexp"
      Height          =   375
      Left            =   2880
      TabIndex        =   26
      Top             =   5760
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   240
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Connect         =   $"frmTranfert.frx":0000
      OLEDBString     =   $"frmTranfert.frx":009E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TABLE_CLIENT"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "cmd2"
      Height          =   375
      Left            =   6360
      TabIndex        =   25
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1"
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   6960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtMontActu 
      BackColor       =   &H00FFFFFF&
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8160
      TabIndex        =   22
      Top             =   2640
      Width           =   2835
   End
   Begin VB.CommandButton cmdVal 
      Caption         =   "val"
      Height          =   255
      Left            =   5520
      TabIndex        =   21
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdRetour 
      Caption         =   "Retour"
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
      Left            =   8400
      TabIndex        =   20
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
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
      Left            =   6960
      TabIndex        =   19
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdVis 
      Caption         =   "Visualiser"
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
      Left            =   4920
      TabIndex        =   18
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ok"
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
      Left            =   1800
      TabIndex        =   17
      Top             =   6960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdTrans 
      Caption         =   "Transférer"
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
      Left            =   3000
      TabIndex        =   16
      Top             =   5040
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmTranfert.frx":013C
      Height          =   1095
      Left            =   7320
      TabIndex        =   15
      Top             =   3840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1931
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   3060,284
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtNo 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      TabIndex        =   11
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox txtMontantBenef 
      BackColor       =   &H00FFFFFF&
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   6960
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.TextBox txtNom 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox txtMontantRetirer 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   2880
      TabIndex        =   3
      Top             =   2640
      Width           =   2820
   End
   Begin VB.TextBox txtMontantRestant 
      BackColor       =   &H00FFFFFF&
      DataField       =   "MontantRestant"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   2880
      TabIndex        =   2
      Top             =   3840
      Width           =   2820
   End
   Begin VB.TextBox txtMontant_Client 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3240
      Width           =   2835
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Montant_Actuel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   4
      Left            =   6000
      TabIndex        =   23
      Top             =   2760
      Width           =   1920
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compte de l'Expéditeur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1320
      TabIndex        =   14
      Top             =   1080
      Width           =   3315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compte du Bénéficiaire"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   7200
      TabIndex        =   13
      Top             =   1080
      Width           =   3285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6120
      TabIndex        =   12
      Top             =   2040
      Width           =   555
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
      ForeColor       =   &H00FFFF00&
      Height          =   300
      Index           =   3
      Left            =   3240
      TabIndex        =   10
      Top             =   7080
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   555
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Montant à Transférer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   255
      TabIndex        =   6
      Top             =   2640
      Width           =   2550
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "MontantRestant:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   2040
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
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   1920
   End
   Begin VB.Line Line1 
      X1              =   5880
      X2              =   5880
      Y1              =   1680
      Y2              =   4560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TRANSFERT D'ARGENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   5820
   End
End
Attribute VB_Name = "frmTranfert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
txtMontActu.Text = Val(txtMontantRetirer) + Val(txtMontantBenef)
End Sub

Private Sub cmd2_Click()
Adodc1.Recordset!Nom_Client = txtNo.Text
Adodc1.Recordset!Montant_Client = txtMontantBenef.Text
Adodc1.Recordset.UpdateBatch
End Sub

Private Sub CmdRetour_Click()
Unload Me
frmMenuClient.Show
End Sub

Private Sub cmdTrans_Click()
miseAjourMontant
'cmd2_Click

End Sub

Private Sub miseAjourMontant()
txtMontantRestant.Text = Val(txtMontant_Client.Text) - Val(txtMontantRetirer.Text)
txtMontant_Client.Text = txtMontantRestant.Text
End Sub

Private Sub cmdVis_Click()
txtNo.Text = Adodc1.Recordset!Nom_Client
txtMontantBenef = Adodc1.Recordset!Montant_Client
End Sub

Private Sub Command1_Click()
    Adodc1.Recordset!Montant_Client = txtMontantRestant.Text
    Adodc1.Recordset.UpdateBatch
End Sub

Private Sub verifierMontant()
If Val(txtMontantRetirer) > Val(txtMontant_Client) Then
  MsgBox "Votre solde ne vous permet pas d'effectuer ce transfert", vbCritical
  txtMontantRetirer.Text = ""
Else
   If MsgBox("Voulez-Vous vraiment faire ce Transfert?", vbYesNo, avertissement) = vbYes Then
   miseAjourMontant
    Adodc1.Recordset!Montant_Client = txtMontant_Client.Text
    Adodc1.Recordset!Nom_Client = txtNo.Text
Adodc1.Recordset!Montant_Client = txtMontantBenef.Text
Adodc1.Recordset.UpdateBatch
    'cmd1_Click
       '    Adodc1.Recordset.UpdateBatch
    txtMontantRetirer.Text = ""
  
    End If
    End If
 
End Sub

Private Sub Command3_Click()
    Adodc1.Recordset!Montant_Client = txtMontActu.Text
    Adodc1.Recordset.UpdateBatch
End Sub

Private Sub Command4_Click()
lbl_Id.Text = Row.Cells(1).Text
End Sub

