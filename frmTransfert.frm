VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTransfert 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   7395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNumCompte 
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   20
      Top             =   960
      Visible         =   0   'False
      Width           =   2295
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
      Connect         =   $"frmTransfert.frx":0000
      OLEDBString     =   $"frmTransfert.frx":009E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TABLE_CLIENT"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtMontActu 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   6240
      Width           =   2835
   End
   Begin VB.CommandButton cmdRetour 
      Caption         =   "Retour"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   17
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   16
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdVis 
      Caption         =   "Afficher le bénéficiare"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   5760
      Width           =   3015
   End
   Begin VB.CommandButton cmdTrans 
      Caption         =   "Transférer"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      Top             =   6840
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmTransfert.frx":013C
      Height          =   1815
      Left            =   360
      TabIndex        =   13
      Top             =   3840
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "Sélectionner le bénéficiaire"
      ColumnCount     =   5
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
      BeginProperty Column04 
         DataField       =   "Ville"
         Caption         =   "Ville"
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
            ColumnWidth     =   2115,213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1289,764
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2174,74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3030,236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2115,213
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   6240
      Width           =   2775
   End
   Begin VB.TextBox txtNom 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox txtMontantRetirer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   2520
      Width           =   2745
   End
   Begin VB.TextBox txtMontantRestant 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "MontantRestant"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   480
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2520
      Width           =   2820
   End
   Begin VB.TextBox txtMontant_Client 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   2835
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Montant_Actuel"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   5220
      TabIndex        =   19
      Top             =   6240
      Width           =   1980
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compte de l'Expéditeur"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   3600
      TabIndex        =   12
      Top             =   1080
      Width           =   3390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compte du Bénéficiaire"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   4200
      TabIndex        =   11
      Top             =   3240
      Width           =   3345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nom"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   6240
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nom"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   585
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Montant à Transférer"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   2685
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "MontantRestant:"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   6255
      TabIndex        =   5
      Top             =   2520
      Width           =   2115
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Montant_Client:"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   6255
      TabIndex        =   4
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TRANSFERT D'ARGENT"
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
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   5865
   End
End
Attribute VB_Name = "frmTransfert"
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

Private Sub cmdActu_Click()
strChaine = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BASE_DE_DONNEES_VB611.mdb;Persist Security Info=False"
Dim cn As New ADODB.Connection
cn.Open strChaine
cn.Execute "update table_client set "
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "select * from table_client", cn, adOpenStatic, adLockBatchOptimistic
'Set DataGrid1.DataSource = rs
Set Adodc1.Recordset = rs
MsgBox rs.RecordCount
End Sub

Private Sub cmdAnnuler_Click()
txtNo.Text = ""
txtMontActu.Text = ""
MsgBox "Resélectionner le bénéficiaire"
txtNo.SetFocus
End Sub

Private Sub cmdRetour_Click()
Unload Me
frmMenuClient.Show
End Sub

Private Sub cmdTrans_Click()
strChaine = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BASE_DE_DONNEES_VB611.mdb;Persist Security Info=False"
Dim cn As New ADODB.Connection
cn.Open strChaine
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "select * from table_client where Numero_de_Compte_Client = " & Val(txtNumCompte.Text), cn, adOpenStatic, adLockBatchOptimistic
If txtMontantRetirer.Text <> "" And txtNo.Text <> "" Then
    montantExpeRestant = Val(txtMontantRestant.Text)
    soldeBenef = Val(txtMontantRetirer.Text) + Val(txtMontActu.Text)
    montantTransfert = Val(txtMontantRetirer.Text)
    benef = txtNo.Text
    expe = txtNom.Text
    If MsgBox("Vous allez envoyer " & montantTransfert & " à " & benef, vbCritical + vbYesNo, avertissement) = vbYes Then
        rs!Montant_Client = montantExpeRestant
        rs.UpdateBatch
        txtMontant_Client.Text = rs!Montant_Client
        'cn.Close
        Adodc1.Recordset!Montant_Client = soldeBenef
        Adodc1.Recordset.UpdateBatch
        Adodc1.Refresh
        Set DataGrid1.DataSource = Adodc1
        txtNo.Text = ""
        txtMontActu.Text = ""
        txtMontantRetirer.Text = ""
        MsgBox "Cher(e) " & expe & " votre compte est débité de " & montantTransfert & " ce jour. Solde Disponible = " & montantExpeRestant, vbInformation + vbOKOnly
    End If
Else
    MsgBox "Montant ou Bénéficiaire manquant"
    txtMontantRetirer.SetFocus
End If
End Sub
Private Sub cmdVis_Click()
txtNo.Text = Adodc1.Recordset!Nom_Client
txtMontActu = Adodc1.Recordset!Montant_Client
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
    txtMontantRetirer.Text = ""
  
    End If
    End If
 
End Sub

Private Sub txtMontActu_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
 KeyAscii = KeyAscii
 Else
  KeyAscii = 0
 End If
End Sub

Private Sub txtMontantRestant_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
 KeyAscii = KeyAscii
 Else
  KeyAscii = 0
 End If
End Sub

Private Sub txtMontantRetirer_Change()
miseAjourMontant
End Sub

Private Sub txtMontantRetirer_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
 KeyAscii = KeyAscii
 Else
  KeyAscii = 0
 End If
End Sub

Private Sub miseAjourMontant()
If Val(txtMontantRetirer) > Val(txtMontant_Client) Then
  MsgBox "Votre solde ne vous permet pas d'effectuer ce transfert", vbCritical
  txtMontantRetirer.Text = ""
  txtMontantRestant.Text = ""
  txtMontantRetirer.SetFocus
Else
    txtMontantRestant.Text = Val(txtMontant_Client.Text) - Val(txtMontantRetirer.Text)
End If
End Sub
