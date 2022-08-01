VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCompte 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Compte"
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13110
   BeginProperty Font 
      Name            =   "Tiro Devanagari Marathi"
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
   ScaleHeight     =   6810
   ScaleWidth      =   13110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNum2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataMember      =   "Command1"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   3000
      TabIndex        =   28
      Top             =   1800
      Width           =   3060
   End
   Begin VB.TextBox txtVille 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   9600
      TabIndex        =   24
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox txtAge 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   3000
      TabIndex        =   22
      Top             =   5040
      Width           =   3015
   End
   Begin VB.TextBox txtNum 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9600
      TabIndex        =   21
      Top             =   4920
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   3000
      TabIndex        =   20
      Top             =   4320
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   53870593
      CurrentDate     =   44715
   End
   Begin VB.ComboBox Cbogenre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9600
      TabIndex        =   16
      Text            =   "Sélectionner votre genre"
      Top             =   4320
      Width           =   3015
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   12720
      TabIndex        =   15
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton cmdModifier 
      Caption         =   "Modifier"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   6000
      Width           =   2175
   End
   Begin VB.TextBox txtNumdeCompte 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataMember      =   "Command1"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   3000
      TabIndex        =   7
      Top             =   2520
      Width           =   3060
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   12480
      Top             =   600
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
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
      Connect         =   $"frmCompte.frx":0000
      OLEDBString     =   $"frmCompte.frx":009E
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
   Begin VB.TextBox txtMontant 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataMember      =   "Command1"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   9600
      TabIndex        =   6
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox txtMotdePasse 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataMember      =   "Command1"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00000000&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   9600
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox txtPrenoms 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataMember      =   "Command1"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   3000
      TabIndex        =   4
      Top             =   3720
      Width           =   3015
   End
   Begin VB.TextBox txtNom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataMember      =   "Command1"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   3000
      TabIndex        =   3
      Top             =   3120
      Width           =   3015
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
      Height          =   495
      Left            =   9120
      TabIndex        =   2
      Top             =   6000
      Width           =   2175
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
      Left            =   12240
      TabIndex        =   1
      Top             =   -120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdEnregistrer 
      Caption         =   "Enregistrer"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   6000
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   9600
      TabIndex        =   26
      Top             =   3600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   53870593
      CurrentDate     =   44715
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "N° de carte d'Identité"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   1800
      Width           =   2685
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date d'Enregistrement"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6600
      TabIndex        =   27
      Top             =   3600
      Width           =   2805
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ville"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6600
      TabIndex        =   25
      Top             =   1800
      Width           =   585
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   5040
      Width           =   465
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numéro de téléphone"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6600
      TabIndex        =   19
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Genre"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6600
      TabIndex        =   18
      Top             =   4320
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date de Naissance"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   4320
      Width           =   2250
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COMPTES ET INFORMATIONS"
      BeginProperty Font 
         Name            =   "Tiro Devanagari Marathi"
         Size            =   35.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1185
      Left            =   1080
      TabIndex        =   14
      Top             =   240
      Width           =   10935
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Compte"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   2385
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   2400
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Montant Client"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6600
      TabIndex        =   11
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Prénoms Client"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   1920
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom Client"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   1440
   End
End
Attribute VB_Name = "frmCompte"
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

Private Sub cmdAnnuler_Click()
txtNom.Text = " "
        txtPrenoms.Text = " "
        txtMontant.Text = " "
        txtMotdePasse.Text = " "
        txtNumdeCompte.Text = ""
         txtVille.Text = ""
'    DTPicker1.Value = ""
'     DTPicker2.Value = ""
     txtNum2.Text = ""
    txtAge.Text = ""
    txtNum.Text = ""
    Cbogenre.Text = ""
        txtNom.SetFocus

End Sub
Private Sub cmdEnregistrer_Click()
If Len(txtMotdePasse.Text) >= 6 Then
 
If (txtNom.Text <> " ") And (txtPrenoms.Text <> " ") Then
    If MsgBox(" Voulez-vous vraiment enrégister", vbYesNo, avertissement) = vbYes Then
       Adodc1.Recordset.AddNew
        Adodc1.Recordset!Nom_Client = txtNom.Text
        Adodc1.Recordset!Prenoms_Client = txtPrenoms.Text
        Adodc1.Recordset!Montant_Client = txtMontant.Text
        Adodc1.Recordset!Password = txtMotdePasse.Text
         Adodc1.Recordset!Numero_de_Compte_Client = txtNumdeCompte.Text
              Adodc1.Recordset!Ville = txtVille.Text
                   Adodc1.Recordset!Naissance = DTPicker1.Value
                        Adodc1.Recordset!Enregistrement = DTPicker2.Value
                          Adodc1.Recordset!identite = txtNum2.Text
                        Adodc1.Recordset!Age = txtAge.Text
                          Adodc1.Recordset!Tel = txtNum.Text
                        Adodc1.Recordset!Genre = Cbogenre.Text
                        
        Adodc1.Recordset.Update
        MsgBox " Enrégistrement effectué avec succès", vbInformation
cmdAnnuler_Click
    End If
    Else
   MsgBox " Veuillez saisir les infos manquantes", vbCritical
   End If
   Else
   MsgBox " Votre Code doit être supérieure à 6 caractères", vbCritical
End If
End Sub

Private Sub cmdModifier_Click()

    Adodc1.Recordset!Nom_Client = txtNom.Text
    Adodc1.Recordset!Prenoms_Client = txtPrenoms.Text
    Adodc1.Recordset!Montant_Client = txtMontant.Text
    Adodc1.Recordset!Password = txtMotdePasse.Text
    Adodc1.Recordset!Numero_de_Compte_Client = txtNumdeCompte.Text
    Adodc1.Recordset!Ville = txtVille.Text
    Adodc1.Recordset!Naissance = DTPicker1.Value
    Adodc1.Recordset!Enregistrement = DTPicker2.Value
    Adodc1.Recordset!identite = txtNum2.Text
    Adodc1.Recordset!Age = txtAge.Text
    Adodc1.Recordset!Tel = txtNum.Text
        Adodc1.Recordset!Genre = Cbogenre.Text
    Adodc1.Recordset.UpdateBatch
    MsgBox "Modification effectuée avec succès", vbInformation
    cmdAnnuler_Click
End Sub

Private Sub cmdRetour_Click()
frmMenuAdmin.Show
frmCompte.Hide
End Sub
Private Sub cmdVisualiser_Click()
 txtNom.Text = Adodc1.Recordset!Nom_Client
 txtPrenoms.Text = Adodc1.Recordset!Prenoms_Client
txtMontant.Text = Adodc1.Recordset!Montant_Client
txtMotdePasse.Text = Adodc1.Recordset!Password
txtNumdeCompte.Text = Adodc1.Recordset!Numero_de_Compte_Client
txtVille.Text = Adodc1.Recordset!Ville
    DTPicker1.Value = Adodc1.Recordset!Naissance
    DTPicker2.Value = Adodc1.Recordset!Enregistrement
    txtNum2.Text = Adodc1.Recordset!identite
    txtAge.Text = Adodc1.Recordset!Age
    txtNum.Text = Adodc1.Recordset!Tel
Cbogenre.Text = Adodc1.Recordset!Genre
End Sub
Private Sub Form_Load()
Cbogenre.AddItem "Homme"
Cbogenre.AddItem "Femme"
End Sub
Private Sub txtAge_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
 KeyAscii = KeyAscii
 Else
  KeyAscii = 0
 End If
End Sub

Private Sub txtMontant_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
 KeyAscii = KeyAscii
 Else
  KeyAscii = 0
 End If
End Sub
Private Sub txtNum_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
 KeyAscii = KeyAscii
 Else
  KeyAscii = 0
 End If
End Sub
Private Sub txtNum2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
 KeyAscii = KeyAscii
 Else
  KeyAscii = 0
 End If
End Sub
Private Sub txtNumdeCompte_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
 KeyAscii = KeyAscii
 Else
  KeyAscii = 0
 End If
End Sub
