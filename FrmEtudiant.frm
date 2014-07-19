VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmEtudiant 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Régistration :  Etudiant"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmEtudiant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   7200
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdEnreg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enregistré"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Enrégistré"
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton CmdRaz 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Effacé"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton CmdSuiv 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Suivant >"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Page Suivante"
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton CmdPrec 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "< Prècedent"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Page Prècedente"
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Informations Suplementaires"
      Height          =   4815
      Left            =   7200
      TabIndex        =   38
      Top             =   3240
      Width           =   6735
      Begin VB.ComboBox ComboNEtude 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   4320
         Sorted          =   -1  'True
         TabIndex        =   57
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox TxtMat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   54
         Top             =   840
         Width           =   3135
      End
      Begin VB.Frame Frame5 
         Caption         =   "Statut"
         Height          =   1095
         Left            =   120
         TabIndex        =   49
         Top             =   3360
         Width           =   6495
         Begin VB.OptionButton OptAband 
            Caption         =   "Abandonné"
            Height          =   270
            Left            =   3480
            TabIndex        =   53
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton OptDiplome 
            Caption         =   "Diplomé"
            Height          =   270
            Left            =   5160
            TabIndex        =   52
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton OptNonRegulier 
            Caption         =   "Non Régulier"
            Height          =   270
            Left            =   1680
            TabIndex        =   51
            Top             =   480
            Width           =   1815
         End
         Begin VB.OptionButton OptRegulier 
            Caption         =   "Régulier"
            Height          =   270
            Left            =   360
            TabIndex        =   50
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Label LabelDateEnreg 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inconnue !"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3360
         TabIndex        =   65
         Top             =   2760
         Width           =   3045
      End
      Begin VB.Label LabelClasse 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Non Inscrit"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3360
         TabIndex        =   64
         Top             =   1800
         Width           =   3045
      End
      Begin VB.Label LabelTotalPayment 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3360
         TabIndex        =   63
         Top             =   2280
         Width           =   3045
      End
      Begin VB.Label LabelNIns 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Non Inscrit"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3360
         TabIndex        =   62
         Top             =   1320
         Width           =   3045
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Payments"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   61
         Top             =   2400
         Width           =   1590
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "N° Inscription"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   56
         Top             =   1440
         Width           =   1440
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Niveau d'Etude"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   55
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "N° Matricule"
         Height          =   285
         Left            =   120
         TabIndex        =   41
         Top             =   960
         Width           =   1290
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Classe"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   40
         Top             =   1920
         Width           =   645
      End
      Begin VB.Label LabelDateEnreg1 
         AutoSize        =   -1  'True
         Caption         =   "Date Enrégistrement"
         Height          =   285
         Left            =   120
         TabIndex        =   39
         Top             =   2880
         Width           =   2145
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   "Informations du Tuteur"
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   7200
      TabIndex        =   28
      Top             =   120
      Width           =   6735
      Begin VB.TextBox TxtPrenomT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   58
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox TxtNomT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   37
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox TxtAddressT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   31
         Top             =   2280
         Width           =   3135
      End
      Begin VB.TextBox TxtEmailT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   30
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox TxtTelT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   29
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Adresse Domicile"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Numéro de Téléphone"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Prénom "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nom "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Informations Civiles"
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   6735
      Begin VB.TextBox TxtNomM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   60
         Top             =   3600
         Width           =   3135
      End
      Begin VB.TextBox TxtPrenom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   59
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox TxtNomP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   46
         Top             =   3120
         Width           =   3135
      End
      Begin VB.TextBox TxtAdress 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   17
         Top             =   5040
         Width           =   3135
      End
      Begin VB.TextBox TxtEmail 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   16
         Top             =   4560
         Width           =   3135
      End
      Begin VB.TextBox TxtNumTel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   15
         Top             =   4080
         Width           =   3135
      End
      Begin VB.ComboBox ComboNat 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   4320
         TabIndex        =   14
         Top             =   2640
         Width           =   2175
      End
      Begin VB.ComboBox ComboLieuNaiss 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   4320
         TabIndex        =   13
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox TxtNom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   12
         Top             =   720
         Width           =   3135
      End
      Begin VB.OptionButton OptionF 
         Appearance      =   0  'Flat
         Caption         =   "Femme"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5280
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton OptionH 
         Appearance      =   0  'Flat
         Caption         =   "Homme"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3360
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPNaiss 
         Height          =   375
         Left            =   4800
         TabIndex        =   18
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   76480513
         CurrentDate     =   41738
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nom et Prénom de la Mère"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   48
         Top             =   3720
         Width           =   2790
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Prénom du Père"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   47
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Adresse Domicile"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Numéro de Téléphone"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   4200
         Width           =   2655
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nationalité"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Lieu de Naissance"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date de Naissance"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Prénom"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nom"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Civilité"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ID"
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.OptionButton OptionCS 
         Appearance      =   0  'Flat
         Caption         =   "Carte Séjour"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4800
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton OptionPass 
         Appearance      =   0  'Flat
         Caption         =   "Passport"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3360
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton OptionCIN 
         Appearance      =   0  'Flat
         Caption         =   "CIN"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox TxtNumID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   1
         Top             =   840
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker DTPExpID 
         Height          =   375
         Left            =   4800
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   76480513
         CurrentDate     =   41738
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Expiration"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Numéro de L'identité"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Type d'identité"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
   End
End
Attribute VB_Name = "FrmEtudiant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdEnreg_Click()
'Quand On Click Sur Enregistrer MODIFIER OU AJOUTER !

  If ChampsEtudiantOk() Then
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM ETUDIANTS WHERE Matricule = '" & UCase(TxtMat.Text) & "' "
    rs.Open SQL, CN, adOpenKeyset
'----------------------------------------------------
    If rs.EOF Then
    'Si C'Est Ajout
       x = MsgBox("Efféctuer l'Enregistrement ?", vbYesNo + vbQuestion)
       If x = vbYes Then
          EnregEtudiantInfos (UCase(TxtMat.Text))
          MsgBox "Registration Effectuée Avec Succées !!", vbInformation
          '--------------------------------------
          If GetSetting("Email_Conf") = "Oui" Then
          x = MsgBox("Voulez Vous Envoyer Un Email de Confirmation à " & GetMyName(TxtMat.Text), vbYesNo + vbQuestion)
          If x = vbYes Then
            If SendWelcomeMessage(TxtMat.Text) Then
              MsgBox "Email Envoyer Avec Succès"
            End If
          End If
          End If
          '--------------------------------------
          Unload Me
       End If
    Else
    'Si C'Est Modification
       x = MsgBox("Efféctuer la modification ?", vbYesNo + vbQuestion)
       If x = vbYes Then
          EnregEtudiantInfos (UCase(TxtMat.Text))
          MsgBox "Modification Effectuée Avec Succées !!", vbInformation
          Unload Me
       End If
    End If
'-----------------------------------------------------
rs.Close
Set rs = Nothing
End If
End Sub

Private Sub CmdPrec_Click()
  QuelPageEtudiant (Page - 1)
End Sub

Private Sub CmdRaz_Click()
Dim C As Control
    For Each C In FrmEtudiant.Controls
      If TypeOf C Is TextBox And C.Name <> "TxtMat" Then
        C.Text = ""
      End If
    Next
End Sub

Private Sub CmdSuiv_Click()
  QuelPageEtudiant (Page + 1)
End Sub

Private Sub Form_Activate()
  ScaleEtudiant
End Sub

Private Sub Form_Load()
  RemplirCombosFrmEtudiant
  '-----------------
  If G_FullScreen = True Then
     Me_Top Me
  End If
  '-----------------
End Sub

'##################### HALIDOU CISSE ##################
