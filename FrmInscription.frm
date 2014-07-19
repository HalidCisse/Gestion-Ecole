VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmInscription 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NOUVELLE INSCRIPTION"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmInscription.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   7215
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Reglement Financier"
      Height          =   1935
      Left            =   240
      TabIndex        =   24
      Top             =   5880
      Width           =   6735
      Begin VB.TextBox TxtPayIns 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   27
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox TxtPayTranch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   26
         Top             =   1320
         Width           =   3135
      End
      Begin VB.ComboBox ComboTypeReg 
         Appearance      =   0  'Flat
         Height          =   405
         ItemData        =   "FrmInscription.frx":27A2
         Left            =   4080
         List            =   "FrmInscription.frx":27AF
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Prix d'Inscription :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   1920
      End
      Begin VB.Label LabelPay_Tranche 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Par Tranche :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   2355
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Type de Reglement :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   2100
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informations de La Classe"
      Height          =   2775
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   6735
      Begin VB.Timer TimerClasseInfo 
         Interval        =   3000
         Left            =   1680
         Top             =   840
      End
      Begin MSComCtl2.DTPicker DTPickerFin 
         Height          =   375
         Left            =   4800
         TabIndex        =   20
         Top             =   2280
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
         Format          =   102105089
         CurrentDate     =   41773
      End
      Begin MSComCtl2.DTPicker DTPickerDeb 
         Height          =   375
         Left            =   4800
         TabIndex        =   19
         Top             =   1800
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
         Format          =   102105089
         CurrentDate     =   41773
      End
      Begin VB.ComboBox ComboAnnee 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   4320
         TabIndex        =   17
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox ComboNiveau 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   4320
         TabIndex        =   14
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox ComboFiliere 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   4320
         TabIndex        =   13
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fin Année Scolaire :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Debut Année Scolaire :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   2355
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Année :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Niveau :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Filière :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   780
      End
   End
   Begin VB.CommandButton CmdEnreg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enrégistré"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8040
      Width           =   1695
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "&H8000000F&"
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.TextBox TxtDetailSus 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   2400
         TabIndex        =   31
         Text            =   "Details de Suspension"
         ToolTipText     =   "Entrez Les Raisons de Suspension de l'inscription"
         Top             =   2040
         Width           =   4095
      End
      Begin VB.OptionButton OptExp 
         Caption         =   "Expiré"
         Height          =   375
         Left            =   5520
         TabIndex        =   23
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton OptSuspendu 
         Caption         =   "Suspendue"
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton OptActiv 
         Caption         =   "Active"
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox ComboClasse 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   8
         Top             =   1080
         Width           =   3135
      End
      Begin VB.ComboBox ComboNumMat 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   4
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox TxtNumIns 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   1
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Statut :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Classe :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "N° Inscription :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Etudiant :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   990
      End
   End
End
Attribute VB_Name = "FrmInscription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdEnreg_Click()

  If ChampsInsOk() Then
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM INSCRIPTIONS WHERE N°_Inscription = '" & UCase(Trim(TxtNumIns.Text)) & "' "
    rs.Open SQL, CN, adOpenKeyset
'----------------------------------------------------
    If rs.EOF Then
       x = MsgBox("Efféctuer L'Inscription ?", vbYesNo + vbQuestion)
       If x = vbYes Then
          EnregInsInfos (UCase(Trim(TxtNumIns.Text)))
          MsgBox "Inscription Effectuée Avec Succées !!", vbInformation
          '------------------
          If GetSetting("Email_Conf") = "Oui" Then
          x = MsgBox("Voulez Vous Envoyer Un Message de Confirmation a " & GetMyName(GetMaMatricule(Trim(TxtNumIns.Text))), vbQuestion + vbYesNo)
          If x = vbYes Then
            If SendInsMessage(Trim(TxtNumIns.Text)) Then
              MsgBox "Message Envoyer Avec Succès !", vbInformation
            End If
          End If
          End If
          '------------------
          Unload Me
       End If
    Else
       x = MsgBox("Modifier L'Inscription ?", vbYesNo + vbQuestion)
       If x = vbYes Then
          EnregInsInfos (UCase(Trim(TxtNumIns.Text)))
          MsgBox "Modification Effectuée Avec Succées !!", vbInformation
          Unload Me
       End If
    End If
'-------------------------------------------------------
rs.Close
Set rs = Nothing
End If
End Sub

Private Sub CmdRaz_Click()
  Dim C As Control
    For Each C In FrmInscription.Controls
      If TypeOf C Is TextBox And C.Name <> "TxtNumIns" Then
        C.Text = ""
      End If
      If TypeOf C Is ComboBox And C.Name <> "ComboTypeReg" Then
        C = ""
      End If
    Next
End Sub

Private Sub ComboClasse_Validate(Cancel As Boolean)
  GetClasseInfos (CStr(UCase(Trim(ComboClasse))))
End Sub

Private Sub ComboNumMat_Validate(Cancel As Boolean)
  Me.Caption = "NOUVELLE INSCRIPTION DE " & UCase(ComboNumMat)
End Sub

Private Sub ComboTypeReg_Validate(Cancel As Boolean)
    If ComboTypeReg = "Unique" Then
        LabelPay_Tranche.Caption = "Total Payement:"
    ElseIf ComboTypeReg = "Mensuel" Then
        LabelPay_Tranche.Caption = "Payement Mensuel :"
    Else
       LabelPay_Tranche.Caption = "Payement Trimestriel :"
    End If
End Sub

Private Sub Form_Load()
  RemplirComboFrmInscription
  If G_FullScreen = True Then
     Me_Top Me
  End If
End Sub

Private Sub OptActiv_Click()
  Frame1.Height = 2055
End Sub

Private Sub OptExp_Click()
   Frame1.Height = 2055
End Sub

Private Sub OptSuspendu_Click()
  Frame1.Height = 2415
End Sub

'##################### HALIDOU CISSE ##################
