VERSION 5.00
Begin VB.Form Main 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GESTION INSCRIPTIONS ETUDIANTS"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   13575
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   20.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00808080&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Main.frx":27A2
   ScaleHeight     =   7260
   ScaleWidth      =   13575
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerStat 
      Interval        =   10000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton CmdListREG 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "LISTE REGISTRATIONS"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   600
      Picture         =   "Main.frx":7515E7
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CommandButton CmdListINS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "LISTE INSCRIPTIONS"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   5520
      Picture         =   "Main.frx":7528FE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CommandButton CmdListPAY 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "LISTE PAYEMENTS"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   10440
      Picture         =   "Main.frx":753CD1
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CommandButton CmdPayment 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "PAYEMENT"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   10440
      Picture         =   "Main.frx":7550C1
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton CmdInscription 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "INSCRIPTION"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   5520
      Picture         =   "Main.frx":756CF3
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton CmdRegEtudiant 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "REGISTRATION"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   600
      MaskColor       =   &H00E0E0E0&
      Picture         =   "Main.frx":7583AB
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label LabNonPay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   11670
      TabIndex        =   10
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label LabNonIns 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6750
      TabIndex        =   9
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label LabListPay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   11670
      TabIndex        =   8
      Top             =   6720
      Width           =   135
   End
   Begin VB.Label LabListIns 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6705
      TabIndex        =   7
      Top             =   6720
      Width           =   135
   End
   Begin VB.Label LabListReg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1710
      TabIndex        =   6
      Top             =   6720
      Width           =   135
   End
   Begin VB.Menu mnProfile 
      Caption         =   "&Profile"
      Begin VB.Menu mnConnecter 
         Caption         =   "&Se Connecter"
      End
      Begin VB.Menu mnDeconnecter 
         Caption         =   "&Se Deconnecter"
         Shortcut        =   ^D
      End
      Begin VB.Menu rien 
         Caption         =   "------------------------"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnGesProfile 
         Caption         =   "&Gestion Des Profiles"
         Begin VB.Menu mnChPass 
            Caption         =   "&Modifié Mon Mot de Passe"
            Shortcut        =   ^C
         End
         Begin VB.Menu mnAddProfile 
            Caption         =   "&Nouvel Profile"
         End
         Begin VB.Menu mnModProfile 
            Caption         =   "&Modifié"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnRecherProfile 
            Caption         =   "&Recherche"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnSupProfile 
            Caption         =   "Supprimé Un Profile"
         End
         Begin VB.Menu mnListProfile 
            Caption         =   "&Liste Des Profiles"
         End
      End
      Begin VB.Menu rien2 
         Caption         =   "-----------------------"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnQuitter 
         Caption         =   "&QUITTER"
      End
   End
   Begin VB.Menu mnEtudiant 
      Caption         =   "&Etudiants"
      Begin VB.Menu mnRegistration 
         Caption         =   "&Registration "
         Shortcut        =   ^R
      End
      Begin VB.Menu mnModifiéETUD 
         Caption         =   "&Modification"
      End
      Begin VB.Menu mnSuppriméETUD 
         Caption         =   "&Suppression"
      End
      Begin VB.Menu mnChercherETUD 
         Caption         =   "&Recherche "
      End
      Begin VB.Menu mnListETUD 
         Caption         =   "&Liste "
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnInscription 
      Caption         =   "&Inscriptions"
      Begin VB.Menu mnAjouterINS 
         Caption         =   "&Nouvelle Inscription"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnModifiéINS 
         Caption         =   "&Modification"
      End
      Begin VB.Menu mnSuppriméINS 
         Caption         =   "&Suppression"
      End
      Begin VB.Menu mnRechercherINS 
         Caption         =   "&Recherche"
      End
      Begin VB.Menu mnListINS 
         Caption         =   "&Liste"
      End
   End
   Begin VB.Menu mnPayments 
      Caption         =   "&Payements"
      Begin VB.Menu MnPay 
         Caption         =   "&Récevoir Un Payement"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnPayAvance 
         Caption         =   "&Payement En Avance"
      End
      Begin VB.Menu MnSupPay 
         Caption         =   "&Suppression de Payement"
      End
      Begin VB.Menu mnListPay 
         Caption         =   "&Liste des Payements"
      End
   End
   Begin VB.Menu mnStats 
      Caption         =   "&Statistiques"
      Begin VB.Menu mnRapActv 
         Caption         =   "&Rapport  Activités"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnGraph 
         Caption         =   "&Graphiques"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnIMP 
      Caption         =   "&Impressions"
      Begin VB.Menu mnImprimEt 
         Caption         =   "&Imprimé Liste Des Etudiants"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnImpIns 
         Caption         =   "&Imprimé Liste Des Inscriptions"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnImpPays 
         Caption         =   "&Imprimé Liste Des Payements"
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu mnOutil 
      Caption         =   "&Outils"
      Begin VB.Menu mnCalc 
         Caption         =   "&Calculatrice"
      End
      Begin VB.Menu mnNotePad 
         Caption         =   "&NotePad"
      End
      Begin VB.Menu mnInternet 
         Caption         =   "&Internet"
         Begin VB.Menu mnGoogle 
            Caption         =   "&Google"
         End
         Begin VB.Menu mnGmail 
            Caption         =   "&Gmail"
         End
         Begin VB.Menu mnFacebook 
            Caption         =   "&Facebook"
         End
      End
      Begin VB.Menu mnOffice 
         Caption         =   "&MS Office"
         Begin VB.Menu mnWord 
            Caption         =   "&Word"
         End
         Begin VB.Menu mnExcel 
            Caption         =   "&Excel"
         End
         Begin VB.Menu mnPPT 
            Caption         =   "&PPT"
         End
         Begin VB.Menu mnAccess 
            Caption         =   "&MS Access"
         End
      End
      Begin VB.Menu mnSystem 
         Caption         =   "&System"
         Begin VB.Menu mnEtendre 
            Caption         =   "&Eteindre La Machine"
         End
         Begin VB.Menu mnRedemarer 
            Caption         =   "&Redemarer"
         End
      End
   End
   Begin VB.Menu mnParametres 
      Caption         =   "&Paramètres"
      Begin VB.Menu mnPref 
         Caption         =   "&Préferences"
      End
   End
   Begin VB.Menu mnAide 
      Caption         =   "&?"
      Begin VB.Menu mnApropos 
         Caption         =   "&A Propos"
      End
      Begin VB.Menu mnAide2 
         Caption         =   "&Rapport Et Soutenance Du Projet ( PFA)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnFeedB 
         Caption         =   "&Feed-Back"
      End
      Begin VB.Menu mnLicence 
         Caption         =   "&Licence"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'######################  MAIN  ###########################3

Private Sub Form_Load()
   If G_FullScreen = True Then
     Me_Top Me
   End If
   SyncData
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If G_FullScreen = True And Not G_Profil_Type = "ADMIN" Then
  MsgBox "Vous Devez Etre ADMIN Pour Pouvoir Fermer Le Programme"
  Cancel = 1
  Exit Sub
End If

  x = MsgBox("Ete Vous Sure de Quitter ? ", vbQuestion + vbYesNo)
  If x = vbNo Then
    Cancel = 1
  Else
    AddEvent ("Profil Deconnecter : " & G_Profil_Login)
    Beep
    UnloadForms
  End If

End Sub

Private Sub Form_Activate()
  UpdateStats
End Sub

Private Sub TimerStat_Timer()
  UpdateStats
End Sub
Private Sub mnQuitter_Click()
  Unload Me
End Sub
Private Sub mnDeconnecter_Click()
  x = MsgBox("Voulez Vous vous Deconnecté ?", vbQuestion + vbYesNo)
 If x = vbYes Then
    DeconnectProfil
 End If
End Sub

Private Sub mnPref_Click()
  FrmParametres.Show 1
End Sub

'##################### AIDE ########################
Private Sub mnFeedB_Click()
  EnvoiEmail_Me "halidcisse@gmail.com", App.Title & ": Ver " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnAide2_Click()
On Error Resume Next
  FrmRapport.Show 1
End Sub
Private Sub mnApropos_Click()
   frmAbout.Show 1
End Sub

Private Sub mnLicence_Click()
  FrmLicence.Show 1
End Sub

'#################   OUTILS  ###################
Private Sub mnCalc_Click()
  Shell "calc.EXE ", vbNormalFocus
End Sub

Private Sub mnNotePad_Click()
  Shell "NOTEPAD.EXE " & App.Path & "\NotePad.txt", vbNormalFocus
End Sub

Private Sub mnGoogle_Click()
  BrowseTo "http://www.Google.com"
End Sub

Private Sub mnFacebook_Click()
  BrowseTo "http://www.Facebook.com"
End Sub

Private Sub mnGmail_Click()
  BrowseTo "http://www.Gmail.com"
End Sub

Private Sub mnWord_Click()
  ShellExecute 0&, vbNullString, "WINWORD", vbNullString, vbNullString, SW_SHOWDEFAULT
End Sub

Private Sub mnExcel_Click()
  ShellExecute 0&, vbNullString, "EXCEL", vbNullString, vbNullString, SW_SHOWDEFAULT
End Sub

Private Sub mnPPT_Click()
  ShellExecute 0&, vbNullString, "POWERPNT", vbNullString, vbNullString, SW_SHOWDEFAULT
End Sub

Private Sub mnAccess_Click()
  ShellExecute 0&, vbNullString, "MSACCESS", vbNullString, vbNullString, SW_SHOWDEFAULT
End Sub

Private Sub mnEtendre_Click()
  Shell ("shutdown -s -f -t 0")
  End
End Sub

Private Sub mnRedemarer_Click()
  Shell ("shutdown -r -f -t 0")
  End
End Sub

'################ PROFILS  #########################
Private Sub mnAddProfile_Click()
  OpenAddProfil
End Sub

Private Sub mnModProfile_Click()
  OpenModProfil
End Sub

Private Sub mnChPass_Click()
  OpenModPass
End Sub

Private Sub mnSupProfile_Click()
  OpenSupProfil
End Sub

Private Sub mnRecherProfile_Click()
  OpenCherProfil
End Sub

Private Sub mnListProfile_Click()
  FrmGridProfil.Show 1
End Sub

'##################### ETUDIANTS ######################
Private Sub CmdListREG_Click()
  FrmGridEtudiant.Show 1
End Sub

Private Sub CmdRechercheEtudiant_Click()
  OpenModEtudiant
End Sub

Private Sub CmdRegEtudiant_Click()
    OpenAddEtudiant
End Sub

Private Sub mnRegistration_Click()
  OpenAddEtudiant
End Sub

Private Sub mnModifiéETUD_Click()
  OpenModEtudiant
End Sub

Private Sub mnSuppriméETUD_Click()
  OpenSupEtudiant
End Sub

Private Sub mnChercherETUD_Click()
  OpenModEtudiant
End Sub

Private Sub mnListETUD_Click()
  FrmGridEtudiant.Show 1
End Sub

'###################### INSCRIPTIONS #########################
Private Sub CmdListINS_Click()
  FrmGridInscription.Show 1
End Sub

Private Sub CmdInscription_Click()
  OpenAddIns
End Sub

Private Sub CmdRechecheInscription_Click()
  FrmGridInscription.Show 1
End Sub

Private Sub mnAjouterINS_Click()
  FrmInscription.Show 1
End Sub

Private Sub mnModifiéINS_Click()
Dim INS As String
DebutSup:
   INS = Trim(InputBox("Entrez Le Matricule de L'Etudiant a Modifier "))
   If INS <> "" Then
    If InsExist(CStr(INS)) Then
      OpenModIns (CStr(INS))
    Else
      y = MsgBox("Inscription Inéxistant", vbCritical + vbRetryCancel)
      If y = vbRetry Then
        GoTo DebutSup
      End If
    End If
   End If
End Sub

Private Sub mnSuppriméINS_Click()
  OpenSupIns
End Sub

Private Sub mnRechercherINS_Click()
  OpenModIns
End Sub

Private Sub mnListINS_Click()
  FrmGridInscription.Show 1
End Sub

'################### PAYMENTS  ######################
Private Sub CmdListPAY_Click()
  If Not G_Profil_Type = "ADMIN" Then
    MsgBox "Vous n'avez pas le privilège de voir les payments effectués !", vbCritical
    Exit Sub
  End If
  FrmGridPayment.Show 1
End Sub
Private Sub CmdPayment_Click()
  FrmPayment.Show 1
End Sub
Private Sub MnPayAvance_Click()
  OpenPayAvance
End Sub
Private Sub mnPay_Click()
  FrmPayment.Show 1
End Sub
Private Sub MnSupPay_Click()
  OpenSupPayment
End Sub
Private Sub mnListPay_Click()
  If Not G_Profil_Type = "ADMIN" Then
    MsgBox "Vous n'avez pas le privilège de voir les payments effectués !", vbCritical
    Exit Sub
  End If
  FrmGridPayment.Show 1
End Sub

'###################### IMPRESSION #####################
Private Sub mnImprimEt_Click()
  Open_DR_Etud
End Sub

Private Sub mnImpIns_Click()
  Open_DR_Ins
End Sub

Private Sub mnImpPays_Click()
  Open_DR_Pay
End Sub

'##################### STATISTIQUES #####################
Private Sub mnGraph_Click()
  FrmGraph.Show 1
End Sub

Private Sub mnRapActv_Click()
  FrmGridEvent.Show 1
End Sub

'##################### HALIDOU CISSE ##################

