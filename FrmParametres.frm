VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmParametres 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paramètres"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmParametres.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   8760
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   3120
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   240
      TabIndex        =   12
      Top             =   6720
      Width           =   8295
      Begin VB.CheckBox ChkConf 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Envoyer Email de Confirmation"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4320
         TabIndex        =   16
         ToolTipText     =   "Selectionner si vous voulez Envoyer un Email de Confirmation de Payment ou d'inscription"
         Top             =   960
         Width           =   3615
      End
      Begin VB.CheckBox ChkSusp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Suspendre L'Inscription en cas de defaut de Payment > 3 Mois"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4320
         TabIndex        =   15
         ToolTipText     =   "Selectionner Si vous  voulez Une Suspension Automatique"
         Top             =   360
         Width           =   3855
      End
      Begin VB.CheckBox ChkFull 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bloquer L'Ecran Au Demarage "
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   3615
      End
      Begin VB.CheckBox chkRun 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Lancer Au Demarage de Windows"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         ToolTipText     =   "Selectionner si vous voulez un demarage automatique"
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.TextBox TxtEcoleSiege 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF8080&
      Height          =   600
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   5760
      Width           =   4815
   End
   Begin VB.TextBox TxtEcoleEmail 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF8080&
      Height          =   390
      Left            =   3720
      TabIndex        =   9
      Text            =   " "
      Top             =   4920
      Width           =   4815
   End
   Begin VB.TextBox TxtEcoleFax 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF8080&
      Height          =   390
      Left            =   3720
      TabIndex        =   7
      Top             =   4200
      Width           =   4815
   End
   Begin VB.TextBox TxtEcoleTEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF8080&
      Height          =   390
      Left            =   3720
      TabIndex        =   5
      Top             =   3480
      Width           =   4815
   End
   Begin VB.TextBox TxtEcoleFormJ 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF8080&
      Height          =   390
      Left            =   3720
      TabIndex        =   3
      Text            =   " "
      Top             =   2760
      Width           =   4815
   End
   Begin VB.TextBox TxtEcoleNom 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF8080&
      Height          =   390
      Left            =   3720
      TabIndex        =   1
      Text            =   " "
      Top             =   2040
      Width           =   4815
   End
   Begin VB.Image EcoleLOGO 
      Appearance      =   0  'Flat
      Height          =   1695
      Left            =   0
      Picture         =   "FrmParametres.frx":27A2
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click Pour Changer"
      Top             =   0
      Width           =   8775
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Adresse Physique"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   5640
      Width           =   1845
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Adresse Electronique"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   4920
      Width           =   2220
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "N° Telephone"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Forme Juridique"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom de L'Ecole"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
End
Attribute VB_Name = "FrmParametres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
  '-----------------------------
  If G_FullScreen = True Then
     Me_Top Me
  End If
  '-----------------------------
  GetParam
  
  On Error Resume Next
  EcoleLOGO.Picture = LoadPicture(App.Path & "\EcoleLogo.jpg")
End Sub

Private Sub EcoleLOGO_DblClick()
'Modifier Le Logo de L'Ecole
Dim Dir
On Error GoTo z
  CommonDialog.DialogTitle = "Selectionner Le Logo de L'Ecole :"
  CommonDialog.InitDir = App.Path
  CommonDialog.Filter = "Pictures (*.jpg)"
  CommonDialog.ShowOpen
  Dir = CommonDialog.FileName
  EcoleLOGO.Picture = LoadPicture(Dir)
  FileCopy Dir, App.Path & "\EcoleLogo.jpg"
  Exit Sub
z:
  MsgBox "Image Non Supporté !", vbCritical
End Sub

Private Sub chkRun_Click()
  SetRunAtStartup App.EXEName, App.Path, (chkRun.Value = vbChecked)
End Sub

Private Sub ChkFull_Click()
  '-------------------
  If ChkFull.Value = vbChecked Then
    Change_Setting "FullScreen", "Oui"
  Else
    Change_Setting "FullScreen", "Non"
  End If
  '--------------------
End Sub
Private Sub ChkSusp_Click()
  '--------------------
  If ChkSusp.Value = vbChecked Then
    Change_Setting "Susp_Ins", "Oui"
  Else
    Change_Setting "Susp_Ins", "Non"
  End If
  '---------------------
End Sub

Private Sub ChkConf_Click()
  '--------------------
  If ChkConf.Value = vbChecked Then
    Change_Setting "Email_Conf", "Oui"
  Else
    Change_Setting "Email_Conf", "Non"
  End If
  '--------------------
End Sub

Function GetParam()
'Verifie Les Parametres
    '-----------------------------------
    TxtEcoleNom.Text = GetSetting("Ecole_nom")
    TxtEcoleFormJ.Text = GetSetting("Ecole_Forme_Juridique")
    TxtEcoleTEL.Text = GetSetting("Ecole_TEL")
    TxtEcoleFax.Text = GetSetting("Ecole_Fax")
    TxtEcoleEmail.Text = GetSetting("Ecole_Email")
    TxtEcoleSiege.Text = GetSetting("Ecole_Siege")
    '-----------------------------------
    If WillRunAtStartup(App.EXEName) Then
        chkRun.Value = vbChecked
    Else
        chkRun.Value = vbUnchecked
    End If
    '-----------------------------------
    If GetSetting("FullScreen") = "Oui" Then
        ChkFull.Value = vbChecked
    Else
        ChkFull.Value = vbUnchecked
    End If
    '------------------------------------
    If GetSetting("Susp_Ins") = "Oui" Then
        ChkSusp.Value = vbChecked
    Else
        ChkSusp.Value = vbUnchecked
    End If
    '------------------------------------
    If GetSetting("Email_Conf") = "Oui" Then
        ChkConf.Value = vbChecked
    Else
        ChkConf.Value = vbUnchecked
    End If
    '------------------------------------
End Function

Function SaveParam()
  '---------------
  Change_Setting "Ecole_nom", TxtEcoleNom.Text
  Change_Setting "Ecole_Forme_Juridique", TxtEcoleFormJ.Text
  Change_Setting "Ecole_TEL", TxtEcoleTEL.Text
  Change_Setting "Ecole_Fax", TxtEcoleFax.Text
  Change_Setting "Ecole_Email", TxtEcoleEmail.Text
  Change_Setting "Ecole_Siege", TxtEcoleSiege.Text
  '---------------
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  SaveParam
End Sub

'##################### HALIDOU CISSE ##################
