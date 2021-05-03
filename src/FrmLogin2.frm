VERSION 5.00
Begin VB.Form FrmLogin2 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Authentification"
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8550
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmLogin2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FrameLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   1560
      Picture         =   "FrmLogin2.frx":27A2
      ScaleHeight     =   1335
      ScaleWidth      =   5295
      TabIndex        =   1
      Top             =   1320
      Width           =   5295
      Begin VB.TextBox TextPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   3
         ToolTipText     =   "Code Secret"
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox TextEmail 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         ToolTipText     =   "Email ou Nom d'utilisateur"
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mot de passe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Timer TimerCharg 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   1440
      Top             =   3960
   End
   Begin VB.PictureBox P2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   8265
      TabIndex        =   9
      Top             =   4320
      Visible         =   0   'False
      Width           =   8295
   End
   Begin VB.Timer TimerWait 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   960
      Top             =   3960
   End
   Begin VB.Timer TimerLoading 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   360
      Top             =   3960
   End
   Begin VB.CommandButton CmdLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Connexion >"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Connecter"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton CmdQuitter 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quitter"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Quitter"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.PictureBox LoadBar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "FrmLogin2.frx":2597F
      ScaleHeight     =   495
      ScaleWidth      =   8295
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   8295
   End
   Begin VB.Label Labelcharg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "chargement ......"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   3600
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Image Image 
      Appearance      =   0  'Flat
      Height          =   780
      Left            =   120
      Picture         =   "FrmLogin2.frx":48B5C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Gestion Des Inscriptions d'Etudiants"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   600
      Left            =   480
      TabIndex        =   0
      Top             =   6000
      Width           =   7965
   End
End
Attribute VB_Name = "FrmLogin2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Turn As String
Private Essaie As Integer
Private T As Integer
Private Pas As Integer

Private Sub CmdLogin_Click()
'Click Sur Le Boutton Connecter
'-------------------------
If TextPass.Text = Empty Then
  TextPass.SetFocus
Exit Sub
End If
'-------------------------
Dim cNET
cNET = Connect_Profil(Trim(TextEmail.Text), TextPass.Text)

  If cNET = "Connecter" Then
    TextEmail.Text = ""
    TextPass.Text = ""
    Loading 50
  ElseIf cNET = "Inactif" Then
    MsgBox "Desolé votre Compte est Inactif " & vbNewLine & "Contacter L'Administration !", vbCritical
  ElseIf cNET = "Pass Incorrect" Then
        MsgBox "Login ou Mot De Passe  Incorrect !"
        '-----------------------------------------
        If Essaie Mod 3 = 0 And Essaie <> 0 Then
            '----------------------------------------------
            x = MsgBox("Voulez Vous Reinitialiser Votre Mot de Passe ?", vbQuestion + vbYesNo)
            If x = vbYes Then
              If ProfilExist(TextEmail.Text) Then
                If EnvoiEMail(GetProfil_RecEmail(TextEmail.Text), "MIAGE RABAT : Récuperation de Votre Mot de Passe ", "Mot de Passe :  " & GetProfilPass(TextEmail.Text)) Then
                  MsgBox "Un Email A éte Envoyer à Votre Adresse Email de Recuperation"
                End If
              Else
                MsgBox "Adresse Email Inexistant  !!!! "
                TextPass = Empty
                TextPass.SetFocus
              End If
            End If
            '-----------------------------------------------
            Essaie = Essaie + 1
        Else
            '------------------
             Essaie = Essaie + 1
             TextPass = Empty
             TextPass.SetFocus
            '------------------
        End If
  End If
End Sub

Private Function Loading(Second As Integer)
  T = Second
  Pas = 100
  Turn = "va"
    '--------------------------
    If G_FullScreen = True Then
      LoadState 0
      FrmLogin2.Hide
      Main.Show 1
    Else
      LoadState 1
    End If
    '--------------------------
End Function

Private Sub TimerCharg_Timer()
  If Labelcharg.Visible = False Then
    Labelcharg.Visible = True
  Else
    Labelcharg.Visible = False
  End If
End Sub

Private Sub TimerWait_Timer()
    T = T - 1
    If T <= 0 Then
         LoadState 0
         '------------------
         FrmLogin2.Hide
         Main.Show 1
         '------------------
    End If
End Sub

Private Sub TimerLoading_Timer()
    AnimBar
End Sub

Function LoadState(i As Integer)
  If i = 1 Then
    'Pour Animation
    TimerLoading.Enabled = True
    TimerWait.Enabled = True
    TimerCharg.Enabled = True
    Labelcharg.Visible = True
    FrameLogin.Visible = False
    CmdLogin.Visible = False
    LoadBar.Visible = True
  Else
   'Pour Login
    CmdQuitter.Visible = True
    CmdLogin.Visible = True
    TimerLoading.Enabled = False
    TimerWait.Enabled = False
    TimerCharg.Enabled = False
    Labelcharg.Visible = False
    FrameLogin.Visible = True
    LoadBar.Visible = False
  End If
End Function

Private Function AnimBar()
'Anime chargement des données
  '------------------
  If Turn = "va" Then
    '---------------
    If LoadBar.Width <= P2.Width - Pas Then
       LoadBar.Width = LoadBar.Width + Pas
    Else
       Turn = "vient"
       LoadBar.Width = P2.Width
    End If
    '--------------
 Else
    '------------
    If LoadBar.Width >= Pas Then
       LoadBar.Width = LoadBar.Width - Pas
    Else
       Turn = "va"
       LoadBar.Width = 0
    End If
    '------------
 End If
 '-------------------
LoadBar.Left = P2.Left + (P2.Width - LoadBar.Width) / 2
End Function

Private Sub CmdQuitter_Click()

If G_FullScreen = True Then
  MsgBox "Vous Devez Vous Connecter Pour Pouvoir Quitter !", vbCritical + vbOKOnly
  Exit Sub
End If
   '---------------
   x = MsgBox("Voulez Vous Quitter Definitivement ?", vbQuestion + vbYesNo)
   If x = vbYes Then
      Unload FrmBkg
      Unload Me
     End
   End If
   '--------------
End Sub

Private Sub Form_Activate()
    If G_FullScreen = True Then
      Me_Top Me
    End If
   LoadState 0
   FrameLogin.Visible = True
   TextEmail.SetFocus
   '----------------
   If NombreEnreg("PROFILES", "SQL", "SELECT COUNT(*) AS Nombre FROM PROFILES WHERE Statut = 'Actif'") = 0 Then
     If Connect_Default_Profil Then
        Loading 50
     End If
   End If
   '---------------
End Sub

Private Sub Form_Load()
    ConnectDB
    LoadBar.Top = 1800
    LoadBar.Width = 0
    Essaie = 1
    
    '-------------------------------------
    If GetSetting("FullScreen") = "Oui" Then
      TakeOver_At_Run
    End If
    '------------------------------------
    
End Sub

Function TakeOver_At_Run()
'Bloquer L'Ecran, No Way Out
    
    Me_FullScreen FrmBkg
    FrmBkg.Show
    
    Me_Top Me
    Me.Show 1
    
End Function

Private Sub Form_Unload(Cancel As Integer)
  UnloadForms
End Sub

'##################### HALIDOU CISSE ##################




