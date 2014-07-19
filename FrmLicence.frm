VERSION 5.00
Begin VB.Form FrmLicence 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Licence d'utilisation"
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10755
   Icon            =   "FrmLicence.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdFermer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "fermer"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox TextLicence 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FrmLicence.frx":2AD05
      Top             =   120
      Width           =   10455
   End
End
Attribute VB_Name = "FrmLicence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFermer_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Main.WindowState = 1
  If G_FullScreen = True Then
     Me_Top Me
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Main.WindowState = 0
End Sub

'##################### HALIDOU CISSE ##################
