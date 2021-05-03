VERSION 5.00
'Object = "AcroExch.Document.11"; "AcroRd32.exe""
Begin VB.Form FrmMessage 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "RAPPORT DU PROJET FIN D'ANNEE"
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8940
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   8940
   StartUpPosition =   1  'CenterOwner
   Begin AcroExchCtl.Document Document1 
      Height          =   12630
      Left            =   0
      OleObjectBlob   =   "FrmMessage.frx":74F2
      TabIndex        =   0
      ToolTipText     =   "Double Click Pour Ouvrir"
      Top             =   0
      Width           =   8925
   End
End
Attribute VB_Name = "FrmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
