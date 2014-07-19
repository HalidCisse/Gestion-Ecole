VERSION 5.00
Object = "AcroExch.Document.11"; "AcroRd32.exe""
Begin VB.Form FrmRapport 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Rapport Du Projet de Fin D'Année"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17610
   Icon            =   "FrmRapport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   17610
   StartUpPosition =   1  'CenterOwner
   Begin AcroExchCtl.Document Document2 
      Height          =   8100
      Left            =   9120
      OleObjectBlob   =   "FrmRapport.frx":014A
      TabIndex        =   1
      ToolTipText     =   "Clicker Pour Ouvrir"
      Top             =   0
      Width           =   14400
   End
   Begin AcroExchCtl.Document Document1 
      Height          =   12630
      Left            =   0
      OleObjectBlob   =   "FrmRapport.frx":525D62
      TabIndex        =   0
      ToolTipText     =   "Clicker Pour Ouvrir"
      Top             =   0
      Width           =   8925
   End
End
Attribute VB_Name = "FrmRapport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

