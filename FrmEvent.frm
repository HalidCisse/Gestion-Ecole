VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmGridEvent 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evènements"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10830
   Icon            =   "FrmEvent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   10830
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.ComboBox Combo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6120
      Sorted          =   -1  'True
      TabIndex        =   3
      Text            =   "Choisir.."
      ToolTipText     =   "Choisir"
      Top             =   240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ComboBox ComboRecItem 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Selectionez Un Champs"
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton CmdAfficher 
      Appearance      =   0  'Flat
      Caption         =   "Afficher >>"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridEVent 
      Bindings        =   "FrmEvent.frx":27A2
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10550
      _ExtentX        =   18600
      _ExtentY        =   8070
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   33023
      BackColorBkg    =   16777215
      GridColor       =   -2147483635
      GridColorFixed  =   255
      GridLinesFixed  =   1
      ScrollBars      =   2
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   5
      _Band(0)._MapCol(0)._Name=   "N°"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(0)._Alignment=   7
      _Band(0)._MapCol(0)._Hidden=   -1  'True
      _Band(0)._MapCol(1)._Name=   "Description"
      _Band(0)._MapCol(1)._RSIndex=   2
      _Band(0)._MapCol(2)._Name=   "Email_Profile"
      _Band(0)._MapCol(2)._Caption=   "Profil Utilisateur"
      _Band(0)._MapCol(2)._RSIndex=   1
      _Band(0)._MapCol(3)._Name=   "EventDate"
      _Band(0)._MapCol(3)._Caption=   "Date"
      _Band(0)._MapCol(3)._RSIndex=   4
      _Band(0)._MapCol(4)._Name=   "EventTime"
      _Band(0)._MapCol(4)._Caption=   "Heure"
      _Band(0)._MapCol(4)._RSIndex=   3
   End
   Begin VB.Label LabelSign 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   4755
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label LabelResult 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   9960
      TabIndex        =   4
      ToolTipText     =   "Nombres de Résultat"
      Top             =   360
      Width           =   630
   End
   Begin VB.Menu PopUp 
      Caption         =   "&PupUpMenu"
      Visible         =   0   'False
      Begin VB.Menu PupSup 
         Caption         =   "&Supprimé"
      End
      Begin VB.Menu PupMod 
         Caption         =   "&Modifié"
      End
   End
End
Attribute VB_Name = "FrmGridEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ResChamps As String
Private Matricule As String

Private Sub CmdAfficher_Click()
  Afficher
End Sub

Private Sub Combo_Change()
  Afficher
End Sub

Private Sub Combo_Validate(Cancel As Boolean)
  Afficher
End Sub

Private Sub ComboRecItem_Validate(Cancel As Boolean)
On Error Resume Next
  Combo.Visible = False
  Remplir Combo, "EVENTS", ResChamps
  Combo.ListIndex = 0
  Combo.Visible = True
  Afficher
End Sub

Private Sub Form_Activate()
  Afficher
End Sub

Private Sub Form_Load()
  ScaleGrid
  RemplirItems
  Afficher
If G_FullScreen = True Then
     Me_Top Me
End If
End Sub

Private Function Afficher()
  RemplirGrid FrmGridEvent, FrmGridEvent.GridEVent, CalculeSQL
End Function

Private Function RemplirItems()
'Remplir les combos des données necessaires
With ComboRecItem
.Clear
     .AddItem ("Tous")
     .AddItem ("Description")
     .AddItem ("Profile D'Utilisateur")
     .AddItem ("Date de L'Evenement")
.Refresh
.ListIndex = 0
End With
End Function

Private Function CalculeSQL() As String
'Calculer le SQL coresspondant aux criteres
    If ComboRecItem = "Tous" Then
       ResChamps = "Email_Profile"
       LabelSign.Visible = False
       Combo.Visible = False
       CalculeSQL = "SELECT * FROM EVENTS ORDER BY CDate(EventDate) DESC, CDate(EventTime) DESC"
       LabelResult = NombreEnreg("EVENTS", "SQL", "SELECT COUNT(*) as Nombre FROM EVENTS") & " EVENEMENTS"
    Else
       LabelSign.Visible = True
       Combo.Visible = True
     If ComboRecItem = "Description" Then
       ResChamps = "Description"
     ElseIf ComboRecItem = "Profile D'Utilisateur" Then
       ResChamps = "Email_Profile"
     ElseIf ComboRecItem = "Date de L'Evenement" Then
       ResChamps = "EventDate"
     End If
     CalculeSQL = "SELECT * FROM EVENTS WHERE " & ResChamps & " Like '%" & UCase(Trim(Combo)) & "%' ORDER BY CDate(EventDate) DESC, CDate(EventTime) DESC"
     LabelResult = NombreEnreg("EVENTS", "SQL", "SELECT COUNT(*) as Nombre FROM EVENTS WHERE " & ResChamps & " Like '%" & UCase(Trim(Combo)) & "%' ") & " EVENEMENTS"
   End If
End Function

Private Sub PupSup_Click()
Dim INS
With GridEVent
  INS = .TextMatrix(.RowSel, 0)
  MsgBox INS
End With
End Sub

Private Sub Timer_Timer()
  CalculeSQL
  
End Sub

Private Function ScaleGrid()
With GridEVent
  .ColWidth(0) = 6000
  .ColWidth(1) = 2000
  .ColWidth(2) = 1000
  .ColWidth(3) = 1300

  .Width = 10550
End With
End Function

'##################### HALIDOU CISSE ##################
