VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmGridProfil 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PROFILES D'UTILISATEURS"
   ClientHeight    =   5550
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12465
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmGridProfil.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   12465
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdAfficher 
      Appearance      =   0  'Flat
      Caption         =   "Afficher >>"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox ComboRecItem 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Selectionez Un Champs"
      Top             =   120
      Width           =   2415
   End
   Begin VB.ComboBox Combo 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   6720
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Choisir.."
      ToolTipText     =   "Choisir"
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   0
      Top             =   120
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Bindings        =   "FrmGridProfil.frx":27A2
      Height          =   4575
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   8070
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   33023
      BackColorBkg    =   16777215
      GridColor       =   -2147483635
      GridColorFixed  =   255
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
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
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   7
      _Band(0)._MapCol(0)._Name=   "Login"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "Nom"
      _Band(0)._MapCol(1)._RSIndex=   2
      _Band(0)._MapCol(2)._Name=   "Prenom"
      _Band(0)._MapCol(2)._RSIndex=   3
      _Band(0)._MapCol(3)._Name=   "AdressEmail"
      _Band(0)._MapCol(3)._Caption=   "Email De Recupération"
      _Band(0)._MapCol(3)._RSIndex=   4
      _Band(0)._MapCol(4)._Name=   "UserType"
      _Band(0)._MapCol(4)._Caption=   "Type Profile"
      _Band(0)._MapCol(4)._RSIndex=   5
      _Band(0)._MapCol(5)._Name=   "Statut"
      _Band(0)._MapCol(5)._RSIndex=   6
      _Band(0)._MapCol(6)._Name=   "Pass"
      _Band(0)._MapCol(6)._RSIndex=   1
      _Band(0)._MapCol(6)._Hidden=   -1  'True
   End
   Begin VB.Label LabBLQ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Bloqué"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   11400
      TabIndex        =   8
      Top             =   5280
      Width           =   540
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   10800
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label LabelAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "ADMIN"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   720
      TabIndex        =   7
      Top             =   5280
      Width           =   555
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   120
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label LabelSTD 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "STANDARD"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   6360
      TabIndex        =   6
      Top             =   5280
      Width           =   870
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   5760
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label LabelResult 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Left            =   11865
      TabIndex        =   5
      ToolTipText     =   "Nombres de Résultat"
      Top             =   240
      Width           =   405
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
      Left            =   5400
      TabIndex        =   4
      Top             =   -120
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Menu PopUp 
      Caption         =   "&PopUp"
      Visible         =   0   'False
      Begin VB.Menu PopMod 
         Caption         =   "&Modifié"
      End
      Begin VB.Menu PopSup 
         Caption         =   "&Supprimé"
      End
      Begin VB.Menu PopDesc 
         Caption         =   "&Desactivé"
      End
   End
End
Attribute VB_Name = "FrmGridProfil"
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
  On Error Resume Next
  Afficher
End Sub

Private Sub Combo_Validate(Cancel As Boolean)
  Afficher
End Sub

Private Sub ComboRecItem_Validate(Cancel As Boolean)
On Error Resume Next
  Combo.Visible = False
  Remplir Combo, "PROFILES", ResChamps
  Combo.ListIndex = 0
  Combo.Visible = True
  Afficher
End Sub

Private Sub Form_Activate()
  Afficher
  GetLabel
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
  Call RemplirGrid(FrmGridProfil, FrmGridProfil.Grid, CalculeSQL)
  GetRowBackColor
End Function

Private Function RemplirItems()
'Remplir les combos des données necessaires
With ComboRecItem
.Clear
     .AddItem ("Tous")
     .AddItem ("Profil Utilisateur")
     .AddItem ("Prenom")
     .AddItem ("Nom")
     .AddItem ("Type d'Utilisateur")
     .AddItem ("Statut")
.Refresh
.ListIndex = 0
End With
End Function

Private Function CalculeSQL() As String
'Calculer le SQL coresspondant aux criteres
    If ComboRecItem = "Tous" Then
       ResChamps = "Login"
       LabelSign.Visible = False
       Combo.Visible = False
       CalculeSQL = "SELECT * FROM PROFILES ORDER BY Login DESC"
       LabelResult = NombreEnreg("PROFILES", "SQL", "SELECT COUNT(*) as Nombre FROM PROFILES") & " Profiles"
    Else
      '----------------------------------------
       LabelSign.Visible = True
       Combo.Visible = True
     If ComboRecItem = "Profil Utilisateur" Then
       ResChamps = "Login"
     ElseIf ComboRecItem = "Prenom" Then
       ResChamps = "Prenom"
     ElseIf ComboRecItem = "Nom" Then
       ResChamps = "Nom"
     ElseIf ComboRecItem = "Type d'Utilisateur" Then
       ResChamps = "UserType"
     ElseIf ComboRecItem = "Statut" Then
       ResChamps = "Statut"
     End If
     '-----------------------------------------
     CalculeSQL = "SELECT * FROM PROFILES WHERE " & ResChamps & " Like '%" & UCase(Trim(Combo)) & "%' ORDER BY Login DESC"
     LabelResult = NombreEnreg("PROFILES", "SQL", "SELECT COUNT(*) as Nombre FROM PROFILES WHERE " & ResChamps & " Like '%" & UCase(Trim(Combo)) & "%' ") & " Profiles"
   End If
End Function

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
' Click Droit
  If Button = 2 Then
    PopupMenu PopUp
  End If
End Sub

Private Sub PopDesc_Click()
  'Desactiver Un Profil
  If Desctiver_Profil(Grid.TextMatrix(Grid.RowSel, 0)) Then
    MsgBox "Profil Desactivé Avec Succès !", vbInformation
  Else
    MsgBox "Erreur !", vbCritical
  End If
  Afficher
End Sub

Private Sub PopMod_Click()
  OpenModProfil Grid.TextMatrix(Grid.RowSel, 0)
  Afficher
End Sub

Private Sub PopSup_Click()
  OpenSupProfil Grid.TextMatrix(Grid.RowSel, 0)
  Afficher
End Sub

Private Sub Timer_Timer()
  CalculeSQL
End Sub

Private Function ScaleGrid()
With Grid
  .ColWidth(0) = 3000
  .ColWidth(1) = 2000
  .ColWidth(2) = 1000
  .ColWidth(3) = 3000
  .ColWidth(4) = 2000
  .ColWidth(5) = 1000
  
  .Width = 12200
End With
End Function

Function GetLabel()
  
  LabelAdmin.Caption = "ADMIN (" & NombreEnreg("PROFILES", "UserType", "ADMIN") & ")"
  LabelSTD.Caption = "STANDARD (" & NombreEnreg("PROFILES", "UserType", "STANDARD") & ")"
  LabBLQ.Caption = "Bloqué (" & NombreEnreg("PROFILES", "Statut", "Inactif") & ")"
  
End Function

Private Function GetRowBackColor()
'Changer La Couleur Des Etudiants Regulier = Blue, Abandonner = Red, Diplomé = Green
Dim lngRow As Long
Dim lngColour As Long
 
With Grid
   .Redraw = False
    .FillStyle = flexFillRepeat
 
    For lngRow = .FixedRows To .Rows - 1
        Select Case .TextMatrix(lngRow, 4)
            Case "ADMIN"
                lngColour = vbBlue
            Case "STANDARD"
                lngColour = vbGreen
            Case Else
                lngColour = vbRed
        End Select
        If .TextMatrix(lngRow, 5) = "Inactif" Then
          lngColour = vbRed
        End If
        
        .Row = lngRow
        .ColSel = .Cols - 1
        .CellForeColor = lngColour
    Next
    .FillStyle = flexFillSingle
   .Redraw = True
End With

End Function

'##################### HALIDOU CISSE ##################

