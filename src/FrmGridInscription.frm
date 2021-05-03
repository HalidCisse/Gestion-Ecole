VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmGridInscription 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INSCRIPTIONS"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10740
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   21.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmGridInscription.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   10740
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton OptDESC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "DESC"
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
      Height          =   225
      Left            =   6840
      TabIndex        =   9
      ToolTipText     =   "Trier Par Nom Descendant"
      Top             =   5400
      Width           =   735
   End
   Begin VB.OptionButton OptASC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ASC"
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
      Height          =   225
      Left            =   5280
      TabIndex        =   8
      ToolTipText     =   "Trier Par Nom Ascendant"
      Top             =   5400
      Width           =   735
   End
   Begin VB.OptionButton OptRecent 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Recent"
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
      Height          =   225
      Left            =   3720
      TabIndex        =   7
      ToolTipText     =   "Trier Par Les Plus Recents"
      Top             =   5400
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   0
      Top             =   0
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
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Selectionez Un Champs"
      Top             =   240
      Width           =   2415
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
      TabIndex        =   0
      ToolTipText     =   "Choisir"
      Top             =   240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridIns 
      Bindings        =   "FrmGridInscription.frx":27A2
      Height          =   4575
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Double Click Pour Modifier"
      Top             =   720
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   8070
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   7
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
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   11
      _Band(0)._MapCol(0)._Name=   "N°_Matricule_Etudiant"
      _Band(0)._MapCol(0)._RSIndex=   1
      _Band(0)._MapCol(1)._Name=   "N°_Inscription"
      _Band(0)._MapCol(1)._RSIndex=   0
      _Band(0)._MapCol(2)._Name=   "Classe"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "Ins_Debut"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(3)._Hidden=   -1  'True
      _Band(0)._MapCol(4)._Name=   "Ins_Fin"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(4)._Hidden=   -1  'True
      _Band(0)._MapCol(5)._Name=   "Annnee_Scolaire"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(6)._Name=   "Payment_Ins"
      _Band(0)._MapCol(6)._RSIndex=   6
      _Band(0)._MapCol(6)._Hidden=   -1  'True
      _Band(0)._MapCol(7)._Name=   "Type_Reglement"
      _Band(0)._MapCol(7)._RSIndex=   7
      _Band(0)._MapCol(8)._Name=   "Payment_Par_Tranche"
      _Band(0)._MapCol(8)._RSIndex=   8
      _Band(0)._MapCol(8)._Hidden=   -1  'True
      _Band(0)._MapCol(9)._Name=   "Date_Inscription"
      _Band(0)._MapCol(9)._RSIndex=   9
      _Band(0)._MapCol(10)._Name=   "Statut"
      _Band(0)._MapCol(10)._RSIndex=   10
   End
   Begin VB.Label LabelSign 
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
      Left            =   5040
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image LPrint 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   9960
      Picture         =   "FrmGridInscription.frx":27B7
      Stretch         =   -1  'True
      ToolTipText     =   "Imprimé"
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   120
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Inscription Expiré"
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
      TabIndex        =   6
      Top             =   5400
      Width           =   1395
   End
   Begin VB.Label LabelResult 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   10125
      TabIndex        =   5
      ToolTipText     =   "Nombres d'Inscriptions"
      Top             =   5400
      Width           =   390
   End
   Begin VB.Label LabelSignx 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5325
      TabIndex        =   4
      Top             =   -480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Menu PopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu PopModIns 
         Caption         =   "&Details"
      End
      Begin VB.Menu PopSupIns 
         Caption         =   "&Supprimé"
      End
      Begin VB.Menu mnImpDossier 
         Caption         =   "&Imprimé Dossier"
         Index           =   2
      End
   End
End
Attribute VB_Name = "FrmGridInscription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ResChamps As String
Private Inscription As String

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
  Remplir Combo, "INSCRIPTIONS", ResChamps
  Combo.ListIndex = 0
  Combo.Visible = True
  Afficher
End Sub

Private Sub Form_Activate()
  Afficher
End Sub

Private Sub Form_Load()
  RemplirNomEtudiant
  ScaleGrid
  RemplirItems
  If G_FullScreen = True Then
     Me_Top Me
     LPrint.Visible = False
  End If
End Sub

Private Function Afficher()
'Affiche La Grid Correspondant a La Grid
  RemplirGrid FrmGridInscription, FrmGridInscription.GridIns, CalculeSQL
  '-------------
  GetEtudiantsName
  GetRowBackColor
  '-------------
End Function

Private Function GetEtudiantsName()
'Permet De Trouver Le Nom De L'Etudiant Correspondant
On Error Resume Next
Dim i
GridIns.TextMatrix(0, 0) = "Nom de L'Etudiant"
  For i = 1 To Val(LabelResult.Caption)
    GridIns.TextMatrix(i, 0) = GetMyName(GridIns.TextMatrix(i, 0))
  Next i
End Function

Private Function RemplirItems()
'Remplir les combos des données necessaires
With ComboRecItem
.Clear
     .AddItem ("Tous")
     .AddItem ("Etudiant")
     .AddItem ("N° Inscription")
     .AddItem ("N° Matricule")
     .AddItem ("Classe")
     .AddItem ("Année Scolaire")
     .AddItem ("STATUT")
     .AddItem ("Type de Reglement")
     .AddItem ("Date d'Inscription")
.Refresh
.ListIndex = 0
End With
End Function

Private Function CalculeSQL() As String
'Calculer le SQL coresspondant aux criteres
    If ComboRecItem = "Tous" Then
       ResChamps = "N°_Inscription"
       LabelSign.Visible = False
       Combo.Visible = False
       CalculeSQL = "SELECT * FROM INSCRIPTIONS "
       LabelResult = NombreEnreg("INSCRIPTIONS", "SQL", "SELECT COUNT(*) as Nombre FROM INSCRIPTIONS") & " Inscriptions"
    Else
       LabelSign.Visible = True
       Combo.Visible = True
     If ComboRecItem = "N° Inscription" Then
       ResChamps = "N°_Inscription"
     ElseIf ComboRecItem = "Etudiant" Then
       ResChamps = "NomEtudiant"
     ElseIf ComboRecItem = "N° Matricule" Then
       ResChamps = "N°_Matricule_Etudiant"
     ElseIf ComboRecItem = "Classe" Then
       ResChamps = "Classe"
       CalculeSQL = "SELECT * FROM INSCRIPTIONS WHERE Statut = 'Active' AND Classe Like '%" & UCase(Trim(Combo)) & "%' "
       LabelResult = NombreEnreg("INSCRIPTIONS", "SQL", "SELECT COUNT(*) as Nombre FROM INSCRIPTIONS WHERE Statut = 'Active' AND  Classe Like '%" & UCase(Trim(Combo)) & "%' ") & " Inscriptions"
       Exit Function
     ElseIf ComboRecItem = "Année Scolaire" Then
       ResChamps = "Annnee_Scolaire"
     ElseIf ComboRecItem = "Type de Reglement" Then
       ResChamps = "Type_Reglement"
     ElseIf ComboRecItem = "Date d'Inscription" Then
       ResChamps = "Date_Inscription"
     ElseIf ComboRecItem = "STATUT" Then
       ResChamps = "Statut"
     End If
     CalculeSQL = "SELECT * FROM INSCRIPTIONS WHERE " & ResChamps & " Like '%" & UCase(Trim(Combo)) & "%' "
     LabelResult = NombreEnreg("INSCRIPTIONS", "SQL", "SELECT COUNT(*) as Nombre FROM INSCRIPTIONS WHERE  " & ResChamps & " Like '%" & UCase(Trim(Combo)) & "%' ") & " Inscriptions"
   End If
   '-----------------------------
   If OptRecent.Value = True Then
     CalculeSQL = CalculeSQL & "ORDER BY Date_Inscription DESC"
   ElseIf OptASC.Value = True Then
     CalculeSQL = CalculeSQL & "ORDER BY NomEtudiant ASC"
   ElseIf OptDESC.Value = True Then
     CalculeSQL = CalculeSQL & "ORDER BY NomEtudiant DESC"
   End If
   '----------------------------
End Function

Private Sub GridIns_DblClick()
'Si on Double Click Sur La Grid Ouvrir La FrmInscription Contenant Les Infos de L'Etudiant Selectionner
  If GridIns.Row > 0 Then
    Inscription = GridIns.TextMatrix(GridIns.RowSel, 1)
    OpenModIns (CStr(Inscription))
  End If
End Sub

Private Sub LPrint_Click()
'Imprimer La Grid
 If ComboRecItem = "Tous" Then
   Open_DR_Ins "SELECT Nom, Prenom, N°_Matricule_Etudiant, Classe, N°_inscription FROM ETUDIANT_INSCRIT ORDER BY Date_Inscription DESC"
 ElseIf ComboRecItem = "Classe" Then
   Open_DR_Ins "SELECT Nom, Prenom, N°_Matricule_Etudiant, Classe, N°_inscription FROM ETUDIANT_INSCRIT WHERE Statut = 'Active' AND Classe Like '%" & UCase(Trim(Combo)) & "%' ORDER BY Date_Inscription DESC"
 Else
   Open_DR_Ins "SELECT Nom, Prenom, N°_Matricule_Etudiant, Classe, N°_inscription FROM ETUDIANT_INSCRIT WHERE " & ResChamps & " Like '%" & UCase(Trim(Combo)) & "%' ORDER BY Date_Inscription DESC"
 End If
End Sub

Private Sub OptASC_Click()
  Afficher
End Sub

Private Sub OptDESC_Click()
  Afficher
End Sub

Private Sub OptRecent_Click()
  Afficher
End Sub

Private Sub Timer_Timer()
  CalculeSQL
End Sub

Private Function ScaleGrid()
With GridIns
  .ColWidth(0) = 2500
  .ColWidth(1) = 1500
  .ColWidth(2) = 2000
  .ColWidth(3) = 1300
  .ColWidth(4) = 1000
  .ColWidth(5) = 1000
  .ColWidth(6) = 1000
  '.Width = 9850
End With
End Function


Private Sub GridIns_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Affiche un PopUp
  If Button = 2 Then
    PopupMenu PopUp
  End If
End Sub

Private Sub PopModIns_Click()
'Si on Click Sur Modifier Ouvrir La FrmInscription Contenant Les Infos de L'Etudiant Selectionner
  OpenModIns (GridIns.TextMatrix(GridIns.RowSel, 1))
  Afficher
End Sub

Private Sub PopSupIns_Click()
'Si on Click Sur Supprimer Ouvrir La FrmInscription Contenant Les Infos de L'Etudiant Selectionner
  OpenSupIns (GridIns.TextMatrix(GridIns.RowSel, 1))
  Afficher
End Sub

Private Sub mnImpDossier_Click(Index As Integer)
  If Not G_FullScreen Then
    Open_DR_Dos_Ins (GridIns.TextMatrix(GridIns.RowSel, 1))
  Else
    MsgBox "Vous Devez Quitter FullScreen Pour Pouvoir Imprimer !"
  End If
End Sub

Private Function RemplirNomEtudiant()
'Pour Satisfaire La Grid
  Dim gSQL
  Dim rs As New ADODB.Recordset
  gSQL = "SELECT N°_Inscription, NomEtudiant FROM INSCRIPTIONS"
  rs.Open gSQL, CN, adOpenDynamic, adLockOptimistic, adCmdText
  
   Do While Not rs.EOF
     rs![NomEtudiant] = GetMyName(GetMaMatricule(rs![N°_Inscription]))
   rs.MoveNext
   Loop
   
   rs.Close
   Set rs = Nothing
End Function

Private Function GetRowBackColor()
'Changer La Couleur Des Etudiants Active = Green, Expiré = Red
Dim lngRow As Long
Dim lngColour As Long
 
With GridIns
   .Redraw = False
    .FillStyle = flexFillRepeat
 
    For lngRow = .FixedRows To .Rows - 1
        Select Case .TextMatrix(lngRow, 6)
            Case "Expiré"
                lngColour = vbRed
            Case "Active"
                lngColour = vbBlue
            Case Else
                lngColour = vbGreen
        End Select
        .Row = lngRow
        .ColSel = .Cols - 1
        .CellForeColor = lngColour
    Next
    .FillStyle = flexFillSingle
   .Redraw = True
End With
End Function

'##################### HALIDOU CISSE ##################

