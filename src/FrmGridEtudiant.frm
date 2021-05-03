VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmGridEtudiant 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LISTE DES ETUDIANTS"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15525
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmGridEtudiant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   15525
   StartUpPosition =   1  'CenterOwner
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
      Left            =   4440
      TabIndex        =   10
      ToolTipText     =   "Trier Par Les Plus Recents"
      Top             =   5400
      Value           =   -1  'True
      Width           =   975
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
      Left            =   6000
      TabIndex        =   9
      ToolTipText     =   "Trier Par Nom Ascendant"
      Top             =   5400
      Width           =   735
   End
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
      Left            =   7560
      TabIndex        =   8
      ToolTipText     =   "Trier Par Nom Descendant"
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton CmdAfficher 
      Appearance      =   0  'Flat
      Caption         =   "Afficher >>"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.ComboBox ComboRecItem 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Selectionez Un Champs"
      Top             =   240
      Width           =   3135
   End
   Begin VB.ComboBox Combo 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   9120
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Choisir"
      Top             =   240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFe 
      Bindings        =   "FrmGridEtudiant.frx":628A
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Double Click Pour Modifier"
      Top             =   720
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   8070
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   19
      FixedCols       =   0
      BackColorFixed  =   33023
      BackColorBkg    =   16777215
      GridColor       =   -2147483635
      GridColorFixed  =   255
      HighLight       =   2
      GridLinesFixed  =   1
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
      _Band(0).Cols   =   19
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   26
      _Band(0)._MapCol(0)._Name=   "TypeID"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(0)._Hidden=   -1  'True
      _Band(0)._MapCol(1)._Name=   "Prenom"
      _Band(0)._MapCol(1)._RSIndex=   5
      _Band(0)._MapCol(2)._Name=   "Nom"
      _Band(0)._MapCol(2)._RSIndex=   4
      _Band(0)._MapCol(3)._Name=   "Classe"
      _Band(0)._MapCol(3)._RSIndex=   22
      _Band(0)._MapCol(4)._Name=   "Matricule"
      _Band(0)._MapCol(4)._RSIndex=   20
      _Band(0)._MapCol(5)._Name=   "N°_Inscription"
      _Band(0)._MapCol(5)._Caption=   "N° Inscription"
      _Band(0)._MapCol(5)._RSIndex=   21
      _Band(0)._MapCol(6)._Name=   "Statut"
      _Band(0)._MapCol(6)._RSIndex=   25
      _Band(0)._MapCol(7)._Name=   "Date_Expire"
      _Band(0)._MapCol(7)._RSIndex=   2
      _Band(0)._MapCol(7)._Hidden=   -1  'True
      _Band(0)._MapCol(8)._Name=   "Sex"
      _Band(0)._MapCol(8)._RSIndex=   3
      _Band(0)._MapCol(8)._Hidden=   -1  'True
      _Band(0)._MapCol(9)._Name=   "Date_Naissance"
      _Band(0)._MapCol(9)._RSIndex=   6
      _Band(0)._MapCol(9)._Hidden=   -1  'True
      _Band(0)._MapCol(10)._Name=   "Lieu_Naissance"
      _Band(0)._MapCol(10)._RSIndex=   7
      _Band(0)._MapCol(10)._Hidden=   -1  'True
      _Band(0)._MapCol(11)._Name=   "Nom_Pere"
      _Band(0)._MapCol(11)._RSIndex=   9
      _Band(0)._MapCol(11)._Hidden=   -1  'True
      _Band(0)._MapCol(12)._Name=   "Nom_Mere"
      _Band(0)._MapCol(12)._RSIndex=   10
      _Band(0)._MapCol(12)._Hidden=   -1  'True
      _Band(0)._MapCol(13)._Name=   "TEL"
      _Band(0)._MapCol(13)._RSIndex=   11
      _Band(0)._MapCol(14)._Name=   "Email"
      _Band(0)._MapCol(14)._RSIndex=   12
      _Band(0)._MapCol(15)._Name=   "Adresse"
      _Band(0)._MapCol(15)._RSIndex=   13
      _Band(0)._MapCol(16)._Name=   "Nom_Tuteur"
      _Band(0)._MapCol(16)._Caption=   "Nom Tuteur"
      _Band(0)._MapCol(16)._RSIndex=   14
      _Band(0)._MapCol(17)._Name=   "Prenom_Tuteur"
      _Band(0)._MapCol(17)._Caption=   "Prenom Tuteur"
      _Band(0)._MapCol(17)._RSIndex=   15
      _Band(0)._MapCol(18)._Name=   "TEL_Tuteur"
      _Band(0)._MapCol(18)._Caption=   "TEL Tuteur"
      _Band(0)._MapCol(18)._RSIndex=   16
      _Band(0)._MapCol(19)._Name=   "Email_Tuteur"
      _Band(0)._MapCol(19)._Caption=   "Email Tuteur"
      _Band(0)._MapCol(19)._RSIndex=   17
      _Band(0)._MapCol(20)._Name=   "Adresse_Tuteur"
      _Band(0)._MapCol(20)._Caption=   "Adresse Tuteur"
      _Band(0)._MapCol(20)._RSIndex=   18
      _Band(0)._MapCol(21)._Name=   "Total_Payment"
      _Band(0)._MapCol(21)._Caption=   "Total Payement"
      _Band(0)._MapCol(21)._RSIndex=   23
      _Band(0)._MapCol(22)._Name=   "Numero_Identite"
      _Band(0)._MapCol(22)._Caption=   "Numero Identite"
      _Band(0)._MapCol(22)._RSIndex=   1
      _Band(0)._MapCol(23)._Name=   "Niveau_Etude"
      _Band(0)._MapCol(23)._Caption=   "Niveau Etude"
      _Band(0)._MapCol(23)._RSIndex=   19
      _Band(0)._MapCol(24)._Name=   "Nationalite"
      _Band(0)._MapCol(24)._RSIndex=   8
      _Band(0)._MapCol(25)._Name=   "Date_Enregistrement"
      _Band(0)._MapCol(25)._Caption=   "Date Registration"
      _Band(0)._MapCol(25)._RSIndex=   24
   End
   Begin VB.Image LPrint 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   14760
      Picture         =   "FrmGridEtudiant.frx":629F
      Stretch         =   -1  'True
      ToolTipText     =   "Imprimé"
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   10080
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label LabelDip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Etudiant Diplomé"
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
      Left            =   10680
      TabIndex        =   7
      Top             =   5400
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   120
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label LabelAb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Etudiant Abandonné"
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
      Width           =   1620
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
      Left            =   14880
      TabIndex        =   5
      ToolTipText     =   "Nombres de Résultat"
      Top             =   5400
      Width           =   390
   End
   Begin VB.Label LabelSign 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   7440
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Menu PopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu PopMod 
         Caption         =   "&Modifié"
      End
      Begin VB.Menu PopSup 
         Caption         =   "&Supprimé"
      End
      Begin VB.Menu mnImpDossier 
         Caption         =   "&Imprimé Dossier"
      End
   End
End
Attribute VB_Name = "FrmGridEtudiant"
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
  Remplir Combo, "ETUDIANTS", ResChamps
  Combo.ListIndex = 0
  Combo.Visible = True
  Afficher
End Sub

Private Sub Form_Activate()
  Afficher
End Sub

Private Sub Form_Load()
  If G_FullScreen = True Then
    Me_Top Me
    LPrint.Visible = False
  End If
'--------------
RemplirItems
LabelDip.Caption = "Etudiant Diplomés(" & GetEtudiantDiplomer & ")"
LabelAb.Caption = "Etudiant Abandonnés(" & GetEtudiantAbandonner & ")"
'--------------
End Sub

Private Function Afficher()
  RemplirGrid FrmGridEtudiant, FrmGridEtudiant.MSHFe, CalculeSQL
  '---------------
  GetRowBackColor
  '---------------
End Function

Private Function RemplirItems()
'Remplir les combos des données necessaires
With ComboRecItem
.Clear
     .AddItem ("Tous")
     .AddItem ("Nom")
     .AddItem ("Prénom")
     .AddItem ("N° Matricule")
     .AddItem ("Statut")
     .AddItem ("N° Identite")
     .AddItem ("Nationalité")
     .AddItem ("Type Identité")
     .AddItem ("Civilité")
     .AddItem ("N° Téléphone")
     .AddItem ("Adresse Email")
     .AddItem ("Adresse Domicile")
     .AddItem ("Date de Naissance")
     .AddItem ("Lieu de Naissance")
     .AddItem ("Nom du Père")
     .AddItem ("Nom de la Mère")
     .AddItem ("Nom du Tuteur")
     .AddItem ("Prénom du Tuteur")
     .AddItem ("N° Télephone du Tuteur")
     .AddItem ("Adresse Email du Tuteur")
     .AddItem ("Adresse Domicile du Tuteur")
     .AddItem ("Date Expiration de l'identité")
     .AddItem ("Date Enregistrement")
.Refresh
.ListIndex = 0
End With
End Function

Private Function CalculeSQL() As String
'Calculer le SQL coresspondant aux criteres
    If ComboRecItem = "Tous" Then
       ResChamps = "Matricule"
       LabelSign.Visible = False
       Combo.Visible = False
       CalculeSQL = "SELECT * FROM ETUDIANTS "
       LabelResult = NombreEnreg("ETUDIANT", "SQL", "SELECT COUNT(*) as Nombre FROM ETUDIANTS") & "  Total Etudiants"
    Else
       LabelSign.Visible = True
       Combo.Visible = True
     If ComboRecItem = "Type Identité" Then
       ResChamps = "TypeID"
     ElseIf ComboRecItem = "N° Matricule" Then
       ResChamps = "Matricule"
     ElseIf ComboRecItem = "N° Identite" Then
       ResChamps = "Numero_Identite"
     ElseIf ComboRecItem = "Date Expiration de l'identité" Then
       ResChamps = "Date_Expire"
     ElseIf ComboRecItem = "Civilité" Then
       ResChamps = "Sex"
     ElseIf ComboRecItem = "Nom" Then
       ResChamps = "Nom"
     ElseIf ComboRecItem = "Prénom" Then
       ResChamps = "Prenom"
     ElseIf ComboRecItem = "Date de Naissance" Then
       ResChamps = "Date_Naissance"
     ElseIf ComboRecItem = "Lieu de Naissance" Then
       ResChamps = "Lieu_Naissance"
     ElseIf ComboRecItem = "Nationalité" Then
       ResChamps = "Nationalite"
     ElseIf ComboRecItem = "Nom du Père" Then
       ResChamps = "Nom_Pere"
     ElseIf ComboRecItem = "Nom de la Mère" Then
       ResChamps = "Nom_Mere"
     ElseIf ComboRecItem = "N° Téléphone" Then
       ResChamps = "TEL"
     ElseIf ComboRecItem = "Adresse Email" Then
       ResChamps = "Email"
     ElseIf ComboRecItem = "Adresse Domicile" Then
       ResChamps = "Adresse"
     ElseIf ComboRecItem = "Nom du Tuteur" Then
       ResChamps = "Nom_Tuteur"
     ElseIf ComboRecItem = "Prénom du Tuteur" Then
       ResChamps = "Prenom_Tuteur"
     ElseIf ComboRecItem = "N° Télephone du Tuteur" Then
       ResChamps = "TEL_Tuteur"
     ElseIf ComboRecItem = "Adresse Email du Tuteur" Then
       ResChamps = "Email_Tuteur"
     ElseIf ComboRecItem = "Adresse Domicile du Tuteur" Then
       ResChamps = "Adresse_Tuteur"
     ElseIf ComboRecItem = "Date Enregistrement" Then
       ResChamps = "Date_Enregistrement"
     ElseIf ComboRecItem = "Statut" Then
       ResChamps = "Statut"
     End If
     
     CalculeSQL = "SELECT * FROM ETUDIANTS WHERE " & ResChamps & " Like '%" & UCase(Trim(Combo)) & "%' "
     LabelResult = NombreEnreg("ETUDIANTS", "SQL", "SELECT COUNT(*) as Nombre FROM ETUDIANTS WHERE " & ResChamps & " Like '%" & UCase(Trim(Combo)) & "%' ") & "  Total Etudiants"
   End If
   '----------------------------
   If OptRecent.Value = True Then
     CalculeSQL = CalculeSQL & "ORDER BY Date_Enregistrement DESC"
   ElseIf OptASC.Value = True Then
     CalculeSQL = CalculeSQL & "ORDER BY Nom ASC"
   ElseIf OptDESC.Value = True Then
     CalculeSQL = CalculeSQL & "ORDER BY Nom DESC"
   End If
   '----------------------------
End Function

Private Sub MSHFe_DblClick()
  If MSHFe.Row > 0 Then
    Matricule = MSHFe.TextMatrix(MSHFe.RowSel, 3)
    OpenModEtudiant (CStr(Matricule))
  End If
End Sub

Private Sub LPrint_Click()
'Imprimer La Grid
Dim ImpSQL

    If ComboRecItem = "Tous" Then
      ImpSQL = "SELECT Prenom, Nom, Matricule, Statut, Niveau_Etude FROM ETUDIANTS "
    Else
      ImpSQL = "SELECT Prenom, Nom, Matricule, Statut, Niveau_Etude FROM ETUDIANTS WHERE " & ResChamps & " Like '%" & UCase(Trim(Combo)) & "%' "
    End If
    '-----------------------------
   If OptRecent.Value = True Then
     ImpSQL = ImpSQL & "ORDER BY Date_Enregistrement DESC"
   ElseIf OptASC.Value = True Then
     ImpSQL = ImpSQL & "ORDER BY Nom ASC"
   ElseIf OptDESC.Value = True Then
     ImpSQL = ImpSQL & "ORDER BY Nom DESC"
   End If
    '-----------------------------
    Open_DR_Etud CStr(ImpSQL)
 
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

Private Sub MSHFe_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Affiche un PopUp
  If Button = 2 Then
    PopupMenu PopUp
  End If
End Sub
Private Sub PopMod_Click()
  OpenModEtudiant (MSHFe.TextMatrix(MSHFe.RowSel, 3))
  Afficher
End Sub

Private Sub PopSup_Click()
  OpenSupEtudiant (MSHFe.TextMatrix(MSHFe.RowSel, 3))
  Afficher
End Sub

Private Sub mnImpDossier_Click()
  If Not G_FullScreen Then
    Open_DR_Dos_Etud (MSHFe.TextMatrix(MSHFe.RowSel, 3))
  Else
    MsgBox "Vous Devez Quitter FullScreen Pour Pouvoir Imprimer !"
  End If
End Sub

Private Function GetEtudiantDiplomer() As Long
'Nombre Total des Etudiants Diplomés
  GetEtudiantDiplomer = NombreEnreg("ETUDIANTS", "Statut", "Diplomé")
End Function

Private Function GetEtudiantAbandonner() As Long
'Nombre Total des Etudiants Diplomés
  GetEtudiantAbandonner = NombreEnreg("ETUDIANTS", "Statut", "Abandonné")
End Function

Private Function GetRowBackColor()
'Changer La Couleur Des Etudiants Regulier = Blue, Abandonner = Red, Diplomé = Green
Dim lngRow As Long
Dim lngColour As Long
'---------
With MSHFe
   .Redraw = False
    .FillStyle = flexFillRepeat
    '---------------------------
    For lngRow = .FixedRows To .Rows - 1
        Select Case .TextMatrix(lngRow, 5)
            Case "Regulier"
                lngColour = vbBlue
            Case "Diplomé"
                lngColour = vbGreen
            Case Else
                lngColour = vbRed
        End Select
        .Row = lngRow
        .ColSel = .Cols - 1
        .CellForeColor = lngColour
    Next
    '---------------------------
   .FillStyle = flexFillSingle
   .Redraw = True
End With
'-------
End Function

'##################### HALIDOU CISSE ##################

