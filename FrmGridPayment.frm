VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmGridPayment 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LISTE DES PAYEMENTS"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmGridPayment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   15360
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdAfficher 
      Appearance      =   0  'Flat
      Caption         =   "Afficher >>"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox ComboRecItem 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Selectionez Un Champs"
      Top             =   240
      Width           =   2415
   End
   Begin VB.ComboBox Combo 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   7920
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Choisir.."
      ToolTipText     =   "Choisir"
      Top             =   240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridPay 
      Bindings        =   "FrmGridPayment.frx":27A2
      Height          =   4575
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   15100
      _ExtentX        =   26644
      _ExtentY        =   8070
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   9
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
      _Band(0).Cols   =   9
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   10
      _Band(0)._MapCol(0)._Name=   "ID_Dette"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(0)._Hidden=   -1  'True
      _Band(0)._MapCol(1)._Name=   "N°_Matricule"
      _Band(0)._MapCol(1)._RSIndex=   2
      _Band(0)._MapCol(2)._Name=   "N°_Payment"
      _Band(0)._MapCol(2)._RSIndex=   1
      _Band(0)._MapCol(3)._Name=   "Designation"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "Somme_Payer"
      _Band(0)._MapCol(4)._RSIndex=   6
      _Band(0)._MapCol(5)._Name=   "Payer_Par"
      _Band(0)._MapCol(5)._RSIndex=   4
      _Band(0)._MapCol(6)._Name=   "Moyen_De_Payment"
      _Band(0)._MapCol(6)._RSIndex=   5
      _Band(0)._MapCol(7)._Name=   "Email_Profile"
      _Band(0)._MapCol(7)._RSIndex=   7
      _Band(0)._MapCol(8)._Name=   "Date_Payment"
      _Band(0)._MapCol(8)._RSIndex=   8
      _Band(0)._MapCol(9)._Name=   "Time_Payment"
      _Band(0)._MapCol(9)._RSIndex=   9
   End
   Begin VB.Image LPrint 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   14640
      Picture         =   "FrmGridPayment.frx":27B7
      Stretch         =   -1  'True
      ToolTipText     =   "Imprimé"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label LabelManque 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Total Manque a Gagner"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   10320
      TabIndex        =   8
      ToolTipText     =   "Total des Dettes Non Payer"
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Etudiant Avec defaut de Payement"
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
      Top             =   5400
      Width           =   2700
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   120
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label LabelSomme 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Payer"
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
      Left            =   6600
      TabIndex        =   6
      ToolTipText     =   "Total Payer"
      Top             =   5400
      Width           =   885
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
      Left            =   14760
      TabIndex        =   5
      ToolTipText     =   "Nombres de Payments"
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
      Left            =   6330
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Menu PopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu PopSupPay 
         Caption         =   "&Supprimé"
      End
      Begin VB.Menu mnImpRecu 
         Caption         =   "&Imprimé Reçue"
      End
   End
End
Attribute VB_Name = "FrmGridPayment"
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
  Call Remplir(Combo, "PAYMENTS", ResChamps)
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
  If G_FullScreen = True Then
     Me_Top Me
     LPrint.Visible = False
  End If
End Sub

Private Function Afficher()
'Affiche Les Payments Selon Les Criteres
  RemplirGrid FrmGridPayment, FrmGridPayment.GridPay, CalculeSQL
  GetRowBackColor
  '----------------------
  Dim MG As Long
  MG = GetManqueAGagner
  If MG > 0 Then
    LabelManque.ForeColor = vbRed
    LabelManque.Caption = "Defaut de Payement: " & Format(MG, "### ### ### ###") & " Dh"
  ElseIf MG = 0 Then
    LabelManque.ForeColor = vbBlack
    LabelManque.Caption = "En Regle"
  Else
    MG = -MG
    LabelManque.ForeColor = vbBlue
    LabelManque.Caption = "En Avance: " & Format(MG, "### ### ### ###") & " Dh"
  End If
  '----------------------
  GetEtudiantsName
End Function

Private Function RemplirItems()
'Remplir les combos des données necessaires
With ComboRecItem
.Clear
     .AddItem ("Tous")
     .AddItem ("Nom Etudiant")
     .AddItem ("N° de Payment")
     .AddItem ("Designation")
     .AddItem ("Moyen de Payment")
     .AddItem ("Somme Payer")
     .AddItem ("Profile d'utilisateur")
     .AddItem ("Date de Payment")
.Refresh
.ListIndex = 0
End With
End Function

Private Function CalculeSQL() As String
'Calculer le SQL coresspondant aux criteres
    If ComboRecItem = "Tous" Then
       ResChamps = "N°_Payment"
       LabelSign.Visible = False
       Combo.Visible = False
       CalculeSQL = "SELECT * FROM PAYMENTS ORDER BY Date_Payment DESC , Time_Payment DESC"
       LabelResult = NombreEnreg("PAYMENTS", "SQL", "SELECT COUNT(*) as Nombre FROM PAYMENTS") & " Payements"
       LabelSomme = "Total Payer: " & Format(NombreEnreg("PAYMENTS", "SQL", "SELECT SUM(Somme_Payer) as Nombre FROM PAYMENTS"), "### ### ### ###") & " Dh"
    Else
       LabelSign.Visible = True
       Combo.Visible = True
    '---------------------------------------------------
     If ComboRecItem = "N° de Payment" Then
       ResChamps = "N°_Payment"
     ElseIf ComboRecItem = "N° de Matricule" Then
       ResChamps = "N°_Matricule"
     ElseIf ComboRecItem = "Designation" Then
       ResChamps = "Designation"
     ElseIf ComboRecItem = "Nom Etudiant" Then
       ResChamps = "Payer_Par"
     ElseIf ComboRecItem = "Moyen de Payment" Then
       ResChamps = "Moyen_De_Payment"
     ElseIf ComboRecItem = "Somme Payer" Then
       ResChamps = "Somme_Payer"
     ElseIf ComboRecItem = "Profile d'utilisateur" Then
       ResChamps = "Email_Profile"
     ElseIf ComboRecItem = "Date de Payment" Then
       ResChamps = "Date_Payment"
     End If
    '------------------------------------------------------
     CalculeSQL = "SELECT * FROM PAYMENTS WHERE " & ResChamps & " Like '%" & UCase(Trim(Combo)) & "%' ORDER BY Date_Payment DESC , Time_Payment DESC "
     LabelResult = NombreEnreg("PAYMENTS", "SQL", "SELECT COUNT(*) as Nombre FROM PAYMENTS WHERE " & ResChamps & " Like '%" & UCase(Trim(Combo)) & "%' ") & " Payments"
     LabelSomme = "Total Payer: " & Format(NombreEnreg("PAYMENTS", "SQL", "SELECT SUM(Somme_Payer) as Nombre FROM PAYMENTS WHERE " & ResChamps & " Like '%" & UCase(Trim(Combo)) & "%' "), "### ### ### ###") & " Dh"
   End If
End Function

Private Sub GridPay_DblClick()
  ComboRecItem.ListIndex = 1
  Combo = GridPay.TextMatrix(GridPay.RowSel, 0)
  Afficher
End Sub

Private Sub LPrint_Click()
  'Imprimer La Grid
 If ComboRecItem = "Tous" Then
   Open_DR_Pay "SELECT Payer_Par, N°_Payment, Designation, Somme_Payer, Date_Payment FROM PAYMENTS ORDER BY Date_Payment DESC"
 Else
   Open_DR_Pay "SELECT Payer_Par, N°_Payment, Designation, Somme_Payer, Date_Payment FROM PAYMENTS WHERE " & ResChamps & " Like '%" & UCase(Trim(Combo)) & "%' ORDER BY Date_Payment DESC"
 End If
End Sub

Private Sub Timer_Timer()
  CalculeSQL
End Sub

Private Function GetEtudiantsName()
'Permet De Trouver Le Nom De L'Etudiant Correspondant Aux Matricules
Dim i
GridPay.TextMatrix(0, 0) = "Nom de L'Etudiant"
'-------------------
  For i = 1 To Val(LabelResult.Caption)
    GridPay.TextMatrix(i, 0) = GetMyName(GridPay.TextMatrix(i, 0))
  Next i
'------------------
End Function

Private Function ScaleGrid()
With GridPay
   '-------------------------
  .ColWidth(0) = 3000  'mat
  .ColWidth(1) = 1500  'N° pay
  .ColWidth(2) = 2000  'Des
  .ColWidth(3) = 1300  'som pay
  .ColWidth(4) = 1500  'py par
  .ColWidth(5) = 1000  'moyen pay
  .ColWidth(6) = 2000  'user
  .ColWidth(7) = 1300  'date
  .ColWidth(8) = 1300  'time
  '--------------------------
  '.Width = 8550
End With
End Function

Private Sub GridPay_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  'Affiche un PopUp
  If Button = 2 Then
    PopupMenu PopUp
  End If
End Sub

Private Sub PopSupPay_Click()
  OpenSupPayment (GridPay.TextMatrix(GridPay.RowSel, 1))
  Afficher
End Sub

Private Sub mnImpRecu_Click()
  If Not G_FullScreen Then
    Open_DR_Recu_Pay (GridPay.TextMatrix(GridPay.RowSel, 1))
  Else
    MsgBox "Vous Devez Quitter FullScreen Pour Pouvoir Imprimer !"
  End If
End Sub

Private Function GetRowBackColor()
'Changer La Couleur Des Payments Avec Dette = red, sans Dette = blue
'---------------------
Dim lngRow As Long
Dim lngColour As Long
Dim DT As Long
'---------------------
With GridPay
   .Redraw = False
    .FillStyle = flexFillRepeat
    '----------------------------
    For lngRow = .FixedRows To .Rows - 1
          DT = GetMaDette(.TextMatrix(lngRow, 0))
          If DT > 0 Then
            lngColour = vbRed
          ElseIf DT < 0 Then
            lngColour = vbBlue
          Else
            lngColour = vbBlack
          End If
        .Row = lngRow
        .ColSel = .Cols - 1
        .CellForeColor = lngColour
    Next
    '----------------------------
    .FillStyle = flexFillSingle
   .Redraw = True
End With
'--------------------------------
End Function

Private Function GetManqueAGagner() As Long
'Calculer le Manque a Gagner
 Dim SQL2
 Dim rs As New ADODB.Recordset
     '----------------------------
     If ComboRecItem = "Tous" Then
       SQL2 = "SELECT DISTINCT N°_Matricule FROM PAYMENTS"
     Else
       SQL2 = "SELECT DISTINCT N°_Matricule FROM PAYMENTS WHERE " & ResChamps & " Like '%" & UCase(Trim(Combo)) & "%' "
     End If
     '---------------------------
    rs.Open SQL2, CN
    '-------------------
    Do While Not rs.EOF
      GetManqueAGagner = GetManqueAGagner + CLng(GetMaDette(rs![N°_Matricule]))
    rs.MoveNext
    Loop
    '------------------
 rs.Close
 Set rs = Nothing
End Function

'##################### HALIDOU CISSE ##################





