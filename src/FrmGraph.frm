VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form FrmGraph 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STATISTIQUES"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15885
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmGraph.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   15885
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Double Click Pour Changer De Type"
      Top             =   0
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   15690
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "INSCRIPTIONS"
      TabPicture(0)   =   "FrmGraph.frx":014A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Chart_INS"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ETUDIANTS"
      TabPicture(1)   =   "FrmGraph.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Chart_ETUD"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "PAYEMENTS"
      TabPicture(2)   =   "FrmGraph.frx":0182
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Chart_PAY"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin MSChart20Lib.MSChart Chart_PAY 
         Height          =   8415
         Left            =   0
         OleObjectBlob   =   "FrmGraph.frx":019E
         TabIndex        =   2
         ToolTipText     =   "Double Click Pour Changer De Type"
         Top             =   360
         Width           =   15855
      End
      Begin MSChart20Lib.MSChart Chart_ETUD 
         Height          =   8415
         Left            =   -75000
         OleObjectBlob   =   "FrmGraph.frx":1C1C
         TabIndex        =   1
         ToolTipText     =   "Double Click Pour Changer De Type"
         Top             =   360
         Width           =   15855
      End
      Begin MSChart20Lib.MSChart Chart_INS 
         Height          =   8415
         Left            =   -75000
         OleObjectBlob   =   "FrmGraph.frx":36BF
         TabIndex        =   3
         ToolTipText     =   "Double Click Pour Changer De Type"
         Top             =   360
         Width           =   15855
      End
   End
End
Attribute VB_Name = "FrmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim chart_Ins_Data() As String
Dim chart_Etud_Data(1 To 5, 1 To 2) As String
Dim chart_Pay_Data(1 To 24, 1 To 2) As String
Private i As Integer


Private Sub Form_Load()
  If G_FullScreen = True Then
     Me_Top Me
  End If
End Sub

Private Sub Form_Activate()
    GetInsData
    GetEtudData
    GetPayData
End Sub

Private Function GetInsData()
'Remplit La Graph Inscription
'--------------------
Dim N As Long
Dim C
N = GetNbreClasse
'----------------------
If N = 0 Then
  Chart_INS.Visible = False
Exit Function
End If
'----------------------
ReDim chart_Ins_Data(1 To N, 1 To 2)
C = 0
  '--------------------
  Dim rs As New ADODB.Recordset
  SQL = "SELECT Classe FROM CLASSES"
  rs.Open SQL, CN
  '--------------------
  Do While Not rs.EOF
    If Not GetClasseEffectif(rs![Classe]) = 0 Then
        chart_Ins_Data(N - C, 1) = CStr(rs![Classe])
        chart_Ins_Data(N - C, 2) = NombreEnreg("INSCRIPTIONS", "SQL", "SELECT COUNT(*) AS Nombre FROM INSCRIPTIONS WHERE Statut = 'Active' AND Classe = '" & rs![Classe] & "'")
        C = C + 1
    End If
  rs.MoveNext
  Loop
  '-------------------
  Chart_INS.ChartData = chart_Ins_Data
  Chart_INS.Footnote = "Nombre d'Etudiants Inscrits Par Classe"
  '-------------------
  rs.Close
  Set rs = Nothing
End Function


Private Function GetEtudData()
'Remplit La Graph Etudiant
Dim y
  y = Year(Date) - 5
  '---------------------
  For i = 5 To 1 Step -1
    chart_Etud_Data(i, 1) = CStr(y + i) & " "
    chart_Etud_Data(i, 2) = NombreEnreg("ETUDIANTS", "SQL", "SELECT COUNT(*) AS Nombre FROM ETUDIANTS WHERE Year(Date_Enregistrement) = '" & y + i & "'")
  Next i
  '---------------------
  Chart_ETUD.ChartData = chart_Etud_Data
  Chart_ETUD.Footnote = "Nombre d'Etudiants Enrégistrés Par Année"
  '---------------------
End Function

Private Function GetPayData()
'Remplit La Graph Inscription
Dim M, S, y, Dat
  Dat = DateSerial(Year(Date) - 2, Month(Date) + 1, 1)
  '---------------------
  For i = 1 To 24
    chart_Pay_Data(i, 1) = Format(Dat, "mmm") & "-" & Year(Dat)
    chart_Pay_Data(i, 2) = NombreEnreg("PAYMENTS", "SQL", "SELECT SUM(Somme_payer) AS Nombre FROM PAYMENTS WHERE Year(Date_Payment) = '" & Year(Dat) & "' AND  Month(Date_Payment) = '" & Month(Dat) & "'")
  Dat = DateSerial(Year(DateAdd("m", 1, Dat)), Month(DateAdd("m", 1, Dat)), Day(Dat))
  Next i
  '---------------------
  Chart_PAY.ChartData = chart_Pay_Data
  Chart_PAY.Footnote = "Total Payements Effectué Par Mois"
  '---------------------
End Function

Private Sub Chart_ETUD_DblClick()
  If Chart_ETUD.chartType = VtChChartType2dBar Then
    Chart_ETUD.chartType = VtChChartType2dLine
  ElseIf Chart_ETUD.chartType = VtChChartType2dLine Then
    Chart_ETUD.chartType = VtChChartType2dPie
  ElseIf Chart_ETUD.chartType = VtChChartType2dPie Then
    Chart_ETUD.chartType = VtChChartType3dLine
  ElseIf Chart_ETUD.chartType = VtChChartType3dLine Then
    Chart_ETUD.chartType = VtChChartType3dBar
  Else
    Chart_ETUD.chartType = VtChChartType2dBar
  End If
End Sub

Private Sub Chart_INS_DblClick()
  If Chart_INS.chartType = VtChChartType2dBar Then
    Chart_INS.chartType = VtChChartType2dLine
  ElseIf Chart_INS.chartType = VtChChartType2dLine Then
    Chart_INS.chartType = VtChChartType2dPie
  ElseIf Chart_INS.chartType = VtChChartType2dPie Then
    Chart_INS.chartType = VtChChartType3dLine
  ElseIf Chart_INS.chartType = VtChChartType3dLine Then
    Chart_INS.chartType = VtChChartType3dBar
  Else
    Chart_INS.chartType = VtChChartType2dBar
  End If
End Sub

Private Sub Chart_PAY_DblClick()
  If Chart_PAY.chartType = VtChChartType2dBar Then
    Chart_PAY.chartType = VtChChartType2dLine
  ElseIf Chart_PAY.chartType = VtChChartType2dLine Then
    Chart_PAY.chartType = VtChChartType2dPie
  ElseIf Chart_PAY.chartType = VtChChartType2dPie Then
    Chart_PAY.chartType = VtChChartType3dLine
  ElseIf Chart_PAY.chartType = VtChChartType3dLine Then
    Chart_PAY.chartType = VtChChartType3dBar
  Else
    Chart_PAY.chartType = VtChChartType2dBar
  End If
End Sub

'##################### HALIDOU CISSE ##################
