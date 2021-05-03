VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmShowUser 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "List des Utilisateurs"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8100
   Icon            =   "FrmShowUser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   8100
   StartUpPosition =   1  'CenterOwner
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid UserGrid 
      Bindings        =   "FrmShowUser.frx":014A
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Double Click Pour Plus d'Informations"
      Top             =   0
      Width           =   18615
      _ExtentX        =   32835
      _ExtentY        =   9551
      _Version        =   393216
      Rows            =   3
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   33023
      BackColorSel    =   14737632
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483643
      GridColorFixed  =   0
      GridColorUnpopulated=   16777215
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      BandDisplay     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Reference Sans Serif"
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
      _Band(0).ColHeader=   1
      _Band(0)._NumMapCols=   5
      _Band(0)._MapCol(0)._Name=   "Prenom"
      _Band(0)._MapCol(0)._RSIndex=   4
      _Band(0)._MapCol(1)._Name=   "Nom"
      _Band(0)._MapCol(1)._RSIndex=   3
      _Band(0)._MapCol(2)._Name=   "UserType"
      _Band(0)._MapCol(2)._Caption=   "Type d'utilisateur"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "Email"
      _Band(0)._MapCol(3)._Caption=   "Adresse Email"
      _Band(0)._MapCol(3)._RSIndex=   0
      _Band(0)._MapCol(4)._Name=   "Pass"
      _Band(0)._MapCol(4)._RSIndex=   1
      _Band(0)._MapCol(4)._Hidden=   -1  'True
   End
End
Attribute VB_Name = "FrmShowUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
   'UserGrid.RowSel = 1
End Sub

Private Sub Form_Load()
   Call ScaleGrid
End Sub

Private Sub UserGrid_DblClick()
 If UserGrid.Rows <> 0 Then
   UserID = UserGrid.TextMatrix(UserGrid.RowSel, 3)
   FillFormUser (UserID)
   FrmUser.TextEmail.Enabled = False
   FrmUser.Show vbModal
 Else
    MsgBox "Vide !!", vbInformation + vbOKOnly, "Gestion Location de Voitures"
 End If
End Sub

Private Function ScaleGrid()
   UserGrid.ColWidth(0) = 1500
   UserGrid.ColWidth(1) = 1500
   UserGrid.ColWidth(2) = 1500
   UserGrid.ColWidth(3) = 3000
   FrmShowUser.Width = 7500
End Function
