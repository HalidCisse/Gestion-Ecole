VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmShowEvents 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Evènements"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9255
   Icon            =   "FrmShowEvents.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   9255
   StartUpPosition =   1  'CenterOwner
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid EventsGrid 
      Height          =   5880
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   10372
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
      GridLinesFixed  =   0
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
      _Band(0)._NumMapCols=   4
      _Band(0)._MapCol(0)._Name=   "Message"
      _Band(0)._MapCol(0)._Caption=   "Evènement"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "UserEmail"
      _Band(0)._MapCol(1)._Caption=   "Utilisateur"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "EventTime"
      _Band(0)._MapCol(2)._Caption=   "Heure"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "EventDate"
      _Band(0)._MapCol(3)._Caption=   "Date"
      _Band(0)._MapCol(3)._RSIndex=   3
   End
End
Attribute VB_Name = "FrmShowEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   ScaleGrid
   Me.Width = EventsGrid.Width + 250
   Me.Height = EventsGrid.Height + 550
End Sub

Private Function ScaleGrid()
  EventsGrid.ColWidth(0) = 4000
  EventsGrid.ColWidth(1) = 2000
  EventsGrid.ColWidth(2) = 1000
  EventsGrid.ColWidth(3) = 1300

  EventsGrid.Width = 8550
End Function
