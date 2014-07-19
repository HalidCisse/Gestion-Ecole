VERSION 5.00
Begin VB.Form FrmPayment 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RECEVOIR UN PAYEMENT"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPayment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6870
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "&H8000000F&"
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
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6615
      Begin VB.ComboBox ComboMotifPay 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   11
         Top             =   1200
         Width           =   3135
      End
      Begin VB.ComboBox ComboMoyPay 
         Appearance      =   0  'Flat
         Height          =   405
         ItemData        =   "FrmPayment.frx":030A
         Left            =   4320
         List            =   "FrmPayment.frx":0317
         TabIndex        =   10
         Top             =   2640
         Width           =   2175
      End
      Begin VB.TextBox TxtPayBy 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   8
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox TxtNumPay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   4
         Top             =   240
         Width           =   3135
      End
      Begin VB.ComboBox ComboNumMat 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   3
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Somme A Payer :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1755
      End
      Begin VB.Label LabSom 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   3360
         TabIndex        =   13
         Top             =   1680
         Width           =   3075
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Motif :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   660
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Moyen de Payement"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   2760
         Width           =   2100
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Etudiant :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "N° Ref Payement :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1860
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Payer Par :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   1140
      End
   End
   Begin VB.CommandButton CmdRaz 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Effacé"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton CmdEnreg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Confirmé"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   1695
   End
End
Attribute VB_Name = "FrmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdEnreg_Click()
   If ChampsPayOK() Then
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM PAYMENTS WHERE N°_Payment = '" & UCase(Trim(TxtNumPay.Text)) & "' "
    rs.Open SQL, CN, adOpenKeyset
'----------------------------------------------------
    If rs.EOF Then
       x = MsgBox("Efféctuer Le Payement ?", vbYesNo + vbQuestion)
       If x = vbYes Then
          EnregPayInfos (UCase(Trim(TxtNumPay.Text)))
          MsgBox "Payement Effectué Avec Succès !!", vbInformation
          '-----------------------------------------------------------
          If GetSetting("Email_Conf") = "Oui" Then
          x = MsgBox("Voulez Vous Envoyez Un Email de Confirmation a " & ComboNumMat, vbQuestion + vbYesNo)
          If x = vbYes Then
            If SendPayEmail(GetMaMat(ComboNumMat), Trim(TxtNumPay.Text)) Then
              MsgBox "Email Envoyer Avec Succès !", vbInformation
            End If
          End If
          End If
          '-----------------------------------------------------------
          Unload Me
       End If
    Else
       MsgBox "Payement Déja Effectuer !", vbInformation
    End If
'-------------------------------------------------------
rs.Close
Set rs = Nothing
End If
End Sub

Private Sub CmdRaz_Click()
  ComboNumMat = ""
  ComboMotifPay = ""
  LabSom.Caption = "0 Dh"
  TxtPayBy.Text = ""
End Sub

Private Sub ComboMotifPay_Validate(Cancel As Boolean)
'Remplit Le Label Avec La Somme A Payer
  LabSom.Caption = GetDette_Som(ComboMotifPay & GetMaMat(ComboNumMat)) & " Dh"
End Sub

Private Sub ComboNumMat_Validate(Cancel As Boolean)
  Me.Caption = "RECEVOIR UN PAYEMENT DE " & UCase(ComboNumMat)
  TxtPayBy.Text = ComboNumMat
  RemplirMesDette GetMaMat(ComboNumMat), ComboMotifPay
  On Error Resume Next
  ComboMotifPay.ListIndex = 0
  LabSom.Caption = GetDette_Som(ComboMotifPay & GetMaMat(ComboNumMat)) & " Dh"
  CmdEnreg.SetFocus
End Sub

Private Sub Form_Activate()
  ComboMoyPay.ListIndex = 0
End Sub

Private Sub Form_Load()
  RemplirEtudiantAvecDette FrmPayment.ComboNumMat
  TxtNumPay = GenNewPayID
  '---------------------
  If G_FullScreen = True Then
     Me_Top Me
  End If
  '---------------------
End Sub

'##################### HALIDOU CISSE ##################
