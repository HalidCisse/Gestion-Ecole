VERSION 5.00
Begin VB.Form FrmBkg 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBkg.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmBkg.frx":0A8A
   ScaleHeight     =   3075
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmBkg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
 
Private WithEvents apiLink As EventVB.APIFunctions
Attribute apiLink.VB_VarHelpID = -1
Private WithEvents hook As EventVB.ApiSystemHook
Attribute hook.VB_VarHelpID = -1
 
Private Const vbKeyLWin As Integer = 91
Private Const vbKeyRWin As Integer = 92
Private Const vbKeyLCtrl As Integer = 162
Private Const vbKeyRCtrl As Integer = 163


Private Sub Form_Load()
On Error Resume Next
  G_FullScreen = True
  If Not Ctrl_Alt(False) Then
    MsgBox "Impossible de Bloquer les Touches !"
  End If
  HideOutil
  HideTaskbar
  Me_Top FrmBkg
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  G_FullScreen = False
  HideTaskbar False
  HideOutil False
  Ctrl_Alt True
  
End Sub

Function Ctrl_Alt(Optional Vrai As Boolean = True) As Boolean
'Sert A Bloquer Les Bouttons Windows, Ctrl + Alt ....
On Error GoTo Error
Ctrl_Alt = True
deb:
  If Not Vrai Then
    Set apiLink = New EventVB.APIFunctions
    Set hook = apiLink.System.Hooks
    On Error GoTo deb
    hook.StartHook WH_KEYBOARD_LL, HOOK_GLOBAL
  Else
    hook.StopHook WH_KEYBOARD_LL
    Set hook = Nothing
    Set apiLink = Nothing
  End If
  
  Exit Function
Error:
  Ctrl_Alt = False
End Function

Private Sub hook_KeyDown(ByVal VKey As Long, ByVal scanCode _
    As Long, ByVal ExtendedKey As Boolean, ByVal AltDown As _
    Boolean, ByVal Injected As Boolean, Cancel As Boolean)
 On Error Resume Next
    'Alt Tab
    If AltDown And VKey = vbKeyTab Then
        Cancel = True
   
    'Alt Esc
    ElseIf AltDown And VKey = vbKeyEscape Then
        Cancel = True
    
    'Ctrl Esc
    ElseIf VKey = vbKeyEscape Then
        'ctrl key down
        If GetKeyState(vbKeyLCtrl) And &HF0000000 Or _
            GetKeyState(vbKeyRCtrl) And &HF0000000 Then
            Cancel = True
        End If
        
    'Windows key (L/R)
    ElseIf VKey = vbKeyLWin Or VKey = vbKeyRWin Then
        Cancel = True
    
    'Windows + Any
    Else
        If GetKeyState(vbKeyLWin) And &HF0000000 Or _
            GetKeyState(vbKeyRWin) And &HF0000000 Then
            Cancel = True
        End If
    End If
    
    If Cancel = True Then
        MessageBeep 0
    End If
    
End Sub

'##################### HALIDOU CISSE ##################
