Attribute VB_Name = "TopMost"

Option Explicit
Global G_FullScreen As Boolean

'#########################################################
Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long
'#########################################################
'ALWAYS ON TOP
      Public Const SWP_NOMOVE = 2
      Public Const SWP_NOSIZE = 1
      Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
      Public Const HWND_TOPMOST = -1
      Public Const HWND_NOTOPMOST = -2
'#########################################################
'FULL SCREEN OVER TASKBAR
Declare Function GetSystemMetrics Lib "user32" _
         (ByVal nIndex As Long) As Long
      Public Const SM_CXSCREEN = 0
      Public Const SM_CYSCREEN = 1
      Public Const HWND_TOP = 0
      Public Const SWP_SHOWWINDOW = &H40
'#########################################################
'HIDE TASKBAR
Declare Function FindWindowA Lib "user32" _
   (ByVal lpClassName As String, _
   ByVal lpWindowName As String) As Long
  Const TOGGLE_HIDEWINDOW = &H80
  Const TOGGLE_UNHIDEWINDOW = &H40
Public handleW1 As Long
'#########################################################

Function SetTopMostWindow(hwnd As Long, TopMost As Boolean) As Long
'ALWAYS ON TOP
    If TopMost = True Then
       SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
       SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
       SetTopMostWindow = False
    End If
End Function
'##########################################################




Function HideTaskbar(Optional Vrai As Boolean = True)
'Faire Disparaitre La Task Bar
   If Vrai = True Then
     handleW1 = FindWindowA("Shell_traywnd", "")
     SetWindowPos handleW1, 0, 0, 0, 0, 0, TOGGLE_HIDEWINDOW
   Else
     SetWindowPos handleW1, 0, 0, 0, 0, 0, TOGGLE_UNHIDEWINDOW
   End If
End Function

Function Me_FullScreen(Form As Form, Optional Vrai As Boolean = True)
'Couvre Tout L'Ecran Meme Taskbar
  Dim cx As Long
  Dim cy As Long
  Dim RetVal As Long
  If Vrai Then
    If Form.WindowState = vbMaximized Then
       Form.WindowState = vbNormal
    End If
    cx = GetSystemMetrics(SM_CXSCREEN)
    cy = GetSystemMetrics(SM_CYSCREEN)
    RetVal = SetWindowPos(Form.hwnd, HWND_TOP, 0, 0, cx, cy, SWP_SHOWWINDOW)
  Else
    Form.Hide
  End If
End Function

Function Me_Top(Form As Form, Optional Vrai As Boolean = True)
'Top Most
  Dim lR As Long
  If Vrai = True Then
    lR = SetTopMostWindow(Form.hwnd, True)
  Else
    lR = SetTopMostWindow(Form.hwnd, False)
  End If
End Function

Function HideOutil(Optional Vrai As Boolean = True)
'Cache Les Outils
With Main
  If Vrai Then
    .mnCalc.Enabled = False
    .mnNotePad.Enabled = False
    
    .mnFacebook.Enabled = False
    .mnGoogle.Enabled = False
    .mnGmail.Enabled = False
    
    .mnAccess.Enabled = False
    .mnWord.Enabled = False
    .mnExcel.Enabled = False
    .mnPPT.Enabled = False
    
    .mnImpIns.Enabled = False
    .mnImpPays.Enabled = False
    .mnImprimEt.Enabled = False
    
    .mnAide2.Enabled = False
    .mnLicence.Enabled = False
    .mnFeedB.Enabled = False
    .mnApropos.Enabled = False
    
  Else
    .mnCalc.Enabled = True
    .mnNotePad.Enabled = True
    
    .mnFacebook.Enabled = True
    .mnGoogle.Enabled = True
    .mnGmail.Enabled = True
    
    .mnAccess.Enabled = True
    .mnWord.Enabled = True
    .mnExcel.Enabled = True
    .mnPPT.Enabled = True
    
    .mnImpIns.Enabled = True
    .mnImpPays.Enabled = True
    .mnImprimEt.Enabled = True
    
    .mnAide2.Enabled = True
    .mnLicence.Enabled = True
    .mnFeedB.Enabled = True
    .mnApropos.Enabled = True
    
  End If
End With
End Function

'##################### END OF STORY   ##################























