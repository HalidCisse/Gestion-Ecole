Attribute VB_Name = "RunAtStartUp"

'Lancer Au Demarage De Windows

Option Explicit

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal _
  hKey As Long) As Long

Declare Function RegSetValueEx Lib "advapi32.dll" _
 Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName _
 As String, ByVal Reserved As Long, ByVal dwType As Long, _
 lpData As Any, ByVal cbData As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName _
As String, ByVal lpReserved As Long, lpType As Long, _
lpData As Any, lpcbData As Long) As Long

Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
   "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
   As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
   As Long, phkResult As Long, lpdwDisposition As Long) As Long
   
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
   "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
   Long) As Long

Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

'------------------------------

Public Const HKEY_CURRENT_USER = &H80000001
Public Const KEY_WRITE = &H20006
Public Const REG_SZ = 1
Public Const ERROR_SUCCESS = 0


'################################################################

Function SetRunAtStartup(ByVal app_name As String, ByVal app_path As String, Optional ByVal Run_At_StartUp As Boolean = True)
'Modifier Le Registre Pour Que Le Programme Soit Lancer Au Demarage
'-------------------
Dim hKey As Long
Dim key_value As String
Dim status As Long
'-------------------
    On Error GoTo SetStartupError
    '---------------------------
    If RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", ByVal 0&, ByVal 0&, ByVal 0&, KEY_WRITE, ByVal 0&, hKey, ByVal 0&) <> ERROR_SUCCESS Then
        MsgBox "Error " & err.Number & " opening key" & vbCrLf & err.Description
        Exit Function
    End If
    '---------------------------
    If Run_At_StartUp Then
        key_value = app_path & "\" & app_name & ".exe" & vbNullChar
        status = RegSetValueEx(hKey, App.EXEName, 0, REG_SZ, ByVal key_value, Len(key_value))
        If status <> ERROR_SUCCESS Then
            MsgBox "Error " & err.Number & " setting key" & vbCrLf & err.Description
        End If
    Else
        RegDeleteValue hKey, app_name
    End If
    '---------------------------
    RegCloseKey hKey
    Exit Function

SetStartupError:
    MsgBox err.Number & " " & err.Description
    Exit Function
End Function

Function WillRunAtStartup(ByVal app_name As String) As Boolean
'Verifie Si L'Application Demarre Avec Wiindows
Dim hKey As Long
Dim value_type As Long
    If RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", 0, KEY_READ, hKey) = ERROR_SUCCESS Then
        WillRunAtStartup = (RegQueryValueEx(hKey, app_name, ByVal 0&, value_type, ByVal 0&, ByVal 0&) = ERROR_SUCCESS)
        RegCloseKey hKey
    Else
        WillRunAtStartup = False
    End If
End Function

'##################### HALIDOU CISSE ##################

