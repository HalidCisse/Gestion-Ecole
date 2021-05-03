Attribute VB_Name = "Validate"
Option Explicit


Function isString(Mess As String) As Boolean
isString = True
Dim i
Mess = Trim(Mess)

If Len(Mess) < 3 Then
  isString = False
  Exit Function
Else
  For i = 1 To Len(Mess)
      '--------------------------
      If IsNumeric(Mid(Mess, i, 1)) Then
        isString = False
        Exit Function
      End If
      '--------------------------
  Next
End If
End Function

Function ValidEmail(ByVal Address_Email As String) As Boolean
'Verifie si l'email est valide

Dim strCheck As String
Dim bCK As Boolean
Dim strDomainType As String
Dim strDomainName As String
Const sInvalidChars As String = "!#$%^&*()=+{}[]|\;:'/?>,< "
Dim i As Integer
strCheck = Address_Email

bCK = Not InStr(1, strCheck, Chr(34)) > 0 'verifie s'il ya un double quote ""
If Not bCK Then GoTo ExitFunction

bCK = Not InStr(1, strCheck, "..") > 0 'verifie qu'il nya pas double point
If Not bCK Then GoTo ExitFunction

' verifie qu'il nya pas de caractere invalide
If Len(strCheck) > Len(sInvalidChars) Then
    For i = 1 To Len(sInvalidChars)
        If InStr(strCheck, Mid(sInvalidChars, i, 1)) > 0 Then
            bCK = False
            GoTo ExitFunction
        End If
    Next
Else
    For i = 1 To Len(strCheck)
        If InStr(sInvalidChars, Mid(strCheck, i, 1)) > 0 Then
            bCK = False
            GoTo ExitFunction
        End If
    Next
End If

If InStr(1, strCheck, "@") > 1 Then 'verifi qu'il ya un @
    bCK = Len(Left(strCheck, InStr(1, strCheck, "@") - 1)) > 0
Else
    bCK = False
End If
If Not bCK Then GoTo ExitFunction

strCheck = Right(strCheck, Len(strCheck) - InStr(1, strCheck, "@"))
bCK = Not InStr(1, strCheck, "@") > 0 'verifi qu il nya pas plus de deux @
If Not bCK Then GoTo ExitFunction

strDomainType = Right(strCheck, Len(strCheck) - InStr(1, strCheck, "."))
bCK = Len(strDomainType) > 0 And InStr(1, strCheck, ".") < Len(strCheck)
If Not bCK Then GoTo ExitFunction

strCheck = Left(strCheck, Len(strCheck) - Len(strDomainType) - 1)
Do Until InStr(1, strCheck, ".") <= 1
    If Len(strCheck) >= InStr(1, strCheck, ".") Then
        strCheck = Left(strCheck, Len(strCheck) - (InStr(1, strCheck, ".") - 1))
    Else
        bCK = False
        GoTo ExitFunction
    End If
Loop
If strCheck = "." Or Len(strCheck) = 0 Then bCK = False

ExitFunction:
ValidEmail = bCK
End Function


'####################### Halidou Cisse ###########################
