Attribute VB_Name = "SendEmail"
Option Explicit
Declare Function ShellExecute _
            Lib "shell32.dll" _
            Alias "ShellExecuteA" ( _
            ByVal hwnd As Long, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) _
            As Long

Public Function BrowseTo(ByRef pstrURL As String)
' Ouvre Le Navigateur Par Defaut
    Call ShellExecute(Main.hwnd, "Open", pstrURL, "", "", True)
End Function

Function SendPayEmail(Matricule As String, Payment_ID As String) As Boolean
'Envoi Un Message de Confirmation D'Inscription
  
  SendPayEmail = False
  Dim Destinateur
  Dim Sujet
  Dim strHTML
  
  Destinateur = GetEtudiantEmail((CStr(Matricule)))
  Sujet = GetSetting("Ecole_nom") & ": Confirmation de reception de payment "
  
  '-----------------------------------------------------------
  strHTML = "<HTML>"
  strHTML = strHTML & "<HEAD>"
  strHTML = strHTML & "<BODY>"
  strHTML = strHTML & "<center>"
  strHTML = strHTML & "<b> Confirmation de Reception de Payment </b></br>"
  strHTML = strHTML & "</center>"
  
  strHTML = strHTML & "<h3>Informations Etudiant</h3>"
  
  strHTML = strHTML & "<ul>"
  strHTML = strHTML & "<li>Nom : " & GetMyName(CStr(Matricule)) & " </li>"
  strHTML = strHTML & "<li>Numéro de Matricule : " & Matricule & " </li>"
  strHTML = strHTML & "<li>Numéro d'Inscription : " & GetMonIns(CStr(Matricule)) & " </li>"
  strHTML = strHTML & "<li>Classe : " & GetMaClasse(GetMonIns(CStr(Matricule))) & "</li>"
  strHTML = strHTML & "<li>Niveau : " & GetClasseNiveau(GetMaClasse((GetMonIns(CStr(Matricule))))) & "</li>"
  strHTML = strHTML & "<li>Année Scolaire : " & GetMonAnneeScolaire(GetMonIns(CStr(Matricule))) & " </li>"
  strHTML = strHTML & "</ul>"
  
  strHTML = strHTML & "<h3>Informations Payment</h3>"
  
  strHTML = strHTML & "<ul>"
  strHTML = strHTML & "<li>N° de Payment : " & Payment_ID & "</li>"
  strHTML = strHTML & "<li>Designation : " & GetPayDes(CStr(Payment_ID)) & "</li>"
  strHTML = strHTML & "<li>Somme Reçue : " & GetPaySum(CStr(Payment_ID)) & " Dh</li>"
  strHTML = strHTML & "<li>Date Reception : " & GetPayDate(CStr(Payment_ID)) & "</li>"
  strHTML = strHTML & "</ul>"
  
  strHTML = strHTML & "<footer>"
  strHTML = strHTML & " <h6> Copyright <i>" & GetSetting("Ecole_nom") & "</i> </h6>"
  strHTML = strHTML & "</footer>"
  strHTML = strHTML & "</BODY>"
  strHTML = strHTML & "</HTML>"
  '-----------------------------------------------------------
  
  If EnvoiEMail(CStr(Destinateur), CStr(Sujet), , CStr(strHTML)) Then
   SendPayEmail = True
  End If
End Function

Function SendInsMessage(Inscription As String) As Boolean
'Envoi Un Message de Confirmation D'Inscription
  
  SendInsMessage = False
  Dim Destinateur
  Dim Sujet
  Dim strHTML
  
  Destinateur = GetEtudiantEmail(GetMaMatricule(CStr(Inscription)))
  Sujet = GetSetting("Ecole_nom") & " : Confirmation d'Inscription"
  
  '-----------------------------------------------------------
  strHTML = "<HTML>"
  strHTML = strHTML & "<HEAD>"
  strHTML = strHTML & "<BODY>"
  strHTML = strHTML & "<center>"
  strHTML = strHTML & "<b> Confirmation d'inscription a Miage Rabat </b></br>"
  strHTML = strHTML & "</center>"
  strHTML = strHTML & "<h3>Informations</h3>"
  strHTML = strHTML & "<ul>"
  strHTML = strHTML & "<li>Nom : " & GetMyName(GetMaMatricule(CStr(Inscription))) & " </li>"
  strHTML = strHTML & "<li>Numéro de Matricule : " & GetMaMatricule(CStr(Inscription)) & " </li>"
  strHTML = strHTML & "<li>Numéro d'Inscription : " & Inscription & " </li>"
  strHTML = strHTML & "<li>Classe : " & GetMaClasse(CStr(Inscription)) & "</li>"
  strHTML = strHTML & "<li>Niveau : " & GetClasseNiveau(GetMaClasse((CStr(Inscription)))) & "</li>"
  strHTML = strHTML & "<li>Année Scolaire : " & GetMonAnneeScolaire(CStr(Inscription)) & " </li>"
  strHTML = strHTML & "</ul>"
  strHTML = strHTML & "<footer>"
  strHTML = strHTML & " <h6> Copyright <i>" & GetSetting("Ecole_nom") & "</i> </h6>"
  strHTML = strHTML & "</footer>"
  strHTML = strHTML & "</BODY>"
  strHTML = strHTML & "</HTML>"
  '-----------------------------------------------------------
  
  If EnvoiEMail(CStr(Destinateur), CStr(Sujet), , CStr(strHTML)) Then
    SendInsMessage = True
  End If
End Function

Function SendWelcomeMessage(Matricule As String) As Boolean
'Envoyer Message de Bienvenue
SendWelcomeMessage = False
  Dim Destinateur
  Dim Sujet
  Dim strHTML
  
  Destinateur = GetEtudiantEmail(CStr(Matricule))
  Sujet = GetSetting("Ecole_nom") & " : Registration Confirmé"
  
  '------------------------------------------------------------------------------
  strHTML = "<HTML>"
  strHTML = strHTML & "<HEAD>"
  strHTML = strHTML & "<BODY>"
  strHTML = strHTML & "<center>"
  strHTML = strHTML & "<b> Confirmation de Registration a Miage Rabat </b></br>"
  strHTML = strHTML & "</center>"
  strHTML = strHTML & "<h3>Informations</h3>"
  strHTML = strHTML & "<ul>"
  strHTML = strHTML & "<li>Nom et Prenom : " & GetMyName(CStr(Matricule)) & " </li>"
  strHTML = strHTML & "<li>Numéro de Matricule : " & Matricule & " </li>"
  strHTML = strHTML & "</ul>"
  strHTML = strHTML & "<footer>"
  strHTML = strHTML & " <h6> Copyright <i>" & GetSetting("Ecole_nom") & "</i> </h6>"
  strHTML = strHTML & "</footer>"
  strHTML = strHTML & "</BODY>"
  strHTML = strHTML & "</HTML>"
  '-------------------------------------------------------------------------------
  
 If EnvoiEMail(CStr(Destinateur), CStr(Sujet), , CStr(strHTML)) Then
   SendWelcomeMessage = True
 End If
End Function

Function EnvoiEmail_Me(Email_Destinateur As String, Sujet As String)
'Sert A Ouvrir Le Navigateur Par Default Pour Envoyer Un Message A Moi

On Error Resume Next
Dim EMail As String

  EMail = "mailto:" & Email_Destinateur & "?subject=" & Sujet

  ShellExecute 0&, vbNullString, EMail, vbNullString, vbNullString, vbNormalFocus

End Function

Function EnvoiEMail(Destinateur As String, Suject As String, Optional Message As String = "", Optional HTML As String = "") As Boolean
'Permet La Re_Initiation Du Mot De Passe Du Profil En Cas D'Oubli
EnvoiEMail = True
    Dim Schemas
    Dim iMsg As Object
    Dim iConf As Object
    Dim Flds As Variant
    
    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
    Schemas = "http://schemas.microsoft.com/cdo/configuration/"
    iConf.Load -1
    Set Flds = iConf.Fields
    
    With Flds
        .Item(Schemas & "smtpusessl") = True
        .Item(Schemas & "smtpauthenticate") = 1
        .Item(Schemas & "sendusername") = "gestioninscriptions@gmail.com"
        .Item(Schemas & "sendpassword") = "halidwalid"
        .Item(Schemas & "smtpserver") = "smtp.gmail.com"
        .Item(Schemas & "sendusing") = 2
        .Item(Schemas & "smtpserverport") = 465 ' 587 '465 '25 587
        .Item(Schemas & "smtpconnectiontimeout") = 10
        .Update
    End With
    
    With iMsg
        Set .Configuration = iConf
        .To = Destinateur
        .CC = ""
        .BCC = ""
        .From = "Gestioninscriptions@gmail.com"
        .Subject = Suject
        ' .AddAttachment FileName1
        If Not HTML = "" Then
          .HTMLBody = HTML
        Else
          .TextBody = Message
        End If
        On Error GoTo errorMsg:
        .Send
    End With
Exit Function
errorMsg:
  EnvoiEMail = False
  MsgBox err.Description
End Function



'##################### HALIDOU CISSE ##################
