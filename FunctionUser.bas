Attribute VB_Name = "Profile"
Option Explicit
Global G_Profil_Login As String
Global G_Profil_Type As String

Function ChampsAddUserOk() As Boolean
'Valide Les Informations Entrées
'-----------------------------------------
With FrmUser
   ChampsAddUserOk = False

   If .TextNom.Text <> "" Then
     If .TextPrenom.Text <> "" Then
       If .TextLogin.Text <> "" Then
         If Len(.TextPass.Text) >= 4 Then
           If Trim(.TextComfirmPass.Text) = (.TextPass.Text) Then
              ChampsAddUserOk = True
           Else
              MsgBox "Mot de passe de confirmation Incorrect !!", vbInformation + vbOKOnly
              .TextComfirmPass.Text = ""
              .TextPass.Text = ""
              .TextPass.SetFocus
           End If
         Else
            MsgBox "Mot de passe très court !!", vbInformation + vbOKOnly
            .TextPass.SetFocus
         End If
       Else
          MsgBox "Le Nom d'utilisateur ne doit pas etre Vide !!", vbInformation + vbOKOnly
          .TextLogin.SetFocus
       End If
     Else
        MsgBox "Prénom Vide !!", vbInformation + vbOKOnly
        .TextPrenom.SetFocus
     End If
   Else
      MsgBox "Nom Vide !!", vbInformation + vbOKOnly
      .TextNom.SetFocus
   End If
  
End With
End Function

'-----------------------Interfaces---------------------------------

Function OpenCherProfil()
'Permet de chercher Un Profil Utilisateur
Dim x
Dim UserCher
   If G_Profil_Type = "ADMIN" Then
debut:
    UserCher = Trim(InputBox("Entrez Le Login du Profil :"))
    If UserCher <> "" Then
        If GetProfilInfos(CStr(UserCher)) Then
           FrmUser.Caption = "INFORMATIONS DE " & UserCher
           FrmUser.TextLogin.Enabled = False
           FrmUser.Show 1
        Else
          x = MsgBox("Profil Inexistant !!!", vbCritical + vbRetryCancel)
          If x = vbRetry Then
             GoTo debut
          End If
        End If
    End If
Else
  MsgBox "Vous N'Avez Pas Le Privilège de Voir Les Informations d'un Profil", vbInformation
End If
End Function

Function OpenSupProfil(Optional Login As String = "")
'Ouvrir Un Dialogue Pour Supprimer Un Profil D'Utilisateur
Dim y
  If Login = "" Then
DebutSup:
    Login = Trim(InputBox("Entrer Le Login de l'utilisateur à Supprimer :"))
  ElseIf Login = "" Then
  ElseIf Not ProfilExist(CStr(Login)) Then
    y = MsgBox("Utilisateur Inéxistant", vbInformation + vbRetryCancel)
    If y = vbRetry Then
      GoTo DebutSup
    End If
  ElseIf GetProfilType(CStr(Login)) = "ADMIN" And Not Login = G_Profil_Login Then
    MsgBox "Vous Ne Pouvez pas Supprimer un Compte d'utilisateur Administrateur !!", vbCritical + vbOKOnly
  Else
    y = MsgBox("Supprimer Cet Profil d'utilisateur ?  " & GetProfilName(CStr(Login)), vbYesNo + vbQuestion)
    If y = vbYes Then
      AddEvent ("Profil Supprimé : " & GetProfilName(CStr(Login)) & " " & Login)
      Call SupprimerProfil(CStr(Login))
      MsgBox "Profil Supprimé Avec Succès", vbOKOnly + vbInformation
    End If
  End If
End Function

Function OpenModPass()
'Pour Modifier Le Mot De Passe de L'Utilisateur Actuel
Dim ProfilCurrentPass
Dim NewPass
Dim VerNewPass
Dim x
debut:
     ProfilCurrentPass = Trim(InputBox("Entrer votre mot de passe actuel : "))
        If Not ProfilCurrentPass = "" Then
             If ProfilCurrentPass = GetProfilPass(CStr(G_Profil_Login)) Then
DebutNewPass:
                  NewPass = Trim(InputBox("Entrez votre nouvel mot de passe : "))
                 If Not NewPass = "" Then
                    If Len(NewPass) > 3 Then
                           VerNewPass = Trim(InputBox("Re-entrez votre nouvel mot de passe : "))
                           '----------------------------------------
                           If Not VerNewPass = "" Then
                            If VerNewPass = NewPass Then
                               x = MsgBox("Ete Vous Sure de Changer Votre Mot de Passe ?", vbInformation + vbYesNo)
                               If x = vbYes Then
                                   If ModProfilPass(G_Profil_Login, CStr(NewPass)) Then
                                      MsgBox "Votre mot de passe à été changé avec succès", vbInformation + vbOKOnly
                                    Else
                                      MsgBox "Erreur , Votre Mot de Passe n'a pas été Changer !", vbCritical + vbOKOnly
                                    End If
                               End If
                            Else
                               MsgBox "Mot de Passe de verification Incorrect !!", vbInformation + vbOKOnly
                               GoTo DebutNewPass
                            End If
                           End If
                           '-------------------------------------------
                    Else
                      MsgBox "Mot de Passe doit etre superieur à 4 chiffres ou lettres !!!", vbInformation + vbOKOnly
                       GoTo DebutNewPass
                    End If
                 End If
             Else
                 x = MsgBox("Mot de passe Incorrect !!!", vbInformation + vbRetryCancel)
                 If x = vbRetry Then
                      GoTo debut
                 End If
            End If
       End If
End Function

Function OpenModProfil(Optional Login As String = "")
'Permet De Modifier Les Informations D'Un Profil
Dim x
With FrmUser
  If G_Profil_Type = "ADMIN" Then
debut:
     If Login = "" Then
       Login = Trim(InputBox("Entrer Le Login de l'utilisateur a Modifier :"))
     End If
     If Login <> "" Then
         If GetProfilInfos(CStr(Login)) Then
             .Caption = "MODIFIER LES INFORMATIONS DU PROFIL DE " & GetProfilName(CStr(Login))
             .TextLogin.Enabled = False
             .Show 1
        Else
            x = MsgBox("UTILISATEUR INEXISTANT !!!", vbCritical + vbRetryCancel)
            If x = vbRetry Then
                 GoTo debut
            End If
        End If
    End If
 Else
        GetProfilInfos CStr(G_Profil_Login)
        .Caption = "MODIFIER LES INFORMATIONS DU PROFIL DE " & GetProfilName(CStr(Login))
        '---------------------------
        .TextLogin.Enabled = False
        .OptAdm.Enabled = False
        .OptST.Enabled = False
        .OptBlok.Enabled = False
        .OptAct.Enabled = False
        '---------------------------
        .Show 1
 End If
End With
End Function

Function OpenAddProfil()
'Pour Ajouter Un Nouvel Profil
With FrmUser
   If G_Profil_Type = "ADMIN" Then
   .Caption = "NOUVEL PROFIL"
   .TextLogin.Enabled = True
   .OptAct.Value = True
   .OptST.Value = True
   FrmUser.Show 1
Else
  MsgBox "Vous devez avoir un profil Administrateur pour pouvoir Ajouter un Profil ! ", vbInformation
End If
End With
End Function


'################### MIS A JOUR BASE DE DONNEE #####################

Function EnregProfil(Login As String)
'Permet D'Enregistrer Les Données D'un Nouvel Profil
With FrmUser
  Dim rs As New ADODB.Recordset
  SQL = "SELECT * FROM PROFILES WHERE Login = '" & Login & "'"
  rs.Open SQL, CN, adOpenKeyset, adLockOptimistic, adCmdText
  If rs.EOF Then
    rs.AddNew
    AddEvent ("Nouvel Profil : " & Login)
  End If
    '---------------------------------
    rs![Nom] = UCase(Trim(.TextNom))
    rs![Prenom] = FormatMe(.TextPrenom)
    rs![Login] = .TextLogin
    rs![AdressEmail] = .TxtEmailRec
    rs![Pass] = .TextPass
    '--------------------------------
    If .OptAdm.Value = True Then
       rs![UserType] = "ADMIN"
    Else
       rs![UserType] = "STANDARD"
    End If
    '--------------------------------
    If .OptAct.Value = True Then
       rs![Statut] = "Actif"
    Else
       rs![Statut] = "Desactivé"
    End If
    '--------------------------------
    
    rs.Update
    rs.Close
    Set rs = Nothing
End With
AddEvent ("Profil Modifier: " & Login)
End Function

Function GetProfilInfos(Login As String) As Boolean
With FrmUser
GetProfilInfos = False
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM PROFILES WHERE Login = '" & Login & "' "
    rs.Open SQL, CN
    If Not rs.EOF Then
        '--------------------------------
        .TextNom.Text = rs![Nom]
        .TextPrenom.Text = rs![Prenom]
        .TextLogin = rs![Login]
        .TxtEmailRec = rs![AdressEmail]
        .TextComfirmPass = rs![Pass]
        .TextPass = rs![Pass]
        '--------------------------------
        If rs![UserType] = "ADMIN" Then
          .OptAdm = True
        Else
          .OptST = True
        End If
        '--------------------------------
        If rs![Statut] = "Actif" Then
          .OptAct = True
        Else
          .OptBlok = True
        End If
        '--------------------------------
        GetProfilInfos = True
    End If
    rs.Close
    Set rs = Nothing
End With
End Function

Function ModProfilPass(Login As String, NewPass As String) As Boolean
'Modifie Le Mot de Pass D'un Profil
    On Error GoTo err
    ModProfilPass = True
      ChangeFieldValue "PROFILES", "Pass", "Login", Login, NewPass
    AddEvent ("Changement de Mot de Passe: " & Login)
    Exit Function
err:
    ModProfilPass = False
End Function

Function SupprimerProfil(Login As String)
'Permet de Supprimer Un Profil D'Utilisateur
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM PROFILES WHERE Login='" & Login & "' "
    rs.Open SQL, CN, adOpenKeyset
        If Not rs.EOF Then
           CN.Execute ("DELETE FROM PROFILES WHERE Login = '" & Login & "'")
        End If
 rs.Close
 Set rs = Nothing
 AddEvent ("Profil Supprimer: " & Login)
End Function


'#################  CONNECTION  #################################


Function Connect_Profil(Login As String, Pass As String) As String
'Permet de Connecter Un Profil
  
  If Not GetProfilPass(CStr(Login)) = Pass Then
    Connect_Profil = "Pass Incorrect"
    AddEvent ("Erreur de Login : " & Login)
  ElseIf Not ProfilActif(CStr(Login)) Then
    Connect_Profil = "Inactif"
    AddEvent ("Profil Bloquée : " & Login)
  Else
    '--------------------------------------
    G_Profil_Login = Login
    G_Profil_Type = GetProfilType(CStr(Login))
    Main.mnConnecter.Caption = GetProfilName(CStr(Login))
    Main.mnConnecter.Checked = True
    Main.mnConnecter.Enabled = False
    Main.mnDeconnecter.Visible = True
    AddEvent ("Profil Connecter: " & Login)
    '--------------------------------------
    
    If G_Profil_Type = "ADMIN" Then
        AddAdminParam
    Else
        RmvAdminParam
    End If
    
    Connect_Profil = "Connecter"
  End If
  
End Function

Function Connect_Default_Profil() As Boolean
'Si Aucun Profil Exist Ce Qu il Faut Faire
Connect_Default_Profil = False

G_Profil_Type = "ADMIN"
G_Profil_Login = "ADMINISTRATEUR"

AddAdminParam

    With Main
      .mnConnecter.Visible = False
      .mnDeconnecter.Visible = False
      .mnAddProfile.Visible = True
      .mnSupProfile.Visible = False
      .mnListProfile.Visible = False
      .mnChPass.Visible = False
    End With
        
Connect_Default_Profil = True
AddEvent ("Profil Connecter: " & G_Profil_Login)
End Function

Function DeconnectProfil() As Boolean
'Deconnecter Le Profil Actuel
With Main
DeconnectProfil = False
AddEvent ("Deconnecter: " & G_Profil_Login)
    RmvAdminParam
    G_Profil_Login = ""
    G_Profil_Type = ""
    .mnConnecter.Enabled = True
    .mnConnecter.Caption = "Se Connecter"
    .mnConnecter.Checked = False
    .mnDeconnecter.Visible = False
     '-----------------
     Main.Hide
     If G_FullScreen = True Then
       FrmLogin2.Show 1
     Else
       FrmLogin2.Show
     End If
     '----------------
DeconnectProfil = True
End With
End Function

Function RmvAdminParam()
'Charger Les Parametres D'un Profil Standard
With Main
  '-------- Profils -----------
  .mnAddProfile.Visible = False
  .mnSupProfile.Visible = False
  .mnListProfile.Visible = False
  '-------- Inscriptions --------
  .mnSuppriméINS.Visible = False
  .mnListINS.Visible = False
  '-------- Registrations -------
  .mnSuppriméETUD.Visible = False
  .mnListETUD.Visible = False
  '-------- Payments ------------
  .MnSupPay.Visible = False
  .mnListPay.Visible = False
  '------- Stats ----------------
  .mnStats.Visible = False
  '-------- Impressions ---------
  .mnIMP.Visible = False
  '-------- Parametres ----------
  .mnParametres.Visible = False
  
End With
End Function

Function AddAdminParam()
'Charger Les Parametres D'un Profil Admin
With Main
  '-------- Profils -----------
  .mnAddProfile.Visible = True
  .mnSupProfile.Visible = True
  .mnListProfile.Visible = True
  '-------- Inscriptions --------
  .mnSuppriméINS.Visible = True
  .mnListINS.Visible = True
  '-------- Registrations -------
  .mnSuppriméETUD.Visible = True
  .mnListETUD.Visible = True
  '-------- Payments ------------
  .MnSupPay.Visible = True
  .mnListPay.Visible = True
  '------- Stats ----------------
  .mnStats.Visible = True
  '-------- Impressions ---------
  .mnIMP.Visible = True
  '-------- Parametres ----------
  .mnParametres.Visible = True
End With
End Function

'###################### SERVICES #####################

Function GetMyProfilType(Login As String) As String
'Renvoie Le Type D'Utilisateur Admin ,Standard
  GetMyProfilType = GetFieldValue("PROFILES", "UserType", "Login", Login)
End Function

Function GetProfilName(Login As String) As String
'Renvoie Le Nom De L'Utilisateur Du Profil
  GetProfilName = GetFieldValue("PROFILES", "Prenom", "Login", Login)
  GetProfilName = GetProfilName & " " & GetFieldValue("PROFILES", "Nom", "Login", Login)
End Function

Function GetProfil_RecEmail(Login As String) As String
'Renvoi L'Email de Recuperation de l'utilisateur
  GetProfil_RecEmail = GetFieldValue("PROFILES", "AdressEmail", "Login", Login)
End Function

Function GetProfilPass(Login As String) As String
'Renvoi le Mot De Pass de l'utilisateur
  GetProfilPass = GetFieldValue("PROFILES", "Pass", "Login", Login)
End Function

Function GetProfilType(Login As String) As String
'Renvoie Le Type De Profil
   GetProfilType = GetFieldValue("PROFILES", "UserType", "Login", Login)
End Function

Function ProfilExist(Login As String) As Boolean
'Verifie Si Un Profil Est Bloqué Ou Pas
  If GetFieldValue("PROFILES", "Login", "Login", Login) = Login Then
    ProfilExist = True
  Else
    ProfilExist = False
  End If
End Function

Function Desctiver_Profil(Login As String) As Boolean
'Desactive Le Profil Du Login
  Desctiver_Profil = False
  ChangeFieldValue "PROFILES", "Statut", "Login", Login, "Inactif"
  Desctiver_Profil = True
End Function

Function ProfilActif(Login As String) As Boolean
'Verifie Si Un Profil Est Bloqué Ou Pas
  If GetFieldValue("PROFILES", "Statut", "Login", Login) = "Actif" Then
    ProfilActif = True
  Else
    ProfilActif = False
  End If
End Function

'##################### HALIDOU CISSE ##################
