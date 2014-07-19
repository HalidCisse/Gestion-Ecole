Attribute VB_Name = "Stats"
Option Explicit

'############################ UPDATES #########################


Function UpdateStats()
'Les Label Sur La Form D'Accueil
With Main
  .LabListReg.Caption = GetNbreRegEnreg
  .LabListIns.Caption = GetNbreInsEnCours
  .LabListPay.Caption = GetNbrePayEnreg
  
  .LabNonIns.Caption = GetNbreEtudiantNonInscrit
  .LabNonPay.Caption = GetNbreDetteNonPayer
End With
End Function

Function SyncData()
'Mettre A Jour Et Sync Tous Les Statut Des Informations Des Etudiants
  Dim rs As New ADODB.Recordset
    SQL = "SELECT Matricule FROM ETUDIANTS"
    rs.Open SQL, CN, adOpenKeyset, adLockOptimistic, adCmdText
    '--------------------
    Do While Not rs.EOF
      SyncMesInfos rs![Matricule]
      rs.MoveNext
    Loop
    '--------------------
    rs.Close
    Set rs = Nothing
End Function

Function SyncMesInfos(Matricule As String)
Dim PayParTranche As Long
PayParTranche = CLng(Val(GetFieldValue("DETTES", "Somme_A_Payer", "Matricule", Matricule)))

'Si L'Inscription A Expiré Changer Son Statut A Expiré "
 UpdateMonInsStatut Matricule

'Si Le Numéro D'Inscription Change Le Changer De La Table Etudiant
  ChangeFieldValue "ETUDIANTS", "N°_Inscription", "Matricule", Matricule, GetMonIns(Matricule)

'Si La Classe Change Le Changer De La Table Etudiant
  ChangeFieldValue "ETUDIANTS", "Classe", "Matricule", Matricule, GetMaClasse(GetMonIns(Matricule))
  
'Mettre A Jour Payment Total
  ChangeFieldValue "ETUDIANTS", "Total_Payment", "Matricule", Matricule, GetMonTotalPAYS(Matricule)

'Suspendre En Cas de Defaut de Payment de 3 Mois
If GetSetting("Susp_Ins") = "Oui" And GetMaDette(Matricule) > (3 * PayParTranche) Then
  If Not GetFieldValue("INSCRIPTIONS", "Statut", "N°_Inscription", GetMonIns(Matricule)) = "Suspendue" Then
    ChangeFieldValue "INSCRIPTIONS", "Statut", "N°_Inscription", GetMonIns(Matricule), "Suspendue"
    ChangeFieldValue "INSCRIPTIONS", "DetailSus", "N°_Inscription", GetMonIns(Matricule), "[AUTO] Defaut de Payments"
    AddEvent "[AUTO] Suspension d'Inscription de " & GetMyName(Matricule) & " Pour defaut de Payments"
  End If
End If
End Function

Function UpdateMonInsStatut(Matricule As String)
'Si La Date De Fin De L'Inscription Est Dépasser Changer Son Statut A Expiré
  Dim rs As New ADODB.Recordset
  SQL = "SELECT N°_Inscription FROM INSCRIPTIONS WHERE N°_Matricule_Etudiant = '" & Matricule & "' AND CDate(Ins_Fin) < Date()"
  rs.Open SQL, CN, adOpenDynamic, adLockOptimistic, adCmdText
  '------------------
  Do While Not rs.EOF
    ChangeMonInsStatut rs![N°_Inscription], "Expiré"
  rs.MoveNext
  Loop
  '-----------------
  rs.Close
  Set rs = Nothing
End Function

Function CleanClasse()
'Supprimer Les Classes Sans Effectif
  Dim rs As New ADODB.Recordset
  SQL = "SELECT * FROM CLASSES"
  rs.Open SQL, CN, adOpenDynamic, adLockOptimistic, adCmdText
  '-----------------
  Do While Not rs.EOF
    If GetClasseEffectif(rs![Classe]) = 0 Then
      AddEvent "[AUTO] Classe Supprimée : " & rs![Classe]
      CN.Execute ("DELETE FROM CLASSES WHERE Classe = '" & rs![Classe] & "'")
    End If
  rs.MoveNext
  Loop
  '---------------
  rs.Close
  Set rs = Nothing
End Function

'############################ TESTS ###########################

Function Inscription_Activ(Inscription As String) As Boolean
'Verifie Si L'Inscription Est Active Ou Suspendue
  If Get_Inscription_Statut(CStr(Inscription)) = "Active" Then
    Inscription_Activ = True
  Else
    Inscription_Activ = False
  End If
End Function

Function Etudiant_Regulier(Matricule As String) As Boolean
'Verifie Si L'Etudiant Est Regulier Ou Pas
  If Get_Etudiant_Statut(CStr(Matricule)) = "Regulier" Then
    Etudiant_Regulier = True
  Else
    Etudiant_Regulier = False
  End If
End Function

Function DettePayer(Matricule As String, Designation As String) As Boolean
'Verifie Si La Dette A Ete Payer ou pas
  DettePayer = False
  If Not NombreEnreg("PAYMENTS", "ID_Dette", Designation & Matricule) = 0 Then
    DettePayer = True
  End If
End Function

Function MaDetteExist(Matricule As String, Designation As String) As Boolean
'Verifie si le Numero De Payment Existe
  MaDetteExist = False
  If Not NombreEnreg("DETTES", "ID_Dette", Designation & Matricule) = 0 Then
    MaDetteExist = True
  End If
End Function

Function PayExist(N°_Payment As String) As Boolean
'Verifie si le Numero De Payment Existe
  PayExist = False
  If Not NombreEnreg("PAYMENTS", "N°_Payment", N°_Payment) = 0 Then
    PayExist = True
  End If
End Function

Function MatExist(Matricule As String) As Boolean
'Verifie Si Le Matricule Existe
MatExist = False
If Not NombreEnreg("ETUDIANTS", "Matricule", Matricule) = 0 Then
  MatExist = True
End If
End Function

Function InsExist(Inscription As String) As Boolean
'Verifie si le Numero d'Inscription Existe
InsExist = False
If Not NombreEnreg("INSCRIPTIONS", "N°_Inscription", Inscription) = 0 Then
  InsExist = True
End If
End Function

Function NumIdExist(NumID As String) As Boolean
'Verifie Si Le Numéro De L'Identité Existe
  NumIdExist = False
  If Not NombreEnreg("ETUDIANTS", "Numero_Identite", NumID) = 0 Then
    NumIdExist = True
  End If
End Function

Function EstInscrit(Matricule As String) As Boolean
'Determine Si L'Etudiant Est Inscrit Ou Pas
  If GetMonIns(Matricule) = "NULL" Then
    EstInscrit = False
  Else
    EstInscrit = True
  End If
End Function

Function Ins_Pas_Expire(Inscription As String) As Boolean
'Determine Si L'Inscription Est En Cours Ou Terminer
  Dim rs As New ADODB.Recordset
    SQL = "SELECT Ins_Fin FROM INSCRIPTIONS WHERE N°_Inscription  = '" & Inscription & "' "
    rs.Open SQL, CN, adOpenKeyset
    '---------------------
    If Not rs.EOF Then
      If CDate(rs(0)) < Date Then
        Ins_Pas_Expire = False
      Else
        Ins_Pas_Expire = True
      End If
    Else
      Ins_Pas_Expire = False
    End If
    '-------------------
    rs.Close
    Set rs = Nothing
End Function


'############################ GET  ############################


Function GetNbreDetteNonPayer() As Integer
'Renvie Le Nombre de Dettes Non Payer
Dim SQL2
    GetNbreDetteNonPayer = 0
    Dim rs As New ADODB.Recordset
    Dim RS2 As New ADODB.Recordset
    SQL = "SELECT Matricule FROM ETUDIANTS"
    rs.Open SQL, CN
      Do While Not rs.EOF
        SQL2 = "SELECT Designation FROM DETTES WHERE Matricule = '" & rs![Matricule] & "' AND CDate(Date_Payment) <= Date() "
        RS2.Open SQL2, CN
            Do While Not RS2.EOF
              If Not DettePayer(rs![Matricule], RS2![Designation]) Then
                GetNbreDetteNonPayer = GetNbreDetteNonPayer + 1
              End If
            RS2.MoveNext
            Loop
        RS2.Close
        Set RS2 = Nothing
      rs.MoveNext
      Loop
    rs.Close
    Set rs = Nothing
End Function

Function GetNbrePayEnreg() As Integer
  GetNbrePayEnreg = NombreEnreg("PAYMENTS", "", "Tous")
End Function

Function GetNbreRegEnreg() As Integer
  GetNbreRegEnreg = NombreEnreg("ETUDIANTS", "", "Tous")
End Function

Function GetNbreEtudiantNonInscrit() As Integer
'Renvoi Le Nombre d'Etudiant Non Inscrit
GetNbreEtudiantNonInscrit = 0
  Dim rs As New ADODB.Recordset
    SQL = "SELECT Matricule FROM ETUDIANTS WHERE Statut = 'Regulier' "
    rs.Open SQL, CN
      Do While Not rs.EOF
        If Not EstInscrit(rs![Matricule]) Then
          GetNbreEtudiantNonInscrit = GetNbreEtudiantNonInscrit + 1
        End If
      rs.MoveNext
      Loop
rs.Close
Set rs = Nothing
End Function

Function GetNbreInsEnCours() As Integer
'Renvoi Le Nombre D'Inscriptions En Cours
  GetNbreInsEnCours = NombreEnreg("INSCRIPTIONS", "SQL", "SELECT COUNT(*) AS Nombre FROM INSCRIPTIONS WHERE CDate(Ins_Fin) > Date()")
End Function

Function GetCurrentAnneeScolaire() As String
'Renvoi L'Année Scolaire En Cours
  If Month(Date) < 8 Then
    GetCurrentAnneeScolaire = Year(Date) - 1 & "/" & Year(Date)
  Else
    GetCurrentAnneeScolaire = Year(Date) & "/" & Year(Date) + 1
  End If
End Function

'---------------------------------------------------------

Function Get_Etudiant_Statut(Matricule As String) As String
'Renvoie Le Statut De L'Etudiant
  Get_Etudiant_Statut = GetFieldValue("ETUDIANTS", "Statut", "Matricule", Matricule)
End Function

Function Get_Inscription_Statut(Inscription As String) As String
'Renvoie Le Statut De L'Inscription
  Get_Inscription_Statut = GetFieldValue("INSCRIPTIONS", "Statut", "N°_Inscription", Inscription)
End Function

Function GetPayMat(PayID As String) As String
'Return Le Matricule De Celui A Effectuer Le Payment
  GetPayMat = GetFieldValue("PAYMENTS", "N°_Matricule", "N°_Payment", PayID)
End Function

Function GetDette_Som(ID_Dette As String) As String
'Return La Designation Du N° De La Dette
  GetDette_Som = GetFieldValue("DETTES", "Somme_A_Payer", "ID_Dette", ID_Dette)
End Function

Function GetDette_Des(ID_Dette As String) As String
'Return La Designation Du N° De La Dette
  GetDette_Des = GetFieldValue("DETTES", "Designation", "ID_Dette", ID_Dette)
End Function

Function GetPay_Des(PayID As String) As String
'Return La Designation Du N° De Payment
  GetPay_Des = GetFieldValue("PAYMENTS", "Designation", "N°_Payment", PayID)
End Function

Function GetClasseEffectif(Classe As String) As Integer
'Renvoie L'Effectif Actuelle De La Classe
  'GetClasseEffectif = NombreEnreg("INSCRIPTIONS", "Classe", Classe)
  GetClasseEffectif = NombreEnreg("INSCRIPTIONS", "SQL", "SELECT COUNT(*) AS Nombre FROM INSCRIPTIONS WHERE Statut = 'Active' AND Classe = '" & Classe & "'")
End Function

Function GetNbreClasse() As Integer
'Return Le Nombre De Classe Active avec effectif
GetNbreClasse = 0
  Dim rs As New ADODB.Recordset
  SQL = "SELECT DISTINCT Classe FROM INSCRIPTIONS WHERE Statut = 'Active'"
  rs.Open SQL, CN, adOpenKeyset
  '-------------
  Do While Not rs.EOF
    GetNbreClasse = GetNbreClasse + 1
  rs.MoveNext
  Loop
  '------------
  rs.Close
  Set rs = Nothing
End Function

Function GetMyName(Matricule As String) As String
'Trouver Le Nom Et Prenom De L'Etudiant A Partir Du Matricule
  Dim rs As New ADODB.Recordset
    SQL = "SELECT Nom, Prenom FROM ETUDIANTS WHERE Matricule = '" & UCase(Matricule) & "' "
    rs.Open SQL, CN, adOpenKeyset
      '----------------
      If Not rs.EOF Then
        GetMyName = FormatMe(rs![Prenom]) & " " & UCase(rs![Nom])
      Else
        GetMyName = ""
      End If
      '---------------
    rs.Close
    Set rs = Nothing
End Function

Function GetMaMat(Name As String) As String
'Trouver Mon Matricule A Partir de Mon Nom
GetMaMat = ""
Dim N
Dim rs As New ADODB.Recordset
    SQL = "SELECT Matricule, Nom, Prenom FROM ETUDIANTS "
    rs.Open SQL, CN
      '--------------------
      Do While Not rs.EOF
        N = rs![Prenom] & " " & rs![Nom]
        If LCase(N) = LCase(Name) Then
          GetMaMat = rs![Matricule]
          Exit Do
        End If
      rs.MoveNext
      Loop
      '-------------------
rs.Close
Set rs = Nothing
End Function

Function GetMonIns(Matricule As String) As String
'Return Le Numero D'Inscription de L'Etudiant Du Matricule
  Dim rs As New ADODB.Recordset
    SQL = "SELECT N°_Inscription FROM INSCRIPTIONS WHERE N°_Matricule_Etudiant = '" & Matricule & "' AND CDate(Ins_Fin) > Date () "
    rs.Open SQL, CN, adOpenKeyset
    '---------------------
      If Not rs.EOF Then
        GetMonIns = rs(0)
      Else
        GetMonIns = "NULL"
      End If
    '---------------------
    rs.Close
    Set rs = Nothing
End Function

Function GetMaMatricule(Inscription As String) As String
'Return Le Numéro De Matricule Du N° De L'Inscription
  GetMaMatricule = GetFieldValue("INSCRIPTIONS", "N°_Matricule_Etudiant", "N°_Inscription", Inscription)
End Function

Function GetMaClasse(Inscription As String) As String
'Return La Classe De L'Etudiant De L'Inscription
  GetMaClasse = GetFieldValue("INSCRIPTIONS", "Classe", "N°_Inscription", Inscription)
End Function

Function GetClasseNiveau(Classe As String) As String
'Return Le Niveau De L'Etudiant De L'Inscription
  GetClasseNiveau = GetFieldValue("CLASSES", "Niveau_Filiere", "Classe", Classe)
End Function

Function GetEtudiantEmail(Matricule As String) As String
'Renvoi L'Adresse Email De L'Etudiant
  GetEtudiantEmail = GetFieldValue("ETUDIANTS", "Email", "Matricule", Matricule)
End Function

Function GetMonAnneeScolaire(Inscription As String) As String
'Returne Mon Année Scolaire Actuel
  GetMonAnneeScolaire = GetFieldValue("INSCRIPTIONS", "Annnee_Scolaire", "N°_Inscription", Inscription)
End Function

Function GetInscriptionStatut(Inscription As String) As String
'Renvoi le statut de l'inscription
  GetInscriptionStatut = GetFieldValue("INSCRIPTIONS", "Statut", "N°_Inscription", Inscription)
End Function

Function GetPayDes(Payment_ID As String) As String
'Renvoi La Designation Du Payment
  GetPayDes = GetFieldValue("PAYMENTS", "Designation", "N°_Payment", Payment_ID)
End Function

Function GetPaySum(Payment_ID As String) As Long
'Renvoi La Somme Du Payment
  GetPaySum = CLng(GetFieldValue("PAYMENTS", "Somme_Payer", "N°_Payment", Payment_ID))
End Function

Function GetPayDate(Payment_ID As String) As Date
'Renvoi La Date Du Payment
  GetPayDate = CDate(GetFieldValue("PAYMENTS", "Date_Payment", "N°_Payment", Payment_ID))
End Function

Function GetMaDette(Matricule As String) As Long
'Return La Dette De L'Etudiant
GetMaDette = CLng(GetMonTotalDette(Matricule)) - CLng(GetMonTotalPAYS(Matricule))
End Function

Function GetMonTotalDette(Matricule) As Long
'Return La Somme De Tous Les Dettes Affectes A Ce Etudiant
GetMonTotalDette = 0
Dim rs As New ADODB.Recordset
    SQL = "SELECT SUM(Somme_A_Payer) FROM DETTES WHERE Matricule  = '" & Matricule & "' AND CDate(Date_Payment) <= Date()"
    rs.Open SQL, CN, adOpenKeyset
      GetMonTotalDette = CLng(Val(rs(0) & ""))
    rs.Close
    Set rs = Nothing
End Function

Function GetMonTotalPAYS(Matricule As String) As Long
'Return La Somme De Tous Les Payments Effectuées Par ce Etudiant
GetMonTotalPAYS = 0
  Dim rs As New ADODB.Recordset
    SQL = "SELECT SUM(Somme_Payer) FROM PAYMENTS WHERE N°_Matricule  = '" & UCase(Matricule) & "' "
    rs.Open SQL, CN, adOpenKeyset
      GetMonTotalPAYS = CLng(Val(rs(0) & ""))
    rs.Close
    Set rs = Nothing
End Function

Function GetSetting(Parametre As String) As String
'Return La Valeur D'Un Enregistrement
  Dim rs As New ADODB.Recordset
    SQL = "SELECT Valeur FROM PARAMETRES WHERE Parametre = '" & Parametre & "'"
    rs.Open SQL, CN
      '---------------
      If Not rs.EOF Then
        GetSetting = Trim(rs(0) & " ")
      Else
        GetSetting = ""
      End If
      '--------------
    rs.Close
    Set rs = Nothing
End Function

Function GetFieldValue(TABLE As String, SelectedField As String, Key As String, KeyVal As String) As String
'Return La Valeur D'Un Enregistrement
  Dim rs As New ADODB.Recordset
    SQL = "SELECT " & SelectedField & " FROM " & TABLE & " WHERE " & Key & "  = '" & UCase(KeyVal) & "' "
    rs.Open SQL, CN, adOpenKeyset
      '-------------------
      If Not rs.EOF Then
        GetFieldValue = rs(0)
      Else
        GetFieldValue = ""
      End If
      '-------------------
    rs.Close
    Set rs = Nothing
End Function

'############################ CHANGE ##########################

Function ChangeMonInsStatut(Inscription As String, Statut As String)
'Permet de Changer Le Statut De L'Inscription : Active ,Suspendu , Expiré
  Call ChangeFieldValue("INSCRIPTIONS", "Statut", "N°_Inscription", Inscription, Statut)
End Function

Function Change_Setting(Parametre As String, Valeur As String)
'Return La Valeur D'Un Enregistrement
  Dim rs As New ADODB.Recordset
    SQL = "SELECT Valeur FROM PARAMETRES WHERE Parametre = '" & Parametre & "' "
    rs.Open SQL, CN, adOpenKeyset, adLockOptimistic, adCmdText
      '--------------------
      If Not rs.EOF Then
        rs.Fields(0) = Valeur
        rs.Update
      End If
      '--------------------
    rs.Close
    Set rs = Nothing
End Function

Function ChangeFieldValue(TABLE As String, SelectedField As String, Key As String, KeyVal As String, SeFdNewValue As String)
'Changer La Valeur D'Une Cellule
  Dim rs As New ADODB.Recordset
    SQL = "SELECT " & SelectedField & " FROM " & TABLE & " WHERE " & Key & "  = '" & UCase(KeyVal) & "' "
    rs.Open SQL, CN, adOpenKeyset, adLockOptimistic, adCmdText
      '----------------
      If Not rs.EOF Then
        rs.Fields(0) = SeFdNewValue
        rs.Update
      End If
      '----------------
    rs.Close
    Set rs = Nothing
End Function

'######################### Forms ###########################################

Function LoadForms()
    Load Main
End Function

Function UnloadForms()
'Unload Tous Les Forms Dans Le Project
On Error Resume Next
    Dim frm As Form
    Dim ctl As Control
    For Each frm In Forms
        frm.Hide
        For Each ctl In frm.Controls
            Set ctl = Nothing
        Next ctl
        Unload frm
        Set frm = Nothing
    Next frm
End Function

Function FormatMe(Str As String) As String
  FormatMe = Trim(UCase(Mid(Str, 1, 1)) + LCase(Mid(Str, 2, Len(Str) - 1)))
End Function

'##################### HALIDOU CISSE ##################
