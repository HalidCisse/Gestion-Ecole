Attribute VB_Name = "Etudiant"
Option Explicit
Public Options As String


Function ChampsEtudiantOk() As Boolean
'Verifie Si Les Informations Entrées Sur La Forme Sont Valident
  ChampsEtudiantOk = False
  With FrmEtudiant
    '-----------------
    If Options = "Ajouter" Then
    'S'Il S'Agit D'Ajouter, Faut Verifier Le Matricule (Unique !)
      If MatExist(Trim(UCase(.TxtMat.Text))) Then
         MsgBox "Un Etudiant Est Déja Enrégistreé Avec Ce Matricule : " & GetMyName(Trim(.TxtMat.Text))
         QuelPageEtudiant 2
        .TxtMat.SetFocus
         Exit Function
      End If
    End If
    '-----------------
  If Trim(.TxtNumID.Text) = "" Then
    MsgBox "Le Numéro D'Identité Ne Doit Pas Etre Null !"
    QuelPageEtudiant (1)
    .TxtNumID.SetFocus
  ElseIf Not ValidEmail(Trim(.TxtEmail.Text)) And Not Trim(.TxtEmail.Text) = "" Then
    MsgBox "L'Adress Email N'Est Pas Valid !"
    QuelPageEtudiant 1
    .TxtEmail.SetFocus
  ElseIf Trim(.TxtMat.Text) = "" Then
     MsgBox "Le Numéro De Matricule Ne Doit Pas Etre Null !"
     QuelPageEtudiant 2
    .TxtMat.SetFocus
  ElseIf Not isString(.TxtNom.Text) Then
    MsgBox "Le Nom N'Est Pas Valide !"
    QuelPageEtudiant (1)
    .TxtNom.SetFocus
  ElseIf Not isString(.TxtPrenom.Text) Then
    MsgBox "Le Prénom N'Est Pas Valide !"
    QuelPageEtudiant (1)
    .TxtPrenom.SetFocus
  Else
   ChampsEtudiantOk = True
  End If
  '------------------
End With
End Function


'############### OPERATIONS DE GESTION ######################

Function OpenAddEtudiant()
'Ouvrir La Form Pour Ajouter Un Etudiant
Options = "Ajouter"
  With FrmEtudiant
    .Caption = "NOUVELLE REGISTRATION ETUDIANT"
    .TxtMat.Text = GenNewMAT
    .TxtMat.ToolTipText = "Automatiquement Génerer !"
    .OptionCIN.Value = True
    .OptionH.Value = True
    .DTPNaiss.MaxDate = Date
    .DTPNaiss.Value = DateSerial(1990, 1, 1)
    .ComboNat = "Marocaine"
    .ComboNEtude = "BAC"
    .LabelDateEnreg1.Visible = False
    .LabelDateEnreg.Visible = False
    .OptRegulier.Value = True
  End With
FrmEtudiant.Show 1
End Function

Function OpenModEtudiant(Optional Matricule As String = "")
'Ouvrir la Form Pour Modifier les Informations d'un Etudiant
Options = "Modifier"
If Matricule = "" Then
  Dim MAT As String
  Dim y
DebutSup:
   MAT = Trim(InputBox("Entrez Le Matricule de L'Etudiant : "))
   If MAT <> "" Then
    If MatExist(CStr(MAT)) Then
        FrmEtudiant.Caption = "REGISTRATION DE " & UCase(GetMyName(Matricule))
        GetEtudiantInfos (MAT)
        FrmEtudiant.TxtMat.Enabled = False
        FrmEtudiant.Show 1
    Else
      y = MsgBox("Etudiant Non Trouver !", vbCritical + vbRetryCancel)
      If y = vbRetry Then
        GoTo DebutSup
      End If
    End If
   End If
Else
  If MatExist(Matricule) Then
    FrmEtudiant.Caption = "REGISTRATION DE " & UCase(GetMyName(Matricule))
    GetEtudiantInfos (Matricule)
    FrmEtudiant.TxtMat.Enabled = False
    FrmEtudiant.Show 1
  Else
    MsgBox "Les Informations de Ce Etudiant N'Existe Plus !", vbCritical + vbRetryCancel
  End If
End If
End Function

Function OpenSupEtudiant(Optional MAT As String = "")
'Ouvrir Un Dialogue Pour Supprimer Les Infos D'Un Etudiant
Dim y
  If MAT = "" Then
DebutSup:
     MAT = Trim(InputBox("Entrez Le Matricule de L'Etudiant a Supprimer "))
   End If
   If MAT <> "" Then
    If MatExist(CStr(MAT)) Then
       y = MsgBox("Confirmé La Suppression de la Registration de " & GetMyName(CStr(MAT)), vbYesNo + vbQuestion)
       If y = vbYes Then
         If SupprimerEtudiant(CStr(MAT)) Then
           MsgBox "Suppression réussie", vbOKOnly + vbInformation
         Else
           MsgBox "Desolé ,Vous n'avez pas le privilège de supprimé un Etudiant !", vbOKOnly + vbInformation
         End If
       End If
    Else
      y = MsgBox("Matricule Inéxistant", vbCritical + vbRetryCancel)
      If y = vbRetry Then
        GoTo DebutSup
      End If
    End If
   End If
End Function


'##################### MIS A JOUR COMPLETE BASE DE DONNEE ############################

Function GetEtudiantInfos(Matricule As String) As Boolean
'Permet de Remplir La Forme des Informations Enregistrées de l'Etudiant
GetEtudiantInfos = False
  Dim rs As New ADODB.Recordset
  SQL = "SELECT * FROM ETUDIANTS WHERE Matricule = '" & Matricule & "'"
  rs.Open SQL, CN, adOpenKeyset, adLockOptimistic, adCmdText
  If Not rs.EOF Then
    SyncMesInfos Matricule
    With FrmEtudiant
    
    '--- Frame 1 -------------
     If rs![TypeID] = "Carte Sejour" Then
             .OptionCS.Value = True
          ElseIf rs![TypeID] = "Passport" Then
             .OptionPass.Value = True
          ElseIf rs![TypeID] = "CIN" Then
             .OptionCIN.Value = True
     End If
     .TxtNumID = rs![Numero_Identite]
     .DTPExpID.Value = rs![Date_Expire]
     
     '--- Frame 2 ------------
     If rs![Sex] = "Femme" Then
             .OptionF.Value = True
          ElseIf rs![Sex] = "Homme" Then
             .OptionH.Value = True
     End If
     .TxtNom.Text = rs![Nom]
     .TxtPrenom.Text = FormatMe(rs![Prenom])
     .DTPNaiss.Value = rs![Date_Naissance]
     .ComboLieuNaiss = rs![Lieu_Naissance]
     .ComboNat = rs![Nationalite]
     .TxtNomP.Text = rs![Nom_Pere]
     .TxtNomM.Text = rs![Nom_Mere]
     .TxtNumTel.Text = rs![TEL]
     .TxtEmail.Text = rs![EMail]
     .TxtAdress.Text = rs![Adresse]
     
     '---- Frame 3 ----------
     .TxtNomT.Text = rs![Nom_Tuteur]
     .TxtPrenomT.Text = rs![Prenom_Tuteur]
     .TxtTelT.Text = rs![TEL_Tuteur]
     .TxtEmailT.Text = rs![Email_Tuteur]
     .TxtAddressT.Text = rs![Adresse_Tuteur]
     
     '---- Frame 4 ----------
     .ComboNEtude = rs![Niveau_Etude] & ""
     .TxtMat.Text = rs![Matricule]
     .LabelNIns.Caption = rs![N°_Inscription] & ""
     .LabelClasse.Caption = rs![Classe]
     .LabelTotalPayment.Caption = rs![Total_Payment] & "  Dh"
     .LabelDateEnreg.Caption = rs![Date_Enregistrement]
     
     If rs![Statut] = "Regulier" Then
         .OptRegulier.Value = True
     ElseIf rs![Statut] = "Non Regulier" Then
         .OptNonRegulier.Value = True
     ElseIf rs![Statut] = "Abandonné" Then
         .OptAband.Value = True
     ElseIf rs![Statut] = "Diplomé" Then
         .OptDiplome.Value = True
     End If
    
    rs.Close
    GetEtudiantInfos = True
    End With
  End If
End Function

Function EnregEtudiantInfos(Matricule As String)
'Permet des Operation de Modification et de D'Ajout de nouveau Etudiants
  Dim rs As New ADODB.Recordset
  SQL = "SELECT * FROM ETUDIANTS WHERE Matricule = '" & Matricule & "'"
  rs.Open SQL, CN, adOpenKeyset, adLockOptimistic, adCmdText
  With FrmEtudiant
  If rs.EOF Then
    'Si L'Etudiant N'Exist Pas L'Ajouter !
    rs.AddNew
    rs![Matricule] = Matricule
    rs![Date_Enregistrement] = Date
    AddEvent ("Etudiant Ajouté : " & Matricule)
  End If
   '--------- Frame 1  ---------------
   If .OptionCS.Value = True Then
           rs![TypeID] = "Carte Sejour"
        ElseIf .OptionPass.Value = True Then
           rs![TypeID] = "Passport"
        Else
           rs![TypeID] = "CIN"
   End If
   rs![Numero_Identite] = UCase(.TxtNumID)
   rs![Date_Expire] = .DTPExpID.Value
   
   '--- Frame 2 ---------------------
   If .OptionF = True Then
           rs![Sex] = "Femme"
        Else
           rs![Sex] = "Homme"
   End If
   rs![Nom] = UCase(Trim(.TxtNom.Text))
   rs![Prenom] = FormatMe(Trim(.TxtPrenom.Text))
   rs![Date_Naissance] = .DTPNaiss.Value
   rs![Lieu_Naissance] = UCase(.ComboLieuNaiss)
   rs![Nationalite] = UCase(.ComboNat)
   rs![Nom_Pere] = UCase(.TxtNomP.Text)
   rs![Nom_Mere] = UCase(.TxtNomM.Text)
   rs![TEL] = .TxtNumTel.Text
   rs![EMail] = .TxtEmail.Text
   rs![Adresse] = .TxtAdress.Text
   
   '---- Frame 3 ----------
   rs![Nom_Tuteur] = UCase(.TxtNomT.Text)
   rs![Prenom_Tuteur] = UCase(.TxtPrenomT.Text)
   rs![TEL_Tuteur] = .TxtTelT.Text
   rs![Email_Tuteur] = .TxtEmailT.Text
   rs![Adresse_Tuteur] = UCase(.TxtAddressT.Text)
   
   '---- Frame 4 ----------
   rs![Niveau_Etude] = .ComboNEtude

   If .OptRegulier.Value = True Then
        rs![Statut] = "Regulier"
    ElseIf .OptNonRegulier.Value = True Then
        rs![Statut] = "Non Regulier"
    ElseIf .OptAband.Value = True Then
        rs![Statut] = "Abandonné"
    ElseIf .OptDiplome.Value = True Then
        rs![Statut] = "Diplomé"
   End If
   '----------------------
  rs.Update
  rs.Close
End With
AddEvent "Registration Modifié : " & GetMyName(Matricule) & " MAT: " & Matricule
End Function

Function SupprimerEtudiant(Matricule As String) As Boolean
'Supprimer l'etudiant de la matricule
Dim N
'---------------------------------
If Not G_Profil_Type = "ADMIN" Then
  SupprimerEtudiant = False
  AddEvent ("Suppression Registration Bloqué : " & GetMyName(Matricule))
  Exit Function
End If
'---------------------------------
N = GetMyName(Matricule) & " MAT: " & Matricule
'---------------------------------
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM ETUDIANTS WHERE Matricule = '" & UCase(Matricule) & "' "
    rs.Open SQL, CN, adOpenKeyset
      If Not rs.EOF Then
        CN.Execute ("DELETE FROM ETUDIANTS WHERE Matricule = '" & UCase(Matricule) & "'")
      End If
    rs.Close
    Set rs = Nothing
    '-----------------------------
AddEvent ("Registration Supprimé : " & N)
SupprimerEtudiant = True
End Function

'##############################################################################

Function ScaleEtudiant()
'Organiser Les Pages de La Forme
 With FrmEtudiant
    Page = 1
    .Width = 7335
    .Height = 9255
    .CmdPrec.Top = 8280
    .CmdPrec.Left = .Frame1.Left
    .CmdSuiv.Top = .CmdPrec.Top
    .CmdSuiv.Left = 5400
    .CmdRaz.Left = .CmdPrec.Left
    .CmdRaz.Top = .CmdPrec.Top
    .CmdEnreg.Left = .CmdSuiv.Left
    .CmdEnreg.Top = .CmdSuiv.Top
    .Frame3.Left = .Frame1.Left
    .Frame3.Top = .Frame1.Top
    .Frame4.Left = .Frame1.Left
   QuelPageEtudiant (Page)
 End With
End Function

Function QuelPageEtudiant(PageCourant As Integer)
'Monter Quel Page ?
 With FrmEtudiant
    If PageCourant = 1 Then
        .Frame1.Visible = True
        .Frame2.Visible = True
        '----------------------
        .Frame3.Visible = False
        .Frame4.Visible = False
        '----------------------
        .CmdRaz.Visible = True
        .CmdPrec.Visible = False
        .CmdEnreg.Visible = False
        .CmdSuiv.Visible = True
        Page = 1
    ElseIf PageCourant = 2 Then
        .Frame1.Visible = False
        .Frame2.Visible = False
        '----------------------
        .Frame3.Visible = True
        .Frame4.Visible = True
        '----------------------
        .CmdEnreg.Visible = True
        .CmdPrec.Visible = True
        .CmdSuiv.Visible = False
        .CmdRaz.Visible = False
        Page = 2
    End If
 End With
End Function

Function GenNewMAT() As String
'Generez un nouveau numéro de matricule
Dim i As Integer
Dim Random As Integer
Dim NewMat As String
rep:
    For i = 1 To 3
        Randomize
        Random = CStr(CInt(89 * Rnd) + 10)
        NewMat = NewMat & CStr(Random)
    Next i
   GenNewMAT = "M" & "-" & UCase(NewMat)
   If MatExist(CStr(GenNewMAT)) Then
      GoTo rep
   End If
End Function

'##################### HALIDOU CISSE ##################
