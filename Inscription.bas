Attribute VB_Name = "Inscription"
Option Explicit
Public OprionsIns As String


Function ChampsInsOk() As Boolean
'Verifie Les Informations Entrées Sur La Form Inscription
Dim MAT
Dim NM
MAT = GetMaMat(FrmInscription.ComboNumMat)
'-----------------------------------------
  ChampsInsOk = False
  With FrmInscription
  '----------------------------------
  If Options = "Ajouter" Then
    'S'Il S'Agit D'Ajouter, Faut Verifier Le N° Inscription (Unique !)
      If InsExist(Trim(UCase(.TxtNumIns.Text))) Then
         MsgBox "Un Etudiant Est Déja Inscrit Avec Ce N° : " & GetMyName(GetMaMatricule(Trim(.TxtNumIns.Text)))
        .TxtNumIns.SetFocus
         Exit Function
      End If
    End If
  '----------------------------------
  NM = DateDiff("m", .DTPickerDeb.Value, .DTPickerFin.Value)
  If .ComboTypeReg = "Trimestriel" And Not (NM Mod 3 = 0) Then
    MsgBox "Si le Type de Reglement est Trimestriel, Le Nombre de Mois Doit Etre Multiple de Trois !"
    .DTPickerFin.SetFocus
  Exit Function
  End If
  '----------------------------------
  If Not Trim(.TxtNumIns.Text) = "" Then
    If MatExist(Trim(UCase(MAT))) Then
      If Not Trim(.ComboClasse) = "" Then
        If Not Trim(.ComboNiveau) = "" Then
          If Not Trim(.ComboFiliere) = "" Then
            If Not Trim(.ComboAnnee) = "" Then
                If .DTPickerDeb < .DTPickerFin Then
                  ChampsInsOk = True
                Else
                  MsgBox "La Date De Fin Doit Etre Supérieur Au Date De Début !"
                End If
            Else
              MsgBox "L'Année Ne Doit Pas Etre Null !"
              .ComboAnnee.SetFocus
            End If
          Else
            MsgBox "La Filière Ne Doit Pas Etre Null !"
            .ComboFiliere.SetFocus
          End If
        Else
          MsgBox "Le Niveau Ne Doit Pas Etre Null !"
          .ComboNiveau.SetFocus
        End If
      Else
        MsgBox "Le Numéro Classe Ne Doit Pas Etre Null !"
        .ComboClasse.SetFocus
      End If
    Else
     MsgBox "Aucun Etudiant Enrégistrer Avec Ce Matricule !"
     .ComboNumMat.SetFocus
    End If
  Else
    MsgBox "Le Numéro D'Inscription Ne Doit Etre Null !"
    .TxtNumIns.SetFocus
  End If
End With
End Function

'######################### FUNCTIONS GESTIONS ####################################

Function OpenAddIns()
'Prepare La frmInscription Pour Recevoir Des Nouvelles Informations
Options = "Ajouter"
  With FrmInscription
    .TxtNumIns.ToolTipText = "Automatiquement Genérer !"
    .TxtNumIns.Text = GenNewINS
    .OptActiv.Value = True
    .DTPickerDeb.Value = DateSerial(Year(Date), 10, 1)
    .DTPickerFin.Value = DateSerial(Year(Date) + 1, 6, 1)
    .TxtPayIns.Text = "00 Dh"
    .ComboTypeReg.ListIndex = 1
    .TxtPayTranch.Text = "00 Dh"
    .LabelPay_Tranche.Caption = "Payement Mensuel :"
  End With
  FrmInscription.Show 1
End Function

Function OpenModIns(Optional Inscription As String = "")
'Permet D'Ouvrir La frmInscription Avec Des Informations Existants
Options = "Modifier"
'------------------------
If Inscription = "" Then
Dim INS
Dim y
DebutOps:
   INS = Trim(InputBox("Entrez Le N° d'Inscription de L'Etudiant : "))
   If Not INS = "" Then
        If Not InsExist(CStr(INS)) Then
            y = MsgBox("N° D'inscription Non Trouver ! !", vbCritical + vbRetryCancel)
            If y = vbRetry Then
              GoTo DebutOps
            End If
        Else
            Inscription = INS
        End If
   Else
     Exit Function
   End If
End If
'------------------------
  With FrmInscription
    .Caption = "INSCRIPTION DE " & UCase(GetMyName(GetMaMatricule(Inscription)))
    .TxtNumIns.Enabled = False
    .ComboNumMat.Enabled = False
    .DTPickerDeb.Enabled = False
    .DTPickerFin.Enabled = False
    .TxtPayIns.Enabled = False
    .ComboTypeReg.Enabled = False
    .TxtPayTranch.Enabled = False
  End With
  GetInsInfos (Inscription)
  FrmInscription.Show 1
End Function

Function OpenSupIns(Optional Inscription As String)
'Permet D'0uvrir Un InputBox Puis Supprimer Les Informations Du Numéro D'Inscription Saisie
Dim y
  If Inscription = "" Then
DebutSup:
   Inscription = InputBox("Entrer Le Numéro D'Inscription A Supprimer : ")
  End If
   If Not Trim(Inscription) = "" Then
    If InsExist(CStr(Inscription)) Then
       y = MsgBox("Supprimer Les Informations De Cette Inscription ?  " & vbNewLine & "       " & Inscription, vbYesNo + vbQuestion)
       If y = vbYes Then
         If SupprimerInscription(CStr(Inscription)) Then
           MsgBox "Suppression réussie", vbOKOnly + vbInformation
         Else
           MsgBox "Desolé, Vous n'avez pas le privilège de supprimer une Inscription", vbOKOnly + vbInformation
         End If
       End If
    Else
      y = MsgBox("Inscription Inéxistant : " & Inscription, vbCritical + vbRetryCancel)
      If y = vbRetry Then
        GoTo DebutSup
      End If
    End If
   End If
End Function

'##################### MIS A JOUR COMPLETE BASE DE DONNEE #############################

Function EnregInsInfos(Inscription As String)
'Permet D'Enregistrer Une Nouvelle ou de Modifier Une Inscription
With FrmInscription
  Dim rs As New ADODB.Recordset
  SQL = "SELECT * FROM INSCRIPTIONS WHERE N°_Inscription = '" & Inscription & "'"
  rs.Open SQL, CN, adOpenKeyset, adLockOptimistic, adCmdText
  If rs.EOF Then
  'Si L'Inscription N'Existait Pas L'Ajouter !
    rs.AddNew
    rs![N°_Inscription] = Inscription
    
    '------- Frame 3 ----------------
    rs![Payment_Ins] = CStr(Val(.TxtPayIns.Text))
    rs![Type_Reglement] = .ComboTypeReg
    rs![Payment_Par_Tranche] = CStr(Val(.TxtPayTranch.Text))
    rs![Date_Inscription] = Date
    AddEvent "Nouvelle Inscription : " & .ComboNumMat & " ID: " & Inscription
   End If
  
   '--------- Frame 1 ---------------
   rs![N°_Matricule_Etudiant] = GetMaMat(.ComboNumMat)
   rs![Classe] = Trim(.ComboClasse)
   If .OptActiv.Value = True Then
          rs![Statut] = "Active"
    ElseIf .OptSuspendu.Value = True Then
        rs![Statut] = "Suspendue"
        rs![DetailSus] = .TxtDetailSus.Text
    ElseIf .OptExp.Value = True Then
        rs![Statut] = "Expiré"
   End If
   
   '-------- Frame 2 -----------------
   EnregClasseInfos (FrmInscription.ComboClasse)
   
   rs![Ins_Debut] = .DTPickerDeb.Value
   rs![Ins_Fin] = .DTPickerFin.Value
   rs![Annnee_Scolaire] = Year(.DTPickerDeb) & "/" & Year(.DTPickerFin)
   
  rs.Update
  rs.Close
  Set rs = Nothing
End With
'-------- Creer Ma Dette------------------------
CreerMaDette
'-------- Ajouter Evenement --------------------
AddEvent "Inscription Modifié : " & GetMyName(GetMaMatricule(Inscription)) & " ID: " & Inscription
End Function

Function GetInsInfos(Inscription As String) As Boolean
'Permet de Reprendre Les Informations Enrégistrées D'Une Inscription
  GetInsInfos = False
  UpdateMonInsStatut (GetMaMatricule(Inscription))
  Dim rs As New ADODB.Recordset
  SQL = "SELECT * FROM INSCRIPTIONS WHERE N°_Inscription = '" & Inscription & "'"
  rs.Open SQL, CN, adOpenKeyset, adLockOptimistic, adCmdText
  With FrmInscription
  If Not rs.EOF Then
    'Si L'Inscription Exist RePrendre Les Informations !
    
    '--------- Frame 1 ---------------
    .TxtNumIns.Text = rs![N°_Inscription]
    .ComboNumMat = GetMyName(rs![N°_Matricule_Etudiant])
    .ComboClasse = rs![Classe]
     If rs![Statut] = "Active" Then
       .OptActiv.Value = True
     ElseIf rs![Statut] = "Suspendue" Then
       .OptSuspendu.Value = True
       .TxtDetailSus.Text = rs![DetailSus] & " "
     ElseIf rs![Statut] = "Expiré" Then
       .OptExp.Value = True
     End If
     
     '--------- Frame 2 -----------
     GetClasseInfos (FrmInscription.ComboClasse)
     
     .DTPickerDeb.Value = CDate(rs![Ins_Debut])
     .DTPickerFin.Value = CDate(rs![Ins_Fin])
    
     '--------- Frame 3 ------------
     .TxtPayIns.Text = rs![Payment_Ins] & ""
     
     If rs![Type_Reglement] = "Unique" Then
          .ComboTypeReg.ListIndex = 0
        ElseIf rs![Type_Reglement] = "Mensuel" Then
          .ComboTypeReg.ListIndex = 1
        ElseIf rs![Type_Reglement] = "Trimestriel" Then
          .ComboTypeReg.ListIndex = 2
     End If
     
     .TxtPayTranch.Text = rs![Payment_Par_Tranche] & ""
     
     GetInsInfos = True
    End If
  rs.Close
  End With
End Function

Function SupprimerInscription(Inscription As String) As Boolean
'Supprimer L'Inscription De L'Etudiant
If Not G_Profil_Type = "ADMIN" Then
  SupprimerInscription = False
  AddEvent ("Suppression Inscription Bloqué : " & GetMyName(GetMaMatricule(Inscription)) & " ID: " & Inscription)
  Exit Function
End If
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM INSCRIPTIONS WHERE N°_Inscription = '" & UCase(Inscription) & "' "
    rs.Open SQL, CN, adOpenKeyset
      If Not rs.EOF Then
        CN.Execute ("DELETE FROM INSCRIPTIONS WHERE N°_Inscription = '" & UCase(Inscription) & "'")
      End If
    rs.Close
    Set rs = Nothing
AddEvent ("Inscription Supprimé : " & GetMyName(GetMaMatricule(Inscription)) & " ID: " & Inscription)
SupprimerInscription = True
End Function

Function EnregClasseInfos(Classe As String)
'Enregistrer ou Modifier Les Informations D'Une Classe
  Dim rs As New ADODB.Recordset
  SQL = "SELECT * FROM CLASSES WHERE Classe = '" & Classe & "'"
  rs.Open SQL, CN, adOpenKeyset, adLockOptimistic, adCmdText
  With FrmInscription
  If rs.EOF Then
    'Si La Classe N'Exist Pas L'Ajouter !
    rs.AddNew
    rs![Classe] = Trim(.ComboClasse)
    AddEvent "Nouvelle Classe " & Classe
  End If
  '--------- Frame 2 ---------------
     rs![Classe] = Trim(.ComboClasse)
     rs![Niveau_Filiere] = .ComboNiveau
     rs![Nom_Filiere] = .ComboFiliere
     rs![Annee_Classe] = .ComboAnnee
     
    rs.Update
    rs.Close
  End With
End Function

Function GetClasseInfos(Classe As String) As Boolean
'Permet De Reprendre les Informations D'Une Classe de La Base de Donnee
GetClasseInfos = False
  Dim RSClasse As New ADODB.Recordset
  SQL = "SELECT * FROM CLASSES WHERE Classe = '" & Classe & "'"
  RSClasse.Open SQL, CN, adOpenKeyset, adLockOptimistic, adCmdText
  With FrmInscription
  If Not RSClasse.EOF Then
    'Si La Classe Exist RePrendre Les Informations !
   .ComboClasse = RSClasse![Classe]
   .ComboNiveau = RSClasse![Niveau_Filiere]
   .ComboFiliere = RSClasse![Nom_Filiere]
   .ComboAnnee = RSClasse![Annee_Classe]
    GetClasseInfos = True
  End If
  RSClasse.Close
  End With
End Function

'################################################################################

Function GenNewINS() As String
'Generez Un Nouveau Numéro D'Inscription
Dim i As Integer
Dim Random As Integer
Dim NewGen As String
rep:
    For i = 1 To 3
        Randomize
        Random = CStr(CInt(89 * Rnd) + 10)
        NewGen = NewGen & CStr(Random)
    Next i
   GenNewINS = "I" & "-" & UCase(NewGen)
   If InsExist(CStr(GenNewINS)) Then
      GoTo rep
   End If
End Function

'##################### HALIDOU CISSE ##################
