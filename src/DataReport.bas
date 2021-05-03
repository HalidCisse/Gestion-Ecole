Attribute VB_Name = "DataReport"
Option Explicit


Function Open_DR_Dos_Etud(Matricule As String)
'Ouvre Le DataReport Pour Imprimer Le Dossier D'un Etudiants

SyncMesInfos (CStr(Matricule))

Dim DR_SQL As String

DR_SQL = "SELECT TypeID, Numero_Identite, Date_Expire, Sex, Nom, Prenom, Date_Naissance, Lieu_Naissance, Nationalite, Nom_Pere, Nom_Mere, TEL,Email, Adresse, Nom_Tuteur, Prenom_Tuteur, TEL_Tuteur, Email_Tuteur, Adresse_Tuteur, Niveau_Etude, Matricule, N°_Inscription, Classe, Total_Payment, Date_Enregistrement, Statut FROM ETUDIANTS WHERE Matricule = '" & Matricule & "'"

'--------------------------------------------
  If DE.rsCMD_ETUD.State = adStateOpen = False Then
    DE.rsCMD_ETUD.Open
  End If
  DE.Commands("CMD_ETUD").CommandText = DR_SQL
  DE.Commands("CMD_ETUD").ActiveConnection = CN
  DE.Commands("CMD_ETUD").Execute
  DE.rsCMD_ETUD.Requery
'--------------------------------------------
'--------------------------------------------
  With DR_Dos_Etud
    .Caption = "DOSSIER DE REGISTRATION DE " & GetMyName(CStr(Matricule))
    .Sections("Section4").Controls("Header").Caption = .Caption
    .ReportWidth = 1000
    .LeftMargin = 800
    .RightMargin = 1440
    .Refresh
    .Show 1
  End With
'--------------------------------------------
If DE.rsCMD_ETUD.State = adStateOpen = True Then
    DE.rsCMD_ETUD.Close
End If
AddEvent "Impression Dossier de Registration de " & GetMyName(CStr(Matricule))
End Function

Function Open_DR_Dos_Ins(Inscription As String)
'Ouvre Le DataReport Pour Imprimer Le Dossier d'Inscription
SyncMesInfos (CStr(GetMaMatricule(Inscription)))
'--------------------
Dim DR_SQL
DR_SQL = "SELECT * FROM INSCRIPTIONS WHERE N°_Inscription = '" & Inscription & "'"
'-------------------
'--------------------------------------------
  If DE.rsCmd_INS.State = adStateOpen = False Then
    DE.rsCmd_INS.Open
  End If
  DE.Commands("Cmd_INS").CommandText = DR_SQL
  DE.Commands("Cmd_INS").ActiveConnection = CN
  DE.Commands("Cmd_INS").Execute
  DE.rsCmd_INS.Requery
'--------------------------------------------
'--------------------------------------------
  With DR_Dos_Ins
    .Caption = "DOSSIER D'INSCRIPTION DE " & GetMyName(GetMaMatricule((Inscription)))
    .Sections("Section4").Controls("Header").Caption = .Caption
    .ReportWidth = 1000
    .LeftMargin = 800
    .RightMargin = 1440
    .Refresh
    .Show 1
  End With
'-------------------------------------------
If DE.rsCmd_INS.State = adStateOpen = True Then
    DE.rsCmd_INS.Close
End If
AddEvent "Impression Dossier Inscription de " & GetMyName(GetMaMatricule(CStr(Inscription)))
End Function

Function Open_DR_Recu_Pay(PayID As String)
'Ouvre Le DataReport Pour Imprimer Les Payments
'--------------------
Dim DR_SQL
DR_SQL = "SELECT * FROM PAYMENTS WHERE N°_Payment = '" & PayID & "'"
'-------------------
'--------------------------------------------
  If DE.rsCMD_PAY.State = adStateOpen = False Then
    DE.rsCMD_PAY.Open
  End If
  DE.Commands("CMD_PAY").CommandText = DR_SQL
  DE.Commands("CMD_PAY").ActiveConnection = CN
  DE.Commands("CMD_PAY").Execute
  DE.rsCMD_PAY.Requery
'--------------------------------------------
'--------------------------------------------
  With DR_Recu_PAY
    .Caption = "RECUE DE PAYEMENT DE " & GetMyName(GetPayMat(PayID))
    .Sections("Section4").Controls("Header").Caption = .Caption
    .ReportWidth = 1000
    .LeftMargin = 800
    .RightMargin = 1440
    .Refresh
    .Show 1
  End With
'--------------------------------------------
If DE.rsCMD_PAY.State = adStateOpen = True Then
    DE.rsCMD_PAY.Close
End If
AddEvent "Impression Reçue de " & GetMyName(GetPayMat(PayID)) & " N°: " & PayID
End Function



'###################### DR IMPRIMER LISTE #####################################

Function Open_DR_Ins(Optional DR_SQL As String = "")
'Ouvre Le DataReport Pour Imprimer Les Inscription
'--------------------
If DR_SQL = "" Then
  DR_SQL = "SELECT Nom, Prenom, N°_Matricule_Etudiant, Classe, N°_inscription FROM ETUDIANT_INSCRIT ORDER BY Date_Inscription DESC"
End If
'-------------------
'--------------------------------------------
  If DE.rsCmd.State = adStateOpen = False Then
    DE.rsCmd.Open
  End If
  DE.Commands("Cmd").CommandText = DR_SQL
  DE.Commands("Cmd").ActiveConnection = CN
  DE.Commands("Cmd").Execute
  DE.rsCmd.Requery
'--------------------------------------------
'--------------------------------------------
  With DR_INS
    .ReportWidth = 1000
    .LeftMargin = 800
    .RightMargin = 1440
    .Refresh
    .Show 1
  End With
'-------------------------------------------
If DE.rsCmd.State = adStateOpen = True Then
    DE.rsCmd.Close
End If
AddEvent "Impression Liste Inscriptions"
End Function

Function Open_DR_Etud(Optional DR_SQL As String = "")
'Ouvre Le DataReport Pour Imprimer La Liste Des Etudiants
'--------------------
If DR_SQL = "" Then
  DR_SQL = "SELECT Prenom, Nom, Matricule, Statut, Niveau_Etude FROM ETUDIANTS ORDER BY Nom ASC"
End If
'-------------------
'--------------------------------------------
  If DE.rsCMD_ETUD.State = adStateOpen = False Then
    DE.rsCMD_ETUD.Open
  End If
  DE.Commands("CMD_ETUD").CommandText = DR_SQL
  DE.Commands("CMD_ETUD").ActiveConnection = CN
  DE.Commands("CMD_ETUD").Execute
  DE.rsCMD_ETUD.Requery
'--------------------------------------------
'--------------------------------------------
  With DR_ETUD
    .ReportWidth = 1000
    .LeftMargin = 800
    .RightMargin = 1440
    .Refresh
    .Show 1
  End With
'--------------------------------------------
If DE.rsCMD_ETUD.State = adStateOpen = True Then
    DE.rsCMD_ETUD.Close
End If
AddEvent "Impression Liste Etudiants"
End Function

Function Open_DR_Pay(Optional DR_SQL As String = "")
'Ouvre Le DataReport Pour Imprimer Les Payments
'--------------------
If DR_SQL = "" Then
  DR_SQL = "SELECT Payer_Par, N°_Payment, Designation, Somme_Payer, Date_Payment FROM PAYMENTS ORDER BY Date_Payment DESC"
End If
'-------------------
'--------------------------------------------
  If DE.rsCMD_PAY.State = adStateOpen = False Then
    DE.rsCMD_PAY.Open
  End If
  DE.Commands("CMD_PAY").CommandText = DR_SQL
  DE.Commands("CMD_PAY").ActiveConnection = CN
  DE.Commands("CMD_PAY").Execute
  DE.rsCMD_PAY.Requery
'--------------------------------------------
'--------------------------------------------
  With DR_PAY
    .ReportWidth = 1000
    .LeftMargin = 800
    .RightMargin = 1440
    .Refresh
    .Show 1
  End With
'--------------------------------------------
If DE.rsCMD_PAY.State = adStateOpen = True Then
    DE.rsCMD_PAY.Close
End If
AddEvent "Impression Liste Payements"
End Function

'##################### HALIDOU CISSE ##################
