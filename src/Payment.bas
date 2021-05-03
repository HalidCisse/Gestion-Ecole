Attribute VB_Name = "Payment"
Option Explicit


Function ChampsPayOK() As Boolean
Dim MAT
ChampsPayOK = False
'--------------
With FrmPayment
  MAT = GetMaMat(.ComboNumMat)
  If Not MatExist(CStr(MAT)) Then
     MsgBox "Etudiant Non Trouver !", vbInformation
    .ComboNumMat.SetFocus
  ElseIf Not MaDetteExist(CStr(MAT), .ComboMotifPay) Then
    MsgBox "Choisir de La Liste !"
    .ComboMotifPay.SetFocus
  ElseIf Not .ComboMotifPay = .ComboMotifPay.List(0) Then
    MsgBox "Vous Devez Payer " & .ComboMotifPay.List(0) & " d'Abord"
  ElseIf Trim(.TxtNumPay.Text) = "" Then
    MsgBox "Le Numéro De Payment Ne Doit Pas Etre Null !"
    .TxtNumPay.SetFocus
  ElseIf Trim(.ComboNumMat) = "" Then
         MsgBox "Le Numéro De Matricule Ne Doit Pas Etre Null !"
        .ComboNumMat.SetFocus
  ElseIf Trim(.ComboMotifPay) = "" Then
    MsgBox "Le Motif Ne Doit Pas Etre Null !"
    .ComboMotifPay.SetFocus
  ElseIf Trim(.TxtPayBy.Text) = "" Then
    MsgBox "Le Nom Du Payeur Ne Doit Pas Etre Null !"
    .TxtPayBy.SetFocus
  Else
   ChampsPayOK = True
  End If
End With
'------------
End Function

'#################### GESTION  ###############################

Function OpenSupPayment(Optional N°Payment As String = "")
'Ouvrir Un Dialogue Pour Supprimer Les Infos D'Un Payment
Dim N As String
Dim y
N = N°Payment
   If N = "" Then
DebutSup:
     N = Trim(InputBox("Entrez Le N° De Payment A Supprimer "))
   End If
   If N <> "" Then
    If PayExist(CStr(N)) Then
       y = MsgBox("Supprimer Les Informations De Ce Payment: " & N, vbYesNo + vbQuestion)
       If y = vbYes Then
         If SupprimerPayment(CStr(N)) Then
           MsgBox "Suppression réussie", vbOKOnly + vbInformation
         Else
           MsgBox "Vous n'avez pas le privilège de supprimer un payment !"
         End If
       End If
    Else
      y = MsgBox("N° De Payement Introuvable !", vbCritical + vbRetryCancel)
      If y = vbRetry Then
        GoTo DebutSup
      End If
    End If
   End If
End Function

Function OpenPayAvance()
'Payment En Avance
Dim MAT
deb:
  MAT = InputBox("Entrer Le Matricule de L'Etudiant :")
  If MatExist(CStr(Trim(MAT))) Then
    With FrmPayment
      .ComboNumMat = GetMyName(CStr(MAT))
      .ComboNumMat.Locked = True
      .Show 1
    End With
  ElseIf Not Trim(MAT) = "" Then
    MsgBox "Etudiant Non Trouvé !", vbInformation
    GoTo deb
  End If
End Function

'##################### MIS A JOUR COMPLETE BASE DE DONNEE ############

Function EnregPayInfos(N°_Payment As String)
  'Enregistrer Les Informations D'Un Payment
With FrmPayment
Dim N_Dette
Dim MAT
MAT = GetMaMat(.ComboNumMat)
N_Dette = .ComboMotifPay & MAT

  If Not DettePayer(CStr(MAT), CStr(N_Dette)) Then
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM PAYMENTS WHERE N°_Payment = '" & N°_Payment & "'"
    rs.Open SQL, CN, adOpenKeyset, adLockOptimistic, adCmdText
    
    If rs.EOF Then
      'Si Le Payment N'Exist Pas L'Ajouter !
      rs.AddNew
      rs![ID_Dette] = N_Dette
      rs![N°_Payment] = Trim(UCase(.TxtNumPay))
      rs![N°_Matricule] = MAT
      rs![Designation] = .ComboMotifPay
      rs![Payer_Par] = .TxtPayBy.Text
      rs![Somme_Payer] = GetDette_Som(CStr(N_Dette))
      rs![Moyen_De_Payment] = .ComboMoyPay
      rs![Email_Profile] = G_Profil_Login
      rs![Date_Payment] = Date
      rs![Time_Payment] = Time
    rs.Update
    rs.Close
    AddEvent "Payement " & .ComboMotifPay & " Effectué Par " & .ComboNumMat
    End If
  End If
End With
End Function

Function SupprimerPayment(Payment_ID As String) As Boolean
'Supprimer Le Payment De L'Etudiant
Dim PayD
Dim MAT
PayD = GetPay_Des(Payment_ID) & " N° " & Payment_ID
MAT = GetPayMat(Payment_ID)
'----------------------------------
If Not G_Profil_Type = "ADMIN" Then
  SupprimerPayment = False
  AddEvent ("Suppression De Payement Bloqué : " & PayD)
  Exit Function
End If
'---------------------------------
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM PAYMENTS WHERE N°_Payment = '" & Payment_ID & "' "
    rs.Open SQL, CN, adOpenKeyset
      If Not rs.EOF Then
        CN.Execute ("DELETE FROM PAYMENTS WHERE N°_Payment = '" & Payment_ID & "'")
      End If
    rs.Close
    Set rs = Nothing
    '---------------
SupprimerPayment = True
AddEvent ("Payement de " & GetMyName(CStr(MAT)) & " Supprimé : " & PayD)
End Function

Function AddDette(Matricule As String, Designation As String, Somme_A_Payer As String, Date_Payment As String)
'Ajouter Une Dette A Payer Par L'Etudiant

If CLng(Somme_A_Payer) = 0 Then
  Exit Function
End If

  Dim rs As New ADODB.Recordset
  SQL = "SELECT * FROM DETTES WHERE ID_Dette = '" & Designation & Matricule & "'"
  rs.Open SQL, CN, adOpenKeyset, adLockOptimistic, adCmdText
    '-----------------
    If rs.EOF Then
      rs.AddNew
    End If
    '-----------------
    rs![ID_Dette] = Designation & Matricule
    rs![Designation] = Designation
    rs![Matricule] = Matricule
    rs![Somme_A_Payer] = Somme_A_Payer
    rs![Date_Payment] = Date_Payment
    '-----------------
  rs.Update
  rs.Close
  Set rs = Nothing
  
End Function

Function CreerMaDette()
'Génerer Les Dettes Scolaires De L'Etudiant
'---------------------
Dim DatePayment As Date
Dim Reglement As Integer
Dim MAT
DatePayment = FrmInscription.DTPickerDeb.Value
'---------------------
 With FrmInscription
    MAT = GetMaMat(.ComboNumMat)
    '--- Creer L'Inscription
     AddDette CStr(MAT), "Inscription Annuelle " & Year(.DTPickerDeb.Value) & "/" & Year(.DTPickerFin.Value), CStr(Val(.TxtPayIns.Text)), DateSerial(Year(DatePayment), Month(DatePayment) - 1, Day(DatePayment) + 27)
    '---------------------------------------------------------
    If .ComboTypeReg = "Unique" Then
    'Creer 1 Dette Unique
      AddDette CStr(MAT), "Payement " & Year(.DTPickerDeb.Value) & "/" & Year(.DTPickerFin.Value) & " - Annuel", CStr(Val(.TxtPayTranch.Text)), CStr(DatePayment)
    ElseIf .ComboTypeReg = "Mensuel" Then
    'Creer Des Dettes Mensuels
      For Reglement = 0 To DateDiff("m", .DTPickerDeb.Value, .DTPickerFin.Value)
        AddDette CStr(MAT), Format(DatePayment, "mmm") & " " & Year(DatePayment) & " - Mensuel", CStr(Val(.TxtPayTranch.Text)), CStr(DatePayment)
        DatePayment = DateSerial(Year(DateAdd("m", 1, DatePayment)), Month(DateAdd("m", 1, DatePayment)), 1)
      Next Reglement
    ElseIf .ComboTypeReg = "Trimestriel" Then
    'Creer Des Dettes Trimestriel
      For Reglement = 1 To (DateDiff("m", .DTPickerDeb.Value, .DTPickerFin.Value)) / 3
        AddDette CStr(MAT), Format(DatePayment, "mmm") & " " & Year(DatePayment) & " - Trimestriel", CStr(Val(.TxtPayTranch.Text)), CStr(DatePayment)
        DatePayment = DateSerial(Year(DateAdd("m", 3, DatePayment)), Month(DateAdd("m", 3, DatePayment)), 1)
      Next Reglement
    End If
    '-------------------------------------------------------
 End With
End Function

Function GenNewPayID() As String
'Generez Un Nouveau Numéro De Payment
Dim i As Integer
Dim Random As Integer
Dim NewGen As String
rep:
    For i = 1 To 3
        Randomize
        Random = CStr(CInt(89 * Rnd) + 10)
        NewGen = NewGen & CStr(Random)
    Next i
   GenNewPayID = "P" & "-" & UCase(NewGen)
   If PayExist(CStr(GenNewPayID)) Then
      GoTo rep
   End If
End Function

'##################### HALIDOU CISSE ##################
