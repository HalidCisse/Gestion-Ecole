Attribute VB_Name = "DBModule"
Option Explicit
Global Page As Integer
Global CN As New ADODB.Connection
Global SQL As String


Function OpenCN(DB_Name As String) As Boolean
On Error Resume Next
Set CN = New ADODB.Connection
'CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fdb
CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_Name & ";Persist Security Info=False;Jet OLEDB:Database Password= " & "halid" & ""
CN.Open
    If CN.State = adStateOpen Then
       OpenCN = True
    Else
       OpenCN = False
    End If
End Function

Function ConnectDB()
    Dim DB As String
      DB = App.Path
    If Not Right(DB, 1) = "\" Then
      DB = DB & "\"
    End If
      DB = DB & "EcoleDB.mdb"
    If Not OpenCN(DB) Then
      MsgBox "Impossible d'établir une connexion avec la base de donnée ", vbCritical
      End
    Else
      'MsgBox "Connexion - Gestion Inscritions Etudiants - Effectuée ", vbInformation
    End If
End Function

'------------------------########################-------------------------

Function AddEvent(Description As String)
'Ajouter Un Evenement Dans L'Historique
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM EVENTS"
    rs.Open SQL, CN, adOpenDynamic, adLockOptimistic, adCmdText
     
    rs.AddNew
        rs![Email_Profile] = G_Profil_Login
        rs![Description] = Description
        rs![EventTime] = Time
        rs![EventDate] = Date
    rs.Update
    
    rs.Close
    Set rs = Nothing
End Function

Function NombreEnreg(TABLE As String, CHAMPS As String, Statut As String) As Long
Dim rs As New ADODB.Recordset
    If CHAMPS = "SQL" Then
      SQL = Statut
    ElseIf Statut = "Tous" Then
      SQL = "SELECT COUNT(*) AS Nombre FROM " & TABLE
    Else
      SQL = "SELECT COUNT(*) AS Nombre FROM " & TABLE & " WHERE " & CHAMPS & " = '" & Statut & "'"
    End If
rs.Open SQL, CN
        NombreEnreg = CLng(Val(rs(0) & ""))
rs.Close
Set rs = Nothing
End Function


'###################### MSHFGRID ###############################

Function RemplirGrid(Form As Form, Grid As MSHFlexGrid, SQL As String)
     Dim rs As New ADODB.Recordset
     rs.Open SQL, CN, adOpenKeyset
     Set Grid.DataSource = rs
     rs.Close
     Set rs = Nothing
End Function

'#################### REMPLISSAGE ComboBox ##########################################

Function RemplirEtudiantAvecDette(Combo As ComboBox)
'Remplit Le Combo Avec Les Etudiant Endetter
Combo.Clear
    Dim rs As New ADODB.Recordset
    SQL = "SELECT DISTINCT N°_Matricule_Etudiant FROM INSCRIPTIONS"
    rs.Open SQL, CN
    '-------------------
    Do While Not rs.EOF
      If GetMaDette(rs![N°_Matricule_Etudiant]) > 0 Then
        Combo.AddItem (GetMyName(rs![N°_Matricule_Etudiant]))
      End If
    rs.MoveNext
    Loop
    '--------------------
    rs.Close
    Set rs = Nothing
End Function

Function RemplirMesDette(Matricule As String, Combo As ComboBox)
'Remplit Le Combo Avec Les Designation Des Dettes De L'Etudiant
Combo.Clear
    Dim rs As New ADODB.Recordset
    SQL = "SELECT Designation FROM DETTES WHERE Matricule = '" & Matricule & "' ORDER BY CDate(Date_Payment) ASC"
    rs.Open SQL, CN
    '-----------------
    Do While Not rs.EOF
      If Not DettePayer(Matricule, rs![Designation]) Then
        Combo.AddItem (rs![Designation])
      End If
    rs.MoveNext
    Loop
    '-----------------
    rs.Close
    Set rs = Nothing
End Function

Function RemplirComboFrmInscription()
  With FrmInscription
    CleanClasse
    RemplirComboEtudiantNonInscrit
    Remplir .ComboClasse, "CLASSES", "Classe"
    
    Remplir .ComboNiveau, "CLASSES", "Niveau_Filiere"
    Remplir .ComboFiliere, "CLASSES", "Nom_Filiere"
    Remplir .ComboAnnee, "CLASSES", "Annee_Classe"
  End With
End Function

Function RemplirCombosFrmEtudiant()
  Remplir FrmEtudiant.ComboNat, "PAYS", "PAYS"
  Remplir FrmEtudiant.ComboLieuNaiss, "VILLES", "VILLES"
  Remplir FrmEtudiant.ComboNEtude, "ETUDIANTS", "Niveau_Etude"
End Function

Function RemplirComboEtudiantNonInscrit()
'Return Les Etudiants Non Inscrit
  With FrmInscription
    .ComboNumMat.Clear
    Dim rs As New ADODB.Recordset
    SQL = "SELECT Matricule FROM ETUDIANTS"
    rs.Open SQL, CN, adOpenKeyset
      Do While Not rs.EOF
        If Not EstInscrit(rs![Matricule]) Then
          If Etudiant_Regulier(rs![Matricule]) Then
            .ComboNumMat.AddItem (GetMyName(rs![Matricule]))
          End If
        End If
        rs.MoveNext
      Loop
    rs.Close
    Set rs = Nothing
  End With
End Function

Function Remplir(Combo As ComboBox, TABLE As String, Optional NomChamps As String = "", Optional NumChamps As Integer = 0)
'Remplit Une Grid Avec des Données
Combo.Clear
    Dim rs As New ADODB.Recordset
    If Not NomChamps = "" Then
      SQL = "SELECT DISTINCT " & NomChamps & " FROM " & TABLE & " ORDER BY " & NomChamps & " ASC "
      NumChamps = 0
    Else
      SQL = "SELECT DISTINCT * FROM " & TABLE & ""
    End If
    rs.Open SQL, CN, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not rs.EOF
        Combo.AddItem (rs.Fields(NumChamps))
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Function



'##################### HALIDOU CISSE ##################
