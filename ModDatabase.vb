Option Strict Off
Option Explicit On
Imports System.Data.OleDb

Module ModDatabase
    Sub ConnectToDatabase(ByRef strPath As String)
        'Connect to database
        'gcConn = New ADODB.Connection
        gcConn = New OleDb.OleDbConnection
        gcConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & strPath & ";" & "Persist Security Info=False"
        gcConn.Open()
    End Sub

    Public Function rsGetCalendrier(ByVal iTeamNo As Short) As OleDbDataReader
        Dim strSQL As String


        strSQL = "SELECT * FROM qry_intranet_calendrier "
        If iTeamNo > -1 Then
            strSQL = strSQL & "WHERE tbl�quipes.No�quipe = " & Str(iTeamNo) & " OR tbl�quipes_1.No�quipe = " & Str(iTeamNo) & " "
        End If
        strSQL = strSQL & "ORDER BY NoPartie"
        Dim command As New OleDbCommand(strSQL, gcConn)
        Dim reader As OleDbDataReader = command.ExecuteReader()
        ' rs.ExecuteReader(strSQL)

        '' rs = gcConn.Execute(strSQL)
        'While reader.Read()
        '    Console.WriteLine(reader(0).ToString())
        '    Console.WriteLine(reader(1).ToString())
        '    Console.WriteLine(reader(11).ToString())
        'End While


        rsGetCalendrier = reader

    End Function

    Public Function rsGetTeamPtsPerGame(ByRef iGameNo As Short, ByRef iTeamNo As Short) As OleDbDataReader
        Dim strSQL As String
        strSQL = "SELECT tblR�pondants.NoPartie, tblR�pondants.No�quipe,  Sum(IIf(tblR�pondants!PtsAlternatifs=0,tblS�ries!Points,tblS�ries!PtsAlternatifs)) AS SumOfPoints " & "FROM tblR�pondants INNER JOIN tblS�ries ON (tblR�pondants.NoS�rie = tblS�ries.NoS�rie) AND (tblR�pondants.NoQuestion = tblS�ries.NoQuestion) " & "GROUP BY tblR�pondants.NoPartie, tblR�pondants.No�quipe " & "HAVING tblR�pondants.NoPartie = " & Str(iGameNo) & "AND tblR�pondants.No�quipe = " & Str(iTeamNo) & " " & "ORDER BY tblR�pondants.NoPartie, tblR�pondants.No�quipe"
        'Dim rs As ADODB.Recordset
        Dim command As New OleDbCommand(strSQL, gcConn)
        Dim reader As OleDbDataReader = command.ExecuteReader()

        rsGetTeamPtsPerGame = reader

    End Function

    'Public Function rsGetPlayerPtsForAGame(ByRef iGameNo As Short, ByRef iTeamNo As Short) As OleDbDataReader
    '    Dim strSQL As String
    '    'Essayer plut�t avec qryPointsJoueurParParties - version Covid
    '    strSQL = "SELECT tblR�pondants.NoPartie, IIf([tblJoueurs.Pr�nomJoueur] Is Null,[tblJoueurs.NomJoueur],[tblJoueurs.Pr�nomJoueur]+' '+[tblJoueurs.NomJoueur]) AS Nom, Sum(IIf(tblR�pondants!PtsAlternatifs=0,tblS�ries!Points,tblS�ries!PtsAlternatifs)) AS SumOfPoints " & "FROM tblJoueurs INNER JOIN (tblR�pondants INNER JOIN tblS�ries ON (tblR�pondants.NoQuestion = tblS�ries.NoQuestion) AND (tblR�pondants.NoS�rie = tblS�ries.NoS�rie)) ON (tblJoueurs.NoJoueur = tblR�pondants.NoJoueur) AND (tblJoueurs.No�quipe = tblR�pondants.No�quipe) " & "GROUP BY tblR�pondants.NoPartie, IIf([tblJoueurs.Pr�nomJoueur] Is Null,[tblJoueurs.NomJoueur],[tblJoueurs.Pr�nomJoueur]+' '+[tblJoueurs.NomJoueur]), tblJoueurs.No�quipe " & "HAVING (((tblR�pondants.NoPartie)=" & Str(iGameNo) & ") AND ((tblJoueurs.No�quipe)=" & Str(iTeamNo) & ")) " & "ORDER BY Sum(IIf(tblR�pondants!PtsAlternatifs=0,tblS�ries!Points,tblS�ries!PtsAlternatifs)) DESC"


    '    Dim command As New OleDbCommand(strSQL, gcConn)
    '    Dim reader As OleDbDataReader = command.ExecuteReader()

    '    rsGetPlayerPtsForAGame = reader

    'End Function


    Public Function rsGetPlayerPtsForAGame(ByRef iGameNo As Short) As OleDbDataReader
        Dim strSQL As String
        'Essayer plut�t avec qryPointsJoueurParParties - version Covid
        strSQL = "SELECT NoPartie, Nom�quipe, joueur, Points " & "FROM qryPointsParJoueursParties  " & " Where NoPartie=" & Str(iGameNo) & " ORDER BY Points desc, nom�quipe, joueur"

        Dim command As New OleDbCommand(strSQL, gcConn)
        Dim reader As OleDbDataReader = command.ExecuteReader()

        rsGetPlayerPtsForAGame = reader

    End Function

    Public Function rsGetPossiblePtsForAGame(ByRef iGameNo As Short) As ADODB.Recordset
        Dim strSQL As String
        Dim rs As ADODB.Recordset

        strSQL = "SELECT tblR�pondants.NoPartie, Sum(tblS�ries.Points) AS PtsPossJoueurs " & "FROM tblR�pondants INNER JOIN tblS�ries ON (tblR�pondants.NoQuestion = tblS�ries.NoQuestion) AND (tblR�pondants.NoS�rie = tblS�ries.NoS�rie) " & "Where (((tblR�pondants.NoJoueur) <> 99)) " & "GROUP BY tblR�pondants.NoPartie " & "HAVING (((tblR�pondants.NoPartie)=" & Str(iGameNo) & "))"

        rs = New ADODB.Recordset
        '  ' rs = gcConn.Execute(strSQL)
        rsGetPossiblePtsForAGame = rs

    End Function

    Public Function rsGetTeams(ByRef blnAllTeams As Boolean) As OleDbDataReader
        Dim strSQL As String

        If blnAllTeams Then
            strSQL = "SELECT tbl�quipes.No�quipe, tbl�quipes.Nom�quipe From tbl�quipes order by tbl�quipes.No�quipe ASC"
        Else
            strSQL = "SELECT TOP 1 tblParties.No�quipeA, tbl�quipes.Nom�quipe, tblParties.No�quipeB, tbl�quipes_1.Nom�quipe " & "FROM tbl�quipes AS tbl�quipes_1 INNER JOIN (tbl�quipes INNER JOIN tblParties ON tbl�quipes.No�quipe = tblParties.No�quipeA) ON tbl�quipes_1.No�quipe = tblParties.No�quipeB " & "Where (((tblParties.No�quipeA) Is Not Null)) " & "ORDER BY tblParties.NoPartie DESC; "
        End If

        Dim command As New OleDbCommand(strSQL, gcConn)

        Dim reader As OleDbDataReader = command.ExecuteReader()

        rsGetTeams = reader

    End Function

    Public Function rsGetClassement() As OleDbDataReader
        Dim strSQL As String

        strSQL = "SELECT * FROM qryRapClassement_final"

        Dim command As New OleDbCommand(strSQL, gcConn)

        Dim reader As OleDbDataReader = command.ExecuteReader()

        rsGetClassement = reader

    End Function


    Public Function rsGetCompteurs(ByVal iTeamNo As Short) As OleDb.OleDbDataReader
        Dim strSQL As String
        'Dim rs As OleDb.OleDbDataReader


        If iTeamNo > -1 Then
            strSQL = "SELECT * " & "FROM qryStatsPtsTotJoueurs_Tous_Final " & "WHERE qryStatsPtsTotJoueurs_Tous_Final.No�quipe = " & Str(iTeamNo)

        Else
            strSQL = "SELECT tbl�quipes.No�quipe, qryStatPtsTotJoueurs.* " & "FROM qryStatPtsTotJoueurs " & "INNER JOIN tbl�quipes ON qryStatPtsTotJoueurs.Nom�quipe = tbl�quipes.Nom�quipe "
        End If
        strSQL = strSQL & " ORDER BY ptsParPartie DESC, PtsTotJoueurs DESC, NomJoueur"
        Dim command As New OleDbCommand(strSQL, gcConn)

        Dim reader As OleDbDataReader = command.ExecuteReader()
        'rs = New ADODB.Recordset
        ' rs = gcConn.Execute(strSQL)
        rsGetCompteurs = reader

    End Function

    Public Function rsGetCompteurs_Old(ByVal iTeamNo As Short) As ADODB.Recordset
        Dim strSQL As String
        Dim rs As ADODB.Recordset

        If iTeamNo > -1 Then
            strSQL = "SELECT tbl�quipes.No�quipe, qryStatPtsTotJoueurs_tous.* " & "FROM qryStatPtsTotJoueurs_tous " & "INNER JOIN tbl�quipes ON qryStatPtsTotJoueurs_tous.Nom�quipe = tbl�quipes.Nom�quipe " & "WHERE tbl�quipes.No�quipe = " & Str(iTeamNo)
        Else
            strSQL = "SELECT tbl�quipes.No�quipe, qryStatPtsTotJoueurs.* " & "FROM qryStatPtsTotJoueurs " & "INNER JOIN tbl�quipes ON qryStatPtsTotJoueurs.Nom�quipe = tbl�quipes.Nom�quipe "
        End If
        strSQL = strSQL & " ORDER BY qryStatPtsTotJoueurs_1.Pourcentage DESC"

        rs = New ADODB.Recordset
        ' rs = gcConn.Execute(strSQL)
        rsGetCompteurs_Old = rs

    End Function

    'Public Function rsGetSalles() As ADODB.Recordset
    'Dim strSQL As String
    'Dim rs As ADODB.Recordset
    '
    'strSQL = "SELECT distinct tblParties.Salle " & _
    ''         "FROM tblParties " & _
    ''         "WHERE tbParties.Salle <> 'JT 11 E'"
    '
    'Set rs = New ADODB.Recordset
    'Set ' rs = gcConn.Execute(strSQL)
    'Set rsGetSalles = rs
    '
    'End Function

    Public Function rsGetQuest(ByRef intQuestNo As Short) As ADODB.Recordset
        Dim strSQL As String
        Dim rs As ADODB.Recordset

        strSQL = "SELECT qryRapFeuilleDeMatch.* FROM qryRapFeuilleDeMatch "

        rs = New ADODB.Recordset
        ' rs = gcConn.Execute(strSQL)
        rsGetQuest = rs

    End Function
End Module