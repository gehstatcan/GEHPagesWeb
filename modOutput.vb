Option Strict Off
Option Explicit On
Module modOutput
	Sub subOutputCalendrier(ByVal iTeamNo As Short, ByRef strTitle As String, ByRef strFileName As String)

		Dim file = My.Computer.FileSystem.OpenTextFileWriter(
		strFileName, False, Text.Encoding.UTF8)

		file.WriteLine("<HTML><HEAD><TITLE></TITLE></HEAD>")
		file.WriteLine("<BODY TOPMARGIN=10 BGPROPERTIES=""FIXED"" BGCOLOR = " & DOC_BACKGROUND_COLOR & "><div id=Outline>")
		file.WriteLine("<CENTER><FONT FACE=VERDANA COLOR = black SIZE=5><B>" & strTitle & "</FONT></CENTER>")
		file.WriteLine("<BR>")
		' file.WriteLine("JT = Jean-Talon, 11i�me, salle E")
		file.WriteLine("Saison 2021 de la maison.")
		file.WriteLine("<BR>")
		file.WriteLine("Utilisez le calendrier pour avoir le lien vers la r�union MS Teams ET vers votre bouton")
		file.WriteLine("<BR>")
		' file.WriteLine("RHC = Coats, 16i�me, salle 3 (Ou RHC 11)")
		'file.WriteLine("<BR>")
		file.WriteLine("<BR>")
		file.WriteLine(funGetCalendrier(iTeamNo))
		file.WriteLine("<BR>")
		file.WriteLine(funGetScript())
		file.WriteLine("</HTML>")
		file.Close()

	End Sub
	
	Sub subOutputClassement(ByRef strTitle As String, ByRef strFileName As String)

		Dim rsClassement As OleDb.OleDbDataReader
		Dim ColWidth(13) As Short
		Dim I As Short
		Dim strHREF As String

		ColWidth(1) = 30
		ColWidth(2) = 170
		ColWidth(3) = 50
		ColWidth(4) = 50
		ColWidth(5) = 50
		ColWidth(6) = 50
		ColWidth(7) = 50
		ColWidth(8) = 50
		ColWidth(9) = 50
		ColWidth(10) = 50
		ColWidth(11) = 70
		ColWidth(12) = 70
		ColWidth(13) = 70

		rsClassement = rsGetClassement()

		Dim file = My.Computer.FileSystem.OpenTextFileWriter(
		strFileName, False, Text.Encoding.UTF8)

		'd�but du document
		file.WriteLine("<!DOCTYPE HTML PUBLIC -//W3C//DTD HTML 4.0 Transitional//EN><HTML><HEAD><TITLE></TITLE>")
		file.WriteLine("<body TOPMARGIN=10 BGPROPERTIES=""FIXED"" BGCOLOR = " & DOC_BACKGROUND_COLOR & ">")
		file.WriteLine("<CENTER>")
		file.WriteLine("<FONT FACE=VERDANA COLOR = black SIZE=5><B>" & strTitle & "</FONT><BR><BR>")

		'ent�te du classement

		file.WriteLine("<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=3>")
		file.WriteLine("<TR align=center bgcolor = " & COL_HEADER_COLOR & ">")
		file.WriteLine("<TD WIDTH=" & Str(ColWidth(1)) & " ROWSPAN=2 VALIGN=MIDDLE><FONT FACE=VERDANA SIZE=2 COLOR=white><B>Pos </B></FONT></TD>")
		file.WriteLine("<TD WIDTH=" & Str(ColWidth(2)) & " ROWSPAN=2 VALIGN=middle ALIGN=LEFT><FONT FACE=VERDANA SIZE=2 COLOR=white><B>�quipes</B></FONT></TD>")
		file.WriteLine("<TD COLSPAN=8 VALIGN=middle><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>Parties</B></FONT></TD>")
		file.WriteLine("<TD COLSPAN=3 VALIGN=middle><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>Pointage</B></FONT></TD>")
		file.WriteLine("<TR BGCOLOR=darkcyan align=center>")
		file.WriteLine("<TD WIDTH=" & Str(ColWidth(3)) & " BGCOLOR=darkcyan><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>PJ</B></FONT></TD>")
		file.WriteLine("<TD WIDTH=" & Str(ColWidth(4)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>G</B></FONT></TD>")
		file.WriteLine("<TD WIDTH=" & Str(ColWidth(5)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>GP</B></FONT></TD>")
		file.WriteLine("<TD WIDTH=" & Str(ColWidth(6)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>PN</B></FONT></TD>")
		file.WriteLine("<TD WIDTH=" & Str(ColWidth(7)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>PP</B></FONT></TD>")
		file.WriteLine("<TD WIDTH=" & Str(ColWidth(8)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>P</B></FONT></TD>")
		file.WriteLine("<TD WIDTH=" & Str(ColWidth(9)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>Pts</B></FONT></TD>")
		file.WriteLine("<TD WIDTH=" & Str(ColWidth(10)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>%*</B></FONT></TD>")
		file.WriteLine("<TD WIDTH=" & Str(ColWidth(11)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>PP</B></FONT></TD>")
		file.WriteLine("<TD WIDTH=" & Str(ColWidth(12)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>PC</B></FONT></TD>")
		file.WriteLine("<TD WIDTH=" & Str(ColWidth(13)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>Diff</B></FONT></TD>")
		file.WriteLine("</TR></TABLE>")
		file.WriteLine("")

		file.WriteLine("<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=3>")

		While rsClassement.Read()


			System.Windows.Forms.Application.DoEvents()

			file.WriteLine("<TR ALIGN=CENTER BGCOLOR = white>")
			file.WriteLine("<TD WIDTH=30><FONT FACE=VERDANA SIZE=2 COLOR=Black>" & Trim(Str(I)) & "</FONT></TD>")
			strHREF = "<A HREF = """ & "equipe_" & Trim(Str(rsClassement(0).ToString)) & ".htm" & """  > " & rsClassement(1).ToString & "</A>"
			file.WriteLine("<TD WIDTH=170 ALIGN=LEFT><FONT FACE=VERDANA SIZE=2 COLOR=black>" & strHREF & "</FONT></TD>")
			file.WriteLine("<TD WIDTH=50><FONT FACE=VERDANA SIZE=2 COLOR=black>" & rsClassement(2).ToString & "</FONT></TD>")
			file.WriteLine("<TD WIDTH=50><FONT FACE=VERDANA SIZE=2 COLOR=black>" & rsClassement(3).ToString & "</FONT></TD>")
			file.WriteLine("<TD WIDTH=50><FONT FACE=VERDANA SIZE=2 COLOR=black>" & rsClassement(4).ToString & "</FONT></TD>")
			file.WriteLine("<TD WIDTH=50><FONT FACE=VERDANA SIZE=2 COLOR=black>" & rsClassement(5).ToString & "</FONT></TD>")
			file.WriteLine("<TD WIDTH=50><FONT FACE=VERDANA SIZE=2 COLOR=black>" & rsClassement(6).ToString & "</FONT></TD>")
			file.WriteLine("<TD WIDTH=50><FONT FACE=VERDANA SIZE=2 COLOR=black>" & rsClassement(7).ToString & "</FONT></TD>")
			file.WriteLine("<TD WIDTH=50><FONT FACE=VERDANA SIZE=2 COLOR=black>" & rsClassement(8).ToString & "</FONT></TD>")
			file.WriteLine("<TD WIDTH=50><FONT FACE=VERDANA SIZE=2 COLOR=black>" & rsClassement(9).ToString & "</FONT></TD>")
			file.WriteLine("<TD WIDTH=70><FONT FACE=VERDANA SIZE=2 COLOR=black>" & rsClassement(10).ToString & "</FONT></TD>")
			file.WriteLine("<TD WIDTH=70><FONT FACE=VERDANA SIZE=2 COLOR=black>" & rsClassement(11).ToString & "</FONT></TD>")
			file.WriteLine("<TD WIDTH=70><FONT FACE=VERDANA SIZE=2 COLOR=black>" & rsClassement(12).ToString & "</FONT></TD>")
			file.WriteLine("</TR>")

		End While
		file.WriteLine("</TABLE>")
		file.WriteLine("<TABLE WIDTH=850 BORDER=0>")
		file.WriteLine("<TR><TD><FONT FACE=VERDANA SIZE=2>Tri� par : 1-%, 2 -Dicoff. Ceci afin de tenir compte du nombre in�gale de parties jou�es.</TD></TR>")
		file.WriteLine("<TR><TD>&nbsp</TD></TR>")
		file.WriteLine("<TR><TD><FONT FACE=VERDANA SIZE=2><B>G</B> : Victoire par plus de 40 pts. (4 pts)</TD></TR>")
		file.WriteLine("<TR><TD><FONT FACE=VERDANA SIZE=2><B>GP</B>: Victoire par 40 pts ou moins. (3 pts)</TD></TR>")
		file.WriteLine("<TR><TD><FONT FACE=VERDANA SIZE=2><B>PN</B>: Partie Nulle. (2 pts)</TD></TR>")
		file.WriteLine("<TR><TD><FONT FACE=VERDANA SIZE=2><B>PP</B>: D�faite par 40 pts ou moins. (1 pts)</TD></TR>")
		file.WriteLine("<TR><TD><FONT FACE=VERDANA SIZE=2><B>P</B> : D�faite par plus de 40 pts. (0 pt)</TD></TR>")
		file.WriteLine("<TR><TD><FONT FACE=VERDANA SIZE=2><B>%</B> : C'est le pourcentage des points acquis sur les points possibles</TD></TR>")
		file.WriteLine("</TABLE>")

		file.Close()

	End Sub

	Sub subOutputCompteurs(ByRef strTitle As String, ByRef strFileName As String)


		Dim file = My.Computer.FileSystem.OpenTextFileWriter(
		strFileName, False, Text.Encoding.UTF8)

		'd�but du document
		file.WriteLine("<!DOCTYPE HTML PUBLIC -//W3C//DTD HTML 4.0 Transitional//EN><HTML><HEAD><TITLE></TITLE>")
		file.WriteLine("<body TOPMARGIN=10 BGPROPERTIES=""FIXED"" BGCOLOR = " & DOC_BACKGROUND_COLOR & ">")
		file.WriteLine("<CENTER>")
		file.WriteLine("<FONT FACE=VERDANA COLOR = black SIZE=5><B>" & strTitle & "</B></FONT><BR>")
		'Print #1, "(Minimum de 2 parties jou�es)<BR><BR>"
		file.WriteLine("<BR><BR>")

		file.WriteLine(funGetCompteurs(-1))
		file.WriteLine("<TABLE WIDTH=850 BORDER=0>")
		file.WriteLine("<TR><TD><FONT FACE=VERDANA SIZE=2><B>*</B>* % des points tot = Points du joueur / Nombre total de points atribu�s aux joueurs (pas �quipes) durant les parties de ce joueur. Comme �a, on tient en compte que certains questionnaires seront plus difficiles que d'autres. Il est � noter que les points possible tiennent compte de tous les parties jou�es avec un m�me questionnaire. 10 points sont ajout�s � ""Pts Poss."" si une question est r�ussie dans une des deux partie (ou les deux). Par exemple, pour la s�rie 2, si la question 1 est r�pondue dans la partie 1 mais pas dans la partie 2, que la question 2 est r�pondue dans la partie 2 seulement et que la question 3 est r�ussie dans les deux parties, les points possible seront de 30.</TD></TR>")
		file.WriteLine("<TR><TD><FONT FACE=VERDANA SIZE=2><B>Note:</B> Les questions s'adressant � l'�quipe ne comptent pour aucun joueur</TD></TR>")
		file.WriteLine("</TABLE>")

		file.Close()

	End Sub

	Function funGetCompteurs(ByVal iTeamNo As Short) As String

		Dim rsCompteurs As OleDb.OleDbDataReader
		Dim ColWidth(8) As Short
		Dim I As Short
		Dim S As String
		Dim strHREF As String
		'Dim MaxLigne As Short

		ColWidth(1) = 30
		ColWidth(2) = 210
		ColWidth(3) = 170
		ColWidth(4) = 40
		ColWidth(5) = 50
		ColWidth(6) = 90
		ColWidth(7) = 90
		ColWidth(8) = 100


		rsCompteurs = rsGetCompteurs(iTeamNo)

		S = ""

		'ent�te des compteurs
		S = S & "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=3>"
		S = S & "<TR align=center bgcolor = " & COL_HEADER_COLOR & ">"
		S = S & "<TD WIDTH=" & Str(ColWidth(1)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>Pos</B></FONT></TD>"
		S = S & "<TD WIDTH=" & Str(ColWidth(2)) & " ALIGN=LEFT><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>Nom</B></FONT></TD>"
		If iTeamNo = -1 Then
			S = S & "<TD WIDTH=" & Str(ColWidth(3)) & " ALIGN=LEFT><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>�quipe</B></FONT></TD>"
		End If
		S = S & "<TD WIDTH=" & Str(ColWidth(4)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>PJ</B></FONT></TD>"
		S = S & "<TD WIDTH=" & Str(ColWidth(5)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>Pts</B></FONT></TD>"
		S = S & "<TD WIDTH=" & Str(ColWidth(6)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>Pts Poss.</B></FONT></TD>"
		S = S & "<TD WIDTH=" & Str(ColWidth(7)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>Pts par PJ</B></FONT></TD>"
		S = S & "<TD WIDTH=" & Str(ColWidth(8)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>% des Pts Tot. *</B></FONT></TD>"
		S = S & "</TR>"

		While rsCompteurs.Read()
			'If Not rsCompteurs.EOF Then
			'	rsCompteurs.MoveFirst()
			'	I = 1
			'	While Not rsCompteurs.EOF
			System.Windows.Forms.Application.DoEvents()
				S = S & "<TR BGCOLOR = white ALIGN=CENTER>"
				S = S & "<TD><FONT FACE=VERDANA SIZE=2>" & Trim(Str(I)) & "</FONT></TD>"
			S = S & "<TD ALIGN=LEFT><FONT FACE=VERDANA SIZE=2>" & rsCompteurs(1).ToString & " " & rsCompteurs(2).ToString & "</FONT></TD>"
			If iTeamNo = -1 Then
				strHREF = "<A HREF = """ & "equipe_" & Trim(Str(rsCompteurs(0).ToString)) & ".htm" & """  > " & rsCompteurs(3).ToString & "</A>"
				S = S & "<TD ALIGN=LEFT><FONT FACE=VERDANA SIZE=2>" & strHREF & "</FONT></TD>"
			End If
			S = S & "<TD><FONT FACE=VERDANA SIZE=2>" & rsCompteurs(4).ToString & "</FONT></TD>"
			S = S & "<TD><FONT FACE=VERDANA SIZE=2>" & rsCompteurs(5).ToString & "</FONT></TD>"
			S = S & "<TD><FONT FACE=VERDANA SIZE=2>" & rsCompteurs(6).ToString & "</FONT></TD>"
			'S = S & "<TD><FONT FACE=VERDANA SIZE=2>" & System.Math.Round(rsCompteurs(7).ToString, 2) & "</FONT></TD>"
			'S = S & "<TD><FONT FACE=VERDANA SIZE=2>" & System.Math.Round(rsCompteurs(8).ToString, 2) & "</FONT></TD>"
			S = S & "<TD><FONT FACE=VERDANA SIZE=2>" & rsCompteurs(7).ToString & "</FONT></TD>"
			S = S & "<TD><FONT FACE=VERDANA SIZE=2>" & rsCompteurs(8).ToString & "</FONT></TD>"
			S = S & "</TR>"
		End While
		'Else
		'	'C'est le dubut de la saison. Mettre des lignes vides
		'	For I = 1 To 15
		'		S = S & "<TR BGCOLOR = white ALIGN=CENTER>"
		'		S = S & "<TD>&nbsp</TD><TD></TD><TD></TD><TD></TD><TD></TD><TD></TD><TD></TD><TD></TD>"
		'		S = S & "</TR>"
		'	Next
		'End If
		S = S & "</TABLE>"

		funGetCompteurs = S

	End Function

	'Function funGetCompteurs_old(ByVal iTeamNo As Integer) As String
	'
	'   Dim rsCompteurs As ADODB.Recordset
	'   Dim ColWidth(8) As Integer
	'   Dim I As Integer
	'   Dim S As String
	'   Dim strHREF As String
	'
	'   ColWidth(1) = 30
	'   ColWidth(2) = 210
	'   ColWidth(3) = 170
	'   ColWidth(4) = 40
	'   ColWidth(5) = 50
	'   ColWidth(6) = 90
	'   ColWidth(7) = 90
	'   ColWidth(8) = 100
	'
	'   Set rsCompteurs = rsGetCompteurs(iTeamNo)
	'
	'   S = ""
	'
	'   'ent�te des compteurs
	'   S = S & "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=3>"
	'   S = S & "<TR align=center bgcolor = " & COL_HEADER_COLOR & ">"
	'   S = S & "<TD WIDTH=" & Str(ColWidth(1)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>Pos</B></FONT></TD>"
	'   S = S & "<TD WIDTH=" & Str(ColWidth(2)) & " ALIGN=LEFT><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>Nom</B></FONT></TD>"
	'   If iTeamNo = -1 Then
	'      S = S & "<TD WIDTH=" & Str(ColWidth(3)) & " ALIGN=LEFT><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>�quipe</B></FONT></TD>"
	'   End If
	'   S = S & "<TD WIDTH=" & Str(ColWidth(4)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>PJ</B></FONT></TD>"
	'   S = S & "<TD WIDTH=" & Str(ColWidth(5)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>Pts</B></FONT></TD>"
	'   S = S & "<TD WIDTH=" & Str(ColWidth(6)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>Pts Poss.</B></FONT></TD>"
	'   S = S & "<TD WIDTH=" & Str(ColWidth(7)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>Pts par PJ</B></FONT></TD>"
	'   S = S & "<TD WIDTH=" & Str(ColWidth(8)) & "><FONT FACE=VERDANA SIZE=2 COLOR='FFFFFF'><B>% des Pts Tot. *</B></FONT></TD>"
	'   S = S & "</TR></TABLE>"
	'   S = S & ""
	'
	'   S = S & "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=3>"
	'
	'   rsCompteurs.MoveFirst
	'   I = 1
	'   While Not rsCompteurs.EOF
	'      DoEvents
	'      S = S & "<TR BGCOLOR = white ALIGN=CENTER>"
	'      S = S & "<TD WIDTH=" & Str(ColWidth(1)) & "><FONT FACE=VERDANA SIZE=2>" & Trim(Str(I)) & "</FONT></TD>"
	'      S = S & "<TD WIDTH=" & Str(ColWidth(2)) & " ALIGN=LEFT><FONT FACE=VERDANA SIZE=2>" & rsCompteurs.Fields(1) & "</FONT></TD>"
	'      If iTeamNo = -1 Then
	'         strHREF = "<A HREF = """ & "equipe_" & Trim(Str(rsCompteurs.Fields(0))) & ".htm" & """  > " & rsCompteurs.Fields(2) & "</A>"
	'         S = S & "<TD WIDTH=" & Str(ColWidth(3)) & " ALIGN=LEFT><FONT FACE=VERDANA SIZE=2>" & strHREF & "</FONT></TD>"
	'      End If
	'      S = S & "<TD WIDTH=" & Str(ColWidth(4)) & "><FONT FACE=VERDANA SIZE=2>" & rsCompteurs.Fields(3) & "</FONT></TD>"
	'      S = S & "<TD WIDTH=" & Str(ColWidth(5)) & "><FONT FACE=VERDANA SIZE=2>" & rsCompteurs.Fields(4) & "</FONT></TD>"
	'      S = S & "<TD WIDTH=" & Str(ColWidth(6)) & "><FONT FACE=VERDANA SIZE=2>" & rsCompteurs.Fields(5) & "</FONT></TD>"
	'      S = S & "<TD WIDTH=" & Str(ColWidth(7)) & "><FONT FACE=VERDANA SIZE=2>" & Round(rsCompteurs.Fields(6), 2) & "</FONT></TD>"
	'      S = S & "<TD WIDTH=" & Str(ColWidth(8)) & "><FONT FACE=VERDANA SIZE=2>" & Round(rsCompteurs.Fields(7), 2) & "</FONT></TD>"
	'      S = S & "</TR>"
	'      rsCompteurs.MoveNext
	'      I = I + 1
	'   Wend
	'   S = S & "</TABLE>"
	'
	'   funGetCompteurs_old = S
	'
	'End Function


	Function funGetCalendrier(ByVal iTeamNo As Short) As String
		Dim strTeamBPts As Object
		Dim strTeamAPts As Object
		Dim rsGames As OleDb.OleDbDataReader
		Dim rsPlayersTeamAPts As OleDb.OleDbDataReader
		'Dim rsPlayersTeamBPts As OleDb.OleDbDataReader
		Dim ColWidth(13) As Short
		Dim strTeamName As String
		Dim strLine As String
		Dim strDate As String
		Dim strPlayerPercent As String
		Dim strHREF As String
		'Dim I As Short
		Dim S As String
		Dim TextToDisplay As String

		'		Dim dt As New DataTable

		ColWidth(1) = 20
		ColWidth(2) = 30
		ColWidth(3) = 80
		ColWidth(4) = 55
		ColWidth(5) = 170
		ColWidth(6) = 70
		ColWidth(7) = 170
		ColWidth(8) = 70
		ColWidth(9) = 140
		ColWidth(10) = 170
		ColWidth(11) = 150
		ColWidth(12) = 150
		ColWidth(13) = 150

		'Lecture de la table parties
		rsGames = rsGetCalendrier(iTeamNo)




		S = ""

		'ent�te du calendrier
		S = S & "<TABLE style=""width:100%"" NOWRAP CELLPADDING=1 CELLSPACING=0>"
		S = S & "<TR BGCOLOR = " & COL_HEADER_COLOR & ">"
		S = S & "<TD WIDTH=" & Str(ColWidth(1)) & ">" & "<FONT FACE=VERDANA COLOR = white SIZE=2></TD>"
		S = S & "<TD WIDTH=" & Str(ColWidth(2)) & ">" & "<FONT FACE=VERDANA COLOR = white SIZE=2><B>#</TD>"
		S = S & "<TD WIDTH=" & Str(ColWidth(3)) & ">" & "<FONT FACE=VERDANA COLOR = white SIZE=2><B>Date</TD>"
		S = S & "<TD WIDTH=" & Str(ColWidth(4)) & ">" & "<FONT FACE=VERDANA COLOR = white SIZE=2><B>Lieu</TD>"
		S = S & "<TD WIDTH=" & Str(ColWidth(5)) & ">" & "<FONT FACE=VERDANA COLOR = white SIZE=2><B>�quipe A</TD>"
		S = S & "<TD WIDTH=" & Str(ColWidth(6)) & ">" & "<FONT FACE=VERDANA COLOR = white SIZE=2><B>Pts</TD>"
		S = S & "<TD WIDTH=" & Str(ColWidth(7)) & ">" & "<FONT FACE=VERDANA COLOR = white SIZE=2><B>�quipe B</TD>"
		S = S & "<TD WIDTH=" & Str(ColWidth(8)) & ">" & "<FONT FACE=VERDANA COLOR = white SIZE=2><B>Pts</TD>"
		S = S & "<TD WIDTH=" & Str(ColWidth(9)) & ">" & "<FONT FACE=VERDANA COLOR = white SIZE=2><B>Animateur</TD>"
		S = S & "<TD WIDTH=" & Str(ColWidth(10)) & ">" & "<FONT FACE=VERDANA COLOR = white SIZE=2><B>Questionnaire</TD>"
		S = S & "<TD WIDTH=" & Str(ColWidth(11)) & ">" & "<FONT FACE=VERDANA COLOR = white SIZE=2><B>Lien R�union</TD>"
		S = S & "<TD WIDTH=" & Str(ColWidth(12)) & ">" & "<FONT FACE=VERDANA COLOR = white SIZE=2><B>Lien animateur</TD>"
		S = S & "<TD WIDTH=" & Str(ColWidth(13)) & ">" & "<FONT FACE=VERDANA COLOR = white SIZE=2><B>Lien joueurs</TD>"
		S = S & "</TR></TABLE>"
		S = S & ""

		While rsGames.Read()


			System.Windows.Forms.Application.DoEvents()
			strLine = "<TABLE style=""width:100%"" NOWRAP CELLPADDING=1 CELLSPACING=0><TR BGCOLOR = white>"
			'colonne 1 - symbole +
			If IsDBNull(rsGames(4).ToString) Then
				'la partie n'a pas encore eu lieu, ne pas mettre le symbole +
				strLine = strLine & "<TD WIDTH=" & Str(ColWidth(1)) & " ALIGN=center></TD>"
			Else
				'la partie a eu lieu, mettre le symbole +
				strLine = strLine & "<TD WIDTH=" & Str(ColWidth(1)) & " ALIGN=center><img src=""" & "plus.gif""" & " id=partie" & Trim(Str(rsGames(0).ToString)) & " style=""" & "cursor:hand""" & " class=Outline></TD>"
			End If
			'colonne 2 - num�ro de la partie
			strLine = strLine & "<TD WIDTH=" & Str(ColWidth(2)) & ">" & "<FONT FACE=VERDANA COLOR = black SIZE=2>" & Trim(Str(rsGames(0).ToString)) & "</TD>"
			'colonne 3 - Date
			strDate = funFormatDate(rsGames(1).ToString)
			strLine = strLine & "<TD WIDTH=" & Str(ColWidth(3)) & ">" & "<FONT FACE=VERDANA COLOR = black SIZE=2>" & strDate & "</TD>"
			'colonne 4 - Lieu
			strLine = strLine & "<TD WIDTH=" & Str(ColWidth(4)) & ">" & "<FONT FACE=VERDANA COLOR = black SIZE=2>" & Trim(rsGames(10).ToString) & "</TD>"
			'colonne 5 - Nom equipe A
			If IsDBNull(rsGames(2).ToString) Then
				strTeamName = ""
				strHREF = ""
			Else
				strTeamName = rsGames(2).ToString
				strHREF = "<A HREF = """ & "equipe_" & Trim(Str(rsGames(3).ToString)) & ".htm" & """  > " & strTeamName & "</A>"
			End If
			strLine = strLine & "<TD WIDTH=" & Str(ColWidth(5)) & ">" & "<FONT FACE=VERDANA COLOR = black SIZE=2>" & strHREF & "</TD>"
			'colonne 6 - Pts equipe A
			If IsDBNull(rsGames(4).ToString) Then
				strTeamAPts = ""
			Else
				strTeamAPts = rsGames(4).ToString
			End If
			strLine = strLine & "<TD WIDTH=" & Str(ColWidth(6)) & ">" & "<FONT FACE=VERDANA COLOR = black SIZE=2><B>" & strTeamAPts & "</TD>"
			'colonne 7 - Nom equipe B
			If IsDBNull(rsGames(5).ToString) Then
				strTeamName = ""
			Else
				strTeamName = rsGames(5).ToString
				strHREF = "<A HREF = """ & "equipe_" & Trim(Str(rsGames(6).ToString)) & ".htm" & """  > " & strTeamName & "</A>"
			End If
			strLine = strLine & "<TD WIDTH=" & Str(ColWidth(7)) & ">" & "<FONT FACE=VERDANA COLOR = black SIZE=2>" & strHREF & "</TD>"
			'colonne 8 - Pts equipe B
			If IsDBNull(rsGames(7).ToString) Then
				strTeamBPts = ""
			Else
				strTeamBPts = rsGames(7).ToString
			End If
			strLine = strLine & "<TD WIDTH=" & Str(ColWidth(8)) & ">" & "<FONT FACE=VERDANA COLOR = black SIZE=2><B>" & strTeamBPts & "</TD>"

			'colonne 9 - Animateur
			strLine = strLine & "<TD WIDTH=" & Str(ColWidth(9)) & ">" & "<FONT FACE=VERDANA COLOR = black SIZE=2>" & Trim(rsGames(8).ToString) & "</TD>"

			'colonne 10 - equipe questionnaire
			If IsDBNull(rsGames(9).ToString) Then
				strTeamName = ""
			Else
				strTeamName = rsGames(9).ToString
				strHREF = "<A HREF = '" & F_DE_MATCH & Trim(Str(rsGames(0).ToString)) & ".htm'><img src='sheet.gif' STYLE='cursor: hand' BORDER=0></A>"
			End If
			' Colonne 10 - si la partie n'a pas eu lieu, ne pas montrer l'icone "sheet" 
			If strTeamAPts = "" Then
				strHREF = ""
			End If
			strLine = strLine & "<TD WIDTH=" & Str(ColWidth(10)) & ">" & "<FONT FACE=VERDANA COLOR = black SIZE=2>" & strHREF & "&nbsp" & strTeamName & "</TD>"

			'colonne 11 - Lien vers la r�uion
			If IsDBNull(rsGames(12).ToString) Then
				TextToDisplay = ""
				strHREF = ""
			Else
				TextToDisplay = "Lien vers la r�union"
				strHREF = "<A HREF = """ & Trim(rsGames(12).ToString) & """" & "> " & TextToDisplay & "</A>"
			End If
			'strLine = strLine & "<TD style=""white-space:nowrap"" WIDTH=" & Str(ColWidth(11)) & ">" & "<FONT FACE=VERDANA COLOR = black SIZE=2>" & Trim(rsGames(12).ToString) & "</TD>"
			strLine = strLine & "<TD WIDTH=" & Str(ColWidth(11)) & ">" & "<FONT FACE=VERDANA COLOR = black SIZE=2>" & strHREF & "</TD>"

			'colonne 12 - Lien vers la console de l'animateur
			If IsDBNull(rsGames(13).ToString) Then
				TextToDisplay = ""
				strHREF = ""
			Else
				TextToDisplay = "Lien animateur"
				strHREF = "<A HREF = """ & Trim(rsGames(13).ToString) & """" & "> " & TextToDisplay & "</A>"
			End If
			strLine = strLine & "<TD WIDTH=" & Str(ColWidth(12)) & ">" & "<FONT FACE=VERDANA COLOR = black SIZE=2>" & strHREF & "</TD>"

			'colonne 13 - Lien vers la console de l'animateur
			If IsDBNull(rsGames(14).ToString) Then
				TextToDisplay = ""
				strHREF = ""
			Else
				TextToDisplay = "Lien joueurs"
				strHREF = "<A HREF = """ & Trim(rsGames(14).ToString) & """" & "> " & TextToDisplay & "</A>"
			End If
			strLine = strLine & "<TD WIDTH=" & Str(ColWidth(13)) & ">" & "<FONT FACE=VERDANA COLOR = black SIZE=2>" & strHREF & "</TD>"

			'Fin de la ligne partie
			strLine = strLine & "</TR></TABLE>"
			S = S & strLine

			'section detail de la partie
			If (strTeamAPts > "0") And (strTeamBPts > "0") Then
				'la partie a eu lieu, afficher le d�tail
				'aller chercher le pointage de chaque joueurs de l'�quipe A et B
				rsPlayersTeamAPts = rsGetPlayerPtsForAGame(rsGames(0).ToString)

				strLine = "<div id=partie" & rsGames(0).ToString & "d style=""" & "display:None" & """ > "
				strLine = strLine & "<TABLE NOWRAP CELLPADDING=1 CELLSPACING=0>"
				strLine = strLine & "<TR BGCOLOR = darkgray>"

				While rsPlayersTeamAPts.Read()
					'rsPlayersTeamAPts.Item(0).value

					'	While rsPlayersTeamBPts.Read()
					strLine = strLine & "<TR BGCOLOR = whitesmoke>"
					strLine = strLine & "<TD WIDTH=" & Str(ColWidth(1)) & "></TD>"
					'strLine = strLine & "<TD WIDTH=" & Str(ColWidth(2)) & "></TD>"
					strLine = strLine & "<TD WIDTH=" & Str(ColWidth(3)) & "></TD>"
					strLine = strLine & "<TD WIDTH=" & Str(ColWidth(4)) & "></TD>"

					strLine = strLine & "<TD WIDTH=" & Str(ColWidth(5)) & " BGCOLOR = white><FONT FACE=VERDANA COLOR = black SIZE=1>" & rsPlayersTeamAPts("Nom�quipe").ToString & "</TD>"
					strLine = strLine & "<TD WIDTH=" & Str(ColWidth(5)) & " BGCOLOR = white><FONT FACE=VERDANA COLOR = black SIZE=1>" & rsPlayersTeamAPts("joueur").ToString & "</TD>"

					strPlayerPercent = Str(System.Math.Round(rsPlayersTeamAPts("Points").ToString / rsGames(11).ToString * 100, 1))
					strLine = strLine & "<TD WIDTH=" & Str(40) & " BGCOLOR = white><FONT FACE=VERDANA COLOR = black SIZE=1>" & rsPlayersTeamAPts("Points").ToString & "</TD>"
					strLine = strLine & "<TD WIDTH=" & Str(40) & " BGCOLOR = white><FONT FACE=VERDANA COLOR = black SIZE=1>" & strPlayerPercent & "%</TD>"

					'TODO : V�rifier ce qui se passe si le joueur n'a pas fait de pts. il faut mettre des cellules vides mettre des cellules vides
					strLine = strLine & "<TD WIDTH=" & Str(ColWidth(9)) & "></TD>"
					strLine = strLine & "<TD WIDTH=" & Str(ColWidth(10)) & "></TD>"
					strLine = strLine & "</TR>"
				End While
				'End While
				strLine = strLine & "</TABLE></DIV>"
				S = S & strLine
				rsPlayersTeamAPts.Close()
				'rsPlayersTeamBPts.Close()
			End If

		End While

		funGetCalendrier = S
		rsGames.Close()
	End Function

	Sub subOutputEquipe(ByVal iTeamNo As Short, ByRef strTitle As String, ByRef strFileName As String)


		Dim file = My.Computer.FileSystem.OpenTextFileWriter(
		strFileName, False, Text.Encoding.UTF8)

		file.WriteLine("<HTML><HEAD><TITLE></TITLE></HEAD>")
		file.WriteLine("<BODY TOPMARGIN=10 BGPROPERTIES=""FIXED"" BGCOLOR = " & DOC_BACKGROUND_COLOR & "><div id=Outline>")
		file.WriteLine("<CENTER><FONT FACE=VERDANA COLOR = black SIZE=5><B>" & strTitle & "</FONT></CENTER>")
		file.WriteLine("<BR><BR>")
		file.WriteLine("<FONT FACE=VERDANA COLOR = black SIZE=4><B>calendrier</FONT><BR>")
		file.WriteLine("<BR>")

		file.WriteLine(funGetCalendrier(iTeamNo))
		file.WriteLine("<BR><BR>")
		file.WriteLine("<FONT FACE=VERDANA COLOR = black SIZE=4><B>joueurs</FONT><BR>")
		file.WriteLine("<BR>")
		file.WriteLine(funGetCompteurs(iTeamNo))
		file.WriteLine("<BR>")
		file.WriteLine(funGetScript())
		file.WriteLine("</HTML>")
		file.Close()
	End Sub

	Private Function funGetScript() As Object
		Dim S As String

		S = ""
		S = "<script>" & vbCrLf
		S = S & "<!--" & vbCrLf
		S = S & "var img1, img2;" & vbCrLf
		S = S & "img1 = new Image();" & vbCrLf
		S = S & "img1.src = """ & "plus.gif" & """;" & vbCrLf
		S = S & "img2 = new Image();" & vbCrLf
		S = S & "img2.src = """ & "minus.gif" & """;" & vbCrLf
		S = S & "" & vbCrLf
		S = S & "function doOutline() {" & vbCrLf
		S = S & "  var targetId, srcElement, targetElement;" & vbCrLf
		S = S & "  srcElement = window.event.srcElement;" & vbCrLf
		S = S & "  if (srcElement.className == """ & "Outline" & """) {" & vbCrLf
		S = S & "     targetId = srcElement.id + """ & "d" & """;" & vbCrLf
		S = S & "     targetElement = document.all(targetId);" & vbCrLf
		S = S & "     if (targetElement.style.display == """ & "none" & """) {" & vbCrLf
		S = S & "        targetElement.style.display = """ & """;" & vbCrLf
		S = S & "        if (srcElement.tagName == """ & "IMG" & """) {" & vbCrLf
		S = S & "           srcElement.src = """ & "minus.gif" & """;" & vbCrLf
		S = S & "        }" & vbCrLf
		S = S & "     } else {" & vbCrLf
		S = S & "        targetElement.style.display = """ & "none" & """;" & vbCrLf
		S = S & "        if (srcElement.tagName == """ & "IMG" & """) {" & vbCrLf
		S = S & "            srcElement.src = """ & "plus.gif" & """;" & vbCrLf
		S = S & "        }" & vbCrLf
		S = S & "     }" & vbCrLf
		S = S & "  }" & vbCrLf
		S = S & "}" & vbCrLf
		S = S & "" & vbCrLf
		S = S & "Outline.onclick = doOutline;" & vbCrLf
		S = S & "-->" & vbCrLf
		S = S & "</script>"

		'UPGRADE_WARNING: Couldn't resolve default property of object funGetScript. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		funGetScript = S
	End Function

	'Private Function funGetDayOfWeek(ByVal strD As String) As String
	'	'Retourne le jour de la semaine de la date strD
	'	Dim strJour As String

	'	strJour = Mid(VB6.Format(strD, "dddd, mmmm dd, yyyy"), 1, InStr(1, VB6.Format(strD, "dddd, mmmm dd, yyyy"), ",") - 1)

	'	Select Case strJour
	'		Case "Monday"
	'			strJour = "lun"
	'		Case "Tuesday"
	'			strJour = "mar"
	'		Case "Wednesday"
	'			strJour = "mer"
	'		Case "Thursday"
	'			strJour = "jeu"
	'		Case "Friday"
	'			strJour = "ven"
	'		Case "Saturday"
	'			strJour = "sam"
	'		Case "Sunday"
	'			strJour = "dim"
	'	End Select

	'	funGetDayOfWeek = strJour

	'End Function

	Private Function funFormatDate(ByVal strD As String) As String
		'change le format de la date de : YYYY-MM-DD � DD-sept-YY
		Dim strDate As String

		strDate = Mid(strD, 9, 2)

		Select Case Mid(strD, 6, 2)
			Case "01"
				strDate = strDate & " jan "
			Case "02"
				strDate = strDate & " fev "
			Case "03"
				strDate = strDate & " mar "
			Case "04"
				strDate = strDate & " avr "
			Case "05"
				strDate = strDate & " mai "
			Case "06"
				strDate = strDate & " jun "
			Case "07"
				strDate = strDate & " jui "
			Case "08"
				strDate = strDate & " aou "
			Case "09"
				strDate = strDate & " sep "
			Case "10"
				strDate = strDate & " oct "
			Case "11"
				strDate = strDate & " nov "
			Case "12"
				strDate = strDate & " dec "
		End Select

		strDate = strDate & Mid(strD, 3, 2)
		funFormatDate = strDate

	End Function

	Sub subOutputQuestionnaire(ByRef intQuestNo As Short, ByRef strFileName As String)

		Dim rsQuest As ADODB.Recordset
		Dim I As Short

		rsQuest = rsGetQuest(intQuestNo)

		Dim file = My.Computer.FileSystem.OpenTextFileWriter(
		strFileName, False, Text.Encoding.UTF8)


		'd�but du document
		file.WriteLine("<!DOCTYPE HTML PUBLIC -//W3C//DTD HTML 4.0 Transitional//EN><HTML><HEAD><TITLE></TITLE>")
		file.WriteLine("<body TOPMARGIN=10 BGPROPERTIES=""FIXED"" BGCOLOR = " & DOC_BACKGROUND_COLOR & ">")
		file.WriteLine("<CENTER>")
		file.WriteLine("<FONT FACE=VERDANA COLOR = black SIZE=5></FONT><BR>test<BR>")
		file.WriteLine("</CENTER>")

		rsQuest.MoveFirst()
		I = 1
		While Not rsQuest.EOF
			System.Windows.Forms.Application.DoEvents()

			'ent�te de la s�rie
			file.WriteLine("<TABLE BORDER=0 CELLSPACING=1 CELLPADING=3>")
			file.WriteLine("<TR BGCOLOR=darkCyan><TD COLSPAN=2><FONT FACE=VERDANA SIZE=4 COLOR=White><B><I>S�rie " & rsQuest.Fields(8).Value & "</I></B></TD>")
			file.WriteLine("<TD COLSPAN=1><FONT FACE=VERDANA SIZE=3 COLOR=White><B><I>" & rsQuest.Fields(9).Value & "</I></B></TD></TR>")
			file.WriteLine("<TR BGCOLOR=darkCyan><TD COLSPAN=2><FONT FACE=VERDANA SIZE=2 COLOR=White><B></B></TD>")
			file.WriteLine("<TD COLSPAN=1><FONT FACE=VERDANA SIZE=2 COLOR=White><B>" & rsQuest.Fields(10).Value & "</B></TD></TR>")
			file.WriteLine("</TABLE>")
			file.WriteLine("<BR>")

			I = I + 1
			rsQuest.MoveNext()
		End While

		'Print #1, "</TABLE>"

		file.Close()

	End Sub
End Module