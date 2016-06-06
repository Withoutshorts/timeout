<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/header_inc.asp"-->	
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/vmenu.asp"-->
<!--#include file="../inc/regular/rmenu.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"--> 

<!-- ska tester -->

<div id="sindhold" style="position:absolute; left:170; top:80; width:60%; height:600; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
    <td valign="top"><br>
	<font class="pageheader">Status for indtasninger:</font>
	 <!-------------------------------Sideindhold------------------------------------->
<br>
<br>
<%
dim varbookmark
 
dim avarFields
dim strFields
Dim strMTjobnr
DIm strDatoCheck
Dim strDubletTjeck
strDubletTjeck = 0
x = 0
	
	oRec.Open "SELECT MTdato, MTjobnr, MTjobnavn, MTknavn, MTimer, Mtid.id AS MT_id, MTaktivitetId, aktiviteter.navn AS Anavn FROM Mtid, aktiviteter WHERE MTmnavn= '"& Session("user") &"' AND aktiviteter.id = MTaktivitetId", oConn, 3
	
	%>
	<table width=80% border="0" cellspacing="0" cellpadding="0">
	<%
	
	If oRec.BOF then
	strEmpty = 1
	Response.write "<tr><td>Der er ingen indtastninger at lægge ned i databasen" & "<br>" & "Klik på <a href='pwcon.asp'><i>tast ny dag</i></a> for at oprette en ny timeregistrering.</td></tr>"
	Else
	
	%>
	<tr>
		<td><b>Dato</b></td><td><b>Jobnr</b></td><td><b>Jobnavn</b></td><td><b>Aktivitet</b></td><td>&nbsp;</td>
	</tr>
	<%
	
	avarFields = oRec.GetRows
	strFields = Ubound(avarFields, 2)
	
	oRec.moveFirst
	varbookmark = oRec.Bookmark
	
	Do While x <= strFields 
	
	strAktivitetNavn = oRec("Anavn")
	strAktivitetId = oRec("MTAktivitetId")
	strMTjobnr = oRec("MTjobnr")
	strDatoCheck = oRec("MTdato")
	strJobnavn89 = oRec("MTjobnavn")
	
	%>
	<tr>
		<td bgcolor="#cccccc" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr>
		<td><%
	Response.write convertDate(strDatoCheck) &"</td><td>"& strMTjobnr &"</td><td>" & strJobnavn89 &"</td><td>" & strAktivitetNavn &"</td>"
	
	varbookmark = oRec.Bookmark
	 
		Set odubRec = Server.CreateObject ("ADODB.Recordset")
		odubRec.open "SELECT * FROM Timer WHERE Tmnavn = '"&session("user")&"' AND TAktivitetId = "& strAktivitetId &" AND Tjobnr = '"&strMTjobnr&"' AND Tdato = #"&strDatoCheck&"# ", oConn, 3
		
		if odubRec.BOF Then
		Response.write "<td>&nbsp;</td>"
		strDubletTjeck = strDubletTjeck
		else
		Response.write "<td><FONT FACE='verdana,helvetica,arial' SIZE='2' COLOR='darkRed'>" &"<b>"&"Fejl!"&"</b></font>&nbsp;&nbsp;&nbsp;<a href='slet_db.asp?func=afslut&id="&oRec("MT_id")&"'>slet</a>" &"</td>"
		strDubletTjeck = strDubletTjeck + 1
		end if 
		
		odubRec.Close
		Set odubRec = Nothing
		
	oRec.movenext
	x = x + 1
	
	%></td></tr><%
	
	loop
	end if

	oRec.Close
	 %>
	 <tr><td colspan="5">
	<%
			
			
			
			
			If strDubletTjeck < 1 then
				
			
			oRec.Open "SELECT * FROM mtid WHERE MTmnavn = '"& session("user")& "'", oConn, 3
			While not oRec.EOF
			
			oConn.Execute ("INSERT INTO timer (TJobnr, Tjobnavn, Tmnr, Tmnavn, Tknr, Tknavn, Tdato, Timer, Timerkom, Tfaktim, TAktivitetId) VALUES ('" & oRec("MTjobnr") & "', '" & oRec("MTjobnavn") & "' , '" & oRec("MTmnr") & "', '" & oRec("MTmnavn") & "', '" & oRec("MTknr") & "' , '" & oRec("MTknavn") & "', '" & oRec("MTdato") & "', '" & oRec("MTimer") & "', '" & oRec("MTimerkom") & "', "& oRec("MTfaktim") &", "& strAktivitetId &")")
			
			oRec.Movenext
 			Wend 
			
			oRec.Close
			
			oConn.Execute ("DELETE FROM Mtid WHERE MTmnavn = '"& session("user")& "'") 
		 
			if strEmpty <> 1 then
			Response.write  "<br><br><br><br>" & "Dine indtastninger er godkendt."
			end if
			
			else
			Response.write  "<br><br><br><br><br><font color=darkRed><b>Dine indtastninger er ikke godkendt!</b></font><br><br>Hvis du får en <b>fejl</b>, betyder det at denne aktivitet er allerede indtastet 1 gang med den pågældende dato.<br>For at <a href='Javascript:history.back()'>rette fejlen kan du klike her</a>.<br><br> Du kan fra oversigten over tastede dag altid redigere en tidligere indtastet aktivitet.<br><br>"
			end if
			 
%>

 </TD> 
</TR>
</table>
<br>
 
</TD> 
</TR>
</TABLE>
</div>
<!--#include file="../inc/regular/footer_inc.asp"-->
 
