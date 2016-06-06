<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/header_inc.asp"-->
<%
	'ktv/qwert Persits
	Set Upload = Server.CreateObject("Persits.Upload.1")
	'On Error Resume Next
	
	'Persists
	Upload.OverwriteFiles = False
	
	'ktv
	'Upload.Save "E:\www\timeout_xp\wwwroot\ver2_1\inc\upload\"&lto&"\" '44
	'Dell
	'Upload.Save "C:\www\cortex\wwwroot\ver3_0\upload\"
	'Timeout DB server
	Upload.Save "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\upload\"&lto&"\" '58
	
	strConnect = strConn
	
	%>
<div id="sindhold" style="position:absolute; left:20; top:20; visibility:visible;">
<%
	'Persists
	For Each File in Upload.Files
	antalLto = len(lto)
	varAntChar = len(File.Path)
	varRemChar = (varAntChar - (58 + antalLto + 1))
	strFilename = Right(File.Path, varRemChar)
	
	
	
		'Persists
		Set oCmd = server.createobject("ADODB.command")
			oCmd.activeconnection = strConnect
			oCmd.commandText = "UPDATE filer SET filnavn = '" & strFilename & "', editor = '"& session("user") &"', dato = '"& date &"' WHERE oprses = 'x'"
			oCmd.commandType = 1
			oCmd.Execute
			
		set oCmd.activeconnection = Nothing
		
		'File.ToDatabase strConnect, "UPDATE filer SET navn = '" & strFilename & "', editor = '"& session("user") &"', dato = '"& date &"' WHERE oprses = 'x'"
		if Err <> 0 Then '3
			Response.Write "Error saving the file: " & Err.Description
			%>
			<br>&nbsp;&nbsp;&nbsp;<font color="darkRed">Fejl!</font>
				<table cellspacing="4" cellpadding="2" border="0">
				<tr>
				<td>
			<%
			Response.Write "<font face=verdana,arial,helvetica size=2>"&"<b>Fejl</b><BR><BR><BR>"
			Response.Write "Et billede med samme navn ("& strFilename &") eksisterer allerede" &"<BR><BR><BR>"
			Response.Write "Husk! at der kun kan uploades office dokumenter, gif og jpg filer.<br><br>"
			%><a href="Javascript:history.back()"><img src="../ill/left.gif" width="20" height="13" alt="" border="0">&nbsp;&nbsp;Prøv igen</a><% 
		
		Else
			'Persists
			'Set File = nothing
			
			Set oCmd = server.createobject("ADODB.command")
			oCmd.activeconnection = strConnect
			oCmd.commandText = "UPDATE filer SET oprses = '' WHERE oprses = 'x'"
			oCmd.commandType = 1
			oCmd.Execute
			
			set oCmd.activeconnection = Nothing
			
			
			strFiltype = Right(strFilename, 3)
			'Response.write strFiltype
			
			if strFiltype = "doc" OR strFiltype = "xls" OR strFiltype = "txt" then '4
			%>
			<br>&nbsp;&nbsp;&nbsp;Office dokumentet er lagt ud på Webserveren.
			<table cellspacing="4" cellpadding="2" border="0">
			<tr>
			<td>
			<%
            Response.Write "<br><br><font face='verdana,arial,helvetica' size='-2'>"&"Du har uploadet dokumentet <b>" & strFilename & "</b>."
            Response.Write "<br><br>"
			Response.write "<a href='Javascript:window.close()'>"&"Luk vindue"&"</a><br><br>"
			else
			%>
			<br>&nbsp;&nbsp;&nbsp;Billedet er lagt ud på Webserveren.
			<table cellspacing="4" cellpadding="2" border="0">
			<tr>
			<td>
			<%
            Response.Write "<br><br><font face='verdana,arial,helvetica' size='-2'>"&"Du har uploadet billedet <b>" & strFilename & "</b>."
            Response.Write "<br><br>"
			%>
			<img src="../inc/upload/<%=lto%>/<%=strFilename%>" alt="" border="0"><br><br>
			<%
			Response.write "<a href='Javascript:window.close()'>"&"Luk vindue"&"</a><br><br>"
			End If '4
			
			
		End If '3
	Next
		%>
</td></tr></table>
</div>