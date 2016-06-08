<!--#include file="../connection/conn_db_inc.asp"-->
<!--#include file="../errors/error_inc.asp"-->
<html>
		<head>
		<title>timeOut 2.1</title>
		<LINK rel="stylesheet" type="text/css" href="../style/timeout_style_print_fak.css">
		</head>
<body topmargin="0" leftmargin="0" class="regular">
<table cellspacing="0" cellpadding="0" border="0" width="880">
	<tr>
		<td bgcolor="#003399" width="650"><img src="../../ill/logo_topbar_print.gif" alt="" border="0"></td>
		<td bgcolor="#FFFFFF" align=right><a href="javascript:window.close()"><img src="../../ill/luk_xp.gif" width="30" height="28" alt="" border="0">&nbsp;Luk</a><img src="../../ill/blank.gif" width="30" height="1" alt="" border="0"><a href="javascript:window.print()"><img src="../../ill/print_xp.gif" width="28" height="30" alt="" border="0">&nbsp;Print</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"></td>
	</tr>
</table>
<br><br>
<table cellspacing="0" cellpadding="0" border="0" width="580">
	<tr><td style="padding-left:30;" colspan=2>Herunder følger en samlet log.<br> For at oprette en ny log skal der oprettes en ny aktion.<br><br>
	<b>Kunde log:</b></td></tr>
<tr><td width="25"><img src="../../ill/blank.gif" width="1" height="1" alt="" border="0"></td><td bgcolor="#003399" height="1" width="555"><img src="../../ill/blank.gif" width="1" height="1" alt="" border="0"></td></tr>

<%
'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
Session.LCID = 1030

strSQL = "SELECT editor, kdato, beskrivelse FROM kunder WHERE kid = " & request("usekid")
oRec.open strSQL, oConn, 3
if not oRec.EOF then
%>
<tr><td style="padding-left:30; padding-top:2;" colspan=2><b><%=oRec("kdato")%></b><br><%=oRec("beskrivelse")%><br><i><%=oRec("editor")%></i></td></tr>
<tr><td width="25"><img src="../../ill/blank.gif" width="1" height="1" alt="" border="0"></td><td bgcolor="#003399" height="1" width="555"><img src="../../ill/blank.gif" width="1" height="1" alt="" border="0"></td></tr>
<%
end if
oRec.close

%><tr><td style="padding-left:30;" colspan=2><br><br>
<b>Aktions log</b></td></tr>
<tr><td width="25"><img src="../../ill/blank.gif" width="1" height="1" alt="" border="0"></td><td bgcolor="#003399" height="1" width="555"><img src="../../ill/blank.gif" width="1" height="1" alt="" border="0"></td></tr>
<%

strSQL = "SELECT editor, crmdato, komm FROM crmhistorik WHERE kundeid = " & request("usekid") &" ORDER BY id DESC"
oRec.open strSQL, oConn, 3

while not oRec.EOF 
if len(oRec("komm")) <> 0 then
%>
<tr><td style="padding-left:30; padding-top:2;" colspan=2><b><%=formatdatetime(oRec("crmdato"), 1)%></b><br><%=oRec("komm")%><br><i><%=oRec("editor")%></i><br><br>&nbsp;</td></tr>
<tr><td width="25"><img src="../../ill/blank.gif" width="1" height="1" alt="" border="0"></td><td bgcolor="#003399" height="1" width="555"><img src="../../ill/blank.gif" width="1" height="1" alt="" border="0"></td></tr>
<%
end if
oRec.movenext
wend
oRec.close
%>
</table>

<!--#include file="footer_inc.asp"-->
