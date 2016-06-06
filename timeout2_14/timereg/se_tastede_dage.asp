<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/header_inc.asp"-->	
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/vmenu.asp"-->
<!--#include file="../inc/regular/rmenu.asp"-->
<!--#include file="inc/convertDate.asp"--> 
<%
if session("user") = "" then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<%
	errortype = 5
	call showError(errortype)
	else
	%>
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	</script>
	<div id="sindhold" style="position:absolute; left:190; top:50; height:600; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
    <td valign="top"><br>
	<!-------------------------------Sideindhold------------------------------------->
	<font class="pageheader">Medarbejder Log </font>
<%
if len(request("selmedarb")) > 0 then
sel_medarb = request("selmedarb")

strSQL = "SELECT Mnavn, Mid FROM medarbejdere WHERE Mid="& sel_medarb &"" 
			oRec.Open strSQL, oConn, 3
			if Not oRec.EOF then
			strMnavn = oRec("Mnavn")
			end if
			oRec.Close
else
strMnavn = Session("user")
sel_medarb = session("Mid")
end if

'oRec.Open "SELECT Tdato, TasteDato, Tjobnr, Tjobnavn, aktiviteter.navn AS Anavn, Tknavn, Timer, Tid, Tfaktim, Timerkom FROM timer, aktiviteter, job WHERE Tmnr = "& sel_medarb &" AND Tid > 0 AND aktiviteter.id = TAktivitetId ORDER BY Tdato DESC", oConn, 3
strSQL = "SELECT Tdato, TasteDato, Tjobnr, Tjobnavn, TAktivitetNavn AS Anavn, TAktivitetId, Tknavn, Tmnr, Tmnavn, Timer, Tid, Tfaktim, TimePris, Timerkom FROM timer WHERE Tmnr = "& sel_medarb &" AND Tid > 0 AND tfaktim <> 5 ORDER BY Tdato DESC, Tmnavn"
oRec.open strSQL, oConn, 0, 1

	
%><br>
<Font color="darkRed"> <%=strMnavn%></font><br>
<div style="position:absolute; top:10; left:463; width:32; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:3px;"><a href="word.asp?eks=<%=request("eks")%>&jobnr=<%=intJobnr%>&medlog=1"><img src="../ill/word.gif" width="23" height="22" alt="" border="0"></a></div>
<div style="position:absolute; top:10; left:499; width:32; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:3px;"><a href="Javascript:window.print()"><img src="../ill/small_print.gif" width="23" height="22" alt="" border="0"></a></div>

<br>

<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>

<br>
<table border="0" cellpadding="0" cellspacing="0" style="position:absolute; width:655; !border: 1px;  border-color: Black; border-style: solid; z-index:0; border-bottom:0; visibility:visible;">
<tr height="20">
	<td width=120><font size="1">&nbsp;&nbsp;Uge&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Dato</td><td width=80><font size="1">Kunde</td><td align="center" width="60"><font size="1">Jobnr</td><td width="80"><font size="1">Jobnavn</td><td width="110"><font size="1">Aktivitet</td><td align="center" width="50"><font size="1">Timer</td><td align="center" width="100"><font size="1">Taste dato</td><td align="center" width="40">&nbsp;</td>
</tr>
	
	<%
	while not oRec.EOF
	strTdato = oRec("Tdato")
	convertDate(strTdato )
	strWeekNum = datepart("ww", strTdato,2,2)
	%>
	<tr>
		<td bgcolor="#003399" colspan="9"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr onmouseover="mOvr('gift',this,'#5582D2');" onmouseout="mOut(this,'');">
	<td height="18"><font size=1>&nbsp;&nbsp;<%=strWeekNum%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=strTdato%></td>
	<td><font size=1 color="#708090"><%=left(oRec("Tknavn"), 16)%></td>
	<td align="center"><font size=1><%=oRec("Tjobnr")%></td>
	<td><font size=1 color="gray"><%=left(oRec("Tjobnavn"),16)%></td>
	<%if oRec("Tfaktim") = "1" OR oRec("Tfaktim") = "2" then%>
	<td><font size=1><%=left(oRec("Anavn"), 20)%></td>
	<%else%>
	<td><font size=1>&nbsp;-</td>
	<%end if%>
	<td align="center"><font size=1><%=oRec("Timer")%></td>
	<td align="center"><font color="gray" size="1"><%=convertDate(oRec("TasteDato"))%></font></td>
	<td><a href="rediger_tastede_dage.asp?id=<%=oRec("Tid")%>"><img src="../ill/rediger.gif" width="12" height="12" alt="" border="0"></a>&nbsp;&nbsp;
	<%if len(trim(oRec("Timerkom"))) <> "0" then%>
	<img src="../ill/comments.gif" width="20" height="15" alt="" border="0">&nbsp;
	<%end if%>
	</td>
	</tr>
	<%
	v = v + 1
	oRec.movenext
	wend	  
	oRec.Close
	Set oRec = Nothing

if v < 10 then	
q = 0
for q = 0 to 10%>
<tr>
		<td bgcolor="#003399" colspan="9"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
<tr>
		<td colspan="9"><img src="ill/blank.gif" width="1" height="12" border="0" alt=""></td>
	</tr>
<%next
end if%>
<tr>
		<td bgcolor="LightBlue" colspan="9"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
<tr>
		<td colspan="9"><img src="ill/blank.gif" width="1" height="12" border="0" alt=""></td>
	</tr>
</table>
<br><br><br>
<br>
</TD>
</TR>
</TABLE>
</div>
<%end if %>
<!--#include file="../inc/regular/footer_inc.asp"-->