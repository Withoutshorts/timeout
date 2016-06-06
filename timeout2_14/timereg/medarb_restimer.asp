<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
		

else
print = request("print")%>
<html>
		<head>
		<title>TimeOut 2.1</title>
		<%if print <> "j" then%>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print.css">
		<%else%>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
		<%end if%>
</head>
<body topmargin="0" leftmargin="0" class="regular">
<%
'*** Sætter lokal dato/kr format. *****
Session.LCID = 1030
medarbid = request("FM_medarb")
medarbnavn = request("FM_medarbnavn")

	if len(request("FM_start_dag")) <> 0 then
	strMrd = request("FM_start_mrd")
	strDag = request("FM_start_dag")
	strAar = right(request("FM_start_aar"),2) 
	strDag_slut = request("FM_slut_dag")
	strMrd_slut = request("FM_slut_mrd")
	strAar_slut = right(request("FM_slut_aar"),2)
	else
	
		strMrd = month(now)
		strDag = day(now)
		strAar = right(year(now), 2)
		
		strDag_slut_temp = dateadd("d", 30, date)
		strDag_slut = day(strDag_slut_temp)
		strMrd_slut_temp = dateadd("m", 1, date)
		strMrd_slut = month(strMrd_slut_temp)
		
		if strMrd_slut <> 1 then
		strAar_slut = right(year(now), 2)
		else
		strAar_slut_temp = dateadd("yyyy", 1, date)
		strAar_slut = right(year(strAar_slut_temp),2)
		end if
		
	end if
	
	
	strStartDato = strAar&"/"&strMrd&"/"&strDag
	strSlutDato = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut

%>

<%if print <> "j" then%>
<div style="position:absolute; left:20; top:20">
<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="575">
<tr bgcolor="#5582D2">
	<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
	<td colspan=4 valign="top"><img src="../ill/tabel_top.gif" width="559" height="1" alt="" border="0"></td>
	<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
</tr>
<tr bgcolor="#5582D2">
	<td colspan=4 valign="top" class="alt">&nbsp;<b>Job og periode:</b></td>
</tr>	
<tr>
	<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
	<form action="medarb_restimer.asp?menu=medarb" method="post">
	<td><input type="hidden" name="FM_medarb" value="<%=medarbid%>">
	<input type="hidden" name="FM_medarbnavn" value="<%=medarbnavn%>"></td>
	<!--#include file="inc/weekselector_b.asp"-->
	<td><input type="image" src="../ill/statpil.gif" vspace=5></td>
	<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
</tr>
</form>
<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan=4 valign="bottom"><img src="../ill/tabel_top.gif" width="559" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
</tr>
</table>
<br><br>
<%else%>
<div style="position:absolute; left:0; top:0">
<table cellspacing="0" cellpadding="0" border="0" width="680">
				<tr>
					<td bgcolor="#003399" width="680"><img src="../ill/logo_topbar_print.gif" alt="" border="0"></td></tr>
					<tr style="padding-right:70; padding-top:5;"><td bgcolor="#FFFFFF" align=right><a href="javascript:window.close()"><img src="../ill/luk_xp.gif" width="30" height="28" alt="" border="0">&nbsp;Luk</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"><a href="javascript:window.print()"><img src="../ill/print_xp.gif" width="28" height="30" alt="" border="0">&nbsp;Print</a></td>
				</tr>
</table>
</div>
<div style="position:absolute; left:20; top:90">
<%end if%>

<table cellspacing="0" cellpadding="2" border="0">
	<tr><td colspan="7"><a href="Javascript:history.back()" class=vmenu><img src="../ill/soeg-knap_tilbage.gif" width="16" height="16" alt="" border="0">&nbsp;Tilbage</a>
	<br>Følgende timer er projekteret til <b><%=medarbnavn%></b> i den valgte periode:
	<%if print <> "j" then%>
	<img src="../ill/blank.gif" width="110" height="1" alt="" border="0"><a href="medarb_restimer.asp?FM_medarb=<%=medarbid%>&FM_medarbnavn=<%=medarbnavn%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>&print=j" class=vmenu>Printer venlig version&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
	<%else%>
	<br>Periode: <b><%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 0)%></b> til <b><%=formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 0)%></b>
	<%end if%>
	</td></tr>
	<tr><td valign="top">
	
	
	<%
	
	
	'*** tildelte timer ***
	strSQL3 = "SELECT DISTINCT(ressourcer.id), ressourcer.dato AS rdato, timer AS sumtimer, job.jobnavn, jobid, job.jobnr AS jnr, job.id AS jid, aktiviteter.navn AS anavn FROM ressourcer LEFT JOIN job ON (job.id = ressourcer.jobid) LEFT JOIN aktiviteter ON (aktiviteter.id = aktid) WHERE (ressourcer.dato >= '"&strStartDato&"' AND ressourcer.dato <= '"&strSlutDato&"') AND mid=" & medarbid &" GROUP BY ressourcer.id ORDER BY ressourcer.dato, jobnavn"
	oRec3.open strSQL3, oConn, 3 
	x = 0
	y = 0
	lastresdato = 0
	lastweek = 0
	while not oRec3.EOF 
	if x = 0 then%>
	</td></tr>
	<tr>
			<td valign=top colspan=7 style="border-bottom:1 #003399 solid;"><br><b>Uge: <%=datepart("ww", oRec3("rdato"),2,2)%></b></td>
	</tr>
	<tr><td valign="top">
	<%end if
		
		if x > 0 AND (lastresdato <> oRec3("rdato")) then%>
		</table></td>
			<%if lastweek <> datepart("ww", oRec3("rdato"),2,2) then%>
			</tr><tr>
					<%if lastweek <> datepart("ww", oRec3("rdato"),2,2) OR y = 0 then%>
					<td colspan=7 valign=top style="border-bottom:1 #003399 solid;"><br><b>Uge: <%=datepart("ww", oRec3("rdato"),2,2)%></b></td>
					</tr><tr>
					<%end if%>
			<%end if%>
			<td valign="top">
		<%end if%>
		
			<%if lastresdato <> oRec3("rdato") then
				select case weekday(oRec3("rdato"))
				case 1, 7 
				bgthis = "Silver"
				'case 2
				'bgthis = "#ffffe1"
				case else
				bgthis = "#ffffff"
				end select
				%>
				<table cellspacing="0" cellpadding="0" border="0" width=111 bgcolor="<%=bgthis%>" style="border:1 #003399 solid;">
				<tr>
					<td colspan=2 style="padding-left:3;"><font class="lillesort"><b><%=left(weekdayname(weekday(oRec3("rdato"))),3)%>&nbsp;<%=formatdatetime(oRec3("rdato"),2)%></b></td></tr>
				<%
				y = y + 1
			end if
		
		if lastjobid <> oRec3("jid") OR (lastresdato <> oRec3("rdato")) then%>
		<tr><td colspan=2 class=lille style="padding-left:3; padding-right:3;"><u>(<%=oRec3("jnr")%>) <%=oRec3("jobnavn")%></u></td></tr>
		<%end if%>
		<tr><td class=lille style="padding-left:3; padding-right:3;"><%=oRec3("anavn")%></td><td class=lille style="padding-left:2; padding-right:3;"><%=formatnumber(oRec3("sumtimer"),2)%></td></tr>
	<%
	timertottildelt = timertottildelt + oRec3("sumtimer")
	lastjobid = oRec3("jid")
	lastresdato = oRec3("rdato")
	lastweek = datepart("ww", oRec3("rdato"),2,2)
	x = x + 1
	oRec3.movenext
	wend
	oRec3.close
	%>
	</table></td></tr>
	</table>
	<%if x = 0 then%>
	<div style="position:relative; left:30; top:0;">
	<table>
	<tr><td colspan=2 bgcolor="#ffffe1" style="padding:5; border:1px darkred solid;"><img src="../ill/blank.gif" width="500" height="1" alt="" border="0"><br><br><img src="../ill/alert_lille.gif" width="22" height="19" alt="" border="0">Der er ikke tildelt ressource timer på den valgte medarbejder i den valgte periode.<br>Vælg en ny periode i toppen af denen side.<br><br>Der kan tildeles ressource timer direkte fra <b>ressource belægnings</b> oversigten,<br> eller ved at redigere et job og klikke på ressourcer.<br><br>&nbsp;</td>
	</tr>
	</table></div>
	<%end if%>
	<br><br>
	<a href="Javascript:window.close()">Luk dette vindue</a>
	</div>
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->