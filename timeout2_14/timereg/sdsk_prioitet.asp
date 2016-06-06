<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/sdsk_func.asp"-->
<%

sub crmaktstatheader
%>
<table border=0 cellpadding=0 cellspacing=0 width="450">
	<tr>
	<td valign="top" width="163"><img src="../ill/logo_bg.gif" width="163" height="53" alt="" border="0"></td>
	<td valign="bottom"><b>Timeout Kontrolpanel - Incident prioitet</b><br>
	Tilføj, fjern eller rediger Incident prioitet.</td>
	</tr>
	</table><br>
<%
end sub


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	thisfile = "sdsk_kontrolpanel"
	kview = "j"
	
	if len(request("lastedit")) <> 0 then
	lastedit = request("lastedit")
	else
	lastedit = 0
	end if
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:40; visibility:visible;">
	<%
	call  sdsktopmenu()
	%>
	<br><br><br>
	<div id="slet" style="position:relative; left:0; top:10; visibility:visible; border:1px red dashed; background-color:#ffff99; width:300px; padding:10px;">
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> en Type. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><a href="sdsk_prioitet.asp?menu=tok&func=sletok&id=<%=id%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
	</tr>
	</table>
	</div>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM sdsk_prioitet WHERE id = "& id &"")
	Response.redirect "sdsk_prioitet.asp?menu=tok&shokselector=1"
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		<%
		errortype = 8
		call showError(errortype)
		
		else
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		function SQLBless2(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless2 = tmp
		end function
		
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		rsptid = SQLBless2(request("FM_rsptid"))
		
		dblfaktor = SQLBless2(request("FM_faktor"))
		if len(request("FM_kunweekend")) <> 0 then
		intkunweekend = 1
		else
		intkunweekend = 0
		end if
		
		if len(request("FM_prio_grp")) <> 0 then
		prio_grp = request("FM_prio_grp")
		else
		prio_grp = 1
		end if
		
		
		
		
		intadvisering = SQLBless2(request("FM_advisering"))
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO sdsk_prioitet (navn, editor, dato, responstid, faktor, kunweekend, advisering, priogrp, type) "_
		&" VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& rsptid &", "_
		&" "& dblfaktor &", "& intkunweekend &", "& intadvisering &", "& prio_grp &")")
				
				
				strSQL = "SELECT id FROM sdsk_prioitet ORDER by id DESC LIMIT 0, 1"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				lastedit = oRec("id")
				oRec.movenext
				wend
				
				oRec.close
				
		else
		oConn.execute("UPDATE sdsk_prioitet SET navn ='"& strNavn &"', editor = '" &strEditor &"', "_
		&" dato = '" & strDato &"', responstid = '"& rsptid &"', faktor = "& dblfaktor &", "_
		&" kunweekend = "& intkunweekend &", advisering = "& intadvisering &", priogrp = "& prio_grp &" WHERE id = "&id&"")
		
		lastedit = id
		
		end if
		
		
		
		Response.redirect "sdsk_prioitet.asp?menu=tok&shokselector=1&lastedit="&lastedit&"&FM_prio_grp="&prio_grp
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	rsptid = 0
	intkunweekend = 1
	
	else
	strSQL = "SELECT navn, editor, dato, responstid, faktor, kunweekend, advisering, type FROM sdsk_prioitet WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	rsptid = oRec("responstid")
	dblfaktor = oRec("faktor")
	intkunweekend = oRec("kunweekend")
	dblAdvi = oRec("advisering")
	
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	
	if len(request("prio_grp")) <> 0 then
	prio_grp = request("prio_grp")
	else
	prio_grp = 1
	end if
	
	
	if intkunweekend = 1 then
	kwChecked = "CHECKED"
	else
	kwChecked = ""
	end if
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:40; visibility:visible;">
	<%
	call  sdsktopmenu()
	%>
	<br><br><br>
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr><form action="sdsk_prioitet.asp?menu=tok&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="FM_prio_grp" value="<%=prio_grp%>">
	<input type="hidden" name="id" value="<%=id%>">
    	<td valign="top" colspan="2"><h3>Incident Type - <%=varbroedkrumme%></h3></td>
	</tr>
	<%if dbfunc = "dbred" then%>
	<tr>
		<td colspan="2" valign="bottom" style="height:30;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></td>
	</tr>
	<%end if%>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td style="padding-top:10px;"><b>Navn:</b></td>
		<td style="padding-top:10px;"><input type="text" name="FM_navn" value="<%=strNavn%>" size="40" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td style="padding-top:10px;"><b>Responstid:</b> (0 = ingen)</td>
		<td style="padding-top:10px;"><input type="text" name="FM_rsptid" value="<%=rsptid%>" size="5" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"> timer</td>
	</tr>
	<tr>
		<td style="padding-top:10px;"><b>Responder kun indenfor normal arbejds / åbningstid?:</b><br>
		Arbejdstid sættes i kontrolpanelet.</td>
		<td style="padding-top:10px;"><input type="checkbox" name="FM_kunweekend" value="1" <%=kwChecked%>></td>
	</tr>
	<tr>
		<td style="padding-top:10px;"><b>Advisering efter:</b><br>
		Efter antal dage med inaktivitet.</td>
		<td style="padding-top:10px;"><input type="text" name="FM_advisering" value="<%=dblAdvi%>" size="5" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"> dage</td>
	</tr>
	<tr>
		<td style="padding-top:10px;"><b>Faktor:</b><br>
		Bruges til at bestemme hvor mange klip der<br> skal tages fra en aftale, for hver time der registreres.</td>
		<td style="padding-top:10px;"><input type="text" name="FM_faktor" value="<%=dblFaktor%>" size="5" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td colspan="2"><br><br><img src="ill/blank.gif" width="100" height="1" alt="" border="0"><input type="image" src="../ill/<%=varSubVal%>.gif"></td>
	</tr>
	</form>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%case else
	
	if len(request("FM_prio_grp")) <> 0 then
	prio_grp = request("FM_prio_grp")
	else
	prio_grp = 1
	end if
	%>
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	</script>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:40; visibility:visible;">
	<%
	call  sdsktopmenu()
	%>
	<h3>Incident Type</h3><br>
	<a href="sdsk_prioitet.asp?menu=tok&func=opret&prio_grp=<%=prio_grp%>">Opret ny Type <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a><br>
	<br>
	
	
	<form method="post" action="sdsk_prioitet.asp">
	<b>Aftalegruppe:</b> &nbsp;<select name="FM_prio_grp" id="FM_prio_grp">
	<%
	strSQL = "SELECT id, navn FROM sdsk_prio_grp WHERE id <> 0"
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
		
		
		if cint(prio_grp) = oRec("id") then
		pgSEL = "SELECTED"
		else
		pgSel = ""
		end if%>
	<option value="<%=oRec("id")%>" <%=pgSel%>><%=oRec("navn")%></option>
	<%
	oRec.movenext
	wend
	
	oRec.close
	%>
	</select>&nbsp;<input type="submit" value="Vælg">
	</form>
	
	
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2><img src="../ill/tabel_top_left.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan=5 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td width="8" align=right valign=top rowspan=2><img src="../ill/tabel_top_right.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td width="50" class=alt><b>Id</b></td>
		<td width="300" class=alt><b>Type</b> (Antal inci.)</td>
		<td class=alt><b>Faktor</b></td>
		<td class=alt><b>Responstid (Timer)</b></td>
		<td width="50" class=alt>&nbsp;</td>
	</tr>
	<%
	strSQL = "SELECT p.id, p.navn, p.responstid, p.faktor, p.priogrp, COUNT(s.id) AS antal FROM sdsk_prioitet p "_
	&" LEFT JOIN sdsk s ON (s.prioitet = p.id) WHERE p.priogrp = "& prio_grp &" GROUP BY p.id ORDER BY p.navn"
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
	if cint(oRec("id")) = cint(lastedit) then
	bgCol = "#ffff99"
	else
	bgCol = "#ffffff"
	end if%>
	<tr>
		<td bgcolor="#003399" colspan="7"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bgCol%>" onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'<%=bgCol%>');">
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td><%=oRec("id")%></td>
		<td height="20"><a href="sdsk_prioitet.asp?menu=tok&func=red&id=<%=oRec("id")%>&prio_grp=<%=oRec("priogrp")%>"><%=oRec("navn")%> </a>&nbsp;&nbsp;(<%=oRec("antal")%>)</td>
		<td><%=oRec("faktor")%></td>
		<td><%=oRec("responstid")%></td>
		<td>
		<%if cint(oRec("antal")) = 0 then%>
		<a href="sdsk_prioitet.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="" border="0"></a>
		<%end if%>&nbsp;</td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>
	<tr bgcolor="#5582D2">
		<td width="8" valign=top height=20 style="border-bottom:1px #003399 solid; border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=5 valign="top" style="border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top style="border-bottom:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>	
	</table>
	<br><br><br>
	<a href="Javascript:window.close()" class=vmenu>Luk vindue, og vend tilbage til kontrolpanel</a><br><br>
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
