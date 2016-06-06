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
	<td valign="bottom"><b>Timeout Kontrolpanel - Prioitetsklasser.</b><br>
	Tilføj, fjern eller rediger Prioitetsklasser.</td>
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
	
	thisfile = "sdsk_kontrolpanel"
	kview = "j"
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
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
	    <td>Du er ved at <b>slette</b> en Prioitet. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><a href="sdsk_prio_typ.asp?menu=tok&func=sletok&id=<%=id%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
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
	oConn.execute("DELETE FROM sdsk_prio_typ WHERE id = "& id &"")
	oConn.execute("DELETE FROM sdsk_prioitet WHERE priogrp = "& id &"")
	Response.redirect "sdsk_prio_typ.asp?menu=tok&shokselector=1"
	
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
		
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO sdsk_prio_typ (navn, editor, dato) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"')")
		else
		oConn.execute("UPDATE sdsk_prio_typ SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"' WHERE id = "&id&"")
		end if
		
		Response.redirect "sdsk_prio_typ.asp?menu=tok&shokselector=1"
		end if
		
	case "kopier"
		
		
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		strSQL = "SELECT navn, id FROM sdsk_prio_typ WHERE id = "& id
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
						
						
						'*** Kopierer grp ***
						strSQLkopi = "INSERT INTO sdsk_prio_typ (editor, dato, navn) VALUES "_
						&" ('"& strEditor &"', '"& strDato &"', '"& oRec("navn") &" (kopi)')"
						oConn.execute(strSQLkopi)
						
						strSQL3 = "SELECT id FROM sdsk_prio_typ WHERE id <> 0 ORDER BY id DESC"
						oRec3.open strSQL3, oConn, 3
						if not oRec3.EOF then
						
						nygrpId = oRec3("id")
						
						end if
						oRec3.close
						
						'** Henter prioiteter ***
						strSQL2 = "SELECT navn, responstid, faktor, kunweekend, advisering FROM sdsk_prioitet WHERE priogrp =" & oRec("id")
						oRec2.open strSQL2, oConn, 3
						
						while not oRec2.EOF 
						
						strNavn = oRec2("navn")
						rsptid = replace(oRec2("responstid"), ",", ".")
						dblfaktor = replace(oRec2("faktor"), ",", ".")
						intkunweekend = oRec2("kunweekend")
						dblAdvi = replace(oRec2("advisering"), ",", ".")
						
							
								strSQLkopi2 = "INSERT INTO sdsk_prioitet ( editor, dato, navn, responstid, faktor, "_
								&" kunweekend, advisering, priogrp) "_
								&" VALUES ('"& strEditor &"', '"& strDato &"', '"& strNavn &"', "& rsptid &", "& dblfaktor &", "_
								&" "& intkunweekend &", "& dblAdvi &", "& nygrpId &")"
								oConn.execute(strSQLkopi2)
						
						
						oRec2.movenext
						wend
						oRec2.close
					
					
		end if
		
		oRec.close	
		Response.redirect "sdsk_prio_typ.asp?menu=tok&shokselector=1"
		
		
		
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	
	else
	strSQL = "SELECT navn, editor, dato FROM sdsk_prio_typ WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
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
	<tr><form action="sdsk_prio_typ.asp?menu=tok&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
    	<td valign="top" colspan="2"><h3>Prioitet  - <%=varbroedkrumme%></h3></td>
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
		<td><b>Navn:</b></td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
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
	<%case else%>
	
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
	<h3>Prioiteter</h3><br>
	<a href="sdsk_prio_typ.asp?menu=tok&func=opret">Opret ny Prioitet <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a><br>
	<br>
	
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2><img src="../ill/tabel_top_left.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan=3 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top rowspan=2><img src="../ill/tabel_top_right.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td width="40" class=alt><b>Id</b></td>
		<td class=alt><b>Prioitet</b> (Antal Incidents)</td>
		<td width="50" class=alt>&nbsp;</td>
	</tr>
	<%
	
	strSQL = "SELECT s.id, s.navn, COUNT(p.id) AS antal FROM sdsk_prio_typ s "_
	&" LEFT JOIN sdsk_prioitet p ON (p.type = s.id) WHERE s.id <> 0 GROUP BY s.id ORDER BY s.navn"
	
	'Response.write strSQL
	'Response.flush
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	%>
	<tr>
		<td bgcolor="#003399" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="#ffffff" onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'#ffffff');">
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td><%=oRec("id")%></td>
		<td height="20"><a href="sdsk_prio_typ.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a>
		<%if oRec("id") = 1 then%>
		&nbsp;- Standard Priotetsklasse. 
		<%end if%>
		
		&nbsp;&nbsp;<b>(<%=oRec("antal")%>)</b>
		</td>
		<td>
		<%if oRec("id") <> 1 AND cint(oRec("antal")) = 0 then%>
		<a href="sdsk_prio_typ.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="" border="0"></a></td>
		<%end if%>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>
	<tr bgcolor="#5582D2">
		<td width="8" valign=top height=20 style="border-bottom:1px #003399 solid; border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=3 valign="top" style="border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top style="border-bottom:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>	
	</table>
	<br><br><br>
	<a href="Javascript:window.close()" class=vmenu>Luk vindue, og vend tilbage til kontrolpanel</a><br><br>
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
