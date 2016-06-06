<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/erp_func.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/konto_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
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
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20px; top:20px; visibility:visible;">
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> en <b>nøgletalskode</b>. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><a href="nogletalskoder.asp?menu=kon&func=sletok&id=<%=id%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM nogletalskoder WHERE id = "& id &"")
	Response.redirect "nogletalskoder.asp?menu=kon"
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
		
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
		
		
		ap_type = request("FM_ap_type")
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO nogletalskoder (navn, editor, dato, ap_type) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& ap_type &")")
		else
		oConn.execute("UPDATE nogletalskoder SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', ap_type = "& ap_type &" WHERE id = "&id&"")
		end if
		
		Response.redirect "nogletalskoder.asp?menu=kon"
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "Opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	ap_type = 0
	
	else
	strSQL = "SELECT navn, editor, dato, ap_type FROM nogletalskoder WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	ap_type = oRec("ap_type")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdaterpil" 
	end if
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20px; top:20px; visibility:visible;">
	<h4>Nøgletalskoder.</h4>
	<table cellspacing="0" cellpadding="2" border="0" width="600">
	<form action="nogletalskoder.asp?menu=tok&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
	<%if dbfunc = "dbred" then%>
	<tr>
		<td colspan="2" valign="bottom" style="height:30;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></td>
	</tr>
	<%end if%>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td width="120"><b>Navn:</b></td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" style="width:240px;"></td>
	</tr>
	<%
	select case cint(ap_type)
	case 1
	keySel1 = "SELECTED"
	keySel2 = ""
	keySel0 = ""
	case 2
	keySel2 = "SELECTED"
	keySel1 = ""
	keySel0 = ""
	case else
	keySel0 = "SELECTED"
	keySel2 = ""
	keySel1 = ""
	end select
	%>
	<tr>
		<td><b>Aktiv / Passiv type:</b></td>
		<td><select name="FM_ap_type">
		<option value="0" <%=keySel0 %>>Ingen</option>
	    <option value="1" <%=keySel1 %>>Aktiver</option>
		<option value="2" <%=keySel2 %>>Passiver</option>
		
		</select></td>
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
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	

	

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20px; top:20px; visibility:visible;">
	<h4>Nøgletalskoder.</h4>
	Opret og vedligehold Nøgletalskoder her. <br />
	Nøgletalskoderne bruges til resultatopgørelsen, således at konti bliver grupperet mellem aktiver og passiver.
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr>
    <td valign="top"><img src="../ill/blank.gif" width="490" height="1" alt="" border="0">
	<a href="nogletalskoder.asp?menu=kon&func=opret" class=vmenu>Opret ny <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a><br>
	<br>
	<table cellspacing="0" cellpadding="0" border="0" width="440" bgcolor="#ffffff">
	 <tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=2 valign="top"><img src="../ill/tabel_top.gif" width="524" height="1" alt="" border="0"></td>
		<td align=right valign=top rowspan=2><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan=2 class=alt><b>Nøgletalskoder.</b></td>
	</tr>
	<tr>
	<td width="5" style="border-left:1px #003399 solid;">&nbsp;</td>
	<td height="30" width=350><b>Navn</b> (type)</td>
	<td width="30">&nbsp;</td>
	<td width="5" style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<%
	strSQL = "SELECT id, navn, ap_type FROM nogletalskoder ORDER BY navn"
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	%>
	<tr>
		<td bgcolor="#cccccc" colspan="4" style="border-left:1px #003399 solid; border-right:1px #003399 solid;"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td height="20"><a href="nogletalskoder.asp?menu=kon&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a>
		
		<%select case oRec("ap_type")
		case 1
		ap_type = "Aktiver" 
		case 2
		ap_type = "Passiver"
		case else
		ap_type = "Ingen"
		end select
		%>
		&nbsp;(<%=ap_type %>)
		</td>
		<td><a href="nogletalskoder.asp?menu=kon&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="" border="0"></a></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>
	<tr><td colspan=4 style="border-bottom:1px #003399 solid; border-left:1px #003399 solid; border-right:1px #003399 solid;">&nbsp;</td></tr>	
	</table></td>
	</tr></table>
	

	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
