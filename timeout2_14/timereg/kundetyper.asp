<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%




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
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:150px; top:102px; visibility:visible;">
	<h4>Kundetyper / Segmenter</h4>
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> en status. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><a href="kundetyper.asp?menu=tok&func=sletok&id=<%=id%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
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
	oConn.execute("DELETE FROM kundetyper WHERE id = "& id &"")
	Response.redirect "kundetyper.asp?menu=tok&shokselector=1"
	
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
		
		intRabat = request("FM_rabat")
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO kundetyper (navn, editor, dato, rabat) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& intRabat &")")
		else
		oConn.execute("UPDATE kundetyper SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', rabat = "& intRabat &" WHERE id = "&id&"")
		end if
		
		Response.redirect "kundetyper.asp?menu=tok&shokselector=1"
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "Opret" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	intRabat = 0
	
	else
	strSQL = "SELECT navn, editor, rabat, dato FROM kundetyper WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	intRabat = oRec("rabat")
	
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdater" 
	end if
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->

    <%call menu_2014() %>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; visibility:visible;">
	<h4>Kundetyper / Segmenter</h4>
        <form action="kundetyper.asp?menu=tok&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
	<table cellspacing="0" cellpadding="2" border="0" width="100%">
	
    
	
	<%if dbfunc = "dbred" then%>
	<tr>
		<td colspan="2" valign="bottom" style="height:30;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></td>
	</tr>
	<%end if%>
	
	<tr>
		<td>Navn:</td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" style="width:350px;"></td>
	</tr>
	<tr>
		<td>Rabat:</td>
		<td>
		
		<%
        rSel0 = ""
        rSel10 = ""
        rSel15 = ""
        rSel25 = ""
        rSel30 = ""
        rSel40 = ""
        rSel50 = ""
        rSel60 = ""
        rSel75 = ""

		'**** husk også at rette på faktura oprettelse ***
		select case intRabat
		case 0
		rSel0 = "SELECTED"
        case 10
        rSel10 = "SELECTED"
        case 15
        rSel15 = "SELECTED"
        case 20
        rSel20 = "SELECTED"
        case 25
        rSel25 = "SELECTED"
        case 30
        rSel30 = "SELECTED"
        case 35
        rSel35 = "SELECTED"
        case 40
        rSel40 = "SELECTED"
        case 50
        rSel50 = "SELECTED"
        case 60
        rSel60 = "SELECTED"
        case 75
        rSel75 = "SELECTED"
		end select
		%>
            <select id="FM_rabat" name="FM_rabat">
                <option value=0  <%=rSel0%>>0%</option>
                <option value=10 <%=rSel10%>>10%</option>
                <option value=15 <%=rSel15%>>15%</option>
                <option value=20 <%=rSel20%>>20%</option>
                <option value=25 <%=rSel25%>>25%</option>
                <option value=30 <%=rSel30%>>30%</option>
                <option value=35 <%=rSel35%>>35%</option>
                <option value=40 <%=rSel40%>>40%</option>
                <option value=50 <%=rSel50%>>50%</option>
                <option value=60 <%=rSel60%>>60%</option>
                <option value=75 <%=rSel75%>>75%</option>
             </select>
        </td>
	</tr>
	
	<tr>
		<td colspan="2" align="right"><br /><br /><input type="submit" value="<%=varSubVal%>" /></td>
	</tr>
	
	</table>
            </form>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%case else%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	

    <%call menu_2014() %>

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:82px; visibility:visible;">
	<h4>Kundetyper / Segmenter</h4>
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr>
    <td valign="top">
	Sortér efter <a href="kundetyper.asp?menu=tok&sort=navn">Navn</a> eller <a href="kundetyper.asp?menu=tok&sort=nr">Id nr.</a>
	<img src="../ill/blank.gif" width="320" height="1" alt="" border="0">
	<a href="kundetyper.asp?menu=tok&func=opret">Opret ny type <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a><br>
	<br>
	</td></tr>
	</table>
	
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2><img src="../ill/tabel_top_left.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan=4 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top rowspan=2><img src="../ill/tabel_top_right.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td width="50" class=alt><b>Id</b></td>
		<td width="300" class=alt><b>Kundetype</b></td>
		<td class=alt><b>Rabat %</b></td>
		<td width="50" class=alt>&nbsp;</td>
	</tr>
	<%
	sort = Request("sort")
	if sort = "navn" then
	strSQL = "SELECT id, navn, rabat FROM kundetyper ORDER BY navn"
	else
	strSQL = "SELECT id, navn, rabat FROM kundetyper ORDER BY id"
	end if
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	%>
	<tr>
		<td bgcolor="#CCCCCC" colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="#ffffff">
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td><%=oRec("id")%></td>
		<td height="20">
        <%if oRec("id") <> 200 then %>
        <a href="kundetyper.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a>
        <%else %>
        <%=oRec("navn")%> (system)
        <%end if %>
        </td>
		<td><%=oRec("rabat") %>%</td>
		<td>
        <%if oRec("id") <> 200 then %>
        <a href="kundetyper.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="" border="0"></a>
        <%end if %>
        &nbsp;</td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>
	<tr bgcolor="#5582D2">
		<td width="8" valign=top height=20 style="border-bottom:1px #003399 solid; border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=4 valign="top" style="border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top style="border-bottom:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>	
	</table>
	
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
