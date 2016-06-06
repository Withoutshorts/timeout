<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/mat_func.asp"-->

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
	
	rdir = request("rdir")
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:210; top:150; background-color:#ffffff; visibility:visible; padding:20px;"">
	<h4>Leverandør- og Service -partnere</h4>
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> en Leverandør/Servicepartner. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><a href="leverand.asp?menu=mat&func=sletok&id=<%=id%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
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
	oConn.execute("DELETE FROM leverand WHERE id = "& id &"")
	Response.redirect "leverand.asp?menu=mat&shokselector=1"
	
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
		
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		strAdr = SQLBless(request("FM_adr"))
		strPostnr = request("FM_postnr")
		strCity = request("FM_city")
		strLand = request("FM_land")
		strEmail = request("FM_email")
		strTlf = request("FM_tlf")
		strFax = request("FM_fax")
		strLevnr = request("FM_lnr")
		strBesk = request("FM_besk")
		
		intType = request("FM_type")
		
		strKpers1 = SQLBless(request("FM_kpers1"))
		strKperstlf1 = request("FM_kperstlf1")
		strKpersemail1 = request("FM_kpersemail1")
		
		strKpers2 = SQLBless(request("FM_kpers2"))
		strKperstlf2 = request("FM_kperstlf2")
		strKpersemail2 = request("FM_kpersemail2")
		
		
		
		
		if func = "dbopr" then
		strSQL = "INSERT INTO leverand (editor, dato, navn, "_
		&" adresse, postnr, city, land, email, tlf, fax, besk, levnr, type, "_
		&" kpers1, kperstlf1, kpersemail1, kpers2, kperstlf2, kpersemail2 "_
		&" ) VALUES "_
		&" ('"& strEditor &"', '"& strDato &"', '"& strNavn &"', "_
		&" '"&strAdr&"', '"&strPostnr&"', '"&strCity&"', '"&strLand&"', '"&strEmail&"', "_
		&" '"&strTlf&"', '"&strFax&"', '"&strBesk&"', "_
		&" '"&strLevnr&"', "&intType&", '"&strKpers1&"', '"&strKperstlf1&"', "_
		&" '"&strKpersemail1&"', '"&strKpers2 &"', "_
		&" '"&strKperstlf2 &"', '"&strKpersemail2 &"')"
		
		else
		
		strSQL = "UPDATE leverand SET editor = '"& strEditor &"', dato = '"& strDato &"', navn = '"& strNavn &"', "_
		&" adresse = '"&strAdr&"', postnr = '"&strPostnr&"', city = '"&strCity&"', land = '"&strLand&"', "_
		&" email = '"&strEmail&"', tlf = '"&strTlf&"' , fax = '"&strFax&"', besk = '"&strBesk&"', levnr = '"&strLevnr&"', type = "&intType&", "_
		&" kpers1 = '"&strKpers1&"', kperstlf1 = '"&strKperstlf1&"', kpersemail1 = '"&strKpersemail1&"', "_
		&" kpers2 = '"&strKpers2 &"', kperstlf2 = '"&strKperstlf2 &"', kpersemail2 = '"&strKpersemail2 &"'"_
		&" WHERE id = "& id
		end if
		
		oConn.execute(strSQL)
		
		
		if rdir = "mat" then
		Response.redirect "materialer.asp?menu=mat"
		else
		Response.redirect "leverand.asp?menu=mat&shokselector=1"
		end if
		
		end if '** validering
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "Opret" 
	varbroedkrumme = "Opret ny."
	dbfunc = "dbopr"
	
	else
	strSQL = "SELECT editor, dato, navn, "_
	&" adresse, postnr, city, land, email, tlf, fax, besk, levnr, type, "_
	&" kpers1, kperstlf1, kpersemail1, kpers2, kperstlf2, kpersemail2 "_
	&" FROM leverand WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
		strNavn = oRec("navn")
		strDato = oRec("dato")
		strEditor = oRec("editor")
		
		strAdr = oRec("adresse")
		strPostnr = oRec("postnr")
		strCity = oRec("city")
		strLand = oRec("land")
		strEmail = oRec("email")
		strTlf = oRec("tlf")
		strFax = oRec("fax")
		strLevnr = oRec("levnr")
		strBesk = oRec("besk")
		
		intType = oRec("type")
		
		strKpers1 = oRec("kpers1")
		strKperstlf1 = oRec("kperstlf1")
		strKpersemail1 = oRec("kpersemail1")
		
		strKpers2 = oRec("kpers2")
		strKperstlf2 = oRec("kperstlf2")
		strKpersemail2 = oRec("kpersemail2")
		
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
	<div id="sindhold" style="position:absolute; left:90px; top:102px; visibility:visible; width:600px; padding:20px; background-color:#ffffff;">
	<h4>Leverandør- og Service -partnere <span style="font-size:9px;"> - <%=varbroedkrumme%></span></h4>
    
        
    <form action="leverand.asp?menu=mat&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
	<input type="hidden" name="rdir" value="<%=rdir%>">

	<table cellspacing="0" cellpadding="2" border="0" width="100%">
	
	
    
	<%if dbfunc = "dbred" then%>
	<tr>
		
		<td colspan="2" valign="bottom" style="height:30;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></td>
		
	</tr>
	<%end if%>
	
	<tr>
		
		<td><font color=red>*</font> Navn:</td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" size="50" ></td>
	</tr>
	<tr>
		<td>Email:</td>
		<td><input type="text" name="FM_email" value="<%=strEmail%>" size="25" ></td>
	</tr>
	<tr>
		<td>Leverandør nr:</td>
		<td><input type="text" name="FM_lnr" value="<%=strLevnr%>" size="10" ></td>
	</tr>
	<tr>
		<td>Leverandør / Servicepartner:</td>
		<td><select name="FM_type" id="FM_type">
		<%select case intType 
		case 1
		sel1 = "SELECTED"
		sel2 = ""
		sel3 = ""
		case 2
		sel1 = ""
		sel2 = "SELECTED"
		sel3 = ""
		case 3
		sel1 = ""
		sel2 = ""
		sel3 = "SELECTED"
		case else
		sel1 = "SELECTED"
		sel2 = ""
		sel3 = ""
		end select%>
		<option value="1" <%=sel1%>>Leverandør</option>
		<option value="2" <%=sel2%>>Servicepartner</option>
		<option value="3" <%=sel3%>>Leverandør & Servicepartner</option>
		</select></td>
	</tr>
	<tr>
		<td>Adresse:</td>
		<td><input type="text" name="FM_adr" value="<%=strAdr%>" size="40" ></td>
	</tr>
	<tr>
		<td>Postnr:</td>
		<td><input type="text" name="FM_postnr" value="<%=strPostnr%>" size="4" ></td>
	</tr>
	<tr>
		<td>By:</td>
		<td><input type="text" name="FM_city" value="<%=strCity%>" size="40" ></td>
	</tr>
	<tr>
		<td>Land:</td>
		<td><select name="FM_land" >
		<%if func = "red" then%>
		<option checked><%=strLand%></option>
		<%else%>
		<option checked>Danmark</option>
		<%end if%>
		<!--#include file="inc/inc_option_land.asp"-->
		</select>
	</tr>
	<tr>
		<td>Tlf:</td>
		<td><input type="text" name="FM_tlf" value="<%=strTlf%>" size="12" ></td>
	</tr>
	<tr>
		<td>Fax:</td>
		<td><input type="text" name="FM_fax" value="<%=strFax%>" size="12" ></td>
	</tr>
	<tr>
		<td colspan="2"><b>Bemærkninger:</b><br>
		<textarea cols="60" rows="5" name="FM_besk"><%=strBesk%></textarea></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td><img src="../ill/ac0026-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Kontaktperson 1:</b></td>
		<td><input type="text" name="FM_kpers1" value="<%=strKpers1%>" size="30" ></td>
	</tr>
	<tr>
		<td>Email 1:</td>
		<td><input type="text" name="FM_kpersemail1" value="<%=strKpersemail1%>" size="25" ></td>
	</tr>
	<tr>
		<td>Tlf 1:</td>
		<td><input type="text" name="FM_kperstlf1" value="<%=strKperstlf1%>" size="12" ></td>
	</tr>
	<tr>
		<td><img src="../ill/ac0057-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Kontaktperson 2:</b></td>
		<td><input type="text" name="FM_kpers2" value="<%=strKpers2%>" size="30" ></td>
	</tr>
	<tr>
		<td>Email 2:</td>
		<td><input type="text" name="FM_kpersemail2" value="<%=strKpersemail2%>" size="25" ></td>
	</tr>
	<tr>
		<td>Tlf 2:</td>
		<td><input type="text" name="FM_kperstlf2" value="<%=strKperstlf2%>" size="12" ></td>
	</tr>
	
	<tr>
		<td colspan="2" align=right><br><br><input type="submit" value="<%=varSubVal%> >>" /><br><br>&nbsp;</td>
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
	
<!--
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	
	<%'call tsamainmenu(11)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	'if showonejob <> 1 then
	'	call mattopmenu()
	'end if
	%>
	</div>
-->

    <%call menu_2014() %>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; visibility:visible; background-color:#ffffff; width:800px; padding:20px;">
	<h4>Leverandør- og Service -partnere</h4>
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
    <td valign="top" align=right><a href="leverand.asp?menu=mat&func=opret">Opret ny leverandør / servicepartner <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a></td>
	</tr>
	</table>
	
	<table cellspacing="0" cellpadding="2" border="0" width="100%">

	<tr bgcolor="#5582D2">
	<td width="90" class=alt><b>Lev. nr.</b></td>
	<td width="300"><a href="leverand.asp?menu=mat&sort=navn"  class=alt>Navn</a></td>
	<td class=alt><a href="leverand.asp?menu=mat&sort=lev"  class=alt>Leverand. / Servicepart.</a></td>
	<td class=alt>&nbsp;</td>
	</tr>
	<%
	sort = Request("sort")
	if sort = "navn" then
	strSQL = "SELECT id, navn, type, levnr FROM leverand WHERE type = 1 OR type = 2 OR type = 3 ORDER BY navn"
	else
	strSQL = "SELECT id, navn, type, levnr FROM leverand WHERE type = 1 OR type = 2 OR type = 3 ORDER BY type, navn"
	end if
	
	'Response.write strSQL
	'Response.flush
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
	antalordrer2 = 0
	%>
	
	<tr bgcolor="#eff3ff" onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'#eff3ff');">
	
		<td style="border-bottom:1px #999999 solid;"><%=oRec("levnr")%></td>
		<td style="border-bottom:1px #999999 solid;"><a href="leverand.asp?menu=mat&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
		<td style="border-bottom:1px #999999 solid;">
		<%select case oRec("type")
		case 1
		strType = "Leverand."
		case 3
		strType = "Lev. & Ser.part."
		case else
		strType = "Servicepart."
		end select
		
		Response.write strType%>
		</td>
		<td style="border-bottom:1px #999999 solid;"><a href="leverand.asp?menu=mat&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="" border="0"></a></td>
	
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>	
	
	</table>
	<br><br><br>
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
