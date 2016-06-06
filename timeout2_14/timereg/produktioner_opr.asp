<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/mat_func.asp"-->
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
	
	if len(id) <> 0 then
	id = id
	else
	id = 0
	end if
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:60; visibility:visible;">
	<img src="../ill/header_progrp.gif" width="600" height="50" alt="" border="0">
	<br><br><br>
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> en produktgruppe. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><a href="produktioner.asp?menu=mat&func=sletok&id=<%=id%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
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
	oConn.execute("DELETE FROM produktioner WHERE id = "& id &"")
	Response.redirect "produktioner.asp?menu=mat&shokselector=1"
	
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
		intGtype = request("FM_gtype")
		intPic_1 = request("FM_pic_1")
		strBesk = request("FM_besk")
		
		if len(request("FM_z")) <> 0 then
		intZ = request("FM_z")
		else
		intZ = 0
		end if
		
		if len(request("FM_left")) <> 0 then
		intLeft = request("FM_left")
		else
		intLeft = 0
		end if
		
		if len(request("FM_top")) <> 0 then
		intTop = request("FM_top")
		else
		intTop = 0
		end if
		
		
		if len(request("FM_width")) <> 0 then
		intW = request("FM_width")
		else
		intW = 0
		end if
		
		if len(request("FM_height")) <> 0 then
		intH = request("FM_height")
		else
		intH = 0
		end if
		
		if len(request("FM_left2")) <> 0 then
		intLeft2 = request("FM_left2")
		else
		intLeft2 = 0
		end if
		
		if len(request("FM_top2")) <> 0 then
		intTop2 = request("FM_top2")
		else
		intTop2 = 0
		end if
		
		
		if len(request("FM_width2")) <> 0 then
		intW2 = request("FM_width2")
		else
		intW2 = 0
		end if
		
		if len(request("FM_height2")) <> 0 then
		intH2 = request("FM_height2")
		else
		intH2 = 0
		end if
		
		if len(request("FM_sortorder")) <> 0 then
		intSortorder = request("FM_sortorder")
		else
		intSortorder = 0
		end if
		
		
		
				
		if func = "dbopr" then
		oConn.execute("INSERT INTO produktioner (navn, editor, dato, gruppetype, pic_1, "_
		&" besk, zindex, leftpx, toppx, wpx, hpx, leftpx2, toppx2, wpx2, hpx2, sortorder) "_
		&" VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& intGtype &", "& intPic_1 &", "_
		&" '"& strBesk &"', "& intZ &", "& intLeft &", "& intTop &", "& intW &", "& intH &""_
		&", "& intLeft2 &", "& intTop2 &", "& intW2 &", "& intH2 &", "& intSortorder &")")
			
			lastId = 0
			
			'*** Finder netop oprettede id ***
			strSQL = "SELECT id FROM produktioner ORDER BY id DESC"
			oRec.open strSQL, oConn, 3 
			if not oRec.EOF then
				lastId = oRec("id")
			end if
			oRec.close 
		
		
		
		else
		oConn.execute("UPDATE produktioner SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', "_
		&" gruppetype = "& intGtype &", pic_1 = "& intPic_1 &", besk = '"& strBesk &"', zindex = "& intZ &", "_
		&" leftpx = "& intLeft &", toppx = "& intTop &", wpx = "& intW &", hpx = "& intH &", "_
		&" leftpx2 = "& intLeft2 &", toppx2 = "& intTop2 &", wpx2 = "& intW2 &", hpx2 = "& intH2 &", sortorder = "& intSortorder &" WHERE id = "&id&"")
		
			lastId = id
		end if
		
		'*** produktionerelationer ***
		oConn.execute("DELETE FROM produktion_mat_rel WHERE mainprodrgp = "& lastId &"")
		
		strproduktion_mat_rel = split(request("FM_produktion_mat_rel"), ",")
		For b = 0 to Ubound(strproduktion_mat_rel)
			oConn.execute("INSERT INTO produktion_mat_rel (mainprodrgp, relprodrgp) VALUES ("& lastId &", "& strproduktion_mat_rel(b) &")")
		Next
		
		
		Response.redirect "produktioner.asp?menu=mat&shokselector=1"
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny produktgruppe."
	dbfunc = "dbopr"
	intGtype = request("gtype") '2
	strPic_1 = 1
	strBesk = ""
	intZ = 0
	intLeft = 0
	intTop = 0
	intW = 0
	intH = 0
	intLeft2 = 0
	intTop2 = 0
	intW2 = 0
	intH2 = 0
	intSortorder = 0
	
	
	else
	strSQL = "SELECT navn, editor, dato, gruppetype, pic_1, besk, zindex, leftpx, toppx, wpx, hpx, leftpx2, toppx2, wpx2, hpx2, sortorder FROM produktioner WHERE id=" & id
	
	'Response.write strSQL
	'Response.flush
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	intGtype = oRec("gruppetype")
	strPic_1 = oRec("pic_1")
	strBesk = oRec("besk")
	intZ = oRec("zindex")
	intTop = oRec("toppx")
	intLeft = oRec("leftpx")
	intW = oRec("wpx")
	intH = oRec("hpx")
	intTop2 = oRec("toppx2")
	intLeft2 = oRec("leftpx2")
	intW2 = oRec("wpx2")
	intH2 = oRec("hpx2")
	intSortorder = oRec("sortorder")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:60; visibility:visible;">
	<img src="../ill/header_progrp.gif" width="600" height="50" alt="" border="0">
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr><form action="produktioner.asp?menu=mat&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
    	<td valign="top" colspan="2"><br><font class="pageheader"><%=varbroedkrumme%></font></td>
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
		<td align=right>Navn:&nbsp;&nbsp;</td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td align=right valign=top>Beskrivelse: (kort)&nbsp;&nbsp;</td>
		<td><textarea cols="30" rows="3" name="FM_besk"><%=strBesk%></textarea></td>
	</tr>
	<tr>
		<td align=right><br>Gruppetype:&nbsp;&nbsp;</td>
		<td><br>
			<%select case intGtype
			case 1
			gtype = "Hovedgruppe"
			case 2
			gtype = "Undergruppe 1"
			case else
			gtype = "Undergruppe 2 (Billeder fra denne gruppe vises på web.)"
			end select%>
		<%=gtype%>	
		<input type="hidden" name="FM_gtype" id="FM_gtype" value="<%=intGtype%>">
		
		
		
		<!--
		<select name="FM_gtype" id="FM_gtype">
			
			<select case intGtype
			case 1
			sel1 = "SELECTED"
			sel2 = ""
			sel3 = ""
			case 2
			sel1 = ""
			sel2 = "SELECTED"
			sel3 = ""
			case else
			sel1 = ""
			sel2 = ""
			sel3 = "SELECTED"
			end select%>
			<option value="1" <=sel1%>>Hovedgruppe</option>
			<option value="2" <=sel2%>>Undergruppe 1</option>
			<option value="3" <=sel3%>>Undergruppe 2</option>
	</select>-->
	
	</td>
	</tr>
	<tr>
		<td align=right><br>
		Rækkefølge:&nbsp;&nbsp;</td>
		<td><br><input type="text" name="FM_sortorder" value="<%=intSortorder%>" size="10" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		&nbsp;</td>
	</tr>
	<!--<tr>
		<td align=right><br>
		Z-index:&nbsp;&nbsp;</td>
		<td><br><input type="text" name="FM_z" value="<=intZ%>" size="10" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td colspan=2 style="padding-left:76px;"><u>Set forfra:</u>
		Left:&nbsp;&nbsp;<input type="text" name="FM_left" value="<=intLeft%>" size="2" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"> px.
		&nbsp;&nbsp;&nbsp;&nbsp;Top:&nbsp;&nbsp;<input type="text" name="FM_top" value="<=intTop%>" size="2" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"> px. 
		&nbsp;&nbsp;&nbsp;&nbsp;Width:&nbsp;&nbsp;<input type="text" name="FM_width" value="<=intW%>" size="2" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"> px.
		&nbsp;&nbsp;&nbsp;&nbsp;Height:&nbsp;&nbsp;<input type="text" name="FM_height" value="<=intH%>" size="2" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"> px. 
		</td>
	</tr>
	<tr>
		<td colspan=2 style="padding-left:76px;"><u>Set fra siden:</u>
		Left:&nbsp;&nbsp;<input type="text" name="FM_left2" value="<=intLeft2%>" size="2" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"> px.
		&nbsp;&nbsp;&nbsp;&nbsp;Top:&nbsp;&nbsp;<input type="text" name="FM_top2" value="<=intTop2%>" size="2" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"> px. 
		&nbsp;&nbsp;&nbsp;&nbsp;Width:&nbsp;&nbsp;<input type="text" name="FM_width2" value="<=intW2%>" size="2" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"> px.
		&nbsp;&nbsp;&nbsp;&nbsp;Height:&nbsp;&nbsp;<input type="text" name="FM_height2" value="<=intH2%>" size="2" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"> px. 
		<br><br>&nbsp;</td>
	</tr>-->
	
	<%
	call grafik(1, strPic_1, 9, "Prod.gruppe billede")
	
	if intGtype < 3 then
	%>
	<tr>
		<td colspan=2 valign=top><br><b>Relationer:</b> <br>Hvilke <u>underliggende</u> produktioner skal vises når denne (<b><%=strNavn%></b>) produktgruppe vises på web?
		<!--<u>Hvis dette er en hovedgruppe:</u><br>
		Vælg hvilke undergrupper der skal vises når denne hovegruppe vises på websitet.<br><br>
		<u>Hvis dette er en undergruppe:</u><br>
		Vælg hvilken (Maks 1) gruppe som denne undergruppe skal være i familie med.
		Eks: Kalecher basis farver er relateret (i famile med) Kalecher special farver.--></td>
	</tr>
	<tr>
		<td colspan=2 valign=top><br>
		<%
		gtypeplus_1 = intGtype + 1
		strSQL = "SELECT id, navn FROM produktioner WHERE id <> "& id &" AND gruppetype = "& gtypeplus_1 &" ORDER BY navn"
		oRec.open strSQL, oConn, 3 
		t = 1
		while not oRec.EOF 
		select case right(t, 1)
		case 5, 0
		Response.write "<br>"
		end select
			
			foundrel = 0
			strSQL2 = "SELECT id FROM produktion_mat_rel WHERE mainprodrgp = "& id &" AND relprodrgp = "& oRec("id")
			oRec2.open strSQL2, oConn, 3 
			if not oRec2.EOF then
			foundrel = 1
			end if
			oRec2.close 
			
			
			if foundrel = 1 then
			chkthis = "checked"
			else
			chkthis = ""
			end if%>
		<input type="checkbox" name="FM_produktion_mat_rel" id="FM_produktion_mat_rel" value="<%=oRec("id")%>" <%=chkthis%>><%=oRec("navn")%>&nbsp;&nbsp;
		<%t = t + 1
		oRec.movenext
		wend
		oRec.close 
		%><br>
		</td>
	</tr>
	<%end if%>
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
	
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	</script>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:80; visibility:visible;">
	<img src="../ill/header_produktioner.gif" width="755" height="48" alt="" border="0">
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr>
    <td valign="top">
	<br>Sortér efter <a href="produktioner.asp?menu=mat&sort=navn">Navn</a> eller <a href="produktioner.asp?menu=mat&sort=nr">Gruppetype</a>
	<img src="../ill/blank.gif" width="200" height="1" alt="" border="0">
	<!--<a href="produktioner.asp?menu=mat&func=opret">Opret ny produktgruppe <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>--><br>
	<br>&nbsp;
	</td>
	</tr>
	</table>
	
	Opret ny:
	<b>
	<a href="produktioner.asp?menu=mat&func=opret&gtype=1"><font color="limegreen">Hovedgruppe <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></font></a> &nbsp;&nbsp;|&nbsp;&nbsp;
	<a href="produktioner.asp?menu=mat&func=opret&gtype=2"><font color="orange">Undergrp. 1 <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></font></a> &nbsp;&nbsp;|&nbsp;&nbsp;
	<a href="produktioner.asp?menu=mat&func=opret&gtype=3"><font color="gray">Undergrp. 2 <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></font></a> </b>
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2><img src="../ill/tabel_top_left.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan=4 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top rowspan=2><img src="../ill/tabel_top_right.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
	<td width="50" class=alt><b>Id</b></td>
	<td width="300" class=alt><b>Produktgruppe navn</b></td>
	<td class=alt><b>Under grp. / Er med i</b></td>
	<td class=alt>&nbsp;</td>
	</tr>
	<%
	sort = Request("sort")
	if sort = "navn" then
	strSQL = "SELECT count(rel.id) AS antalrel, p.id, p.navn, gruppetype FROM produktioner p LEFT JOIN produktion_mat_rel rel ON (mainprodrgp = p.id) GROUP BY p.id ORDER BY navn"
	else
	strSQL = "SELECT count(rel.id) AS antalrel, p.id, p.navn, gruppetype FROM produktioner p LEFT JOIN produktion_mat_rel rel ON (mainprodrgp = p.id) GROUP BY p.id ORDER BY gruppetype, navn"
	end if
	
	'Response.write strSQL
	'Response.flush
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	%>
	<tr>
		<td bgcolor="#003399" colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="#eff3ff" onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'#eff3ff');">
		<td height=20 style="border-left:1px #003399 solid;">&nbsp;</td>
		<td><%=oRec("id")%></td>
		<td><a href="produktioner.asp?menu=mat&func=red&id=<%=oRec("id")%>">
		<%select case oRec("gruppetype")
		case 1
		fcol = "limegreen"
		case 2
		fcol = "orange"
		case else
		fcol = "gray"
		end select%>
		<font color="<%=fcol%>">
		<%=oRec("navn")%></font> 
		</a></td>
		<td>(<%=oRec("antalrel")%>)
		<%
		antalsubrel = 0
		strSQL2 = "SELECT count(id) AS antalsubrel FROM produktion_mat_rel WHERE relprodrgp = "& oRec("id")
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
			antalsubrel = oRec2("antalsubrel")
		end if
		oRec2.close 
		
		Response.write "&nbsp;&nbsp;("& antalsubrel & ")"
		%></td>
		<td>
		<%
		if oRec("antalrel") = "0" AND antalsubrel = "0" then%>
		<a href="produktioner.asp?menu=mat&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="" border="0"></a>
		<%else%>
		&nbsp;
		<%end if%></td>
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
	<br><br><br>
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
