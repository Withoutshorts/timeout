<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/kontakter_func.asp"-->
<% 'Include the DevEdit class file %>
<!-- #INCLUDE file="de/class.devedit.asp" -->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	thisfile = "infobase"
	func = request("func")
	kview = request("usekview")
	'uselogin = request("uselogin")
	
	
	if len(request("kontaktid")) <> 0 then
	kontaktid = request("kontaktid")
	else
	kontaktid = 0
	end if
	
	
	if len(request("id")) <> 0 then
	id = request("id")
	else
	id = 0
	end if
	
	if len(request("infid")) <> 0 then
	infid = request("infid")
	else
	infid = 0
	end if
	
	select case func
	case "slet"
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<div id="sindhold" style="position:absolute; left:210; top:180; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#FFFFF1"">
	<tr>
	    <td style="!border: 1px; border-color: #990200; border-style: solid; padding-right:10; padding-left:10; padding-top:5; padding-bottom:5;"><img src="../ill/alert.gif" width="44" height="45" alt="" border="0"><br>
		Du er ved at <b>slette</b> et infobase punkt. <br>
		Du vil samtidig slette alle underliggende infobase punkter.<br>
		<br>Er du sikker på at du vil slette dette infobase punkt?<br><br>
		<a href="infobase.asp?menu=kund&id=<%=id%>&func=sletok&infid=<%=infid%>"><font color=green>Ja</font></a><img src="../ill/blank.gif" width="200" height="1" alt="" border="0"><a href="javascript:history.back()"><font color=red>Nej, jeg vil ikke slette dette punkt!</font></a></td>
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
	
	oConn.execute("DELETE FROM infobase WHERE id = "& infid &" OR parent = "& infid &"")
	Response.redirect "infobase.asp?menu=kund&id="&id
	
	
	case "dbopr", "dbred"
	
	function SQLBless(s)
	dim tmp
	tmp = s
	tmp = replace(tmp, "'", "''")
	SQLBless = tmp
	end function
	
	
	'**** sætter værdier ***
	
	
	lasteredigerval = request("lastredigervalue")
	
	if func = "dbopr" then
	strNavn = SQLBless(request("FM_navn"))
	strBesk = request("FM_besk") 
	else
	strNavn = SQLBless(request("navn_"&lasteredigerval))
	strBesk = request("FM_besk_"&lasteredigerval) 'myDE.GetValue(true)
	end if
	
	
	strEditor = session("user")
	strDato = year(date)&"/"&month(date)&"/"&day(date) '&"/"&time(now)
	intParent = request("FM_parent")
	intOprParent = request("FM_oprparent")
	intKundeid = request("FM_kundeid")
	
	
	'**** Hvis punktet bliver oprettet i den globale kun er intKundeid = 0 ***
	if len(request("FM_global")) <> 0 OR intKundeid = 0 then
	intGlobal = 1
	else
	intGlobal = 0
	end if
	
	'**** Validering ****
	if len(trim(strNavn)) = 0 then 
		
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
		<%
		errortype = 8
		call showError(errortype)
	
	else
	
	if func = "dbopr" then
	strSQL3 = "INSERT INTO infobase (navn, editor, dato, parent, beskrivelse, kuid, globalinfobase) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& intParent &", '"& strBesk &"', "& intKundeid &", "& intGlobal &")"
	else
	strSQL3 = "UPDATE infobase SET navn = '"& strNavn &"',  editor = '"& strEditor &"', dato = '"& strDato &"', beskrivelse = '"& strBesk &"' WHERE id=" & lasteredigerval
	end if
	
	'Response.write strSQL3
	
	oConn.execute(strSQL3)
	
	Response.redirect "infobase.asp?menu=kund&id="&id
	
	
	end if '** Validering
	
		
	case "opr", "red"
	if func = "red" then
		strSQL = "SELECT id, navn, beskrivelse, kuid, parent, globalinfobase FROM infobase WHERE id = " & infid
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		
		infid = oRec("id")
		strNavn = oRec("navn")
		strBesk = oRec("beskrivelse")
		intKuid = oRec("kuid")
		intParent = oRec("parent") 
		intGlobal = oRec("globalinfobase")
		dbfunc = "dbred" 
		
		end if
		oRec.close
	else
		
		infid = 0
		strNavn = ""
		strBesk = ""
		intKuid = request("id")
		intParent = request("parent")
		intGlobal = 0
		dbfunc = "dbopr" 
		
	end if
	%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	<%call tsamainmenu(5)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call kundtopmenu()
	%>
	</div>
		
	<!--#include file="inc/infobase_inc.asp"-->
	
	<div id="sindhold" style="position:absolute; left:20; top:132; visibility:visible;">
	<h3>Infobase</h3>
	<table cellspacing="0" cellpadding="0" border="0" width="400">
	<form method="POST" action="infobase.asp">
	<input type="hidden" name="infid" value="<%=infid%>">
	<input type="hidden" name="id" value="<%=id%>">
	<input type="hidden" name="func" value="<%=dbfunc%>">
	<input type="hidden" name="FM_parent" value="<%=intParent%>">
	<tr><td><b>Infobase:</b></td><td style="padding-left:5px;">
	<%
		strSQL = "SELECT kkundenavn FROM kunder WHERE kid =" & intKuid
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		Response.write oRec("kkundenavn")
		else
		Response.write "Global"
		end if
		oRec.close
	%>
	<input type="hidden" name="FM_kundeid" value="<%=intKuid%>">
	</td></tr>
	<tr>
		<td><b>Navn:</b></td><td style="padding-left:5px;"><input type="text" name="FM_navn" size="30" value="<%=strNavn%>"></td></tr>
	<tr>
	<tr><td colspan=2><b>Tekst:</b><br>
	<textarea cols="57" rows="20" name="FM_besk" id="FM_besk" style="border:1px #003399 solid;"></textarea>
	</td></tr>
	<tr><td colspan=2 align=center valign=top><br><br><input type="image" src="../ill/opretpil.gif"></td></tr>
	</form>
	</table>
	<br><br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	case "eks"
	'******************************************************************************
	'Print hele infobasen / eksport
	'*******************************************************************************
	
	%>
	<html>
	<head>
		<title>TimeOut 2.1</title>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
	</head>
	<body topmargin="0" leftmargin="0" class="regular">
	<table cellspacing="0" cellpadding="0" border="0" width="880">
	<tr>
		<td bgcolor="#003399" width="650"><img src="../ill/logo_topbar_print.gif" alt="" border="0"></td>
		<td bgcolor="#FFFFFF" align=right><a href="javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0">&nbsp;Tilbage</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"><a href="javascript:window.print()"><img src="../ill/print_xp.gif" width="28" height="30" alt="" border="0">&nbsp;Print</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"></td>
	</tr>
	</table>
	<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#FFFFFF">
	<tr>
		<td style="padding-left : 20px; padding-top:10;" valign="top">

	<%
	if id <> 0 then
		strSQL = "SELECT kkundenavn FROM kunder WHERE kid =" & id
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		varinfobasenavn = oRec("kkundenavn")
		end if
		oRec.close
	else
	varinfobasenavn = "Global"
	end if
	
	Response.write "<br>Infobase: "
	Response.write "<b>"& varinfobasenavn &"</b>"
	Response.write "<br><br>" 
	
	if cint(id) <> 0 then
	kkri = "kuid = "& id &" AND parent = 0" 
	else
	kkri = "(globalinfobase = 1) AND parent = 0"
	end if
	
	
	
	Response.write "<br><br><br>"
	Response.write "<font size=4><b>Indholds fortegnelse:</b></font><br><img src='../ill/hrstreg_600.gif' width='500' height='1' alt='' border='0'><br>"
	
	
	'**** Indholdsfortegnelse ****
	strSQL = "SELECT id, navn, beskrivelse, kuid, parent, globalinfobase FROM infobase WHERE "& kkri &" ORDER BY navn"
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	Response.write "<br><b>"
	Response.write oRec("navn") &"</b><br>"
	
		strSQL2 = "SELECT id, navn, beskrivelse, parent, globalinfobase FROM infobase WHERE parent = " & oRec("id")
		oRec2.open strSQL2, oConn, 3
		while not oRec2.EOF
		
			Response.write "<img src='../ill/blank.gif' width='40' height='1' alt='' border='0'>"& oRec2("navn") &"<br>"
			
			
			strSQL3 = "SELECT id, navn, beskrivelse, parent, globalinfobase FROM infobase WHERE parent = " & oRec2("id")
			oRec3.open strSQL3, oConn, 3
			while not oRec3.EOF
		
				Response.write "<img src='../ill/blank.gif' width='80' height='1' alt='' border='0'>"& oRec3("navn") &"<br>"
				
			oRec3.movenext
			wend
			oRec3.close
		
			
		oRec2.movenext
		wend
		oRec2.close
	
	
	oRec.movenext
	wend
	oRec.close
	
	Response.write "<br><br><img src='../ill/hrstreg_600.gif' width='500' height='1' alt='' border='0'><br><br><br><br><br><br>"
	Response.write "</td></tr></table>"
	
	
	'**** niveau 1 ****
	strSQL = "SELECT id, navn, beskrivelse, kuid, parent, globalinfobase FROM infobase WHERE "& kkri &" ORDER BY navn"
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	

	%><table cellspacing=0 cellpadding=0 border=0 style="page-break-before:always;" width="600"><tr>
	<td style="padding-left : 20px; padding-top:10;" valign="top"><%
	Response.write "<font size=3><img src='../ill/pil_kalender.gif' width='10' height='12' alt='' border='0'>&nbsp;<b>" &oRec("navn")& "</b></font><br><img src='../ill/hrstreg_600.gif' width='500' height='1' alt='' border='0'><br><br>"
	Response.write oRec("beskrivelse") &"<br><br>"
	
	
		'**** niveau 2 ****	
		strSQL2 = "SELECT id, navn, beskrivelse, parent, globalinfobase FROM infobase WHERE parent = " & oRec("id")
		oRec2.open strSQL2, oConn, 3
		while not oRec2.EOF
		
			Response.write "<br><br><img src='../ill/pil_kalender.gif' width='10' height='12' alt='' border='0'>&nbsp;<img src='../ill/pil_kalender.gif' width='10' height='12' alt='' border='0'>&nbsp;<font size=3><b>" &oRec2("navn")& "</b></font><br>"
			Response.write oRec2("beskrivelse")
			Response.write "<br><br>"
			
			'**** niveau 3 ****
			strSQL3 = "SELECT id, navn, beskrivelse, parent, globalinfobase FROM infobase WHERE parent = " & oRec2("id")
			oRec3.open strSQL3, oConn, 3
			while not oRec3.EOF
		
				Response.write "<br><br><img src='../ill/pil_kalender.gif' width='10' height='12' alt='' border='0'>&nbsp;<img src='../ill/pil_kalender.gif' width='10' height='12' alt='' border='0'>&nbsp;<img src='../ill/pil_kalender.gif' width='10' height='12' alt='' border='0'>&nbsp;<font size=2><b>" &oRec3("navn")& "</b></font><br>"
				Response.write oRec3("beskrivelse")
				Response.write "<br><br>"
			
			oRec3.movenext
			wend
			oRec3.close
		
			
		oRec2.movenext
		wend
		oRec2.close
	
	
	Response.write "</td></tr></table>"
	
	oRec.movenext
	wend
	oRec.close
	
	'************************************ 
	'Print et enkelt punkt
	'************************************
	case "print"
	infopunktid = request("FM_print_id")
	%>
	<html>
	<head>
		<title>TimeOut 2.1</title>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
	</head>
	<body topmargin="0" leftmargin="0" class="regular">
	<table cellspacing="0" cellpadding="0" border="0" width="880">
	<tr>
		<td bgcolor="#003399" width="650"><img src="../ill/logo_topbar_print.gif" alt="" border="0"></td>
		<td bgcolor="#FFFFFF" align=right><a href="javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0">&nbsp;Tilbage</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"><a href="javascript:window.print()"><img src="../ill/print_xp.gif" width="28" height="30" alt="" border="0">&nbsp;Print</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"></td>
	</tr>
	</table>
	<br><br><br>
	<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#FFFFFF">
	<tr>
		<td style="padding-left : 20px; padding-top:10;" valign="top">
		<%
		strSQL = "SELECT id, navn, beskrivelse, kuid, parent, globalinfobase FROM infobase WHERE id = "& infopunktid &""
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		
		Response.write "<font size=3><b>" &oRec("navn")& "</b></font><br><img src='../ill/hrstreg_600.gif' width='500' height='1' alt='' border='0'><br><br>"
		Response.write oRec("beskrivelse") &"<br><br>"
		
		end if
		oRec.close
		%>
	</td></tr>
	</table>
	
	
	<%
	case else
	%>
	
	<%if kview <> "j" then%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	<%call tsamainmenu(5)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call kundtopmenu()
	%>
	</div>
	
	
	
	<%else%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
			
			<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
			<%
			call kundelogin_mainmenu(4, lto, kontaktid)%>
			</div>
			
			
	<%end if%>
	
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	
	
	 if (document.images){
			plus = new Image(200, 200);
			plus.src = "ill/plus.gif";
			minus = new Image(200, 200);
			minus.src = "ill/minus2.gif";
			}
	
			function expand (de){
				if (document.all("D" + de)){
					if (document.all("D" + de).style.display == "none"){
					document.all("D" + de).style.display = "";
					document.images["plus" + de].src = minus.src;
				}else{
					document.all("D" + de).style.display = "none";
					document.images["plus" + de].src = plus.src;
				}
			}
		}
	
	function showHTML(thismesHTML){
	document.getElementById("besk_HTML_"+thismesHTML).style.visibility = "visible"
	document.getElementById("besk_HTML_"+thismesHTML).style.display = ""
	}
	
	function hideHTML(thismesHTML){
	document.getElementById("besk_HTML_"+thismesHTML).style.visibility = "hidden"
	document.getElementById("besk_HTML_"+thismesHTML).style.display = "none"
	}
	
	function showmes(thismes){
	
	lastredigervalue = document.getElementById("lastredigervalue").value
	
	//if (lastredigervalue > 0) {
	document.getElementById("besk_"+lastredigervalue).style.visibility = "hidden"
	document.getElementById("besk_"+lastredigervalue).style.display = "none"
	//}
	
	document.getElementById("besk_"+thismes).style.visibility = "visible"
	document.getElementById("besk_"+thismes).style.display = ""
	
	document.getElementById("lastredigervalue").value = thismes
	
	}
	
	function hidesearcht(){
	document.all["FM_search"].value = ""
	}
	//-->
	</script>
	<%
	function SQLBless2(s)
	dim tmp
	tmp = s
	tmp = replace(tmp, "'", "")
	SQLBless2 = tmp
	end function
	
	
	function beskdiv(infobase_record_ID, infobase_record_besk, infobase_record_navn, x, topx, editor, dato)
	
			if x = 0 then 
			vzb = "visible"
			dsp = ""
			%>
			<input type="hidden" name="lastredigervalue" id="lastredigervalue" value="<%=infobase_record_ID%>">
			<%
			else
			vzb = "hidden"
			dsp = "none"
			end if
			%>
			<div style="position:absolute; width:400; left:282; top:<%=topx%>; visibility:<%=vzb%>; display:<%=dsp%>;" name="besk_<%=infobase_record_ID%>" id="besk_<%=infobase_record_ID%>">
			<table><tr><td>
			<font class=megetlillesort>Sidst redigeret af: <%=editor%>, d. <%=formatdatetime(dato, 1)%></font> <br>
			<input type="text" name="navn_<%=infobase_record_ID%>" id="navn_<%=infobase_record_ID%>" value="<%=infobase_record_navn%>" style="border:1px #003399 solid; font-size:9px; width:200px;">
			<%if kview <> "j" then%>
			&nbsp;<a href="infobase.asp?menu=kund&id=<%=id%>&func=slet&infid=<%=infobase_record_ID%>" class=red>Slet</a>&nbsp;|
			<%end if%>
			&nbsp;
			<a href="#" onClick="hideHTML('<%=infobase_record_ID%>')" class=kal_g_bold>Se som Tekst</a>
			&nbsp;
			|&nbsp;
			<a href="#" onClick="showHTML('<%=infobase_record_ID%>')" class=vmenu>Se som HTML</a>
			<textarea cols="57" rows="19" name="FM_besk_<%=infobase_record_ID%>" id="FM_besk_<%=infobase_record_ID%>" style="border:1px #003399 solid;"><%=trim(infobase_record_besk)%></textarea>
			
			<%if kview <> "j" then%>
			<img src="ill/blank.gif" width="175" height="1" border="0"><input type="image" src="../ill/opdaterpil.gif">
			<%end if%>
			</td></tr></table>
			</div>
			<div id="besk_html_<%=infobase_record_ID%>" name="besk_html_<%=infobase_record_ID%>" style="position:absolute; width:460; height:308; left:285; top:67; visibility:hidden; display:none; border:1px #003399 solid; background-color:#ffffff;"><%=trim(infobase_record_besk)%></div>
	<%
	end function
	
	
				
			
			
		
	
	
	if kview = "j" then
	dtop = 130
	dleft = 190
	sqlKundKri = " AND kid = " & kontaktid
			
		'** Firma logo ***
		call visfirmalogo(20, 205, kontaktid)
			
	else
	dtop = 132
	dleft = 20
	sqlKundKri = ""
	end if
	%>
	<div id="sindhold" style="position:absolute; left:<%=dleft%>; top:<%=dtop%>; visibility:visible;">
	<h3>Infobase</h3>
	<div style="position:relative; top:0px; left:0px; padding:10px; background-color:#ffffe1; border:1px #5582d2 solid;">
	<table border=0 cellpadding=0 cellspacing=0 width="255" bgcolor="#ffffe1">
	<form action="infobase.asp?menu=kund&usekview=<%=kview%>&kontaktid=<%=kontaktid%>" method="post" name="search" id="search">
	<tr>
		<td style="padding-left:5px;"><b>Kontakt:</b><br>
		<%
		
		
		strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'e' "& sqlKundKri &" ORDER BY Kkundenavn"
		'Response.write strSQL & "<br>"
		'Response.flush		
		%>
		<select name="id" id="id" style="font-size : 11px; width:255px;">
		<%
				
				oRec.open strSQL, oConn, 3
				k = 0
				while not oRec.EOF
				
				if k = 0 then
				
				if cint(id) = 0 then
				isGlbSel = "SELECTED"
				else
				isGlbSel = ""
				end if
				%>
				<option value="0" <%=isGlbSel%>>Alle (Global Infobase)</option>
				<%end if
				
				if cint(id) = cint(oRec("kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%>&nbsp;(<%=oRec("Kkundenr")%>)</option>
				<%
				k = k + 1
				oRec.movenext
				wend
				oRec.close
				
				
				%>
		</select>
	</td>
	</tr>

	
	</tr>
	<td style="padding-left:5px; padding-top:5px;">
	<%
	'if len(trim(request("FM_search"))) = 0 then
	'strsoegindhold = "Søg i infobasen.."
	'else
	strsoegindhold = request("FM_search")
	'end if
	
	%>
	<b>Søg:</b>&nbsp;<input type="text" name="FM_search" id="FM_search" value="<%=strsoegindhold%>" style="font-size : 11px; width:105px;" onFocus="hidesearcht()">
	</td>
	</tr>
	<tr><td align=right style="padding-top:20px; padding-top:5px;">
	<input type="submit" value="Vis Infobase">
	</td>
	</tr>
	</form>
	</table>
	
	
	
	<%if kview <> "j" then%>
	<div id="opretny" style="position:absolute; left:640; top:0; visibility:visible;">
	<a href='infobase.asp?menu=kund&func=opr&infid=0&id=<%=id%>&parent=0' class='vmenu'><img src="../ill/opretpkt.gif" width="116" height="30" alt="" border="0"></a>
	</div>
	<%end if%>
	
	<%
	'**** Eksport/Print funktion ****
	if kview <> "j" then
	prtop = 0
	prleft = 290
	bgcol = "#d6dff5"
	else
	prtop = 40
	prleft = 500
	bgcol = "#ffffff"
	end if
	%>
	
	<%if kview <> "j" then%>
	<div id="eksport" style="position:absolute; left:<%=prleft%>; top:<%=prtop%>; visibility:visible;">
	<a href='infobase.asp?menu=kund&func=eks&infid=0&id=<%=id%>&parent=0' class='vmenu'>Eksportér/print infobase <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
	</div>
	<%end if%>
	
	<br><br>
	Valgt Infobase:
	<%if id <> 0 then
		strSQL = "SELECT kkundenavn FROM kunder WHERE kid =" & id
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		Response.write "<b>"& oRec("kkundenavn") &"</b>"
		end if
		oRec.close
	else%>
	<b>Global</b>
	<%end if%>
	
	<form action="infobase.asp?menu=kund&func=dbred&id=<%=id%>" method="post" name="opdater" id="opdater">
	
	
	<%
	dim oRec2id, oRec2besk, oRec2navn, oRec2x, oRec2Editor, oRec2Dato
	Redim oRec2id(0), oRec2besk(0), oRec2navn(0), oRec2x(0), oRec2Editor(0), oRec2Dato(0)
	
	'*** Kundelogin skal kunne se den globale ***
	if kview <> "j" then
	
		if cint(id) <> 0 then
		kkri = "kuid = "& id &" AND (parent = 0)"
		else
		kkri = "(globalinfobase = 1 AND parent = 0)"
		end if
	
	else
	kkri = "kuid = "& id &" AND (parent = 0) OR (globalinfobase = 1 AND parent = 0)"
	end if
	
	if len(trim(request("FM_search"))) = 0 then
	kkri = kkri
	else
	kkri = kkri &" AND (navn LIKE '%"&request("FM_search")&"%' OR navn LIKE '"&request("FM_search")&"' OR beskrivelse LIKE '%"&request("FM_search")&"%' OR beskrivelse LIKE '"&request("FM_search")&"')"
	usesearch = 1
	end if 
	
	strSQL = "SELECT id, navn, beskrivelse, kuid, parent, globalinfobase, editor, dato FROM infobase WHERE "& kkri &" ORDER BY navn" 
	'Response.write strSQL
	
	oRec.open strSQL, oConn, 3
	x = 0
	y = 1
	while not oRec.EOF
	select case right(x, 1)
	case 0, 2, 4, 6, 8
	bc = "#FFFFFF"
	case else
	bc = "#EFF3FF"
	end select
	%>
	<table cellspacing="0" cellpadding="0" border="0" width="250" bgcolor="#5582d2">
	<tr>
		<td bgcolor="#d6dff5" colspan="2"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bc%>">
		<td height=22 style="padding-left:5;">
		<%if usesearch <> 1 then%>
				<%
				count1 = 0
				strSQL2 = "SELECT count(id) AS antal FROM infobase WHERE parent = "& oRec("id") &" ORDER BY id" 
				oRec2.open strSQL2, oConn, 3
				if not oRec2.EOF then
				count1 = oRec2("antal")
				end if
				oRec2.close
				
				if count1 <> 0 then
				%> (<%=count1%>) <a href="javascript:expand('<%=oRec("id")%>');" class='vmenu'><img src="ill/plus.gif" width="9" height="9" border="0" name="plus<%=oRec("id")%>"></a>
				<%else%>
				<img src="ill/blank.gif" width="12" height="9" border="0">
				<%end if%>
		&nbsp;<a href="#" class='vmenu' onClick="showmes('<%=oRec("id")%>')"><%=oRec("navn")%></a>
			<%if kview <> "j" then%>
			&nbsp;<a href='infobase.asp?menu=kund&func=opr&infid=0&id=<%=id%>&parent=<%=oRec("id")%>' class='vmenu'><img src="../ill/addmore55.gif" width="10" height="13" alt="Tilføj element" border="0"></a>
			<%end if%>
		<%else%>
		<a href="#" class='vmenu' onClick="showmes('<%=oRec("id")%>')"><%=oRec("navn")%></a>
		<%end if%>
		</td>
	</tr>
	</table>
	
		
			<%
			call beskdiv(oRec("id"), oRec("beskrivelse"), oRec("navn"), x, 30, oRec("editor"), oRec("dato"))
			
			
				
				
				if usesearch <> 1 then%>
				<DIV ID="D<%=oRec("id")%>" Style="position: relative; display: none; width:100;">
				<%
				
				strSQL2 = "SELECT id, navn, beskrivelse, kuid, parent, editor, dato FROM infobase WHERE parent = "& oRec("id") &" ORDER BY navn" 
				oRec2.open strSQL2, oConn, 3
				while not oRec2.EOF
				x = x + 1
				%>
				<table cellspacing="0" cellpadding="0" border="0" width="250" bgcolor="<%=bc%>">
				<tr>
					<td width=180 height=22 style="padding-left:20;">
					<img src="ill/blank.gif" width="9" height="9" border="0">
					<a href="#" class='vmenu' onClick="showmes('<%=oRec2("id")%>')">
					<%=oRec2("navn")%></a></td>
				</table>
				
				<%
				Redim preserve oRec2id(y), oRec2besk(y), oRec2navn(y), oRec2x(y), oRec2Editor(y), oRec2Dato(y) 
				
				oRec2id(y) = oRec2("id")
				oRec2besk(y) = oRec2("beskrivelse")
				oRec2navn(y) = oRec2("navn")
				oRec2x(y) = x
				oRec2Editor(y) = oRec2("editor")
				oRec2Dato(y) = oRec2("dato")
				
				y = y + 1
				oRec2.movenext
				wend
				oRec2.close
				%>	
				</div>
				<%
				end if 'usesearch%>
	<%
	x = x + 1
	oRec.movenext
	wend
	oRec.close
	
	for y = 1 to y - 1
	 call beskdiv(oRec2id(y), oRec2besk(y), oRec2navn(y), oRec2x(y), 30, oRec2Editor(y), oRec2Dato(y))
	next
	
	if x = 0 then%>	
	<table cellspacing="0" cellpadding="0" border="0" width="460" bgcolor="<%=bgcol%>">
	<tr>
		<td bgcolor="<%=bgcol%>"><font color=red><b>
		<%if len(trim(request("FM_search"))) <> 0 then%>
		Der blev ikke fundet nogen punkter der matcher de valgte søgekriteier, eller infobasen er tom!.<br><br>
		<%else%>
		Denne infobase er tom!
		<%end if%>
		</b></font><br>
		
		
		<%if kview <> "j" then%>
		Der er ikke oprettet nogen menupunkter i den valgte infobase.<br>
		Opret indold i denne infobase ved at klikke på opret nyt pkt. knappen i øverst til højre.
		<%else%>
		<br>
		Kontakt evt. <b><%=lto%></b> for flere oplysninger vedr. brug af infobase.
		<%end if%></td>
	</tr>
	</table>
	<%end if%>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	
	</form>
	</div>
	
	</div>
	
	<!--<div id="print" name="print" style="position:absolute; width:40; height:15; left:715; top:195; visibility:visible; display:;">
	<form action="infobase.asp?func=print" method="POST" name="printthis" id="printthis">
	<input type="hidden" name="FM_print_id" id="FM_print_id" value="0">
	<input type="image" src="../ill/print_xp.gif" alt="Print det valgte menupkt.">
	</form>
	</div>-->
	
	<br><br>&nbsp;
	</div>
	<%
	end select
	%>
	
	


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
