<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
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
	uselogin = request("uselogin")
	
	
	
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
	
	'Create a new DevEdit class object
	dim myDE
	set myDE = new DevEdit

	'Set the name of this DevEdit class
	myDE.SetName "myDevEditTimeout"
				
	'Once the form has been submitted, GetValue will return the HTML code.
	'GetValue accepts 1 parameter, and this specifies whether to convert
	'quotes from ' to '' and " to "". If you want to save the HTML to a database,
	'pass true to the GetValue function. If not, pass false.
	
	'**** sætter værdier ***
	strNavn = SQLBless(request("FM_navn"))
	strBesk = myDE.GetValue(true)
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
		
		'*************************************************************************
		'*** Er der skiftet parent? 										  ****
		'*** Denne sættest til at følge parent med hensyn til global infobase ****
		'*************************************************************************
		
		'if intKundeid = 0 then
			'intGlobal = 1
		'else
		
			'if intOprParent <> intParent then
				'strSQL = "SELECT globalinfobase FROM infobase WHERE id = " & intParent
				'oRec.open strSQL, oConn, 3
			'	
				'if not oRec.EOF then
				'intGlobal = oRec("globalinfobase") 
				'else
				'intGlobal = intGlobal
				'end if
				
				'oRec.close
			'else
			'intGlobal = intGlobal
			'end if
			
		'end if
	
	strSQL3 = "UPDATE infobase SET navn = '"& strNavn &"',  editor = '"& strEditor &"', dato = '"& strDato &"', parent = "& intParent &", beskrivelse = '"& strBesk &"', kuid = "& intKundeid &", globalinfobase = "& intGlobal &" WHERE id=" & infid
	end if
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
		<!--#include file="../inc/regular/header_inc.asp"-->
		<!--#include file="../inc/regular/topmenu_inc.asp"-->
		<!--#include file="../inc/regular/vmenu.asp"-->
		<!--#include file="../inc/regular/rmenu.asp"-->
		
		
		<!-------------------------------Sideindhold------------------------------------->
		<div id="sindhold" style="position:absolute; left:190; top:80; visibility:visible;">
		<table cellspacing="0" cellpadding="0" border="0" width="600">
		<tr>
	    	<td valign="top"><img src="../ill/header_infobase.gif" alt="" border="0" width="758" height="45"></TD>
		</TR>
		</TABLE>
		
	<!--#include file="inc/infobase_inc.asp"-->
	<span style="position:absolute; left:410; top:138; padding:5px; background-color:#FFFFFF; width:300;"><font size=1 color=#5582d2>NB) Hvis du modtager en besked, om du vil se både "secure and unsecure items" og ønsker at fjerne denne gøres ved at:<br> Vælg menuen Tools i din browser --> Security --> Custom level --><br> Sæt "Launching program and files in IFRAME" = enable, og<br> sæt "Display mixed content" = enable. </font>
	</span>
	<!--
	<table border=0 cellpadding=0 cellspacing=0 width="600">
	<tr>
	<td valign="top" width="163"><img src="../ill/logo_bg.gif" width="163" height="53" alt="" border="0"></td>
	<td valign="top"><b>Infobase.</b><br>&nbsp;Angiv navn og beskrivelse på det nye infobase pkt. <br>
	<br>&nbsp;Vælg hvilken kunde dette pkt. relaterer sig til og angiv om det skal være tilgængeligt i den globale Infobase. 
	</td>
	</tr>
	</table>-->
	<br><br>
	<table cellspacing="0" cellpadding="0" border="0" width="400">
	<form method="POST" action="infobase.asp">
	<input type="hidden" name="infid" value="<%=infid%>">
	<input type="hidden" name="id" value="<%=id%>">
	<input type="hidden" name="func" value="<%=dbfunc%>">
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
	<tr>
		<td colspan=2><br><b>Tilgængelig i den Globale infobase?</b><br>
		<%
		'**************************************************************************************
		'*** kun level 0 kan overføres til Global infobase 									***
		'*** Underpunkter sættes til det samme som parent. Dette sættes når der opdateres.  ***
		'**************************************************************************************
		if intParent = 0 then
			if cint(intGlobal) = 1 OR id = 0 then
			chkclobal = "checked"
			else
			chkclobal = ""
			end if
			%>
			Ja, <input type="checkbox" name="FM_global" value="1" <%=chkclobal%>>&nbsp;&nbsp;Gør dette infobasepunkt tilgængelig i den Globale infobase.<br>
			(underliggende punkter følger automatisk deres parent.)
		<%
		else%>
		Kun infobase punkter på 1 niveau kan skifte tilgængelighed i den <b>Globale infobase.</b> Underliggende punkter følger automatisk deres parent.
		<%end if%>
		</td>
	</tr>
	<%
	'** Tjekker at vi ikke redigerer et punkt der ikke hører til den globale infobase *** 
	if func = "red" AND id <> 0 then 
	%>
	<tr>
		<td colspan=2><br><b>Parent:</b><br>
		<select name="FM_parent">
		<option value="0">0'te niveau (Root)</option>
		<%
		strSQL = "SELECT id, kuid, navn, parent FROM infobase WHERE kuid = " & intKuid &" AND id <> "& infid
		oRec.open strSQL, oConn, 3
		
		While Not oRec.EOF 
		if oRec("id") = intParent then
		parSel = "SELECTED"
		else
		parSel = ""
		end if
		
		if oRec("parent") = 0 then 
		%>
		<option value="<%=oRec("id")%>" <%=parSel%>><%=oRec("navn")%></option>
		<%
		else	
		
				'** tjekker at den nye parent er min. 2 niveau ****
				strSQL2 = "SELECT parent FROM infobase WHERE id = " & oRec("parent") 
				oRec2.open strSQL2, oConn, 3
		
				if not oRec2.EOF then
					if oRec2("parent") = 0 then  
					%>
					<option value="<%=oRec("id")%>" <%=parSel%>><%=oRec("navn")%></option>
					<%
					end if
				end if
				oRec2.close
				
		
		end if
		
		oRec.movenext
		Wend
		oRec.close%>
		</select>
		<input type="hidden" name="FM_opr_parent" value="<%=intParent%>">
		</td>
	</tr>
	<%else%>
	<tr><td colspan=2><br><b>Parent:</b>&nbsp;&nbsp;&nbsp;&nbsp;(kan ikke skiftes her.)
	</td></tr>
	<input type="hidden" name="FM_parent" value="<%=intParent%>">
	<%end if%>
	</table>
	<table cellspacing="0" cellpadding="0" border="0" width="800">
	<tr>
	<td colspan=2 valign="top" style="padding-top:20;"><b>Beskrivelse:</b><br>
		<%'*** skal kaldes mellem form tags ****
		call deveditset
		%>
		</td>
	</tr>
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
	
	Response.write "kview" & kview%>
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
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	<%else%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
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
	
	
	
	function showmes(thismes){
	var thistxt; 
	var thisnv;
	
	lastredigervalue = document.all["lastredigervalue"].value
	if (lastredigervalue > 0) {
	document.all["rediger_"+lastredigervalue+""].style.visibility = "hidden"
	document.all["rediger_"+lastredigervalue+""].style.display = "none"
	}
	
	thistxt = (document.all["besk_"+thismes+""].innerHTML);
	thisnv = (document.all["navn_"+thismes+""].value);
	
	frames.editor.document.designMode="On"
	editor.document.open()
	frames.editor.document.write(thistxt) 
	editor.document.close()
	document.all["showthisnavn"].innerText = thisnv
	document.all["rediger_"+thismes+""].style.visibility = "visible"
	document.all["rediger_"+thismes+""].style.display = ""
	document.all["FM_print_id"].value = thismes
	document.all["lastredigervalue"].value = thismes
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
	
	if kview <> "j" then
	dtop = 80
	dleft = 190
	else
	dtop = 50
	dleft = 20
	end if
	%>
	<div id="sindhold" style="position:absolute; left:<%=dleft%>; top:<%=dtop%>; visibility:visible;">
		
		<%if kview <> "j" then%>
		<table cellspacing="0" cellpadding="0" border="0" width="600">
		<tr>
	    	<td valign="top"><img src="../ill/header_infobase.gif" alt="" border="0" width="758" height="45"></TD>
		</TR>
		</TABLE>
		<%end if%>
		
	<table border=0 cellpadding=0 cellspacing=0 width="600">
	<tr>
	<%if kview <> "j" then%>
	<td valign="top" width="163"><img src="../ill/logo_bg.gif" width="163" height="53" alt="" border="0"></td>
	<%else%>
	<td>&nbsp;</td>
	<%end if%>
	
	<td valign="top" style="padding-left:2px;"><b>Infobase.</b><br>
	<%if kview <> "j" then%>
	Hvis der ikke er valgt nogen kunde, vises den globale infobase. <br>
	Hvis der er valgt kunde (ude til venstre) vises de kunde specifikke infobase punkter.<br><br></td>
	<%else%>
	I infobasen kan i finde oplysninger om jeres virksomhed. Det kan f.eks være information om jeres netværk, designmanualer,	kontaktpersoner, udviklingsforløb eller anden information der er vigtig altid at have ved hånden.   
	<br>
	<a href="joblog_k.asp?useKid=<%=id%>&uselogin=n" class="vmenu">Tilbage til joboversigten...</a>
	<br><br>&nbsp;</td>
	
	<%end if%>
	</tr>
	<tr>
	<td colspan="2" style="padding-left:4px;">
	<%
	if len(trim(request("FM_search"))) = 0 then
	strsoegindhold = "Søg i infobasen.."
	else
	strsoegindhold = request("FM_search")
	end if
	
	%>
	<form action="infobase.asp?menu=kund&id=<%=id%>&usekview=<%=kview%>&uselogin=<%=uselogin%>" method="post" name="search" id="search">
	<input type="text" name="FM_search" id="FM_search" value="<%=strsoegindhold%>" size="35" onFocus="hidesearcht()">
	&nbsp;<input type="image" src="../ill/pilstorxp.gif">
	</form>
	</td>
	</tr>
	</table>
	
	<%if kview <> "j" then%>
	<div id="opretny" style="position:absolute; left:640; top:50; visibility:visible;">
	<a href='infobase.asp?menu=kund&func=opr&infid=0&id=<%=id%>&parent=0' class='vmenu'><img src="../ill/opretpkt.gif" width="116" height="30" alt="" border="0"></a>
	</div>
	<%else%>
	<div id="opretny" style="position:absolute; left:-180; top:10; width:150; visibility:visible;">
	<b>Andre links:</b><br>
	<a href='joblog_k.asp?useKid=<%=id%>&uselogin=<%=uselogin%>' class='vmenu'>Joboversigten</a>
	</div>
	<%end if
	
	'**** Eksport/Print funktion ****
	if kview <> "j" then
	prtop = 135
	prleft = 600
	else
	prtop = 40
	prleft = 700
	end if
	%>
	<div id="eksport" style="position:absolute; left:<%=prleft%>; top:<%=prtop%>; visibility:visible;">
	<a href='infobase.asp?menu=kund&func=eks&infid=0&id=<%=id%>&parent=0' class='vmenu'>Eksportér/print infobase <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
	</div>
	
	Infobase:
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
	
	<%
	if cint(id) <> 0 then
	kkri = "kuid = "& id &" AND parent = 0"
	else
	kkri = "(globalinfobase = 1 AND parent = 0)"
	end if
	
	if len(trim(request("FM_search"))) = 0 then
	kkri = kkri
	else
	kkri = "(navn LIKE '%"&request("FM_search")&"%' OR navn LIKE '"&request("FM_search")&"' OR beskrivelse LIKE '%"&request("FM_search")&"%' OR beskrivelse LIKE '"&request("FM_search")&"')"
	usesearch = 1
	end if 
	
	strSQL = "SELECT id, navn, beskrivelse, kuid, parent, globalinfobase FROM infobase WHERE "& kkri &" ORDER BY navn" 
	oRec.open strSQL, oConn, 3
	x = 0
	while not oRec.EOF
	select case right(x, 1)
	case 0, 2, 4, 6, 8
	bc = "#FFFFFF"
	case else
	bc = "#EFF3FF"
	end select
	%>
	<table cellspacing="0" cellpadding="0" border="0" bordercolor="#FFFFFF" width="260" bgcolor="#5582d2">
	<form>
	<tr>
		<td bgcolor="#d6dff5" colspan="2"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bc%>">
		<td width=180 height=22 style="padding-left:5;">
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
				%><a href="javascript:expand('<%=oRec("id")%>');" class='vmenu'><img src="ill/plus.gif" width="9" height="9" border="0" name="plus<%=oRec("id")%>"></a>
				<%else%>
				<img src="ill/blank.gif" width="9" height="9" border="0">
				<%end if%>
		&nbsp;<a href="#" class='vmenu' onClick="showmes('<%=oRec("id")%>')"><%=oRec("navn")%></a>
		<%else%>
		<a href="#" class='vmenu' onClick="showmes('<%=oRec("id")%>')"><%=oRec("navn")%></a>
		<%end if%>
		</td>
		<td><div id="rediger_<%=oRec("id")%>" name="rediger_<%=oRec("id")%>" style="position:relative; left:0; top:0; visibility:hidden; display:none;">
			<%if kview <> "j" then%>
			<table cellspacing="0" cellpadding="0" border="0"><tr>
				<td><a href="infobase.asp?menu=kund&id=<%=id%>&parent=<%=oRec("id")%>&func=red&infid=<%=oRec("id")%>" class=vmenu><img src="../ill/blyant.gif" width="12" height="11" alt="Rediger" border="0"></a></td>
				<td style="padding-left:4;"><a href="infobase.asp?menu=kund&id=<%=id%>&parent=<%=oRec("id")%>&func=opr&infid=0" class=vmenu><img src="../ill/addmore3.gif" width="18" height="18" alt="Opret nyt underpkt." border="0"></a></td>
				<td style="padding-left:2;"><a href="infobase.asp?menu=kund&id=<%=id%>&func=slet&infid=<%=oRec("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="Slet" border="0"></a></td>
				</tr>
			</table>
			<%end if%>
			</div>
			<div style="display:none;" name=besk_<%=oRec("id")%> id=besk_<%=oRec("id")%>>
			<%=oRec("beskrivelse")%>
			</div>
			<input type="hidden" name="navn_<%=oRec("id")%>" id="navn_<%=oRec("id")%>" value="<%=oRec("navn")%>">
		</td>
	</tr>
	</table>
	
	
				<%if usesearch <> 1 then%>
				<DIV ID="D<%=oRec("id")%>" Style="position: relative; display: none;">
				<%
				strSQL2 = "SELECT id, navn, beskrivelse, kuid, parent FROM infobase WHERE parent = "& oRec("id") &" ORDER BY navn" 
				oRec2.open strSQL2, oConn, 3
				while not oRec2.EOF
				%>
				<table cellspacing="0" cellpadding="0" border="0" width="260" bgcolor="<%=bc%>">
				<tr>
					<td width=180 height=22 style="padding-left:20;">
					<%
					count2 = 0
					strSQL3 = "SELECT count(id) AS antal FROM infobase WHERE parent = "& oRec2("id") &" ORDER BY id" 
					oRec3.open strSQL3, oConn, 3
					if not oRec3.EOF then
					count2 = oRec3("antal")
					end if
					oRec3.close
					
					if count2 <> 0 then%>
					<a href="javascript:expand('<%=oRec2("id")%>');"><img src="ill/plus.gif" width="9" height="9" border="0" name="plus<%=oRec2("id")%>"></a>&nbsp;
					<%else%>
					<img src="ill/blank.gif" width="9" height="9" border="0">
					<%end if%>
					
					<a href="#" class='vmenu' onClick="showmes('<%=oRec2("id")%>')">
					<%=oRec2("navn")%></a></td><td>
						<div id="rediger_<%=oRec2("id")%>" name="rediger_<%=oRec2("id")%>" style="position:relative; visibility:hidden; display:none;">
						<%if kview <> "j" then%>
						<table cellspacing="0" cellpadding="0" border="0"><tr>
							<td><a href="infobase.asp?menu=kund&id=<%=id%>&parent=<%=oRec2("id")%>&func=red&infid=<%=oRec2("id")%>" class=vmenu><img src="../ill/blyant.gif" width="12" height="11" alt="Rediger" border="0"></a></td>
							<td style="padding-left:4;"><a href="infobase.asp?menu=kund&id=<%=id%>&parent=<%=oRec2("id")%>&func=opr&infid=0" class=vmenu><img src="../ill/addmore3.gif" width="18" height="18" alt="Opret nyt underpkt." border="0"></a></td>
							<td style="padding-left:2;"><a href="infobase.asp?menu=kund&id=<%=id%>&func=slet&infid=<%=oRec2("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="Slet" border="0"></a></td>
						</tr>
						</table>
						<%end if%>
						</div>
						<div style="display:none;" name=besk_<%=oRec2("id")%> id=besk_<%=oRec2("id")%>>
						<%=oRec2("beskrivelse")%>
						</div>
						<input type="hidden" name="navn_<%=oRec2("id")%>" id="navn_<%=oRec2("id")%>" value="<%=oRec2("navn")%>">
					</td>
				</tr>
				</table>
				
				
				
									<DIV ID="D<%=oRec2("id")%>" Style="position: relative; display: none;">
									<%
									strSQL3 = "SELECT id, navn, beskrivelse, kuid, parent FROM infobase WHERE parent = "& oRec2("id") &" ORDER BY navn" 
									oRec3.open strSQL3, oConn, 3
									while not oRec3.EOF
									%>
									<table cellspacing="0" cellpadding="0" border="0" width="260" bgcolor="<%=bc%>">
									<tr>
										<td height=22 style="padding-left:40;">
										<a href="#" class='vmenu' onClick="showmes('<%=oRec3("id")%>')"><%=oRec3("navn")%></a>
										</td><td width="50"><p id="rediger_<%=oRec3("id")%>" name="rediger_<%=oRec3("id")%>" style="position:relative; visibility:hidden; display:none;">
											<%if kview <> "j" then%>
											<table cellspacing="0" cellpadding="0" border="0">
												<tr>
													<td><a href="infobase.asp?menu=kund&id=<%=id%>&parent=<%=oRec3("id")%>&func=red&infid=<%=oRec3("id")%>" class=vmenu><img src="../ill/blyant.gif" width="12" height="11" alt="Rediger" border="0"></a></td>
													<td style="padding-left:4;"><a href="infobase.asp?menu=kund&id=<%=id%>&func=slet&infid=<%=oRec3("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="Slet" border="0"></a></td>
												</tr>
											</table>
											<%end if%></p>
											<div style="display:none;" name=besk_<%=oRec3("id")%> id=besk_<%=oRec3("id")%>>
											<%=oRec3("beskrivelse")%>
											</div>
											<input type="hidden" name="navn_<%=oRec3("id")%>" id="navn_<%=oRec3("id")%>" value="<%=oRec3("navn")%>"></td>
									</tr>
									</table>
									<%
									oRec3.movenext
									wend
									oRec3.close
									%>	
									</div>
									
				
				<%
				oRec2.movenext
				wend
				oRec2.close
				%>	
				</div>
				<%end if 'usesearch%>
	<%
	x = x + 1
	oRec.movenext
	wend
	oRec.close
	%>	
	<table cellspacing="0" cellpadding="0" border="0" width="260" bgcolor="#d6dff5">
	<tr>
		<td bgcolor="#d6dff5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	<input type="hidden" name="lastredigervalue" id="lastredigervalue" value="0">
	</form>
	</div>
	
	<div id="mesnavn" name="mesnavn" style="position:absolute; width:212; height:15; left:475; top:220; visibility:visible; display:;">
	<!-- job faneblad -->
	<table cellspacing="0" cellpadding="0" border="0">
		<tr>
			<td bgcolor="#5582D2" width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
			<td valign="top"><img src="../ill/tabel_top.gif" width="196" height="1" alt="" border="0"></td>
			<td bgcolor="#5582D2" align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
		</tr>
		<tr bgcolor="#5582D2">
			<td style=padding-top:2px; class=alt>
			<p name="showthisnavn" id="showthisnavn">..</p>
		</td>
		</tr>
	</table>
	<!-- slut -->
	</div>
	
	<div id="print" name="print" style="position:absolute; width:40; height:15; left:695; top:212; visibility:visible; display:;">
	<form action="infobase.asp?func=print" method="POST" name="printthis" id="printthis">
	<input type="hidden" name="FM_print_id" id="FM_print_id" value="0">
	<input type="image" src="../ill/print_xp.gif" alt="Print det valgte menupkt.">
	<img src="" width="28" height="30" alt="Print det valgte punkt." border="0">
	</form>
	</div>
	
	<IFRAME id=editor frameBorder=0 height=500 width=480 marginWidth=5 src="about:blank" style="line-height: 10px; position: absolute; left:460; top:246; z-index:10;"></IFRAME><br>
	<TEXTAREA name=frmContent style="DISPLAY: none" wrap="hard"></TEXTAREA>
	<!--<div id="mes" name="mes"></div>-->
	<div id="mes" name="mes" style="position:absolute; width:490; height:508; left:455; top:241; background-color:#FFFFFF; filter:alpha(opacity=100); z-index:5; visibility:visible; display:; border-left:1 #003399 solid; border-top:1 #003399 solid; border-bottom:1 #003399 solid; border-right:1 #003399 solid; overflow:auto; padding-left:10; padding-top:10; padding-right:10; padding-bottom:2;">
	</div>
	
	<div id="luft" name="luft" style="position:absolute; width:500; height:35; left:455; top:810; visibility:visible; display:; padding-left:10; padding-top:10; padding-right:10; padding-bottom:2;">
	<br><br>&nbsp;
	</div>
	<%
	end select
	%>
	
	


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
