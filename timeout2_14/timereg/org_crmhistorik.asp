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
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	kl = request("kl")
	crmaktion = request("crmaktion")
	thisfile = "crmhistorik"
	
	select case func
	case "slet"
	
	'*** Her spørges om det er ok at der slettes en aktion ***
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:120; visibility:visible;">
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td><img src="../ill/alert.gif" width="44" height="45" alt="" border="0">&nbsp;Du er ved at <b>slette</b> en aktion. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><img src="../ill/blank.gif" width="44" height="1" alt="" border="0">&nbsp;<a href="crmhistorik.asp?menu=crm&func=sletok&id=<%=id%>&crmaktion=<%=crmaktion%>&ketype=e&emner=<%=request("emner")%>&status=<%=request("status")%>&medarb=<%=request("medarb")%>&selpkt=<%=request("selpkt")%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
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
	'*** Her slettes en crmaktion ***
	oConn.execute("DELETE FROM crmhistorik WHERE id = "& crmaktion &"")
	
	if request("selpkt") <> "kal" then
	Response.redirect "crmhistorik.asp?menu=crm&shokselector=1&ketype=e&id="&id&"&func=hist&selpkt=hist&emner="&request("emner")&"&status="&request("status")&"&medarb="&request("medarb")&""
	else
	Response.redirect "crmkalender.asp?menu=crm&shokselector=1&id="&id&"&medarb="&request("medarb")&""
	end if
	
	
	case "dbopr", "dbred"
	personal = request("personal") 
	
	'*** Her indsættes en ny aktion i db ****
	'validering af tidspunkt *** kan vist ikke bruse mere da tidspunkt vælges fra DD. 
	if len(request("klok_timer")) = 0 then
	strKloktimer = "12"
	else
	strKloktimer = request("klok_timer")
	end if
	
	if len(request("klok_min")) = 0 then
	strKlokmin = "00"
	else
	strKlokmin = request("klok_min")
	end if
	
	if len(request("klok_timer_slut")) = 0 then
	strKloktimer_slut = "12"
	else
	strKloktimer_slut = request("klok_timer_slut")
	end if
	
	if len(request("klok_min_slut")) = 0 then
	strKlokmin_slut = "00"
	else
	strKlokmin_slut = request("klok_min_slut")
	end if
	
	if strKloktimer > 24 OR strKloktimer < 0 OR strKlokmin < 0 OR strKlokmin > 60 OR strKloktimer_slut > 24 OR strKloktimer_slut < 0 OR strKlokmin_slut < 0 OR strKlokmin_slut > 60 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 31
	call showError(errortype)
	else
		
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		'** henter medarbejder id ****
		strSQL = "SELECT mid FROM medarbejdere WHERE mnavn = '"& session("user") &"'"
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		intMid = oRec("mid")
		End if
		oRec.close
		''**''
		
		if personal <> "y" then
			strNavn = SQLBless(request("FM_navn"))
			strEmne = request("FM_emne")
			intKundeId = request("id")
			strBesk = SQLBless(request("FM_besk"))
			strEditor = session("user")
			strDato = year(now)&"/"&month(now)&"/"&day(now)
			strCrmdato = request("FM_start_aar")&"/"&request("FM_start_mrd")&"/"&request("FM_start_dag")
			strStatus = request("FM_status")
			strKontaktform = request("FM_kform")
			strKpers = request("FM_kpers")
			strKlokkeslet = strKloktimer &":"& strKlokmin&":"&"01"
			strKlokkeslet_slut = strKloktimer_slut &":"& strKlokmin_slut &":"&"01"
			
			'** Hvis slut klokkeslet er før starttid
			if strKlokkeslet > strKlokkeslet_slut then
			strKlokkeslet_slut = strKlokkeslet 
			end if
		
		else
			strNavn = ""
			strEmne = 0
			intKundeId = -1
			strBesk = SQLBless(request("FM_note_personal"))
			strEditor = session("user")
			strDato = year(now)&"/"&month(now)&"/"&day(now)
			strCrmdato = request("FM_start_aar_per")&"/"&request("FM_start_mrd_per")&"/"&request("FM_start_dag_per")
			strStatus = 0
			strKontaktform = 0
			strKpers = ""
			strKlokkeslet = request("FM_klslet_personal")&":"&"01"
			strKlokkeslet_slut = dateadd("h", 1, strKlokkeslet)
		end if
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO crmhistorik (editor, dato, crmdato, kontaktemne, kontaktpers, status, komm, navn, kundeid, kontaktform, editorid, crmklokkeslet, crmKlokkeslet_slut) VALUES ('"& strEditor &"', '"& strDato &"', '"& strCrmdato &"', "& strEmne &", '"& strKpers &"', '"& strStatus &"', '"& strBesk &"', '"& strNavn &"', "& intKundeId &", "& strKontaktform &", "& intMid &", '"& strKlokkeslet &"', '"& strKlokkeslet_slut &"')")
			
			strSQL = "SELECT id FROM crmhistorik ORDER BY id DESC"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
				aktionsid = oRec("id")
			end if
			oRec.close
		else
			oConn.execute("UPDATE crmhistorik SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', crmdato = '"& strCrmdato &"', kontaktemne = "& strEmne &", kontaktpers = '"& strKpers &"', status = '"& strStatus &"', komm = '"& strBesk &"', navn = '"& strNavn &"', kundeid = "& intKundeId &", kontaktform = "& strKontaktform &", editorid = "& intMid &", crmklokkeslet = '"& strKlokkeslet &"', crmklokkeslet_slut = '"& strKlokkeslet_slut &"' WHERE id = "& crmaktion &"")
			aktionsid = crmaktion
		end if
		
		
		'**** Opdaterer aktions relationer ****
		oConn.execute("DELETE FROM aktionsrelationer WHERE aktionsid = "& aktionsid &"")
			
		if len(request("FM_medarbrel")) <> 0 then
			Dim medarbRelationer
			Dim b
			b = 0
			medarbRelationer = Split(request("FM_medarbrel"), ", ")
				For b = 0 to Ubound(medarbRelationer)
				oConn.execute("INSERT INTO aktionsrelationer (aktionsid, medarbid) VALUES ("& aktionsid &", "& medarbRelationer(b) &") ")
				next
		else
		oConn.execute("INSERT INTO aktionsrelationer (aktionsid, medarbid) VALUES ("& aktionsid &", "& intMid &") ")
		end if
		'*** slut opdatering af realtioner ***
		
		
			if request("selpkt") = "kal" OR personal = "y" then
				if len(request("kalviskunde")) <> 0 then
				kalviskunde = request("kalviskunde")
				else
				kalviskunde = 0
				end if
			Response.redirect "crmkalender.asp?menu=crm&shokselector=1&strdag="&request("FM_start_dag_per")&"&strmrd="&request("FM_start_mrd_per")&"&straar="&request("FM_start_aar_per")&"&medarb="&request("kalmedarb")&"&id="&kalviskunde&"&ketype=e"
			else 
			Response.redirect "crmhistorik.asp?selpkt=hist&menu=crm&shokselector=1&id="& intKundeId
			end if
		
	end if
		
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny aktion ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "Tilføj!" 
	varbroedkrumme = "Opret ny aktion"
		
		'** Hvis der først skal vælges kunde ***
		if func = "opret" AND id = 0 then
		dbfunc = "opret"
		varbroedkrumme = "Opret ny aktion step 1/2"
		else
		dbfunc = "dbopr"
		varbroedkrumme = "Opret ny aktion"
		end if
		'*** ****
		
	else
	
	strSQL = "SELECT editor, dato, crmdato, kontaktemne, kontaktpers, status, komm, navn, kundeid, kontaktform, crmklokkeslet, crmklokkeslet_slut FROM crmhistorik WHERE id=" & crmaktion 
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	crmDato = oRec("crmDato")
	StrTdato = crmDato
	intEmne = oRec("kontaktemne")
	intStatus = oRec("status")
	strBesk = oRec("komm")
	intKontaktform = oRec("kontaktform")
	strKlokkeslet = oRec("crmklokkeslet")
	strKlokkeslet_slut = oRec("crmklokkeslet_slut")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdater!" 
	end if
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	<%
	if func = "red" then
	id = id
	'Response.write strKlokkeslet
	
	formatstrKlokkeslet = FormatDateTime(strKlokkeslet,4)
	klok_timer = left(formatstrKlokkeslet,2)
	klok_min = right(formatstrKlokkeslet,2)
	
	formatstrKlokkeslet_slut = FormatDateTime(strKlokkeslet_slut,4)
	klok_timer_slut = left(formatstrKlokkeslet_slut,2)
	klok_min_slut = right(formatstrKlokkeslet_slut,2)
	%>
	<!--#include file="inc/dato2.asp"-->
	<% 
	else
	
	if len(request("strdag")) <> 0 then
	strDag = request("strdag")
	strMrd = request("strmrd")
	strAar = request("straar")
	else
	strDag = day(now)
	strMrd = month(now)
	strAar = year(now)
	end if
	%>
	<!--#include file="inc/dato2_b.asp"-->
	<%
	
	
	klok_timer = left(kl, 2)
	if len(klok_timer) <> 0 then
	klok_timer = klok_timer
	else
	klok_timer = left(FormatDateTime(time,4),2)
	end if
	klok_timer_slut = klok_timer
	end if
	
	if request("selpkt") = "kal" then
	selpkt = "kal"
	kalvalues = "&kaldato="&request("kaldato")&"&kalstatus="&request("kalstatus")&"&kalemner="&request("kalemner")&"&kalmedarb="&request("kalmedarb")&"&kalviskunde="&request("kalviskunde")&""
	else
	selpkt = "hist"
	kalvalues = " "
	end if
	%>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:50; width:60%; height:600; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="550">
	<tr>
    <td valign="top"><br>
	<img src="../ill/header_crmosigt.gif" alt="" border="0"><hr align="left" width="730" size="1" color="#000000" noshade>
	</td></tr></table>
	<br>
	<table cellspacing="0" cellpadding="0" border="0">
	<tr bgcolor="#5582D2">
	<td width="3"><img src="../ill/venstre_hjorne.gif" alt="" border="0"></td>
	<td width="724" style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid;"><font class="stor-hvid">&nbsp;Historik -- &nbsp;<%=varbroedkrumme%></font><img src="../ill/blank.gif" width="690" height="1" alt="" border="0"></td>
	<td  width="3" align="right"><img src="../ill/hojre_hjorne.gif" alt="" border="0"></td>
	</tr>
	</table>
	<br>
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr><form action="crmhistorik.asp?menu=crm&func=<%=dbfunc%>&ketype=e&selpkt=<%=selpkt&kalvalues%>" method="post">
	<input type="hidden" name="crmaktion" value="<%=crmaktion%>">
		<td valign="top" colspan="2">&nbsp;</td>
	</tr>
	<%if dbfunc = "dbred" then%>
	<tr>
		<td colspan="2" valign="top" style="height:30;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></td>
	</tr>
	<%end if
	if func = "opret" AND id = 0 then
	varSubVal = " Step 2 "
	%>
	<tr>
	<td colspan="2">Kunde:&nbsp;&nbsp;<select name="id">
		<%
		strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'k' ORDER BY Kkundenavn"
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		%>
		<option value="<%=oRec("Kid")%>"><%=left(oRec("Kkundenavn"), 9)%></option>
		<%
		oRec.movenext
		wend
		oRec.close
		%>
	</select>
	<input type="hidden" name="kl" value="<%=kl%>">
	<input type="hidden" name="strdag" value="<%=strDag%>">
	<input type="hidden" name="strmrd" value="<%=strMrd%>">
	<input type="hidden" name="straar" value="<%=strAar%>"></td></tr>
	<%
	else%>
	<input type="hidden" name="id" value="<%=id%>">
	<tr>
		<td>Dato:</td>
		<td><select name="FM_start_dag">
		<option value="<%=strDag%>"><%=strDag%></option> 
		<option value="1">1</option>
	   	<option value="2">2</option>
	   	<option value="3">3</option>
	   	<option value="4">4</option>
	   	<option value="5">5</option>
	   	<option value="6">6</option>
	   	<option value="7">7</option>
	   	<option value="8">8</option>
	   	<option value="9">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	   	<option value="22">22</option>
	   	<option value="23">23</option>
	   	<option value="24">24</option>
	   	<option value="25">25</option>
	   	<option value="26">26</option>
	   	<option value="27">27</option>
	   	<option value="28">28</option>
	   	<option value="29">29</option>
	   	<option value="30">30</option>
		<option value="31">31</option></select>&nbsp;&nbsp;
		
		<select name="FM_start_mrd">
		<option value="<%=strMrd%>"><%=strMrdNavn%></option>
		<option value="1">jan</option>
	   	<option value="2">feb</option>
	   	<option value="3">mar</option>
	   	<option value="4">apr</option>
	   	<option value="5">maj</option>
	   	<option value="6">jun</option>
	   	<option value="7">jul</option>
	   	<option value="8">aug</option>
	   	<option value="9">sep</option>
	   	<option value="10">okt</option>
	   	<option value="11">nov</option>
	   	<option value="12">dec</option></select>
		
		
		<select name="FM_start_aar">
		<option value="<%=strAar%>">
		<%if id <> 0 AND func <> "opret" then%>
		20<%=strAar%>
		<%else%>
		<%=strAar%>
		<%end if%></option>
		<option value="02">2002</option>
		<option value="03">2003</option>
	   	<option value="04">2004</option>
	   	<option value="05">2005</option>
		<option value="06">2006</option>
		<option value="07">2007</option></select></td>
	</tr>
		<tr><td>Start tidspunkt:</td><td>
		<select name="klok_timer">
		<option value="<%=klok_timer%>" SELECTED><%=klok_timer%></option>
		<option value="07">7</option>
	   	<option value="08">8</option>
	   	<option value="09">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	  	</select>&nbsp;:
		<%if func = "opret" then%>
		<select name="klok_min">
		<option value="00">00</option>
		<option value="15">15</option>
	   	<option value="30">30</option>
	   	<option value="45">45</option>
		</select>
		<%else%>
		<input type="text" name="klok_min" value="<%=klok_min%>" size="2" maxlength="2">
		<%end if%>
		&nbsp;&nbsp;(tt:mm)</td></tr>
	<tr>
	</tr>
		<tr><td>Slut tidspunkt:</td><td>
		<select name="klok_timer_slut">
		<option value="<%=klok_timer_slut%>" SELECTED><%=klok_timer_slut%></option>
		<option value="07">7</option>
	   	<option value="08">8</option>
	   	<option value="09">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	  	</select>&nbsp;:
		<%if func = "opret" then%>
		<select name="klok_min_slut">
		<option value="00">00</option>
		<option value="15">15</option>
	   	<option value="30" SELECTED>30</option>
	   	<option value="45">45</option>
		</select>
		<%else%>
		<input type="text" name="klok_min_slut" value="<%=klok_min_slut%>" size="2" maxlength="2">&nbsp;&nbsp;(tt:mm)</td></tr>
		<%end if%>
	<tr>
		<td>Emne:</td>
		<td><select name="FM_emne">
		<%strSQL = "SELECT id, navn FROM crmemne ORDER BY navn"
		oRec.open strSQL, oConn, 3
		While not oRec.EOF %>
			<%if intEmne = oRec("id") then
			selected = "selected"
			else
			selected = ""
			end if%> 
		<option value="<%=oRec("id")%>" <%=selected%>><%=oRec("navn")%></option>
		<%oRec.movenext
		wend
		oRec.close%>
	</select></td>
	</tr>
	<tr>
		<td>Kommentar til aktion:</td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"> (ikke nødvendig)</td>
	</tr>
	<tr>
		<td>Kontakt person:</td>
		<td><select name="FM_kpers">
		<%strSQL = "SELECT kid, kontaktpers1, kontaktpers2, kontaktpers3, kontaktpers4, kontaktpers5 FROM kunder WHERE Kid = "& id
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then%>
			<%if len(trim(oRec("kontaktpers1"))) <> 0 then %>
			<option value="<%=oRec("kontaktpers1")%>"><%=oRec("kontaktpers1")%></option>
			<%end if%>
			<%if len(trim(oRec("kontaktpers2"))) <> 0 then %>
			<option value="<%=oRec("kontaktpers2")%>"><%=oRec("kontaktpers2")%></option>
			<%end if%>
			<%if len(trim(oRec("kontaktpers3"))) <> 0 then %>
			<option value="<%=oRec("kontaktpers3")%>"><%=oRec("kontaktpers3")%></option>
			<%end if%>
			<%if len(trim(oRec("kontaktpers4"))) <> 0 then %>
			<option value="<%=oRec("kontaktpers4")%>"><%=oRec("kontaktpers4")%></option>
			<%end if%>
			<%if len(trim(oRec("kontaktpers5"))) <> 0 then %>
			<option value="<%=oRec("kontaktpers5")%>"><%=oRec("kontaktpers5")%></option>
			<%end if%>
		<%end if
		oRec.close%>
		<option value="Ingen">Ingen</option>
	</select></td>
	</tr>
	<tr><td colspan=2>&nbsp;</td></tr>
	<tr>
		<td valign=top colspan="2">Beskrivelse/ Log:<br>
	<textarea cols="70" rows="12" name="FM_besk" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><%=strBesk%></textarea><br>&nbsp;</td>
	</tr>
	<tr>
		<td>Status:</td>
		<td><select name="FM_status">
		<%strSQL = "SELECT id, navn FROM crmstatus ORDER BY navn"
		oRec.open strSQL, oConn, 3
		While not oRec.EOF %>
		<%if intStatus = oRec("id") then
		selected = "selected"
		else
		selected = ""
		end if%> 
		<option value="<%=oRec("id")%>" <%=selected%>><%=oRec("navn")%></option>
		<%oRec.movenext
		wend
		oRec.close%>
	</select></td>
	</tr>
	<tr>
		<td>Kontaktform:</td>
		<td><select name="FM_kform">
		<%strSQL = "SELECT id, navn FROM crmkontaktform ORDER BY navn"
		oRec.open strSQL, oConn, 3
		While not oRec.EOF %>
		<%if intKontaktform = oRec("id") then
		selected = "selected"
		else
		selected = ""
		end if%> 
		<option value="<%=oRec("id")%>" <%=selected%>><%=oRec("navn")%></option>
		<%oRec.movenext
		wend
		oRec.close%>
	</select></td>
	</tr>
	<tr>
		<td colspan="2"><br>Tilføj medarbejdere til denne aktion:<br></td>
	</tr>
	<%
		
		strSQL = "SELECT mid, mnavn FROM medarbejdere WHERE Brugergruppe = 1 OR Brugergruppe = 3 OR Brugergruppe = 6 ORDER BY mnavn"
		oRec.open strSQL, oConn, 0, 1
		while not oRec.EOF
			if func = "red" then
				strSQL2 = "SELECT medarbid FROM aktionsrelationer WHERE aktionsid = "& crmaktion &" AND medarbid = "& oRec("mid")
				oRec2.open strSQL2, oConn, 0, 1
				if not oRec2.EOF then
				strcheckmedarb = "CHECKED"
				else
				strcheckmedarb = " "
				end if
				oRec2.close
			else
				if oRec("mnavn") = session("user") then
				strcheckmedarb = "CHECKED"
				else
				strcheckmedarb = " "
				end if
			end if
		
		%> 
		<tr><td><%=oRec("mnavn")%></td><td><input type="checkbox" name="FM_medarbrel" value="<%=oRec("mid")%>" <%=strcheckmedarb%> style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td></tr>
		<%
		oRec.movenext
		wend
		oRec.close
		%>
		
	<%end if%>
	<tr>
		<td colspan="2"><br><br><img src="ill/blank.gif" width="100" height="1" alt="" border="0"><input type="submit" value="<%=varSubVal%>" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid; cursor:hand; width:116;"></td>
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
	emner = request("emner")
	status = request("status")
	medarb = request("medarb")
	
	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
	end function
	
	'** Hvis der søges på telefon nummer **
	if len(request("sogekri")) <> 0 then
	sogekri = trim(SQLBless(request("sogekri")))
	kundeKritlf = " Kkundenavn LIKE '%"& sogekri &"%' OR telefon LIKE '" & sogekri &"%'"_
	&" OR telefonmobil LIKE '%"& sogekri &"%'"_
	&" OR telefonalt LIKE '"& sogekri & "%' OR email LIKE '"& sogekri &"%'"_
	&" OR kontaktpers1 LIKE '%"& sogekri &"%'"_
	&" OR kontaktpers2 LIKE '%"& sogekri &"%'"_
	&" OR kontaktpers3 LIKE '%"& sogekri &"%'"_
	&" OR kontaktpers4 LIKE '%"& sogekri &"%'"_
	&" OR kontaktpers5 LIKE '%"& sogekri &"%'"_
	&" OR kpersemail1 LIKE '"& sogekri &"%'"_
	&" OR kpersemail2 LIKE '"& sogekri &"%' OR kpersemail3 LIKE '"& sogekri &"%'"_
	&" OR kpersemail4 LIKE '"& sogekri &"%' OR kpersemail5 LIKE '"& sogekri &"%'"_
	&" OR kperstlf1 LIKE '"& sogekri &"%' OR kperstlf2 LIKE '"& sogekri &"%'"_
	&" OR kperstlf3 LIKE '"& sogekri &"%'"_
	&" OR kperstlf4 LIKE '"& sogekri &"%' OR kperstlf5 LIKE '"& sogekri &"%'"_
	&" OR kpersmobil1 LIKE '"& sogekri &"%'"_
	&" OR kpersmobil2 LIKE '"& sogekri &"%' OR kpersmobil3 LIKE '"& sogekri &"%'"_
	&" OR kpersmobil4 LIKE '"& sogekri &"%' OR kpersmobil5 LIKE '"& sogekri &"%' "
	else
	sogekri = ""
	kundeKritlf = " Kid =" & id &""
	end if
	%>
	
	<!--#include file="../inc/regular/header_inc.asp"-->
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	</script>
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:50; width:60%; height:600; visibility:visible;">
	<br>
	<%
	'**************************** Søgefunktion *********************************************'
	%>
	<span style="position:absolute; top:5; left:538;">
	<table cellspacing="0" cellpadding="0" border="0" width="120">
	<tr>
		<td><img src="../ill/v-menu_soeg_t.gif" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td height="30" valign="middle" align="center">	
		<table>
		<tr><form action="crmhistorik.asp?menu=crm&func=hist&selpkt=hist" method="post">
		<td><input type="text" name="sogekri" size="28" maxlength="28" style="font-size : 11px;"></td>
		<td><input type="image" src="../ill/pil_groen_lille.gif" border="0"></td>
		</tr></form>
		</table></td>
		</tr>
	</table>
	</span>
	<%
	'**************************************************************************************
	%>

	
	
	<table cellspacing="0" cellpadding="0" border="0" width="740">
	<tr>
    <td valign="top"><br>
	<img src="../ill/header_crmhist.gif" alt="" border="0"><hr align="left" width="740" size="1" color="#000000" noshade>
	</td></tr></table>
	<br>
	
	<%
	'****************************************************************************************** 
	'SEARCH RESULTS.... Finder kunde udfra telefon, email navn etc. 
	'*******************************************************************************************
	if len(sogekri) <> 0 then
	Response.write "Der blev søgt på: <b>" & sogekri &"</b><br>" 
	Response.write "Følgende firmaer matchede søgningen:<br><br>"
	end if
	
	strSQL = "SELECT kkundenavn, Kid FROM kunder WHERE "& kundeKritlf& " ORDER BY kkundenavn"
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	valgtKunde = oRec("kkundenavn")
	id = oRec("Kid")
	
	if len(sogekri) <> 0 then
	Response.write "<a href='crmhistorik.asp?menu=crm&id="&oRec("Kid")&"&func=hist&selpkt=hist'>"& oRec("kkundenavn") & "&nbsp;</a><br>" 
	x = 1
	end if
	
	oRec.movenext
	wend
	oRec.close

	if x <> 1 AND len(sogekri) <> 0 then
	Response.write "<img src='../ill/alert.gif' width='44' height='45' alt='' border='0'>Der blev <b>ikke</b> fundet nogen emner der matchede søgekriteriet!"
	end if
	
	x = 0
	
	'********************************************************************************************
	' Udskriver aktions historik liste
	'********************************************************************************************
	if len(sogekri) = 0 then%>
	<div id="opretny" style="position:absolute; left:390; top:40; visibility:visible;">
	<a href="crmhistorik.asp?menu=crm&func=opret&id=<%=id%>&ketype=e">Opret ny aktion <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
	</div>
	<table cellspacing="0" cellpadding="0" border="0" width="740">
	<tr bgcolor="#5582D2">
	<td width="3"><img src="../ill/venstre_hjorne.gif" alt="" border="0"></td>
	<td width="730" style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid;" class='alt'>&nbsp;Aktions historik&nbsp;&nbsp;&nbsp;<b><%=valgtKunde%></b>
	<img src="../ill/blank.gif" width="375" height="1" alt="" border="0"></td>
	<td width="3" align="right"><img src="../ill/hojre_hjorne.gif" alt="" border="0"></td>
	</tr>
	</table>
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr>
    <td valign="top">
	<table cellspacing="0" cellpadding="0" border="0" width="740" bgcolor="#EFF3FF">
	<% 
	'************************************************************************************ 
	'Table header 
	'************************************************************************************
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030%>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td width="160" height=20>&nbsp;&nbsp;<a href="crmhistorik.asp?menu=crm&sort=firm&func=hist&id=<%=id%>&selpkt=hist&emner=<%=emner%>&status=<%=status%>&medarb=<%=medarb%>"><b>Firma</b></a></td>
		<td width="120"><a href="crmhistorik.asp?menu=crm&sort=dato&func=hist&id=<%=id%>&selpkt=hist&emner=<%=emner%>&status=<%=status%>&medarb=<%=medarb%>"><b>Dato</b></a></td>
		<td width="110"><a href="crmhistorik.asp?menu=crm&sort=emne&func=hist&id=<%=id%>&selpkt=hist&emner=<%=emner%>&status=<%=status%>&medarb=<%=medarb%>"><b>Emne</b></a></td>
		<td width="100"><b>Kommentar</b></td>
		<td><a href="crmhistorik.asp?menu=crm&sort=status&func=hist&id=<%=id%>&selpkt=hist&emner=<%=emner%>&status=<%=status%>&medarb=<%=medarb%>"><b>Status</b></a></td>
		<td align=center>&nbsp;</td>
		<td width="90"><a href="crmhistorik.asp?menu=crm&sort=editor&func=hist&id=<%=id%>&selpkt=hist&emner=<%=emner%>&status=<%=status%>&medarb=<%=medarb%>"><b>Oprettet af</b></a></td>
		<td>&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	select case level 
	case "1", "2"
			if medarb = 0 OR len(medarb) = 0 then
			usemedarbKri = " AND editorid <> 0 "
			else
			usemedarbKri = " AND aktionsid = crmhistorik.id AND medarbid = '" & medarb & "' "
			end if
	case else
	usemedarbKri = " AND aktionsid = crmhistorik.id AND medarbid = '" & session("Mid") & "' "
	end select
	
	if id = 0 OR len(id) = 0 then
	useidKri = " kundeid <> " & id & ""
	else
	useidKri = " kundeid = " & id & ""
	end if
	
	if emner = 0 OR len(emner) = 0 then
	useemneKri = " "
	else
	useemneKri = " AND kontaktemne = " & emner & ""
	end if
	
	if status = 0 OR len(status) = 0 then
	usestatKri = " "
	else
	usestatKri = " AND status = " & status & ""
	end if
	
	
	sort = Request("sort")
	select case sort
	case "emne" 
	strSQL = "SELECT DISTINCT(crmhistorik.id), kid, kkundenavn, crmhistorik.editor, crmdato, kontaktemne, crmhistorik.navn, crmemne.navn AS enmenavn, crmstatus.navn AS statusnavn, crmkontaktform.navn AS kontaktform, ikon, crmhistorik.kontaktpers FROM crmhistorik, crmemne, crmstatus, crmkontaktform, kunder, aktionsrelationer WHERE "& useidKri &" "& useemneKri &""& usestatKri &""& usemedarbKri &" AND crmemne.id = kontaktemne AND crmstatus.id = status AND crmkontaktform.id = kontaktform AND kid = kundeid ORDER BY kontaktemne, crmhistorik.crmdato DESC"
	case "editor" 
	strSQL = "SELECT DISTINCT(crmhistorik.id), kid, kkundenavn, crmhistorik.editor, crmdato, kontaktemne, crmhistorik.navn, crmemne.navn AS enmenavn, crmstatus.navn AS statusnavn, crmkontaktform.navn AS kontaktform, ikon, crmhistorik.kontaktpers FROM crmhistorik, crmemne, crmstatus, crmkontaktform, kunder, aktionsrelationer WHERE "& useidKri &" "& useemneKri &""& usestatKri &""& usemedarbKri &" AND crmemne.id = kontaktemne AND crmstatus.id = status AND crmkontaktform.id = kontaktform AND kid = kundeid ORDER BY crmhistorik.editor, crmhistorik.crmdato DESC"
	case "status"
	strSQL = "SELECT DISTINCT(crmhistorik.id), kid, kkundenavn, crmhistorik.editor, crmdato, kontaktemne, crmhistorik.navn, crmemne.navn AS enmenavn, crmstatus.navn AS statusnavn, crmkontaktform.navn AS kontaktform, ikon, crmhistorik.kontaktpers FROM crmhistorik, crmemne, crmstatus, crmkontaktform, kunder, aktionsrelationer WHERE "& useidKri &" "& useemneKri &""& usestatKri &""& usemedarbKri &" AND crmemne.id = kontaktemne AND crmstatus.id = status AND crmkontaktform.id = kontaktform AND kid = kundeid ORDER BY statusnavn, crmhistorik.crmdato DESC"
	case "firm"
	strSQL = "SELECT DISTINCT(crmhistorik.id), kid, kkundenavn, crmhistorik.editor, crmdato, kontaktemne, crmhistorik.navn, crmemne.navn AS enmenavn, crmstatus.navn AS statusnavn, crmkontaktform.navn AS kontaktform, ikon, crmhistorik.kontaktpers FROM crmhistorik, crmemne, crmstatus, crmkontaktform, kunder, aktionsrelationer WHERE "& useidKri &" "& useemneKri &""& usestatKri &""& usemedarbKri &" AND crmemne.id = kontaktemne AND crmstatus.id = status AND crmkontaktform.id = kontaktform AND kid = kundeid ORDER BY kkundenavn, crmhistorik.crmdato"
	case else
	strSQL = "SELECT DISTINCT(crmhistorik.id), kid, kkundenavn, crmhistorik.editor, crmdato, kontaktemne, crmhistorik.navn, crmemne.navn AS enmenavn, crmstatus.navn AS statusnavn, crmkontaktform.navn AS kontaktform, ikon, crmhistorik.kontaktpers FROM crmhistorik, crmemne, crmstatus, crmkontaktform, kunder, aktionsrelationer WHERE "& useidKri &" "& useemneKri &""& usestatKri &""& usemedarbKri &" AND crmemne.id = kontaktemne AND crmstatus.id = status AND crmkontaktform.id = kontaktform AND kid = kundeid ORDER BY crmhistorik.crmdato DESC"
	end select
	
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	'************************************************************************************ 
	'Table Rows indhold 
	'************************************************************************************
	%>
	<tr>
		<td bgcolor="#5582D2" colspan="10"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'');">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td height=20>&nbsp;&nbsp;<a href="kunder.asp?menu=crm&func=red&id=<%=oRec("Kid")%>"><%=left(oRec("kkundenavn"), 18) &"</a><br>&nbsp;&nbsp;<font size='1' color='#999999'>" &oRec("kontaktpers")&"</font>"%></td>
		<td><font size="1"><%=formatdatetime(oRec("crmdato"), 1)%></font></td>
		<td><a href="crmhistorik.asp?selpkt=hist&menu=crm&func=red&id=<%=oRec("Kid")%>&crmaktion=<%=oRec("id")%>" class='vmenuglobal'><u><%=left(oRec("enmenavn"), 12)%></u></a>&nbsp;</td>
		<td><a href="crmhistorik.asp?selpkt=hist&menu=crm&func=red&id=<%=oRec("Kid")%>&crmaktion=<%=oRec("id")%>" class=rmenu>
		<%if len(oRec("navn")) > 16 then
		Response.write left(oRec("navn"), 16) &"..."
		else
		Response.write oRec("navn")
		end if
		%>
		</a>&nbsp;&nbsp;</td>
		<td><font size="1"><%=left(oRec("statusnavn"), 16)%></td>
		<td>&nbsp;&nbsp;<img src="../ill/<%=oRec("ikon")%>" border="0" alt="<%=oRec("kontaktform")%>"></td>
		<td><font size="1" color="#999999"><%=left(oRec("editor"), 16)%></td>
		<td><a href="crmhistorik.asp?selpkt=hist&menu=crm&func=slet&id=<%=oRec("Kid")%>&crmaktion=<%=oRec("id")%>&ketype=e&emner=<%=emner%>&status=<%=status%>&medarb=<%=medarb%>"><img src="../ill/slet_eks.gif" width="20" height="20" alt="" border="0"></a></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="8" valign="bottom"><img src="../ill/tabel_top.gif" width="721" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>	
	</table>
	</TD>
	</TR></TABLE>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%end if%>
<%end select%>
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
