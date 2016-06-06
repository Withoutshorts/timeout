<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<!--#include file="inc/sdsk_inc.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->


<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	%>
	
	<%
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	func = request("func")
	
	thisfile = "sdsk_knowledge.asp"
	
	if func = "dbopr" then
	id = 0
	else
		if len(request("id")) <> 0 then
		id = request("id")
		else
		id = 0
		end if
	end if
	
	dato = day(now)&"/"&month(now)&"/"&year(now)
	tid = time
	
	
	
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
		
		function SQLBless3(s, sogeord)
			tmpthis = s
			tmpthis = replace(lcase(tmpthis), ""&lcase(sogeord)&"", "<font class=bq>"& sogeord &"</font>")
			SQLBless3 = tmpthis
		end function
	%>
	

<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->

	<script>
	function showbesk(besk_id) {
	var lastOpen = document.getElementById("divopen").value
	
	if (lastOpen != 0) {
	document.getElementById("b_"+lastOpen).style.display = "none";
	document.getElementById("b_"+lastOpen).style.visibility = "hidden";
	}
	
	document.getElementById("b_"+besk_id).style.display = "";
	document.getElementById("b_"+besk_id).style.visibility = "visible";
	
	document.getElementById("FM_div_id_txtbox_"+besk_id).focus()
	//document.getElementById("FM_div_id_txtbox_"+besk_id).scrollintoview()

	
	document.getElementById("divopen").value = besk_id
	}
	
	
	function hidebesk(besk_id) {
	document.getElementById("b_"+besk_id).style.display = "none";
	document.getElementById("b_"+besk_id).style.visibility = "hidden";
	document.getElementById("divopen").value = besk_id
	}
	
	</script>
	
	
	
	
	
	<%
	

        call menu_2014()

	
	'* Type
	if len(request("FM_type")) <> 0 AND request("FM_type") <> 0 then
	typeKri = " AND s.type = "& request("FM_type")
	typeSel = request("FM_type")
	else
	typeSel = 0
	typeKri = ""
	end if
	
	
	
	'* Kunde
	if len(request("FM_kontakt")) <> 0 AND request("FM_kontakt") <> 0 then
	kontaktKri = " s.kundeid = "& request("FM_kontakt")
	selKid = request("FM_kontakt")
	else
	selKid = 0
	kontaktKri = " s.kundeid <> 0"
	end if
	

	
	
	'* sogeKri
	if len(request("FM_sog")) <> 0 then
	showSogeKri = request("FM_sog") 
	sogeKri = " AND (s.emne LIKE '"& showSogeKri &"%' OR s.besk LIKE '%"& showSogeKri &"%' OR s.sogeord1 LIKE '"& showSogeKri &"%' OR s.sogeord2 LIKE '"& showSogeKri &"%' OR s.sogeord3 LIKE '"& showSogeKri &"%' OR s.sogeord4 LIKE '"& showSogeKri &"%')"
	else
	showSogeKri = ""
	sogeKri = ""
	end if
	
	
	
	%>
	
	<div id="sideindhold" style="position:absolute; left:90px; top:102px; visibility:visible;">
	
	<h4>Servicedesk - Knowledgebase</h4>
	<div style="position:relative; width:500px; padding:20px; background-color:#ffffff;">
	<table cellspacing=0 cellpadding=2 border=0 width=100%>
	<form action="sdsk_knowledge.asp" method="post">
	<tr><td valign=top bgcolor="#ffffff" style="padding-top:10px;">
	
	<table cellspacing=0 cellpadding=4 border=0>
	<tr>
		<td align=right style="padding-right:5px;"><b>Søg på emne ell. søgeord:</b></td><td><input type="text" name="FM_sog" id="FM_sog" value="<%=showSogeKri%>" style="width:240px;"></td>
	</tr>
	<tr>
		<td align=right style="padding-right:5px;"><b>Vælg kontakt:</b></td><td>
		<select name="FM_kontakt" style="width:240px; ">
			<option value="0">Alle</option>
			<%
			
			ketypeKri = " ketype <> 'e'"
			
					strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE "& ketypeKri &" AND sdskpriogrp <> 0 ORDER BY Kkundenavn"
					'Response.write strSQL
					'Response.flush
					oRec.open strSQL, oConn, 3
					while not oRec.EOF
					
					if cint(selKid) = cint(oRec("Kid")) then
					isSelected = "SELECTED"
					else
					isSelected = ""
					end if
					%>
					<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%> (<%=oRec("Kkundenr")%>)</option>
					<%
					oRec.movenext
					wend
					oRec.close
					%>
			</select>
		</td>
	</tr>
	<tr>
		<td align=right style="padding-right:5px;"><b>Type:</b></td><td><select name="FM_type" style=width:240px;">
							<option value="0">Alle</option>
							
							<%
							strSQL = "SELECT id, navn FROM sdsk_typer WHERE id <> 0 ORDER BY navn"
							oRec.open strSQL, oConn, 3
							while not oRec.EOF
							if cint(typeSel) = cint(oRec("id")) then
							isSelected = "SELECTED"
							else
							isSelected = ""
							end if%>
							<option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("navn")%></option>
							<%
							oRec.movenext
							wend
							
							oRec.close
							%>
							</select>
	</td></tr>
	</table>
	
	</td>
	</tr>
	<tr bgcolor="#ffffff"><td colspan=2 align=right style="padding-right:50px;"><input type="submit" value="  Søg  "></td></tr>
	</form>
	</table>
	</div>
	
	 <br><br>
	
	
	<%
	
	strSQL = "SELECT s.editor, s.public, s.id, s.emne, s.besk, s.dato, s.tidspunkt, p.navn AS pnavn, p.responstid, s.type, s.ansvarlig, "_
	&" t.navn AS type, k.kkundenavn, k.kkundenr, k.kid, s.kundeid, "_
	&" s.prioitet, s.status, st.navn AS statusnavn, st.farve, s.sogeord1, s.sogeord2, s.sogeord3, s.sogeord4 FROM sdsk s "_
	&" LEFT JOIN sdsk_prioitet p ON (p.id = s.prioitet) "_
	&" LEFT JOIN sdsk_typer t ON (t.id = s.type) "_
	&" LEFT JOIN kunder k ON (k.kid = s.kundeid) "_
	&" LEFT JOIN sdsk_status st ON (st.id = s.status) "
	
	if request("visikke") <> "1" then
	strSQL =  strSQL &" WHERE "& kontaktKri  & typeKri & ansvKri & sogeKri & datoKri & statusKri &""
	else
	strSQL =  strSQL &" WHERE s.id = 0"
	end if
	
	strSQL =  strSQL &" ORDER by id DESC"
	
	'response.write strSQL
	'response.flush
	
	i = 0
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
		
		
		
		oprettet = oRec("dato") &" "& formatdatetime(oRec("tidspunkt"), 3)
		'nu = dato &" "& tid
		'responsTidspunkt = dateadd("h", oRec("responstid"), oprettet)
		'timerTilrespons = dateDiff("h", nu, responsTidspunkt)
		
	
	%>
	<table cellspacing=0 cellpadding=2 border=0 bgcolor="#ffffff" width=900>
	<tr>
	<td bgcolor="#eff3ff" style="border-top:1px #8caae6 solid; border-left:1px #8caae6 solid; padding-left:20px; padding-top:12px; padding-bottom:15px;">
	ID: <b><%=oRec("id")%></b>&nbsp;&nbsp;Dato: <%=oRec("dato")%>&nbsp;<%=formatdatetime(oRec("tidspunkt"), 3)%>&nbsp;(<%=oRec("editor")%>)<br>
	Type: <b><%=oRec("type")%></b><br>
	Emne: <b>
	<%if len(oRec("emne")) <> 0 then%>
	<%=SQLBless3(oRec("emne"), showSogeKri)%>
	<%else%>
	<%=oRec("emne")%>
	<%end if%> </b><br>
	Kontakt: <b><%=oRec("kkundenavn")%> (<%=oRec("kkundenr")%>)</b> <br>
	Status: <b><%=oRec("statusnavn")%></b><br>
	Søgeord: 
	<%if len(trim(oRec("sogeord1"))) <> 0 then%>
	&nbsp;<%=SQLBless3(oRec("sogeord1"), showSogeKri)%>
	<%end if%>
	
	<%if len(trim(oRec("sogeord1"))) <> 0 then%>
	, <%=SQLBless3(oRec("sogeord2"), showSogeKri)%>
	<%end if%>
	
	<%if len(trim(oRec("sogeord1"))) <> 0 then%>
	, <%=SQLBless3(oRec("sogeord3"), showSogeKri)%> 
	<%end if%>
	
	<%if len(trim(oRec("sogeord1"))) <> 0 then%>
	, <%=SQLBless3(oRec("sogeord4"), showSogeKri)%> 
	<%end if%>
	
	<br><br>
	<b>Beskrivelse:</b><br>
	<%if len(oRec("besk")) <> 0 then%>
	<%=SQLBless3(oRec("besk"), showSogeKri)%>
	<%else%>
	<%=oRec("besk")%>
	<%end if%> 
	</td>
	<td bgcolor="#eff3ff" valign=top width=300 style="border-top:1px #8caae6 solid; border-right:2px #5582d2 solid; padding-left:20px; padding-right:20px;  padding-top:12px; padding-bottom:15px;">
		<table>
		<%
			strSQL2 = "SELECT id, filnavn FROM filer WHERE incidentid = " & oRec("id") 
			oRec2.open strSQL2, oConn, 3
			while not oRec2.EOF
			%>
			<tr><td><img src="../ill/addmore55.gif" width="10" height="13" alt="" border="0">&nbsp;<a href="../inc/upload/<%=lto%>/<%=oRec2("filnavn")%>" class='vmenulille' target="_blank"><%=oRec2("filnavn")%></a>&nbsp;</td></tr>
			<%
			oRec2.movenext
			wend
			
			oRec2.close
	%>
		</table>
	</td>
	</tr>
	
	
	<%
	strSQL2 = "SELECT besk, sdskdato, sdsktidspunkt, id, editor, public, "_
	&" losning FROM sdsk_rel WHERE sdsk_rel = "& oRec("id") &" ORDER BY id DESC"
	oRec2.open strSQL2, oConn, 3
	while not oRec2.EOF
	%>
	<tr><td colspan=2 style="border-left:1px #8caae6 solid; border-right:2px #5582d2 solid; border-bottom:1px #999999 dashed; padding-top:10px; padding-bottom:5px; padding-left:50px;">
	<%if oRec2("losning") = 1 then%>
	<img src="../ill/ac0060-16.gif" width="16" height="16" alt="Godkendt som løsning" border="0">
	<%end if%>
	
	Dato: <b><%=oRec2("sdskdato")%> Kl. <%=formatdatetime(oRec2("sdsktidspunkt"), 3)%></b> (<%=oRec2("editor")%>)<br>
	<%=SQLBless3(oRec2("besk"), showSogeKri)%><br>
	</td>
	</tr>
	<%
	oRec2.movenext
	wend
	
	
	oRec2.close
	
	
	%>
	<tr><td colspan=2 style="border-bottom:2px #5582d2 solid; border-right:2px #5582d2 solid; border-left:1px #8caae6 solid;">&nbsp;</td></tr>
	</table>
	<br><br>
	<%
	
	i = i + 1
	oRec.movenext
	wend
	
	oRec.close
	
	%>
	
	<br>
	Der er ialt fundet: <b><%=i%></b> Incidents.
	<br>
	<br>
	*) Timer til respons.
	
	<form><input type="hidden" name="divopen" id="divopen" value="<%=lastopenDiv%>"></form>
	
	</div>
	
	
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
	
	
