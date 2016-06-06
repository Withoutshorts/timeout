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
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	%>
	
	<%
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	func = request("func")
	
	thisfile = "sdsk"
	
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
	
	
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<%call menu_2014() %>
	
	
	<div id="sideindhold" style="position:absolute; left:90px; top:102px; visibility:visible;">
	
	
	
	
			
	<%
	
	
	
	
	
	'* Kunde
	if len(request("FM_kontakt")) <> 0 AND request("FM_kontakt") <> 0 then
	kontaktKri = " kundeid = "& request("FM_kontakt")
	selKid = request("FM_kontakt")
	else
	selKid = 0
	kontaktKri = " kundeid <> 0"
	end if
	
	'* Ansavrlig
	if len(request("FM_ansv")) <> 0 then
		if request("FM_ansv") <> 0 then
		ansvKri = " AND s.ansvarlig = "& request("FM_ansv")
		ansSel = request("FM_ansv")
		else
		ansSel = 0
		ansvKri = ""
		end if
	else
	ansSel = 0
	ansvKri = " AND s.ansvarlig <> "& ansSel
	end if
	
	'* Dato
	
	call datocookie
	stDatoSQL = strAar &"/"& strMrd &"/"& strDag
	slDatoSQL = strAar_slut &"/"& strMrd_slut &"/"& strDag_slut
	
	if request("usedatokri") = "j" then
	datoKriJa = "CHECKED"
	datoKriNej = ""
	datoKri = " AND s.dato BETWEEN '"& stDatoSQL &"' AND '"& slDatoSQL &"'"
	else
	datoKri = ""
	datoKriJa = ""
	datoKriNej = "CHECKED"
	end if
	%>	
	
	<%call filterheader_2013(0,0,600,pTxt) %>
	
        <h4>Servicedesk - Incident Statistik</h4>

	<table cellspacing=0 cellpadding=2 border=0 width=100%>
	<form action="sdsk_stat.asp" method="post">
	<tr bgcolor="#ffffff">
		<td><b>Vælg kontakt:</b></td><td colspan=2>
		<select name="FM_kontakt" style="width:240px;">
			<option value="0">Alle</option>
			<%
			
			ketypeKri = " ketype <> 'e'"
			
					strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE "& ketypeKri &" ORDER BY Kkundenavn"
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
					<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%>&nbsp;(<%=oRec("Kkundenr")%>)</option>
					<%
					oRec.movenext
					wend
					oRec.close
					%>
			</select>
		</td>
		</tr>
		<tr>
		<td><b>Ansvarlig (Intern):</b></td><td colspan=2><select name="FM_ansv" style="width:240px;">
							<option value="0">Alle</option>
							
							<%
							strSQL = "SELECT mid, mnavn FROM medarbejdere WHERE mid <> 0 AND mansat <> 2 ORDER BY mnavn"
							oRec.open strSQL, oConn, 3
							while not oRec.EOF
							if cint(ansSel) = cint(oRec("mid")) then
							isSelected = "SELECTED"
							else
							isSelected = ""
							end if%>
							<option value="<%=oRec("mid")%>" <%=isSelected%>><%=oRec("mnavn")%></option>
							<%
							oRec.movenext
							wend
							
							oRec.close
							%>
							</select>
	</td>
	</tr>
	<tr bgcolor="#ffffff">
		<td><b>Benyt datointerval:</b><br>
		<input type="radio" name="usedatokri" value="j" <%=datoKriJa%>>&nbsp;Ja&nbsp;&nbsp;
		<input type="radio" name="usedatokri" value="n" <%=datoKriNej%>>&nbsp;Nej</td>
	
		<!--#include file="inc/weekselector_b.asp"-->
	</tr>
	<tr bgcolor="#ffffff"><td colspan=3 align=right style="padding-right:50px;"><input type="submit" value=" Søg >> "></td></tr>
	</form>
	</table>
	
	
	<!-- filterDiv -->
	</td>
	</tr>
	</table>
	
	</div>
			
			
	
	
	<%
	
	tTop = 20
	tLeft = 0
	tWdth = 1004
	
    call tableDiv(tTop,tLeft,tWdth)
	
	%>
	
	<table cellspacing=5 cellpadding=2 border=0 width=100%>
	<tr>
		<td valign=top bgcolor="#eff3ff" style="padding:5px;">
		<h4>Incidents</h4>
		<b>Antal fordelt på Status:</b><br />
		
		<table cellspacing=0 cellpadding=1 border=0 width=100% bgcolor="#ffffff">
	<%
	
	intTotalall = 1
	
	antalTot = 0
	strSQL = "SELECT count(s.id) AS antal, st.navn AS navn "_
	&" FROM sdsk s"_
	&" LEFT JOIN sdsk_status st ON (st.id = s.status)"_
	&" WHERE s.id <> 0 AND "& kontaktKri &" "& ansvKri &" "& datoKri &" GROUP BY s.status ORDER BY antal DESC"
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	Response.write "<tr><td>"& oRec("navn") & ": <b>" & oRec("antal") &"</b></td></tr>"
	
	antalTot = antalTot + cint(oRec("antal"))
	oRec.movenext
	wend
	
	oRec.close
	%>
	
	<tr><td>Ialt: <b><%=antalTot%></b> incidents.</td></tr>
	</table>
	
	</td>
	
	<td bgcolor="#eff3ff" valign=top style="padding:5px;"><b>Antal fordelt på Kategori:</b><br> (Top 10 med flest incidents)<br>
	
	<table cellspacing=0 cellpadding=1 border=0 width=100% bgcolor="#ffffff">
	<%
	strSQL = "SELECT count(s.id) AS antal, tp.navn AS navn "_
	&" FROM sdsk s"_
	&" LEFT JOIN sdsk_typer tp ON (tp.id = s.type)"_
	&" WHERE s.id <> 0 AND "& kontaktKri &" "& ansvKri &" "& datoKri &" GROUP BY s.type ORDER BY antal DESC"
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	Response.write "<tr><td style='border-bottom:1px #cccccc solid;'>"& oRec("navn") & ":</td><td style='border-bottom:1px #cccccc solid;'> <b>" & oRec("antal") &"</b></td></tr>"
	
	oRec.movenext
	wend
	
	oRec.close
	%>
	</table>
	</td>
	
	
	<td valign=top bgcolor="#f3efff"><b>Antal fordelt på Type:</b><br>
	<table cellspacing=0 cellpadding=1 border=0 width=100% bgcolor="#ffffff">
	
	<%
	strSQL = "SELECT count(s.id) AS antal, p.navn AS navn "_
	&" FROM sdsk s"_
	&" LEFT JOIN sdsk_prioitet p ON (p.id = s.prioitet)"_
	&" WHERE s.id <> 0 AND "& kontaktKri &" "& ansvKri &" "& datoKri &" GROUP BY p.navn ORDER BY antal DESC"
	
	'Response.write strSQL 
	'Response.flush
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	Response.write "<tr><td>"& oRec("navn") & ": <b>" & oRec("antal") &"</b></td></tr>"
	
	oRec.movenext
	wend
	
	oRec.close
	%>
	</table>
	</td>
	
	<td valign=top bgcolor="#ffffe1"><b>Antal uden tildelt ansvarlig:</b><br>
	<table cellspacing=0 cellpadding=1 border=0 width=100% bgcolor="#ffffff">
	
	<%
	strSQL = "SELECT count(s.id) AS antal "_
	&" FROM sdsk s"_
	&" WHERE s.ansvarlig = 0 AND "& kontaktKri &" "& datoKri 
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	Response.write "<tr><td>Der er ialt: <b>" & oRec("antal") &"</b> incidents u. ansvarlige.</td></tr>"
	
	oRec.movenext
	wend
	
	oRec.close
	%>
	</table>
	</td>
	</tr>
	<tr>
	
	<td colspan=2 valign=top bgcolor="#Eff3ff">
	<br />
	<h4>Medarbejdere fordeling</h4>
	<b>Antal fordelt på medarbejdere og status:</b><br>
	<table cellspacing=0 cellpadding=1 border=0 width=100% bgcolor="#ffffff">
	<tr><td>Medarbejder</td><td>Status</td><td align=right>Antal</td><td align=right>Est. tid</td></tr>
	
	<%
	strSQL = "SELECT count(s.id) AS antal, SUM(s.esttid) AS sumtid, ansvarlig, "_
	&" s.status, m.mnavn, m.mnr, ss.navn AS statusnavn "_
	&" FROM sdsk s"_
	&" LEFT JOIN medarbejdere m on (mid = s.ansvarlig)" _
	&" LEFT JOIN sdsk_status ss on (ss.id = s.status)" _
	&" WHERE "& kontaktKri &" "& ansvKri &" "& datoKri & " GROUP BY s.ansvarlig, s.status ORDER BY s.ansvarlig" 
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	if oRec("ansvarlig") <> 0 then
	Response.write "<tr><td>"& oRec("mnavn") &" ("& oRec("mnr") &")</td>"
	else
	Response.write "<tr><td>-</td>"
	end if
	
	Response.Write "<td>"& oRec("statusnavn") &"</td>"_
	&"<td align=right>" & oRec("antal") &"</td><td align=right>"& formatnumber(oRec("sumtid"), 2) &"</td></tr>"
	
	oRec.movenext
	wend
	
	oRec.close
	%>
	</table>
	
	</td>
	<td colspan=2 valign=top bgcolor="#Eff3ff">
	<br />
	<h4>Medarbejdere Sum</h4>
	<b>Antal ialt fordelt på medarbejdere, </b><br />hvor status med i forvalgt vis på liste. (dvs. aktive status typer):<br>
	<table cellspacing=0 cellpadding=1 border=0 width=100% bgcolor="#ffffff">
	<tr><td>Medarbejder</td><td align=right>Antal</td><td align=right>Est. tid</td></tr>
	
	<%
	strSQL = "SELECT count(s.id) AS antal, SUM(s.esttid) AS sumtid, ansvarlig, "_
	&" s.status, m.mnavn, m.mnr, ss.navn AS statusnavn "_
	&" FROM sdsk s"_
	&" LEFT JOIN medarbejdere m on (mid = s.ansvarlig)" _
	&" LEFT JOIN sdsk_status ss on (ss.id = s.status)" _
	&" WHERE "& kontaktKri &" "& ansvKri &" "& datoKri & " AND ss.vispahovedliste = 1 GROUP BY s.ansvarlig ORDER BY s.ansvarlig" 
	
	'Response.Write strSQL
	'Response.flush
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	if oRec("ansvarlig") <> 0 then
	Response.write "<tr><td>"& oRec("mnavn") &" ("& oRec("mnr") &")</td>"
	else
	Response.write "<tr><td>-</td>"
	end if
	
	Response.Write ""_
	&"<td align=right>" & oRec("antal") &"</td><td align=right>"& formatnumber(oRec("sumtid"), 2) &"</td></tr>"
	
	oRec.movenext
	wend
	
	oRec.close
	%>
	</table>
	
	</td>
	</tr>
	
	
	<tr>
	<td colspan=2 bgcolor="#F3FFEF" style="padding:5px;" valign=top>
	<br />
	<h4>Påbegyndt</h4>
	
	
	
	
	
							<%
							'************************************************************************
							'** Påbegyndt
							'************************************************************************
							%>


				
							<font color=red>(Denne statistik er ikke aktiveret.)</font><br />
							<b>Antal påbegyndt til tiden / Antal påbegyndt forsent.</b><br>
							Fordelt på Type.
							<table cellspacing=0 cellpadding=2 border=0 width=100% bgcolor="#ffffff">
							<tr>
							<td valign=top><br><u>Antal påbegyndt forsent:</u><br>
							<%
							antalOvertiden = 0
							strSQL = "SELECT count(s.id) AS antal, p.responstid, p.navn AS navn "_
							&" FROM sdsk s"_
							&" LEFT JOIN sdsk_prioitet p ON (p.id = s.prioitet)"_
							&" WHERE s.id <> 0 AND "& kontaktKri &" "& ansvKri &" "& datoKri &" AND (p.responstid < s.rsptid_a_tid AND s.rsptid_a_id <> 0) GROUP BY s.prioitet ORDER BY antal DESC"
							
							'Response.write strSQL 
							'Response.flush
							oRec.open strSQL, oConn, 3
							while not oRec.EOF
							
							Response.write oRec("navn") & ": <b>" & oRec("antal") &"</b><br>"
							antalOvertiden = antalOvertiden + cint(oRec("antal"))
							oRec.movenext
							wend
							
							oRec.close
							%></td>
							
							<td valign=top><br><u>Antal påbegyndt. til tiden:</u><br>
							<%
							antalTiltiden = 0
							strSQL = "SELECT count(s.id) AS antal, p.responstid, p.navn AS navn "_
							&" FROM sdsk s"_
							&" LEFT JOIN sdsk_prioitet p ON (p.id = s.prioitet)"_
							&" WHERE s.id <> 0 AND "& kontaktKri &" "& ansvKri &" "& datoKri &" AND (p.responstid >= s.rsptid_a_tid AND s.rsptid_a_id <> 0) GROUP BY s.prioitet ORDER BY antal DESC"
							
							'Response.write strSQL 
							'Response.flush
							oRec.open strSQL, oConn, 3
							while not oRec.EOF
							
							Response.write oRec("navn") & ": <b>" & oRec("antal") &"</b><br>"
							antalTiltiden = antalTiltiden + cint(oRec("antal")) 
							oRec.movenext
							wend
							
							oRec.close
							
				
				
				if antalTiltiden <> 0 AND intTotalall <> 0 then
				'tiltidenPieStr = CStr(Round(antalTiltiden / intTotalall * 360))
				tiltidenPro = FormatPercent(antalTiltiden / intTotalall, 1) 
				else
				'tiltidenPieStr = 0
				tiltidenPro = FormatPercent(0, 1) 
				end if
				
				if antalOvertiden <> 0 AND intTotalall <> 0 then
				'overtidenPieStr = CStr(Round(antalOvertiden / intTotalall * 360))
				overtidenPro = FormatPercent(antalOvertiden / intTotalall, 1) 
				else
				'overtidenPieStr = 0
				overtidenPro = FormatPercent(0, 1) 
				end if
				%>
			    <br /><br />
				<b>Incidents ialt:  <%=intTotalall %> </b><br />
			    <B>Påbegyndt til tiden. <%= tiltidenPro %></B> <br />
				<B>Påbegyndt forsent. <%= overtidenPro %></B>
				
							</td>
							</tr>
							</table>
			    
			    
			    </td>
			    
			    <td colspan=2 bgcolor="#EFFFF3" style="padding:5px;" valign=top>
				<br />
				<h4>Afsluttet</h4>
				
				<%
				'************************************************************************
				'** Afsluttet 
				'************************************************************************
				%>
				
				
				
				
							<b>Antal incidents afsluttet forsent / til tiden.</b><br>
							Fordelt på Type.
							<table cellspacing=0 cellpadding=2 border=0 width=100% bgcolor="#ffffff">
							<tr>
								<td valign=top><br><b>Antal afsl. forsent:</b><br>
								<%
								afslForsent = 0
								strSQL = "SELECT count(s.id) AS antal, p.responstid, p.navn AS navn "_
								&" FROM sdsk s"_
								&" LEFT JOIN sdsk_prioitet p ON (p.id = s.prioitet)"_
								&" WHERE s.id <> 0 AND "& kontaktKri &" "& ansvKri &" "& datoKri &" AND (p.responstid < s.rsptid_b_tid AND s.rsptid_b_id <> 0) GROUP BY s.prioitet ORDER BY antal DESC"
								
								'Response.write strSQL 
								'Response.flush
								oRec.open strSQL, oConn, 3
								while not oRec.EOF
								
								Response.write oRec("navn") & ": <b>" & oRec("antal") &"</b><br>"
								afslForsent = afslForsent + cint(oRec("antal"))
								oRec.movenext
								wend
								
								oRec.close
								%></td>
								
								<td valign=top><br><b>Antal afsl. til tiden:</b><br>
								<%
								afslTiltiden = 0
								strSQL = "SELECT count(s.id) AS antal, p.responstid, p.navn AS navn "_
								&" FROM sdsk s"_
								&" LEFT JOIN sdsk_prioitet p ON (p.id = s.prioitet)"_
								&" WHERE s.id <> 0 AND "& kontaktKri &" "& ansvKri &" "& datoKri &" AND (p.responstid >= s.rsptid_b_tid AND s.rsptid_b_id <> 0) GROUP BY s.prioitet ORDER BY antal DESC"
								
								'Response.write strSQL 
								'Response.flush
								oRec.open strSQL, oConn, 3
								while not oRec.EOF
								
								Response.write oRec("navn") & ": <b>" & oRec("antal") &"</b><br>"
								afslTiltiden = afslTiltiden + cint(oRec("antal"))
								oRec.movenext
								wend
								
								oRec.close
								
				                if afslTiltiden <> 0 AND intTotalall <> 0 then
				                'afslTiltidenPieStr = CStr(Round(afslTiltiden / intTotalall * 360))
				                afslTiltidenPro = FormatPercent(afslTiltiden / intTotalall, 1) 
				                else
				                'afslTiltidenPieStr = 0
				                afslTiltidenPro = FormatPercent(0, 1) 
				                end if
                				
				                if afslForsent <> 0 AND intTotalall <> 0 then
				                'afslForsentPieStr = CStr(Round(afslForsent / intTotalall * 360))
				                afslForsentPro = FormatPercent(afslForsent / intTotalall, 1) 
				                else
				                'afslForsentPieStr = 0
				                afslForsentPro = FormatPercent(0, 1) 
				                end if
				                %>
								<br /><br />
								<b>Incidents ialt:  <%=intTotalall %> </b><br />
								<B>Afsluttet til tiden: <%= afslTiltidenPro %></B><br />
								<B>Afsluttet forsent: <%= afslForsentPro %></B>
								</td>
								
								</tr>
							</table>
							
				
				
				
				
                
                </td>
	
	</tr>
	
	
	</table>
	
				
				
			
				</div><!--table div-->
				
				
				
				


<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
				&nbsp;<br />
	
	
</div>
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
	
	
