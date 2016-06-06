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
	if func = "dbopr" OR func = "opret" then
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
	<div id="sindhold" style="position:absolute; left:210; top:150; background-color:#ffffff; visibility:visible; padding:20px;">
	<h4>Standard Produktioner</h4>
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> en Standard Produktion. Er dette korrekt?</td>
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
	oConn.execute("DELETE FROM produktion_mat_rel WHERE produktionsId = "& id &"")
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
		strGrpNr = SQLBless(request("FM_nr"))
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO produktioner (navn, editor, dato, nummer) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', '"& strGrpNr &"')")
			
			
			lastId = 0
			
			'*** Finder netop oprettede id ***
			strSQL = "SELECT id FROM produktioner ORDER BY id DESC"
			oRec.open strSQL, oConn, 3 
			if not oRec.EOF then
				lastId = oRec("id")
			end if
			oRec.close 
			
		else
		oConn.execute("UPDATE produktioner SET navn ='"& strNavn &"', "_
		&" nummer = '"& strGrpNr &"', editor = '" &strEditor &"', dato = '" & strDato &"' WHERE id = "&id&"")
		
			lastId = id
		
		end if
		
		'*** produktionerelationer ***
		oConn.execute("DELETE FROM produktion_mat_rel WHERE produktionsId = "& lastId &"")
		
		'*** Opdaterer db **********
		strproduktion_mat_rel = split(request("FM_produktion_mat_rel"), ",")
		For b = 0 to Ubound(strproduktion_mat_rel)
			
			'** Afhængighed ***
			afhidThis = 0
			afhidThis = request("FM_afh_"&trim(strproduktion_mat_rel(b))&"")
			if len(afhidThis) <> 0 then
			afhidThis = afhidThis
			else
			afhidThis = 0
			end if
			
			strSQL = "INSERT INTO produktion_mat_rel (produktionsId, materialeId, afhId) VALUES ("& lastId &", "& strproduktion_mat_rel(b) &", "& afhidThis &")"
			'Response.write strSQL
			'Response.flush
			oConn.execute(strSQL)
		Next	
		
		
		Response.redirect "produktioner.asp?menu=mat&shokselector=1&id="&lastId
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	strGrpNr = ""
	sidst_redigeret = ""
	
	else
	strSQL = "SELECT navn, editor, dato, nummer FROM produktioner WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	strGrpNr = oRec("nummer")
	
	end if
	oRec.close
	
	sidst_redigeret = "Sidst opdateret den <b>"&formatdatetime(strDato, 1)&"</b> af <b>"&strEditor&"</b>"
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
<!--
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	
	<%call tsamainmenu(11)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	if showonejob <> 1 then
		call mattopmenu()
	end if
	%>
	</div>
    -->

    <%call menu_2014() %>

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; width:600px; visibility:visible; padding:20px; background-color:#ffffff;">
	<h4>Standard Produktioner <span style="font-size:9px;">- <%=varbroedkrumme%></span></h4>
	<table cellspacing="0" cellpadding="2" border="0" width="100%" bgcolor="#ffffff">
	<form action="produktioner.asp?menu=mat&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">

  	<tr>
		
		<td colspan="2" valign="middle" style="height:30;"><%=sidst_redigeret%>&nbsp;</td>
		
	</tr>
	<tr>
		
		<td align=right><font color=red><b>*</b></font>&nbsp;Navn:</td>
		<td style="padding-left:5px;"><input type="text" name="FM_navn" value="<%=strNavn%>" size="30" /></td>
		
	</tr>
	<tr>
		
		<td align=right>Produktions id / nr:</td>
		<td style="padding-left:5px;"><input type="text" name="FM_nr" value="<%=strGrpNr%>" size="30"></td>
		
	</tr>
	<tr>
		
		<td colspan=2><br>
		<b>Vælg hvilke materialer der indgår i denne produktion:</b><br><br>
		<div id=mate name=mate style="position:relative; left:0; top:0; height:200px; overflow:auto;">
		<table cellspacing=0 cellpadding=0 border=0>
				<tr bgcolor="#c4c4c4"><td height=20 width=120>&nbsp;<b>Materiale</b></td><td width=80><b>Mat. nr</b></td><td width=60><b>På lager</b></td><td width=110><b>Arr.dato/Prod.tid</b></td><td width=130><b>Afhængig af:</b></td></tr>
			
		<%
		strSQL = "SELECT m.matgrp, m.id AS id, m.varenr, m.navn, m.ptid, "_
		&"m.ptid_arr, m.arrestordredato, m.antal, mg.navn AS grpnavn FROM materialer m LEFT JOIN materiale_grp mg ON (mg.id = m.matgrp) WHERE m.id <> 0 ORDER BY m.matgrp, m.navn"
		oRec.open strSQL, oConn, 3 
		t = 1
		lastmatgrp = 0
		matListe = ""
		while not oRec.EOF 
		
		if cint(lastmatgrp) <> oRec("matgrp") then
		%><tr><td colspan=5 bgcolor="#d6dff5">&nbsp;<b><%=oRec("grpnavn")%></b></td></tr>
		<%end if
		
		
		select case right(t, 1)
		case 1, 3, 5, 7, 9
		bgThis = "#ffffff"
		case else
		bgThis = "#ffffff"
		end select 
		
			
			foundrel = 0
			AfhId = 0
			strSQL2 = "SELECT id, afhid FROM produktion_mat_rel WHERE produktionsId = "& id &" AND materialeId = "& oRec("id")
			oRec2.open strSQL2, oConn, 3 
			'Response.write strSQL2
			if not oRec2.EOF then
			foundrel = 1
			AfhId = oRec2("afhid")
			end if
			oRec2.close 
			'Response.write "<br>"& AfhId
			
			
			if foundrel = 1 then
			chkthis = "checked"
			matListe = matListe & "<li>"& oRec("navn") & " (" & oRec("varenr") &")<br>"
			else
			chkthis = ""
			end if%>
		<tr bgcolor="<%=bgThis%>">
		<td class=lille><input type="checkbox" name="FM_produktion_mat_rel" id="FM_produktion_mat_rel" value="<%=oRec("id")%>" <%=chkthis%>><%=oRec("navn")%>&nbsp;</td>
		<td class=lille><%=oRec("varenr")%></td>
		<td class=lille>
		<%=oRec("antal")%> stk.
		</td>
		<td class=lille>
		<%if oRec("antal") <> 0 then
		Response.write "("
		end if%>
		<%if oRec("ptid_arr") = 1 then
			if len(oRec("arrestordredato")) <> 0 then
			Response.write formatdatetime(oRec("arrestordredato"), 1)
			end if
		else
			Response.write oRec("ptid") & " dage"
		end if%>
		<%if oRec("antal") <> 0 then
		Response.write ")"
		end if%>
		</td>
		<td><select name="FM_afh_<%=oRec("id")%>" id="FM_afh_<%=oRec("id")%>" style="width:100px; font-size:9px;">
		<option value="0">Ingen</option>
		<%
		strSQL2 = "SELECT id, navn, varenr FROM materialer WHERE id <> 0 ORDER BY navn"
		oRec2.open strSQL2, oConn, 3 
		while not oRec2.EOF 
		
		if cint(afhId) = oRec2("id") then
		selAfhThis = "SELECTED"
		else
		selAfhThis = ""
		end if%>
		<option value="<%=oRec2("id")%>" <%=selAfhThis%>><%=oRec2("navn")%></option>
		<%
		oRec2.movenext
		wend
		oRec2.close 
		%>
		</select>
		<%
		
		if cint(afhId) <> 0 then%>
		<font class="roed"><i><b>V</b></i></font>
		<%end if%>
		</td>
		</tr>
		
		<%
		lastmatgrp = oRec("matgrp")
		t = t + 1
		oRec.movenext
		wend
		oRec.close 
		%>
		</table>
		</div></td>
	
	</tr>
	
	
	<tr>
		<td colspan="2" align=right><br><br><input type="image" src="../ill/<%=varSubVal%>.gif"></td>
	</tr>
	</form>
	</table>
	<div id=matliste name=matliste style="position:absolute; left:660px; top:0px; background-color:#ffffff; border:0px #003399 solid; width:200px; padding:20px;">
	<b>Materialeliste:</b><br>
	<br>
	<ul>
	<%=matListe%>
	</ul>
	<br><br>
	<font class=megetlillesort>(Produktion skal opdateres før ændringer vises på materialelisten)</font>
	</div>
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
	

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; visibility:visible; background-color:#FFFFFF; width:800px; padding:20px;">
	<h4>Standard Produktioner</h4>
	<table cellspacing="0" cellpadding="0" border="0" width="100%"><form action="produktioner.asp?menu=mat" method="post">
	<tr>
		<!--
		<td><b>Søg efter gruppe:</b>&nbsp;&nbsp;(Gruppenavn ell. nummer)<br>
		<input type="text" name="FM_sog" id="FM_sog" value="<%=request("FM_sog")%>" style="width:200px;"> <input type="image" src="../ill/pilstorxp.gif"></td>
		-->
		<td align=right><a href="produktioner.asp?menu=mat&func=opret">Ny Standard Produktion <img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a><br></td>
	</tr>
	<tr><td><br>
	<b>På lager?</b> Denne produktion er umiddelbar klar til levering.<br><br>
	<b>På lager om: (Estimeret*)</b> Hvis standard produktionen ikke er på lager, angives den dato hvor produktionen er klar til levering.
	Datoen beregnes udfra de materialer der indgår i produktionen. 
	Hvis et materiale er i restordre frem til d. 12 dec. og har en intern produktions tid på 3 dage, vil produktionen være klar d. 15. dec.
	<br><br><font class=megetlillesort>*) Hvis flere materialer er i restordre benyttes kun den restordredato der er længst ude i fremtiden. Ved de andre materialer adderes kun deres produktionstid. </font>
	</td></tr></form>
	</table>
	<br>
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	
	
	<tr bgcolor="#5582D2">
		<td><a href="produktioner.asp?menu=mat&sort=nr" class=alt><b>Prod. id / nr.</b></a></td>
		<td><a href="produktioner.asp?menu=mat&sort=navn" class=alt><b>Navn</b></a></td>
		<td class=alt>&nbsp;På lager?</td>
		<td class=alt>&nbsp;På lager om:</td>
		<td class=alt>&nbsp;</td>
        
	</tr>
	<%
	sort = Request("sort")
	if sort = "navn" then
	strSQL = "SELECT id, navn, nummer FROM produktioner ORDER BY navn"
	else
	strSQL = "SELECT id, navn, nummer FROM produktioner ORDER BY nummer"
	end if
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
	if cint(id) = oRec("id") then
	bgthis = "#ffffe1"
	else
	bgthis = "#eff3ff"
	end if%>
	
	<tr bgcolor="<%=bgthis%>">
		
		<td><%=oRec("nummer")%></td>
		<td ><a href="produktioner.asp?menu=mat&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
		<td>&nbsp;
			<%
			'*** Finder materialer i produktionen ***
			strSQL2 = "SELECT pmr.id, pmr.afhid, m.antal, m.navn, m.ptid, m.arrestordredato, m.ptid_arr FROM produktion_mat_rel pmr"_
			&" LEFT JOIN materialer m ON (m.id = pmr.materialeId) WHERE produktionsId = "& oRec("id") &" AND m.antal = 0 GROUP BY m.id ORDER BY arrestordredato DESC"
			
			'Response.write strSQL2
			'Response.flush
			
			'*** Estimerer produktions tid. *** 
			'*** Kun de materiale med den højeste arrestordre dato benyttes. De andre opsumeres kun med deres produktionstid.
			
			oRec2.open strSQL2, oConn, 3 
			palager = 1
			ptid = 0
			x = 0
			while not oRec2.EOF 
			
			'*** Finder højeste arrestordre dato ***
			if x = 0 then
			arrestordredatoDageHojeste = datediff("d", now, oRec2("arrestordredato"))
			else
			arrestordredatoDageDenne = datediff("d", now, oRec2("arrestordredato"))
			
				if arrestordredatoDageDenne > arrestordredatoDageHojeste AND oRec2("ptid_arr") <> 0 then
				arrestordredatoDageHojeste = arrestordredatoDageDenne
				else
				arrestordredatoDageHojeste = arrestordredatoDageHojeste 
				end if
				
			end if
			'*****************************************	
				
				if oRec2("antal") > 0 then
					
					palager = palager
					ptid = ptid
				
				else
					
					palager = 0
					
					'*** Afhængig af ***
					if oRec2("afhid") <> 0 then
						
						'*** Finder produktionstid på det afhængige materiale ***
						strSQL3 = "SELECT ptid, arrestordredato, ptid_arr, antal, ptid, navn FROM materialer WHERE id = " & oRec2("afhid")
						oRec3.open strSQL3, oConn, 3 
						'Response.write strSQL3
						'Response.flush
						if not oRec3.EOF then
						
							if oRec3("antal") = 0 then
								
								if oRec3("ptid_arr") = 0 then
								ptid = ptid + oRec3("ptid")
								else
								
								dagetil = datediff("d", now, oRec3("arrestordredato"))
								'Response.write "#"& dagetil &" > "& arrestordredatoDageHojeste & "<br>"
								if dagetil > arrestordredatoDageHojeste then
								ptid = ptid + dagetil + oRec3("ptid")
								else
								ptid = ptid + oRec3("ptid")
								end if
								
								end if
							
							else '** Afhængig materiale findes på lager ***
							ptid = ptid
							end if
						end if
						
						'Response.write "afh:" & oRec3("navn") & " ptid: "& ptid  &" useArrd: "& oRec3("ptid_arr") &" aDato: "& oRec3("arrestordredato") & "<br>" 
			
						oRec3.close 
						
					end if
						
						'** Plus det oprindelige materiale (hvis det er afhængigt af andet materiale) ***
						if oRec2("ptid_arr") <> 0 then
								
								dagetil = datediff("d", now, oRec2("arrestordredato"))
								'Response.write "#"& dagetil &" > "& arrestordredatoDageHojeste & "<br>"
								
								if dagetil > arrestordredatoDageHojeste OR x = 0 then
								ptid = ptid + dagetil + oRec2("ptid")
								else
								ptid = ptid + oRec2("ptid")
								end if
								
						else
						ptid = ptid + oRec2("ptid")
						end if
					
					
				end if 'antal
			
			'Response.write oRec2("navn") & " ptid: "& ptid  &" useArrd: "& oRec2("ptid_arr")  &" aDato:"& oRec2("arrestordredato") & "<br>" 
			
			x = x + 1
			oRec2.movenext
			wend
			oRec2.close 
			
			if palager = 1 then
			%>
			Ja
			<%else%>
			Nej
			<%end if%>
		</td>
		
		<td>&nbsp;ca. <b><%=ptid%></b> dage. (<%=formatdatetime(dateadd("d", ptid, now), 1)%>)</td>
		<td><a href="produktioner.asp?menu=mat&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="" border="0"></a></td>
	
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>	
	
	</table>
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
