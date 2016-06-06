<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
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
	case "slet", "sletfolder"
		
		
		gruppeid = request("gruppeid") 
		gruppenavn = request("gruppenavn")	
		
		if func = "sletfolder" then
		folderellergruppe = 1
		sletfunc = "sletfolderok"
		txt = "folder"
		else
		folderellergruppe = 0
		sletfunc = "sletok"
		txt = "standardfolder gruppe"
		end if
	
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190px; top:102px; visibility:visible;">
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> en <b><%=txt%></b>. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><a href="stfolder_gruppe.asp?menu=job&func=<%=sletfunc%>&id=<%=id%>&gruppeid=<%=gruppeid%>&gruppenavn=<%=gruppenavn%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	'*** Her slettes ***
	case "sletok", "sletfolderok"
	
	if func = "sletfolderok" then
	
	gruppeid = request("gruppeid") 
	gruppenavn = request("gruppenavn")	
	
	oConn.execute("DELETE FROM foldere WHERE id = "& id &"")
	Response.redirect "stfolder_gruppe.asp?menu=job&func=visgruppe&gruppeid="&gruppeid&"&gruppenavn="&gruppenavn&"&id=0"
			
	else
	
	oConn.execute("DELETE FROM folder_grupper WHERE id = "& id &"")
	Response.redirect "stfolder_gruppe.asp?menu=job"
	end if
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		
		<%
		errortype = 8
		useleftdiv = "c"
		call showError(errortype)
		
		else
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		if len(request("folderellergruppe")) <> 0 then
		folderellergruppe = request("folderellergruppe")
		else
		folderellergruppe = 0
		end if
		
					if len(request("FM_kundese")) <> 0 then
					kundese = request("FM_kundese")
					else
					kundese = 0
					end if
			
					gruppeid = request("gruppeid") 
					gruppenavn = request("gruppenavn")	
		
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		if func = "dbopr" then
				if folderellergruppe = 1 then
					oConn.execute("INSERT INTO folder_grupper (navn, editor, dato) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"')")
					
					'**** Finder id ****
					strSQL = "SELECT id FROM folder_grupper WHERE id <> 0 ORDER BY id DESC"
					oRec.open strSQL, oConn, 3
					if not oRec.EOF then
					
					id = oRec("id")
					
					end if
					oRec.close
				
				else
					strSQL = "INSERT INTO foldere (kundese, navn, editor, dato, stfoldergruppe) VALUES ("& kundese &", '"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& gruppeid &")"
					oConn.execute(strSQL)
					
					
					'**** Finder id ****
					strSQL = "SELECT id FROM foldere WHERE id <> 0 ORDER BY id DESC"
					oRec.open strSQL, oConn, 3
					if not oRec.EOF then
					
					id = oRec("id")
					
					end if
					oRec.close
					
				end if
		else
				if folderellergruppe = 1 then
					oConn.execute("UPDATE folder_grupper SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"' WHERE id = "&id&"")
				else
					oConn.execute("UPDATE foldere SET kundese = "& kundese &", navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"' WHERE id = "&id&"")
				end if
		end if
		
			if folderellergruppe = 1 then
			Response.redirect "stfolder_gruppe.asp?menu=job&shokselector=1&id="&id
			else
			Response.redirect "stfolder_gruppe.asp?menu=job&func=visgruppe&gruppeid="&gruppeid&"&gruppenavn="&gruppenavn&"&id="&id
			end if
		end if
	
	
	
	
	case "opret", "red", "opretfolder", "redfolder"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" OR func = "opretfolder" then
		
		strNavn = ""
		strTimepris = ""
		varSubVal = "Opret" 
		kundese = 0
		dbfunc = "dbopr"
		
	else
		
		if func = "redfolder" then
			
		strSQL = "SELECT kundese, navn, editor, dato FROM foldere WHERE id=" & id
		oRec.open strSQL,oConn, 3
		if not oRec.EOF then
		strNavn = oRec("navn")
		strDato = oRec("dato")
		strEditor = oRec("editor")
		kundese = oRec("kundese")
		end if
		oRec.close
		
		else
		
		
		strSQL = "SELECT navn, editor, dato FROM folder_grupper WHERE id=" & id
		oRec.open strSQL,oConn, 3
		if not oRec.EOF then
		strNavn = oRec("navn")
		strDato = oRec("dato")
		strEditor = oRec("editor")
		end if
		oRec.close
		
		end if
		
		dbfunc = "dbred"
		varSubVal = "Opdater" 

	end if
	
	
	if func = "opret" OR func = "red" then
		folderellergruppe = 1
			if func = "opret" then
			varbroedkrumme = " grupper -- Opret ny"
			else
			varbroedkrumme = " grupper -- Rediger"
			end if
		txt1 = "Gruppe:"
		formfelt2 = ""
	else
		folderellergruppe = 0
			
			if func = "opretfolder" then
			varbroedkrumme = " -- Opret ny"
			else
			varbroedkrumme = " -- Rediger"
			end if
		
		txt1 = "Folder:"
		
		if kundese = 1 then
		kundeseCHK = "CHECKED"
		else
		kundeseCHK = ""
		end if
		
		formfelt2 = "<br><input type='checkbox' name='FM_kundese' id='FM_kundese' value='1' "& kundeseCHK &"> Folder tilgængelig for eksterne kontakter."
		gruppeid = request("gruppeid")
	end if
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->

    <%call menu_2014() %>

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; visibility:visible;">
	<h4>Standardfolder <%=varbroedkrumme%></h4>
	<table cellspacing="0" cellpadding="0" border="0" width="400">
	<tr><form action="stfolder_gruppe.asp?menu=crm&func=<%=dbfunc%>&folderellergruppe=<%=folderellergruppe%>&gruppeid=<%=gruppeid%>&gruppenavn=<%=gruppenavn%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
    	<td valign="top" colspan="2">&nbsp;</td>
	</tr>
	<%if dbfunc = "dbred" then%>
	<tr>
		<td colspan="2" valign="bottom" style="height:30;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></td>
	</tr>
	<%end if%>
	<tr>
		<td valign=top align=right style="padding:5px;"><font class=roed><b>*</b></font>&nbsp;<%=txt1%></td>
		<td valign=top><input type="text" name="FM_navn" value="<%=strNavn%>" size="30">
		<%=formfelt2%>
		</td>
	</tr>
	<tr>
		<td colspan="2" align="right">
		<input type="submit" value="<%=varSubVal%> >>"></td>
	</tr>
	</form>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	
	
	
	
	<%case "visgruppe"
	
	gruppenavn = request("gruppenavn")
	gruppeid = request("gruppeid")
	
	
	
	%> 
        <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <%call menu_2014() %>

	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; visibility:visible;">
	<h4>Foldere i gruppen: <%=gruppenavn%></h4>
	<br>
	
	<a href="stfolder_gruppe.asp?menu=job&func=opretfolder&gruppeid=<%=gruppeid%>&gruppenavn=<%=gruppenavn%>">Opret ny folder&nbsp;&nbsp;<img src="../ill/pil_groen_lille.gif" width="17" height="17" alt="" border="0"></a><br>
	<br>
	<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#EFF3FF">
 	<tr bgcolor="5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=4 valign="top"><img src="../ill/tabel_top.gif" width="587" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class='alt'><b>Id</b></td>
		<td class='alt'><b>Folder</b></td>
		<td class='alt'><b>Ekstern adgang?</b></td>
		<td class='alt'><b>Slet folder?</b></td>
	</tr>
	<%
	strSQL = "SELECT id, navn, kundese FROM foldere WHERE stfoldergruppe = "& gruppeid &" ORDER BY navn"
	'Response.write strSQL 
	'Response.flush
	c = 0
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	%>
	<tr>
		<td bgcolor="#5582D2" colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<%
	if cint(id) = oRec("id") then
		bgc = "#ffff99"
	else
		select case right(c, 1)
		case 0, 2, 4, 6, 8
		bgc = "#ffffff"
		case else
		bgc = "#EFF3FF"
		end select
	
	end if
	%>
	<tr bgcolor=<%=bgc%>>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td><%=oRec("id")%></td>
		<td height="20"><img src="../ill/folder_ikon.gif" width="17" height="15" alt="" border="0">&nbsp;<a href="stfolder_gruppe.asp?menu=job&func=redfolder&gruppeid=<%=gruppeid%>&gruppenavn=<%=gruppenavn%>&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
		<td><%if oRec("kundese") = 1 then%>
		<b><i>V</i></b>
		<%else%>
		&nbsp;
		<%end if%>
		</td>
		<td><a href="stfolder_gruppe.asp?menu=job&func=sletfolder&gruppeid=<%=gruppeid%>&gruppenavn=<%=gruppenavn%>&id=<%=oRec("id")%>"><img src="../ill/slet_eks.gif" width="20" height="20" alt="" border="0"></a></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	c = c + 1
	x = 0
	oRec.movenext
	wend
	%>	
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="4" valign="bottom"><img src="../ill/tabel_top.gif" width="587" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="stfolder_gruppe.asp?id=<%=gruppeid%>"><img src="../ill/soeg-knap_tilbage.gif" width="16" height="16" alt="" border="0">&nbsp;Tilbage til oversigt over standardfolder grupper</a>
	<br>
	<br>
	&nbsp;
	</div>
	
	
	
	
	
	<%case else%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->

    <%call menu_2014() %>
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; visibility:visible;">
	<h4>Standardfolder grupper</h4>
	<br />
	Standard foldere er en gruppe af filmapper der til tilknyttes en kunde når kunden oprettes. <br />
	Derved vil denne gruppe altid være tilgængelig på job der er oprettet på denne kunde.
	<br />
	<br>
	<a href="stfolder_gruppe.asp?menu=job&func=opret">Opret ny gruppe&nbsp;&nbsp;<img src="../ill/pil_groen_lille.gif" width="17" height="17" alt="" border="0"></a><br>
	<br>
	<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#EFF3FF">
 	<tr bgcolor="5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=4 valign="top"><img src="../ill/tabel_top.gif" width="587" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class='alt'><b>Id</b></td>
		<td class='alt'><b>Gruppe</b></td>
		<td class='alt'><b>Foldere</b></td>
		<td class='alt'><b>Slet gruppe?</b></td>
	</tr>
	<%
	strSQL = "SELECT id, navn FROM folder_grupper ORDER BY navn"
	c = 0
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	%>
	<tr>
		<td bgcolor="#5582D2" colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<%
	if cint(id) = oRec("id") then
		bgc = "#ffff99"
	else
		select case right(c, 1)
		case 0, 2, 4, 6, 8
		bgc = "#ffffff"
		case else
		bgc = "#EFF3FF"
		end select
	end if
	%>
	<tr bgcolor=<%=bgc%>>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td><%=oRec("id")%></td>
					
					<%
					'** Henter aktiviteter i den aktuelle gruppe ****
					strSQL2 = "SELECT count(id) AS antal FROM foldere WHERE stfoldergruppe = "& oRec("id") 
					oRec2.open strSQL2, oConn, 3
					if not oRec2.EOF then
					intAntal = oRec2("antal")
					end if
					oRec2.close 
					%>
		
		<%if oRec("id") = 2 then%>
		<td height="20"><%=oRec("navn")%></td>
		<%else%>
		<td height="20"><a href="stfolder_gruppe.asp?menu=job&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%></a></td>
		<%end if%>
		<td>&nbsp;(<%=intAntal%>)&nbsp;<a href='stfolder_gruppe.asp?menu=job&func=visgruppe&gruppeid=<%=oRec("id")%>&gruppenavn=<%=oRec("navn")%>' class='vmenu'><img src="../ill/pil_groen_lille.gif" width="17" height="17" alt="" border="0"></a></td>
		
		<%if oRec("id") <= 2 then%>
		<td>&nbsp;</td>
		<%else%>
			<%if intAntal = 0 then%>
			<td><a href="stfolder_gruppe.asp?menu=job&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet_eks.gif" width="20" height="20" alt="" border="0"></a></td>
			<%else%>
			<td>&nbsp;</td>
			<%end if%>
		<%end if%>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	x = 0
	c = c + 1
	oRec.movenext
	wend
	%>	
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="4" valign="bottom"><img src="../ill/tabel_top.gif" width="587" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	&nbsp;
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
