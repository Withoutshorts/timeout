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
	if func = "dbopr" OR func = "opret" then
	id = 0
	else
	id = request("id")
	end if
	
	strjobnr = trim(Request("jobid"))
	jobid = strjobnr
	
	public strTdato
	public strUdato
	public intjobnrthis
	public totaltimertildelt
	
	function budgettimerTildelt(jobidthis)
	strSQL = "SELECT jobstartdato, jobslutdato, budgettimer, jobnr FROM job WHERE id = " & jobidthis
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	intjobnrthis = oRec("jobnr")
	strTdato = oRec("jobstartdato")
	strUdato = oRec("jobslutdato")
	jobtotaltimer = oRec("budgettimer")
	end if
	oRec.close
	
	if len(jobtotaltimer) <> 0 then
	jobtotaltimer = jobtotaltimer
	else
	jobtotaltimer = 0
	end if
	
	strSQL = "SELECT sum(budgettimer) AS akttimer FROM aktiviteter WHERE job = " & jobidthis &" ORDER BY budgettimer"
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
		akttotaltimer = oRec("akttimer")
	end if
	oRec.close
	
	if len(akttotaltimer) <> 0 then
	akttotaltimer = akttotaltimer
	else
	akttotaltimer = 0
	end if
	
	
	totaltimertildelt = (jobtotaltimer - akttotaltimer)
	end function
	
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:210; top:180; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF">
	<tr>
	    <td style="!border: 1px; border-color: #8CAAE6; border-style: solid; padding-right:5; padding-left:5; padding-top:5; padding-bottom:5;"><img src="../ill/alert.gif" width="44" height="45" alt="" border="0"><br>
		Du er ved at <b>slette</b> en aktivitet. <br>
		Du vil samtidig slette alle timeregistreringer på denne aktivitet.<br>
		Timeregistreringerne vil <b>ikke kunne genskabes</b>. <br>
		<br>Er du sikker på at du vil slette denne aktivitet?<br><br>
		<a href="aktiv.asp?menu=job&func=sletok&jobid=<%=jobid%>&jobnavn=<%=request("jobnavn")%>&id=<%=id%>">Ja</a><img src="../ill/blank.gif" width="200" height="1" alt="" border="0"><a href="javascript:history.back()">Nej, Annuler!</a></td>
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
	'*** Her slettes en aktivitet ***
	oConn.execute("DELETE FROM aktiviteter WHERE id = "& id &"")
	oConn.execute("DELETE FROM timer WHERE TAktivitetId = "& id &"")
	Response.redirect "aktiv.asp?menu=job&jobid="&jobid&"&jobnavn="&request("jobnavn")&""
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny aktivitet i db ****
	
	'*tjekker om startdag eksisterer ** 
				if Request("FM_start_dag") > 28 then 
				select case Request("FM_start_mrd")
				case "2"
				strStartDay = 28
				case "4", "6", "9", "11"
				strStartDay = 30
				case else
				strStartDay = Request("FM_start_dag")
				end select
				else
				strStartDay = Request("FM_start_dag")
				end if
				
				
				'*tjekker om slutdag eksisterer ** 
				if Request("FM_slut_dag") > 28 then 
				select case Request("FM_slut_mrd")
				case "2"
				strSlutDay = 28
				case "4", "6", "9", "11"
				strSlutDay = 30
				case else
				strSlutDay = Request("FM_slut_dag")
				end select
				else
				strSlutDay = Request("FM_slut_dag")
				end if
	
	startDatoNum = cdate(strStartDay &"/"& Request("FM_start_mrd") &"/"& Request("FM_start_aar")) 
	slutDatoNum = cdate(strSlutDay &"/"& Request("FM_slut_mrd")  &"/"& Request("FM_slut_aar"))
		
	strSQl = "SELECT jobstartdato, jobslutdato FROM job WHERE id =" & request("FM_jnr")	
	oRec.open strSQL, oConn, 3	
	if not oRec.EOF Then
	jobstdate = oRec("jobstartdato") 
	jobsldate = oRec("jobslutdato")
	end if
	oRec.close
	
	
	if dateDiff("d", startDatoNum, jobstdate) > 0 then
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
		<!--#include file="../inc/regular/topmenu_inc.asp"-->
		<%
		errortype = 32
		call showError(errortype)
		else
		if len(request("FM_budgettimer")) = 0 then
		useBudgettimer = 0
		else
		useBudgettimer = request("FM_budgettimer")
		end if
		
		if dateDiff("d", slutDatoNum, jobsldate) < 0 then
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
		<!--#include file="../inc/regular/topmenu_inc.asp"-->
		<%
		errortype = 33
		call showError(errortype)
		else
		if cdbl(request("FM_timertildelt")) < cdbl(useBudgettimer) then
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
		<!--#include file="../inc/regular/topmenu_inc.asp"-->
		<%
		errortype = 35
		call showError(errortype)
		else
		if len(request("FM_navn")) = 0 OR len(request("FM_jnr")) = 0 OR startDatoNum > slutDatoNum then
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
		<!--#include file="../inc/regular/topmenu_inc.asp"-->
		<%
		errortype = 23
		call showError(errortype)
		else
			%>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(useBudgettimer)
					if isInt > 0 then
					%>
					<!--#include file="../inc/regular/header_inc.asp"-->
					<!--#include file="../inc/regular/topmenu_inc.asp"-->
					<%
					errortype = 20
					call showError(errortype)
					isInt = 0
		
				else
				function SQLBless(s)
				dim tmp
				tmp = s
				tmp = replace(tmp, "'", "''")
				SQLBless = tmp
				end function
		
				strNavn = request("FM_navn")
				strbeskrivelse = SQLBless(request("FM_beskrivelse"))
				strjnr = request("FM_jnr")
				strOLDaktId = request("FM_OLDaktId")
				
				if request("FM_fakturerbart") = 1 then
				strFakturerbart = request("FM_fakturerbart")
				else
				strFakturerbart = 0
				end if
				
				'if request("FM_favorit") = 1 then
				int_aktFav = request("FM_favorit")
				'else
				'int_aktFav = 0
				'end if
				
				'if len(trim(request("FM_budgettimer"))) = "0" then
				'strBudgettimer = 0
				'else
				'strBudgettimer = request("FM_budgettimer")
				'end if
				
				intaktstatus = request("FM_aktstatus")
				strEditor = session("user")
				strDato = session("dato")
				
				startDato = Request("FM_start_aar") & "/" & Request("FM_start_mrd") & "/" & strStartDay 
				slutDato = Request("FM_slut_aar") & "/" & Request("FM_slut_mrd") & "/" & strSlutDay 
				
				if request("FM_alle") = "1" then
				strSQL = "SELECT projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM job WHERE id =" & strjnr
				oRec.open strSQL, oConn, 3
					if not oRec.EOF then
					strProjektgr1 = oRec("projektgruppe1")
					strProjektgr2 = oRec("projektgruppe2")
					strProjektgr3 = oRec("projektgruppe3")
					strProjektgr4 = oRec("projektgruppe4")
					strProjektgr5 = oRec("projektgruppe5")
					strProjektgr6 = oRec("projektgruppe6")
					strProjektgr7 = oRec("projektgruppe7")
					strProjektgr8 = oRec("projektgruppe8")
					strProjektgr9 = oRec("projektgruppe9")
					strProjektgr10 = oRec("projektgruppe10")
					end if
				oRec.close
				else
				strProjektgr1 = request("FM_projektgruppe_1")
				strProjektgr2 = request("FM_projektgruppe_2")
				strProjektgr3 = request("FM_projektgruppe_3")
				strProjektgr4 = request("FM_projektgruppe_4")
				strProjektgr5 = request("FM_projektgruppe_5")
				strProjektgr1 = request("FM_projektgruppe_6")
				strProjektgr2 = request("FM_projektgruppe_7")
				strProjektgr3 = request("FM_projektgruppe_8")
				strProjektgr4 = request("FM_projektgruppe_9")
				strProjektgr5 = request("FM_projektgruppe_10")
				end if
				
						if func = "dbopr" then
						oConn.execute("INSERT INTO aktiviteter (navn, beskrivelse, job, fakturerbar, aktfavorit, aktstartdato, aktslutdato, editor, dato, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, budgettimer, aktstatus) VALUES "_
						&" ('"& strNavn &"', "_
						&"'"& strBeskrivelse &"', "_ 
						&""& strjnr &", "_  
						&""& strFakturerbart &", "_ 
						&""& int_aktFav &", "_ 
						&"'"&startDato &"', "_ 
						&"'"&slutDato &"', "_
						&"'"&strEditor &"', "_
						&"'"&strDato &"', "_ 
						&""&strProjektgr1 &", "_ 
						&""&strProjektgr2 &", "_
						&""&strProjektgr3 &", "_
						&""&strProjektgr4 &", "_
						&""&strProjektgr5 &", "_      
						&""&useBudgettimer &", "_
						&""&intaktstatus &")") 
						else
						oConn.execute("UPDATE aktiviteter SET "_
						&" navn = '"& strNavn &"', "_
						&" beskrivelse ='"& strBeskrivelse &"', "_ 
						&" job = "& strjnr &", "_  
						&" fakturerbar = "& strFakturerbart &", "_ 
						&" aktfavorit = "& int_aktFav &", "_ 
						&" projektgruppe1 = "& strProjektgr1 &", "_ 
						&" projektgruppe2 = "& strProjektgr2 &", "_ 
						&" projektgruppe3 = "& strProjektgr3 &", " _
						&" projektgruppe4 = "& strProjektgr4 &", " _
						&" projektgruppe5 = "& strProjektgr5 &", " _
						&" projektgruppe6 = "& strProjektgr6 &", "_ 
						&" projektgruppe7 = "& strProjektgr7 &", "_ 
						&" projektgruppe8 = "& strProjektgr8 &", " _
						&" projektgruppe9 = "& strProjektgr9 &", " _
						&" projektgruppe10 = "& strProjektgr10 &", " _
						&" aktstartdato = '"& startDato &"', "_ 
						&" aktslutdato = '"& slutDato &"', "_
						&" editor = '"& strEditor &"', "_
						&" budgettimer = "& useBudgettimer &", "_
						&" dato = '"& strDato &"', "_
						&" aktstatus = "& intaktstatus &" WHERE id = "&id&"")
						
						if strFakturerbart = "1" then
						tfaktimvalue = 1
						else
						tfaktimvalue = 2
						end if
						
						oConn.execute("UPDATE timer SET "_
						& " TAktivitetNavn = '"& strNavn &"', "_
						& " TFaktim = "& tfaktimvalue &""_
						& " WHERE TAktivitetId = "& strOLDaktId &""_
						& "")
						
						end if
						
						Response.redirect "aktiv.asp?menu=job&jobid="&strjnr&"&jobnavn="&request("jobnavn")&""
					
					 end if '*validering
					end if '*validering
				   end if '*validering
				end if
			end if
	
	
	'***********************************************************************************************	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af Aktivitet ***
	if func = "opret" then
	strNavn = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny aktivitet"
	dbfunc = "dbopr"
	strFakturerbart = 1
	intaktstatus = 1
	intAktfavorit = 0
	
	call budgettimerTildelt(strjobnr)
	
	else
	strSQL = "SELECT id, navn, beskrivelse, job, fakturerbar, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, aktstartdato, aktslutdato, editor, dato, budgettimer, aktfavorit, aktstatus FROM aktiviteter WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strBeskrivelse = oRec("beskrivelse")
	strjobnr = oRec("job")
	strTdato = oRec("aktstartdato")
	strUdato = oRec("aktslutdato")
	strProj_1 = oRec("projektgruppe1")
	strProj_2 = oRec("projektgruppe2")
	strProj_3 = oRec("projektgruppe3")
	strProj_4 = oRec("projektgruppe4")
	strProj_5 = oRec("projektgruppe5")
	strProj_6 = oRec("projektgruppe6")
	strProj_7 = oRec("projektgruppe7")
	strProj_8 = oRec("projektgruppe8")
	strProj_9 = oRec("projektgruppe9")
	strProj_10 = oRec("projektgruppe10")
	strDato = oRec("dato")
	strLastUptDato = oRec("dato") 
	strEditor = oRec("editor")
	strFakturerbart = oRec("fakturerbar")
	intAktfavorit = oRec("aktfavorit")
	strBudgettimer = oRec("budgettimer")
	intaktstatus = oRec("aktstatus")
	end if
	
	oRec.close
	
	
	call budgettimerTildelt(strjobnr)
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger aktivitet"
	varSubVal = "opdaterpil" 
	end if
	
	
		if strFakturerbart = 1 then
		varFakEks = "checked"
		else
		varFakEks = ""
		end if
		
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<script>
	function expand() {
		if (document.all("progrp").style.display == "none"){
			document.all("progrp").style.display = "";
		}else{
			document.all("progrp").style.display = "none";
		}
	}
	
	</script>
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	
	
	<%
	'** Ændrer ID så det passer til Dato2 fil.
	if func = "opret" then
	id = 1
	else
	id = id
	end if 
	%>
	<!--#include file="inc/dato2.asp"-->
	<%
	'** Ændrer ID tilbage til opr ID.
	if func = "opret" then
	id = 0
	nedarvdato = "j"
	else
	id = id
	end if
	%>
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:50;">
	<br><img src="../ill/header_akt.gif" alt="" border="0" width="279" height="34">
	<hr align="left" width="600" size="1" color="#000000" noshade>
	<div id="jobnavn" style="position:absolute; left:0; top:76; visibility:visible; z-index:1200;">
	<a href="jobs.asp?menu=job&func=red&id=<%=jobid%>&int=1"><%=request("jobnavn")%></a>
	</div>
	<br><br>
	<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#EFF3FF">
	<form action="aktiv.asp?menu=job&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
	<input type="hidden" name="jobnavn" value="<%=request("jobnavn")%>">
	<input type="hidden" name="FM_OLDaktId" value="<%=id%>">
	<input type="hidden" name="FM_timertildelt" value="<%=totaltimertildelt%>">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=2 valign="top"><img src="../ill/tabel_top.gif" width="584" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<%if dbfunc = "dbred" then%>
	<tr bgcolor="#5582D2">
		<td colspan="2" valign="bottom" class="alt"><b><%=varbroedkrumme%></b><img src="../ill/blank.gif" width="200" height="1" alt="" border="0">Sidst opdateret den <b><%=strLastUptDato%></b> af <b><%=strEditor%></b></td>
	</tr>
	<%else%>
	<tr bgcolor="#5582D2">
		<td colspan="2" valign="bottom" class="alt"><b><%=varbroedkrumme%></b><img src="../ill/blank.gif" width="250" height="1" alt="" border="0"></td>
	</tr>
	<%end if%>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td colspan="2">&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td style="padding-left:40;">Aktivitet:</td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="90" alt="" border="0"></td>
		<td style="padding-left:40;" valign="TOP">Beskrivelse:</td>
		<td><textarea name="FM_beskrivelse" rows="5" cols="35" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><%=strBeskrivelse%></textarea></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="90" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td style="padding-left:40;">Timer:&nbsp;(enheder)</td>
		<td><input type="text" name="FM_budgettimer" value="<%=strBudgettimer%>" size="10" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;&nbsp;Budgetterede timer til rådighed på job:&nbsp;<b><%=totaltimertildelt%></b></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td style="padding-left:40;">Status:</td>
		<td><select name="FM_aktstatus">
		<%if dbfunc = "dbred" then 
		select case intaktstatus
		case 1
		strStatusNavn = "Aktiv"
		case 2
		strStatusNavn = "Passiv"
		case 0
		strStatusNavn = "Lukket"
		end select
		%>
		<option value="<%=intaktstatus%>" SELECTED><%=strStatusNavn%></option>
		<%end if%>
	<option value="1">Aktiv</option>
	<option value="2">Passiv</option>
	<option value="0">Lukket</option>
</select></td>
	<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="50" alt="" border="0"></td>
		<td colspan="2" style="padding-left:40;"><br><br><b>Stam-aktivitet?</b></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="50" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td colspan="2" style="padding-left:40;">Skal denne aktivitet fungere som en Stam-aktivitet?</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	strSQL = "SELECT id, navn FROM akt_gruppe WHERE id <> 2 ORDER BY navn"
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
		if st = 0 then
			if intAktfavorit = 0 then
			varAktFchecked = "checked"
			else
			varAktFchecked = ""
			end if
			%>
			<tr>
				<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="90" alt="" border="0"></td>
				<td valign=top style="padding-left:40;"><b>Nej!</b><br><br>
				Opret denne aktivitet i en af <br>nedenstående <br>stam-aktivitets grupper:<br>
				<b>Ja!</b></td>
				<td valign=top>&nbsp;<input type="radio" name="FM_favorit" value="0" <%=varAktFchecked%>></td>
				<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="90" alt="" border="0"></td>
			</tr>
		<%end if%>
	<%if intAktfavorit = oRec("id") then
	varAktFchecked = "checked"
	else
	varAktFchecked = ""
	end if%>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td class="blaa" style="padding-left:40;"><%=oRec("navn")%></td>
		<td>&nbsp;<input type="radio" name="FM_favorit" value="<%=oRec("id")%>" <%=varAktFchecked%>></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	st = 1
	oRec.movenext
	wend
	oRec.close%>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td colspan="2"><br>&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td colspan="2" style="padding-left:40;"><br><b>Er dette en fakturerbar aktivitet?</b></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td style="padding-left:40;">Fakturerbar?</td>
		<td>Ja:&nbsp;<input type="checkbox" name="FM_fakturerbart" value="1" <%=varFakEks%>></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td><br>&nbsp;</td>
		<td><br>&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td colspan="2" style="padding-left:40;"><br><b>Start og slut dato:</b>&nbsp;&nbsp;
			<%if func = "opret" then
			Response.write "(nedarves fra job)"
			end if%>
		</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td style="padding-left:40;">Start dato:&nbsp;</td><td><select name="FM_start_dag">
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
		<option value="31">31</option></select>&nbsp;
		
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
		<%if id <> 0 OR nedarvdato = "j" then%>
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
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		</tr>
		<tr>
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
			<td style="padding-left:40;">Slut dato:&nbsp;</td><td><select name="FM_slut_dag">
		<option value="<%=strDag_slut%>"><%=strDag_slut%></option> 
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
		<option value="31">31</option></select>&nbsp;
		
		<select name="FM_slut_mrd">
		<option value="<%=strMrd_slut%>"><%=strMrdNavn_slut%></option>
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
		
		
		<select name="FM_slut_aar">
		<option value="<%=strAar_slut%>">
		<%if id <> 0 OR nedarvdato = "j" then%>
		20<%=strAar_slut%>
		<%else%>
		<%=strAar_slut%>
		<%end if%></option>
		<option value="02">2002</option>
		<option value="03">2003</option>
	   	<option value="04">2004</option>
	   	<option value="05">2005</option>
		<option value="06">2006</option>
		<option value="07">2007</option></select></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td><br>&nbsp;</td>
		<td><br>&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td valign="top" style="padding-left:40;"><b>Denne aktivitet tilhører:</b></td>
		<td><%
		strSQL = "SELECT jobnavn, jobnr, id FROM job WHERE id="& strjobnr
		oRec.open strSQL, oConn, 3
		
		if not oRec.EOF then
		Response.write oRec("jobnr") &"&nbsp;&nbsp;" & oRec("jobnavn")
		end if
		%>
		<input type="hidden" name="FM_jnr" value="<%=oRec("id")%>">
		<%oRec.close%><br>&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%if dbfunc <> "dbred" then%>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="100" alt="" border="0"></td>
		<td colspan="2" style="padding-left:40;"><br><b>Projektgrupper:</b><br>Vælg <b>alle</b> de projketgrupper der har adgang til ovenstående job? 
		&nbsp;&nbsp;&nbsp;&nbsp;<b>ja:</b> <input type="checkbox" name="FM_alle" value="1" CHECKED onclick="expand();"><br>
		eller angiv de projektgrupper der skal have adgang til denne aktivitet.<br>
		&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="100" alt="" border="0"></td>
	</tr>
	<%end if%>
	
	<%if func = "red" then
	disp = "show"
	%>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td colspan="2" style="padding-left:40;"><br><b>Projektgrupper:</b><br></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	else
	disp = "none"
	end if%>
	
	<tr>
	<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="180" alt="" border="0"></td>
	<td colspan="2">
			<DIV ID="progrp" name="progrp" Style="position: relative; display: <%=disp%>;">
			<table>
			<tr>
				<td valign="top" style="padding-left:40;">Projektgruppe 1:</td>
				<td><select name="FM_projektgruppe_1">
				<option value="1">Ingen</option>
					<%
					strSQL = "SELECT projektgrupper.id AS id, projektgrupper.navn AS navn, job.projektgruppe1, job.projektgruppe2, job.projektgruppe3, job.projektgruppe4, job.projektgruppe5 FROM projektgrupper, job WHERE projektgrupper.id = job.projektgruppe1 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe2 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe3 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe4 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe5 AND job.id = "& strjobnr &" ORDER BY projektgrupper.id"
					oRec.open strSQL, oConn, 3
					
					While not oRec.EOF 
					projgId_1 = oRec("id")
					projgNavn_1 = oRec("navn")
					
					if trim(strProj_1) = trim(projgId_1) then
					varSelected_1 = "SELECTED"
					else
					varSelected_1 = ""
					end if
					
					if varLastName <> projgId_1 then
					%>
					<option value="<%=projgId_1%>" <%=varSelected_1%>><%=projgNavn_1%></option>
					<%
					end if
					varLastName = projgId_1
					oRec.movenext
					wend
					varLastName = ""
					oRec.close%>
		</select></td>
			</tr>
			
			
			<tr>
				<td valign="top" style="padding-left:40;">Projektgruppe 2:</td>
				<td><select name="FM_projektgruppe_2">
				<option value="1">Ingen</option>
				<%
					strSQL = "SELECT projektgrupper.id AS id, projektgrupper.navn AS navn, job.projektgruppe1, job.projektgruppe2, job.projektgruppe3, job.projektgruppe4, job.projektgruppe5 FROM projektgrupper, job WHERE projektgrupper.id = job.projektgruppe1 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe2 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe3 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe4 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe5 AND job.id = "& strjobnr &" ORDER BY projektgrupper.id"
					oRec.open strSQL, oConn, 3
					
					While not oRec.EOF 
					projgId_2 = oRec("id")
					projgNavn_2 = oRec("navn")
					
					if trim(strProj_2) = trim(projgId_2) then
					varSelected_2 = "SELECTED"
					else
					varSelected_2 = ""
					end if
					
					if varLastName <> projgId_2 then
					%>
					<option value="<%=projgId_2%>" <%=varSelected_2%>><%=projgNavn_2%></option>
					<%
					end if
					
					varLastName = projgId_2
					oRec.movenext
					wend
					varLastName = ""
					oRec.close%>
		</select></td>
			</tr>
			<tr>
				<td valign="top" style="padding-left:40;">Projektgruppe 3:</td>
				<td><select name="FM_projektgruppe_3">
				<option value="1">Ingen</option>
				<%
					strSQL = "SELECT projektgrupper.id AS id, projektgrupper.navn AS navn, job.projektgruppe1, job.projektgruppe2, job.projektgruppe3, job.projektgruppe4, job.projektgruppe5 FROM projektgrupper, job WHERE projektgrupper.id = job.projektgruppe1 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe2 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe3 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe4 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe5 AND job.id = "& strjobnr &" ORDER BY projektgrupper.id"
					oRec.open strSQL, oConn, 3
					
					While not oRec.EOF 
					projgId_3 = oRec("id")
					projgNavn_3 = oRec("navn")
					
					if trim(strProj_3) = trim(projgId_3) then
					varSelected_3 = "SELECTED"
					else
					varSelected_3 = ""
					end if
					
					if varLastName <> projgId_3 then
					%>
					<option value="<%=projgId_3%>" <%=varSelected_3%>><%=projgNavn_3%></option>
					<%
					end if
					
					varLastName = projgId_3
					oRec.movenext
					wend
					varLastName = ""
					oRec.close%>
		</select></td>
			</tr>
				<tr>
				<td valign="top" style="padding-left:40;">Projektgruppe 4:</td>
				<td><select name="FM_projektgruppe_4">
				<option value="1">Ingen</option>
				<%
					strSQL = "SELECT projektgrupper.id AS id, projektgrupper.navn AS navn, job.projektgruppe1, job.projektgruppe2, job.projektgruppe3, job.projektgruppe4, job.projektgruppe5 FROM projektgrupper, job WHERE projektgrupper.id = job.projektgruppe1 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe2 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe3 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe4 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe5 AND job.id = "& strjobnr &" ORDER BY projektgrupper.id"
					oRec.open strSQL, oConn, 3
					
					While not oRec.EOF 
					projgId_4 = oRec("id")
					projgNavn_4 = oRec("navn")
					
					if trim(strProj_4) = trim(projgId_4) then
					varSelected_4 = "SELECTED"
					else
					varSelected_4 = ""
					end if
					
					if varLastName <> projgId_4 then
					%>
					<option value="<%=projgId_4%>" <%=varSelected_4%>><%=projgNavn_4%></option>
					<%
					end if
					
					varLastName = projgId_4
					oRec.movenext
					wend
					varLastName = ""
					oRec.close%>
		</select></td>
			</tr>
				<tr>
				<td valign="top" style="padding-left:40;">Projektgruppe 5:</td>
				<td><select name="FM_projektgruppe_5">
				<option value="1">Ingen</option>
				<%
					strSQL = "SELECT projektgrupper.id AS id, projektgrupper.navn AS navn, job.projektgruppe1, job.projektgruppe2, job.projektgruppe3, job.projektgruppe4, job.projektgruppe5 FROM projektgrupper, job WHERE projektgrupper.id = job.projektgruppe1 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe2 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe3 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe4 AND job.id = "& strjobnr &" OR projektgrupper.id = job.projektgruppe5 AND job.id = "& strjobnr &" ORDER BY projektgrupper.id"
					oRec.open strSQL, oConn, 3
					
					While not oRec.EOF 
					projgId_5 = oRec("id")
					projgNavn_5 = oRec("navn")
					
					if trim(strProj_5) = trim(projgId_5) then
					varSelected_5 = "SELECTED"
					else
					varSelected_5 = ""
					end if
					
					if varLastName <> projgId_5 then
					%>
					<option value="<%=projgId_5%>" <%=varSelected_5%>><%=projgNavn_5%></option>
					<%
					end if
					
					varLastName = projgId_5
					oRec.movenext
					wend
					varLastName = ""
					oRec.close%>
			</select></td>
			</tr>
			</table>
			</div>
		</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="180" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan=2 valign="bottom"><img src="../ill/tabel_top.gif" width="584" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	</table>
	<table>
	<tr>
		<td><br><br><img src="ill/blank.gif" width="200" height="1" alt="" border="0"><input type="image" src="../ill/<%=varSubVal%>.gif"></td>
	</tr>
	</form>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%case "favorit"%>
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
	<!--#include file="inc/convertDate.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:50; visibility:visible;">
	<br><img src="../ill/job_stam_header.gif" alt="" border="0" width="600" height="49"><br><br>
	Stam aktiviteter oprettes, ved at "tilføje til favorit aktivitet" under den enkelte aktivitet.<br>
	For nemmere at oprette et job og tilknytte aktiviteter, kan der, når jobbet oprettes vælges at "tilføje stam aktiviteter".
	<br><br>
	På listen over stam aktiviteter er netop nu:<br>
	<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#EFF3FF">
 	<tr bgcolor="5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=2 valign="top"><img src="../ill/tabel_top.gif" width="587" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class='alt'><b>Navn</b></td>
		<td class='alt'><b>Fjern</b></td>
	</tr>
	<%
	if id <> 0 then
	sqlKriakt = "aktfavorit = "& id &""
	else
	sqlKriakt = "aktfavorit <> 0"
	end if
	
	strSQL = "select a.navn, a.id, aktfavorit, b.navn AS gnavn, b.id AS gid FROM aktiviteter a, akt_gruppe b WHERE "& sqlKriakt &" AND b.id = a.aktfavorit ORDER BY aktfavorit"
	a = 0
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	if lastfavgp <> oRec("aktfavorit") then%>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td height="16" valign="middle" colspan="2">&nbsp;<b>Stam aktivitetsgruppe:&nbsp;<%=oRec("gnavn")%></b></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%end if%>
	<tr>
		<td bgcolor="#5582D2" colspan="4"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'');">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td height="16" valign="middle">&nbsp;<%=oRec("navn")%></td>
		<%if oRec("gid") <> 2 then%>
		<td><a href="aktiv.asp?func=favorit_fjern&id=<%=oRec("id")%>"><img src='../ill/slet_eks.gif' width='20' height='20' alt='Fjern' border='0'></a></td>
		<%else%>
		<td>&nbsp;</td>
		<%end if%>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	lastfavgp = oRec("aktfavorit")
	a = a + 1
	oRec.movenext
	wend
	oRec.close
	
	if a = 0 then%>
	<tr>
		<td bgcolor="#5582D2" colspan="4"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr onmouseover="mOvr('gift',this,'#EFF3FF');" onmouseout="mOut(this,'');">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="76" alt="" border="0"></td>
		<td height="76" valign="top" colspan="2">&nbsp;<br><img src="../ill/alert.gif" width="44" height="45" alt="" border="0">Der er ingen stam-aktiviteter i denne gruppe.</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="76" alt="" border="0"></td>
	</tr>
	<%end if%>
	<tr>
		<td bgcolor="#5582D2" colspan="4"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="2" valign="bottom"><img src="../ill/tabel_top.gif" width="587" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	</table>
	<br>NB) Disse aktiviter bliver, hvis du tilvælger det, automatisk tilføjet når du opretter et nyt job.
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%case "favorit_fjern"
	
	strSQL = "UPDATE aktiviteter SET aktfavorit = 0 WHERE id =" & id
	oRec.open strSQL, oConn, 3
	
	Response.redirect "aktiv.asp?menu=job&func=favorit"
	
	case else%>
	
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
	<!--#include file="inc/convertDate.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:50; visibility:visible;">
	<%
	jobid = request("jobid")
	%>
	<br><img src="../ill/header_akt.gif" alt="" border="0" width="279" height="34">
	<div id="jobnavn" style="position:absolute; left:0; top:76; visibility:visible; z-index:1200;">
	<a href="jobs.asp?menu=job&func=red&id=<%=jobid%>&int=1"><%=request("jobnavn")%></a>
	</div>
	<hr align="left" width="730" size="1" color="#000000" noshade>
	<br><br>
	
	<%
	call budgettimerTildelt(strjobnr)

	if totaltimertildelt > 0 OR len(trim(totaltimertildelt)) = 0 then %>
	<div id="opretny" style="position:absolute; left:590; top:30; visibility:visible;">
	<a href="aktiv.asp?menu=job&func=opret&jobid=<%=jobid%>&id=0&jobnavn=<%=request("jobnavn")%>">Opret ny aktivitet&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
	</div>
	<%else%>
	<div id="opretny" style="position:absolute; left:440; top:5; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF">
	<tr>
	    <td style="!border: 1px; border-color: #8CAAE6; border-style: solid; padding-right:5; padding-left:5; padding-top:2; padding-bottom:2;">
		Der kan ikke oprettes nye aktiviteter, da der ikke er<br>
		flere budgetterede timer til rådighed på dette job.<br>
		<a href="jobs.asp?menu=job&func=red&id=<%=jobid%>&int=1" class="vmenuglobal">Tildel flere timer...</a>
		</td>
	</tr></table>
	</div>
	<%end if%>
	
	<table cellspacing="0" cellpadding="0" border="0" width="730" bgcolor="#EFF3FF">
	<tr bgcolor="5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="31" alt="" border="0"></td>
		<td colspan="8" valign="top"><img src="../ill/tabel_top.gif" width="714" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="31" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
	<td width="220" class='alt'>Aktivitet</td>
	<td colspan="2" width="120" class='alt'>Fuldført</td>
	<td class='alt'>Timer tildelt/ Forbrugt</td>
	<td class='alt'>&nbsp;&nbsp;<a href="aktiv.asp?menu=job&sort=slutdato&jobid=<%=jobid%>&jobnavn=<%=request("jobnavn")%>" class='alt'><b>Slut dato</b></a></td>
	<td align="center" class='alt'>Status</td>
	<td align="center" class='alt'>&nbsp;&nbsp;Fakturerbar?</td>
	<td>&nbsp;&nbsp;&nbsp;</td>
	</tr>
	<%
	sort = Request("sort")
	select case sort
	case "slutdato"
	strSQL = "SELECT id, navn, job, fakturerbar, projektgruppe1, projektgruppe2, projektgruppe3, aktstartdato, aktslutdato, budgettimer, aktstatus FROM aktiviteter WHERE job = "& jobid &" ORDER BY aktslutdato"
	case else
	strSQL = "SELECT id, navn, job, fakturerbar, projektgruppe1, projektgruppe2, projektgruppe3, aktstartdato, aktslutdato, budgettimer, aktstatus FROM aktiviteter WHERE job = "& jobid &" ORDER BY navn"
	end select
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
	strSQL = "SELECT sum(timer) AS proafslut FROM timer WHERE TaktivitetId=" & oRec("id")
	oRec3.open strSQL, oConn, 3
	
	if len(oRec3("proafslut")) <> 0 then
	proaf = oRec3("proafslut")
	else
	proaf = "0"
	end if
	
	oRec3.close
	
	select case oRec("fakturerbar")
	case "1" 
	fakikon = "<img src='../ill/fakbarikon.gif' width='26' height='25' alt='' border='0'>"
	case "2"
	fakikon = "<img src='../ill/blank.gif' width='26' height='25' alt='' border='0'>"
	case else
	fakikon = "<img src='../ill/notfakbarikon.gif' width='26' height='25' alt='' border='0'>"
	end select
	
	if len(trim(oRec("budgettimer"))) = "0" OR oRec("budgettimer") = "0" then 
	budgettimer = 0
	else
	budgettimer = oRec("budgettimer")
	end if
	
	if budgettimer > 0 then
	projektcomplt = ((proaf/budgettimer)*100)
	else
	projektcomplt = 100
	end if
	
	if projektcomplt > 100 then
	projektcomplt = 100
	end if
	%>
	<tr>
		<td bgcolor="#5582D2" colspan="11"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'');">
	<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	<%if oRec("fakturerbar") <> 2 then 'Kørsel%>
	<td height="20">&nbsp;&nbsp;<a href="aktiv.asp?menu=job&func=red&id=<%=oRec("id")%>&jobnavn=<%=request("jobnavn")%>"><%=left(oRec("navn"), 24)%> </a></td>
	<td align="left"><div style="!border: 1 px;background-color:#ffffff; border-color: black; border-style: solid; height:6px; width:100px;"><img src="../ill/completed.gif" width="<%=left(projektcomplt, 4)%>" height="6" alt="" border="0"></div></td>
	<td class='lillegray'>&nbsp;<%=left(projektcomplt, 4)%>%&nbsp;</td>
	<td>&nbsp;<%=budgettimer%>&nbsp;/&nbsp;<b><%=proaf%></b>&nbsp;</td>
	<td>&nbsp;<%=convertDate(oRec("aktslutdato"))%></td>
	<%else%>
	<td height="20" colspan="5">&nbsp;&nbsp;<%=left(oRec("navn"), 24)%></td>
	<%end if%>
	<td align="center">
	<%select case oRec("aktstatus")
			case "1"
			if (oRec("aktslutdato") - date) < 0 then
			strStatImg = "<img src='../ill/status_slut.gif' width='16' height='19' alt='' border='0'>"
			else
			strStatImg = "<img src='../ill/status_groen.gif' width='14' height='19' alt='' border='0'>"
			end if
			case "2"
			strStatImg = "<img src='../ill/status_graa.gif' width='14' height='20' alt='' border='0'>"
			case "0"
			strStatImg = "<img src='../ill/status_roed.gif' width='14' height='19' alt='' border='0'>"
			end select
	Response.write strStatImg%></td>
	<td align="center"><%=fakikon%></td>
	<td align="center"><a href="aktiv.asp?menu=job&func=slet&jobid=<%=jobid%>&jobnavn=<%=request("jobnavn")%>&id=<%=oRec("id")%>"><img src="../ill/slet_eks.gif" width="20" height="20" alt="" border="0"></a></td>
	<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="8" valign="bottom"><img src="../ill/tabel_top.gif" width="714" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>		
	</table>
	<%
	'*** fakturerbare timer tildelt på aktiviteter **** 
	strSQL3 = "SELECT sum(budgettimer) AS akttimer FROM aktiviteter WHERE job = " & jobid &" AND fakturerbar = 1 ORDER BY budgettimer"
	oRec3.open strSQL3, oConn, 3
	if not oRec3.EOF then
		akttfaktimtildelt = oRec3("akttimer")
	end if
	oRec3.close
	
	if len(akttfaktimtildelt) <> 0 then
	akttfaktimtildelt = akttfaktimtildelt
	else
	akttfaktimtildelt = 0
	end if
	
	
	'***Ikke fakturerbare timer tildelt på aktiviteter **** 
	strSQL3 = "SELECT sum(budgettimer) AS aktnotftimer FROM aktiviteter WHERE job = " & jobid &" AND fakturerbar = 0 ORDER BY budgettimer"
	oRec3.open strSQL3, oConn, 3
	if not oRec3.EOF then
		akttnotfaktimtildelt = oRec3("aktnotftimer")
	end if
	oRec3.close
	
	if len(akttnotfaktimtildelt) <> 0 then
	akttnotfaktimtildelt = akttnotfaktimtildelt
	else
	akttnotfaktimtildelt = 0
	end if
	%>
	<br><br>
	Antal fakturerbare timer tildelt: <b><%=akttfaktimtildelt%></b><br>
	Antal <b>ikke</b> fakturerbare timer tildelt: <b><%=akttnotfaktimtildelt%></b>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="jobs.asp?menu=job">Joboversigt</a>
	<br>
	<br>
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
