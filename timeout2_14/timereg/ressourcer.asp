<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="inc/helligdage_func.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function
	
	func = request("func")
	id = request("id")
	
	select case func 
	case "editdate" 
			
			medarbId = request("mid")
			stdagiuge = request("stdagiuge")
			sldagiuge = request("sldagiuge")
			intc = request("intc")
			
			strSQL = "SELECT id, jobnavn, fakturerbart, jobstartdato, jobslutdato FROM job WHERE id="&id
			
			oRec.open strSQL, oConn, 3 
			if not oRec.EOF then	
			jobnavn = oRec("jobnavn")
			jid = id
			fakbart = oRec("fakturerbart")
			useLowerdatoKri = cdate(oRec("jobstartdato"))
			useUpperdatoKri = cdate(oRec("jobslutdato"))
			end if
			oRec.close
			
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
			<td bgcolor="#FFFFFF" align=right><a href="javascript:window.close()"><img src="../ill/luk_xp.gif" width="30" height="28" alt="" border="0">&nbsp;Luk</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"><a href="javascript:window.print()"><img src="../ill/print_xp.gif" width="28" height="30" alt="" border="0">&nbsp;Print</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"></td>
		</tr>
		</table>
				<div id="editdate" name="editdate" style="position:absolute; left:30; top:80;">
				<table cellspacing="1" cellpadding="0" border="0" width="680" bgcolor="#5582d2">
				<form action="ressourcer.asp?menu=job&func=upd" target="_self" method="post" name="updated_<%=cnt%>" id="updated_<%=cnt%>">
				<input type="hidden" name="mid" id="mid" value="<%=medarbId%>">
				<input type="hidden" name="id" id="id" value="<%=id%>">
				<input type="hidden" name="intc" id="intc" value="<%=intc%>">
				<tr>
					<td bgcolor="#ffffff" colspan="9" style="Padding-left:5; Padding-right:5; Padding-top:4;"><b>Tildel timer på de valgte dage:</b></td>
				</tr>
				<tr bgcolor="#ffffff">
					<td widht=80 valign="top" align="right" style="Padding-right:5; Padding-top:3;" class=lille>&nbsp;</td>
					<td widht=80 valign="top" align="right" style="Padding-right:5; Padding-top:3;" class=lille>S <%=stdagiuge%><br><%call helligdage(stdagiuge)%></td>
					<td widht=80 valign="top" align="right" style="Padding-right:5; Padding-top:3;" class=lille>M <%=dateadd("d", 1, stdagiuge)%><br><%call helligdage(dateadd("d", 1, stdagiuge))%></td>
					<td widht=80 valign="top" align="right" style="Padding-right:5; Padding-top:3;" class=lille>T <%=dateadd("d", 2, stdagiuge)%><br><%call helligdage(dateadd("d", 2, stdagiuge))%></td>
					<td widht=80 valign="top" align="right" style="Padding-right:5; Padding-top:3;" class=lille>O <%=dateadd("d", 3, stdagiuge)%><br><%call helligdage(dateadd("d", 3, stdagiuge))%></td>
					<td widht=80 valign="top" align="right" style="Padding-right:5; Padding-top:3;" class=lille>T <%=dateadd("d", 4, stdagiuge)%><br><%call helligdage(dateadd("d", 4, stdagiuge))%></td>
					<td widht=80 valign="top" align="right" style="Padding-right:5; Padding-top:3;" class=lille>F <%=dateadd("d", 5, stdagiuge)%><br><%call helligdage(dateadd("d", 5, stdagiuge))%></td>
					<td widht=80 valign="top" align="right" style="Padding-right:5; Padding-top:3;" class=lille>L <%=sldagiuge%><br><%call helligdage(sldagiuge)%></td>
					<td widht=80 valign="top" style="Padding-right:5; Padding-top:3;" align=center class=lille>&nbsp;Timer tildelt.</td>
					<input type="hidden" name="FM_son_date" value="<%=stdagiuge%>">
					<input type="hidden" name="FM_man_date" value="<%=dateadd("d", 1, stdagiuge)%>">
					<input type="hidden" name="FM_tir_date" value="<%=dateadd("d", 2, stdagiuge)%>">
					<input type="hidden" name="FM_ons_date" value="<%=dateadd("d", 3, stdagiuge)%>">
					<input type="hidden" name="FM_tor_date" value="<%=dateadd("d", 4, stdagiuge)%>">
					<input type="hidden" name="FM_fre_date" value="<%=dateadd("d", 5, stdagiuge)%>">
					<input type="hidden" name="FM_lor_date" value="<%=sldagiuge%>">
				</tr>
					<tr bgcolor="#ffffff">
						<td class="blaa" valign="middle" style="padding-left:2;"><img src="../ill/blank.gif" width="120" height="1" alt="" border="0"><br>
						<%
						if fakbart = "1" then
						inteksimg = "<img src='../ill/eksternt_job_trans.gif' width='15' height='15' alt='"&formatdatetime(useLowerdatoKri, 1)&" til "& formatdatetime(useUpperdatoKri, 1)&"' border='0'>"
						else
						inteksimg = "<img src='../ill/internt_job_trans.gif' width='15' height='15' alt='"&formatdatetime(useLowerdatoKri, 1)&" til "& formatdatetime(useUpperdatoKri, 1)&"' border='0'>"
						end if
						Response.write inteksimg%>
						&nbsp;
						
						<b><%=left(jobnavn, 20)%></b></td>
					
					<td style="Padding-top:3;" align=center>&nbsp;</td>
					<td style="Padding-top:3;" align=center>&nbsp;</td>
					<td style="Padding-top:3;" align=center>&nbsp;</td>
					<td style="Padding-top:3;" align=center>&nbsp;</td>
					<td style="Padding-top:3;" align=center>&nbsp;</td>
					<td style="Padding-top:3;" align=center>&nbsp;</td>
					<td style="Padding-top:3;" align=center>&nbsp;</td>
					<td style="Padding-top:3;" align=center class=lille>&nbsp;Valgt med. / Alle</td>
					</tr>
				<%
				
				'************** udskriver aktiviteter *********************
				strSQL = "SELECT aktiviteter.job, aktiviteter.navn AS anavn, aktstartdato, aktslutdato, aktstatus, fakturerbar, aktiviteter.id AS aid FROM aktiviteter WHERE aktiviteter.job = "& id &" AND fakturerbar <> 2 ORDER BY navn"
				'Response.write strSQL
	
				oRec.open strSQL, oConn, 3
				x = 0
				while not oRec.EOF
				%>
				<input type="hidden" name="FM_aktid_<%=x%>" value="<%=oRec("aid")%>">
				<tr bgcolor="#ffffff">
					<td class="blaa" valign=middle style="padding-left:2;"><img src="../ill/blank.gif" width="120" height="1" alt="" border="0"><br>
					<%
					if oRec("fakturerbar") = 1 then
					aktfakimg = "<img src='../ill/fakbarikon_14px.gif' width='14' height='14' alt='"&formatdatetime(oRec("aktstartdato"), 1)&" til "& formatdatetime(oRec("aktslutdato"), 1)&"' border='0'>"
					else
					aktfakimg = "<img src='../ill/notfakbarikon_14px.gif' width='14' height='14' alt='"&formatdatetime(oRec("aktstartdato"), 1)&" til "& formatdatetime(oRec("aktslutdato"), 1)&"' border='0'>"
					end if
					Response.write aktfakimg
					%>
					&nbsp;<font class="lillesort"><%=left(oRec("anavn"), 20)%></font>
					</td>
					
							<%
							for ugedage = 0 to 6
							select case ugedage
							case 0
							usedagnavn = "son"
							case 1
							usedagnavn = "man" 
							case 2
							usedagnavn = "tir"
							case 3
							usedagnavn = "ons"
							case 4
							usedagnavn = "tor"
							case 5
							usedagnavn = "fre"
							case 6
							usedagnavn = "lor"
							end select 
							
							usedagher = dateadd("d", ugedage, stdagiuge)
							
							if usedagher < useLowerdatoKri OR usedagher > useUpperdatoKri then
							thisMlenght = 0
							thisSize = 4
							else
							thisMlenght = 5
							thisSize = 4
							end if
							
							if usedagnavn = "son" OR usedagnavn = "lor" then
							bgher = "#CCCCCC"
							else
							bgher = "#FFFFFF"
							end if
							
							'*** henter timer på akt ****************'
							strSQL3 = "SELECT timer AS sumtimerakt FROM ressourcer WHERE dato = '"&convertDateYMD(usedagher)&"' AND mid=" & medarbId &" AND jobid ="&id&" AND aktid ="& oRec("aid")
							oRec3.open strSQL3, oConn, 3 
							if not oRec3.EOF then
							timerakt = SQLBless(oRec3("sumtimerakt")) 
							end if
							oRec3.close
							%>
							<td bgcolor="<%=bgher%>" style="Padding-top:3;" align=center><input type="text" name="FM_<%=usedagnavn%>_<%=x%>" value="<%=timerakt%>" size="<%=thisSize%>" maxlength="<%=thisMlenght%>"></td>
							<%
							timerakt = ""
							next%>
							<td bgcolor="#ffffff" align=center class=lille>
							<%
							strSQL3 = "SELECT sum(timer) AS sumtottimeraktmed FROM ressourcer WHERE mid=" & medarbId &" AND jobid ="&id&" AND aktid ="& oRec("aid")
							oRec3.open strSQL3, oConn, 3 
							if not oRec3.EOF then
							timertotaktmed = oRec3("sumtottimeraktmed") 
							end if
							oRec3.close
					
							%> <%=timertotaktmed%> / 
							
							<%
							strSQL3 = "SELECT sum(timer) AS sumtottimerakt FROM ressourcer WHERE jobid ="&id&" AND aktid ="& oRec("aid")
							oRec3.open strSQL3, oConn, 3 
							if not oRec3.EOF then
							timertotakt = oRec3("sumtottimerakt") 
							end if
							oRec3.close
							%> <%=timertotakt%>
							</td>
				</tr>
				<%
				x = x + 1
				oRec.movenext
				wend
				
				oRec.close
				%>	
				<tr>
					<td bgcolor="#ffffff" colspan="9" style="Padding-left:5; Padding-right:5;"><br>
					<input type="hidden" name="antalx" id="antalx" value="<%=x%>">
					<font class=lillesort>Flyt alle tildelte ressource timer i denne uge <input type="text" name="FM_antal_flytdage" id="FM_antal_flytdage" size="2" maxlength="2" value="0" style="font-size:10;"> dage frem ell. (-)tilbage.<br>
					<input type="checkbox" name="FM_notuseweekend" value="1"> Brug først kommende mandag til timer der bliver lagt på en lørdag eller en søndag?  <br>
					<img src="../ill/blank.gif" width="530" height="1" alt="" border="0"><input type="image" src="../ill/opdaterpil.gif">
					</td>
				</tr>
				</form>
				</table><br><br>
				<br>&nbsp;&nbsp;<a href="Javascript:history.back()" class="vmenu"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0">&nbsp;Tilbage</a><br><br>	
			</div>
	
	<%
	case "upd" 
	%>
	<!--#include file="inc/isint_func.asp"-->
	<%
					
					intc = request("intc")
					if len(request("FM_antal_flytdage")) <> 0 then
						call erDetInt(request("FM_antal_flytdage"))
						if isInt > 0 then
						flytreg = 0
						else
						flytreg = request("FM_antal_flytdage")
						end if
					else
					flytreg = 0
					end if
					
					datoson = convertdateYMD(request("FM_son_date"))
					datoman = convertdateYMD(request("FM_man_date"))
					datotir = convertdateYMD(request("FM_tir_date"))
					datoons = convertdateYMD(request("FM_ons_date"))
					datotor = convertdateYMD(request("FM_tor_date"))
					datofre = convertdateYMD(request("FM_fre_date"))
					datolor = convertdateYMD(request("FM_lor_date"))
					
					medarbid = request("mid")
					antaljogakt = request("antalx")
					
					'**** Opdaterer ressourcer ***
					strSQL = "DELETE FROM ressourcer WHERE mid = "& medarbid & " AND jobid=" & id &" AND (dato = '"& datoson &"' OR dato = '"& datoman &"' OR dato = '"& datotir &"' OR dato = '"& datoons &"' OR dato = '"& datotor &"' OR dato = '"& datofre &"' OR dato = '"& datolor &"')" 
					oConn.execute(strSQL)
					
					'Korrigerer dato ved flyt
					
					'** ikke brug lørdag og søndag ***
					ikkeBruglorogson = request("FM_notuseweekend")
					if len(ikkeBruglorogson) <> 0 then
					ikkeBruglorogson = 1
					else
					ikkeBruglorogson = 0
					end if
					
					datoson = dateadd("d", flytreg, convertdateYMD(request("FM_son_date")))
					if ikkeBruglorogson = 1 AND flytreg <> 0 then
						if weekday(datoson) = 1 then
						datoson = dateadd("d", flytreg+1, convertdateYMD(request("FM_son_date")))
						end if
						if weekday(datoson) = 7 then
						datoson = dateadd("d", flytreg+2, convertdateYMD(request("FM_son_date")))
						end if
					end if
					
					datoman = dateadd("d", flytreg, convertdateYMD(request("FM_man_date")))
					if ikkeBruglorogson = 1 AND flytreg <> 0 then
						if weekday(datoman) = 1 then
						datoman = dateadd("d", flytreg+1, convertdateYMD(request("FM_man_date")))
						end if
						if weekday(datoman) = 7 then
						datoman = dateadd("d", flytreg+2, convertdateYMD(request("FM_man_date")))
						end if
					end if
					
					datotir = dateadd("d", flytreg, convertdateYMD(request("FM_tir_date")))
					if ikkeBruglorogson = 1 AND flytreg <> 0 then
						if weekday(datotir) = 1 then
						datotir = dateadd("d", flytreg+1, convertdateYMD(request("FM_tir_date")))
						end if
						if weekday(datotir) = 7 then
						datotir = dateadd("d", flytreg+2, convertdateYMD(request("FM_tir_date")))
						end if
					end if
					
					datoons = dateadd("d", flytreg, convertdateYMD(request("FM_ons_date")))
					if ikkeBruglorogson = 1 AND flytreg <> 0 then
						if weekday(datoons) = 1 then
						datoons = dateadd("d", flytreg+1, convertdateYMD(request("FM_ons_date")))
						end if
						if weekday(datoons) = 7 then
						datoons = dateadd("d", flytreg+2, convertdateYMD(request("FM_ons_date")))
						end if
					end if
					
					datotor = dateadd("d", flytreg, convertdateYMD(request("FM_tor_date")))
					if ikkeBruglorogson = 1 AND flytreg <> 0 then
						if weekday(datotor) = 1 then
						datotor = dateadd("d", flytreg+1, convertdateYMD(request("FM_tor_date")))
						end if
						if weekday(datotor) = 7 then
						datotor = dateadd("d", flytreg+2, convertdateYMD(request("FM_tor_date")))
						end if
					end if
					
					datofre = dateadd("d", flytreg, convertdateYMD(request("FM_fre_date")))
					if ikkeBruglorogson = 1 AND flytreg <> 0 then
						if weekday(datofre) = 1 then
						datofre = dateadd("d", flytreg+1, convertdateYMD(request("FM_fre_date")))
						end if
						if weekday(datofre) = 7 then
						datofre = dateadd("d", flytreg+2, convertdateYMD(request("FM_fre_date")))
						end if
					end if
					
					datolor = dateadd("d", flytreg, convertdateYMD(request("FM_lor_date")))
					if ikkeBruglorogson = 1 AND flytreg <> 0 then
						if weekday(datolor) = 1 then
						datolor = dateadd("d", flytreg+1, convertdateYMD(request("FM_lor_date")))
						end if
						if weekday(datolor) = 7 then
						datolor = dateadd("d", flytreg+2, convertdateYMD(request("FM_lor_date")))
						end if
					end if
					
					
					dagen = 1
					for dagen = 1 to 7 
						select case dagen
						case 1
						useday = datoson
						usetimer = "son"
						case 2
						useday = datoman
						usetimer = "man"
						case 3
						useday = datotir
						usetimer = "tir"
						case 4
						useday = datoons
						usetimer = "ons"
						case 5
						useday = datotor
						usetimer = "tor"
						case 6
						useday = datofre
						usetimer = "fre"
						case 7
						useday = datolor
						usetimer = "lor"
						end select
						
					useday = convertdateYMD(useday)
					
						for intcounter = 0 to antaljogakt
						
						usetimerakt = SQLBless(request("FM_"&usetimer&"_"&intcounter&"")) 
						useaktid = request("FM_aktid_"&intcounter&"") 
							call erDetInt(usetimerakt)
							if isInt = 0 then
								if len(useaktid) <> 0 AND len(usetimerakt) then
									strSQL3 = "SELECT timer FROM ressourcer WHERE aktid = " & useaktid &" AND dato = '"& useday &"' AND mid = "& medarbid
									oRec3.open strSQL3, oConn, 3
									if not oRec3.EOF then
									usetimerakt_pre = oRec3("timer")
									else
									usetimerakt_pre = 0
									end if
									oRec3.close
									
									usetimerakt = (usetimerakt + usetimerakt_pre)
									
									minutter = (usetimerakt * 60)
									useStarttp = "09:00:00"
									useSluttp = dateadd("n", minutter, useStarttp)
									'Response.write useStarttp&"<br>"& minutter &"<br>"& useSluttp & "<br>"
									
									if usetimerakt_pre = 0 then
									strSQL2 = "INSERT INTO ressourcer (jobid, aktid, dato, timer, mid, starttp, sluttp) VALUES ("& id &", "& useaktid &", '"& useday &"', "& usetimerakt &","& medarbid&", '"& useStarttp &"', '"& useSluttp &"')"
									else
									strSQL2 = "UPDATE ressourcer SET jobid = "& id &", aktid = "& useaktid &", dato = '"& useday &"', timer = "& usetimerakt &", mid = "& medarbid&" WHERE aktid = " & useaktid &" AND dato = '"& useday &"' AND mid = "& medarbid
									end if
									oConn.execute(strSQL2)
								end if
							isInt = 0 
							end if
						next
					next
					
	
	
	Response.redirect "ressourcer.asp?id="&id&"&intc="&intc
	
	case else
	%>
	<html>
		<head>
		<title>TimeOut 2.1</title>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
		<script LANGUAGE="javascript">
		function NewWin_popupcal(url) {
		window.open(url, 'Calpick', 'width=250,height=250,scrollbars=no,toolbar=no');    
		}
		</script>
		
		<script language="javascript">
		//function showeditdate(thiscnt){
		//opennow = document.all["lastmedarb"].value
		//document.all["editdate_"+opennow+"_"+thiscnt+""].style.display = "";
		//document.all["editdate_"+opennow+"_"+thiscnt+""].style.visibility = "visible";
		//}
		
		//function closeeditdate(thiscnt){
		//opennow = document.all["lastmedarb"].value
		//document.all["editdate_"+opennow+"_"+thiscnt+""].style.display = "none";
		//document.all["editdate_"+opennow+"_"+thiscnt+""].style.visibility = "hidden";
		//}
		
		function showuser(thiscnt){
		opennow = document.all["lastmedarb"].value
		document.all["medarb_"+opennow+""].style.display = "none";
		document.all["medarb_"+opennow+""].style.visibility = "hidden";
		document.all["medarb_"+thiscnt+""].style.display = "";
		document.all["medarb_"+thiscnt+""].style.visibility = "visible";
		document.all["lastmedarb"].value = thiscnt
		}
		</script>
		</head>
	<body topmargin="0" leftmargin="0" class="regular">
	
	<div id="loading" name="loading" style="position:absolute; top:80; left:490; width:300; height:20; z-index:1000; visibility:visible; padding-left:4px; padding-top:5px;">
	<!--<img src="../ill/info.gif" width="42" height="38" alt="" border="0">--><b>Henter information (kan tage optil 10 sek.)...</b><br><br>
	<img src="../ill/loaderbar.gif" width="174" height="13" alt="" border="0">
	<!--Denne handling kan tage op til 20 sek. alt efter din forbindelse.-->
	</div>
	
	
	<%
	'*** Sætter lokal dato/kr format. *****
	Session.LCID = 1030
	
	leftPos = 0
	topPos = 0
	
	dim medarbId
	dim medarbNavn
	dim medarbnormtimer
	%>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=leftPos%>; top:<%=topPos%>; visibility:visible;">
	
	<table cellspacing="0" cellpadding="0" border="0" width="880">
		<tr>
			<td bgcolor="#003399" width="650"><img src="../ill/logo_topbar_print.gif" alt="" border="0"></td>
			<td bgcolor="#FFFFFF" align=right><a href="javascript:window.close()"><img src="../ill/luk_xp.gif" width="30" height="28" alt="" border="0">&nbsp;Luk</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"><a href="javascript:window.print()"><img src="../ill/print_xp.gif" width="28" height="30" alt="" border="0">&nbsp;Print</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"></td>
		</tr>
		</table>
	<br>
	<div id="ressourcer" name="ressourcer" style="position:absolute; left:15; top:138; visibility:visible; display:;">
	<table cellspacing="0" cellpadding="0" border="0" width="200">
		<tr>
			<td>
			<table cellspacing="1" cellpadding="0" border="0" bgcolor="silver" width="100%">
			<tr bgcolor="#ffffff"><td valign=bottom width=110>&nbsp;<font class=megetlillesort>Medarbejder (norm. uge)</td><!--<td valign=bottom><font class=megetlillesort>Se tim.</td>--><td valign=bottom width=50>&nbsp;<font class=megetlillesort>Tild. på job</td></tr>
			
			
			<%
			strSQL = "SELECT ProjektgruppeId, MedarbejderId, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, "_
			&" timepris, kostpris, normtimer_son, normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, normtimer_fre, normtimer_lor, "_
			&" medarbejdertype, mid, mnavn, jobnavn, jobnr, job.fakturerbart, jobstartdato, jobslutdato, jobstatus "_
			&" FROM job, progrupperelationer LEFT JOIN medarbejdere ON (mid = MedarbejderId) LEFT JOIN medarbejdertyper ON (medarbejdertyper.id = medarbejdertype) WHERE job.id="& id &""_
			&" AND (ProjektgruppeId = projektgruppe1 OR ProjektgruppeId = projektgruppe2 OR ProjektgruppeId = projektgruppe3 OR ProjektgruppeId = projektgruppe4 OR ProjektgruppeId = projektgruppe5 OR ProjektgruppeId = projektgruppe6 OR ProjektgruppeId = projektgruppe7 OR ProjektgruppeId = projektgruppe8 OR ProjektgruppeId = projektgruppe9 OR ProjektgruppeId = projektgruppe10)"_
			&" GROUP BY mid"
			
			'Response.write strSQL
			oRec.open strSQL, oConn, 3
			
			x = 0
			Redim medarbId(x)
			Redim medarbNavn(x)
			Redim medarbnormtimer(x)
			
			
			while not oRec.EOF 
			if len(oRec("mid")) <> 0 then
				
			Redim preserve medarbId(x)
			Redim preserve medarbNavn(x)
			Redim preserve medarbnormtimer(x)
			
			if x = 0 then
			jobnavn = oRec("jobnavn")
			jid = id
			fakbart = oRec("fakturerbart")
			useLowerdatoKri = cdate(oRec("jobstartdato"))
			useUpperdatoKri = cdate(oRec("jobslutdato"))
			
			strDofyear = datepart("y", oRec("jobstartdato"))
			strDofyear_slut = datepart("y", oRec("jobslutdato"))
			
			jobdiff = datediff("w", useLowerdatoKri, useUpperdatoKri) 
			end if		
			
			medarbNavn(x) = oRec("mnavn")
			medarbId(x) = oRec("MedarbejderID") 'oRec("mid")
			 
			normtimer_son = normtimer_son + oRec("normtimer_son")
			normtimer_man = normtimer_man + oRec("normtimer_man")
			normtimer_tir = normtimer_tir + oRec("normtimer_tir")
			normtimer_ons = normtimer_ons + oRec("normtimer_ons")
			normtimer_tor = normtimer_tor + oRec("normtimer_tor")
			normtimer_fre = normtimer_fre + oRec("normtimer_fre")
			normtimer_lor = normtimer_lor + oRec("normtimer_lor")
			
			normtimermedarbweek = (oRec("normtimer_son") + oRec("normtimer_man") + oRec("normtimer_tir") + oRec("normtimer_ons") + oRec("normtimer_tor") + oRec("normtimer_fre") + oRec("normtimer_lor"))
			medarbnormtimer(x) = normtimermedarbweek
			
					strSQL3 = "SELECT sum(timer) AS sumtimerjob_medarb FROM ressourcer WHERE mid = "& medarbId(x) &" AND jobid ="&id
					'Response.write strSQL3
					oRec3.open strSQL3, oConn, 3 
					if not oRec3.EOF then
					timertotjob_medarb = oRec3("sumtimerjob_medarb") 
					end if
					oRec3.close
			
			%>
			<tr bgcolor="#ffffff"><td>&nbsp;<a href="#" onClick=showuser('<%=x%>') class=vmenu><%=left(oRec("mnavn"),18)%></a>&nbsp;&nbsp;<font class=megetlillesort>(<%=normtimermedarbweek%>)</td>
			<!--<td style="padding-top:3;" align=center><a href="javascript:NewWin_todo('medarb_restimer.asp?menu=medarb&func=show&FM_medarb=<=medarbId(x)%>&FM_medarbnavn=<=oRec("mnavn")%>') target=_top"><img src="../ill/popupcal.gif" width="16" height="15" alt="" border="0"></a></td>-->
			<td align=right style="padding-right:3;">
			<%Response.write "&nbsp;<font class=megetlillesort>"& timertotjob_medarb &"</td></tr>"
			x = x + 1
			end if
			oRec.movenext
			wend
			oRec.close
			
			
			totalstyrkeweek = (normtimer_son + normtimer_man + normtimer_tir + normtimer_ons + normtimer_tor + normtimer_fre + normtimer_lor)
			if len(totalstyrkeweek) <> 0 then
			totalstyrkeweek = totalstyrkeweek
			else
			totalstyrkeweek = 0
			end if
			
					strSQL3 = "SELECT sum(timer) AS sumtimerjob_all FROM ressourcer WHERE jobid ="&id
					oRec3.open strSQL3, oConn, 3 
					if not oRec3.EOF then
					timertotjob_ALL = oRec3("sumtimerjob_all") 
					end if
					oRec3.close
					
					if len(timertotjob_ALL) <> 0 then
					timertotjob_ALL = timertotjob_ALL
					else
					timertotjob_ALL = 0
					end if
					
					strSQL3 = "SELECT budgettimer, ikkebudgettimer FROM job WHERE id ="&id
					oRec3.open strSQL3, oConn, 3 
					if not oRec3.EOF then
					budgettimerThis = oRec3("budgettimer")
					ikkebudgettimerThis = oRec3("ikkebudgettimer")
					end if
					oRec3.close
					
					'*** tildelte timer på job ***
					totalbudgettimer = formatnumber((budgettimerThis + ikkebudgettimerThis),2)
					
					if len(totalbudgettimer) <> 0 then
					totalbudgettimer = totalbudgettimer
					else
					totalbudgettimer = formatnumber(0, 2)
					end if
					
					
					if len(ikkebudgettimerThis) <> 0 then
					ikkebudgettimerThis = formatnumber(ikkebudgettimerThis, 2)
					else
					ikkebudgettimerThis = formatnumber(0, 2)
					end if
					
					if len(budgettimerThis) <> 0 then
					budgettimerThis = formatnumber(budgettimerThis, 2)
					else
					budgettimerThis = formatnumber(0, 2)
					end if
				
		%>
		
			<tr bgcolor="#eff3ff"><td>&nbsp;<font class=lillesort><b>Total: (<%=formatnumber(totalstyrkeweek, 2)%>)</b></td><td align=right>&nbsp;<font class=lillesort><b><%=formatnumber(timertotjob_ALL, 2)%></b>&nbsp;</td></tr>
		</table>
		</td></tr></table>
		</div>
			
	<table cellspacing="1" cellpadding="1" border="0">
	<tr>
		<td style="padding-left:20;"><b>Ressource timer:</b><br>
		<font class=lillesort>
		Der kan tildeles timer på de medarbejdere, der er medlem af en af de <br>projektgrupper, der har adgang til dette job.<br>
		</font><br><b>Medarbejdere:</b>
		</td>
	</tr>
	</table>
	<%
	'*** sidst redigeret medarb ***
	intc = request("intc")
	if len(intc) <> 0 then
	intc = intc
	else
	intc = 0
	end if
	
	
	for intcounter = 0 to x - 1
	
	if cint(intcounter) = cint(intc) then
	vizdiv = "visibility"
	dispdiv = ""
	else
	vizdiv = "hidden"
	dispdiv = "none"
	end if
	%>
	<div id="medarb_<%=intcounter%>" name="medarb_<%=intcounter%>" style="position:absolute; left:230; top:122; visibility:<%=vizdiv%>; display:<%=dispdiv%>;">
	<table border=0 cellspacing=0 cellpadding=0 width=350 bgcolor="#eff3ff">
	<tr>
		<td colspan=2 bgcolor="#ffffff"><b>Job nøgletal.</b></td></tr>
	<tr>
		<td width=120 style="padding-left:4px; border-top:1px #5582d2 solid; border-left:1px #5582d2 solid;"><font class=lillesort>Periode: </td>
		<td style="border-top:1px #5582d2 solid; border-right:1px #5582d2 solid;"><font class=lillesort><b><%=formatdatetime(useLowerdatoKri, 1)%></b> - <b><%=formatdatetime(useUpperdatoKri, 1)%></b></td>
	</tr>
	<tr>
		<td style="padding-left:4px; border-left:1px #5582d2 solid; border-right:1px #5582d2 solid;" colspan=2><font class=lillesort><b>Forkalkuleret:</b></td>
	</tr>
	<tr>
		<td style="padding-left:4px; border-left:1px #5582d2 solid;"><font class=lillesort>Fakturerbare timer:</td>
		<td style="border-right:1px #5582d2 solid;"><font class=lillesort><b><%=budgettimerThis%></b></td>
	</tr>
	<tr>
		<td style="padding-left:4px; border-left:1px #5582d2 solid;"><font class=lillesort>Ikke fakturerbare timer:</td>
		<td style="border-right:1px #5582d2 solid;"><font class=lillesort><b><%=ikkebudgettimerThis%></b></td></tr>
	<tr>
		<td style="padding-left:4px; border-left:1px #5582d2 solid;"><font class=lillesort>Ialt: </td>
		<td style="border-right:1px #5582d2 solid;"><font class=lillesort><u><b><%=totalbudgettimer%></b></u></td>
	</tr>
	</table>
	
	<table cellspacing="1" cellpadding="0" border="0" bgcolor=#5582d2>
	<tr><td colspan="2" bgcolor=#ffffe1 style="Padding-left:5;"><font class=lillesort><b><%=left(medarbNavn(intcounter), 18)%></b></td>
	<td colspan="5" align="right" bgcolor=#ffffff style="Padding-right:5;"><font class=lillesort>tildelt totalt | <b>tildelt på job</b> | rest. timer i uge
	
	
				<%
				cnt = 0 
				ctd = 1
				for cnt = 0 to jobdiff
				
				getnyDatoMrd = dateadd("d", cnt*7, useLowerdatoKri)
				if lastm <> month(getnyDatoMrd) then
					if cnt > 0 then
					select case ctd 
					case 1
					Response.write "<td bgcolor=#ffffff colspan=4>&nbsp;</td>"
					case 2
					Response.write "<td bgcolor=#ffffff colspan=3>&nbsp;</td>"
					case 3
					Response.write "<td bgcolor=#ffffff colspan=2>&nbsp;</td>"
					case 4
					Response.write "<td bgcolor=#ffffff>&nbsp;</td>"
					end select
					
					end if%>
					</tr>
				<tr><td width="50" bgcolor=#ffffff style="padding-left:3; padding-right:3;" valign="top"><b><%=left(monthname(month(getnyDatoMrd)), 3)%>&nbsp;<%=right(year(getnyDatoMrd),2)%></b></td>
				<%
				ctd = 0
				end if%>
				
				<%	
					thisweek = datepart("ww", getnyDatoMrd)
					thisDay = Weekday(getnyDatoMrd)
					
					select case thisDay
					case 1
					stdagiuge = dateadd("d", 0, getnyDatoMrd)
					sldagiuge = dateadd("d", 6, getnyDatoMrd)
					case 2
					stdagiuge = dateadd("d", -1, getnyDatoMrd)
					sldagiuge = dateadd("d", 5, getnyDatoMrd)
					case 3
					stdagiuge = dateadd("d", -2, getnyDatoMrd)
					sldagiuge = dateadd("d", 4, getnyDatoMrd)
					case 4
					stdagiuge = dateadd("d", -3, getnyDatoMrd)
					sldagiuge = dateadd("d", 3, getnyDatoMrd)
					case 5
					stdagiuge = dateadd("d", -4, getnyDatoMrd)
					sldagiuge = dateadd("d", 2, getnyDatoMrd)
					case 6
					stdagiuge = dateadd("d", -5, getnyDatoMrd)
					sldagiuge = dateadd("d", 1, getnyDatoMrd)
					case 7
					stdagiuge = dateadd("d", -6, getnyDatoMrd)
					sldagiuge = dateadd("d", 0, getnyDatoMrd)
					end select
					
					strSQL3 = "SELECT sum(timer) AS sumtimer FROM ressourcer WHERE (dato >= '"&convertDateYMD(stdagiuge)&"' AND dato <= '"&convertDateYMD(sldagiuge)&"') AND mid=" & medarbId(intcounter)
					oRec3.open strSQL3, oConn, 3 
					if not oRec3.EOF then
					timertot =  oRec3("sumtimer")
					timerrest = (medarbnormtimer(intcounter) - oRec3("sumtimer"))
					end if
					oRec3.close
					
					strSQL3 = "SELECT sum(timer) AS sumtimerjob FROM ressourcer WHERE (dato >= '"&convertDateYMD(stdagiuge)&"' AND dato <= '"&convertDateYMD(sldagiuge)&"') AND mid=" & medarbId(intcounter) &" AND jobid ="&id
					oRec3.open strSQL3, oConn, 3 
					if not oRec3.EOF then
					timertotjob = oRec3("sumtimerjob") 
					end if
					oRec3.close
					
					if len(timertotjob) <> 0 then
						if timerrest > 0 then
						bgthis = "LightGreen"
						else
						bgthis = "mistyrose"
						end if
					else
						bgthis = "#ffffff"
					end if
					
					%>
					<td bgcolor=<%=bgthis%> style="padding-left:3;" width=100><a href="ressourcer.asp?id=<%=id%>&mid=<%=medarbId(intcounter)%>&func=editdate&stdagiuge=<%=stdagiuge%>&sldagiuge=<%=sldagiuge%>&intc=<%=intcounter%>"><%=thisweek%></a>
					<%
					Response.write "<font class=megetlillesilver><br>"& stdagiuge &"<br>"& sldagiuge &"<br>"
					Response.write "<font class=lillesort>" & timertot &"&nbsp;|&nbsp;</font>"
					Response.write "<font class=lillesort><b>" & timertotjob &"</b>&nbsp;|&nbsp;</font>"
					Response.write "<font class=lillesort>"& timerrest
					ctd = ctd + 1
				%>
		</td>		
		<%
		Response.flush
		lastm = month(getnyDatoMrd)
		'lasty = year(getnyDatoMrd)
		next
		%>
		<td valign="top" width=20 colspan="2" class=alt align="right">&nbsp;</td>
		</tr>
		</table>
	</div>
	<%
	lastm = 0
	next%>
	<form><input type="hidden" name="lastmedarb" id="lastmedarb" value="<%=intc%>"></form>
	<br><br>
	<br>
	</div>
	
	<script>
	document.all["loading"].style.visibility = "hidden";
	</script>
	<%
	end select
	
	
	
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
