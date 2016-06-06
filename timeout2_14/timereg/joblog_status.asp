<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<%
if session("user") = "" then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<%
	errortype = 5
	call showError(errortype)
	else
		if request("FM_medarb") = "" OR request("FM_job") = "" then%>
		<!--#include file="../inc/regular/header_inc.asp"-->
		<!--#include file="../inc/regular/topmenu_inc.asp"-->
		<%
		errortype = 40
		call showError(errortype)
		else
	%>
	<%if request("print") <> "j" then%>
	<!--#include file="../inc/regular/header_inc.asp"-->	
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	<%else%>
	<html>
	<head>
		<title>TimeOut 2.1</title>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
	</head>
	<body topmargin="0" leftmargin="0" class="regular">
	<%end if%>
	<!--#include file="inc/convertDate.asp"-->
	
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	</script>
	<%
	linket = request("linket")
	FM_medarb = request("FM_medarb")
	
	selmedarb = FM_medarb 'request("selmedarb")
	if len(selmedarb) = "0" then
	selmedarb = 0
	else
	selmedarb = selmedarb
	end if
	
	selaktid = request("selaktid")
	FM_job = request("FM_job")
	intJobnr = FM_job 'request("jobnr")
	
	if len(Request("mrd")) <> "0" AND request("mrd") <> 0 then
	strReqMrd = Request("mrd")
	else
	strReqMrd = 0
	end if
	
	'**** finder ud af om der er valgt et år ***
	if len(Request("year")) <> "0" then
		if Request("year") = "-1" then
		strReqAar = 0
		else
		strReqAar = Request("year")
		end if	
	else
	strReqAar = right(year(now), 2)
	end if
	
	thisfile = "joblog_status"
	if request("print") <> "j" then
	%>
	<!--#include file="inc/joblog_z_mrdk.asp"-->
	<!-- mrd knapper -->
	<%end if
	%>
	<!--include file="inc/stat_submenu.asp"-->
	<div style="position:absolute; left:700; top:165;">
	<a href="javascript:NewWin('joblog_status.asp?menu=stat&print=j&mrd=<%=strReqMrd%>&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>')" target="_self" class='rmenu'>&nbsp;Printer venlig version&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
	</div>
	<%
	if request("print") <> "j" then
	pleft = 190
	ptop = 55
	else
	pleft = 20
	ptop = 10
	end if
	%>	
	<div id="sindhold" style="position:absolute; left:<%=pleft%>; top:<%=ptop%>; visibility:visible; z-index:100;">
	<br><font class="pageheader"><b>Nøgletal og status</b></font><br>
	Oversigt over nøgletal og status på de valgte job.<br>
	Timeforbrug og omsætning for den valgte periode <b>omfatter kun de valgte medarbejdere</b>.
	<!-------------------------------Sideindhold------------------------------------->
<br><br>
<%'*** Sætter lokal dato/kr format. *****
Session.LCID = 1030

if len(Request("mrd")) <> "0" AND request("mrd") <> 0 then
	useSTdate = formatdatetime("1/"&strReqMrd&"/"&strReqAar, 2)
	useEndDate = formatdatetime(dateAdd("m", 1, useSTdate), 2) 
	else
	useSTdate = formatdatetime("1/1/"&strReqAar, 2)
	useEndDate = formatdatetime(dateAdd("m", 12, useSTdate), 2) 
end if
	
useSTdate = convertDateYMD(useSTdate)
useEnddate = convertDateYMD(useEnddate)

'Response.write "<br><br>" & useSTdate & "## " & useEnddate

'*** Finder de medarbejdere og job der er valgt ***
Dim intMedArbVal 
Dim b
Dim intJobnrKriValues 
Dim i

Dim j
j = 0
Dim totaltimer()
Dim timerTotpajob_mrd()
Dim timerTotpajob_mrd_fakbar()
Dim beloebTotpaajob()
Dim jobnavn()
Dim kundenavn()
Dim slutdato()
Dim jobnr()
Dim totaltimer_fakbar()
Dim jbudtim()
Dim fakturerbart()
Dim jobid()
Dim jbudget()
Dim jfastpris()
Dim beloebTotpaajob_mrd()
Dim kostprispaajob()
Dim kostprispaajob_mrd()
Dim antalrecords()
Dim avrtimepris()
Dim avrkostpris()
Dim antalFakbarerecords()
Dim jikkebudtim()
			
			
			'****** Finder de valgte medarbejdre og job til SQL kriterie **************
			intMedArbVal = Split(selmedarb, ", ")
			For b = 0 to Ubound(intMedArbVal)
				intJobnrKriValues = Split(intJobnr, ", ")
				For i = 0 to Ubound(intJobnrKriValues)
				if selmedarb <> "0" then
				jobMedarbKri = jobMedarbKri & " Tmnr = " & intMedArbVal(b) &" AND Tjobnr = " & intJobnrKriValues(i) &" AND tfaktim <> 5 OR "
				else
				jobMedarbKri = jobMedarbKri & " Tmnr <> 0 AND Tjobnr = " & intJobnrKriValues(i) &" AND tfaktim <> 5 OR "
				end if
				next
			next
			intJobMedarbKriCount = len(jobMedarbKri)
			trimIntJobMedarbKri = left(jobMedarbKri, (intJobMedarbKriCount-3))
			jobMedarbKri = trimIntJobMedarbKri & " AND "
			'***************************************************************************
			
			
			'************************************************************
			'*** Split medarbejdereKri igen. 
			'*** Denne bruges til sum funktionerne under hverenkelt job 
			'************************************************************
			medarbKri2 = "("
			For b = 0 to Ubound(intMedArbVal)
				if selmedarb <> "0" then
				medarbKri2 = medarbKri2 & " Tmnr = " & intMedArbVal(b) &" OR"
				else
				medarbKri2 = medarbKri2 & " Tmnr <> 0 "
				end if
			next
			intMedarbKriCount = len(medarbKri2)
			trimIntmedarbKri = left(medarbKri2, (intMedarbKriCount-3))
			medarbKri2 = trimIntmedarbKri &") "
			'**************************************************************

if len(request("selaktid")) <> "0" then
selaktidKri = "AND TAktivitetId = " & selaktid &""
else
selaktidKri = ""
end if
	
	lastjob = 0
	
	ReDim totaltimer(j)
	ReDim timerTotpajob_mrd(j)
	ReDim timerTotpajob_mrd_fakbar(j)
	ReDim beloebTotpaajob(j)
	ReDim jobnavn(j)
	ReDim kundenavn(j)
	ReDim slutdato(j)
	ReDim jobnr(j)
	ReDim totaltimer_fakbar(j)
	ReDim jbudtim(j)
	Redim fakturerbart(j)
	Redim jobid(j)
	Redim jbudget(j)
	Redim jfastpris(j)
	Redim beloebTotpaajob_mrd(j)
	Redim akttfaktimtildelt(j)
	Redim akttnotfaktimtildelt(j)
	Redim kostprispaajob(j)
	Redim kostprispaajob_mrd(j)
	Redim jikkebudtim(j)
	
	strSQL = "SELECT Tdato, TasteDato, Tjobnr, Tjobnavn, TAktivitetNavn AS Anavn, TAktivitetId, Tknavn, Tmnr, Tmnavn, Timer, Tid, Tfaktim, TimePris, timer.kostpris, Timerkom, timer.fastpris AS fastpris, job.jobTpris, job.budgettimer, job.jobslutdato, job.fakturerbart, job.id AS jid, ikkebudgettimer FROM timer LEFT JOIN job ON (job.jobnr = timer.Tjobnr) WHERE "& jobMedarbKri &" Tid > 0 "& selaktidKri &" ORDER BY Tjobnavn, Tjobnr"
	oRec.open strSQL, oConn, 3
	
	while not oRec.EOF
	
	''************** Finder de valgte job ***********************
			ReDim preserve totaltimer(j)
			ReDim preserve totaltimer_fakbar(j)
			ReDim preserve timerTotpajob_mrd(j)
			ReDim preserve timerTotpajob_mrd_fakbar(j)
			ReDim preserve beloebTotpaajob(j)
			ReDim preserve totaltimer_mrd_fakbar(j)
			ReDim preserve jobnavn(j)
			ReDim preserve jobnr(j)
			ReDim preserve kundenavn(j)
			ReDim preserve slutdato(j)
			ReDim preserve jbudtim(j)
			Redim preserve fakturerbart(j)
			Redim preserve jobid(j)
			Redim preserve jbudget(j)
			Redim preserve jfastpris(j)
			Redim preserve beloebTotpaajob_mrd(j)
			Redim preserve akttfaktimtildelt(j)
			Redim preserve akttnotfaktimtildelt(j)
			Redim preserve kostprispaajob(j)
			Redim preserve kostprispaajob_mrd(j)
			Redim preserve jikkebudtim(j)
			
			if lastjob <> oRec("Tjobnr") then
				
				kundenavn(j) = oRec("Tknavn")
				jobnavn(j) = oRec("Tjobnavn")
				jobnr(j) = oRec("Tjobnr")
				jbudtim(j) = oRec("budgettimer")
				if oRec("ikkebudgettimer") > 0 then
				jikkebudtim(j) = oRec("ikkebudgettimer")
				else
				jikkebudtim(j) = 0
				end if
				slutdato(j) = oRec("jobslutdato")
				fakturerbart(j) = oRec("fakturerbart")
				jobid(j) = oRec("jid")
				jbudget(j) = oRec("jobTpris")
				jfastpris(j) = oRec("fastpris")
					
					
					
					'******** Finder total omsætning på job ***
					strSQL3 = "SELECT timer, timepris FROM timer WHERE "& medarbKri2 &" AND Tjobnr=" & oRec("Tjobnr") &" AND tfaktim = 1"
					oRec3.open strSQL3, oConn, 3
					
					belpaajob = 0		
					while not oRec3.EOF 
					
						if oRec("fastpris") = "1" then
						belpaajob = belpaajob + ((oRec("jobTpris")/oRec("budgettimer")) * oRec3("timer"))
						else
						belpaajob = belpaajob + (oRec3("timer") * oRec3("timepris"))
						end if
					
					oRec3.movenext
					wend
					oRec3.close
					
					beloebTotpaajob(j) = belpaajob
					
					'******** Finder total omsætning på job i perioden ***
					strSQL3 = "SELECT timer, timepris FROM timer WHERE "& medarbKri2 &" AND Tjobnr=" & oRec("Tjobnr") &" AND tfaktim = 1 AND tdato BETWEEN '"& useSTdate &"' AND '"& useEnddate &"'"
					oRec3.open strSQL3, oConn, 3
					
					belpaajob_mrd = 0		
					while not oRec3.EOF 
					
						if oRec("fastpris") = "1" then
						belpaajob_mrd = belpaajob_mrd + ((oRec("jobTpris")/oRec("budgettimer")) * oRec3("timer"))
						else
						belpaajob_mrd = belpaajob_mrd + (oRec3("timer") * oRec3("timepris"))
						end if
					
					oRec3.movenext
					wend
					oRec3.close
					
					beloebTotpaajob_mrd(j) = belpaajob_mrd
			
					
					'******** Finder total kostpris på job ***
					strSQL3 = "SELECT timer, kostpris FROM timer WHERE "& medarbKri2 &" AND Tjobnr=" & oRec("Tjobnr") &" AND tfaktim <> 5"
					oRec3.open strSQL3, oConn, 3
					
					kostpaajob = 0		
					while not oRec3.EOF 
						
						if len(oRec3("kostpris")) <> 0 then  
						kostpaajob = kostpaajob + (oRec3("kostpris") * oRec3("timer"))
						else
						kostpaajob = kostpaajob
						end if
						
						
					oRec3.movenext
					wend
					oRec3.close
					
					kostprispaajob(j) = kostpaajob
					
					'******** Finder total kostpris på job i perioden ***
					strSQL3 = "SELECT timer, kostpris FROM timer WHERE "& medarbKri2 &" AND Tjobnr=" & oRec("Tjobnr") &" AND tfaktim <> 5 AND tdato BETWEEN '"& useSTdate &"' AND '"& useEnddate &"'"
					oRec3.open strSQL3, oConn, 3
					
					kostpaajob_mrd = 0		
					while not oRec3.EOF 
						
						if len(oRec3("kostpris")) <> 0 then  
						kostpaajob_mrd = kostpaajob_mrd + (oRec3("kostpris") * oRec3("timer"))
						else
						kostpaajob_mrd = kostpaajob_mrd
						end if
						
					oRec3.movenext
					wend
					oRec3.close
					
					kostprispaajob_mrd(j) = kostpaajob_mrd
					
					
					'** Alle timer **
					strSQL3 = "SELECT sum(timer) AS timerTotpajob FROM timer WHERE "& medarbKri2 &" AND Tjobnr=" & oRec("Tjobnr") &" AND tfaktim <> 5"
						oRec3.open strSQL3, oConn, 3
						
						if not oRec3.EOF then
						totaltimer(j) = oRec3("timerTotpajob")
						else
						totaltimer(j) = 0
						end if
						
					oRec3.close
					
					'*** Alle timer i periode ***
					strSQL3 = "SELECT sum(timer) AS timerTotpajob_mrd FROM timer WHERE "& medarbKri2 &" AND Tjobnr=" & oRec("Tjobnr") &" AND tfaktim <> 5 AND tdato BETWEEN '"& useSTdate &"' AND '"& useEnddate &"'"
						oRec3.open strSQL3, oConn, 3
						
						if not oRec3.EOF then
						timerTotpajob_mrd(j) = oRec3("timerTotpajob_mrd")
						else
						timerTotpajob_mrd(j) = 0
						end if
						
					oRec3.close
					%>
					
					
					<%'** Kun Fakturerbare **
					strSQL3 = "SELECT sum(timer) AS timerTotFaktimerpajob FROM timer WHERE "& medarbKri2 &" AND Tjobnr=" & oRec("Tjobnr") &" AND Tfaktim  = 1"
						oRec3.open strSQL3, oConn, 3
						
						if not oRec3.EOF then
						totaltimer_fakbar(j) = oRec3("timerTotFaktimerpajob")
						end if
						
					oRec3.close
					
					
					'** Kun Fakturerbare i periode **
					strSQL3 = "SELECT sum(timer) AS timerTotFaktimerpajob_mrd FROM timer WHERE "& medarbKri2 &" AND Tjobnr=" & oRec("Tjobnr") &" AND Tfaktim  = 1 AND tdato BETWEEN '"& useSTdate &"' AND '"& useEnddate &"'"
						oRec3.open strSQL3, oConn, 3
						
						if not oRec3.EOF then
						timerTotpajob_mrd_fakbar(j)  = oRec3("timerTotFaktimerpajob_mrd")
						else
						timerTotpajob_mrd_fakbar(j) = 0
						end if
						
					oRec3.close
					
					
					
					'*** fakturerbare timer tildelt på aktiviteter **** 
					strSQL3 = "SELECT sum(budgettimer) AS akttimer FROM aktiviteter WHERE job = " & jobid(j) &" AND fakturerbar = 1 ORDER BY budgettimer"
					oRec3.open strSQL3, oConn, 3
					if not oRec3.EOF then
						akttfaktimtildelt(j) = oRec3("akttimer")
					end if
					oRec3.close
					
					if len(akttfaktimtildelt(j)) <> 0 then
					akttfaktimtildelt(j) = akttfaktimtildelt(j)
					else
					akttfaktimtildelt(j) = 0
					end if
					
					'***Ikke fakturerbare timer tildelt på aktiviteter **** 
					strSQL3 = "SELECT sum(budgettimer) AS aktnotftimer FROM aktiviteter WHERE job = " & jobid(j) &" AND fakturerbar = 0 ORDER BY budgettimer"
					oRec3.open strSQL3, oConn, 3
					if not oRec3.EOF then
						akttnotfaktimtildelt(j) = oRec3("aktnotftimer")
					end if
					oRec3.close
					
					if len(akttnotfaktimtildelt(j)) <> 0 then
					akttnotfaktimtildelt(j) = akttnotfaktimtildelt(j)
					else
					akttnotfaktimtildelt(j) = 0
					end if
					
			end if
			
					
	if lastjob <> oRec("Tjobnr") then
	lastjob = oRec("Tjobnr")
	j = j + 1
	end if
	
	
	oRec.movenext
	wend	  
	
	oRec.Close
	Set oRec = Nothing
%>
<br><br><br>				

<%

for j = 0 to j - 1%>
			<table border="0" width="420" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
					<tr bgcolor="#5582D2">
						<td width="8" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="20" alt="" border="0"></td>
						<td colspan="3" valign="top"><img src="../ill/tabel_top.gif" width="404" height="1" alt="" border="0"></td>
						<td align=right valign=top><img src="../ill/tabel_top_right.gif" width="8" height="20" alt="" border="0"></td>
					</tr>
					<%if fakturerbart(j) = 1 then%>
					<tr height="20">
						<td valign="top" rowspan="28"><img src="../ill/tabel_top.gif" width="1" height="540" alt="" border="0"></td>
						<td class='alt' colspan="3">&nbsp;&nbsp;<img src="../ill/eksternt_job_trans.gif" width="15" height="15" alt="" border="0">&nbsp;&nbsp;<a href="jobs.asp?menu=job&func=red&id=<%=jobid(j)%>&int=<%=fakturerbart(j)%>"><%=jobnr(j)%>&nbsp;<b><%=jobnavn(j)%></b></a>&nbsp;(<%=kundenavn(j)%>)<br>
						<%
						'*** Jobansvarlige ***
						
						strSQL2 = "SELECT mnavn, mnr, mid, job.id, jobans1 FROM job, medarbejdere WHERE job.id = "& jobid(j) &" AND mid = jobans1"
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						jobans1txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
						end if
						oRec2.close
						
						
						strSQL2 = "SELECT mnavn, mnr, mid, jobans2 FROM job, medarbejdere WHERE job.id = "& jobid(j) &" AND mid = jobans2"
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						jobans2txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
						end if
						oRec2.close
						
						
						%>
						<font class=megetlillesilver>
						Jobansvarlig 1: <%=jobans1txt%><br>
	   					Jobansvarlig 2: <%=jobans2txt%></font>
						
						</td>
						<td valign="top" align=right rowspan="28"><img src="../ill/tabel_top.gif" width="1" height="540" alt="" border="0"></td>
					</tr>	
					<tr><td colspan="3"><br>
							<%if jfastpris(j) = 1 then%>
							<b>Fastpris:</b>
							<%else%>
							<b>Budget: (Ikke fastpris)</b>
							<%end if%>
							&nbsp;&nbsp;<%=formatcurrency(jbudget(j),2)%>
							</td></tr>
							
							<tr bgcolor="#EFF3FF"><td colspan="3"><b>Forkalkuleret timer:</b></td></tr>
							<tr><td width="120" valign=bottom>Timer tildelt </td><td width="100">Fakturerbare</td><td width="200">Ikke fak. bare:</td></tr>
							<tr><td valign=top><%=cdbl(jbudtim(j)+jikkebudtim(j))%></td><td valign=top><font color="DarkOrange"><%=jbudtim(j)%>&nbsp;(<%=akttfaktimtildelt(j)%>)&nbsp;</font></td><td valign=top><font color="LimeGreen">&nbsp;<%=jikkebudtim(j)%>&nbsp;(<%=akttnotfaktimtildelt(j)%>)</font>
							</td></tr>
							<tr><td colspan="3"><%if jfastpris(j) = 1 then
							Response.write "Beregnet timepris:&nbsp;<u>" & formatcurrency((jbudget(j)/jbudtim(j)),2) &"</u><br>"
							end if%>&nbsp;</td></tr>
							
							<tr bgcolor="#EFF3FF"><td colspan="3"><b>Timeforbrug:</b></td></tr>
							<tr><td>Total</td><td>Fakturerbare</td><td>Ikke fak. bare:</td></tr>
							<tr><td valign=top>
							<%=totaltimer(j)%></td><td valign=top><font color="DarkOrange"><%
							if len(totaltimer_fakbar(j)) <> 0 then
							response.write totaltimer_fakbar(j)
							else
							totaltimer_fakbar(j) = 0
							Response.write totaltimer_fakbar(j)
							end if%></font></td><td valign=top><font color="LimeGreen">
							<%
							totaltimer_ikkefak = (totaltimer(j) - totaltimer_fakbar(j))
							Response.write totaltimer_ikkefak&"</font>"
							%></td></tr>
							<tr><td colspan="3"><%if jfastpris(j) = 1 then
							Response.write "Faktisk timepris:&nbsp;"
							faktimepris = (jbudget(j)/totaltimer_fakbar(j))
							Response.write "<u>"& formatcurrency(faktimepris, 2) &"</u><br>"
							end if%>&nbsp;</td></tr>
							
							
							<tr bgcolor="#EFF3FF"><td colspan="3"><b>Resterende timer:</b></td></tr>
							<tr><td>Total</td><td>Fakturerbare</td><td>Ikke fak. bare:</td></tr>
							<tr><td>
							<%resttimertot = ((jbudtim(j)+jikkebudtim(j)) -totaltimer(j))
							Response.write resttimertot%></td><td><font color="DarkOrange"><%
							resttimerfak = (jbudtim(j) - totaltimer_fakbar(j))
							Response.write resttimerfak
							%></font></td><td><font color="LimeGreen">
							<%
							resttimer_ikkefak = (jikkebudtim(j) - totaltimer_ikkefak)
							Response.write resttimer_ikkefak&"</font>"
							%></td>
							</tr>
							
							<tr height="10">
								<td class='alt' colspan="3">&nbsp;<img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
							</tr>
							
							<tr>
								<td valign="middle"><b>% af job udført:</b></td>
								<td valign="top" colspan="2"><%
								totbudgettimer = (jbudtim(j)+jikkebudtim(j))
								
								if len(totbudgettimer) <> 0 AND totbudgettimer > 0 then
								totbudgettimer = totbudgettimer
								else
								totbudgettimer = 1
								end if
								
								overh = ((totaltimer(j) / totbudgettimer) * 100)
								if totaltimer(j) <> 0 then
								projektcomplt = FormatPercent(totaltimer(j) / totbudgettimer, 1)
								else
								projektcomplt = "0 %"
								end if
								
								if overh > 100 then%>
								100 %
								<%else%>
								<%=projektcomplt%>
								<%end if%>
								<%
								if overh > 100 then
								complgif = 100
								else
								complgif = projektcomplt
								end if
								%>
								<div style="!border: 1 px;background-color:#ffffff; border-color: black; border-style: solid; height:6px; width:100px;">
								<img src="../ill/completed.gif" width="<%=complgif%>" height="6" alt="" border="0"></div></td>
							</tr>
							
							<tr><td><b>Forventet slut dato:</b></td><td colspan="2"><%=formatdatetime(slutdato(j), 1)%></td></tr>
							<tr height="10">
								<td class='alt' colspan="3">&nbsp;<img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
							</tr>
							<tr bgcolor="WhiteSmoke">
								<td valign="top" style="padding-top:3px; padding-left:5px; border:1px DarkOrange solid; border-right:0px;"><b>Omsætning:</b>
								<%if jfastpris(j) = 1 then%><br>
								<font size=1>Beregnet omsætning:<br>
								<u>Timer x (fastprisbeløbet / tildelte fak.bare timer)</u></font>
								<%else%>
								<font size=1><br><u>timer x medarbejdertypes timepris</u></font>
								<%end if%></td>
								<td colspan="2" valign="top" style="padding-top:3px; border:1px DarkOrange solid; border-left:0px;">
								<%if jfastpris(j) = 1 then
									Response.write formatcurrency(jbudget(j),2) &"&nbsp;&nbsp;"
									
									if len(beloebTotpaajob(j)) <> 0 then
									Response.write "<br><font size=1>"& formatcurrency(beloebTotpaajob(j), 2)
									else
									Response.write "<br><font size=1>"& formatcurrency(0, 2)
									end if
								
								else
									
									if len(beloebTotpaajob(j)) <> 0 then
										Response.write formatcurrency(beloebTotpaajob(j), 2)
									else
										Response.write formatcurrency(0, 2)
									end if
								
								end if%></td>
							</tr>
							
							<tr height="2">
								<td class='alt' colspan="3"><img src="../ill/blank.gif" width="1" height="2" alt="" border="0"></td>
							</tr>
							
							<tr bgcolor="WhiteSmoke"><td valign=top style="padding-top:3px; padding-left:5px; border:1px silver solid; border-right:0px;"><b>Balance:</b>
							<%if jfastpris(j) = 1 then%><br>
							<font size=1><u>Fastprisbeløbet - beregnet omsætning</u></font>
							<%else%><br>
							<font size=1><u>Budget - totalomsætning</u></font>
							<%end if%></td>
							<td colspan="2" valign=top style="padding-top:3px; border:1px silver solid; border-left:0px;">
							<%if jfastpris(j) = 1 then
							'Response.write formatcurrency(0, 2) &"<br>"
							rebudget = (jbudget(j) - beloebTotpaajob(j))
							Response.write formatcurrency(rebudget, 2) 
							else
							rebudget = (jbudget(j) - beloebTotpaajob(j))
							Response.write formatcurrency(rebudget, 2)
							end if
							%>
							</td></tr>
							
							
							<tr><td style="padding-top:3px; padding-left:5px;"><b>Intern kostpris:</b> 
							<!--<br><font size="1">(Kun timer indtastet efter d. 19/9 2003, der kan have en kostpris. Kostprisen sættes på medarbejder-typen.)-->
							</td>
							<td valign="bottom" colspan="2" style="padding-top:3px; border-bottom : 1px; border-top : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : dashed;">
							<font color="darkred">
							<%
							if len(kostprispaajob(j)) <> 0 then
								Response.write formatcurrency(kostprispaajob(j), 2)
							else
								kostprispaajob(j) = 0
								Response.write formatcurrency(kostprispaajob(j), 2) 
							end if%></td></tr>
							
							
							
							<tr height="10">
								<td class='alt' colspan="3">&nbsp;<img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
							</tr>
							<tr bgcolor="#EFF3FF"><td colspan="3"><b>Timeforbrug i den valgte periode:</b></td></tr>
							<tr><td>Total</td><td>Fakturerbare</td><td>Ikke fak. bare:</td></tr>
							<tr><td><%
							if len(timerTotpajob_mrd(j)) <> 0 then
							Response.write timerTotpajob_mrd(j)
							Else
							timerTotpajob_mrd(j) = 0
							Response.write timerTotpajob_mrd(j)
							end if
							%></td><td><font color="DarkOrange"><%
							if len(timerTotpajob_mrd_fakbar(j)) <> 0 then
							Response.write timerTotpajob_mrd_fakbar(j)
							else
							timerTotpajob_mrd_fakbar(j) = 0
							Response.write timerTotpajob_mrd_fakbar(j)
							end if
							%></font></td><td><font color="LimeGreen">
							<%
							timerTotpajob_mrd_ikkefakbar = (timerTotpajob_mrd(j) - timerTotpajob_mrd_fakbar(j))
							Response.write timerTotpajob_mrd_ikkefakbar%>
							</font>
							</td></tr>
							
							<tr><td><b>Oms. i periode:</b></td><td colspan="2" style="border-top : 1px; border-bottom : 1px; border-left : 0px; border-right : 0px; border-color : #CCCCCC; border-style : solid;"><%
							if len(beloebTotpaajob_mrd(j)) <> 0 then
							Response.write formatcurrency(beloebTotpaajob_mrd(j),2) 
							else
							Response.write formatcurrency(0, 2)
							end if%>
							</td></tr>
							
							<tr><td><b>Kostpris i periode:</b></td><td colspan="2" style="border-top : 0px; border-bottom : 1px; border-left : 0px; border-right : 0px; border-color : #CCCCCC; border-style : dashed;">
							<font color="darkred"><%
							if len(kostprispaajob_mrd(j)) <> 0 then
							Response.write formatcurrency(kostprispaajob_mrd(j),2)
							else
							Response.write formatcurrency(0, 2)
							end if%>
							</td></tr>
							
							
							
					<%else%>
							<tr height="20">
								<td valign="top" rowspan="4"><img src="../ill/tabel_top.gif" width="1" height="70" alt="" border="0"></td>
								<td class='alt' colspan="3">&nbsp;&nbsp;<img src="../ill/internt_job_trans.gif" width="15" height="15" alt="" border="0">&nbsp;&nbsp;<a href="jobs.asp?menu=job&func=red&id=<%=jobid(j)%>&int=<%=fakturerbart(j)%>"><%=jobnr(j)%>&nbsp;<b><%=jobnavn(j)%></b></a>&nbsp;(<%=kundenavn(j)%>)</td>
								<td valign="top" align=right rowspan="4"><img src="../ill/tabel_top.gif" width="1" height="70" alt="" border="0"></td>
							</tr>	
							<tr><td colspan="3"><b>Timeforbrug total:</b>&nbsp;&nbsp;&nbsp;<%=totaltimer(j)%></td></tr>
							<tr><td colspan="3"><b>Timeforbrug i den valgte periode:</b>&nbsp;&nbsp;&nbsp;<%
							if len(timerTotpajob_mrd(j)) <> 0 then
							Response.write timerTotpajob_mrd(j)
							else
							Response.write 0
							end if%></td></tr>
					<%end if%>
					<tr height="20">
						<td class='alt' colspan="3">&nbsp;<img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
					</tr>
					<tr height="20" bgcolor="#5582D2">
						<td class='alt' colspan="5">&nbsp;<img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
					</tr>
					</table><br><br>
<%next%>
<br><br>

<%if j = 0 then%>
	<div id="ingenregi" style="position:absolute; left:-110; top:100; width:60%; height:600; visibility:visible; z-index:100;">
	<table width="480" celpadding="2" cellspacing="1" border="0">
	<tr><td height="50" align="center" valign="middle"><img src="../ill/alert.gif" width="44" height="45" alt="" border="0">&nbsp;&nbsp;Ingen registreringer i den valgte periode.<br><b><%=noJobMessage%></b></td></tr></table>
	</div>
<%end if%>

</div>
<br>
<br>
<%end if 'validering%>
<%end if %>
<!--#include file="../inc/regular/footer_inc.asp"-->