<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
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
		if request("print") <> "j" then%>	
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
		<script>
		//Aktiviteter expand
		 if (document.images){
				plus = new Image(200, 200);
				plus.src = "ill/plus.gif";
				minus = new Image(200, 200);
				minus.src = "ill/minus2.gif";
				}
		
				function expand (nummer){
					if (document.all(nummer)){
						if (document.all(nummer).style.display == "none"){
						document.all(nummer).style.display = "";
						document.images[nummer+"ikon"].src = minus.src;
					}else{
						document.all(nummer).style.display = "none";
						document.images[nummer+"ikon"].src = plus.src;
					}
				}
			}
		</script>
<%
'****************************************************************************************************
'****************   Variable og kriterier for medarbejdere, job og periode   ************************
'****************************************************************************************************

	thisfile = "joblog_z_b"
	vis_medarb = request("FM_vis_medarb")
	vis_akt = request("FM_vis_akt")
	vis_medarb_k = "j" 'request("FM_vis_medarb_k")
	
	if len(request("FM_medarb")) <> 0 then
	FM_medarb = request("FM_medarb")
	else
	FM_medarb = request("FM_medarb_alle")
	end if
	
	'*** Er der valgt alle job eller et enkelt job? ******* 
	erdervalgtalle = instr(request("FM_job"), ",")
	
	seljob = request("FM_seljob")
	if len(seljob) <> 0 then
	seljob = seljob 
	else
	seljob = -1
	end if
	
	
	'Response.write seljob &"<br>" & erdervalgtalle &"<br>"
	
	'*** Hvis der kun er valgt 1 job *****
	if erdervalgtalle = 0 AND len(request("FM_job")) <> 0 then
		if seljob <> 0 then  'Hvis der valgt 1 job fra stat siden.
			FM_job = request("FM_job")
			showallealle = "n"
			'Response.write "A"
		else				'Hvis der valgt 1 job fra stat siden og der er valg alle/alle i dd på joblog_z_b
			showallealle = "j"
			FM_job = request("FM_job")
			'Response.write "B"
		end if
	else
	'*** Hvis der er valgt mere et 1 job, eller slet ingen (alle)'**
		if len(request("FM_job")) <> 0 then
			if seljob > 0 then 
			FM_job = request("FM_job")
			showallealle = "n"
			'Response.write "C"
			else
			showallealle = "j"
			FM_job = request("FM_job")
			'Response.write "D"
			end if
		else
		showallealle = "j"
		FM_job = request("FM_job_alle")
		'Response.write "E"
		end if
	end if
	
	
	intJobnr = FM_job 'request("jobnr")
	func = request("func")
	 
	'linket = request("linket")
	
	'**** De valgte medarbejdere *************
		selmedarb = FM_medarb 'request("selmedarb")
		if len(selmedarb) = "0" then
		selmedarb = 0
		else
		selmedarb = selmedarb
		end if
	'*****************************************
	
	'**** til normtimer ****
	dim cntthisjobid
	
	Dim medarbnavn()
	m = 0
	Redim medarbnavn(m)
	
	'*** Finder de medarbejdere og job der er valgt *******************
	Dim intMedArbVal 
	Dim b
	Dim intJobnrKriValues 
	Dim i
	
	'** Bruges til DDown og opbygning af valgte job ***
	antalvalgtejob = 0 'antal valgte job
	intJobnrDD = Split(intJobnr, ", ")
	printhvilkejob = "<tr><td colspan=5>Følgende job er omfattet af totaler pr. medarbjer i denne udskrift:</td></tr>"
		Redim cntthisjobid(i)
		For i = 0 to Ubound(intJobnrDD)
		
			strSQL = "SELECT id, jobnavn, jobstartdato, jobslutdato, kkundenavn, jobtpris, budgettimer, fakturerbart, fastpris, ikkebudgettimer FROM job, kunder WHERE jobnr = " & intJobnrDD(i) &" AND kid = jobknr"
			'Response.write strSQL
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			Redim preserve cntthisjobid(i)
			cntthisjobid(i) = oRec("id")
			strJobnavnThis = oRec("jobnavn")
			strKundenavnThis = oRec("kkundenavn")
			jStartdato = oRec("jobstartdato")
			jSlutdato = oRec("jobslutdato")
			jTpris = oRec("jobtpris")
			jbudgettimer = oRec("budgettimer")
			jikkebudgettimer = oRec("ikkebudgettimer")
			jFakturerbart = oRec("fakturerbart")
			jFastpris = oRec("fastpris")
			jid = oRec("id")
			else
			strJobnavnThis = "jobnavn ikke angivet!"
			end if
			oRec.close
			
			if seljob = intJobnrDD(i) OR seljob = -1 then
				if showallealle = "n" then
				selthis = "SELECTED"
				else
				selthis = ""
				end if
			
			useJobnavn = strJobnavnThis
			useJobnr = intJobnrDD(i)
			useKundenavn = strKundenavnThis
			useStdato = jStartdato
			useSldato = jSlutdato
			useTpris = jTpris
			
					useBudgettimer = jbudgettimer
					if jikkebudgettimer > 0 then
					useIkkebudgettimer = jikkebudgettimer
					else
					useIkkebudgettimer = 0
					end if
					
					useFakbart = jFakturerbart
					if useFakbart = 1 then
					fakbargif = "<img src='../ill/eksternt_job_trans.gif' width='15' height='15' alt='Eksternt job' border='0'>"
					else
					fakbargif = "<img src='../ill/internt_job_trans.gif' width='15' height='15' alt='Internt job' border='0'>"
					end if
					
			useFastpris = jFastpris
			useJid = jid
			else
			selthis = ""
			end if
			
			
			ddjobnr = ddjobnr &"<option value="&intJobnrDD(i)&" "& selthis &">"&intJobnrDD(i)&"&nbsp;"&left(strJobnavnThis, 30)&"</option>" 
			antalvalgtejob = antalvalgtejob + 1
			if request("print") = "j" then
				if (right(i, 1) = 5 OR right(i, 1) = 0) then 
				printhvilkejob = printhvilkejob &"</tr><tr><td width=150><b>"& intJobnrDD(i) &"</b>&nbsp;" & left(strJobnavnThis, 20) &"</td>"
				else
				printhvilkejob = printhvilkejob &"<td width=150><b>"& intJobnrDD(i) &"</b>&nbsp;" & left(strJobnavnThis, 20) &"</td>"
				end if
			end if
		next
		printhvilkejob = printhvilkejob & "</tr></table>"
		'******************************************************************
	
	
	
	
	
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	useTotal = request("FM_vis_total")
	'FM_Aar = request("FM_Aar") '???
	weekselector = request("weekselector") 'Er weekselector (periodevælger) brugt? ret:j/null
	strYear = request("year") 'Det valgte år
	yearsel = request("yearsel") 'Er det klikket på et år? ret: j/null
	
	'*** periode vælger er brugt ****
	 if len(request("FM_start_dag")) <> 0 then
	'* Bruges i weekselector *
	strMrd =  request("FM_start_mrd")
	strDag  = request("FM_start_dag")
	if len(request("FM_start_aar")) > 2 then
	strAar = right(request("FM_start_aar"), 2)
	else
	strAar = request("FM_start_aar")
	end if
	
	strMrd_slut =  request("FM_slut_mrd")
	strDag_slut  =  request("FM_slut_dag")
	
	if len(request("FM_slut_aar")) > 2 then
	strAar_slut = right(request("FM_slut_aar"),2)
	else
	strAar_slut = request("FM_slut_aar")
	end if
	
	
	select case strMrd
	case 2
		if strDag > 28 then
		strDag = 28
		else
		strDag = strDag
		end if
	case 4, 6, 9, 11
		if strDag > 30 then
		strDag = 30
		else
		strDag = strDag
		end if
	end select
	
	select case strMrd_slut
	case 2
		if strDag_slut > 28 then
		strDag_slut = 28
		else
		strDag_slut = strDag_slut
		end if
	case 4, 6, 9, 11
		if strDag_slut > 30 then
		strDag_slut = 30
		else
		strDag_slut = strDag_slut
		end if
	end select
	
	'*** Den valgte start og slut dato ***
	StrTdato = strDag &"/" & strMrd & "/20" & strAar
	StrUdato = strDag_slut &"/" & strMrd_slut & "/20" & strAar_slut
	
	else
	StrTdato = date-31
	StrUdato = date 
	
	'* Bruges i weekselector *
	if month(now()) = 1 then
	strMrd = 12
	else
	strMrd = month(now()) - 1
	end if
	
	strDag = day(now())
	
	if month(now()) = 1 then
	strAar = right(year(now()) - 1, 2)
	else
	strAar = right(year(now()), 2) 
	end if
	
	strMrd_slut = month(now())
	strAar_slut = right(year(now()), 2) 
	
	if strDag > "28" then
	strDag_slut = "1"
	strMrd_slut = strMrd_slut + 1
	else
	strDag_slut = day(now())
	end if
	end if
	
	periodeinterval = datediff("d", StrTdato, StrUdato)
	periodeintervalWeeks = datediff("ww", StrTdato, StrUdato)
	
	'** Tjek om periode er negativ **
	'********************************
	
	if len(Request("mrd")) <> "0" then
		strReqMrd = Request("mrd")
	else
	strReqMrd = month(now)
	end if
	
	
	'** Indsætter cookie **
	Response.Cookies("st_dag") = strDag
	Response.Cookies("st_dag").Expires = date + 40
	
	Response.Cookies("st_md") = strMrd
	Response.Cookies("st_md").Expires = date + 40
	
	Response.Cookies("st_aar") = strAar
	Response.Cookies("st_aar").Expires = date + 40
	
	
	Response.Cookies("sl_dag") = strDag_slut
	Response.Cookies("sl_dag").Expires = date + 40
	
	Response.Cookies("sl_md") = strMrd_slut
	Response.Cookies("sl_md").Expires = date + 40
	
	Response.Cookies("sl_aar") = strAar_slut
	Response.Cookies("sl_aar").Expires = date + 40
	

'**************************************************************************************
'Funktioner der bruges på siden.					 								  *
'**************************************************************************************





'************************************************
'****** Medarbejer Fleks regnskab funktion ******
function medarbejderFleksNorm(jobaktid, jobElaktValgt)
select case jobElaktValgt 
case "J"
bgthis = "#EFF3FF"
jobaktidKri = " jobid = "& jobaktid
case "A"
jobaktidKri = " aktid = "& jobaktid 
bgthis = "#FFFFF1"
case "J_all"
bgthis = "#EFF3FF"
	jobaktidKri = " ( "
	for intcounter = 0 to jobaktid - 1
	jobaktidKri = jobaktidKri & " jobid = "& cntthisjobid(intcounter) &" OR "
	next
	
	LEN_jobaktidKri = len(jobaktidKri)
	temp_jobaktidKri= left(jobaktidKri, (LEN_jobaktidKri - 3))
	temp_jobaktidKri = temp_jobaktidKri &" ) "
	jobaktidKri = temp_jobaktidKri
end select
%>
<table border="0" cellspacing="1" cellpadding="0" width="550">
								<tr>
									<td width="300" bgcolor=<%=bgthis%> valign=bottom style="padding-top:2; padding-left:2; border:#8CAAE6 1px solid;">
									<%if jobElaktValgt <> "A" then%>
									<b>Normeret tid / Flekstid:</b>
									<br><font class=megetlillesort>I den valgte periode.</font>
									<%else%>
									<b>Ressource timer på aktivitet.</b>
									<%end if%></td>
									<td width="50" bgcolor=<%=bgthis%> align=center valign=bottom style="padding-top:2; padding-left:2; border:#8CAAE6 1px solid;">&nbsp;</td>
									<td width="80" bgcolor=<%=bgthis%> align=center valign=bottom style="padding-top:2; padding-left:2; border:#8CAAE6 1px solid;">
									<%if jobElaktValgt <> "A" then%>
									<b>Timer brugt</b><br>
									<%end if%>
									<font color="darkred">&sum;</font> Tot.</td>
									<td width="120" bgcolor=<%=bgthis%> align=center valign=bottom style="padding-top:2; padding-right:5; border:#8CAAE6 1px solid;">
									<%if jobElaktValgt <> "A" then%>
									<b>Flekstid / Balance <br>
									<%end if%>
									+ / -</b></td>
								</tr>
								<%
								'****************** fleks regnskab / normeret tid ********************
								strSQL = "SELECT mid, medarbejdertype, normtimer_son, id, normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, normtimer_fre, normtimer_lor FROM medarbejdertyper, medarbejdere WHERE medarbejdere.mid = "& intMedArbValThis(m) &" AND id = medarbejdertype"  
								oRec.open strSQL, oConn, 3 
								if not oRec.EOF then
								
								sontimer = antalson * oRec("normtimer_son")
								mantimer = antalman * oRec("normtimer_man")
								tirtimer = antaltir * oRec("normtimer_tir")
								onstimer = antalons * oRec("normtimer_ons")
								tortimer = antaltor * oRec("normtimer_tor")
								fretimer = antalfre * oRec("normtimer_fre")
								lortimer = antallor * oRec("normtimer_lor")
								
								
								thisnormeret = (sontimer + mantimer + tirtimer + onstimer +tortimer + fretimer + lortimer)
								balancethis_normeret = thisnormeret - timertot   
								
								if jobElaktValgt <> "A" then
 									Response.write "<tr><td valign=top align=right style=""border:1px #8CAAE6 solid; padding:1px;"">Normeret tid på medarbejdertype:<font class=megetlillesort><br><u>Flekstid / Balance:</u> Kun timer på de(t) valgte job er medregnet.</font></td>"
									Response.write "<td valign=top bgcolor=#FFFFFF align=right style=""border:1px #8CAAE6 solid; padding-right:5px;"">"& formatnumber(thisnormeret,2) &"</td>"
									Response.write "<td valign=top bgcolor=#FFFFFF align=right style=""border:1px #8CAAE6 solid; padding-right:5px;""><font color='darkred'>"& formatnumber(timertot,2) &"</td>"
									Response.write "<td valign=top bgcolor=#FFFFFF align=right style=""border:1px #8CAAE6 solid; padding-right:5px;"">"& formatnumber(balancethis_normeret, 2) &"</td></tr>"
								end if
								
								Response.write "<tr><td align=right style=""border:1px #8CAAE6 solid; padding:1px;"">Tildelte ressource timer."
								if jobElaktValgt <> "A" then
								Response.write "<br><font class=megetlillesort>Kun på <u>de(t) valgte job</u>.</font></td>"
								else
								Response.write "<br><font class=megetlillesort>Denne medarbejder på <u>ovenstående aktivitet</u> i den valgte periode.</font></td>"
								end if
								Response.write "<td valign=top align=right bgcolor=#FFFFFF style=""border:1px #8CAAE6 solid; padding-right:5px;"">"
								
								strSQL3 = "SELECT sum(timer) AS sumtimer FROM ressourcer WHERE "& jobaktidKri &" AND mid =" &intMedArbValThis(m)
								oRec3.open strSQL3, oConn, 3 
								if not oRec3.EOF then
								timertot_medarb = oRec3("sumtimer")
								end if
								oRec3.close
								if len(timertot_medarb) <> 0 then
								timertot_medarb = timertot_medarb
								else
								timertot_medarb = 0
								end if
								
								Response.write formatnumber(timertot_medarb, 2)
								balancethis_ress =  timertot_medarb - timertot  
								Response.write "<td valign=top bgcolor=#FFFFFF align=right style=""border:1px #8CAAE6 solid; padding-right:5px;""><font color='darkred'>"& formatnumber(timertot,2) &"</td>"
								Response.write "</td><td bgcolor=#FFFFFF valign=top align=right style=""border:1px #8CAAE6 solid; padding-right:5px;"">"& formatnumber(balancethis_ress, 2) &"</td></tr>"
								
								end if
								oRec.close
								%></table><br><br>
								<%
								antalson = 0
								antalman = 0
								antaltir = 0
								antalons = 0
								antaltor = 0
								antalfre = 0
								antallor = 0
							
end function







'****************************************
'*** Tildelte ressourcetimer funktion ***
function ressourcetimer(thisjobid, aktelljob)
if aktelljob = "J" then
aktjobkri = " jobid = "
bgthis = "#EFF3FF"
txt = "job"
else
aktjobkri = " aktid = "
bgthis = "#FFFFF1"
txt = "aktivitet"
end if%>
<table cellpadding=0 cellspacing=1 border=0 width=610>
					<tr>
						<td bgcolor="<%=bgthis%>" width=300 valign=bottom style="padding-top:2; padding-left:2; border:#8CAAE6 1px solid;"><b>Ressourcetimer tildelt på <%=txt%>:</b><br>
						<font color=darkred>&sum;</font> <font class=megetlillesort>Tot. og + / - balance beregnet udfra <u>de valgte medarbejdere</u>.</td>
						<td bgcolor="<%=bgthis%>" width=100 valign=bottom style="padding-top:2; padding-left:2; border:#8CAAE6 1px solid;"><font class=megetlillesort>På <u>valgte</u> medarbejd.</td>
						<td bgcolor="<%=bgthis%>" width=80 valign=bottom style="padding-top:2; padding-left:2; border:#8CAAE6 1px solid;"><font class=megetlillesort>På <u>alle</u> medarb.</td>
						<td bgcolor="<%=bgthis%>" width=80 valign=bottom align="center" style="padding-top:2; padding-left:2; border:#8CAAE6 1px solid;"><b>Timer brugt</b><br><font color=darkred>&sum;</font> <font class=megetlillesort>Periode / <font color=#999999>Total </td>
						<td bgcolor="<%=bgthis%>" width=50 valign=bottom align="center" style="padding-top:2; padding-left:2; border:#8CAAE6 1px solid;"><font class=megetlillesort> + / - </font></td>
					</tr>
						<tr>
						<td align=right style="padding-top:2; padding-right:3; border:#8CAAE6 1px solid;">I valgt periode:</td>
						<td bgcolor="#FFFFFF" align=right valign="top" style="padding-top:2; padding-right:3; border:#8CAAE6 1px solid;"><%
							strSQL3 = "SELECT sum(timer) AS sumtimer FROM ressourcer WHERE (dato >= '"&convertDateYMD(strTdato)&"' AND dato <= '"&convertDateYMD(strUdato)&"') AND "& aktjobkri &" "& thisjobid &" AND " & useResMed_kri
							
							oRec3.open strSQL3, oConn, 3 
							if not oRec3.EOF then
							timertot_medarb = oRec3("sumtimer")
							end if
							oRec3.close
							if len(timertot_medarb) <> 0 then
							timertot_medarb = timertot_medarb
							else
							timertot_medarb = 0
							end if
							
							Response.write formatnumber(timertot_medarb, 2)
							%></td>
						<td align=right bgcolor="#FFFFFF" valign="top" style="padding-top:2; padding-right:3; border:#8CAAE6 1px solid;"><%
							strSQL3 = "SELECT sum(timer) AS sumtimer FROM ressourcer WHERE (dato >= '"&convertDateYMD(strTdato)&"' AND dato <= '"&convertDateYMD(strUdato)&"') AND "& aktjobkri &" "& thisjobid &""
							oRec3.open strSQL3, oConn, 3 
							if not oRec3.EOF then
							sumtimertot = oRec3("sumtimer")
							end if
							oRec3.close
							if len(sumtimertot) <> 0 then
							sumtimertot = sumtimertot
							else
							sumtimertot = 0
							end if
							
							Response.write formatnumber(sumtimertot, 2)
							resttimer = timertot_medarb - timertot
							%></td>
							<td bgcolor="#FFFFFF" align=right valign="top" style="padding-top:2; padding-right:3; border:#8CAAE6 1px solid;"><font color=darkred><%=formatnumber(timertot, 2)%></td>
							<td bgcolor="#FFFFFF" align=right valign="top" style="padding-top:2; padding-right:3; border:#8CAAE6 1px solid;"><%=formatnumber(resttimer, 2)%></td>
					
						
					</tr>
					<td align=right style="padding-top:2; padding-right:3; border:#8CAAE6 1px solid;">Ialt tildelt:</td>
						<td bgcolor="#FFFFFF" align=right valign="top" style="padding-top:2; padding-right:3; border:#8CAAE6 1px solid;"><%
							strSQL3 = "SELECT sum(timer) AS sumtimer FROM ressourcer WHERE "& aktjobkri &" "& thisjobid &" AND " & useResMed_kri
							oRec3.open strSQL3, oConn, 3 
							if not oRec3.EOF then
							timertot_medarb = oRec3("sumtimer")
							end if
							oRec3.close
							if len(timertot_medarb) <> 0 then
							timertot_medarb = timertot_medarb
							else
							timertot_medarb = 0
							end if
							
							Response.write formatnumber(timertot_medarb, 2)
							
							%></td>
							<td bgcolor="#FFFFFF" align=right valign="top" style="padding-top:2; padding-right:3; border:#8CAAE6 1px solid;"><%
							strSQL3 = "SELECT sum(timer) AS sumtimer FROM ressourcer WHERE "& aktjobkri &" "& thisjobid &""
							oRec3.open strSQL3, oConn, 3 
							if not oRec3.EOF then
							sumtimertot = oRec3("sumtimer")
							end if
							oRec3.close
							if len(sumtimertot) <> 0 then
							sumtimertot = sumtimertot
							else
							sumtimertot = 0
							end if
							
							Response.write formatnumber(sumtimertot, 2)
							
							
							if aktelljob = "J" then
								'**** timer tot. på job ******
								strSQLT = "SELECT jobnr FROM job WHERE id="& thisjobid
								oRec.open strSQLT, oConn, 3 
								if not oRec.EOF then
								jnrthis = oRec("jobnr")
								end if
								oRec.close
							end if
							
							if aktelljob = "J" then
							strSQL3 = "SELECT sum(timer) AS sumtimertotal FROM timer WHERE Tjobnr =" & jnrthis &" AND tfaktim <> 5"
							else
							strSQL3 = "SELECT sum(timer) AS sumtimertotal FROM timer WHERE TAktivitetId =" & thisjobid &" AND tfaktim <> 5"
							end if
							
							oRec3.open strSQL3, oConn, 3 
							if not oRec3.EOF then
							jobsumtimertot = oRec3("sumtimertotal")
							end if
							oRec3.close
							if len(jobsumtimertot) <> 0 then
							jobsumtimertot = jobsumtimertot
							else
							jobsumtimertot = 0
							end if
							resttimer = timertot_medarb - jobsumtimertot
							%></td>
							<td bgcolor="#FFFFFF" align=right valign="top" style="padding-top:2; padding-right:3; border:#8CAAE6 1px solid;"><font color=#999999><%=formatnumber(jobsumtimertot, 2)%></td>
							<td bgcolor="#FFFFFF" align=right valign="top" style="padding-top:2; padding-right:3; border:#8CAAE6 1px solid;"><%=formatnumber(resttimer, 2)%></td>
					</tr>
					</table>
<%
end function
'******************************************************************************




'******* Periode funktion *******************************************************
'* Henter Periode
'* Altid del op i en periode på 4 intervaler.
'* Daysinperiod <= 31 = 1 mrd 
'* Daysinperiod > 31 < ?
'* Regn totaler sammen for den specifikke periode 
'********************************************************************************
public antalson, antalman, antaltir, antalons, antaltor, antalfre, antallor 
public function periode(interval, intervaluger, startDato, slutDato, countdays)
%>
	<%
		'** Viser dag for dag / uge for uge
		if 0 <= interval AND interval <= 31 then 'AND useTotal <> "j" then%>
		<table border=0 cellspacing=0 cellpadding=0>
			<tr><%
				for dag = 0 to interval 
				
				dato = datevalue(startDato) + cint(dag) 
				usedato = cdate(dato)
				ugedag = Weekday(usedato)
				ugedagnavn = WeekdayName(ugedag, 1)
				
				'*** normeret tid **
				if countdays = "j" then
					Select case ugedag
					case 1
					antalson = antalson + 1
					case 2
					antalman = antalman + 1
					case 3
					antaltir = antaltir + 1
					case 4
					antalons = antalons + 1
					case 5
					antaltor = antaltor + 1
					case 6
					antalfre = antalfre + 1
					case 7
					antallor = antallor + 1
					end select
				end if 
				
				if ugedag = 1 and dag <> 0 then%>
				<td bgcolor="<%=bgcolhvbl%>" class='<%=tdcls%>' valign="bottom" style="!border: 0px; border-color: #8CAAE6; padding-left:2; padding-right:4;"><font color="darkred">&sum;</font></td>
				<%end if
				
				'*** ombrydning af tabel ***
				select case dag
				case 24
				%>
				</tr><tr style="position:relative; top:28; left:0; z-index:1000;" height="20">
				<%
				end select
				'******************
				
				if ugedag = 1 OR dag = 0 then%>
					<td class='<%=tdcls%>' bgcolor="#EFF3FF" valign="top" width="30" align="center" style="!border: 1px; border-color: #8CAAE6;  border-style: solid;  border-bottom:0; padding-left:2; padding-right:2;"><b>uge 
					<%=datepart("ww", usedato,2,2)%><br>
					<%=year(usedato)%>
					</b></td>
				<%end if%>
					<td class='<%=tdcls%>' bgcolor="#EFF3FF" align=center width="20" style="!border: 1px; border-color: #8CAAE6; border-style: solid; border-bottom:0; border-left:0;"><img src="../ill/blank.gif" width="20" height="1" alt="" border="0"><br><%=ugedagnavn%><br>
					<%
					antalch = instr(formatdatetime(usedato, 1),".")
					usedatotrim = left(formatdatetime(usedato, 1), (antalch + 4))
					Response.write usedatotrim%></td>
				<%next%>
				<td valign="bottom" bgcolor="<%=bgcolhvbl%>" class='<%=tdcls%>' width=40 style="padding-left:2;"><img src="../ill/blank.gif" width="40" height="1" alt="" border="0"><br><font color="darkred">&sum;</font>&nbsp;<b>Tot.</b></td>
				<td valign="bottom" bgcolor="<%=bgcolhvbl%>" class='<%=tdcls%>' width=60 style="padding-left:2;"><img src="../ill/blank.gif" width="60" height="1" alt="" border="0"><br><b>Omsætning:</b></td>
			</tr>
		<%else
		'** Viser uge for uge / mrd for mrd%>
			<table border=0 cellspacing=0 cellpadding=0>
			<tr><%
				for uger = 0 to intervaluger 
				dato = datevalue(startDato) + cint(7*uger) 
				usedato = cdate(dato)
				uge = datepart("ww", usedato,2,2)
				
				
				'*** normeret tid **
				if countdays = "j" then
				
						antalson = antalson + 1
						antalman = antalman + 1
						antaltir = antaltir + 1
						antalons = antalons + 1
						antaltor = antaltor + 1
						antalfre = antalfre + 1
						antallor = antallor + 1
				
				
					if uger = 0 then
						select case weekday(startDato)
						case 1
						
						case 2
						antalson = antalson - 1
						case 3
						antalson = antalson - 1
						antalman = antalman - 1
						case 4
						antalson = antalson - 1
						antalman = antalman - 1
						antaltir = antaltir - 1
						case 5
						antalson = antalson - 1
						antalman = antalman - 1
						antaltir = antaltir - 1
						antalons = antalons - 1
						case 6
						antalson = antalson - 1
						antalman = antalman - 1
						antaltir = antaltir - 1
						antalons = antalons - 1
						antaltor = antaltor - 1
						case 7
						antalson = antalson - 1
						antalman = antalman - 1
						antaltir = antaltir - 1
						antalons = antalons - 1
						antaltor = antaltor - 1
						antalfre = antalfre - 1
						end select 
					end if
					
					if uger = intervaluger then
					
						select case weekday(slutDato)
						case 7
						
						case 6
						antallor = antallor - 1
						case 5
						antalfre = antalfre - 1
						antallor = antallor - 1
						case 4
						antaltor = antaltor - 1
						antalfre = antalfre - 1
						antallor = antallor - 1
						case 3
						antalons = antalons - 1
						antaltor = antaltor - 1
						antalfre = antalfre - 1
						antallor = antallor - 1
						case 2
						antaltir = antaltir - 1
						antalons = antalons - 1
						antaltor = antaltor - 1
						antalfre = antalfre - 1
						antallor = antallor - 1
						case 1
						antalman = antalman - 1
						antaltir = antaltir - 1
						antalons = antalons - 1
						antaltor = antaltor - 1
						antalfre = antalfre - 1
						antallor = antallor - 1
						end select 
					end if
					
					
				end if
				
				if month(usedato) <> lastmonth AND uger <> 0 then%>
				<td bgcolor="<%=bgcolhvbl%>" class='<%=tdcls%>' valign="bottom" style="!border: 1px; border-color: #8CAAE6; border-style: solid; border-left:0; border-top:0; border-right:0; padding-left:2; padding-right:4;"><font color="darkred">&sum;</font></td>
				<%end if
					'*** ombrydning af tabel ***
					if (uger = 0 OR uger = 18 OR uger = 36 OR uger = 54 OR uger = 72 OR uger = 90 OR uger = 108 OR uger = 126 OR uger = 144) AND periodeinterval > 180 then
							
							select case uger 
							case 0
							toppx = "0"
							case 18
							toppx = "24"
							case 36
							toppx = "48"
							case 54
							toppx = "72"
							case 72
							toppx = "96"
							case 90
							toppx = "120"
							case 108
							toppx = "144"
							case 126
							toppx = "168"
							case 144
							toppx = "192"
							end select%>
					</tr><tr style="position:relative; top:<%=toppx%>; left:0; z-index:1000;" height="20">
					<%end if
					
				if month(usedato) <> lastmonth OR uger = 0 then%>
					<td class='<%=tdcls%>' bgcolor="#EFF3FF" valign="top" width="30" align="center" style="!border: 1px; border-color: #8CAAE6; border-style: solid; padding-left:2; padding-right:2;"><b>
					<%=left(monthname(datepart("m", usedato)), 3)%><br>
					<%=year(usedato)%>
					</b></td>
				<%end if%>
					<td class='<%=tdcls%>' bgcolor="#EFF3FF" align=center width="20" style="!border: 1px; border-color: #8CAAE6; border-style: solid; border-left:0;">
					<img src="../ill/blank.gif" width="20" height="1" alt="" border="0"><br>
					<%=uge%></td>
				<%
				lastmonth = month(usedato)
				lastyear = year(usedato)
				next%>
				<td valign="bottom" bgcolor="<%=bgcolhvbl%>" class='<%=tdcls%>' width=40 style="!border: 1px; border-color: #8CAAE6; border-style: solid; border-left:0; border-top:0; border-right:0; padding-left:2;"><img src="../ill/blank.gif" width="40" height="1" alt="" border="0"><br><font color="darkred">&sum;</font>&nbsp;<b>Tot.</b></td>
				<td valign="bottom" bgcolor="<%=bgcolhvbl%>" class='<%=tdcls%>' width=60 style="!border: 1px; border-color: #8CAAE6; border-style: solid; border-left:0; border-top:0; border-right:0; padding-left:2;"><img src="../ill/blank.gif" width="60" height="1" alt="" border="0"><br><b>Omsætning:</b></td>
			</tr>
		<%end if
	
end function


'*** SQLBless **********************
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
		end function

'** Function til at fjerne decimaler
	Public left_intX
	Public function kommaFunc(x)
	if len(x) <> 0 then
	instr_komma = SQLBless(instr(x, "."))
		
		if instr_komma > 0 then
		left_intX = left(x, instr_komma + 2)
		else
		left_intX = x
		end if
	else
	left_intX = 0
	end if
	
	Response.write left_intX 
	end function


'******* Time funktion ******************************************************************
'* 		 Finder timer i den valgte periode på de(t) valgte job.							*
'* 		 Tabellen fra periode funktioen bliver afsluttet i denne funktion.				*
'****************************************************************************************
public timertot_dag()
public timertot_week()
Public timertot
Function findtimer(hvilken, hvmedarb, hvaktid)

			
			timertot = 0
			totOms = 0
			timertot_week_this = 0
			
			'*** <= 31 = 1 måned, dag for dag *****
			if periodeinterval >= 0 AND periodeinterval <= 31 then%>
			<tr><%
			Redim timertot_dag(0)
			
			For SQLdag = 0 to periodeinterval
				
				Redim preserve timertot_dag(SQLdag)
				
				SQLdato = DateAdd("d", SQLdag, StrTdato)
				SQLenddato = datevalue(StrUdato)  
				SQLdatodate = DateValue(SQLdato)
				SQLusedato = ConvertDateYMD(SQLdatodate)
				SQLugedag = Weekday(SQLusedato)
				SQLugedagnavn = WeekdayName(SQLugedag, 1)
				
				if SQLugedag = 1 AND SQLdag <> 0 then%>
					<td class='<%=tdcls%>' align="center" bgcolor="#EFF3FF" style="!border: 1px; border-color: #8CAAE6; border-style: solid; border-left:0;  border-right:0;">&nbsp;<%
					if SQLdag <> 0 then
					Response.write "<br><font color=darkred>"&timertot_week_this&"</font>"
					'timertot_prmedarb = timertot_prmedarb + timertot_week_this
					timertot_week_this = 0
					end if
					%>&nbsp;</td>
				<%end if
				
				'*** ombrydning af tabel ***
				select case SQLdag 
				case 0 
					if periodeinterval > 23 then%>
					</tr><tr style="position:relative; top:-36; left:0; z-index:1000;" height="20"><!-- -36 -->
					<%else%>
					</tr><tr style="position:relative; top:0; left:0; z-index:1000;" height="20">
					<%end if
				case 24
				%>
				</tr><tr style="position:relative; top:3; left:0; z-index:1000;" height="20">
				<%end select
				'******************
				
				if SQLugedag = 1 OR SQLdag = 0 then%>
					<td bgcolor="<%=bgcolhvbl%>" class='<%=tdcls%>' align="center" style="!border: 1px; border-color: #8CAAE6; border-style: solid; border-bottom:0;">&nbsp;</td>
				<%end if%>
				
					<td class='<%=tdcls%>' align=center width="20" style="!border: 1px; background-color:#ffffff; border-color: #8CAAE6; border-style: solid; border-left:0;"><img src="../ill/blank.gif" width="20" height="1" alt="" border="0"><br><%=ugedagnavn%><br>
					<%
					select case hvilken
					case "jobtotal"
					SQLthisA = "SELECT timer, tfaktim, fastpris, timepris FROM timer WHERE "& hvmedarb &" AND tjobnr = "& useJobnr &" AND tfaktim <> 5 AND tdato = '"& SQLusedato &"'"
					case "medarbjobtotal"
					SQLthisA = "SELECT timer, tfaktim, fastpris, timepris FROM timer WHERE tjobnr = "& useJobnr &" AND tfaktim <> 5  AND tmnr = "& hvmedarb &" AND tdato = '"& SQLusedato &"'"
					case "allealleudspec"
					'hvaktid = jobnr
					SQLthisA = "SELECT timer, tfaktim, timer.fastpris, timepris, job.jobtpris, job.budgettimer FROM timer LEFT JOIN job ON(jobnr = tjobnr) WHERE "& hvmedarb &" AND tjobnr = "& hvaktid &" AND tfaktim <> 5 AND tdato = '"& SQLusedato &"'"
					case "medarbjobtotalAlle", "allealle"
						medArbAllejobSQL = " ("
						intJobnrAlle = Split(intJobnr, ", ")
						For i = 0 to ubound(intJobnrAlle)
						medArbAllejobSQL = medArbAllejobSQL & "tjobnr = " & intJobnrAlle(i) &" OR "
						next
						LEN_medArbAllejobSQL = len(medArbAllejobSQL)
						useMedArbAllejobSQL = left(medArbAllejobSQL, (LEN_medArbAllejobSQL - 3))
						useMedArbAllejobSQL = useMedArbAllejobSQL &") "
						if hvilken = "medarbjobtotalAlle" then
						SQLthisA = "SELECT timer, tfaktim, timer.fastpris, timepris, job.jobtpris, job.budgettimer FROM timer LEFT JOIN job ON(jobnr = tjobnr) WHERE "& useMedArbAllejobSQL &" AND tmnr = "& hvmedarb &" AND tfaktim <> 5 AND tdato = '"& SQLusedato &"'"
						else
						SQLthisA = "SELECT timer, tfaktim, timer.fastpris, timepris, job.jobtpris, job.budgettimer FROM timer LEFT JOIN job ON(jobnr = tjobnr) WHERE "& useMedArbAllejobSQL &" AND " & hvmedarb &" AND tfaktim <> 5 AND tdato = '"& SQLusedato &"'"
						end if
					case "akttotal"
					SQLthisA = "SELECT timer, tfaktim, fastpris, timepris FROM timer WHERE "& hvmedarb &" AND taktivitetid = "& hvaktid &" AND tdato = '"& SQLusedato &"'"
					case "medprakt"
					SQLthisA = "SELECT timer, tfaktim, fastpris, timepris FROM timer WHERE taktivitetid = "& hvaktid &" AND tmnr = "& hvmedarb &" AND tfaktim <> 5 AND tdato = '"& SQLusedato &"'"
					end select
					
					strSQL = SQLthisA 
					oRec.open strSQL, oConn, 0, 1
					while not oRec.EOF 
					timertot = timertot + oRec("timer")
					timertot_dag(SQLdag) = timertot_dag(SQLdag) + oRec("timer")
					
						if cstr(oRec("tfaktim")) = "1" then
							if cstr(oRec("fastpris")) = "1" then
							
								if hvilken = "medarbjobtotalAlle" OR hvilken = "allealle" OR hvilken = "allealleudspec" then
									if oRec("budgettimer") > 0 then
									thistp = (oRec("jobtpris")/oRec("budgettimer"))
									else
									thistp = 0
									end if
								else
									if cint(useBudgettimer) > 0 then
									thistp = (int(useTpris)/int(useBudgettimer))
									else
									thistp = 0
									end if
								end if
								totOms = totOms + (thistp * oRec("timer"))									
							else 
								totOms = totOms + (oRec("timer") * oRec("timepris"))
							end if
						else
							if totOms > 0 then 
							totOms = totOms
							else
							totOms = 0
							end if
						end if
					
					oRec.movenext
					wend
					oRec.close
						
					Response.write timertot_dag(SQLdag)
					timertot_week_this = timertot_week_this + timertot_dag(SQLdag)
					%>
					</td>
				<%next%>
				<td class='<%=tdcls%>' valign=bottom width=40 style="!border: 1px; background-color:#ffffff; border-color: #8CAAE6; border-style: solid; border-left:0; padding-left:2;"><img src="../ill/blank.gif" width="40" height="1" alt="" border="0"><br><font color=darkred><%=kommaFunc(timertot)%>&nbsp;&nbsp;</td>
				<td class='<%=tdcls%>' valign=bottom width=60 style="!border: 1px; background-color:#ffffff; border-color: #8CAAE6; border-style: solid; border-left:0; padding-left:2;"><img src="../ill/blank.gif" width="60" height="1" alt="" border="0"><br>
				<% if int(totOms) > 0 then
				Response.write formatcurrency(totOms,2)
				else
				response.write totOms
				end if%></td>
			</tr>
			<%
			else%>
			<tr><%
			'*** > 31 = uge for uge *****
			Redim timertot_week(0)
			
			'Redim timertot_month(0)
			For SQLuger = 0 to periodeintervalWeeks
				
				Redim preserve timertot_week(SQLuger)
				'Redim preserve timertot_month(SQLuger)
				SQLdato = DateAdd("d", (SQLuger*7), StrTdato)
				SQLdatodate = DateValue(SQLdato)
				SQLusedato = ConvertDateYMD(SQLdatodate)
				SQLugedag = Weekday(SQLusedato)
				SQLugedagnavn = WeekdayName(SQLugedag, 1)
				
				if month(SQLusedato) <> lastmonth AND SQLuger <> 0 then%>
					<td class='<%=tdcls%>' bgcolor="#EFF3FF" align="center" style="!border: 1px; border-color: #8CAAE6; border-style: solid; border-left:0; border-right:0; border-top:0;">&nbsp;<%
					if SQLuger <> 0 then
					Response.write "<br><font color=darkred>"& timertot_month &"</font>"
					timertot_month = 0
					end if%>&nbsp;</td>
				<%end if
				
				'*** ombrydning *****
				if(SQLuger = 0 OR SQLuger = 18 OR SQLuger = 36 OR SQLuger = 54 OR SQLuger = 72 OR SQLuger = 90 OR SQLuger = 108 OR SQLuger = 126 OR SQLuger = 144) AND periodeinterval > 180 then
					
							if periodeinterval > 180 AND periodeinterval < 246 then '335 then 242
								select case SQLuger 
								case 0
								toppxS = "-26" 
								case 18
								toppxS = "-2"
								end select
							end if
							if periodeinterval => 246 AND periodeinterval < 375 then '242
								select case SQLuger 
								case 0
								toppxS = "-50" 
								case 18
								toppxS = "-26"
								case 36
								toppxS = "-2"
								end select
							end if
							if periodeinterval => 375 AND periodeinterval < 501 then
								select case SQLuger 
								case 0
								toppxS = "-74" 
								case 18
								toppxS = "-50"
								case 36
								toppxS = "-26"
								case 54
								toppxS = "-2"
								end select
							end if
							if periodeinterval => 501 AND periodeinterval < 627 then
								select case SQLuger 
								case 0
								toppxS = "-98" 
								case 18
								toppxS = "-74"
								case 36
								toppxS = "-50"
								case 54
								toppxS = "-26"
								case 72
								toppxS = "-2"
								end select
							end if
							if periodeinterval => 627 AND periodeinterval < 756 then
								select case SQLuger 
								case 0
								toppxS = "-122" 
								case 18
								toppxS = "-98"
								case 36
								toppxS = "-74"
								case 54
								toppxS = "-50"
								case 72
								toppxS = "-26"
								case 90
								toppxS = "-2"
								end select
							end if
							if periodeinterval => 756 AND periodeinterval < 855 then 'Ca. 2½ år. Dette er pt den maksimale tidsperiode der kan vises.
								select case SQLuger 
								case 0
								toppxS = "-146" 
								case 18
								toppxS = "-122"
								case 36
								toppxS = "-98"
								case 54
								toppxS = "-74"
								case 72
								toppxS = "-50"
								case 90
								toppxS = "-26"
								case 108
								toppxS = "-2"
								case 126
								toppxS = "22"
								case 144
								toppxS = "46"
								end select
							end if
						
					%>
					</tr><tr style="position:relative; top:<%=toppxS%>; left:0;" height="20">
				<%end if
				
				if month(SQLusedato) <> lastmonth OR SQLuger = 0 then%>
					<td bgcolor="<%=bgcolhvbl%>" class='<%=tdcls%>' align="center" style="!border: 1px; border-color: #8CAAE6; border-style: solid; border-top:0; border-bottom:0;">&nbsp;</td>
				<%end if%>
					<td class='<%=tdcls%>' align=center width="20" valign="top" style="!border: 1px; background-color:#ffffff; border-color: #8CAAE6; border-style: solid; border-left:0; border-top:0; padding-left:2;"><img src="../ill/blank.gif" width="20" height="1" alt="" border="0"><br><%=ugedagnavn%><br>
					<%
					'** første og sidste dag i uge ***
					forsteDagiUge = DateAdd("d", -(SQLugedag-1), SQLdatodate) 
					forsteDagiUge = ConvertDateYMD(forsteDagiUge)
					sidsteD = cint(7-SQLugedag)
					SidsteDagiUge = DateAdd("d", (sidsteD), SQLdatodate) 
					SidsteDagiUge = ConvertDateYMD(SidsteDagiUge)
					
					'*********************************
					select case hvilken
					case "jobtotal"
					SQLthisB = "SELECT timer AS ztimer, tfaktim, fastpris, timepris FROM timer WHERE "& hvmedarb &" AND tjobnr = "& useJobnr &" AND tfaktim <> 5 AND tdato BETWEEN '"& forsteDagiUge &"' AND '"& SidsteDagiUge &"'"
					case "medarbjobtotal"
					SQLthisB = "SELECT timer AS ztimer, tfaktim, fastpris, timepris FROM timer WHERE tjobnr = "& useJobnr &" AND tmnr = "& hvmedarb &" AND tfaktim <> 5 AND tdato BETWEEN '"& forsteDagiUge &"' AND '"& SidsteDagiUge &"'"
					case "allealleudspec"
					'hvaktid = jobnr
					SQLthisB = "SELECT timer AS ztimer, tfaktim, timer.fastpris, timepris, job.jobtpris, job.budgettimer FROM timer LEFT JOIN job ON(jobnr = tjobnr) WHERE "& hvmedarb &" AND tjobnr = "& hvaktid &" AND tfaktim <> 5 AND tdato BETWEEN '"& forsteDagiUge &"' AND '"& SidsteDagiUge &"'"
					case "medarbjobtotalAlle", "allealle"
						medArbAllejobSQL = " ("
						intJobnrAlle = Split(intJobnr, ", ")
						For i = 0 to ubound(intJobnrAlle)
						medArbAllejobSQL = medArbAllejobSQL & "tjobnr = " & intJobnrAlle(i) &" OR "
						next
						LEN_medArbAllejobSQL = len(medArbAllejobSQL)
						useMedArbAllejobSQL = left(medArbAllejobSQL, (LEN_medArbAllejobSQL - 3))
						useMedArbAllejobSQL = useMedArbAllejobSQL &") "
						
						if hvilken = "medarbjobtotalAlle" then
						SQLthisB = "SELECT timer AS ztimer, tfaktim, timer.fastpris, timepris, job.jobtpris, job.budgettimer FROM timer LEFT JOIN job ON(jobnr = tjobnr) WHERE "& useMedArbAllejobSQL &" AND tmnr = "& hvmedarb &" AND tfaktim <> 5 AND tdato BETWEEN '"& forsteDagiUge &"' AND '"& SidsteDagiUge &"' ORDER BY tdato"
						else
						SQLthisB = "SELECT timer AS ztimer, tfaktim, timer.fastpris, timepris, job.jobtpris, job.budgettimer FROM timer LEFT JOIN job ON(jobnr = tjobnr) WHERE "& useMedArbAllejobSQL &" AND " & hvmedarb &" AND tfaktim <> 5 AND tdato BETWEEN '"& forsteDagiUge &"' AND '"& SidsteDagiUge &"' ORDER BY tdato"
						end if
					
					case "akttotal"
					SQLthisB = "SELECT timer AS ztimer, tfaktim, fastpris, timepris FROM timer WHERE "& hvmedarb &" AND taktivitetid = "& hvaktid &" AND tdato BETWEEN '"& forsteDagiUge &"' AND '"& SidsteDagiUge &"'"
					case "medprakt"
					SQLthisB = "SELECT timer AS ztimer, tfaktim, fastpris, timepris FROM timer WHERE taktivitetid = "& hvaktid &" AND tmnr = "& hvmedarb &" AND tfaktim <> 5 AND tdato BETWEEN '"& forsteDagiUge &"' AND '"& SidsteDagiUge &"'"
					end select
					
					strSQL = SQLThisB
					'Response.write strSQL 
					oRec.open strSQL, oConn, 0, 1
					
					while not oRec.EOF
					timertot = timertot + oRec("ztimer")
					timertot_week(SQLuger) = timertot_week(SQLuger) + oRec("ztimer")
				 	
						if oRec("tfaktim") = "1" then
							if oRec("fastpris") = "1" then
								if hvilken = "medarbjobtotalAlle" OR hvilken = "allealle" OR hvilken = "allealleudspec" then
									if oRec("budgettimer") > 0 then
									thistp = (oRec("jobtpris")/oRec("budgettimer"))
									else
									thistp = 0
									end if
								else
									if cint(useBudgettimer) > 0 then
									thistp = (int(useTpris)/int(useBudgettimer))
									else
									thistp = 0
									end if
								end if
								totOms = totOms + (thistp * oRec("ztimer"))
							else 
							totOms = totOms + (oRec("ztimer") * oRec("timepris"))
							'Response.write oRec("ztimer") &" # " & oRec("timepris")
							end if
						else
							if totOms > 0 then 
							totOms = totOms
							else
							totOms = 0
							end if
						end if
					
					
					oRec.movenext
					wend
					oRec.close
					
					Response.write timertot_week(SQLuger)
					timertot_month = timertot_month + timertot_week(SQLuger)
					
					%></td>
				<%
				lastmonth = month(SQLusedato) 
			next%>
			<td class='<%=tdcls%>' valign=bottom width=40 style="!border: 1px; background-color:#ffffff; border-color: #8CAAE6; border-style: solid; border-left:0; border-top:0; padding-left:2;"><img src="../ill/blank.gif" width="40" height="1" alt="" border="0"><br><font color=darkred><%=kommaFunc(timertot)%>&nbsp;&nbsp;</td>
			<td class='<%=tdcls%>' valign=bottom width=60 style="!border: 1px; background-color:#ffffff; border-color: #8CAAE6; border-style: solid; border-left:0; border-top:0; padding-left:2;"><img src="../ill/blank.gif" width="60" height="1" alt="" border="0"><br>
			<% if int(totOms) > 0 then
			Response.write formatcurrency(totOms,2)
			else
			response.write totOms
			end if
			%></td>
		</tr>
		<%end if%>
		
		<%
end function



'***************************************************
'* 			Timefordelings grafik					
'***************************************************
sub timefordelingsgrafik

%>
<tr>
	<td>
	<table border=0 cellspacing=0 cellpadding=0 style="position:absolute; left:0; top:-72;">
	<tr>
	<td valign="bottom"><b>Timefordeling:</b><br>
	<%
	showIntervalStartDato = StrTdato
	showIntervalSlutDato = StrUdato
	if periodeinterval <= 31 then
	Response.write formatdatetime(showIntervalStartDato, 1) 
	else
	Response.write "Uge:&nbsp;"& datepart("ww", showIntervalStartDato,2,2) &" - 20"& strAar 
	end if
	%>&nbsp;til&nbsp;
	<%
	if periodeinterval <= 31 then
	Response.write formatdatetime(showIntervalSlutDato, 1)
	else
	Response.write datepart("ww", showIntervalSlutDato,2,2) &" - 20"& strAar_slut
	end if
	%></td>
	</tr>
	</table>
	<table border=0 cellspacing=0 cellpadding=0 style="position:absolute; left:<%=lefttfg%>; top:-42;">
	<tr>
	<%
	if periodeinterval <= 31 then
		counterint = periodeinterval
	else
		counterint = periodeintervalWeeks
	end if
	
	For x = 0 to counterint%>
	<td height="50" bgcolor="#FFFFFF" width="4" valign="bottom" align="center" style="!border: 1px; border-color: Darkred; border-style: solid; border-bottom:0; border-top:0; border-left:0;">
	<%if periodeinterval <= 31 then
		if cint(timertot_dag(x)) < 50 then
		heightthis = timertot_dag(x)
		else
		heightthis = 50
		end if
	else
		if cint(timertot_week(x)) < 50 then
		heightthis = timertot_week(x)
		else
		heightthis = 50
		end if
	end if%>
	<img src="../ill/zscala.gif" width="2" height="<%=heightthis%>" alt="" border="0"></td>
	<%next%>
	
	</tr>
	</table>
	</td>
</tr>
<%end sub

'***************************     Funktioner slut      *************************************




'*************************************************************************************************
'*    Følgende udskrives til skærm  																 *
'*************************************************************************************************
if request("print") <> "j" then
pleft = 190
ptop = 55
tdcls = "lille-print"
bgcolhvbl = "#D6DFF5"
else
tdcls = "lille-print"
pleft = 20
ptop = 0
bgcolhvbl = "#FFFFFF"
end if

if request("print") <> "j" then%>	
<!--include file="inc/stat_submenu.asp"-->
<div style="position:absolute; left:800; top:110;">
<a href="javascript:NewWin_large('joblog_z_b.asp?menu=stat&print=j&FM_seljob=<%=usepseljob%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>&FM_vis_medarb=<%=vis_medarb%>&FM_vis_akt=<%=vis_akt%>&FM_vis_medarb_k=<%=vis_medarb_k%>')" target="_self" class='rmenu'>&nbsp;Print venlig version&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
</div>
<%else%>
<table cellspacing="0" cellpadding="0" border="0" width="880">
	<tr>
		<td bgcolor="#003399" width="650"><img src="../ill/logo_topbar_print.gif" alt="" border="0"></td>
		<td bgcolor="#FFFFFF" align=right><a href="javascript:window.close()"><img src="../ill/luk_xp.gif" width="30" height="28" alt="" border="0">&nbsp;Luk</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"><a href="javascript:window.print()"><img src="../ill/print_xp.gif" width="28" height="30" alt="" border="0">&nbsp;Print</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"></td>
	</tr>
	</table>
<%end if%>
<div id="sideindhold" style="position:absolute; left:<%=pleft%>; top:<%=ptop%>; visibility:visible; z-index:100;">
<%
if request("print") <> "j" then
grafstartleft = 0
grafstarttop = 80 
tablestartleft = 0 '-180
else
grafstartleft = 0
grafstarttop = -100
tablestartleft = 0
end if

if request("print") <> "j" then%>
<img src="../ill/header_joblog_z.gif" width="758" height="45" alt="" border="0">
<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="600" style="position:absolute; left:<%=grafstartleft%>; top:<%=grafstarttop-27%>;">
<tr bgcolor="#5582D2">
	<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
	<td colspan=4 valign="top"><img src="../ill/tabel_top.gif" width="584" height="1" alt="" border="0"></td>
	<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
</tr>
<tr bgcolor="#5582D2">
	<td colspan=4 valign="top" class="alt">&nbsp;<b>Job og periode:</b></td>
</tr>	
<tr><form action="joblog_z_b.asp?menu=stat" method="post">
	<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	<td><input type="hidden" name="FM_medarb" value="<%=FM_medarb%>">
	<input type="hidden" name="FM_job" value="<%=FM_job%>">
	<select name="FM_seljob" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
	<%if showallealle = "j" then
	selAlleAlle = "SELECTED"
	else
	selAlleAlle = ""
	end if
	%>
	<option value="0" <%=selAlleAlle%>>-- Alle valgte job --</option>
	<%=ddjobnr%>
	</select></td>
	<!--#include file="inc/weekselector_b.asp"-->
	<td>&nbsp;&nbsp;<!--Total:&nbsp;<input type="checkbox" name="FM_vis_total" value="j" <=tot_ch%>>--></td>
	<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
</tr>
<tr>
<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="70" alt="" border="0"></td>
<td colspan="2"><br><b>Vis timefordeling fordelt på medarbejdere:</b>&nbsp;
<%if vis_medarb = "j" then
medarb_ch = "CHECKED"
else
medarb_ch = ""
end if

if vis_akt = "j" then
akt_ch = "CHECKED"
else
akt_ch = ""
end if

if vis_medarb_k = "j" then
medarb_k_ch = "CHECKED"
else
medarb_k_ch = ""
end if

%>
&nbsp;<input type="checkbox" name="FM_vis_medarb" value="j" <%=medarb_ch%>><br>
Vælg denne for at få vist den normerede tid / fleks tid for de valgte<br> medarbejdere på de valgte job.

</td><td colspan="2" valign="bottom">&nbsp;<input type="image" src="../ill/statpil.gif"><br></td>
<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="70" alt="" border="0"></td>
</tr>
<tr>
	<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="48" alt="" border="0"></td>
	<td colspan="4"><b>nb)</b>
	Hvis der vælges en periode på <b>mere end 31 dage</b> vises uge for uge. Ellers vises dag for dag.
	<br>
	I <b>Uge For Uge</b> oversigten kan timerne bilve vist som indtastet i den forrige måned hvis en uge deler sig over 2 måneder.</td>
	<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="48" alt="" border="0"></td>
</tr>
</form>
<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan=4 valign="bottom"><img src="../ill/tabel_top.gif" width="584" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
</tr>
</table>
<%end if%>

<table border=0 cellspacing=0 cellpadding=0 style="position:absolute; left:<%=tablestartleft%>; top:<%=grafstarttop+240%>;">
<%
'************************************************************************************************
'* Henter job totaler for de valgte job og medarbejdere.										*
'************************************************************************************************
ressource_usesel_medarb = " ("
sqlallemedarb = " ("
intMedArbValThis = Split(selmedarb, ", ")
For m = 0 to Ubound(intMedArbValThis)
	Redim preserve medarbnavn(m)
	strSQL = "SELECT mnavn FROM medarbejdere WHERE mid = "& intMedArbValThis(m) 
	oRec.open strSQL, oConn, 0, 1 'adOpenForwardOnly, adLockReadOnly
	if not oRec.EOF then
	medarbnavn(m) = oRec("mnavn")
	sqlallemedarb = sqlallemedarb & " tmnr = "&	 intMedArbValThis(m) &" OR "
	ressource_usesel_medarb = ressource_usesel_medarb & " mid = "& intMedArbValThis(m) &" OR "
	end if
	oRec.close
next

'*** trimer SQL statements ***
LEN_sqlallemedarb = len(sqlallemedarb)
usesqlallemedarb = left(sqlallemedarb, (LEN_sqlallemedarb - 3))
usesqlallemedarb = usesqlallemedarb &") "

LEN_ressource_usesel_medarb = len(ressource_usesel_medarb)
useressource_usesel_medarb= left(ressource_usesel_medarb, (LEN_ressource_usesel_medarb - 3))
useressource_usesel_medarb = useressource_usesel_medarb &") "
useResMed_kri = useressource_usesel_medarb

'***** Udvalgt job ****
if showallealle = "n" AND (periodeinterval >= 0 AND periodeinterval < 855) then%>
<tr>
	<td width="855">
	<br><br><img src="../ill/job.gif" width="35" height="33" alt="" border="0">&nbsp;&nbsp;<b>Det valgte job:</b></td>
</tr>
<tr>
	<td style="padding-bottom:2;">
	<table cellpadding=0 cellspacing=1 border=0 width="500">
	<tr>
		<td colspan="2" bgcolor="#EFF3FF" style="padding-top:2; padding-left:5; border:#8CAAE6 1px solid;"><%=fakbargif%>
		&nbsp;
		<%if request("print") <> "j" then%>
		<a href="jobs.asp?menu=job&func=red&id=<%=useJid%>&int=1"><%=useJobnr%>&nbsp;&nbsp;<%=useJobnavn%></a>&nbsp;&nbsp;(<%=useKundenavn%>)
		<%else%>
		<a href="#"><%=useJobnr%>&nbsp;&nbsp;<%=useJobnavn%></a>&nbsp;&nbsp;(<%=useKundenavn%>)
		<%end if%><br>
		
		<%
						'*** Jobansvarlige ***
						'if jobans1 <> 0 then
						
						strSQL2 = "SELECT mnavn, mnr, mid, job.id, jobans1 FROM job, medarbejdere WHERE job.id = "& useJid &" AND mid = jobans1"
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						jobans1txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
						end if
						oRec2.close
						
						
						
						strSQL2 = "SELECT mnavn, mnr, mid, jobans2 FROM job, medarbejdere WHERE job.id = "& useJid &" AND mid = jobans2"
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						jobans2txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
						end if
						oRec2.close
						
						'end if
		
		%>
		<font class=megetlillesilver>
		Jobansvarlig 1: <%=jobans1txt%><br>
	   	Jobansvarlig 2: <%=jobans2txt%></font>
		</td>
	</tr>
	<tr>
		<td bgcolor="#EFF3FF" style="padding-top:2; padding-right:5; border:#8CAAE6 1px solid;" align=right>Periode:<br>Type:</td>
		<td bgcolor="#FFFFFF" style="padding-top:2; padding-left:5; border:#8CAAE6 1px solid;"><b><%=formatdatetime(useStdato, 1) &"</b>&nbsp;&nbsp;til <b>"&formatdatetime(useSldato, 1)%></b><br>
		<%
			if useFakbart = 1 then
				if useFastpris = 1 then
				Response.write "Faspris <b>" & formatcurrency(jTpris, 2) &"</b>"
				else
				Response.write "Budget (Ikke fastpris.) <b>" & formatcurrency(jTpris, 2) &"</b>"
				end if
			else
			Response.write "(Internt)"
			end if%>
		</td>
	</tr>
	<tr>
		<td bgcolor="#EFF3FF" style="padding-top:2; padding-right:5; border:#8CAAE6 1px solid;" align=right>Forkalkuleret timer:</td>
		<td valign=top bgcolor="#FFFFFF" style="padding-top:2; padding-left:5; border:#8CAAE6 1px solid;">Fakturerbare: <b><%=useBudgettimer%></b>&nbsp;&nbsp;&nbsp;
		Ikke fakturerbare: <b><%=useIkkebudgettimer%></b></td></tr>
	</table></td>
</tr>
<tr>
	<td valign="top">
			<%
			call periode(periodeinterval, periodeintervalWeeks, StrTdato, StrUdato, "n")
			call findtimer("jobtotal", usesqlallemedarb, 0)
			%>
			</table>
	</td>
</tr>
<tr>
	<td valign="top" style="padding-top:5px;">
	<%call ressourcetimer(useJid, "J")%>
</td></tr>
<%
call timefordelingsgrafik

else
if (periodeinterval >= 0 AND periodeinterval < 855) then%>
<tr>
	<td width="855" style="padding-bottom:2;"><br><br><img src="../ill/job.gif" width="35" height="33" alt="" border="0"><div style="background-color:#FFFFFF; width:400; height:50; border:#5582d2 1px solid; padding-left:5; padding-top:2;">
	&nbsp;&nbsp;<b>Alle</b> job i denne kørsel/ <b>Alle </b>de valgte medarbejdere.</div></td>
</tr>
<%
'********* Alle/alle ********
%>
<tr>
	<td valign="top">
			<%
			call periode(periodeinterval, periodeintervalWeeks, StrTdato, StrUdato, "n")
			call findtimer("allealle", usesqlallemedarb, 0)
			%></table>
	</td>
</tr>
<%
call timefordelingsgrafik
%>
			
			
			<%
			if antalvalgtejob = 1 then
			'**** Ved alle/alle vises udspecificering af hvert valgt job. ** 
			intJobnrAlleAlle = Split(intJobnr, ", ")
			For i = 0 to Ubound(intJobnrAlleAlle)
			%>
			<tr><td style="padding-top:42;">
			<%
			strSQL = "SELECT id, jobnavn, jobnr, jobstartdato, jobslutdato, kkundenavn, jobtpris, budgettimer, fakturerbart, fastpris, ikkebudgettimer FROM job, kunder WHERE jobnr = " & intJobnrAlleAlle(i) &" AND kid = jobknr"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			thisjobid = oRec("id")
				if oRec("fakturerbart") = 1 then
				fakbargif = "<img src='../ill/eksternt_job_trans.gif' width='15' height='15' alt='Eksternt job' border='0'>"
				else
				fakbargif = "<img src='../ill/internt_job_trans.gif' width='15' height='15' alt='Internt job' border='0'>"
				end if
				
				if oRec("fakturerbart") = 1 then
					if oRec("fastpris") = 1 then
					typethisj = "Faspris <b>" & formatcurrency(jTpris, 2) &"</b>"
					else
					typethisj = "Budget (Ikke fastpris.) <b>" & formatcurrency(jTpris, 2) &"</b>"
					end if
				else
				typethisj = "(Internt)"
				end if
			
			%>
			<table cellpadding=0 cellspacing=1 border=0 width="500">
			<tr>
				<td colspan="2" bgcolor="#EFF3FF" style="padding-top:2; padding-left:5; border:#8CAAE6 1px solid;"><%=fakbargif%>
				&nbsp;
				<%if request("print") <> "j" then%>
				<a href="jobs.asp?menu=job&func=red&id=<%=oRec("id")%>&int=1">(<%=oRec("jobnr")%>)&nbsp;<%=oRec("jobnavn")%></a>&nbsp;&nbsp;(<%=oRec("kkundenavn")%>)
				<%else%>
				<a href="#">(<%=oRec("jobnr")%>)&nbsp;<%=oRec("jobnavn")%></a>&nbsp;&nbsp;(<%=oRec("kkundenavn")%>)
				<%end if%>
				<br>
				<%
						'*** Jobansvarlige ***
						
						strSQL2 = "SELECT mnavn, mnr, mid, job.id, jobans1 FROM job, medarbejdere WHERE job.id = "& oRec("id") &" AND mid = jobans1"
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						jobans1txtb = oRec2("mnavn") & " ("& oRec2("mnr") &")"
						end if
						oRec2.close
						
						
						
						strSQL2 = "SELECT mnavn, mnr, mid, jobans2 FROM job, medarbejdere WHERE job.id = "& oRec("id") &" AND mid = jobans2"
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						jobans2txtb = oRec2("mnavn") & " ("& oRec2("mnr") &")"
						end if
						oRec2.close
						
		
		%>
		<font class=megetlillesilver>
		Jobansvarlig 1: <%=jobans1txtb%><br>
	    Jobansvarlig 2: <%=jobans2txtb%></font>
		
		<%
		jobans1txtb = ""
		jobans2txtb = ""
		%>
				</td>
			</tr>
			<tr>
				<td bgcolor="#EFF3FF" style="padding-top:2; padding-right:5; border:#8CAAE6 1px solid;" align=right>Periode:<br>Type:</td>
				<td bgcolor="#FFFFFF" style="padding-top:2; padding-left:5; border:#8CAAE6 1px solid;"><b><%=formatdatetime(oRec("jobstartdato"), 1) & "</b> til <b>" & formatdatetime(oRec("jobslutdato"), 1) %></b><br>
				<%=typethisj%></td>
			</tr>
			<tr>
				<td bgcolor="#EFF3FF" style="padding-top:2; padding-right:5; border:#8CAAE6 1px solid;" align=right>Forkalkuleret timer:</td>
				<td valign=top bgcolor="#FFFFFF" style="padding-top:2; padding-left:5; border:#8CAAE6 1px solid;">Fakturerbare: <b><%=oRec("budgettimer")%></b>&nbsp;&nbsp;&nbsp;
				Ikke fakturerbare: <b><%
				if oRec("ikkebudgettimer") > 0 then
				Response.write oRec("ikkebudgettimer")
				else
				Response.write "0"
				end if
				%></b></td></tr>
			</table>
			<%
			end if
			oRec.close
			%>
			</td></tr>
			<tr>
				<td valign="top">
						<%
						call periode(periodeinterval, periodeintervalWeeks, StrTdato, StrUdato, "n")
						call findtimer("allealleudspec", usesqlallemedarb, intJobnrAlleAlle(i))
						%></table>
				</td>
			</tr>
			<tr><td valign="top" style="padding-top:5px;">
					<%if typethisj <> "(Internt)" then
					call ressourcetimer(thisjobid, "J")
					else
					Response.write "<br>"
					end if%>
			</td></tr>
			<%
			next
			else
			%>
			<tr><td><br><b>Udspecificering:</b><br>Vælg et job fra dropdown-listen for at se en udspecificering.<br><br></td></tr>
			<%
			end if
end if 'periodeinterval
end if
%>
<tr>
	<td valign="top">
		<%
		'****************************************************************************************
		'* Henter medarbejdere total 
		'****************************************************************************************
		if vis_medarb = "j" AND (periodeinterval >= 0 AND periodeinterval < 855) then
		if request("print") = "j" then
		%><br><br><br><table cellspacing="0" cellpadding="0" border="0"><%
		Response.write printhvilkejob
		end if%>
		<br>
		<table cellspacing="0" cellpadding="0" border="0">
		<tr>
			<td><br><img src="../ill/medarbikon.gif" width="36" height="30" alt="" border="0">&nbsp;&nbsp;
			<b>Medarbejdere</b><br>
			<%
			if showallealle <> "n" then
			%>
			<b>Udspecificering.</b><br>
			Vælg et enkelt job for at få udspecificeret medarbejder timer på job og aktiviteter.<br>
			<%end if
			%><br></td>
		</tr>
		<tr>
			<td>
				<table cellspacing="0" cellpadding="0" border="0">
				<%
				m = 0
				sqlallemedarb = " ("
				intMedArbValThis = Split(selmedarb, ", ")
				For m = 0 to Ubound(intMedArbValThis)
				Redim preserve medarbnavn(m)
				strSQL = "SELECT mnavn FROM medarbejdere WHERE mid = "& intMedArbValThis(m) 
				oRec.open strSQL, oConn, 0, 1 'adOpenForwardOnly, adLockReadOnly
				if not oRec.EOF then
				medarbnavn(m) = oRec("mnavn")
				end if
				oRec.close
				%>
				<tr>
					<td style="padding-bottom:2;">
						<%if m > 0 then%>
						<br><br>
						<%end if%>
					<b><%=medarbnavn(m)%></b>
					<%if showallealle = "n" then %>
					<br><font class=megetlillesort>På det valgte job.</font>
					<%else%>
					<br><font class=megetlillesort>Summeret på <b>Alle</b> de valgte job.</font>&nbsp;
					<%end if%>
					</td>
				</tr>
				<tr>
					<td valign="top">
							<%if showallealle = "n" then
								call periode(periodeinterval, periodeintervalWeeks, StrTdato, StrUdato, "n")
								call findtimer("medarbjobtotal", intMedArbValThis(m), 0)
								else
								call periode(periodeinterval, periodeintervalWeeks, StrTdato, StrUdato, "j")
								call findtimer("medarbjobtotalAlle", intMedArbValThis(m), 0)
							end if
							%>
								</table>
								<img src="ill/blank.gif" width="1" height="5" alt="" border="0"><br>
								
								<%
								if showallealle = "n" then
									call medarbejderFleksNorm(useJid, "J")
								else
									call medarbejderFleksNorm(antalvalgtejob, "J_all")
								end if
								%>
								
					</td>
				</tr>
				<%next%>
				</table>
		<br><br>&nbsp;</td>
		</tr>
		</table>
		<%end if%>
		
		<%if showallealle = "n" AND (periodeinterval >= 0 AND periodeinterval < 855) AND useFakbart = 1 then%>
		<table cellspacing="0" cellpadding="0" border="0">
		<%
		'******************************************************************************
		'* 							Henter aktiviteter 	     						  *
		'******************************************************************************
		%>
		<tr>
		<td width="855">
			<table border=0 cellspacing=0 cellpadding=0>
			<tr>
				<td><br><br><img src="../ill/aktikon.gif" width="30" height="32" alt="" border="0">&nbsp;&nbsp;&nbsp;<b>Aktiviteter</b></td>
			</tr>
			<tr>
				<td valign="top">
					<table cellspacing="0" cellpadding="0" border="0">
						<%
						strSQL2 = "SELECT id, navn, fakturerbar, budgettimer, aktstartdato, aktslutdato FROM aktiviteter WHERE job = " & useJid &" AND fakturerbar <> 2" 
						oRec2.open strSQL2, oConn, 3
						while not oRec2.EOF 
						%>
						<tr>
							<td valign="top" style="padding-top:25;">
							<table border="0" cellspacing="1" cellpadding="0" width="420"><tr>
								<td colspan=2 bgcolor="#FFFFF1" valign=top style="padding-top:2; padding-left:5; border:#8CAAE6 1px solid;">
								<%if oRec2("fakturerbar") = 1 then%>
								<img src="../ill/fakbarikon.gif" width="26" height="25" alt="" border="0">
								<%else%>
								<img src="../ill/notfakbarikon.gif" width="26" height="25" alt="" border="0">
								<%end if%>
								&nbsp;&nbsp;&nbsp;<b><%=oRec2("navn")%></b></td>
							</tr>
							<tr>
								<td bgcolor="#FFFFF1" valign=top style="padding-top:2; padding-right:5; border:#8CAAE6 1px solid;" align=right>Periode:<br>
								Timer tildelt:</td>
								<td bgcolor="#FFFFF1" style="padding-top:2; padding-left:5; border:#8CAAE6 1px solid;"><b><%=formatdatetime(oRec2("aktstartdato"), 1)%></b> til <b><%=formatdatetime(oRec2("aktslutdato"), 1)%></b><br>
								<b><%=oRec2("budgettimer")%></b></td>
							</tr>
							</table>
							
							</td>
						</tr>
						<tr>
							<td valign="top">
							<%
									call periode(periodeinterval, periodeintervalWeeks, StrTdato, StrUdato, "n")
									call findtimer("akttotal", usesqlallemedarb, oRec2("id"))
							%>
							</table>
							</td>
						</tr>
						<tr><td valign="top" style="padding-top:5px;">
							<%call ressourcetimer(oRec2("id"), "A")%>
						</td></tr>
						<tr>
							<td valign="top">
								<%
								'********************************************
								'* Henter medarbejdere pr. aktivitet		*
								'********************************************
								if vis_medarb = "j" AND (periodeinterval >= 0 AND periodeinterval < 855) then%>
								<table cellspacing="0" cellpadding="0" border="0">
										<%
										m = 0
										sqlallemedarb = " ("
										intMedArbValThis = Split(selmedarb, ", ")
										For m = 0 to Ubound(intMedArbValThis)
											Redim preserve medarbnavn(m)
											strSQL = "SELECT mnavn FROM medarbejdere WHERE mid = "& intMedArbValThis(m) 
											oRec.open strSQL, oConn, 0, 1 'adOpenForwardOnly, adLockReadOnly
											if not oRec.EOF then
											medarbnavn(m) = oRec("mnavn")
											end if
											oRec.close
											%>
											<tr><td><br><b><%=medarbnavn(m)%></b></td></tr>
											<tr>
												<td valign="top">
														<%if showallealle = "n" then
														call periode(periodeinterval, periodeintervalWeeks, StrTdato, StrUdato, "n")
														call findtimer("medprakt", intMedArbValThis(m), oRec2("id"))
														end if
														%>
														</table>
												</td>
											</tr>
											<tr><td valign="top" style="padding-top:5px;">
												<%if showallealle = "n" then
													call medarbejderFleksNorm(oRec2("id"), "A")
												end if%>
												</td></tr>
										<%next%>
								</table>
								<%end if
						oRec2.movenext
						wend
						oRec2.close
						%>
						<tr><td><br><br><br><br><br><br><br><br>&nbsp;</td></tr>
					</table>
					<!--/div-->
					</td>
				</tr>
				</table>
				<%end if%>
			<br><br><br><br><br>&nbsp;</td>
		</tr>
		</table>
	<br><br><br><br><br><br><br><br>&nbsp;</td>
</tr>
</table>
<%
if periodeinterval < 0 OR periodeinterval > 854 then 
Response.write "<br><br><br><br><br><br><br><img src='../ill/alert.gif' width='44' height='45' alt='' border='0'>&nbsp;<font class='error'>Fejl!</font>"_
	&"<br><img src='../ill/blank.gif' width='50' height='1' alt='' border='0'>"_
	&"<b>Der er valgt en ugyldig tidsperiode.</b>"_
	&"<br><br><img src='../ill/blank.gif' width='50' height='1' alt='' border='0'>"_
	&"Perioden skal være minimun på 1 dag. f.eks 4 august 2003 - 5 august 2003."_
	&"<br><img src='../ill/blank.gif' width='50' height='1' alt='' border='0'>"_
	&"og den må maksimal være på 855 dage (ca. 2½ år)."_
	&"<br><img src='../ill/blank.gif' width='50' height='1' alt='' border='0'>"_
	&"Den valgte periode er på: <b>" & periodeinterval & "</b>&nbsp;dage"_
	&"<br><br><img src='../ill/blank.gif' width='50' height='1' alt='' border='0'>"_
	&"Vælg en ny tidsperiode i toppen af siden..."
end if
%>
</div>
<%end if '** validering
end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
