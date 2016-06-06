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
	%>
	<%if request("print") <> "j" then%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	leftloading = "740"
	toploading = "50"
	rowbgcolor = "#5582D2"
	rowbgcolor2 = "#003399"
	else%>
	<html>
	<head>
		<title>timeOut 2.1</title>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
	</head>
	<body topmargin="0" leftmargin="0" class="regular">
	<%
	leftloading = "140"
	toploading = "250"
	rowbgcolor = "#ffffff"
	rowbgcolor2 = "#999999"
	end if%>
	<div id="loading" name="loading" style="position:absolute; top:<%=toploading%>; left:<%=leftloading%>; width:300; height:20; z-index:1000; visibility:visible; padding-left:4px; padding-top:5px;">
	<!--<img src="../ill/info.gif" width="42" height="38" alt="" border="0">--><b>Henter information...</b><br>
	<img src="../ill/loaderbar.gif" width="174" height="13" alt="" border="0">
	<!--Denne handling kan tage op til 20 sek. alt efter din forbindelse.-->
	</div>
	<%if request("print") <> "j" then%>
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	<%end if%>
	<!--#include file="inc/convertDate.asp"-->
	<script Language="JavaScript">
    function NewWin(url)    {
    window.open(url, 'print', 'width=620,height=680,scrollbars=yes,toolbar=no');    
	}
	</script>
	<script language="javascript">
	function showweek(weeknum,antalweeks) {
	for (i=0;i<antalweeks;i++){
	if(i == weeknum){
	document.all["weekdiv_"+weeknum+""].style.visibility = "visible";
	document.all["knap_"+i+""].style.borderColor = "darkRed";
	}
	else{
	document.all["weekdiv_"+i+""].style.visibility = "hidden";
	document.all["knap_"+i+""].style.borderColor = "#5582D2";
	}
	}
	}
	</script> 
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	</script>
	
	<%
	'******************************************************************************************
	'******************************************************************************************
	sub farvekoder
	%>
	<div id="submenu" style="position:absolute; left:755; top:530; visibility:visible; z-index:100;">
	<table cellspacing="0" cellpadding="0" border="0">
	<tr bgcolor="5582D2">
				<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
				<td valign="top"><img src="../ill/tabel_top.gif" width="134" height="1" alt="" border="0"></td>
				<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
		</tr>
		<tr bgcolor="5582D2">
				<td valign="top" class="alt"><b>Omsætning farvekoder:</b></td>
		</tr>
		<tr bgcolor="5582D2">
				<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="64" alt="" border="0"></td>
				<td  valign="top" class="alt"><img src="../ill/fakbarikon_lille.gif" width="16" height="15" alt="" border="0"><font color="#FFCC00">Fastpris</font><br>
				<img src="../ill/fakbarikon_lille.gif" width="16" height="15" alt="" border="0"><font color="#FFFFFF">Timepris</font><br>
				<img src="../ill/notfakbarikon_lille.gif" width="16" height="15" alt="" border="0"><font color="#cccccc">Ikke fakturerbar aktivitet</font></td>
				<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="64" alt="" border="0"></td>
		</tr>
		<tr bgcolor="<%=rowbgcolor%>">
			<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
			<td valign="bottom"><img src="../ill/tabel_top.gif" width="134" height="1" alt="" border="0"></td>
			<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
		</tr>
	</table>
	</div>
	<%
	end sub
	'******************************************************************************************
	'******************************************************************************************
	
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
		end function
				
				'*** func til beregning af omsætning
				public totOms
				public totOmsWeek
				public totOmsAkt
				public totOmsTotalM
				public intZtimer
				public intZtimerTotalAkt
				public intZtimerTotalWeek
				public intZtimerTotalMonth
				
				'** timeberegning **
				public function timeberegning()
				intZtimer = intZtimer + csng(oRec2("Ztimer"))
				intZtimerTotalAkt = intZtimerTotalAkt + csng(oRec2("Ztimer"))
				intZtimerTotalWeek = intZtimerTotalWeek + csng(oRec2("Ztimer"))
				intZtimerTotalMonth = intZtimerTotalMonth + csng(oRec2("Ztimer"))
				'** Sætter total timeforbrug pr aktivitet i den valgte periode ***
				timerAktTper(a) = timerAktTper(a) + csng(oRec2("Ztimer"))
				medarbTtot(t) = medarbTtot(t) + csng(oRec2("Ztimer"))
				end function
				
				
				'** Fastpris **
				public function omsFastpris1()
				totOms = totOms + ((jobTpris(a)/budgetTimer(a)) * csng(oRec2("Ztimer")))
				totOmsWeek = totOmsWeek + ((jobTpris(a)/budgetTimer(a)) * csng(oRec2("Ztimer")))
				totOmsAkt = totOmsAkt + ((jobTpris(a)/budgetTimer(a)) * csng(oRec2("Ztimer")))
				totOmsTotalM = totOmsTotalM  + ((jobTpris(a)/budgetTimer(a)) * csng(oRec2("Ztimer")))
				timerAktOmsper(a) = timerAktOmsper(a) + ((jobTpris(a)/budgetTimer(a)) * csng(oRec2("Ztimer")))
				medarbOmstot(t) = medarbOmstot(t) + ((jobTpris(a)/budgetTimer(a)) * csng(oRec2("Ztimer")))
				end function
				
				''**** ikke fastpris ****
				public function omsIkkeFastpris()
					if oRec2("Tfaktim") = "1" then
						totOms = totOms + (csng(oRec2("Ztimer")) * ccur(oRec2("TimePris")))
						totOmsWeek = totOmsWeek + (csng(oRec2("Ztimer")) * ccur(oRec2("TimePris")))
						totOmsAkt = totOmsAkt + (csng(oRec2("Ztimer")) * ccur(oRec2("TimePris")))
						totOmsTotalM = totOmsTotalM + (csng(oRec2("Ztimer")) * ccur(oRec2("TimePris")))
						'** Sætter total timeforbrug på en aktivitet i den valgte periode ***
						timerAktOmsper(a) = timerAktOmsper(a) + (csng(oRec2("Ztimer")) * ccur(oRec2("TimePris")))
						medarbOmstot(t) = medarbOmstot(t) + (csng(oRec2("Ztimer")) * ccur(oRec2("TimePris")))
					else
						'** Hvis aktiviteten ikke er fakturerbar **
						totOms = totOms 
						totOmsWeek = totOmsWeek 
						totOmsAkt = totOmsAkt 
						totOmsTotalM = totOmsTotalM 
						'** Sætter total timeforbrug på en aktivitet i den valgte periode ***
						timerAktOmsper(a) = timerAktOmsper(a) 
						medarbOmstot(t) = medarbOmstot(t) 
					end if
				end function
		
				'**** sub til genbuge hvis der ikke er omsætning på et job.
				'**** Hvis et job ikke er fakturerbart 
				public function ingenOms()
					if totOms > 0 then 
					totOms = totOms
					else
					totOms = 0
					end if
					
					if totOmsWeek > 0 then
					totOmsWeek = totOmsWeek
					else
					totOmsWeek = 0
					end if
				
					if totOmsAkt > 0 then
					totOmsAkt = totOmsAkt
					else
					totOmsAkt = 0
					end if
					
					if totOmsTotalM > 0 then
					totOmsTotalM = totOmsTotalM
					else
					totOmsTotalM = 0
					end if
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
	
	
	function opbygning()
	'*** Opbygning af oversigt medarbejdere *****
			areMaP = instr(strMedarbPrinted, ", " &oRec("Tmnr")&"#")
			
			if areMaP = "0" then
			Redim preserve medarbejder(t)
			Redim Preserve medarbID(t)
			medarbID(t) = oRec("Tmnr")  
			medarbejder(t) = oRec("Tmnavn")
			Redim preserve medarbTtot(t) 
			Redim preserve medarbOmstot(t)
			t = t + 1
			end if
			
			strMedarbPrinted = strMedarbPrinted & ", " & oRec("Tmnr") &"#" 
			
			
			'****************************************************************************
			'*** Opbygning af oversigt Aktiviteter 									*****
			'****************************************************************************
			areAktP = instr(strAktPrinted, ", " &oRec("TAktivitetId")& "#")
			
			if areAktP = "0" then
			Redim preserve aktivitetID(a)
			Redim preserve aktiviteter(a)
			Redim preserve jobnavn(a)
			Redim preserve jobnr(a)
			Redim preserve kundenavn(a)
			Redim preserve fastpris(a)
			Redim preserve faktim(a)
			Redim preserve imgJob(a)
			Redim preserve jobfakbart(a)
			Redim preserve budgetTimer(a)
			Redim preserve jobTpris(a)
			
			'*** Fra Timer ****
			aktivitetID(a) = oRec("TAktivitetId")
			aktiviteter(a) = oRec("Anavn")
			jobnavn(a) = oRec("Tjobnavn")
			jobnr(a) = oRec("Tjobnr")
			kundenavn(a) = oRec("Tknavn")
			fastpris(a) = oRec("fastpris")
			faktim(a) = oRec("Tfaktim")
			
			'**** Fra Job *****
			jobTpris(a) = oRec("jobTpris")
			budgetTimer(a) = oRec("budgetTimer")
			jobfakbart(a) = oRec("fakturerbart")
			
			if jobfakbart(a) = 1 then
			imgJob(a) = "<img src='../ill/eksternt_job_trans.gif' width='15' height='15' alt='Eksternt job' border='0'>"
			else
			imgJob(a) = "<img src='../ill/internt_job_trans.gif' width='15' height='15' alt='Internt job' border='0'>"
			end if
			
			Redim preserve timerAktTper(a)
			Redim preserve timerAktOmsper(a)
			a = a + 1
			end if
			
			strAktPrinted = strAktPrinted & ", " & oRec("TAktivitetId") &"#" 
			
			
				'*** Opbygning af weeknumbers *****
				areWeekP = instr(strWeekNumPrinted, ", " &strWeekNum&"#")
				if areWeekP = "0" then
				Redim preserve weeknumbers(w)
				weeknumbers(w) = strWeekNum
				w = w + 1
				end if
			
			
				'*** Opbygning af yearnumbers *****
				areYearP = instr(strYearNumPrinted, ", " & strYearNum &"#")
				if areYearP = "0" then
				y = y + 1
				strWeekNumPrinted = "#0"
				yearweeks(y,week) = strWeekNum
				week = week + 1
				else
					if areWeekP = "0" then
					yearweeks(y,week) = strWeekNum
					week = week + 1
					end if
				end if
				
				strWeekNumPrinted = strWeekNumPrinted & ", " & strWeekNum &"#" 
				strYearNumPrinted = strYearNumPrinted & ", " & strYearNum &"#" 
			'**************  opbygning slut ********************
	end function
	
	
	'function AktprMedTot()
	'end function
	
	'*******************************************************************************
	'****************       Start sideopbygning og variable          ***************
	'** Sætter Periode værdier **
	FM_medarb = request("FM_medarb")
	FM_job = request("FM_job")

	if len(trim(FM_job)) = "0" then
	FM_job = 909909909909
	noJobMessage = "<img src='../ill/info.gif' width='42' height='38' alt='' border='0'><br>Der er ikke valgt noget job!<br>vælg job <a href='stat.asp?menu=stat&shokselector=1&FM_kunde=0&FM_progrupper=10'>her</a>"
	else
	noJobMessage = ""
	end if
	
	FM_Aar = request("FM_Aar")
	eks = request("eks")
	func = request("func")
	weekselector = request("weekselector")
	strYear = request("year") 
	yearsel = request("yearsel")
	
	if weekselector = "j" then 
	StrTdato = request("FM_start_mrd") &"/" & request("FM_start_dag") & "/20" & request("FM_start_aar")
	StrUdato = request("FM_slut_mrd") &"/" & request("FM_slut_dag") & "/20" & request("FM_slut_aar")
	UseWeekSelDate = StrUdato 
	strWSelStartDato = cdate(StrTdato)
	strWSelEndDato = cdate(StrUdato) 
	id = "1"
		'if request("FM_start_aar") <> request("FM_slut_aar") then
		strReqAar_slut = request("FM_slut_aar")
		'end if
	else
		if month(now()) <> 1 then
		StrTdato = (month(now()) - 1) &"/1/" & year(now())
		else
		StrTdato = month(now) &"/1/" & year(now)
		end if
	StrUdato = month(now()) &"/" & day(now()) & "/" & year(now())
	id = 1
	strReqAar_slut = 2025
	end if
	'**** 
	
	linket = request("linket")
	
	selmedarb = FM_medarb 'request("selmedarb")
	if len(selmedarb) = "0" then
	selmedarb = 0
	else
	selmedarb = selmedarb
	end if
	
	selaktid = request("selaktid")
	intJobnr = FM_job 'request("jobnr")
	
	if len(Request("mrd")) <> "0" then
		strReqMrd = Request("mrd")
	else
	strReqMrd = month(now)
	end if
	
	'**** finder ud af om der er valgt et år ***
	if weekselector <> "j" then 
		if len(strYear) <> "0" then
			if strYear = "-1" then
			strReqAar = "0"
			else
			strReqAar = strYear
			end if	
		else
		strReqAar = right(year(now), 2)
		end if
		strReqAar_slut = strReqAar
	else
	strReqAar = request("FM_start_aar")
	end if	
	
	if weekselector <> "j" then 
		if strReqAar <> "0" then
			if strReqAar > 10 then
			yearKri = " AND Taar = '2"&strReqAar&"' "
			else
			yearKri = " AND Taar = '20"&strReqAar&"' "
			end if
		else
		yearKri = ""
		end if
	else
	yearKri = ""
	end if	
	
	'Response.write "strReqAar " & strReqAar
	'Response.write "yearKri: " & yearKri
	
	lastFakdag = Request("lastFakdag")
	
	'*** Bruges som markør af mrd knapper **
	if weekselector = "j" then
	framrd = datepart("m", StrTdato)
	tilmrd = datepart("m", StrUdato)
	end if

thisfile = "joblog_z"
if request("print") <> "j" then%>
	<!--#include file="inc/joblog_z_mrdk.asp"-->
	<%
end if 'print

'*** Finder de medarbejdere og job der er valgt ***
Dim intMedArbVal 
Dim b
Dim intJobnrKriValues 
Dim i
		
			intMedArbVal = Split(selmedarb, ", ")
			For b = 0 to Ubound(intMedArbVal)
				intJobnrKriValues = Split(intJobnr, ", ")
				For i = 0 to Ubound(intJobnrKriValues)
				if selmedarb <> "0" then
				jobMedarbKri = jobMedarbKri & " Tmnr = " & intMedArbVal(b) &" AND Tjobnr = " & intJobnrKriValues(i) & yearKri &" OR "
				else
				jobMedarbKri = jobMedarbKri & " Tmnr <> 0 AND Tjobnr = " & intJobnrKriValues(i) & yearKri &" OR "
				end if
				next
			next
			intJobMedarbKriCount = len(jobMedarbKri)
			trimIntJobMedarbKri = left(jobMedarbKri, (intJobMedarbKriCount-3))
			jobMedarbKri = trimIntJobMedarbKri & " "


'*** Finder de aktiviteter der er valgt ***
if len(request("selaktid")) <> "0" then
selaktidKri = "AND TAktivitetId = " & selaktid &""
else
selaktidKri = ""
end if

if request("print") <> "j" then
visWeekSel = "visible"
else
visWeekSel = "hidden"
end if%>	
<!--#include file="inc/weekselector.asp"-->
<%
if request("yearsel") <> "j" then
	'** Fakturerbart / Ikke fakturerbart job. **
	strSQL = "SELECT DISTINCT (tid), Tdato, TasteDato, Tjobnr, Tjobnavn, TAktivitetNavn AS Anavn, TAktivitetId, Tknavn, Tmnr, Tmnavn, Timer, Tfaktim, TimePris, Timerkom, timer.fastpris, fakturerbart, jobTpris, budgetTimer, Taar FROM timer LEFT JOIN job ON (job.jobnr = timer.Tjobnr) WHERE "& jobMedarbKri &" "& selaktidKri &" ORDER BY Tdato, Tjobnavn, Tmnavn" 
	'Ikke fjern ORDER BY tdato da ugevisnigen ellers vil blive defekt
	oRec.open strSQL, oConn, 0, 1 'adOpenForwardOnly, adLockReadOnly
	
	strYearNumPrinted = "0#"
	strWeekNumPrinted = "0#"
	strMedarbPrinted = "0#"
	strAktPrinted = "0#"
	a = 0
	t = 0
	w = 0
	v = 0
	y = -1
	week = 0
	
	Redim medarbID(t)
	Redim medarbejder(t)
	Redim medarbOmstot(t)
	Redim medarbTtot(t)
	Redim aktivitetID(a)
	Redim aktiviteter(a)
	Redim jobnavn(a)
	Redim jobnr(a)
	Redim fastpris(a)
	Redim faktim(a)
	Redim imgJob(a)
	Redim jobfakbart(a)
	Redim weeknumbers(w)
	'Redim yearnumbers(y)
	Dim yearweeks(2,52)
	Redim timerAktTper(a)
	Redim timerAktOmsper(a)
	Redim kundenavn(a)
	Redim fontenakttotal(a)
	Redim jobTpris(a)
	Redim budgetTimer(a) 
	
	while not oRec.EOF
		StrTdato = formatdatetime(oRec("Tdato"), 0)
		'Response.write "StrTdato"& oRec("Tid")&", " & StrTdato & "<br>"
		strWeekNum = datepart("ww", StrTdato)
		strYearNum = datepart("yyyy", StrTdato)
		id = 1
		
		if cint(strReqAar) = cint(strReqAar_slut) AND weekselector <> "j" AND oRec("Tfaktim") <> 5 then
			StrUdato = "31/12/2007"
			%>
			<!--#include file="inc/dato2.asp"-->
			<%
			if cint(strAar) = cint(strReqAar) OR cint(strReqAar) = "0" then
			if cint(strReqMrd) = cint(strMrd) OR cint(strReqMrd) = "0" then
			if datepart("ww", oRec("tdato")) => frauge AND datepart("ww", oRec("tdato")) <= tiluge then 
			
			call opbygning()
			
			end if
			end if
			end if
		else
			StrUdato = UseWeekSelDate
			%>
			<!--#include file="inc/dato2.asp"-->
			<%
			if cDate(StrTdato) => strWSelStartDato AND cDate(StrTdato) <= strWSelEndDato AND oRec("Tfaktim") <> 5 then
				call opbygning()
			end if
		end if
	
	
	oRec.movenext
	wend	  
		
	oRec.Close
	Set oRec = Nothing

'**********************************************************************************************
'**********************************************************************************************	
'*************************       Z-siden        ***********************************************
'**********************************************************************************************
'**********************************************************************************************
%>
<!--#include file="inc/stat_submenu.asp"-->
<%
if request("print") <> "j" then
pleft = 190
ptop = 55
else
pleft = 20
ptop = 10
end if
%>	
<!-- venstre side div -->
<div id="z" style="position:absolute; left:<%=pleft%>; top:<%=ptop%>; width:300; height:600; visibility:visible; z-index:100;">
<table cellspacing="0" cellpadding="0" border="0" width="600">
<tr><td><img src="../ill/header_joblog_z.gif" alt="" border="0" width="758" height="45">
<%
'*** Hvis periodevælger eller yearSel er brugt ***
if weekselector = "j" OR yearsel = "j" then

if weekselector = "j" then
	strOverskrift = "Den valgte periode er:"
	strTekst = showStrTdato &"&nbsp;(uge&nbsp;"& frauge &")&nbsp;" &" - " & showStrUdato &"&nbsp;(uge&nbsp;"& tiluge &")"
end if

if yearsel = "j" then
	strOverskrift = "Se ugeoversigt?"
	strTekst = "Klik på månederne ovenfor får at se timefordeling / uge fordelt på medarbejder(e)"
end if

%>
<div id="periodeShow" style="position:absolute; left:5; top:181; visibility:visible; z-index:100;">
<img src="../ill/info.gif" width="42" height="38" alt="" border="0">
<table border="0" cellpadding="0" cellspacing="0" width="250">
 <tr bgcolor="5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan="3"><img src="../ill/tabel_top.gif" width="234" height="1" alt="" border="0"></td>
		<td align=right width="8" rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
</tr>
<tr bgcolor="#5582D2">
		<td colspan="3" class='alt' valign='middle'><%=strOverskrift%></td>
</tr>
<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="60" alt="" border="0"></td>
		<td colspan="3" class='alt' height='60' valign='middle' align='center'><b><%=strTekst%></b></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="60" alt="" border="0"></td>
</tr>
<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="3" valign="bottom"><img src="../ill/tabel_top.gif" width="234" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
</table></div>
<%
end if
'*** slut periodevælger/yearsel ***
%>
<br></td></tr></table>
<!-------------------------------Sideindhold venstre side------------------------------------->
<%
'** Hvis der findes uger med registreringer i den valgte periode.**

if week > 0 then
	
		if request("print") <> "j" AND weekselector <> "j" AND yearsel <> "j" then
				%>
				<!--#include file="inc/joblog_z_ugek.asp"-->
				<%
		end if
	
		
		intZtimer = 0
		For y = 0 to 1
		if len(trim(week)) <> "0" then
		 	if weekselector = "j" then
				week = 1
			end if
		for week = 0 to week - 1
			if trim(len(yearweeks(y, week))) <> "0" then
			
			if week = g then
				varVisi = "visible"
				else
				varVisi = "hidden"
			end if
				varWeekTop = varTop - 32
				
				if request("print") <> "j" AND weekselector <> "j" AND yearsel <> "j" then%>
				<div id="weekdiv_<%=week%>" name="weekdiv_<%=week%>" style="position:absolute; top:<%=varWeekTop%>; left:0; z-index:100; visibility:<%=varVisi%>;">
				<%end if%>
				<%if weekselector <> "j" AND yearsel <> "j" then%>
				<br><br><br><br><br>
				<table border="0" cellpadding="0" cellspacing="0" width="260">
				 <tr bgcolor="<%=rowbgcolor%>">
					<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
						<td colspan="3"><img src="../ill/tabel_top.gif" width="244" height="1" alt="" border="0"></td>
					<td align=right width="8" rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
				</tr>
				<tr bgcolor="#5582D2">
					<td width="160">&nbsp;&nbsp;<font class=lilleblaa>Aktiviteter</font></td>
					<td align=right width="30"><font class=lilleblaa>Timer</td>
					<td align=right width="70"><font class=lilleblaa>Oms.&nbsp;</td>
				</tr>
				<tr bgcolor="#D4DDF5">
					<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="1" alt="" border="0"></td>
					<td height=1 colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="1" alt="" border="0"></td>
				</tr>
				<%end if%>
				<%
				for a = 0 to a - 1
					%>
					<%if weekselector <> "j" AND yearsel <> "j" then%>
					<tr bgcolor="<%=rowbgcolor%>">
							<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="55" alt="" border="0"></td>
							<td height="20" valign="bottom" colspan="3" style="padding-left : 4px;" class='alt'>
							<%if len(jobnavn(a) ) > 14 then
							strJobnavnThis = left(jobnavn(a) , 14) & ".."
							else
							strJobnavnThis =  jobnavn(a) 
							end if 
							
							if len(kundenavn(a)) > 10 then
							strKnavnThis = left(kundenavn(a), 10) & ".."
							else
							strKnavnThis = kundenavn(a)
							end if 
							
							if len(aktiviteter(a)) > 20 then
							strAktThis = left(aktiviteter(a), 20) & ".."
							else
							strAktThis = aktiviteter(a)
							end if 
							
							
							Response.write "<br>"& imgJob(a) &"&nbsp;&nbsp;<b>" & jobnr(a) & "&nbsp;" & strJobnavnThis &"&nbsp;"_
							& "</b>&nbsp;&nbsp;<font class=lilleblaa>("& strKnavnThis &")</font><br><b>"&strAktThis&":</b>"
						%></td>
					<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="55" alt="" border="0"></td>
						</tr>
						<%end if%>
							<%
							for t = 0 to t - 1
							%>
							<%if weekselector <> "j" AND yearsel <> "j" then%>
							<tr bgcolor="#D4DDF5">
								<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="1" alt="" border="0"></td>
								<td height=1 colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
								<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="1" alt="" border="0"></td>
							</tr>
							<tr bgcolor="<%=rowbgcolor%>">
								<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="20" alt="" border="0"></td>
								<td style="padding-left : 4px;" class='alt'><%=medarbejder(t)%></td>
								<td align=right style="padding-left : 4px; padding-right : 1px;" class='alt'>
							<%end if%>
								<%
								select case faktim(a)
								case "1", "2" 
								strSQL = "SELECT DISTINCT(tid), timer AS Ztimer, tdato, Tfaktim, TAktivitetId, TimePris, fastpris, Taar FROM timer WHERE Tmnr = "&medarbID(t)&" AND TAktivitetId = "& aktivitetID(a) & yearKri &""
								case else
								strSQL = "SELECT DISTINCT(tid), timer AS Ztimer, tdato, Tfaktim, TAktivitetId, TimePris, timer.fastpris AS fastpris, Taar FROM timer WHERE Tmnr = "&medarbID(t)&" AND Tjobnr = "& jobnr(a) & yearKri &""
								end select
									
									oRec2.open strSQL, oConn, 0, 1 'adOpenForwardOnly, adLockReadOnly
									
									while not oRec2.EOF
									if weekselector <> "j" then 
										if datepart("ww", oRec2("tdato")) = yearweeks(y,week) then
										
										
											call timeberegning
											
											if oRec2("Tfaktim") = "1" then
												if oRec2("fastpris") = "1" then
													call omsFastpris1										
												else 
													call omsIkkeFastpris
												end if
											else
												call ingenOms
											end if
										end if
										
									else
										
										'*** Hvis der er valgt perode med periode vælgeren ****
										if oRec2("tdato") => strWSelStartDato AND oRec2("tdato") <= strWSelEndDato then
											
											'Er det en fakbar aktivitet?
											if int_eks_job = 1 then 
											fakbar = oRec2("akt_fakturerbar")
											else
											fakbar = 0
											end if
											
											call timeberegning
											
											if oRec2("Tfaktim") = "1" then
												if oRec2("fastpris") = "1" then
													call omsFastpris1
												else 
													call omsIkkeFastpris
												end if
											else
												call ingenOms
												
											end if
										end if
									end if
									
									tfaktimerne = oRec2("Tfaktim")
									if tfaktimerne = 1 then
										if oRec2("fastpris") = 1 then
											if request("print") <> "j" then
											fonten = "#FFCC00"
											else
											fonten = "#000000"
											end if
										else
											if request("print") <> "j" then
											fonten = "#FFFFFF"
											else
											fonten = "#000000"
											end if
										end if
									else
											if request("print") <> "j" then
											fonten = "#CCCCCC"
											else
											fonten = "#000000"
											end if
									end if
									
									oRec2.movenext
									wend
									oRec2.close
									
									'** Sætter fonten til aktiviteter pr. medarb **
									Redim preserve fontenakttotal(a)
									fontenakttotal(a) = fonten
									
								if weekselector <> "j" AND yearsel <> "j" then
								Response.write SQLBless(intZtimer) &"</td>"
								%><td align="right" style='padding-left : 4px; padding-right : 1px;'><font color="<%=fonten%>"> 
								<%=kommaFunc(totOms)%> dkr.
								</td>
								<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="20" alt="" border="0"></td>
								</tr>
								<%
								end if
								intZtimer = 0
								totOms = 0
							next
							%>
				<%if weekselector <> "j" AND yearsel <> "j" then%>
				<tr bgcolor="#D4DDF5">
					<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="1" alt="" border="0"></td>
					<td height=1 colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="1" alt="" border="0"></td>
				</tr>
				<tr bgcolor="<%=rowbgcolor%>">
					<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="20" alt="" border="0"></td>
					<td style="padding-left : 4px;" class='alt'>Ialt: </td>
					<td align=right style="padding-left : 4px; padding-right : 1px;" class='alt'><%
					Response.write SQLBless(intZtimerTotalAkt) %>
					<td align="right" style='padding-left : 4px; padding-right : 1px;'><font color="<%=fonten%>">
				<%
				Response.write kommaFunc(totOmsAkt) & "&nbsp;dkr."
				%>
				</td>
				<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="20" alt="" border="0"></td>
				</tr>
				<tr bgcolor="#cccccc">
					<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="2" alt="" border="0"></td>
					<td height=2 colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					<td valign="top" align=right><img src="../ill/tabel_top.gif" width="2" height="2" alt="" border="0"></td>
				</tr>
				<%
				end if
				intZtimerTotalAkt = 0
				totOmsAkt = 0
				next
				%>
			<%if weekselector <> "j" AND yearsel <> "j" then%>
			<tr bgcolor="<%=rowbgcolor%>">
				<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
				<td height=20 colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
				<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
			</tr>
			<tr bgcolor="<%=rowbgcolor%>">
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
			<td style="padding-left : 4px;" class='alt'><b>Uge <%=yearweeks(y, week)%> total:</b></td>
			<td align="right" style="padding-left : 4px; padding-right : 4px;" class='alt'><b><%
			Response.write SQLBless(intZtimerTotalWeek) 
			%></b></td>
			<td align="right" style="padding-left : 4px; padding-right : 1px;" class='alt'><b><%= kommaFunc(totOmsWeek)%> dkr.</b></td>
			<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
			</tr>
			<tr bgcolor="#5582D2">
				<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="10" alt="" border="0"></td>
				<td height=10 colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
				<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="10" alt="" border="0"></td>
			</tr>
			</TABLE><br><br><br>
			<%if request("print") <> "j" then%>
			</div>
			<%end if%>
			<%end if
			intZtimerTotalWeek = 0
			totOmsWeek = 0
			end if
		next
	end if
	Next
'end if '** venstre side vises
%>
<!-- slut venstre div -->
</div>
<%


'*** Periode total ****
%>
<!-- Højre side div -->
<%if request("print") <> "j" then%>	
<div id="akttotal" style="position:absolute; left:470; top:188; visibility:visible; z-index:200;">
<%else%>
<div id="akttotal" style="position:absolute; left:320; top:120; visibility:visible; z-index:200;">
<%end if%>
<table cellspacing="1" cellpadding="0" border="0" width="270">
<tr>
	<%if weekselector <> "j" then%>
	<td width="100" style="!border: 1 px; background-color: #FFFFFF; border-color: <%=rowbgcolor%>; border-style: solid; padding-left : 4px;  padding-right : 4px;"><font size="1">fra uge - til uge</td>
	<%else%>
	<td width="100" style="!border: 1 px; background-color: #FFFFFF; border-color: <%=rowbgcolor%>; border-style: solid; padding-left : 4px;  padding-right : 4px;"><font size="1">fra dato - til dato</td>
	<%end if%>
	<td><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	<td align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: <%=rowbgcolor%>; border-style: solid; padding-left : 4px;  padding-right : 4px;"><font size="1">timer</td>
	<td><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	<td align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: <%=rowbgcolor%>; border-style: solid; padding-left : 4px;  padding-right : 4px;"><font size="1">omsætning</td>
</tr>
<tr>
	<%if weekselector <> "j" then%>
	<td style="!border: 1 px; background-color: #FFFFFF; border-color: <%=rowbgcolor%>; border-style: solid; padding-left : 4px; padding-right : 1px;"><font size="1"><b><%=weeknumbers(0)&"&nbsp;&nbsp;20"&strReqAar&"&nbsp;&nbsp;-&nbsp;&nbsp; "& weeknumbers(w-1)%>&nbsp;&nbsp;20<%=strReqAar_slut%></b></td>
	<%else%>
	<td style="!border: 1 px; background-color: #FFFFFF; border-color: <%=rowbgcolor%>; border-style: solid; padding-left : 4px; padding-right : 1px;"><font size="1"><b><%=convertDate(strWSelStartDato)%> - <%=convertDate(strWSelEndDato)%></b></td>
	<%end if%>
	<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	<td valign="bottom" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: <%=rowbgcolor%>; border-style: solid; padding-left : 4px;  padding-right : 4px;">
	<b><%
	for a = 0 to a - 1 
	  timerPerTot = timerPerTot + timerAktTper(a)
	  omsTotPer = omsTotPer + timerAktOmsper(a)
	 next
	
	Response.write SQLBless(timerPerTot)
	timerPerTot = 0
	
	%></b></td>
	<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	<td valign="bottom" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: <%=rowbgcolor%>; border-style: solid; padding-left : 4px;  padding-right : 2px;"><b><%=kommaFunc(omsTotPer)%></b></td>
	</tr>
<tr>
	<td height=20 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
</table>
<%omsTotPer = 0

'*** Aktiviteter total ****
%>
<img src="../ill/aktikon.gif" width="30" height="32" alt="" border="0"><br>
<table cellspacing="0" cellpadding="0" border="0" width="270">
<tr bgcolor="5582D2">
	<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
	<td colspan="3"><img src="../ill/tabel_top.gif" width="254" height="1" alt="" border="0"></td>
	<td align=right width="8" rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
</tr>	
<tr bgcolor="#5582D2">
	<td width="150">&nbsp;&nbsp;<font class=lilleblaa>Aktiviteter</font></td>
	<td align=right width="30"><font class=lilleblaa>Timer</td>
	<td align=right width="90"><font class=lilleblaa>Oms.&nbsp;</td>
</tr>
<tr bgcolor="#D4DDF5">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="1" alt="" border="0"></td>
		<td height=1 colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="1" alt="" border="0"></td>
</tr><%
For a = 0 to a - 1
	%>
		<tr bgcolor="<%=rowbgcolor%>">
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
			<td valign="bottom" style="padding-left : 4px;" class='alt' colspan="3"><%
			if len(jobnavn(a) ) > 14 then
			strJobnavnThis = left(jobnavn(a) , 14) & ".."
			else
			strJobnavnThis =  jobnavn(a) 
			end if 
			
			if len(kundenavn(a)) > 18 then
			strKnavnThis = left(kundenavn(a), 18) & ".."
			else
			strKnavnThis = kundenavn(a)
			end if 
			
			if len(aktiviteter(a)) > 20 then
			strAktThis = left(aktiviteter(a), 20) & ".."
			else
			strAktThis = aktiviteter(a)
			end if
			
			Response.write imgJob(a) &"&nbsp;&nbsp;<b>" & jobnr(a) & "&nbsp;" & strJobnavnThis &"&nbsp;"_
			& "</b>&nbsp;&nbsp;<font class=lilleblaa>("& strKnavnThis &")</font>"
			%></td>
			<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#D4DDF5">
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="1" alt="" border="0"></td>
			<td height=1 colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
			<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr bgcolor="<%=rowbgcolor%>">
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="20" alt="" border="0"></td>
			<td style="padding-left : 4px; padding-right : 1px;" class='alt'><b><%=strAktThis%>:</b></td>
			<td align="right" style="padding-left : 4px; padding-right : 1px;" class='alt'><%
			Response.write SQLBless(timerAktTper(a)) %></td>
			<td align="right" style="padding-left : 4px;  padding-right : 1px;"><font color="<%=fontenakttotal(a)%>"><%=kommaFunc(timerAktOmsper(a))%> dkr.</td>
			<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="20" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#cccccc">
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="2" alt="" border="0"></td>
			<td height=2 colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
			<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="2" alt="" border="0"></td>
	</tr>
	<%
	next%>
	<tr bgcolor="#5582D2">
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="10" alt="" border="0"></td>
			<td height=10 colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
			<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="10" alt="" border="0"></td>
	</tr>
	</table>
	<br><br><br>
	<img src="../ill/medarbikon.gif" width="36" height="30" alt="" border="0"><br>
	<table cellspacing="0" cellpadding="0" border="0" width="270">
	<tr bgcolor="5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan="3"><img src="../ill/tabel_top.gif" width="254" height="1" alt="" border="0"></td>
		<td align=right width="8" rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>	
	<tr bgcolor="#5582D2">
		<td>&nbsp;&nbsp;<font class=lilleblaa>Medarbejder(e)</font></td>
		<td align=right width="30"><font class=lilleblaa>Timer</td>
		<td align=right width="120"><font class=lilleblaa>Oms.&nbsp;</td>
	</tr>
	<%
	'** Medarbejder total **
	For t = 0 to t - 1
	%>
	<tr bgcolor="#D4DDF5">
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="1" alt="" border="0"></td>
			<td height=1 colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
			<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr bgcolor="<%=rowbgcolor%>">
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
			<td style="padding-left : 4px;" class='alt'><b><%=medarbejder(t) %></b></td>
			<td align=right style="padding-left : 4px; padding-right : 4px;" class='alt'>
			<b><%
			Response.write SQLBless(medarbTtot(t)) 
			%></b></td>
		<td align="right" style="padding-left : 4px;  padding-right : 1px;" class='alt'><b><%=kommaFunc(medarbOmstot(t))%></b> dkr.</td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	next
	%>
	<tr bgcolor="#5582D2">
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="10" alt="" border="0"></td>
			<td height=10 colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
			<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="10" alt="" border="0"></td>
	</tr>
	</table><br>
	<br>
	<%
		'*** aktiviteter pr. medarbejder ****
		%><img src="../ill/aktikon.gif" width="30" height="32" alt="" border="0">&nbsp;<img src="../ill/medarbikon.gif" width="36" height="30" alt="" border="0"><br>
		<table cellspacing="0" cellpadding="0" border="0" width="271">
		<tr bgcolor="5582D2">
			<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="28" alt="" border="0"></td>
			<td colspan="3"><img src="../ill/tabel_top.gif" width="264" height="1" alt="" border="0"></td>
			<td align=right width="8" rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="28" alt="" border="0"></td>
		</tr>	
		<tr bgcolor="#5582D2">
			<td width=200>&nbsp;&nbsp;<font class=lilleblaa>Aktiviteter fordelt pr. medarbejder:</font></td>
			<td align=right width="30"><font class=lilleblaa>Timer</td>
			<td align=right width="120"><font class=lilleblaa>Oms.&nbsp;</td>
		</tr>	
		<%
		For t = 0 to t - 1
		%>
		<tr bgcolor="#D4DDF5">
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="1" alt="" border="0"></td>
			<td height=1 colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
			<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="1" alt="" border="0"></td>
		</tr>
		<tr bgcolor="<%=rowbgcolor2%>">
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
			<td colspan="3" valign="bottom" height="30" style="padding-left : 4px;" class='alt'><b><%=medarbejder(t) %></b></td>
			<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		</tr>
		<tr bgcolor="#cccccc">
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="2" alt="" border="0"></td>
			<td height=2 colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
			<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="2" alt="" border="0"></td>
		</tr>
		<%
		For a = 0 to a - 1
		%>
				<tr bgcolor="<%=rowbgcolor%>">
					<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
					<td style="padding-left : 4px;" class='alt' colspan="3" valign="bottom">
					<%if len(jobnavn(a) ) > 14 then
						strJobnavnThis = left(jobnavn(a) , 14) & ".."
						else
						strJobnavnThis =  jobnavn(a) 
						end if 
						
						if len(kundenavn(a)) > 18 then
						strKnavnThis = left(kundenavn(a), 18) & ".."
						else
						strKnavnThis = kundenavn(a)
						end if 
						
						if len(aktiviteter(a)) > 20 then
						strAktThis = left(aktiviteter(a), 20) & ".."
						else
						strAktThis = aktiviteter(a)
						end if 
						Response.write imgJob(a) &"&nbsp;&nbsp;<b>" & jobnr(a) & "&nbsp;" & strJobnavnThis &""_
						& "</b>&nbsp;&nbsp;&nbsp;<font class=lilleblaa>("& strKnavnThis &")</font></td>"%>
						<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
				</tr>
				<tr bgcolor="#D4DDF5">
						<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="1" alt="" border="0"></td>
						<td height=1 colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
						<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="1" alt="" border="0"></td>
				</tr>
				<tr bgcolor="<%=rowbgcolor%>">
					<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="15" alt="" border="0"></td>
					<td style="padding-left : 4px;" class='alt'><b><%=strAktThis%>:</b></td>
					<td width="50" align=right valign="bottom" style="padding-left : 4px; padding-right : 1px;" class='alt'><%
					'call aktprMedTot()
					fs = 0
					select case faktim(a)
					case "1", "2" 
					strSQL = "SELECT DISTINCT(tid), timer AS Ztimer, tdato, Tfaktim, TAktivitetId, TimePris, fastpris, Taar FROM timer WHERE Tmnr = "&medarbID(t)&" AND TAktivitetId = "& aktivitetID(a) & yearKri &""
					case else
					strSQL = "SELECT DISTINCT(tid), timer AS Ztimer, tdato, Tfaktim, TAktivitetId, TimePris, timer.fastpris AS fastpris, Taar FROM timer WHERE Tmnr = "&medarbID(t)&" AND Tjobnr = "& jobnr(a) & yearKri &""
					end select
						
						oRec2.open strSQL, oConn, 0, 1 'adOpenForwardOnly, adLockReadOnly
						while not oRec2.EOF
							
							if weekselector <> "j" then 
								if datepart("ww", oRec2("tdato")) => weeknumbers(0) AND datepart("ww", oRec2("tdato")) <= weeknumbers(w-1) then 
							 	intZtimer = intZtimer + csng(oRec2("Ztimer"))
									if oRec2("Tfaktim") = "1" then
										if oRec2("fastpris") = "1" then
										omsAktMed = omsAktMed + ((jobTpris(a)/budgetTimer(a)) * csng(oRec2("Ztimer")))
										else
											if oRec2("Tfaktim") = "1" then
											omsAktMed = omsAktMed + (csng(oRec2("Ztimer")) * ccur(oRec2("TimePris")))
											else
											omsAktMed = omsAktMed + (csng(oRec2("Ztimer")) * 0)
											end if
										end if
									else
									omsAktMed = 0
									end if
								end if
							else
								if oRec2("tdato") => strWSelStartDato AND oRec2("tdato") <= strWSelEndDato then
								intZtimer = intZtimer + csng(oRec2("Ztimer"))
									if oRec2("Tfaktim") = "1" then
										if oRec2("fastpris") = "1" then
										omsAktMed = omsAktMed + ((jobTpris(a)/budgetTimer(a)) * csng(oRec2("Ztimer")))
										else
											if oRec2("Tfaktim") = "1" then
											omsAktMed = omsAktMed + (csng(oRec2("Ztimer")) * ccur(oRec2("TimePris")))
											else
											omsAktMed = omsAktMed + (csng(oRec2("Ztimer")) * 0)
											end if
										end if
									else
									omsAktMed = 0
									end if
								end if			
							end if
						
						
						tfaktimerne = oRec2("Tfaktim")
						if tfaktimerne = 1 then
							if oRec2("fastpris") = 1 then
								if request("print") <> "j" then
								fonten = "#FFCC00"
								else
								fonten = "#000000"
								end if
							else
								if request("print") <> "j" then
								fonten = "#FFFFFF"
								else
								fonten = "#000000"
								end if
							end if
						else
								if request("print") <> "j" then
								fonten = "#CCCCCC"
								else
								fonten = "#000000"
								end if
						end if
						
						oRec2.movenext
						wend
						oRec2.close
					
					Response.write SQLBless(intZtimer) 
					intZtimer = 0
					%></td>
					<td valign="bottom" align="right" style="padding-left : 4px;  padding-right : 1px;"><font color="<%=fonten%>"><%
					Response.write kommaFunc(omsAktMed) & "&nbsp;dkr."
					%>
					</td>
					<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="15" alt="" border="0"></td>
				</tr>
				<tr bgcolor="#cccccc">
					<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="2" alt="" border="0"></td>
					<td height=2 colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="2" alt="" border="0"></td>
				</tr>
				<%omsAktMed = 0
			next
		next%>
		<tr bgcolor="#5582D2">
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="10" alt="" border="0"></td>
			<td height=10 colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
			<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="10" alt="" border="0"></td>
		</tr>
		</table><br><br><br>
		<!-- slut højreside div -->
		</div>
		<%
		Set oRec = Nothing
		Set oRec2 = Nothing
		oConn.close
		Set oConn = Nothing
	
	
	'***** Farvekode tabel *****
	call farvekoder 
	
	else '*** if week = 0 
	%>
	<div id="ingenregi" style="position:absolute; left:-110; top:100; width:60%; height:600; visibility:visible; z-index:100;">
	<table width="480" celpadding="2" cellspacing="1" border="0">
	<tr><td height="50" align="center" valign="middle"><img src="../ill/alert.gif" width="44" height="45" alt="" border="0">&nbsp;&nbsp;Ingen registreringer i den valgte periode.<br><b><%=noJobMessage%></b></td></tr></table>
	</div>
	<%
	end if
	
	else '*** yearsel = j
	%>
	<img src="../ill/header_joblog_z.gif" alt="" border="0" width="758" height="45" style="position:absolute; left:190; top:55; visibility:visible; z-index:100;">
	<div id="yearselj" style="position:absolute; left:225; top:180; visibility:visible; z-index:100;">
	<table width="480" celpadding="2" cellspacing="1" border="0">
	<tr><td height="50" align="center" valign="middle"><img src="../ill/info.gif" width="42" height="38" alt="" border="0">&nbsp;&nbsp;Klik på et af månederne ovenfor for at få vist timefordelingen...</td></tr></table>
	</div>
	<!--#include file="inc/stat_submenu.asp"-->
	<%
	end if
	
end if 
%>
<!--#include file="../inc/regular/footer_inc.asp"-->
<script>
document.all["loading"].style.visibility = "hidden";
</script>