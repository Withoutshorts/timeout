<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	%>
	<script>
	function showFak() {
	document.all["faktura"].style.visibility = "visible";
	document.all["statknap"].style.visibility = "visible";
	}
	function showStat() {
	document.all["faktura"].style.visibility = "hidden";
	document.all["statknap"].style.visibility = "hidden";
	}
	</script>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	
	<div id="sindhold" style="position:absolute; left:170; top:90; width:60%; height:600; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
    <td valign="top"><font class="pageheader">Statistik</font><br><br>
	<!-------------------------------Sideindhold------------------------------------->
<%
func = request("func")
varMedArbNr = request("FM_medarb") 
varJob = request("FM_job")
varYear = request("FM_aar")
%>
<table cellspacing="0" cellpadding="2" border="0" width="500" bordercolorDark="#C1E1d4">
<form action="stat.asp?menu=stat&func=sel" method="post"><tr>
	<td>Medarbejder(e):&nbsp;<br><select name="FM_medarb" style="font-size:10px;">
	<%if level < 3 then%>
	<option value="0">Alle</option>
	<%end if%>
	<%
	strSQL = "SELECT Mid, Mnavn, Mnr FROM medarbejdere ORDER BY Mnavn"
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
		if func = "sel" then
		usr = varMedArbNr
			if trim(oRec("Mid")) = usr then
			varSel = "SELECTED"
			else
			varSel = ""		
			end if
		else
		usr = Session("user")
			if trim(oRec("Mnavn")) = usr then
			varSel = "SELECTED"
			else
			varSel = ""		
			end if
		end if
	
	if level < 3 then
	showallusr = "1" 
	else
		if cint(oRec("Mid")) = cint(session("Mid")) then
		showallusr = "2"
		else
		showallusr = "0"
		end if
	end if
	
	select case showallusr 
	case "1", "2"
	%>
	<option value="<%=oRec("Mid")%>" <%=varSel%>><%=oRec("Mnavn")%></option>
	<%
	end select
	oRec.movenext
	wend
	oRec.close
	%></select></td>
	
	
	<td>Job:&nbsp;<br><select name="FM_job" style="font-size:10px;">
	<option value="0">Alle</option>
	<%
	strSQL = "SELECT id, jobnr, jobnavn, jobknr, Kkundenavn, kid FROM job, kunder WHERE kid = jobknr ORDER BY Jobnr DESC"
	oRec.open strSQL, oConn, 3
	
	while not oRec.EOF
	
	if func = "sel" then
	jobsel = varJob
			if trim(oRec("id")) = jobsel then
			varSel = "SELECTED"
			else
			varSel = ""		
			end if
	%>
	<option value="<%=oRec("id")%>" <%=varSel%>><%=oRec("jobnr")%>&nbsp;&nbsp;<%=oRec("jobnavn")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=oRec("Kkundenavn")%></option>
	<%else%>
	<option value="<%=oRec("id")%>"><%=oRec("jobnr")%>&nbsp;&nbsp;<%=oRec("jobnavn")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=oRec("Kkundenavn")%></option>
	<%
	end if
	
	oRec.movenext
	wend
	oRec.close
	%>
</select></td>
	<td>
	<input type="hidden" name="FM_aar" value="<%=year(now)%>">
	<!--År:&nbsp;<br><select name="FM_aar" style="font-size:10px;">
	<if func = "sel" then%>
	<option value="<=varYear%>" SELECTED><=varYear%></option>
	<else%>
	<option value="<=year(now)%>" SELECTED><=year(now)%></option>
	<end if%>
	<option value="2002">2002</option>
	<option value="2003">2003</option>
	<option value="2004">2004</option>
	<option value="2005">2005</option>
	<option value="2006">2006</option>
	<option value="2007">2007</option>
</select>-->&nbsp;</td>
	<td valign="bottom"><input type="submit" value="Se statistik!" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid; cursor:hand; width:116; font-size:10px;"></td>
</tr></form>
</table>
<%
if func = "sel" then


'*** Finder ud af hvilke timer der allerede er faktureret ****
	if varJobNr <> "0" then
		strSQL = "SELECT fakDato, faknr, beloeb, Fid FROM fakturaer WHERE fakturaer.jobid=" & varJob &" ORDER BY fakDato"
		oRec.open strSQL, oConn, 3
		
		while not oRec.EOF 
		lastFakdag = oRec("fakDato")
		foundone = "j"
		timerFak = timerFak + oRec("beloeb")
		oRec.movenext
		wend
		oRec.close
		
		if foundone <> "j" then
		lastFakdag = "1/1/00"
		timerFak = 0
		end if
		
		
	end if
	'*** faktureret ****



z = 0
x = 0
aktiviteter = ""
strMedarbejdere = ""
erMedarb = 0
erMedakt = 0
erFaktimer = 0
intomsFak = 0
intomsNotFak = 0

strSQL = "SELECT Mid, Mnavn FROM medarbejdere WHERE Mid=" & varMedArbNr 
oRec.open strSQL, oConn, 3

if not oRec.EOF then
varMedarbNavn = oRec("Mnavn")
else
varMedarbNavn = "Alle"
end if

oRec.close

strSQL = "SELECT id, jobnavn, jobnr, jobTpris FROM job WHERE id =" & varJob
oRec.open strSQL, oConn, 3

if not oRec.EOF then 
varJobNavn = oRec("jobnavn")
varJobNr = oRec("jobnr")
varBudget = oRec("jobTpris")
else
varJobNavn = "Alle"
varJobNr = 0
varBudget = 0
end if

oRec.close


	if varMedarbNavn = "Alle" then
		if varJobNr = "0" then
	 	strSQL = "SELECT Tid, Tdato, Timer, Tfaktim, Tjobnr, Tjobnavn, Taar, TAktivitetId, TAktivitetNavn, Tmnr, Tmnavn, timepris FROM timer WHERE Taar = "& varYear &" ORDER BY Tdato"
		else
		strSQL = "SELECT Tid, Tdato, Timer, Tfaktim, Tjobnr, Tjobnavn, Taar, TAktivitetId, TAktivitetNavn, Tmnr, Tmnavn, timepris FROM timer WHERE Tjobnr = "& varJobNr &" AND Taar = "& varYear &" ORDER BY Tdato"
		end if
	else
		if varJobNr = "0" then
	 	strSQL = "SELECT Tid, Tdato, Timer, Tfaktim, Tjobnr, Tjobnavn, Taar, TAktivitetId, TAktivitetNavn, Tmnr, Tmnavn, timepris FROM timer WHERE Tmnr = "& varMedArbNr &" AND Taar = "& varYear &" ORDER BY Tdato"
		else
		strSQL = "SELECT Tid, Tdato, Timer, Tfaktim, Tjobnr, Tjobnavn, Taar, TAktivitetId, TAktivitetNavn, Tmnr, Tmnavn, timepris FROM timer WHERE Tmnr = "& varMedArbNr &" AND Tjobnr = "& varJobNr &" AND Taar = "& varYear &" ORDER BY Tdato"
		end if
	end if
	
	oRec.Open strSQL, oConn, 3
	
	intFakbareTimer = 0
	totTimer = 0
	antalarbDage = 0
	lastDato_start = "1/1/2002"
	
	Redim aktivitet(x)
	Redim aktivitetID(x)
	Redim maNavn(x)
	Redim medarbID(x)
	
	Redim timers(12)
	timers(1) = 0
	timers(2) = 0
	timers(3) = 0
	timers(4) = 0
	timers(5) = 0
	timers(6) = 0
	timers(7) = 0
	timers(8) = 0
	timers(9) = 0
	timers(10) = 0
	timers(11) = 0
	timers(12) = 0
	
	Redim fakTimer(12)
	fakTimer(1) = 0
	fakTimer(2) = 0
	fakTimer(3) = 0
	fakTimer(4) = 0
	fakTimer(5) = 0
	fakTimer(6) = 0
	fakTimer(7) = 0
	fakTimer(8) = 0
	fakTimer(9) = 0
	fakTimer(10) = 0
	fakTimer(11) = 0
	fakTimer(12) = 0
	
	Redim lastDato(12)
	lastDato(1) = 0
	lastDato(2) = 0
	lastDato(3) = 0
	lastDato(4) = 0
	lastDato(5) = 0
	lastDato(6) = 0
	lastDato(7) = 0
	lastDato(8) = 0
	lastDato(9) = 0
	lastDato(10) = 0
	lastDato(11) = 0
	lastDato(12) = 0
	
	Redim antalReg(12)
	antalReg(1) = 0
	antalReg(2) = 0
	antalReg(3) = 0
	antalReg(4) = 0
	antalReg(5) = 0
	antalReg(6) = 0
	antalReg(7) = 0
	antalReg(8) = 0
	antalReg(9) = 0
	antalReg(10) = 0
	antalReg(11) = 0
	antalReg(12) = 0
	
	while not oRec.EOF 
	showFak = oRec("Tfaktim")
		
		if oRec("Tfaktim") = 1 then
		intFakbareTimer = intFakbareTimer + oRec("timer")
			
			'** Hvilke timer er allerede faktureret.
			if oRec("Tdato") <= lastFakdag AND lastFakdag <> "1/1/00" then
			erFaktimer = erFaktimer + oRec("timer")
			intomsFak = intomsFak + (oRec("timer") * oRec("timepris"))  
			else
			intomsNotFak = intomsNotFak + (oRec("timer") * oRec("timepris")) 
			end if
			
		else
		intFakbareTimer = intFakbareTimer
		end if
		
		totTimer = totTimer + oRec("timer")
		'Slut Årstotal
		
	StrTdato = oRec("Tdato")
	'Response.write StrTdato & "<br>"
	
	' Hvis datoformat er Long Date e.lign.
	justerLangDatoT = instr(StrTdato, "200")
	if justerLangDatoT <> 0 then
	StrTdato_end = right(StrTdato, 2)
	StrTdato_start = left(StrTdato, (justerLangDatoT - 1)) 
	StrTdato = StrTdato_start & StrTdato_end
	end if
	
	if len(StrTdato) > 7 then
		strMrd = left(StrTdato, 2)
		strDag  = mid(StrTdato,4,2)
		strAar = Right(StrTdato,2)
	else
		if len(StrTdato) = 6 then
		strMrd = left(StrTdato, 1)
		strDag  = mid(StrTdato,3,1)
		strAar = Right(StrTdato,2)
		else
			varStregMrd = instr(StrTdato, "/")
			'Response.write varStregMrd
			if varStregMrd < 3 then 
			strMrd = left(StrTdato, 1)
			strDag  = mid(StrTdato,3,2)
			strAar = Right(StrTdato,2)
			'Response.write "H"
			else
			strMrd = left(StrTdato, 2)
			strDag  = mid(StrTdato,4,1)
			strAar = Right(StrTdato,2)
			end if
		end if
	end if
	
	
	aktId = oRec("TAktivitetId")
	aktNavn = oRec("TAktivitetNavn")
	
	erMedakt = instr(aktiviteter, aktId)
	
		if erMedakt = "0" then
		
			Redim preserve aktivitet(x)
			Redim preserve aktivitetID(x)
			aktiviteter = aktiviteter &", "& aktId
			aktivitet(x) = aktNavn
			aktivitetID(x) = aktId
			x = x + 1
		end if
	
	
	medarbNavn = oRec("TMnavn")
	medarbNr = oRec("TMnr")
	
	erMedarb = instr(strMedarbejdere, medarbNr)
	
		if erMedarb = "0" then
			Redim preserve maNavn(z)
			Redim preserve medarbID(z)
			strMedarbejdere = strMedarbejdere &", "& medarbNr
			maNavn(z) = medarbNavn
			medarbID(z) = medarbNr
			z = z + 1
		end if
	
	call beregnTimer(strMrd, oRec("Timer"), oRec("Tfaktim"), oRec("Tdato")) 
	
	if oRec("Tdato") <> lastDato_start then
	lastDato_start = oRec("Tdato")
	antalarbDage = antalarbDage + 1
	else
	antalarbDage = antalarbDage
	end if
	
	intTjekJobnr = oRec("Tjobnr")
	oRec.movenext
	wend
	
	if len(trim(intTjekJobnr)) = 0 then
	intTjekJobnr = 0
	end if
	
	strSQL = "SELECT fakturerbart FROM job WHERE jobnr=" & intTjekJobnr
	oRec3.open strSQL, oConn, 3
	if not oRec3.EOF then
	EksternJob = oRec3("fakturerbart")
	end if
	oRec3.close
	
	'** Tjekker at ingen registreringer er = 0 da der ikke kan dividere med 0 Nedenfor.
	if antalReg(1) = "0" then 
	antalReg(1) = 1
	end if
	
	if antalReg(2) = "0" then
	antalReg(2) = 1
	end if
	
	if antalReg(3) = "0" then
	antalReg(3) = 1
	end if
	
	if antalReg(4) = "0" then
	antalReg(4) = 1
	end if
	
	if antalReg(5) = "0" then
	antalReg(5) = 1
	end if
	
	if antalReg(6) = "0" then
	antalReg(6) = 1
	end if
	
	if antalReg(7) = "0" then
	antalReg(7) = 1
	end if
	
	if antalReg(8) = "0" then
	antalReg(8) = 1
	end if
	
	if antalReg(9) = "0" then
	antalReg(9) = 1
	end if
	
	if antalReg(10) = "0" then
	antalReg(10) = 1
	end if
	
	if antalReg(11) = "0" then
	antalReg(11) = 1
	end if
	
	if antalReg(12) = "0" then
	antalReg(12) = 1
	end if
	
	if intOms = "0" then
	intOms = 0
	end if
	
	if intOmsFak = "0" then
	intOmsFak = 0
	end if
	
	
	function beregnTimer(strMrd, antTimer, antFakTimer, dato)
		
		timers(strMrd) = timers(strMrd) + antTimer
		if antFakTimer = 1 then
		fakTimer(strMrd) = fakTimer(strMrd) + antTimer
		else
		fakTimer(strMrd) = fakTimer(strMrd)
		end if
		
		if dato <> lastDato(strMrd) then
		lastDato(strMrd) = dato
		antalReg(strMrd) = antalReg(strMrd) + 1
		else
		antalReg(strMrd) = antalReg(strMrd)
		end if
		
	end function
	%>
<br><br><br>
<%if z > 0 then%>
<div style="position:absolute; top:90; left:0; width:75; height:15;"><table><tr><td>Gennemsnit:</td><td style='!border: 1px; background-color: DarkRed; border-color: #000000; border-style: solid; padding-left : 4px;padding-right : 4px;'><img src="ill/blank.gif" width="10" height="1" alt="" border="0"></td></tr></table></div>
<div style="position:absolute; top:90; left:98; width:25; height:15;"><table><tr><td>Eksterne:</td><td style='!border: 1px; background-color: #d2691e; border-color: #000000; border-style: solid; padding-left : 4px;padding-right : 4px;'><img src="ill/blank.gif" width="10" height="1" alt="" border="0"></td></tr></table></div>
<div style="position:absolute; top:90; left:180; width:25; height:15;"><table><tr><td>Interne:</td><td style='!border: 1px; background-color: #add8e6; border-color: #000000; border-style: solid; padding-left : 4px;padding-right : 4px;'><img src="ill/blank.gif" width="10" height="1" alt="" border="0"></td></tr></table></div>
<div style="position:absolute; top:90; left:240; width:95; height:15; background-color:#d3d3d3; visibility:hidden;">&nbsp;&nbsp;Total pr. mrd.&nbsp;&nbsp;</div>

<div style="position:absolute; top:313; left:43; width:380; height:1px; z-index:100;"><img src="ill/scalag2.gif" width="385" height="1" alt="" border="0"></div>
<div style="position:absolute; top:344; left:43; width:380; height:1px; z-index:100;"><img src="ill/scalag2.gif" width="385" height="1" alt="" border="0"></div>
<%
if level < 3 then%>
	<div style="position:absolute; top:0; left:440; width:33; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:3px;">
	<a href="joblog.asp?menu=stat&jobnr=<%=varJobNr%>&eks=<%=EksternJob%>&lastFakdag=<%=lastFakdag%>&FM_medarb=<%=varMedArbNr%>&FM_job=<%=varJob%>&FM_aar=<%=varYear%>&selmedarb=<%=varMedArbNr%>"><img src="../ill/joblog.gif" alt="Joblog" border="0" width="23" height="22"></a>
	</div>
	<div style="position:absolute; top:0; left:475; width:33; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:3px;">
	<a href="joblog_z.asp?menu=stat&jobnr=<%=varJobNr%>&eks=<%=EksternJob%>&lastFakdag=<%=lastFakdag%>&FM_medarb=<%=varMedArbNr%>&FM_job=<%=varJob%>&FM_aar=<%=varYear%>&selmedarb=<%=varMedArbNr%>"><img src="../ill/sum.gif" alt="Joblog" border="0" width="23" height="22"></a>
	</div>
<%end if%>

<div style="position:absolute; top:0; left:510; width:32; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:3px;">
<a href="Javascript:window.print()"><img src="../ill/small_print.gif" width="23" height="22" alt="" border="0"></a>
</div>

&nbsp;&nbsp;Arbejdstimer pr. dag:<br>
<table cellspacing="0" cellpadding="0" border="1" bordercolorDark="darkRed">
<tr bgcolor="#ffffff">
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jan</td>
   	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Feb</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Mar</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Apr</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Maj</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jun</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jul</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Aug</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Sep</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Okt</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Nov</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Dec</td>
</tr>
<tr>
    <td width="40" align="right" bgcolor="#ffffff"><img src="ill/scala.gif" width="20" height="240" alt="" border="0"></td>
    <td valign="bottom" width="30"><img src="ill/scalag1.gif" width="10" height="<%=(10*(left(Timers(1),4)/antalReg(1))/z)%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=(10*(left(fakTimer(1),4)/antalReg(1))/z)%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=(10*(left((Timers(1) - fakTimer(1)),4)/antalReg(1))/z)%>" alt="" border="0"></td>
	<td valign="bottom" width="30"><img src="ill/scalag1.gif" width="10" height="<%=(10*(left(Timers(2),4)/antalReg(2))/z)%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=(10*(left(fakTimer(2),4)/antalReg(2))/z)%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=(10*(left((Timers(2) - fakTimer(2)),4)/antalReg(2))/z)%>" alt="" border="0"></td>
	<td valign="bottom" width="30"><img src="ill/scalag1.gif" width="10" height="<%=(10*(left(Timers(3),4)/antalReg(3))/z)%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=(10*(left(fakTimer(3),4)/antalReg(3))/z)%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=(10*(left((Timers(3) - fakTimer(3)),4)/antalReg(3))/z)%>" alt="" border="0"></td>
	<td valign="bottom" width="30"><img src="ill/scalag1.gif" width="10" height="<%=(10*(left(Timers(4),4)/antalReg(4))/z)%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=(10*(left(fakTimer(4),4)/antalReg(4))/z)%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=(10*(left((Timers(4) - fakTimer(4)),4)/antalReg(4))/z)%>" alt="" border="0"></td>
	<td valign="bottom" width="30"><img src="ill/scalag1.gif" width="10" height="<%=(10*(left(Timers(5),4)/antalReg(5))/z)%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=(10*(left(fakTimer(5),4)/antalReg(5))/z)%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=(10*(left((Timers(5) - fakTimer(5)),4)/antalReg(5))/z)%>" alt="" border="0"></td>
	<td valign="bottom" width="30"><img src="ill/scalag1.gif" width="10" height="<%=(10*(left(Timers(6),4)/antalReg(6))/z)%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=(10*(left(fakTimer(6),4)/antalReg(6))/z)%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=(10*(left((Timers(6) - fakTimer(6)),4)/antalReg(6))/z)%>" alt="" border="0"></td>
	<td valign="bottom" width="30"><img src="ill/scalag1.gif" width="10" height="<%=(10*(left(Timers(7),4)/antalReg(7))/z)%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=(10*(left(fakTimer(7),4)/antalReg(7))/z)%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=(10*(left((Timers(7) - fakTimer(7)),4)/antalReg(7))/z)%>" alt="" border="0"></td>
	<td valign="bottom" width="30"><img src="ill/scalag1.gif" width="10" height="<%=(10*(left(Timers(8),4)/antalReg(8))/z)%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=(10*(left(fakTimer(8),4)/antalReg(8))/z)%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=(10*(left((Timers(8) - fakTimer(8)),4)/antalReg(8))/z)%>" alt="" border="0"></td>
	<td valign="bottom" width="30"><img src="ill/scalag1.gif" width="10" height="<%=(10*(left(Timers(9),4)/antalReg(9))/z)%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=(10*(left(fakTimer(9),4)/antalReg(9))/z)%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=(10*(left((Timers(9) - fakTimer(9)),4)/antalReg(9))/z)%>" alt="" border="0"></td>
	<td valign="bottom" width="30"><img src="ill/scalag1.gif" width="10" height="<%=(10*(left(Timers(10),4)/antalReg(10))/z)%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=(10*(left(fakTimer(10),4)/antalReg(10))/z)%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=(10*(left((Timers(10) - fakTimer(10)),4)/antalReg(10))/z)%>" alt="" border="0"></td>
	<td valign="bottom" width="30"><img src="ill/scalag1.gif" width="10" height="<%=(10*(left(Timers(11),4)/antalReg(11))/z)%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=(10*(left(fakTimer(11),4)/antalReg(11))/z)%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=(10*(left((Timers(11) - fakTimer(11)),4)/antalReg(11))/z)%>" alt="" border="0"></td>
	<td valign="bottom" width="30"><img src="ill/scalag1.gif" width="10" height="<%=(10*(left(Timers(12),4)/antalReg(12))/z)%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=(10*(left(fakTimer(12),4)/antalReg(12))/z)%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=(10*(left((Timers(12) - fakTimer(12)),4)/antalReg(12))/z)%>" alt="" border="0"></td> 
</tr>
<tr bgcolor="#ffffff">
    <td width="40" bgcolor="#d2691e" align="center"><font size="1"><%=intFakbareTimer%></td>
    <td valign="bottom" width="30" align="center"><font size="1"><%=fakTimer(1)%></font></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=fakTimer(2)%></font></td> 
	<td valign="bottom" width="30" align="center"><font size="1"><%=fakTimer(3)%></font></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=fakTimer(4)%></font></td> 
	<td valign="bottom" width="30" align="center"><font size="1"><%=fakTimer(5)%></font></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=fakTimer(6)%></font></td> 
	<td valign="bottom" width="30" align="center"><font size="1"><%=fakTimer(7)%></font></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=fakTimer(8)%></font></td> 
	<td valign="bottom" width="30" align="center"><font size="1"><%=fakTimer(9)%></font></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=fakTimer(10)%></font></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=fakTimer(11)%></font></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=fakTimer(12)%></font></td>  
</tr>
<tr bgcolor="#ffffff">
    <td width="40" bgcolor="#add8e6" align="center"><font size="1"><%=(totTimer - intFakbareTimer)%></td>
    <td valign="bottom" width="30" align="center"><font size="1"><%=left((Timers(1)-fakTimer(1)), 4)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=left((Timers(2)-fakTimer(2)), 4)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=left((Timers(3)-fakTimer(3)), 4)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=left((Timers(4)-fakTimer(4)), 4)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=left((Timers(5)-fakTimer(5)), 4)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=left((Timers(6)-fakTimer(6)), 4)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=left((Timers(7)-fakTimer(7)), 4)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=left((Timers(8)-fakTimer(8)), 4)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=left((Timers(9)-fakTimer(9)), 4)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=left((Timers(10)-fakTimer(10)), 4)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=left((Timers(11)-fakTimer(11)), 4)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=left((Timers(12)-fakTimer(12)), 4)%></td>
</tr>
<tr bgcolor="#ffffff">
    <td width="40" bgcolor="DarkRed" align="center"><font size="1" color="#ffffff"><%=totTimer%></td>
    <td valign="bottom" width="30" align="center"><font size="1"><%=timers(1)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=timers(2)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=timers(3)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=timers(4)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=timers(5)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=timers(6)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=timers(7)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=timers(8)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=timers(9)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=timers(10)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=timers(11)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=timers(12)%></td>
</tr>
</table>
<br>
<table cellspacing="0" cellpadding="2" border="0" width="516" bordercolorDark="#C1E1d4">
<tr>
	<td colspan="6">Antal dage med registreringer: <b><%=antalarbDage%></b></td>
</table>
<br>
<%if varJobNr <> "0" then%>
<table cellspacing="0" cellpadding="2" border="0" width="510">
<tr><td valign="top">
<table cellspacing="0" cellpadding="2" border="1" width="305" bordercolorDark="#C1E1d4">
<tr bgcolor="#ffffff">
	<td width=2>&nbsp;</td><td><b>Aktiviteter</b></td><td><b>Timer</b></td>
</tr>
	<%
	oRec.close
	
	y = x - 1
	x = 0
	for x = 0 to y
	if varMedarbNavn = "Alle" then
	strSQL = "SELECT sum(timer) AS sumtimer FROM timer WHERE TaktivitetId = " & aktivitetID(x) 
	else
	strSQL = "SELECT sum(timer) AS sumtimer FROM timer WHERE Tmnr = "& varMedArbNr &" AND TaktivitetId = " & aktivitetID(x) 
	end if
	
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
		
			strSQL = "SELECT Tfaktim, TAktivitetId FROM timer WHERE TaktivitetId = " & aktivitetID(x) 
			oRec2.open strSQL, oConn, 3
			if not oRec2.EOF then
					if oRec2("Tfaktim") = "1" then
					bc = "#d2691e"
					else
					bc = "#add8e6"
					end if
					
					if left(oRec2("TAktivitetId"), 3) = "909" then
					ekst = 0
					else
					ekst = 1
					end if
			Response.write "<tr><td bgcolor='"&bc&"'>&nbsp;</td><td><a href='joblog.asp?menu=stat&selmedarb="&varMedArbNr&"&selaktid="&aktivitetID(x)&"&jobnr="&varJobNr&"&eks="&ekst&"&lastFakdag="&lastFakdag&"&FM_medarb="&varMedArbNr&"&FM_job="&varJob&"&FM_aar="&varYear&"'><img src='../ill/joblog_small.gif' width='11' height='12' alt='Joblog' border='0'>&nbsp;&nbsp;" & aktivitet(x) &"</a></td><td>" & oRec("sumtimer") & "</td></tr>" 
			end if
			oRec2.close
	end if
	oRec.close
	next
	%>
</table>
</td><td valign="top">
<table cellspacing="0" cellpadding="2" border="1" width="205" bordercolorDark="#C1E1d4">
<tr bgcolor="#ffffff">
	<td><b>Medarbejdere</b></td><td><b>Timer</b></td>
</tr>
	<%
	t = 0
	v = z - 1
	z = 0
	for z = 0 to v
	strKriterie = ""
		
		for t = 0 to x - 1
		strKriterie = strKriterie & " Tmnr = " & medarbID(z) & " AND TaktivitetId = " & aktivitetID(t) & " OR "
		intmedarbnr = medarbID(z)
		next
		
		intAntalchar = len(strKriterie)
		strKriterie = left(strKriterie, intAntalchar - 4) 
		
	strSQL = "SELECT sum(timer) AS sumtimer FROM timer WHERE " & strKriterie
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
		if varJob <> "0" then
		Response.write "<tr bgcolor=#d3d3d3><td><a href='joblog.asp?menu=stat&selmedarb="&intmedarbnr&"&jobnr="&varJobNr&"&eks="&ekst&"&lastFakdag="&lastFakdag&"&FM_medarb="&varMedArbNr&"&FM_job="&varJob&"&FM_aar="&varYear&"'><img src='../ill/joblog_small.gif' width='12' height='12' alt='Joblog' border='0'>&nbsp;&nbsp;" & maNavn(z) &"</a></td><td>" & oRec("sumtimer") & "</td></tr>" 
		else
		Response.write "<tr bgcolor=#d3d3d3><td>" & maNavn(z) &"</td><td>" & oRec("sumtimer") & "</td></tr>" 
		end if
	end if
	oRec.close
	next
	%>
</table>
</td></tr>
</table>
<%else%>
<font size="1">Note: Vælg job, for at få oversigt over aktiviteter og timeforbrug pr. medarbjeder.</font>
<%end if%>
<br>
<br>

<%if varJob <> "0"  AND EksternJob = "1" then 'AND varMedArbNr = "0"%>
<div id="fakturaknap" style="position:absolute; top:0; left:348; width:90; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:5px;">
<a href="#" onclick="showFak();">Fakturering&nbsp;<img src="../ill/pil_lysblaa.gif" width="11" height="6" alt="" border="0"></a>
</div>
<div id="statknap" style="position:absolute; top:0; left:348; width:90; height:30px; z-index:100; !border: 1px; visibility:hidden; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:5px;">
<a href="#" onclick="showStat();"><img src="../ill/pil_lysblaa_left.gif" width="11" height="6" alt="" border="0">&nbsp;Statistik</a>
</div>
<div id="faktura" style="position:absolute; left:0; top:85; width:60%; height:600; visibility:hidden; z-index:200; background-color:#fff8dc;">
	
<%
intTimerTilFak = (intFakbareTimer - erFaktimer)
%>

<a href="fak.asp?menu=stat&func=opr&jobid=<%=varJob%>&jobnr=<%=varJobNr%>&ttf=<%=intTimerTilFak%>&ktf=<%=intomsNotFak%>">Indsæt ny faktura&nbsp;&nbsp;<img src="../ill/pil.gif" width="11" height="11" alt="" border="0"></a><br><br>
<table cellspacing="0" cellpadding="2" border="1" width="515" bordercolorDark="#C1E1d4">
<tr bgcolor="#ffffff">
<td><b>Faktura</b></td><td><b>Dato</b></td><td><b>Timer</b></td><td align="right"><b>Beløb</b></td>
</tr>
<%
strSQL = "SELECT * FROM fakturaer WHERE jobid = " & varJob
oRec.open strSQL, oConn, 3
while not oRec.EOF 
DKDato = convertDate(oRec("fakDato"))%>
<tr bgcolor="#d3d3d3">
<td><a href="fak.asp?menu=stat&func=red&id=<%=oRec("Fid")%>&jobid=<%=varJob%>&jobnr=<%=varJobNr%>"><%=oRec("faknr")%></a></td><td><%=DKDato%></td><td><%=oRec("timer")%></td><td align="right"><%=oRec("beloeb")%> kr.</td>
</tr>
<%
intFakturaTimer = intFakturaTimer + oRec("timer")
oRec.movenext
wend

if len(trim(intFakturaTimer)) = "0" then 
intFakturaTimer = 0
else
intFakturaTimer = intFakturaTimer
end if%>
<tr>
	<td colspan="2">Total budget</td>
	<td>&nbsp;</td>
	<td align="right" bgcolor=""><font color="darkRed"><%=varBudget%> kr.</font></td>
</tr>
<tr>
	<td colspan="2">Faktureret&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="gray" size="1">(Time registreret beløb: <%=intomsFak%> kr.)</font></td>
	<td>&nbsp;</td>
	<td align="right" bgcolor=""><%=timerFak%> kr.</td>
	
</tr>
<tr>
	<td colspan="2">Til fakturering</td>
	<td>&nbsp;</td>
	<td align="right"><%=intomsNotFak%> kr.</td>
</tr>  
<tr>
	<td colspan="2">Resterende budget</td>
	<td>&nbsp;</td>
	<td align="right"><%=(varBudget - (timerFak + intomsNotFak)) %> kr.</td>
</tr> 
<tr>
	<td colspan="2">Timer der bør være faktureret /fakturede timer</td>
	<td><%=erFaktimer%>/<%=intFakturaTimer%>
	<%if erFaktimer <> intFakturaTimer then
	%>
	<font color="darkRed"><b>!</b></font>
	<%
	end if
	%></td>
	<td>&nbsp;</td>
</tr>
<tr>
	<td colspan="2">Timer til fakturering</td>
	<td><%=intTimerTilFak%></td>
	<td>&nbsp;</td>
</tr>
</table>
</div>
<%end if%>
<br>
<br>
<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="16" height="17" alt="" border="0"></a>
<br>
<br>
</TD>
</TR>
</TABLE>
</div>


<%
'now close and clean up
	
	oConn.close
	Set oRec = Nothing
	set oConn = Nothing 

else
%>
<div id="ingenregi" style="position:absolute; left:5; top:100; width:60%; height:600; visibility:visible; z-index:100;">
<table width="480" celpadding="2" cellspacing="1" border="0">
<tr><td height="50" align="center" valign="middle" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px; padding-right : 4px;">Ingen registreringer for de(n) ønskede medarbejder(e) i den valgte periode.</td></tr></table>
</div>
<%
end if	
end if 
end if

%>
<!--#include file="../inc/regular/footer_inc.asp"-->
