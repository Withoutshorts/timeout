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
	<!--#include file="../inc/regular/header_inc.asp"-->	
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
	document.all["knap_"+i+""].style.borderColor = "lightBlue";
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
	function opbygning()
	'*** Opbygning af oversigt medarbejdere *****
			areMaP = instr(strMedarbPrinted, ", " &oRec("Tmnr")&"#")
			
			if areMaP = "0" then
			Redim preserve medarbejder(t)
			Redim Preserve medarbID(t)
			medarbID(t) = oRec("Tmnr")  
			medarbejder(t) = oRec("Tmnavn")
			t = t + 1
			end if
			
			strMedarbPrinted = strMedarbPrinted & ", " & oRec("Tmnr") &"#" 
			
			'*** Opbygning af oversigt Aktiviteter *****
			areAktP = instr(strAktPrinted, ", " &oRec("TAktivitetId")&"#")
			
			if areAktP = "0" then
			Redim preserve aktivitetID(a)
			Redim preserve aktiviteter(a)
			Redim preserve jobnavn(a)
			Redim preserve jobnr(a)
			aktivitetID(a) = oRec("TAktivitetId")
			aktiviteter(a) = oRec("Anavn")
			jobnavn(a) = oRec("Tjobnavn")
			jobnr(a) = oRec("Tjobnr")
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
			
			
			'Response.write strYearNum &"-" & strWeekNum &"&nbsp;&nbsp;:"
			'*** Opbygning af yearnumbers *****
			areYearP = instr(strYearNumPrinted, ", " & strYearNum &"#")
			if areYearP = "0" then
			y = y + 1
			strWeekNumPrinted = "#0"
			yearweeks(y,week) = strWeekNum
			'Response.write yearweeks(y,week) & "<br>"
			week = week + 1
			else
				if areWeekP = "0" then
				yearweeks(y,week) = strWeekNum
				'Response.write yearweeks(y,week) & "<br>"
				week = week + 1
				end if
			end if
			
			
			strWeekNumPrinted = strWeekNumPrinted & ", " & strWeekNum &"#" 
			strYearNumPrinted = strYearNumPrinted & ", " & strYearNum &"#" 
			'**************  opbygning slut ********************
	end function
	
	
	function weekberegning()
	strSQL = "SELECT timer AS Ztimer, tdato, Tfaktim, TAktivitetId, TimePris FROM timer WHERE Tmnr = "&medarbID(t)&" AND TAktivitetId = "& aktivitetID(a) &""
		oRec2.open strSQL, oConn, 0, 1  'adOpenForwardOnly, adLockReadOnly
		while not oRec2.EOF
		if datepart("ww", oRec2("tdato")) = weeknumbers(w) then  'yearweeks(y,week) then  ''AND datepart("yyyy", oRec2("tdato")) = yearnumbers(0) then 
			intZtimer = intZtimer + csng(oRec2("Ztimer"))
			intZtimerTotalAkt = intZtimerTotalAkt + csng(oRec2("Ztimer"))
			intZtimerTotalWeek = intZtimerTotalWeek + csng(oRec2("Ztimer"))
			intZtimerTotalMonth = intZtimerTotalMonth + csng(oRec2("Ztimer"))
			if oRec2("Tfaktim") = "1" then
			totOms = totOms + (csng(oRec2("Ztimer")) * ccur(oRec2("TimePris")))
			totOmsWeek = totOmsWeek + (csng(oRec2("Ztimer")) * ccur(oRec2("TimePris")))
			totOmsAkt = totOmsAkt + (csng(oRec2("Ztimer")) * ccur(oRec2("TimePris")))
			totOmsTotalM = totOmsTotalM + (csng(oRec2("Ztimer")) * ccur(oRec2("TimePris")))
			else
			totOms = 0
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
			end if
		end if
		oRec2.movenext
		wend
		oRec2.close
	end function
	
	
	
	'***************          Start sideopbygning og sidevariable              ************
	'** Sætter Periode værdier **
	strReqAar_slut = 2025
	
	if request("weekselector") = "j" then 
	StrTdato = request("FM_start_mrd") &"/" & request("FM_start_dag") & "/" & request("FM_start_aar")
	StrUdato = request("FM_slut_mrd") &"/" & request("FM_slut_dag") & "/" & request("FM_slut_aar")
	id = "1"
		if request("FM_start_aar") <> request("FM_slut_aar") then
		strReqAar_slut = request("FM_slut_aar")
		end if
	else
	StrTdato = (month(now()) - 1) &"/1/" & year(now())
	StrUdato = month(now()) &"/" & day(now()) & "/" & year(now())
	id = "1"
	end if
	'**** 
	
	linket = request("linket")
	
	selmedarb = request("selmedarb")
	selaktid = request("selaktid")
	
	intJobnr = request("jobnr")
	if intJobnr = "0" then
	intJobnrKri = ""
	else
	intJobnrKri = "Tjobnr = " & intJobnr &" AND "
	end if
	
	if len(Request("mrd")) <> "0" then
		strReqMrd = Request("mrd")
	else
	strReqMrd = month(now)
	end if
	
	'**** finder ud af om der er valgt et år ***
	if request("weekselector") <> "j" then 
		if len(Request("year")) <> "0" then
			if Request("year") = "-1" then
			strReqAar = "0"
			else
			strReqAar = Request("year")
			end if	
		else
		strReqAar = right(year(now), 2)
		end if
		strReqAar_slut = strReqAar
	else
	strReqAar = request("FM_start_aar")
	end if	
	
	lastFakdag = Request("lastFakdag")
	
	'*** Bruges som markør af mrd knapper **
	if request("weekselector") = "j" then
	framrd = datepart("m", StrTdato)
	tilmrd = datepart("m", StrUdato)
	end if


if request("print") <> "j" then%>

	<!-- mrd knapper -->
	<%
	leftpix = 200
	mrdnum = 1
	for mrdnum = 1 to 12
	
	Select case mrdnum 
	case 1
	mrdnavn = "Jan"
	pixadd = 37
	case 2
	mrdnavn = "Feb"
	pixadd = 39
	case 3
	mrdnavn = "Mar"
	pixadd = 38
	case 4
	mrdnavn = "Apr"
	pixadd = 38
	case 5
	mrdnavn = "Maj"
	pixadd = 37
	case 6
	mrdnavn = "Jun"
	pixadd = 38
	case 7
	mrdnavn = "Jul"
	pixadd = 36
	case 8
	mrdnavn = "Aug"
	pixadd = 41
	case 9
	mrdnavn = "Sep"
	pixadd = 40
	case 10
	mrdnavn = "Okt"
	pixadd = 39
	case 11
	mrdnavn = "Nov"
	pixadd = 40
	case 12
	mrdnavn = "Dec"
	pixadd = 39
	end select
	
		if request("weekselector") = "j" then
			if cint(mrdnum) => cint(framrd) AND cint(mrdnum) <= cint(tilmrd) then
			strBgCol = "#ffffff"
			else
			strBgCol = "#add8e6"
			end if
		else
			if cint(strReqMrd) = cint(mrdnum) then
			strBgCol = "#ffffff"
			else
			strBgCol = "#add8e6"
			end if
		end if
	
	if strReqAar = "0" then
	%>
	<div id="mrdknap_<%=mrdnavn%>" style="position:absolute; top:140; left:<%=leftpix%>; width:35; height:20; background-color:<%=strBgCol%>;!border: 1px; border-color: Black; border-style: solid; border-bottom:0; z-index:200; visibility:visible;">&nbsp;&nbsp;<%=mrdnavn%>&nbsp;&nbsp;</div>
	<%else%>
	<div id="mrdknap_<%=mrdnavn%>" style="position:absolute; top:140; left:<%=leftpix%>; width:35; height:20; background-color:<%=strBgCol%>;!border: 1px; border-color: Black; border-style: solid; border-bottom:0; z-index:200; visibility:visible;">&nbsp;&nbsp;<a href="joblog_z.asp?menu=stat&mrd=<%=mrdnum%>&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>"><%=mrdnavn%></a>&nbsp;&nbsp;</div>
	<%end if
	leftpix = leftpix + pixadd
	next
	%>
	<div id="mrdknap_under" style="position:absolute; top:159; left:170; width:555; height:600; !border: 1px;  border-color: Black; border-style: solid; z-index:0; border-bottom:0; visibility:visible;"></div>
	<%
	minus1 = right((year(now)-1), 2)
	iaar = right(year(now), 2)
	plus1 = right((year(now)+1), 2)
	
	
	Select case strReqAar
	case minus1
	varBgColy1 = "#ffffff"
	varBgColy2 = "#cccccc"
	varBgColy3 = "#cccccc"
	varBgColy4 = "#cccccc"   
	case iaar
	varBgColy1 = "#cccccc"
	varBgColy2 = "#ffffff"
	varBgColy3 = "#cccccc"
	varBgColy4 = "#cccccc"
	case plus1
	varBgColy1 = "#cccccc"
	varBgColy2 = "#cccccc"
	varBgColy3 = "#ffffff"
	varBgColy4 = "#cccccc"
	case "0"
	varBgColy1 = "#cccccc"
	varBgColy2 = "#cccccc"
	varBgColy3 = "#cccccc"
	varBgColy4 = "#ffffff"
	end select
	%>
	<div id="aarknap1" style="position:absolute; top:122; left:210; width:35; height:38; background-color:<%=varBgColy1%>;!border: 1px; border-color: Black; border-style: solid; border-bottom:0; z-index:199;  visibility:visible;">&nbsp;&nbsp;<a href="joblog_z.asp?menu=stat&mrd=0&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=right((year(now)-1), 2)%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>"><%=year(now)-1%></a>&nbsp;&nbsp;</div>
	<div id="aarknap2" style="position:absolute; top:122; left:257; width:35; height:38; background-color:<%=varBgColy2%>;!border: 1px; border-color: Black; border-style: solid; border-bottom:0; z-index:199;  visibility:visible;">&nbsp;&nbsp;<a href="joblog_z.asp?menu=stat&mrd=0&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=right(year(now), 2)%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>"><%=year(now)%></a>&nbsp;&nbsp;</div>
	<div id="aarknap3" style="position:absolute; top:122; left:304; width:35; height:38; background-color:<%=varBgColy3%>;!border: 1px; border-color: Black; border-style: solid; border-bottom:0; z-index:199;  visibility:visible;">&nbsp;&nbsp;<a href="joblog_z.asp?menu=stat&mrd=0&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=right((year(now)+1), 2)%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>"><%=year(now)+1%></a>&nbsp;&nbsp;</div>
	<div id="aarknap4" style="position:absolute; top:142; left:664; width:35; height:18; background-color:<%=varBgColy4%>;!border: 1px; border-color: Black; border-style: solid; border-bottom:0; z-index:198;  visibility:visible;">&nbsp;&nbsp;<a href="joblog_z.asp?menu=stat&mrd=0&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=-1&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>">Total</a>&nbsp;&nbsp;</div>
	
	<!-- slut knapper -->
<%
end if 'print

'*** Finder de medarbejdere der er valgt ***
if len(request("selmedarb")) <> "0" then
	if request("selmedarb") > "0" then
	selmedarbKri = "Tmnr = " & request("selmedarb") &" AND "
	else
	selmedarbKri = "Tmnr <> 0 AND "
	end if
else
selmedarbKri = "Tmnr <> 0 AND "
end if

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
<!-- Flere knapper , Weekselecter -->
<div id="weekselecter" style="position:absolute; top:242; left:10; width:120; height:80; z-index:100; background-color:#fff8dc; visibility:<%=visWeekSel%>;">
<!--#include file="inc/dato2.asp"-->
<%
	if request("weekselector") = "j" then
		frauge = datepart("ww", strMrd & "/" & strDag & "/" & strAar)
		tiluge = datepart("ww", strMrd_slut & "/" & strDag_slut & "/" & strAar_slut)
		showStrTdato = strDag & ".&nbsp;" & strMrdNavn & "&nbsp;" & strAar
		showStrUdato = strDag_slut & ".&nbsp;" & strMrdNavn_slut & "&nbsp;" & strAar_slut
	else
		frauge = 1
		tiluge = 52
	end if
%>
<table cellspacing="0" cellpadding="0" border="0" width="120">
	<tr><form action="joblog_z.asp?menu=stat&weekselector=j&mrd=0&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>" method="post">
   <td bgcolor="#ffffff" colspan="5"><img src="../ill/periode_top_blaa.gif" width="148" height="23" alt="" border="0"></td>
</tr>
<tr bgcolor="lightBlue">
	<td rowspan="3"><img src="../ill/lightblue.gif" width="1" height="90" alt="" border="0"></td>
	<td><font size="1">&nbsp;Fra:<br>&nbsp;<select name="FM_start_dag" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
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
		<option value="31">31</option></select></td>
		<td><br>
		<select name="FM_start_mrd" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
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
	   	<option value="12">dec</option></select></td>
		<td><br>
		<select name="FM_start_aar" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
		<option value="<%=strReqAar%>"><%=strReqAar%></option>
		<option value="02">02</option>
		<option value="03">03</option>
	   	<option value="04">04</option>
	   	<option value="05">05</option>
		<option value="06">06</option>
		<option value="07">07</option></select></td>
		<td rowspan="3" align="left"><img src="../ill/lightblue.gif" width="1" height="90" alt="" border="0"></td>
		</tr>
		<tr bgcolor="lightBlue">
		<td><font size="1">&nbsp;Til:<br>&nbsp;<select name="FM_slut_dag" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
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
		<option value="31">31</option></select></td>
		<td><br>
		<select name="FM_slut_mrd" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
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
	   	<option value="12">dec</option></select></td>
		<td><br>
		<select name="FM_slut_aar" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
		<option value="<%=strReqAar_slut%>"><%=strReqAar_slut%></option>
		<option value="02">02</option>
		<option value="03">03</option>
	   	<option value="04">04</option>
	   	<option value="05">05</option>
			<option value="06">06</option>
				<option value="07">07</option></select></td>
				</tr>
	<tr bgcolor="lightBlue">
	<td align="center" colspan="3"><input type="submit" value="Vælg!" style="background-color : #ffffff; !border: 1px; border-color: #336699; border-style: solid; font: 10px verdana; cursor:hand;"></td>
</tr>
<tr>
    <td colspan="5" bgcolor="#ffffff"><img src="../ill/bund_blaa.gif" width="148" height="6" alt="" border="0"></td>
</tr></form></table>
</div>
<!-- Weekselecter SLUT -->
<%
StrTdato = "1/1/2001"
StrUdato = "12/12/2007"
	
	'** Fakturerbart / Ikke fakturerbart job. **
	strSQL = "SELECT Tdato, TasteDato, Tjobnr, Tjobnavn, TAktivitetNavn AS Anavn, TAktivitetId, Tknavn, Tmnr, Tmnavn, Timer, Tid, Tfaktim, TimePris, Timerkom FROM timer WHERE "& selmedarbKri & intJobnrKri &" Tid > 0 "& selaktidKri &" ORDER BY Tdato DESC, Tmnavn"
	oRec.open strSQL, oConn, 0, 1
	
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
	Redim aktivitetID(a)
	Redim aktiviteter(a)
	Redim jobnavn(a)
	Redim jobnr(a)
	Redim weeknumbers(w)
	'Redim yearnumbers(y)
	Dim yearweeks(7,52)
	
	Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
	
	while not oRec.EOF
		StrTdato = oRec("Tdato")
		strWeekNum = datepart("ww", StrTdato)
		strYearNum = datepart("yyyy", StrTdato)
		id = 1%>
		<!--#include file="inc/dato2.asp"-->
		<%
		if cint(strReqAar) = cint(strReqAar_slut) then
			if cint(strAar) = cint(strReqAar) OR cint(strReqAar) = "0" then
			if cint(strReqMrd) = cint(strMrd) OR cint(strReqMrd) = "0" then
			if datepart("ww", oRec("tdato")) => frauge AND datepart("ww", oRec("tdato")) <= tiluge then 
			
			call opbygning()
			
			end if
			end if
			end if
		else
			Response.write "<br>WON<br>"
			'if cint(strAar) => cint(strReqAar) AND cint(strAar) <= cint(strReqAar_slut) then
			'if datepart("ww", oRec("tdato")) => frauge AND datepart("ww", oRec("tdato")) <= tiluge then 
				call opbygning()
			'end if
			'end if
		end if
	
	
	oRec.movenext
	wend	  
		
	oRec.Close
	Set oRec = Nothing
	
'*************************       Z        *****************************************************
%>
<!-- venstre side div -->
<%if request("print") <> "j" then
leftloading = "280"
else
leftloading = "140"
end if
%>
<div id="loading" name="loading" style="position:absolute; top:300; left:<%=leftloading%>; width:300; height:50; z-index:1000; visibility:visible; background-color:#ffffff; !border: 1px; border-color: DarkRed; border-style: solid; padding-left:4px; padding-top:5px;">
<b>Henter information...</b><br>
Denne handling kan tage op til 20 sek. alt efter din forbindelse.
</div>
<%if request("print") <> "j" then%>		
<div id="z" style="position:absolute; left:180; top:55; width:300; height:600; visibility:visible; z-index:100;">
<div id="logknap" style="position:absolute; top:35; left:304; width:32; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:3px;">
<a href="word.asp?eks=<%=request("eks")%>&jobnr=<%=intJobnr%>&mrd=<%=strReqMrd%>&year=<%=strReqAar%>"><img src="../ill/word.gif" width="23" height="22" alt="" border="0"></a>
</div>
<div style="position:absolute; top:35; left:338; width:90; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:5px;">
<a href="stat.asp?menu=stat&func=sel&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>"><img src="../ill/pil_lysblaa_left.gif" width="11" height="6" alt="" border="0">&nbsp;Statistik</a>
</div>
<div style="position:absolute; top:35; left:430; width:33; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:3px;">
	<a href="joblog.asp?menu=stat&mrd=<%=strReqMrd%>&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>"><img src="../ill/joblog.gif" alt="Joblog" border="0" width="23" height="22"></a>
</div>
<div style="position:absolute; top:35; left:465; width:33; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:3px;">
	<a href="joblog_z.asp?menu=stat&mrd=<%=strReqMrd%>&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>"><img src="../ill/sum.gif" alt="Joblog" border="0" width="23" height="22"></a>
</div>
<div style="position:absolute; top:35; left:500; width:32; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:3px;">
<a href="javascript:NewWin('joblog_z.asp?menu=stat&print=j&mrd=<%=strReqMrd%>&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>')" target="_self"><img src="../ill/small_print.gif" width="23" height="22" alt="" border="0"></a>
</div>
<%else%>
<div id="z" style="position:absolute; left:20; top:15; width:300; height:600; visibility:visible; z-index:100;">
<div style="position:absolute; top:35; left:466; width:32; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:3px;">
<a href="javascript:window.close()"><img src="../ill/luk.gif" width="23" height="22" alt="" border="0"></a>
</div>
<div style="position:absolute; top:35; left:500; width:32; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:3px;">
<a href="javascript:window.print()"><img src="../ill/small_print.gif" width="23" height="22" alt="" border="0"></a>
</div>
<%end if%>
<table cellspacing="0" cellpadding="0" border="0" width="600">
<tr><td><br><br><font class="pageheader">Timefordeling</font><br>
<%if request("weekselector") = "j" then
Response.write showStrTdato &"&nbsp;["& frauge &"]&nbsp;" &" - " & showStrUdato &"&nbsp;["& tiluge &"]&nbsp;"
end if

if request("print") = "j" then
Response.write "20" & strReqAar 
end if%>
<br></td></tr></table>
<!-------------------------------Sideindhold------------------------------------->

<%

'For y = 0 to UBOUND(yearweeks, 1)
	'for week = 0 to week - 1
	'Response.write y &":" & yearweeks(y, week) & "<BR>"
	'next
'Next

'*** opr
'** Hvis der findes uger med registreringer i den valgte periode.**
if w > 0 then
if request("print") <> "j" then
		'** Uge knapper **
		g = (w - 1)
		antalWeeks = w
		varleft = 32
		varTop = 129
		
				for w = 0 to w - 1
				if w = g then
				varBColor = "darkRed"
				else
				varBColor = "Lightblue"
				end if
				%>
				<div id="knap_<%=w%>" name="knap_<%=w%>" style="position:absolute; top:<%=varTop%>; left:<%=varleft%>; width:45; z-index:1000; visibility:visible; background-color:#ffffff; !border: 1px; border-color: <%=varBColor%>;  border-bottom:0; border-style: solid;">
				<table border="0" cellpadding="0" cellspacing="0" width="45">
				<tr><td align="center" height="18" width="45">
				<a href="#" onclick="javascript:showweek('<%=w%>', '<%=antalWeeks%>');"><%=weeknumbers(w)%></a></td></tr></table>
				</div>
				<%
				select case w
				case 3, 7, 11, 15, 19, 23, 27, 31, 35, 39, 43, 47, 51
				varTop = varTop + 18
				varleft = 32 - w
				case else
				varTop = varTop
				varleft = varleft + 49
				end select
				next
		end if


'** Hvis der er valgt et måned **
'if strReqMrd <> 0 then
intZtimerTotalMonth = 0


for w = 0 to w - 1
		
		if w = g then
		varVisi = "visible"
		else
		varVisi = "hidden"
		end if
		varWeekTop = varTop - 32
		
		if request("print") <> "j" then%>
		<div id="weekdiv_<%=w%>" name="weekdiv_<%=w%>" style="position:absolute; top:<%=varWeekTop%>; left:0; z-index:100; visibility:<%=varVisi%>;">
		<%end if%>
		<table border="0" cellpadding="0" cellspacing="0" width="260" style="page-break-before:always;">
		<tr><td height=30 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
		<tr>
			<td colspan="6" height=20 style="padding-left : 4px;">&nbsp;</td>
		</tr>
		<tr>
			<td height=1 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		</tr>
		<%
				for a = 0 to a - 1
				%>
				<tr><td height=2 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
				<tr>
					<td height="20" colspan="2" width="185" style="!border: 1 px; background-color: LightBlue; border-color: silver; border-style: solid; padding-left : 4px;"><font size=1><%= jobnr(a) &"&nbsp;" & jobnavn(a) &"&nbsp;</font><br><b>" & left(aktiviteter(a), 16)%></b></td>
					<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					<td width="50" align=right valign="bottom" style="!border: 1 px; background-color: LightBlue; border-color: silver; border-style: solid; padding-left : 4px;"><font size=1>timer</td>
					<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					<td width="54" align=right valign="bottom" style="!border: 1 px; background-color: LightBlue; border-color: silver; border-style: solid; padding-left : 4px;"><font size=1>omsætning</td>
				</tr>
					<%
					for t = 0 to t - 1 
					%>
					<tr>
						<tr><td height=1 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
						<td colspan=2 style="!border: 1 px; background-color: #FFFFFF; border-color: #d2691e; border-style: solid; padding-left : 4px;"><%=medarbejder(t)%></td>
						<td width=1><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
						<td width="50" align=right style="!border: 1 px; background-color: #FFFFFF; border-color: silver; border-style: solid; padding-left : 4px; padding-right : 1px;">
						<%
						
						call weekberegning()
						
						Response.write csng(intZtimer) &"</td><td width='1'><img src='ill/blank.gif' width='1' height='1' border='0' alt=''></td>"_
						&"<td width=54 align=right style='!border: 1 px; background-color: #FFFFFF; border-color: silver; border-style: solid; padding-left : 4px; padding-right : 1px;'>" & ccur(totOms) &""
						intZtimer = 0
						totOms = 0
						%></td></tr>
					<%
					next
					%>
					<tr><td height=2 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
					<tr><td colspan=2 style="!border: 1 px; background-color: #FFFFFF; border-color: Darkred; border-style: solid; padding-left : 4px;">Ialt: </td>
					<td width=1><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					<td align=right style="!border: 1 px; background-color: #FFFFFF; border-color: Darkred; border-style: solid; padding-left : 4px; padding-right : 1px;"><%
					Response.write intZtimerTotalAkt
					intZtimerTotalAkt = 0%></td><td width=1><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					<td align="right" style='!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px; padding-right : 1px;'><%=ccur(totOmsAkt)%></td>
					</tr>
					<%totOmsAkt = 0
				next
			%>
			<tr><td height=20 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
					<tr><td colspan=2 style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px;"><b>Uge <%=weeknumbers(w)%> total:</b></td>
					<td width=1><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					<td align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px; padding-right : 4px;"><b><%
					Response.write intZtimerTotalWeek
					intZtimerTotalWeek = 0%></b></td><td width=1><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					<td align="right" style='!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px; padding-right : 1px;'><b><%=ccur(totOmsWeek)%></b></td>
					</tr>
				</TABLE><br><br><br>
				<%if request("print") <> "j" then%>
				</div>
				<%end if%>
			<%totOmsWeek = 0
		next
%>
<!-- slut venstre div -->
</div>
<%
'*** Aktiviteter total ****
%>
<!-- Højre side div -->
<%if request("print") <> "j" then%>	
<div id="akttotal" style="position:absolute; left:450; top:188; visibility:visible; z-index:200;">
<%else%>
<div id="akttotal" style="position:absolute; left:300; top:120; visibility:visible; z-index:200;">
<%end if%>
<table cellspacing="1" cellpadding="0" border="0" width="261">
<tr>
	<td colspan="2"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	<td align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px;  padding-right : 4px;"><font size="1">timer</td>
	<td><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	<td align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px;  padding-right : 4px;"><font size="1">omsætning</td>
</tr>
<tr>
	<td width="115" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px; padding-right : 1px;"><font size="1">uge: <b><%=weeknumbers(w-1)&" - "& weeknumbers(0)%></b></td>
	<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	<td width="59" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px;  padding-right : 4px;">
	<b><%
	if intZtimerTotalMonth <> "0" then
	Response.write intZtimerTotalMonth
	intZtimerTotalMonth = 0
	end if  
	%></b></td>
	<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	<td align="right" width="85" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px;  padding-right : 2px;"><b><%=ccur(totOmsTotalM)%></b></td>
	</tr>
<tr>
	<td height=20 colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
</table>
<%totOmsTotalM = 0%>

<table cellspacing="0" cellpadding="0" border="0" width="261">	
<tr>
	<td colspan="2" style="!border: 1 px; background-color: LightBlue; border-color: silver; border-style: solid; padding-left : 4px;">Aktiviteter:</td>
	<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	<td align="right" style="!border: 1 px; background-color: LightBlue; border-color: silver; border-style: solid; padding-left : 4px;"><font size=1>timer&nbsp;</td>
	<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	<td align="right" style="!border: 1 px; background-color: LightBlue; border-color: silver; border-style: solid; padding-left : 4px;"><font size=1>omsætning&nbsp;</td></tr>
	
<%
For a = 0 to a - 1
	%>
	
	<tr><td height=2 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
			<tr><td colspan="2" width="115" style="!border: 1 px; background-color: #ffffff; border-color: LightBlue; border-style: solid; padding-left : 4px;"><font size="1"><%=jobnr(a) &"&nbsp;" & jobnavn(a) &"</font><b><br>"& left(aktiviteter(a), 16) %></b></td>
			<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
			<td align="right" valign="bottom" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px; padding-right : 1px;"><%
			
			strSQL = "SELECT timer AS Ztimer, tdato, TAktivitetId, TimePris, Tfaktim FROM timer WHERE "& selmedarbKri &" TAktivitetId = "& aktivitetID(a) &""
				oRec2.open strSQL, oConn, 3
				while not oRec2.EOF
					if datepart("ww", oRec2("tdato")) => weeknumbers(w-1) AND datepart("ww", oRec2("tdato")) <= weeknumbers(0) then 
						intZtimer = intZtimer + csng(oRec2("Ztimer"))
						if oRec2("Tfaktim") = "1" then
						omsAkt = omsAkt + (csng(oRec2("Ztimer")) * ccur(oRec2("TimePris")))
						else
						omsAkt = 0
						end if
					end if
				oRec2.movenext
				wend
				oRec2.close
			Response.write intZtimer
			intZtimer = 0
			%></td>
			<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		<td height=20 valign="bottom" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px;  padding-right : 1px;"><%=ccur(omsAkt)%></td>
	</tr>
	<%omsAkt = 0 
next%>
	<tr><td colspan="6" height=20><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
	<tr>
		<td colspan="2" style="!border: 1 px; background-color: #d2691e; border-color: silver; border-style: solid; padding-left : 4px;">Medarbejdere:</td>
		<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		<td align="right" style="!border: 1 px; background-color: #d2691e; border-color: silver; border-style: solid; padding-left : 4px;"><font size=1>timer&nbsp;</td>
		<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		<td align="right" style="!border: 1 px; background-color: #d2691e; border-color: silver; border-style: solid; padding-left : 4px;"><font size=1>omsætning&nbsp;</td></tr>
	</tr>
	<%
	'** Medarbejder total **
	For t = 0 to t - 1
	%>
	
	<tr><td height=2 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
			<tr>
			<td colspan=2 style="!border: 1 px; background-color: #FFFFFF; border-color: #d2691e; border-style: solid; padding-left : 4px;"><%=medarbejder(t) %></td>
			<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
			<td align=right style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px; padding-right : 4px;">
			<%
			strSQL = "SELECT timer AS Ztimer, tdato, TAktivitetId, TimePris, Tfaktim FROM timer WHERE Tmnr = "& medarbID(t) &""
				oRec2.open strSQL, oConn, 3
				while not oRec2.EOF
					for a = 0 to a - 1
					if datepart("ww", oRec2("tdato")) => weeknumbers(w-1) AND datepart("ww", oRec2("tdato")) <= weeknumbers(0) AND oRec2("TAktivitetId") = aktivitetID(a) then 
						intZtimer = intZtimer + csng(oRec2("Ztimer"))
						if oRec2("Tfaktim") = "1" then
						omsAktMedT = omsAktMedT + (csng(oRec2("Ztimer")) * ccur(oRec2("TimePris")))
						else
						omsAktMedT = omsAktMedT
						end if
					end if
					next
				oRec2.movenext
				wend
				oRec2.close
			Response.write intZtimer
			intZtimer = 0
			%></td>
			<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		<td valign="bottom" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px;  padding-right : 1px;"><%=ccur(omsAktMedT)%></td>
	</tr>
	<%omsAktMedT = 0  
	next
	%>
	</table><br>
	<br>
	
<table cellspacing="0" cellpadding="0" border="0" width="261">	
<tr><td colspan="2" height=20 style="padding-left : 4px;">Aktiviteter fordelt pr. medarbejder:</td>
<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		<td valign="bottom" align="right"><font size=1>timer&nbsp;&nbsp;&nbsp;</td>
		<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		<td valign="bottom" align="right"><font size=1>omsætning&nbsp;</td></tr>
</tr>
<%
For t = 0 to t - 1
if t > 0 then
varHeight = 20
else 
varHeight = 2
end if%>
<tr>
	<td height=<%=varHeight%> colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
<tr>
	<td colspan=6 style="!border: 1 px; background-color: #FFFFFF; border-color: #d2691e; border-style: solid; padding-left : 4px;"><b><%=medarbejder(t) %></b></td>
</tr>
<%

For a = 0 to a - 1

	%>
	<tr>
		<td height=2 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
	<tr>
			<td colspan=2 width="125" style="!border: 1 px; background-color: #FFFFFF; border-color: LightBlue; border-style: solid; padding-left : 4px;"><font size="1"><%=jobnr(a) &"&nbsp;" & jobnavn(a) &"</font><b><br>" & left(aktiviteter(a), 16) %></b></td>
			<td width=1><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
			<td width="50" align=right valign="bottom" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px; padding-right : 1px;"><%
			
			strSQL = "SELECT timer AS Ztimer, tdato, TAktivitetId, TimePris, Tfaktim FROM timer WHERE TAktivitetId = "& aktivitetID(a) &" AND Tmnr = "& medarbID(t) &""
				oRec2.open strSQL, oConn, 3
				while not oRec2.EOF
					if datepart("ww", oRec2("tdato")) => weeknumbers(w-1) AND datepart("ww", oRec2("tdato")) <= weeknumbers(0) then 
						intZtimer = intZtimer + csng(oRec2("Ztimer"))
						if oRec2("Tfaktim") = "1" then
						omsAktMed = omsAktMed + (csng(oRec2("Ztimer")) * ccur(oRec2("TimePris")))
						else
						omsAktMed = 0
						end if
					end if
				oRec2.movenext
				wend
				oRec2.close
			Response.write intZtimer
			intZtimer = 0
			%></td>
			<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
			<td width="60" valign="bottom" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px;  padding-right : 1px;"><%=ccur(omsAktMed)%></td>
			</tr>
<%omsAktMed = 0
	next
next%>
</table><br><br><br>
<!-- slut højreside div -->
</div>
<%
else
%>
<div id="ingenregi" style="position:absolute; left:35; top:140; width:60%; height:600; visibility:visible; z-index:100;">
<table width="480" celpadding="2" cellspacing="1" border="0">
<tr><td height="50" align="center" valign="middle" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px; padding-right : 4px;">Ingen registreringer i den valgte periode.</td></tr></table>
</div>
<%
end if
end if 
%>
<!--#include file="../inc/regular/footer_inc.asp"-->
<script>
document.all["loading"].style.visibility = "hidden";
</script>