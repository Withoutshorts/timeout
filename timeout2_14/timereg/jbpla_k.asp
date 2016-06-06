<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
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
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	
	restimerIalt = 0
	arbstyrkeIalt = 0
	forkaltimerTot = 0
	
	if func = "redalle" then
		
		tildeltimer_timer = request("FM_tildeltimer_timer")
		
		
		dim jobidsThis
		Redim jobidsThis(0)
		lastjobid = 0
		
		jobids = split(request("FM_sel_job"), ",")
		for x = 0 to Ubound(jobids)
				
				Redim preserve jobidsThis(x)
				'Response.write jobids(x)
				jobidsThis(x) = jobids(x)
				semmihvor = instr(jobidsThis(x), ";")
				len_jobids = len(jobidsThis(x))
				jobid = left(jobidsThis(x), semmihvor - 1)
				medid = mid(jobidsThis(x), semmihvor + 1, len_jobids - (semmihvor-1))
				
				
				if lastmedid <> medid then
					lastjobid = 0
					lastmedid = medid
				end if
					
					if lastjobid <> jobid then
					
						datoer = split(request("FM_sel_dage"), ",")
						for y = 0 to Ubound(datoer)
							'Response.write jobid & " -- "& medid &": "
							'Response.write datoer(y) & "<br>"
							
							'call normtimer(tildeltimer_timer, datoer(y), medid, jobid)
							
						next
					
					lastjobid = jobid
					end if
					
				
				
		next 
	end if
	
	if func = "red" then
	
	'Response.write request("FM_dato") & "#<br>"
	dato = request("FM_dato") 'day(request("FM_dato"))&"/"&month(request("FM_dato"))&"/"&year(request("FM_dag"))
	'Response.write request("FM_dato")
	
	tildeltimer_medid = request("FM_tildeltimer_mid")
	tildeltimer_jobid = request("FM_tildeltimer_job")
	tildeltimer_timer = request("FM_tildeltimer_timer")
	
	'call normtimer(tildeltimer_timer, dato, tildeltimer_medid, tildeltimer_jobid)
	
	end if
	
	
	'***** Function Normtimer ****
	'*** Henter normeret timer ved 1/1 dag. ***
	public tildeltimer_timer
	function xnormtimer(tildeltimer_timer, dato, medid, jobid)
	
	sqlDato = year(dato)&"/"&month(dato)&"/"&day(dato)
	
	
		if tildeltimer_timer = 7 then
		
				select case weekday(dato)
				case 1
				normdag = "normtimer_son"
				case 2
				normdag = "normtimer_man"
				case 3
				normdag = "normtimer_tir"
				case 4
				normdag = "normtimer_ons"
				case 5
				normdag = "normtimer_tor"
				case 6
				normdag = "normtimer_fre"
				case 7
				normdag = "normtimer_lor"
				case else
				normdag = "normtimer_man"
				end select
				
				
			strSQL = "SELECT medarbejdertype, "& normdag &" AS timer FROM medarbejdere m"_
			&" LEFT JOIN medarbejdertyper t ON (t.id = m. medarbejdertype)"_
			&" WHERE m.mid = " & medid
			
			'Response.write strSQL & "<hr>"
			'Response.flush
			
			oRec.open strSQL, oConn, 3 
			while not oRec.EOF 
				timerThis = oRec("timer")
			oRec.movenext
			wend
			oRec.close 
			
			
			if len(timerThis) <> 0 then
			timerThis = replace(timerThis, ",",".")
			else
			timerThis = 0
			end if
		
		else
		timerThis = tildeltimer_timer
		end if
		
		sttid = "09:00:00"
		sltid = formatdatetime(dateadd("h", timerThis, dato &" "&sttid ), 3)
		
		'*** Sletter gamle ressource registreringer ****
		strSQL_del = "DELETE FROM ressourcer WHERE dato = '"& sqlDato &"' "_
		&" AND jobid = "& jobid &" AND mid = " & medid
		oConn.execute(strSQL_del)
		
		'**** Oprettter nye ressource registreringer ****
		strSQL = "INSERT INTO ressourcer (dato, mid, jobid, aktid, starttp, sluttp, timer) VALUES "_
		&" ('"& sqlDato &"', "& medid &", "& jobid &", 0, '"& sttid &"', '"& sltid &"', "& timerThis &")"
		'Response.write strSQL & "<br><br>"
		'Response.flush
		oConn.execute(strSQL)
		
	end function
	'***************************************************
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	'*** Hvis plus eller minus knap er brugt ***
	if len(request("selectplus")) <> 0 then
	selectplus = request("selectplus")
	else
	selectplus = 0
	end if
	
	if len(request("selectminus")) <> 0 then
	selectminus = request("selectminus")
	else
	selectminus = 0
	end if
	
	
	if len(request("mdselect")) <> 0 then
		monththis = request("mdselect")
		
		if selectminus <> 0 then
		monththis = monththis - 1
		end if
		
		if selectplus <> 0 then
		monththis = monththis + 1
		end if
		
		
		select case monththis 
		case 0
		monththis = 12
		yearthis = request("aarselect")
		yearthis = yearthis - 1
		case 13
		monththis = 1
		yearthis = request("aarselect")
		yearthis = yearthis + 1
		case else
		yearthis = request("aarselect")
		end select
		
	else
	monththis = month(now)
	yearthis = year(now)
	end if
	
	
			'** Antal dage i md ***
			Select case monththis
			case "4", "6", "9", "11"
				mthDays = 30
			case 2
					select case yearthis
					case 2004, 2008, 2012, 2016, 2020, 2024, 2028
					mthDays = 29
					case else 
					mthDays = 28
					end select
			case else
				mthDays = 31
			end select
			
	
	mthDaysUse = mthDays
	monthUse = monththis
	yearUse = yearthis
	
	jobStartKri = yearUse&"/"&monthUse&"/"&"1"
	jobSlutKri = yearUse&"/"&monthUse&"/"&mthDaysUse
	
	daynow = day(now)&"/"&month(now)&"/"&year(now)
	
	'** vis alle uanset datoKri
	if visalle = 1 then
	datoKri = ""
	else
	datoKri = "AND jobstartdato <= '"&jobSlutKri&"' AND jobslutdato >= '"&jobStartKri&"'"
	end if
	
	
	'*** Jobnavn kriterie ***		
	
		'if len(request("mdselect")) <> 0 then ' AND request("jobselect") <> "Jobnavn.. (Alle)" then
			if trim(request("jobselect")) <> "(Alle)" then
			jobnavn = request("jobselect")
			else
			jobnavn = ""
			end if
			jobNavnKri = " AND jobnavn LIKE '"& jobnavn &"%' "
			
			if len(request("jobselect")) <> 0 then
			jobnavn = request("jobselect")
			else
			jobnavn = "(Alle)"
			end if
		'else
		
		'***** Henter 1 jobnavn ****
		'strSQL = "SELECT jobnavn FROM job WHERE fakturerbart = 1 AND jobstatus = 1 "& datoKri &" ORDER BY jobnavn"
		''Response.write strSQL
		'oRec.open strSQL, oConn, 3 
		'if not oRec.EOF then
		'	jobnavn = left(oRec("jobnavn"), 1)
		'end if
		'oRec.close 
		
		'jobNavnKri = " AND jobnavn LIKE '"& jobnavn &"%' "
		'end if
	
	
	
	if len(request("visalle")) <> 0 then
	visalle = 1
	else
	visalle = 0
	end if
	
	%>
	<script language=javascript>
	function clearJnavn() {
	document.getElementById("jobselect").value = ""
	}
	
	function seladd(){
	document.getElementById("selectplus").value = 1 
	}
	
	function selminus(){
	document.getElementById("selectminus").value = 1 
	}
	
	function show(jobid) {
	lastdivval = document.getElementById("FM_lastdiv").value 
	document.getElementById("d_"+lastdivval+"").style.visibility = "hidden"
	document.getElementById("d_"+lastdivval+"").style.display = "none"
	
	document.getElementById("d_"+jobid+"").style.visibility = "visible"
	document.getElementById("d_"+jobid+"").style.display = ""
	
	document.getElementById("FM_lastdiv").value = jobid 
	}
	
	function hide(jobid) {
	document.getElementById("d_"+jobid+"").style.visibility = "hidden"
	document.getElementById("d_"+jobid+"").style.display = "none"
	}
	
	//function opretny(){
	//document.getElementById("opretny").style.visibility = "visible"
	//document.getElementById("opretny").style.display = ""
	//}
	
	</script>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	<%call tsamainmenu(4)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	if showonejob <> 1 then
		call resstopmenu()
	end if
	%>
	</div>
	
		
	
	<%leftPos = 20%>
	<%topPos = 132%>
	<!-------------------------------Sideindhold------------------------------------->
	
	<!-- Oversigts kalender -->
	<div id="overblik" style="position:absolute; left:<%=leftPos%>; top:<%=topPos%>; visibility:visible;">
	<h3>Job Belastning</h3>
	
	
	<table cellspacing="0" cellpadding="0" border="0">
	<form action="jbpla_k.asp" method="post" name="mdselect" id="mdselect">
	<tr>
	<td colspan=3 width=350>
	<%
	if visalle = 1 then
	chkvalle = "CHECKED"
	else
	chkvalle = ""
	end if 
	%>
	<input type="checkbox" name="visalle" id="visalle" value="1" <%=chkvalle%>>
	<font class="megetlillesort">Vis alle job, uanset uanset om jobperioden ligger uden for det valgte interval.
	</td>
	</tr>
	<tr>
	<td><img src="../ill/icon_jobmapper_lysbla_24.gif" width="24" height="24" alt="Vælg Job" border="0"></td>
	<td>
	<input type="text" name="jobselect" id="jobselect" size="15" value="<%=jobnavn%>" onFocus="clearJnavn()">
	&nbsp;<select name="mdselect">
	<option value="<%=monththis%>" SELECTED><%=left(monthname(monththis),3)%></option>
	<option value="1">Jan</option>
	<option value="2">Feb</option>
	<option value="3">Mar</option>
	<option value="4">Apr</option>
	<option value="5">Maj</option>
	<option value="6">Jun</option>
	<option value="7">Jul</option>
	<option value="8">Aug</option>
	<option value="9">Sep</option>
	<option value="10">Okt</option>
	<option value="11">Nov</option>
	<option value="12">Dec</option>
	</select><select name="aarselect">
	<option value="<%=yearthis%>" SELECTED><%=yearthis%></option>
	<option value="2002">2002</option>
	<option value="2003">2003</option>
	<option value="2004">2004</option>
	<option value="2005">2005</option>
	<option value="2005">2006</option>
	<option value="2005">2007</option>
	<option value="2005">2008</option>
	</select></td><td valign="bottom"><input type="image" src="../ill/pilstorxp.gif"></td></tr>
	
	<input type="hidden" name="selectminus" id="selectminus" value="0">
	<input type="hidden" name="selectplus" id="selectplus" value="0">
	
	<div style="position:absolute; left:380; top:<%=topdiv+105%>;">
	<input type="submit" name="submitplus" id="submitplus" value="<- <%=monthname(month(dateadd("m", -1, "1/"&monththis&"/"&yearthis)))%>" style="font-size:9px; border:1px #003399 solid; background-color:#ffffff; width:75px;" onClick="selminus()">
	</div>
	<div style="position:absolute; left:470; top:<%=topdiv+105%>;">
	<input type="submit" name="submitplus" id="submitplus" value="<%=monthname(month(dateadd("m", +1, "1/"&monththis&"/"&yearthis)))%> ->" style="font-size:9px; border:1px #003399 solid; background-color:#ffffff; width:75px;" onClick="seladd()">
	</div>
	</form>
	</table>
	</div>
	
	
	<%
	toppxThisDiv = 235
	leftpxThisDiv = 690
	%>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv-90%>px; background-color:#ffffe1; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv-73%>px; padding:2px;">&nbsp;0 - 1 time tildelt</div>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv+10%>px; background-color:#cccccc; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv+30%>px; padding:2px;">&nbsp;1 - 4 tildelt</div>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv+100%>px; background-color:PeachPuff; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv+120%>px; padding:2px;">&nbsp;4 - 8 tildelt.</div>		
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv+185%>px; background-color:red; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv+205%>px; padding:2px;">&nbsp; > 8 timer tildelt.</div>		
	
	
	<div id="job" style="position:absolute; left:<%=leftPos%>; top:<%=topPos+135%>; visibility:visible;">
	
	<%
	'*************************************************************************************
	'Joboversigten
	'Viser alle job det matcher filter i den valgte måned.
	'*************************************************************************************
	
	%>
	
	<table cellspacing="0" cellpadding="0" border="0">
	<tr bgcolor="#d6dff5"><td bgcolor="#5582d2" colspan=2 align=right valign="top" style="padding-right:2; padding-left:2; border-left:1px #ffffff solid; border-bottom:1px #ffffff solid; border-top:1px #ffffff solid;" class=alt><b><%=monthname(monthUse)%>&nbsp;<%=yearUse%></b></td>
		<td>
				<table cellspacing="0" cellpadding="0" border="0">
				<tr>
							<%for d = 0 to mthDaysUse - 1
							getnyDato = dateadd("d", d, jobStartKri)
							showday = d + 1
							
							if cdate(daynow) = cdate(getnyDato) then
							bgcolordatoer = "#FFCC00"
							else
							
								select case weekday(getnyDato)
								case 1,7
								bgcolordatoer = "#cccccc"
								case else
								bgcolordatoer = "#8caae6"
								end select
							
							end if
							
							%>
							<td bgcolor="<%=bgcolordatoer%>" valign="top" align=center width=20 style="border-left:1px #ffffff solid;"><font class="megetlillesort"><%=showday%></td>
							<%next%>
				<td height=15 width=50 align=right style="padding-left:5px; border-bottom:1px #ffffff solid; border-top:1px #ffffff solid; border-left:1px #ffffff solid" bgcolor="#8caae6"><font class=megetlillehvid>Forkalk.</td>	
				<td height=15 width=50 align=right style="padding-left:5px; border-bottom:1px #ffffff solid; border-top:1px #ffffff solid;" bgcolor="#8caae6"><font class=megetlillehvid>Real.</td>
				<td height=15 width=50 align=right style="padding-left:5px; border-bottom:1px #ffffff solid; border-top:1px #ffffff solid" bgcolor="#8caae6"><font class=megetlillehvid>Ressource</td>	
				<td height=15 width=50 align=right style="padding-left:5px; border-bottom:1px #ffffff solid; border-top:1px #ffffff solid; border-right:1px #ffffff solid" bgcolor="#8caae6"><font class=megetlillehvid>Balance</td>	
				
				</tr>
				</table>
			
		</td>
		</tr>
	
	<%
	x = 0
	dim arrjobid, arrjobnavn, arrjobnr, arrjstdato, arrslutdato, forkaltimer, thisJobResTimerTot
	Redim arrjobid(x)
	Redim arrjobnavn(x)
	Redim arrjobnr(x)
	Redim arrjstdato(x)
	Redim arrslutdato(x)
	Redim forkaltimer(x)
	Redim thisJobResTimerTot(x)
	
		
	strSQL = "SELECT job.id AS jobid, jobnavn, jobnr, jobstartdato, jobslutdato, "_
	&" jobstatus, budgettimer, ikkebudgettimer, k.kkundenavn, k.kkundenr FROM job "_
	&" LEFT JOIN kunder k ON (kid = jobknr)"_
	&" WHERE fakturerbart = 1 AND jobstatus = 1 "& datoKri &" "& jobNavnKri &" ORDER BY jobslutdato, jobnavn"
	'Response.write strSQL
	'Response.flush
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	Redim preserve arrjobid(x)
	Redim preserve arrjobnavn(x)
	Redim preserve arrjobnr(x)
	Redim preserve arrjstdato(x)
	Redim preserve arrslutdato(x)
	Redim preserve forkaltimer(x)
	Redim preserve thisJobResTimerTot(x)
	
	
	arrjobid(x) = oRec("jobid")
	arrjobnavn(x) = oRec("jobnavn")
	arrjobnr(x) = oRec("jobnr")
	arrjstdato(x) = oRec("jobstartdato")
	arrslutdato(x) = oRec("jobslutdato")
	forkaltimer(x) = oRec("budgettimer") + oRec("ikkebudgettimer")
		%>
		<tr>
			
			<td width="200" height=30 bgcolor="#ffffff" style="padding-right:2; padding-left:8; border-bottom:1px #cccccc solid;">
			<a href="ressource_belaeg_jbpla.asp?id=<%=oRec("jobid")%>&showonejob=1" target="_blank"><%=left(oRec("jobnavn"), 20)%> (<%=oRec("jobnr")%>)</a><br>
			<%=oRec("kkundenavn")%> (<%=oRec("kkundenr")%>)
			&nbsp;</td>
			<td bgcolor="#8caae6" style="border-bottom:1px #ffffff solid; padding:5px;"><a href="jobplanner.asp?menu=job&id=<%=oRec("jobid")%>" target="_blank"><img src="../ill/popupcal_small.gif" width="11" height="10" alt="Rediger jobperiode og aktiviteter" border="0"></a></td>
	
		<td height=8 valign="top">
				<table cellspacing="0" cellpadding="0" border="0">
				<tr>
						<%for d = 0 to mthDaysUse - 1
							
							getnyDato = dateadd("d", d, jobStartKri)
							ressourcetimerDato = 0
							
							if cdate(oRec("jobstartdato")) <= cdate(getnyDato) AND cdate(oRec("jobslutdato")) >= cdate(getnyDato) then
							imgthis = "g-prik.gif"
									
									sqlday = d + 1
									sqlresdato = yearthis&"/"&monththis&"/"&sqlday 
									
									strSQL3 = "SELECT sum(timer) AS restimer FROM ressourcer WHERE ressourcer.jobid = "& oRec("jobid") &" AND dato = '"&sqlresdato&"'"
									'Response.write strSQL3
									oRec3.open strSQL3, oConn, 3 
									if not oRec3.EOF then
									ressourcetimerDato = oRec3("restimer")
									end if
									oRec3.close 
									
							else
							imgthis = "blank.gif"
							end if
							
							
							if ressourcetimerDato <> 0 then
							ressourcetimerDato = ressourcetimerDato
							else
							ressourcetimerDato = ""
							end if
							
							select case weekday(getnyDato)
							case 1,7
							bgcolordatoer = "#eff3ff"
							case else
							bgcolordatoer = "#ffffff"
							end select
							
							%>	
							<td valign="top" height=30 align="center" width=20 bgcolor="<%=bgcolordatoer%>" style="border-left:1px #d6dff5 solid; border-top:1px #cccccc solid; padding-top:4px;">
							<img src="../ill/<%=imgthis%>" width="6" height="6" alt="" border="0">
							
							<%
							if len(ressourcetimerDato) <> 0 then
							
							if ressourcetimerDato = 0 then
							bgtimerCol = "#ffffff"
							end if
							
							if ressourcetimerDato > 0 AND ressourcetimerDato <= 1 then
							bgtimerCol = "#ffffe1"
							end if
							
							if ressourcetimerDato > 1 AND ressourcetimerDato <= 4 then
							bgtimerCol = "#cccccc"
							end if
							
							if ressourcetimerDato > 4 AND ressourcetimerDato <= 8 then
							bgtimerCol = "PeachPuff"
							end if
							
							if ressourcetimerDato > 8 then
							bgtimerCol = "red"
							end if
							
							else
							
								bgtimerCol = "#ffffff"
						
							end if
							%>
							
							<div style="position:relative; background-color:<%=bgtimerCol%>;">
							<font class=megetlillesort><%=ressourcetimerDato%>
							</div>
							</td>
							<%
							if len(ressourcetimerDato) <> 0 then
							addTimer = ressourcetimerDato
							else
							addTimer = 0
							end if
							
							thisJobResTimerTot(x) = thisJobResTimerTot(x) + addTimer 
						
						next%>
						
					<td align=right width=50 style="padding-right:5px; border-left:1px #ffffff solid; border-bottom:1px #ffffff solid" bgcolor="#8caae6">
					<font class=lillehvid><%=forkaltimer(x)%>
					</td>
					<td align=right width=50 style="padding-right:5px; border-bottom:1px #ffffff solid" bgcolor="#8caae6">
					<font class=lillehvid><%=thisJobResTimerTot(x)%>
					</td>
					<td align=right width=50 style="padding-right:5px; border-bottom:1px #ffffff solid" bgcolor="#8caae6">
					<font class=lillehvid><%=thisJobResTimerTot(x)%>
					</td>	
					<td align=right width=50 style="padding-right:5px; border-right:1px #ffffff solid; border-bottom:1px #ffffff solid" bgcolor="#8caae6">
					<font class=lillehvid><b><%=forkaltimer(x) - thisJobResTimerTot(x)%></b>
					</td>	
						
				</tr>
				</table>
		</td></tr>
		<%
		forkaltimerTot = forkaltimerTot + forkaltimer(x)
		x = x + 1
		Response.flush
	oRec.movenext
	wend
	oRec.close
	%>
	</tr>	
	</table>
	
	<br><br>
	
	
	
			<table cellspacing="0" cellpadding="0" border="0">
				<tr>
					<td bgcolor="#5582d2" width=220 rowspan=3 align=right valign="top" style="padding-right:2; padding-left:2;" class=alt><b>Dage<br>Ressourtimer tildelt ialt<br>Saml. Normeret arb.styrke</b></td>
				
							<%
							Dim totRestimerDato
							Redim totRestimerDato(d)
							
							for d = 0 to mthDaysUse - 1
							'totRestimerDato = 0
							
							
							getnyDato = dateadd("d", d, jobStartKri)
							showday = d + 1
							
							if cdate(daynow) = cdate(getnyDato) then
							bgcolordatoer = "#FFCC00"
							else
							
								select case weekday(getnyDato)
								case 1,7
								bgcolordatoer = "#cccccc"
								case else
								bgcolordatoer = "#8caae6"
								end select
							
							end if
							
							%>
							<td bgcolor="<%=bgcolordatoer%>" valign="top" align=center width=20 style="border-left:1px #ffffff solid; padding:1px;">
							<b><%=showday%></b><br>
							</td>
							<%next%>
							
							</tr><tr>
							<%
							for d = 0 to mthDaysUse - 1
							
							sqlday = d + 1
							sqlresdato = yearthis&"/"&monththis&"/"&sqlday 
							
							%>
							<td bgcolor="#ffffff" valign="top" align=center width=20 style="border-left:1px #ffffff solid; padding:1px;">
							<%
							strSQL3 = "SELECT sum(timer) AS restimer FROM ressourcer WHERE dato = '"&sqlresdato&"'"
							'Response.write strSQL3
							oRec3.open strSQL3, oConn, 3 
							if not oRec3.EOF then
							 Redim preserve totRestimerDato(d)
							 totRestimerDato(d) = oRec3("restimer")
							end if
							oRec3.close 
							
							if len(totRestimerDato(d)) <> 0 then
							showResTimer = totRestimerDato(d)
							else
							Redim preserve totRestimerDato(d)
							totRestimerDato(d) = 0
							showResTimer = 0
							end if
							%>
							<font class="megetlillesort"><%=showResTimer%></td>
							<%
							restimerIalt = restimerIalt + totRestimerDato(d)
						next%>
				
				
				<%
						'** Finder samlet antal normerede timer ***
						for d = 1 to 7
							
							select case d
							case 1
							normdag = "normtimer_son"
							case 2
							normdag = "normtimer_man"
							case 3
							normdag = "normtimer_tir"
							case 4
							normdag = "normtimer_ons"
							case 5
							normdag = "normtimer_tor"
							case 6
							normdag = "normtimer_fre"
							case 7
							normdag = "normtimer_lor"
							case else
							normdag = "normtimer_man"
							end select
							
							
							strSQL3 = "SELECT mid, medarbejdertype, "& normdag &" AS timer FROM medarbejdere m"_
							&" LEFT JOIN medarbejdertyper t ON (t.id = m. medarbejdertype)"_
							&" WHERE m.mansat = '1' OR m.mansat = 'yes' GROUP BY mid" 
							
							'Response.write strSQL3
							'Response.flush
							
							
							oRec3.open strSQL3, oConn, 3 
							while not oRec3.EOF 
								
								select case d
								case 1
								normdagSon = normdagSon + oRec3("timer")
								case 2
								normdagMan = normdagMan + oRec3("timer")
								case 3
								normdagTir = normdagTir + oRec3("timer")
								case 4
								normdagOns = normdagOns + oRec3("timer")
								case 5
								normdagTor = normdagTor + oRec3("timer")
								case 6
								normdagFre = normdagFre + oRec3("timer")
								case 7
								normdagLor = normdagLor + oRec3("timer")
								end select
								
								
								
							oRec3.movenext
							wend
							oRec3.close 
							%>
							
							<%next%>
				
				
				</tr>
						
						<tr>
							<%for d = 0 to mthDaysUse - 1
							
							getnyDato = dateadd("d", d, jobStartKri)
							showday = d + 1
							
							select case weekday(getnyDato)
							case 1
							normdagShow = normdagSon
							case 2
							normdagShow = normdagMan
							case 3
							normdagShow = normdagTir
							case 4
							normdagShow = normdagOns
							case 5
							normdagShow = normdagTor
							case 6
							normdagShow = normdagFre
							case 7
							normdagShow = normdagLor
							end select
							
							if normdagShow > totRestimerDato(d) then
							bgcolordatoer = "DarkSeaGreen"
							else
							bgcolordatoer = "Red"
							end if
							
							%>
							<td bgcolor="<%=bgcolordatoer%>" valign="top" align=center style="border-left:1px #ffffff solid; padding:1px;">
							
							<font class="megetlillesort"><%=normdagShow%></td>
							<%
							arbstyrkeIalt = arbstyrkeIalt + normdagShow
							next%>
				</tr>
				</table><br>
				<div style="position:relative; top:20px; left:0px; background-color:DarkSeaGreen; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
				<div style="position:relative; top:0px; left:20px; padding:2px;">&nbsp;Ressourcetimer til rådighed</div>
				<div style="position:relative; top:-17px; left:180px; background-color:red; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
				<div style="position:relative; top:-37px; left:200px; padding:2px;">&nbsp;Ressourcetimer opbrugt</div>
	
	<br>
	<br><br>&nbsp;	
	</div>
	
	<div style="position:absolute; top:130px; left:600px; padding:10px; background-color:#ffffff; border:2px silver solid;">
	<table cellspacing=1 cellpadding=2 border=0 bgcolor="#5582d2">
	<tr>
		<td bgcolor="#ffffff" align=right valign=bottom width=60 class=lille>Forkalkuleret<br> på job</td>
		<td bgcolor="#ffffff" align=right valign=bottom width=60 class=lille>Ressourcetimer<br>tildelt</td>
		<td bgcolor="#ffffff" align=right valign=bottom width=60 class=lille>Normeret arbejdstyrke</td>
		<td bgcolor="#ffffff" align=right valign=bottom width=60 class=lille>Balance<br> Forkalk./Res.</td>
		<td bgcolor="#ffffff" align=right valign=bottom width=80 class=lille>Balance<br> Norm. arb.styrke / Res. timer tildelt</td>
	</tr>
	<tr>
		<td bgcolor="#ffffff" align=right><b><%=formatnumber(forkaltimerTot, 2)%></b></td>
		<td bgcolor="#ffffff" align=right><b><%=formatnumber(restimerIalt, 2)%></b></td>
		<td bgcolor="#ffffff" align=right><b><%=formatnumber(arbstyrkeIalt, 2)%></b></td>
		<td bgcolor="#ffffff" align=right>
		<%
		balfkres = (forkaltimerTot - (restimerIalt))
		%>
		<b><%=formatnumber(balfkres, 2)%></b>
		</td>
		<td bgcolor="#ffffff" align=right>
		<%
		balresnorm = (arbstyrkeIalt - (restimerIalt))
		%>
		<b><%=formatnumber(balresnorm, 2)%></b>
		</td>
	</tr>
	</table>
	</div>
	
	
	
	<!--- restimer -->
	
	
	

<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
