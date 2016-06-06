<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
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
	case "-"
	
	case else
	%>
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
	
	<%
	'*** Sætter lokal dato/kr format. *****
	Session.LCID = 1030
	
	leftPos = 20
	topPos = 102
	

	if len(request("FM_md")) <> 0 then
	strMd = request("FM_md")
	strAar = request("FM_aar")
	intMid = request("FM_medarb")
	intJid = request("FM_job")
	else
	intMid = 0
	intJid = 0
	strMd = month(date)
	strAar = year(date)
	end if
	
	select case strMd
	case 1,3,5,7,8,10,12
	usedag = 31
	case 2
	usedag = 28
	case else
	usedag = 30
	end select
	
	startdato = strAar&"/"&strMd&"/1"
	slutdato = strAar&"/"&strMd&"/"&usedag
	
	if len(request("FM_periode")) <> 0 then
	usePer = request("FM_periode")
	else
	usePer = 0
	end if
	
	if cint(usePer) = 1 then
	perKri = " AND ressourcer.dato BETWEEN '"& startdato &"' AND '"& slutdato &"'"
	perKriReal = " AND timer.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"'"
	else
	perKri = ""
	perKriReal = ""
	end if 
	
	
	if cint(intJid) = 0 then
	restimJidKri = " job.id <> "& intJid &""
	realTimJid = " tjobnr <> "
	else
	restimJidKri = " job.id = "& intJid &" "
	realTimJid = " tjobnr = "
	end if
	%>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=leftPos%>px; top:<%=topPos%>px; visibility:visible;">
	<h3>Ressource belægning</h3>
	
	
	
	<%
	toppxThisDiv = 75
	leftpxThisDiv = 620
	%>
	<div style="position:absolute; top:<%=toppxThisDiv+25%>px; left:<%=leftpxThisDiv-90%>px; background-color:#ffffe1; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
	<div style="position:absolute; top:<%=toppxThisDiv+25%>px; left:<%=leftpxThisDiv-75%>px; padding:2px;">&nbsp; 0 - 2 timer tildelt</div>
	<div style="position:absolute; top:<%=toppxThisDiv+25%>px; left:<%=leftpxThisDiv+30%>px; background-color:#cccccc; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
	<div style="position:absolute; top:<%=toppxThisDiv+25%>px; left:<%=leftpxThisDiv+50%>px; padding:2px;">&nbsp;2 - 7 tildelt</div>
	<div style="position:absolute; top:<%=toppxThisDiv+25%>px; left:<%=leftpxThisDiv+125%>px; background-color:PeachPuff; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
	<div style="position:absolute; top:<%=toppxThisDiv+25%>px; left:<%=leftpxThisDiv+145%>px; padding:2px;">&nbsp;> 7 timer tildelt.</div>		

	
	<table border=0 cellspacing=0 cellpadding=0>
	<form action="ressource_belaeg.asp?menu=res" method="post" name="periode" id="periode">
	<tr>
	<td>
	<select name="FM_medarb" id="FM_medarb">
	<!--<option value="0">Alle</option>	-->
	<%
	strSQL = "SELECT mnavn, mid FROM medarbejdere ORDER BY mnavn"
	oRec.open strSQL, oConn, 3 
	x = 0
	while not oRec.EOF 
	if len(request("FM_medarb")) = 0 AND x = 0 then
		selthis = "SELECTED"
		intMid = session("mid") 'oRec("mid")
	else	
		if cint(intMid) = oRec("mid") then
		selthis = "SELECTED"
		else
		selthis = ""
		end if
	end if
	%>
	<option value="<%=oRec("mid")%>" <%=selthis%>><%=oRec("mnavn")%></option>	
	<%
	x = x + 1
	oRec.movenext
	wend
	oRec.close 
	
	%></select>
	</td>
	<td>&nbsp;&nbsp;
	<select name="FM_job" id="FM_job" style="width:200px;">
	<option value="0">Alle</option>	
	<%
	strSQL = "SELECT jobnavn, id FROM job WHERE jobstatus = 1 AND fakturerbart = 1 ORDER BY jobnavn"
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
	if cint(intJid) = oRec("id") then
	selthis = "SELECTED"
	else
	selthis = ""
	end if
	%>
	<option value="<%=oRec("id")%>" <%=selthis%>><%=oRec("jobnavn")%></option>	
	<%
	oRec.movenext
	wend
	oRec.close 
	
	%></select>
	</td>
	<td>&nbsp;
	<input type="hidden" name="FM_periode" id="FM_periode" value="1">
	</td>
	<td>&nbsp;&nbsp;
	<select name="FM_md" id="FM_md">
	<%
	x = 1
	for x = 1 to 12
	select case x
	case 1
	strMDname = "Januar"
	case 2
	strMDname = "Februar"
	case 3
	strMDname = "Marts"
	case 4
	strMDname = "April"
	case 5
	strMDname = "Maj"
	case 6
	strMDname = "Juni"
	case 7
	strMDname = "Juli"
	case 8
	strMDname = "August"
	case 9
	strMDname = "September"
	case 10
	strMDname = "Oktober"
	case 11
	strMDname = "November"
	case 12
	strMDname = "December"
	end select
	
	if cint(strMd) = x then
	selthis = "SELECTED"
	showstrMDname = strMDname
	else
	selthis = ""
	end if
	%>
	<option value="<%=x%>" <%=selthis%>><%=strMDname%></option>
	<%next%>
	</select></td><td>
	&nbsp;&nbsp;
	<select name="FM_aar" id="FM_aar">
	<%
	x = 3
	for x = 3 to 7
	select case x
	case 3
	strAarnameid = 2003
	case 4
	strAarnameid = 2004
	case 5
	strAarnameid = 2005
	case 6
	strAarnameid = 2006
	case 7
	strAarnameid = 2007
	end select
	
	if cint(strAar) = strAarnameid then
	selthis = "SELECTED"
	else
	selthis = ""
	end if
	%>
	<option value="<%=strAarnameid%>" <%=selthis%>><%=strAarnameid%></option>
	<%next%>
	</select></td><td style="padding-top:1;">
	&nbsp;&nbsp;<input type="image" src="../ill/pilstorxp.gif">
	</td></tr>
	</form></table>
	<br>
	Klik på det valgte job for at tildele ressource timer.<br>
	Klik på en medarbejder på at se dennes plan over tildelte timer.
	<br><br>
	
	
	
	<%
	'*** MedarbejderKriterie ***
	'if cint(intMid) = 0 then
	'restimMidKri = " AND medarbejdere.mid <> "& intMid &""
	'realtimerMidKri = " AND tmnr <> "& intMid &""
	'resMedarbKri = " AND ressourcer.mid = medarbejdere.mid "
	'else
	restimMidKri = " AND medarbejdere.mid = "& intMid &""
	realtimerMidKri = " AND tmnr = "& intMid &""
	resMedarbKri = " AND ressourcer.mid = medarbejdere.mid "
	'end if
	
	
	dim showrestimer
	dim showresdato
	dim useJobnr
	'dim realtimer
	dim xmid
	dim mnavn
	dim normeretuge
	
	dim normTimerSon
	dim normTimerMan
	dim normTimerTir
	dim normTimerOns
	dim normTimerTor
	dim normTimerFre
	dim normTimerLor
	
	Redim showrestimer(0)
	Redim showresdato(0)
	Redim useJobnr(0)
	'Redim realtimer(0)
	Redim xmid(0)
	Redim mnavn(0)
	Redim normeretuge(0)
	
	Redim normTimerSon(0)
	Redim normTimerMan(0)
	Redim normTimerTir(0)
	Redim normTimerOns(0)
	Redim normTimerTor(0)
	Redim normTimerFre(0)
	Redim normTimerLor(0)
	
	'realtimer = 0
	lastweek = 0
	lastmedarbId = 0
	
	strSQL = "SELECT jobnavn, mnavn, job.id, job.jobnr, medarbejdertype, "_
	&" normtimer_son, normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, normtimer_fre, "_
	&" normtimer_lor, medarbejdere.mid, ressourcer.timer AS restimer, ressourcer.dato AS resdato "_
	&" FROM job, medarbejdere LEFT JOIN medarbejdertyper ON (medarbejdertyper.id = medarbejdertype) "_
	&" LEFT JOIN ressourcer ON (ressourcer.jobid = job.id "& resMedarbKri &" "& perKri &") "_
	&" WHERE "& restimJidKri &" "& restimMidKri &" AND fakturerbart = 1 "_
	&" ORDER BY medarbejdere.mid, job.jobnavn"
	oRec.open strSQL, oConn, 3 
	x = 0
	
	while not oRec.EOF 
	Redim preserve showrestimer(x)
	Redim preserve showresdato(x)
	Redim preserve useJobnr(x)
	'Redim preserve realtimer(x)
	Redim preserve xmid(x)
	Redim preserve mnavn(x)
	Redim preserve normeretuge(x)
	Redim preserve normTimerSon(x)
	Redim preserve normTimerMan(x)
	Redim preserve normTimerTir(x)
	Redim preserve normTimerOns(x)
	Redim preserve normTimerTor(x)
	Redim preserve normTimerFre(x)
	Redim preserve normTimerLor(x)
	
	useJobnr(x) = oRec("jobnr")
	xmid(x) = oRec("mid")
	mnavn(x) = oRec("mnavn")
	
	if len(oRec("resdato")) <> 0 then
	showrestimer(x) = oRec("restimer")
	showresdato(x) = oRec("resdato")
	else
	showrestimer(x) = 0
	showresdato(x) = "2001/1/1"
	end if
	
	
	if lastmedarbId <> oRec("mid") then 
		normeretuge(x) = oRec("normtimer_son") + oRec("normtimer_man") + oRec("normtimer_tir") + oRec("normtimer_ons") + oRec("normtimer_tor")+ oRec("normtimer_fre")+ oRec("normtimer_lor")
		
		normTimerSon(x) = oRec("normtimer_son")
		normTimerMan(x) = oRec("normtimer_man")
		normTimerTir(x) = oRec("normtimer_tir")
		normTimerOns(x) = oRec("normtimer_ons")
		normTimerTor(x) = oRec("normtimer_tor")
		normTimerFre(x) = oRec("normtimer_fre")
		normTimerLor(x) = oRec("normtimer_lor")
		
		
		strSQL2 = "SELECT sum(timer) AS realiserettimer FROM timer WHERE "& realTimJid & "'"& useJobnr(x) &"' "& realtimerMidKri &" AND tfaktim <> 5"
		'Response.write strSQL2
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
		totalrealtimer = totalrealtimer + oRec2("realiserettimer")
		'Response.write totalrealtimer
		end if
		oRec2.close 
	end if	
	
	
	
	resTimerTot = resTimerTot + oRec("restimer")
	lastJobID = oRec("id")
	lastmedarbId = oRec("mid")
	lastJobnavn = oRec("jobnavn")
	x = x + 1
	oRec.movenext
	wend
	oRec.close 
	
	if cint(intJid) <> 0 then
	'jbTxt = "<a href='javascript:NewWin_large(""ressourcer.asp?menu=res&id="&lastJobID&""")' class=vmenu>"&lastJobnavn&"</a>" 
	jbTxt = lastJobnavn
	else
	jbTxt = "Alle eksterne job" 
	end if
	
	%>
	
<table border=0 cellspacing=1 cellpadding=0 background-color="#5582d2">
<tr>
	<td colspan=40><%=showstrMDname%>&nbsp;<%=right(strAar, 2)%>&nbsp;&nbsp;<b><%=jbTxt%></b></td>
</tr>
<%
numoffdaysinmonth =  datediff("d", startdato, slutdato)
y = 0 
nday = startdato
weekrealtimer = 0
weekrest = 0		
weekbalance = 0


useTopX = (x - 1)
for x = 0 to x - 1
	
	if lastxmid <> xmid(x) then	 'if y = 0 then
		 if x <> 0 then
	 		Response.write "</tr><tr>"_
			&"<td valign=top colspan=40>&nbsp;</td></tr>"
		 end if
	%>
	<tr>
		<td valign=top width=100><font class=megetlillesort>
		<div style="background-color:#eff3ff; border:1px #5582d2 solid; padding-left:2;"><a href="javascript:NewWin_todo('medarb_restimer.asp?menu=res&func=show&FM_medarb=<%=xmid(x)%>&FM_mednavn=<%=mnavn(x)%>')" class=vmenu><%=left(mnavn(x), 14)%></a></div>
		<div style="background-color:#8caae6; padding-left:2;">Realiseret</div>
		<div style="background-color:#ffffff; padding-left:2;">Tildelt</div>
		<div style="background-color:#ffffff; border:1px #ffffff solid; padding-left:1;">+ / - Balance</div>
		<div style="background-color:#8caae6; padding-left:2;">Norm./Real.</div>
		<div style="background-color:#ffffff; padding-left:2;">Norm./Tildelt</div>
		</td>
	<%end if
	
	
			
		y = 0
		for y = 0 TO numoffdaysinmonth%>
		<%
						
				'dblused = 0
				
				if y = 0 then
				nday = dateadd("d", 0, startdato)
				else
				nday = dateadd("d", 1, nday)
				end if
				
				weekthis = datepart("ww", nday,2,2)
				
				if weekthis <> lastweek AND y <> 0 AND lastxmid <> xmid(x) then
				%>
				<td valign=top width="35" align=center><font class=megetlillesort>
				<div style="background-color:#ffffe1; border:1px #5582d2 solid; padding-left:0;"><%=weekthis-1%></div>
				<div style="background-color:#8caae6; padding-left:0;"><%=formatnumber(weekrealtimer, 2)%></div>
				<div style="background-color:#ffffff; padding-left:0;"><%=formatnumber(weekrest, 2)%></div>
				<%
				weekbalance = (weekrest) - (weekrealtimer) 
				normbal = normTimer - (weekrest)
				normreal = normTimer - (weekrealtimer)
				%>
				<div style="background-color:#ffffff; border:1px #000000 solid; padding-left:0;"><%=formatnumber(weekbalance, 2)%></div>
				<div style="background-color:#8caae6;"><%=formatnumber(normreal, 2)%></div>
				<div style="background-color:#ffffff; border:1px #ffffff solid; padding-left:0;"><%=formatnumber(normbal, 2)%></div>
				<%
				mdrealtotal = mdrealtotal + weekrealtimer
				mdrest = mdrest + weekrest
				mdnormbal = mdnormbal + (normbal)
				mdnormreal = mdnormreal + (normreal)
				weekrealtimer = 0
				weekrest = 0
				weekbalance = 0
				normbal = 0
				normreal = 0
				normTimer = 0%>			
				</td><td valign=top align=center width="2"></td>
				<%end if%>
				
				
				<%
				if lastxmid <> xmid(x) then
				%>
				<td valign=top bgcolor="#ffffff" align=center><font class=megetlillesort>
				<%select case weekday(nday)
				case 1
				dbg = "Gainsboro"
				normTimer = normTimer + normTimerSon(x)
				case 2
				dbg = "#eff3ff"
				normTimer = normTimer + normTimerMan(x)
				case 3
				dbg = "#eff3ff"
				normTimer = normTimer + normTimerTir(x)
				case 4
				dbg = "#eff3ff"
				normTimer = normTimer + normTimerOns(x)
				case 5
				dbg = "#eff3ff"
				normTimer = normTimer + normTimerTor(x)
				case 6
				dbg = "#eff3ff"
				normTimer = normTimer + normTimerFre(x)
				case 7
				dbg = "Gainsboro"
				normTimer = normTimer + normTimerLor(x)
				end select
				
				if cdate(date) = cdate(nday) then
				dbg = "#2c962d"
				end if
				%>
				<div style="background-color:<%=dbg%>; border:1px #5582d2 solid; padding-left:2; width:16;"><%=left(formatdatetime(nday, 2), 2)%></div>
				<%
				'Realiseret
					sql3dateKri = year(nday)&"/"&month(nday)&"/"&day(nday)
					rtimerdato = 0
					
					if cint(intJid) = 0 then
					realTimJid2 = " tjobnr <> '0'"
					else
					realTimJid2 = " tjobnr = '"& useJobnr(x) &"'"
					end if
					
					strSQL3 = "SELECT sum(timer) AS rtimerdato FROM timer WHERE "& realTimJid2 &" AND timer.tdato = '"& sql3dateKri &"' AND tmnr = "& xmid(x) &" AND tfaktim <> 5"
					'Response.write strSQL3
					oRec3.open strSQL3, oConn, 3 
					if not oRec3.EOF then
					rtimerdato = oRec3("rtimerdato")
					end if
					oRec3.close 
				
					dayBalance = rtimerdato
					rtimer0 = rtimerdato
						
						
				
						if rtimer0 <> 0 then
						Response.write "<div style='background-color:#ffffff; padding-left:2;'>" & rtimer0 & "</div>"
						else
						Response.write "<div>&nbsp;</div>"
						end if
						
						if len(rtimer0) <> 0 then
						weekrealtimer = weekrealtimer + rtimerdato
						else
						weekrealtimer = weekrealtimer
						end if
					
						
						timerpavalgtejob = 0	
						for l = 0 to useTopX
						'Tildelte timer
							if cdate(showresdato(l)) = cdate(nday) then
						
								if xmid(x) = xmid(l) then 
								timerpavalgtejob = timerpavalgtejob + showrestimer(l)
								weekrest = weekrest + showrestimer(l)
								'dblused = 1
								end if
							else
								weekrest = weekrest 
							end if
						next
						
						if timerpavalgtejob <> "0" then
						showtimerpavalgtejob = timerpavalgtejob
						else
						showtimerpavalgtejob = "&nbsp;"
						end if
						
						if timerpavalgtejob = 0 then
						bgTimer = "#ffffff"
						end if
						
						if timerpavalgtejob > 0 AND timerpavalgtejob <= 2 then
						bgTimer = "#ffffe1"
						end if
						
						if timerpavalgtejob > 2 AND timerpavalgtejob <= 7 then
						bgTimer = "#cccccc"
						end if
						
						if timerpavalgtejob > 7 then
						bgTimer = "PeachPuff"
						end if
						
						Response.write "<div style='background-color:"& bgTimer &"; padding-left:2;'>"&showtimerpavalgtejob&"</div>"
						
						
							'Balance
							if len(rtimer0) <> 0 then
							dayBalance = (timerpavalgtejob - rtimer0)
							else
							dayBalance = timerpavalgtejob
							end if
							
							timerpavalgtejob = 0		
							
							if len(dayBalance) <> 0 then
							dayBalance = dayBalance
							else
							dayBalance = 0
							end if
							
							'Response.write "<u>"&rtimer0&"</u>"
							if len(rtimer0) = 0 OR rtimer0 = 0 then
							Response.write "<div>&nbsp;</div>"
							end if
							
							
							
							'if dblused = 0 then
							'Response.write "<div>&nbsp;</div>"
							'end if
							
							if dayBalance < 0 then
							%><div style="background-color:#ffffff; border:1px #ff3300 solid;"><%=dayBalance%></div><%
							else
							%><div style="background-color:#ffffff; border:1px #009900 solid;"><%=dayBalance%></div><%
							end if
							
							rtimer0 = 0
							
				%>					
				</td><%
				end if	
				
				lastweek = weekthis	
				next%>
				
				
				<%if lastxmid <> xmid(x) then%>
				<td valign=top width="40" align=center><font class=megetlillesort>
				<div style="background-color:#ffffe1; border:1px #5582d2 solid;"><%=formatnumber(weekthis, 2)%></div>
				<div style="background-color:#8caae6;"><%=formatnumber(weekrealtimer, 2)%></div>
				<div style="background-color:#ffffff;"><%=formatnumber(weekrest, 2)%></div>
				<%
				weekbalance = (weekrest) - (weekrealtimer)
				normreal = normTimer - (weekrealtimer) 
				normbal = normTimer - (weekrest) 
				%>
				<div style="background-color:#ffffff; border:1px #000000 solid;"><%=formatnumber(weekbalance, 2)%></div>
				<div style="background-color:#8caae6;"><%=formatnumber(normreal, 2)%></div>
				<div style="background-color:#ffffff; border:1px #ffffff solid;"><%=formatnumber(normbal, 2)%></div>
				</td>
				<td valign=top width="40" align=center><font class=megetlillesort>
				<div style="background-color:#ffffe1; border:1px #5582d2 solid;">Tot.</div> 
				<div style="background-color:#8caae6;"><%=formatnumber(mdrealtotal, 2)%></div>
				<div style="background-color:#ffffff;"><%=formatnumber(mdrest, 2)%></div>
				<%
				mdbal = (mdrest) - (mdrealtotal)
				mdnormreal = mdnormreal + (normreal)
				mdnormbal = mdnormbal + (normbal)
				%>
				<div style="background-color:#ffffff; border:1px #000000 solid;"><%=formatnumber(mdbal, 2)%></div>
				<div style="background-color:#8caae6;"><%=formatnumber(mdnormreal, 2)%></div>
				<div style="background-color:#ffffff; border:1px #ffffff solid;"><%=formatnumber(mdnormbal, 2)%></div>
				</td>
				<%
				
				totalreal = totalreal + (mdrealtotal)
				totalres = totalres + (mdrest)
				totalbal = totalbal + (mdbal)
				'totalnormbal = totalnormbal + (mdnormbal)
				
				mdrealtotal = 0
				mdrest = 0
				mdbal = 0
				mdnormbal = 0
				mdnormreal = 0
				normTimer = 0
				end if%>
				
				<%
		lastxmid = xmid(x)		
		next%>

</table>
<br>
<table cellspacing=1 cellpadding=0 border=0 width=300 bgcolor="#d6dff5">
<tr>
	<td width=150>&nbsp;</td><td bgcolor="#ffffff" align=center width=75 style="border:1px #8caae6 solid;"><b>I periode</b></td><td width=75 bgcolor="#ffffff" style="border:1px #8caae6 solid;" align=center><b>Ialt</b></td>
</tr>
<tr>
	<td bgcolor="#ffffff">&nbsp;Realiserede timer total:</td><td align=right bgcolor="#8caae6" style="padding-right:5;"><b><%=totalreal%></b></td><td align=right bgcolor="#8caae6" style="padding-right:5;"><%=totalrealtimer%></td>
</tr>
<tr>
	<td bgcolor="#ffffff">&nbsp;Ressource timer tildelt:</td><td align=right bgcolor="#ffffff" style="padding-right:5;"><b><%=totalres%></b></td>


<%
if cint(intJid) <> 0 then
resJidKri = "ressourcer.jobid = "& intJid 
else
resJidKri = "ressourcer.jobid <> "& intJid 
end if

if cint(intMid) <> 0 then
resMidKri = " AND ressourcer.mid = "& intMid
else
resMidKri = " AND ressourcer.mid <> "& intMid
end if

strSQL3 = "SELECT sum(timer) AS restimer FROM ressourcer WHERE "& resJidKri & resMidKri
'Response.write strSQL3
oRec3.open strSQL3, oConn, 3 
if not oRec3.EOF then
restimerialt = oRec3("restimer")
end if
oRec3.close 

if len(restimerialt) <> 0 then
restimerialt = restimerialt
else
restimerialt = 0
end if

if len(totalrealtimer) <> 0 then
totalrealtimer = totalrealtimer
else
totalrealtimer = 0
end if
%><td align=right bgcolor="#ffffff" style="padding-right:5;"><%=restimerialt%></td>
</tr>
<%
ialtBal =  formatnumber(restimerialt - (totalrealtimer), 2)
%> 
<tr>
	<td bgcolor="#ffffff">&nbsp;+ / - Balance: </td>
	<td align=right bgcolor="#ffffff" style="padding-right:5; border:1px #ffffff solid;"><u><b><%=totalbal%></b></u></td>
	<td align=right bgcolor="#ffffff" style="padding-right:5; border:1px #ffffff solid;"><%=ialtBal%></td>
</tr>
<!--<tr>
	<td bgcolor="#ffffff">&nbsp;Norm. balance: </td>
	<td align=right bgcolor="#ffffff" style="padding-right:5; border:1px #ffffff solid;"><=totalnormbal%></td>
</tr>-->
</table>
	<br><br>&nbsp;
	
	
	
	
	
	</div>
	<%
	end select
	
	
	
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
