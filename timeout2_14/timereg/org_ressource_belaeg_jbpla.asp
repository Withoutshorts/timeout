<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/ressource_belaeg_jbpla_inc.asp"-->
<!--#include file="inc/helligdage_func.asp"-->
<!--#include file="inc/isint_func.asp"-->

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
	
	thisfile = "res_belaeg"
	'Response.write "func: " & func
	
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	if len(request("periodeselect")) <> 0 then
	periodeSel = request("periodeselect")
	response.cookies("resbel")("valgtperiode") = periodeSel
	else
		if len(request.cookies("resbel")("valgtperiode")) <> 0 then
		periodeSel = request.cookies("resbel")("valgtperiode")
		else
		periodeSel = 12
		end if
	end if
	
	dim timerTotalTotal, maanedY, medarbTotalTimer
	Redim timerTotalTotal(0), maanedY(0), medarbTotalTimer(0)
	
	select case periodeSel 
	case 1
	periodeShow = "1 Måned (dag/dag)"
	Redim preserve timerTotalTotal(31)
	Redim preserve medarbTotalTimer(31)
	Redim preserve maanedY(31)
	
	case 12
	periodeShow = "12 Måneder frem (måned/måned)"
	Redim preserve timerTotalTotal(12)
	Redim preserve medarbTotalTimer(12)
	Redim preserve maanedY(12)
	
	end select
	
	
	
	
	
	if len(request("showonejob")) <> 0 then
	showonejob = request("showonejob")
	else
	showonejob = 0
	end if
	
	if len(request("inklweekend")) <> 0 then
	weekendOn = request("inklweekend")
	else
	weekendOn = 0
	end if 
	
	if len(request("inklweekend_2")) <> 0 then
	weekend_2_On = request("inklweekend_2")
	else
	weekend_2_On = 0
	end if 
	
	
	if len(request("medarbSel")) <> 0 then 
	medarbSel = request("medarbSel")
	else
	medarbSel = 0
	end if
	
	if len(request("medarbSel")) <> 0 then 
	medarbSel = request("medarbSel")
	response.cookies("resbel")("medarb") = medarbSel
	else
		if len(request.cookies("resbel")("medarb")) <> 0 then
		medarbSel = request.cookies("resbel")("medarb")
		else
		medarbSel = session("mid")
		end if
	end if
	
	'Response.write session("mid")
	'Response.flush
	
	response.cookies("resbel").expires = date + 30
	
	medarbIswrt = ""
	
	if func = "redalle" then
	
						
						
						'*** Datoer ****
						datoer = split(request("FM_dato"), ",")
						
						
						tmedid = split(request("FM_medarbid"), ",")
						tjobid = split(request("FM_jobid"), ",")
						
						tTimertildelt_temp = replace(request("FM_timer"), ", #,", ";")
						tTimertildelt_temp2 = replace(tTimertildelt_temp, ", #", "")
						
						'Response.write request("FM_jobid") & "<br>"
						'Response.write tTimertildelt_temp2
						'Response.flush
						
						tTimertildelt = split(tTimertildelt_temp2, ";") 'split(request("FM_timer"), ";")
						
						'*** Validerer ***
						antalErr = 0
						for y = 0 to UBOUND(tTimertildelt)
						
							
							call erDetInt(SQLBless(trim(tTimertildelt(y))))
							if isInt > 0 then
								antalErr = 1
								errortype = 67
								'useleftdiv = "t"
								%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
								call showError(errortype)
								response.end
							end if	
							isInt = 0
						
						next
						
						for y = 0 to Ubound(datoer)
						
						'Response.write "her<br>"
						
						if cint(instr(medarbIswrt, "#"&tmedid(y)&"#")) = 0 then
								'*** Opdaterer medarbejder TimeStamp ***
								medarbStamp = year(now)&"/"&month(now)&"/"&day(now)
								strSQlmed = "UPDATE medarbejdere SET forecaststamp = '"& medarbStamp &"' WHERE mid = "& tmedid(y)		
								
								oConn.execute(strSQLmed)
								
								medarbIswrt = medarbIswrt & "#"& tmedid(y)&"#,"
						
						end if
							
							
												select case periodeSel
												case 1
												
												tDato = year(datoer(y)) &"/"& month(datoer(y))&"/"& day(datoer(y))
												
												monththis = datepart("m", datoer(y))
												yearthis = datepart("yyyy", datoer(y))
												
												tildeltialt = 0
												
												call opd_weekends_yes_no(tTimertildelt(y))
												
														'*** Renser ud i tabellen (ressourer_md) 
														strSQLdel = "DELETE FROM ressourcer_md WHERE jobid = "& tjobid(y) &" AND medid = "& tmedid(y) &""_
														&" AND md = "& monththis &" AND aar =  "& yearthis 
														
														oConn.execute strSQLdel
														
														'*** Opretter nye records ***
														if tildeltialt <> 0 then
																
																timerThisMD =  sqlbless(tildeltialt)
																
																strSQL = "INSERT INTO ressourcer_md"_
																& " (jobid, medid, md, aar, timer) VALUES ("_
																&" "& tjobid(y) &", "& tmedid(y) &", "& monththis &", "& yearthis &", "& timerThisMD &")" 
																
																oConn.execute strSQL
														end if

												
												case 12
												
												'Response.write datoer(y) & ":&nbsp;"
												monththis = datepart("m", datoer(y))
												yearthis = datepart("yyyy", datoer(y))
												
												
												
												'call opd_ressourcer_md(trim(tTimertildelt(y)))
												
												
												'*** Opretter nye records ***
												
													if len(trim(tTimertildelt(y))) <> 0  then
													
																
																if tTimertildelt(y) > 0 then
																
																'*** Renser ud i tabellen 
																strSQLdel = "DELETE FROM ressourcer_md WHERE jobid = "& tjobid(y) &" AND medid = "& tmedid(y) &""_
																&" AND md = "& monththis &" AND aar =  "& yearthis 
												
																oConn.execute strSQLdel
																
															
																timerThisMD =  sqlbless(tTimertildelt(y))
																
																strSQL = "INSERT INTO ressourcer_md"_
																& " (jobid, medid, md, aar, timer) VALUES ("_
																&" "& tjobid(y) &", "& tmedid(y) &", "& monththis &", "& yearthis &", "& timerThisMD &")" 
																
																oConn.execute strSQL
																
																call dageimd(monththis, yearthis)	
																
																
																'** Wekends og helligdage medregnes? **
																if cint(weekend_2_On) = 0 then 
																nyeTimeprDag = formatnumber(tTimertildelt(y)/workingMthDays, 1) 
																else
																nyeTimeprDag = formatnumber(tTimertildelt(y)/mthDays, 1)
																end if
																
																
																tildeltsofar = 0
																
																	for d = 0 to mthDays - 1
																		
																		tDatotemp = dateadd("d", d, datoer(y))
																		tDato = year(tDatotemp) &"/"& month(tDatotemp)&"/"& day(tDatotemp)
																		
																		
																		'Response.write d &" = "& lastWday & "<br>"
																		if (d = mthDays - 1 AND cint(weekend_2_On) = 1) OR (d + 1 = lastWday AND cint(weekend_2_On) = 0) then
																			nyeTimeprDag = (tTimertildelt(y) - tildeltsofar)
																		else
																			nyeTimeprDag = nyeTimeprDag
																		end if
																		
																		
																		call opd_weekends_yes_no(nyeTimeprDag)
																		
																	next
																	
																else
																		
																		'**** Sletter evt. gamle registreringer (Hvis 0 er angivet) ***
																		strSQLdel = "DELETE FROM ressourcer WHERE jobid = "& tjobid(y) &" AND mid = "& tmedid(y) &" AND MONTH(dato) = '"& monththis &"' AND YEAR(dato) = '"& yearthis &"'"
																		'Response.write strSQLdel
																		oConn.execute(strSQLdel)
														
																end if
															end if
												end select
							
												
						next
	
	end if
	
	
	
	
	
	
	
	'************** Indlæser variable til sidevisning *************

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
	
	
	
	
	
	'*******************************************************************************************
	'***** Finder periode udfra de valgte kriterier ********************************************
	'*******************************************************************************************
	
	
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
	
	
	
	select case periodeSel 
	case 1
	
			call dageimd(monththis,yearthis)	
			mthDaysUse = mthDays
			monthUse = monththis
			yearUse = yearthis
			
	jobStartKri = yearthis&"/"&monththis&"/"&"1"
	jobSlutKri = yearUse&"/"&monthUse&"/"&mthDaysUse
			
	case 3
	
	jobSlutTEMP = dateadd("m", 2, "1/"& monththis &"/"& yearthis) '** år ?
	
	monthUse = datepart("m", jobSlutTEMP)
	yearUse = datepart("yyyy", jobSlutTEMP)
	
	call dageimd(monthUse,yearUse)	
	mthDaysUse = mthDays
	
	stDatoTemp = "1/"&monththis&"/"&yearthis
	select case weekday(stDatoTemp, 2)
	case 1
	stThisDato = dateadd("d", 0, stDatoTemp)
	case 2
	stThisDato = dateadd("d", -1, stDatoTemp)
	case 3
	stThisDato = dateadd("d", -2, stDatoTemp)
	case 4
	stThisDato = dateadd("d", -3, stDatoTemp)
	case 5
	stThisDato = dateadd("d", -4, stDatoTemp)
	case 6
	stThisDato = dateadd("d", -5, stDatoTemp)
	case 7
	stThisDato = dateadd("d", -6, stDatoTemp)
	end select
	jobStartKri = year(stThisDato)&"/"&month(stThisDato)&"/"&day(stThisDato)
	
	
	'*** Sidste dag i ugen ***
	slDatoTemp = mthDaysUse&"/"&monthUse&"/"&yearUse
	'Response.write slDatoTemp &"<br>"
	select case weekday(stDatoTemp, 2)
	case 1
	slThisDato = dateadd("d", 6, slDatoTemp)
	case 2
	slThisDato = dateadd("d", 5, slDatoTemp)
	case 3
	slThisDato = dateadd("d", 4, slDatoTemp)
	case 4
	slThisDato = dateadd("d", 3, slDatoTemp)
	case 5
	slThisDato = dateadd("d", 2, slDatoTemp)
	case 6
	slThisDato = dateadd("d", 1, slDatoTemp)
	case 7
	slThisDato = dateadd("d", 0, slDatoTemp)
	end select
	jobSlutKri = year(slThisDato)&"/"&month(slThisDato)&"/"&day(slThisDato)
	
	'Response.write jobSlutKri 
	
	
	
	
	case 12
	
	jobSlutTEMP = dateadd("m", 11, "1/"& monththis &"/"& yearthis) '** år ?
	
	monthUse = datepart("m", jobSlutTEMP)
	yearUse = datepart("yyyy", jobSlutTEMP)
	
	call dageimd(monthUse,yearUse)	
	mthDaysUse = mthDays
	
	jobStartKri = yearthis&"/"&monththis&"/"&"1"
	jobSlutKri = yearUse&"/"&monthUse&"/"&mthDaysUse
	
	end select
	
	
	
	daynow = day(now)&"/"&month(now)&"/"&year(now)
	'***************************************************************************
	
	
	
	
	'Response.write jobStartKri &  " jobSlutKri " & jobSlutKri
	'Response.flush
	
	
		'*** Jobnavn kriterie ***		
		
		if showonejob = 1 then '*** Job/milepæle ***
		
			jobNavnKri = " AND j.id = "& id
		
		else
			
			if trim(request("jobselect")) <> "(Alle)" then
			jobnavn = request("jobselect")
			else
			jobnavn = ""
			end if
			
			jobNavnKri = " AND j.jobnavn LIKE '"& jobnavn &"%' "
			
			if len(request("jobselect")) <> 0 then
			jobnavn = request("jobselect")
			else
			jobnavn = "(Alle)"
			end if
			
		
		end if
		
	
	
		
	'*************************************************************************************	
		
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
	
	function showtildeltimer(thisMid, dato, jobid, persel) {
	document.getElementById("tildeltimer").style.visibility = "visible"
	document.getElementById("tildeltimer").style.display = ""
	document.getElementById("FM_tildeltimer_mid").value = thisMid 
	//document.getElementById("FM_dato_sel").value = dato
	
	document.getElementById("chkboxes").innerHTML = dato
	 
	
	//document.getElementById("FM_dato").style.visibility = "hidden"
	//document.getElementById("FM_dato").style.display = "none"
	
	//if (persel =! 1) {
	//document.getElementById("ugetxt").style.visibility = "visible"
	//document.getElementById("ugetxt").style.display = ""
	//}
	
	
			for(i=0;i<document.forms["timertildel"]["FM_tildeltimer_job"].length;i++){
				if(document.forms["timertildel"]["FM_tildeltimer_job"][i].value == jobid){
					document.forms["timertildel"]["FM_tildeltimer_job"][i].selected = true;
				}
			}
		
	}
	
	
	function checkAll(field) {
	field.checked = true;
	for (i = 0; i < field.length; i++)
		field[i].checked = true ;
	}
	
	function unCheckAll(field) {
	field.checked = true;
	for (i = 0; i < field.length; i++)
		field[i].checked = false ;
	}
	
	function hidetildeltimer() {
	document.getElementById("tildeltimer").style.visibility = "hidden"
	document.getElementById("tildeltimer").style.display = "none"
	}
	
	
	function BreakItUp()
	{
	  //Set the limit for field size.
	  var FormLimit = 102399
	
	  //Get the value of the large input object.
	  var TempVar = new String
	  TempVar = document.theForm.BigTextArea.value
	
	  //If the length of the object is greater than the limit, break it
	  //into multiple objects.
	  if (TempVar.length > FormLimit)
	  {
	    document.theForm.BigTextArea.value = TempVar.substr(0, FormLimit)
	    TempVar = TempVar.substr(FormLimit)
	
	    while (TempVar.length > 0)
	    {
	      var objTEXTAREA = document.createElement("TEXTAREA")
	      objTEXTAREA.name = "BigTextArea"
	      objTEXTAREA.value = TempVar.substr(0, FormLimit)
	      document.theForm.appendChild(objTEXTAREA)
	
	      TempVar = TempVar.substr(FormLimit)
	    }
	  }
	}
	
	
	</script>

	<%
	intMid = 0
	intJid = 0
	
	startdato = jobStartKri
	slutdato =  jobSlutKri
	
	%>
	
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%if showonejob <> 1 then%>
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
	topdiv = 132
	leftdiv = 20
	else
	topdiv = 20
	leftdiv = 20
	end if
	%>
	
	<!-------------------------------Sideindhold------------------------------------->
	
	<div id="overblik" style="position:absolute; left:<%=leftdiv%>; top:<%=topdiv%>; visibility:visible;">
	<%
	
	
	toppxThisDiv = 110
	leftpxThisDiv = 90
	
	select case periodeSel
	case 1
	lav1 = 0
	lav2 = 1
	mellem1 = 1
	mellem2 = 4
	hoj1 = 4 
	
	case 12
	lav1 = 0
	lav2 = 40
	mellem1 = 40
	mellem2 = 80
	hoj1 = 80 
	end select
	
	%>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv-90%>px; background-color:LightGreen; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv-75%>px; padding:2px;">&nbsp; <%=lav1%> - <%=lav2%> time tildelt</div>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv+30%>px; background-color:#cccccc; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv+50%>px; padding:2px;">&nbsp;<%=mellem1%> - <%=mellem2%> tildelt</div>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv+125%>px; background-color:PeachPuff; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv+145%>px; padding:2px;">&nbsp;> <%=hoj1%> timer tildelt.</div>		

	
	
	
	<div style="position:relative; background-color:#ffffff; border:1px #8caae6 solid; border-bottom:2px #8caae6 solid; border-right:2px #8caae6 solid; padding:5px;">
	<table cellspacing="0" cellpadding="0" border="0">
	<form action="ressource_belaeg_jbpla.asp?showonejob=<%=showonejob%>&id=<%=id%>" method="post" name="mdselect" id="mdselect">
	<tr>
	
	<td><b>Medarbejder:</b>&nbsp;
	<%
	strSQL = "SELECT mnavn, mid, mnr FROM medarbejdere m WHERE m.mid <> 0 AND mansat <> 2 ORDER BY mnavn"
	'Response.write strSQL
	'Response.flush
	oRec.open strSQL, oConn, 3 
	%><select name="medarbSel" id="medarbSel" style="width:220px;">
	<option value="0">Alle</option>
	<%
	while not oRec.EOF 
	
	 if cint(medarbSel) = oRec("mid") then
	 mSEL = "SELECTED"
	 else
	 mSEL = ""
	 end if %>
	 
	<option value="<%=oRec("mid")%>" <%=mSEL%>><%=oRec("mnavn")%>&nbsp;(<%=oRec("mnr")%>)</option>
	<%
	oRec.movenext
	wend
	oRec.close %>
	</select>&nbsp;&nbsp;</td>
	
	<%if showonejob = 1 then
		strSQL = "SELECT jobnavn, jobnr FROM job WHERE id = "& id
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then 
			jobnavn = oRec("jobnavn")
		end if
		oRec.close 
	end if
	%>
	
	<%if showonejob <> 1 then%>
	<td><b>job:</b>&nbsp;&nbsp;</td>
	<td>
	<input type="text" name="jobselect" id="jobselect" size="25" value="<%=jobnavn%>" onFocus="clearJnavn()">
	<%else%>
	<td>&nbsp;</td>
	<td>
	<input type="hidden" name="jobselect" id="jobselect" value="<%=jobnavn%>"><b><%=jobnavn%></b>
	<%end if%>
	&nbsp;fra: 
	<select name="mdselect">
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
	<option value="2006">2006</option>
	<option value="2007">2007</option>
	<option value="2008">2008</option>
	<option value="2009">2009</option>
	<option value="2010">2010</option>
	<option value="2011">2011</option>
	<option value="2012">2012</option>
	</select> og 
	<select name="periodeselect" style="font-size:9px; width:190px;">
	<option value="<%=periodeSel%>" SELECTED><%=periodeShow%></option>
	<option value="1">1 Måned (dag/dag)</option>
	<!--<option value="3">3 Måneder frem (uge/uge)</option>
	<option value="4">4 Måneder frem (måned/måned)</option>-->
	<option value="12">12 Måneder frem (måned/måned)</option>
	<!--<option value="6">6 Måneder frem</option>-->
	</select></td>
	</tr>
	<tr>
		<td align=right colspan=4 style="padding-top:4px;"><input type="submit" value="Vis valgte kriterier!"></td>
	</tr>
	
	<input type="hidden" name="selectminus" id="selectminus" value="0">
	<input type="hidden" name="selectplus" id="selectplus" value="0">
	
	<%if showonejob <> 1 then
	knapperTop = topdiv-49
	else
	knapperTop = topdiv+10
	end if%>
	
	<div style="position:absolute; left:708; top:<%=knapperTop%>; z-index:60000; width:150px;">
	<input type="submit" name="submitplus" id="submitplus" value="<- <%=monthname(month(dateadd("m", -1, "1/"&monththis&"/"&yearthis)))%>" style="font-size:9px; border:1px #003399 solid; background-color:#ffffff; width:75px;" onClick="selminus()">
	</div>
	<div style="position:absolute; left:808; top:<%=knapperTop%>; z-index:60000; width:150px;">
	<input type="submit" name="submitplus" id="submitplus" value="<%=monthname(month(dateadd("m", +1, "1/"&monththis&"/"&yearthis)))%> ->" style="font-size:9px; border:1px #003399 solid; background-color:#ffffff; width:75px;" onClick="seladd()">
	</div>
	</form>
	</table>
	</div>
	
	
	</div>
	
	
	<div id="resbelaeg" style="position:absolute; left:20; top:<%=topdiv+60%>; visibility:visible; z-index:10000;">
	<br><h3><img src="../ill/ac0020-24.gif" width="24" height="24" alt="" border="0">&nbsp;Ressourcetimer på medarbejdere.</h3>
	
	
	
	
	
	<%
	
	skrivCsvFil = 1 
	dim csvTxtTop, csvTxt
	public csvDatoer(35)
	
	if skrivCsvFil = 1 then
		csvTxtTop = "Medarbejder;Initialer;Normeret uge;Job;Job nr.;Kunde;Kunde nr." 
		csvTxtTop = csvTxtTop &";Måned;Forecast;Realiseret;Afvigelse %"
	end if
	
	
	dt = 2000
	'mi = 50 
	ji = 3
	
	dim midjiddato 
	redim midjiddato(dt, ji) 
	
	dim showrestimer
	dim showresdato
	dim useJobnr
	dim useJobnavn
	dim xmid
	dim mnavn, minit, mnr
	dim useJobid
	'dim midjiddato
	dim knavn, knr, forecaststamp
	
	
	Redim showrestimer(0), forecaststamp(0)
	Redim showresdato(0)
	Redim useJobnr(0)
	Redim useJobnavn(0)
	Redim xmid(0)
	Redim mnavn(0), mnr(0), minit(0)
	Redim useJobid(0)
	Redim knavn(0), knr(0)
	'Redim midjiddato(0)
	
	lastYear = 0
	lastWeek = 0
	lastMid = 0
	lastJid = 0
	n = 0
	
	'realtimer = 0
	'lastweek = 0
	lastmedarbId = 0
	
	if cint(medarbSel) = 0 Then
	strMedidSQL = " m.mid <> 0 "
	else
	strMedidSQL = " m.mid = " & medarbSel
	end if
	
	
	
	'************************************************************************
	'*** Main SQL ***
	'************************************************************************
	strSQL = "SELECT jobnavn, mnavn, mnr, j.id AS jobid, j.jobnr, medarbejdertype, "_
	&" m.mid, r.timer AS restimer, r.dato AS resdato, r.id AS resid, k.kkundenavn, k.kkundenr, m.init, m.forecaststamp FROM medarbejdere m"_
	&" LEFT JOIN ressourcer r ON (r.mid = m.mid AND r.dato BETWEEN '"& startdato &"' AND '"& slutdato &"')"_
	&" LEFT JOIN job j ON  (j.id = r.jobid "& jobNavnKri &")"_
	&" LEFT JOIN kunder k ON (kid = j.jobknr)"_
	&" WHERE "& strMedidSQL &" AND mansat <> 2 ORDER BY m.mid, j.id, r.dato"
	
	'Response.write strSQL &"<br><br><br>"
	'Response.flush
	
	oRec.open strSQL, oConn, 3 
	x = 0
	
	while not oRec.EOF 
	Redim preserve showresdato(x)
	Redim preserve useJobnr(x)
	Redim preserve useJobnavn(x)
	Redim preserve xmid(x)
	Redim preserve mnavn(x)
	Redim preserve mnr(x)
	Redim preserve minit(x)
	Redim preserve useJobid(x)
	Redim preserve showrestimer(x)
	Redim preserve knavn(x), knr(x)
	Redim preserve forecaststamp(x)
	
	useJobid(x) = oRec("jobid")
	useJobnr(x) = oRec("jobnr")
	useJobnavn(x) = oRec("jobnavn")
	xmid(x) = oRec("mid")
	mnavn(x) = oRec("mnavn")
	showresdato(x) = oRec("resdato")
	showrestimer(x) = oRec("restimer")
	knavn(x) = oRec("kkundenavn")
	knr(x) = oRec("kkundenr")
	mnr(x) = oRec("mnr")
	minit(x) = oRec("init")
	forecaststamp(x) = oRec("forecaststamp")
	
	select case periodeSel 
	case 1
			
			
			if len(oRec("resid")) <> 0 AND len(oRec("jobid")) <> 0 then
			n = n + 1
			midjiddato(n, 0) = showrestimer(x) 
			midjiddato(n, 1) = useJobid(x) 
			midjiddato(n, 2) = xmid(x) 
			midjiddato(n, 3) = showresdato(x) 
			end if
	
	
	
	case 12
		
			
			if len(oRec("resid")) <> 0 AND len(oRec("jobid")) <> 0 then
		
				if cdbl(lastMid) = oRec("mid") AND cdbl(lastJid) = oRec("jobid") AND cint(lastMonth) = datepart("m", oRec("resdato")) AND datepart("yyyy", oRec("resdato")) = lastYear then
				
					midjiddato(n, 0) = midjiddato(n, 0) + showrestimer(x)
					 
				else
					
					n = n + 1
					midjiddato(n, 0) = showrestimer(x) 
					midjiddato(n, 1) = useJobid(x) 
					midjiddato(n, 2) = xmid(x) 
					midjiddato(n, 3) = showresdato(x) 
					
				end if
				
			lastYear = datepart("yyyy", oRec("resdato")) 
			lastMonth = datepart("m", oRec("resdato"), 2 ,3) 
			lastMid = oRec("mid")
			lastJid = oRec("jobid")
		end if
	
		
	end select
									
									
	
	x = x + 1
	oRec.movenext
	wend
	oRec.close 
	
		
	
	%>
	


<form action="ressource_belaeg_jbpla.asp?showonejob=<%=showonejob%>&id=<%=id%>&medarbSel=<%=medarbSel%>" method="post" name="alledagemedarb" id="alledagemedarb">
<input type="hidden" name="jobselect" id="jobselect" value="<%=jobnavn%>">
<input type="hidden" name="mdselect" id="mdselect" value="<%=monththis%>">
<input type="hidden" name="aarselect" id="aarselect" value="<%=yearthis%>">
<input type="hidden" name="periodeselect" id="periodeselect" value="<%=periodeSel%>">
<%if periodeSel <> 12 then%>
<input type="hidden" name="inklweekend_2" id="inklweekend_2" value="1"><img src="../ill/blank.gif" width="770" height="1" alt="" border="0">
<%else%>
<br>
<input type="Checkbox" name="inklweekend_2" id="inklweekend_2" value="1"> <b>Opdatér også lør, søn og hellig -dage.</b>

<%end if%>

<input type="hidden" name="func" id="func" value="redalle" checked><!-- <b>Opdatér felter enkeltvis.</b>--> 
<table border=0 cellspacing=1 cellpadding=0 bgcolor="#8caae6">

<%
select case periodeSel 
case 1
numoffdaysorweeksinperiode = datediff("d", startdato, slutdato)
case 12
numoffdaysorweeksinperiode = datediff("m", startdato, slutdato)
end select

'numoffdaysorweeksinperiode = datediff("d", startdato, slutdato)
y = 0 
nday = startdato
weekrest = 0		
antalxx = x
xx = 0
foundone = 0
daynowrpl = replace(daynow, "/", "-")
lastMonth = monththis



call datoeroverskrift


'useTopX = (x - 1)



strjIdWrt = "0#," '** Bruges kun når der er valgt en enkelt medarbejder


'**** Medarbejdere ****
for x = 0 to x - 1

	
		
	
	if lastxmid <> xmid(x) then	 
		
				select case periodeSel 
				case "1"
				tdcolspan = 2
				case "12"
				tdcolspan = 5
				end select
	
	strDageChkboxOne = "<input type=text size=10 name=FM_sel_dage id=FM_sel_dage value="&daynowrpl&">&nbsp;dd-mm-aaaa"
	
	if x > 0 then				
		call medarbtotal
	end if%>
	
	
	<tr>
		<td bgcolor="#ffffe1" height=20 style="padding-left:5px; padding-top:5px;  padding-bottom:5px;">
		<a href="javascript:NewWin_todo('medarb_restimer.asp?menu=job&func=show&FM_medarb=<%=xmid(x)%>&FM_mednavn=<%=mnavn(x)%>')" class=vmenu><%=mnavn(x)%> (<%=mnr(x)%>)</a>&nbsp;
		<!--<a href="#" onClick="showtildeltimer('<=xmid(x)%>','<=strDageChkboxOne%>','<=id%>','<=periodeSel%>')" class=vmenuglobal>Ny</a>-->
		<%if len(trim(minit(x))) <> 0 then
		Response.write " - " & minit(x)
		end if
		%><br>
		<font class=megetlillesilver>Sidst opd: <%=forecaststamp(x)%></font>
		</td>
		
		<td colspan=4 bgcolor="#eff3ff">&nbsp;</td>
		
		<%
		select case periodeSel
		case 1
		cspan = 32
		case 12
		cspan = 12
		end select
		%>
		
		<td bgcolor="#ffffe1" height=20 align=right style="padding-right:5px;" colspan=<%=cspan%>>Normeret uge: <%
		call normtimerPer(xmid(x))
		Response.write formatnumber(ntimPer, 1)
		
		%>
		&nbsp;
		</td>
	</tr>
	<%
	lastJid = 0
	end if
	
		
		if lastJid <> useJobid(x) OR lastxmid <> xmid(x) AND len(useJobid(x)) <> 0 then%>
		<tr>
		<td bgcolor="#ffffff" style="padding-left:5px; padding-right:5px;" width=200><b><%=useJobnavn(x)%> (<%=useJobnr(x)%>)</b><br>
		<font class=megetlillesilver><%=knavn(x)%> (<%=knr(x)%>)</font></td>
			
		<td valign=top bgcolor="#eff3ff" align=right class=lille style="padding-top:1px; padding-right:5px; border-left:2px #8caae6 solid;">
		Forecast:<br>
		Realiseret:<br>
		Afvigelse:
		
		</td>
				<%
				
				for s = 0 to 2
					
						
						
						select case s
						case 0
						statDatoUse = statdatoM3
						bdleft = 0
						bdright = 0
						thisy = 35
						case 1
						statDatoUse = statdatoM2
						bdleft = 0
						bdright = 0
						thisy = 34
						case 2
						statDatoUse = statdatoM1 
						bdleft = 0
						bdright = 2
						thisy = 33
						end select
					
					%>
					<td valign=top bgcolor="#eff3ff" align=right class=lille style="padding-top:1px; padding-right:5px;  border-right:<%=bdright%> #8caae6 solid; border-left:<%=bdleft%> #8caae6 solid;">
					<%call tildelttimerPer(xmid(x), useJobid(x), statDatoUse, s)%>
					<%=formatnumber(tildtimPer, 1)%><br>
					<%
					
					
					sumTimer = 0
					
					strSQLtimer = "SELECT sum(timer) AS sumtimerFaktisk FROM timer WHERE tmnr = "& xmid(x) &" AND tjobnr = "&useJobnr(x)&" AND (YEAR(tdato) = "& year(statDatoUse) &" AND MONTH(tdato) = "& month(statDatoUse) &")"
					oRec3.open strSQLtimer, oConn, 3
					while not oRec3.EOF
					
					sumTimer = oRec3("sumtimerFaktisk")
					
					oRec3.movenext
					wend
					
					oRec3.close
					
					
					if len(trim(sumTimer)) <> 0 then
					sumTimer = sumTimer 
					else
					sumTimer = 0
					end if%>
					
					<%=formatnumber(sumTimer, 1)%><br>
					
					<%
					
					
					
					'*** Afvigelse ***
					afv = 0
					if sumTimer <> 0 then
						
						if tildtimPer = 0 then
						afv = sumTimer
						else
						afv = 100 -((tildtimPer/sumTimer) * 100)
						end if
					
					'Response.write afv & "<br>"
					
					'if afv >= 100 then
					'afv = afv - 100
					'else
					'afv = 100 - afv
					'end if
					
					else
							if tildtimPer = 0 then
							afv = 0
							else
							afv = tildtimPer '100
							end if
					end if '(tildtimPer - (sumTimer))
					
					Response.write formatnumber(afv, 0) &"%"
					
					
					
					csvTxt = csvTxt & vbcrlf & mnavn(x) &" ("& mnr(x) &")"
					csvTxt = csvTxt & ";"& minit(x) 
					csvTxt = csvTxt & ";"& formatnumber(ntimPer, 2) 
					csvTxt = csvTxt & ";"& useJobnavn(x) & ";" & useJobnr(x) 
					csvTxt = csvTxt & ";"& knavn(x) & ";" & knr(x)
				
					'csvTxt = csvTxt & vbcrlf &";;;;;;;"& csvDatoer(thisy) 
					csvTxt = csvTxt &";"& csvDatoer(thisy) 
					csvTxt = csvTxt & ";" & formatnumber(tildtimPer, 2)
					csvTxt = csvTxt & ";" & formatnumber(sumTimer, 2)
					csvTxt = csvTxt & ";" & formatnumber(afv, 0)
					
					
						select case s
						case 0
						totFaktisk3 = totFaktisk3 + sumTimer
						case 1
						totFaktisk2 = totFaktisk2 + sumTimer
						case 2
						totFaktisk1 = totFaktisk1 + sumTimer
						end select
					%>
					</td>
				<%next%>
				
				
		
		<%
		strjIdWrt = strjIdWrt & "#"&useJobid(x)&"#," 
		lastJid = useJobid(x)
	
					lastMonth = monththis
					y = 0
					for y = 0 TO numoffdaysorweeksinperiode
					
						call antaliperiode(periodeSel, y, startdato, monthUse)
					
								select case periodeSel 
								case 1
									timerThis = 0	
									For n = 0 to UBOUND(midjiddato, 1)
										if cdate(midjiddato(n, 3)) = cdate(nday) AND midjiddato(n, 1) = useJobid(x) AND midjiddato(n, 2) = xmid(x) then
										timerThis = midjiddato(n, 0)
										foundone = 1
										end if
									next
								
								
								
								case 12
									
									timerThis = 0	
									For n = 0 to UBOUND(midjiddato, 1)
										if (datepart("m", midjiddato(n, 3)) = datepart("m", nday)) AND midjiddato(n, 1) = useJobid(x) AND midjiddato(n, 2) = xmid(x) AND datepart("yyyy", midjiddato(n, 3)) = datepart("yyyy", nday) then
										timerThis = midjiddato(n, 0)
										foundone = 1
										end if
									next
									
								end select
								
						foundone = 0
						xx = 0
					
					
					if timerThis = 0 then
					bgTimer = "#ffffff"
					end if
					
					if timerThis > lav1 AND timerThis <= lav2 then
					bgTimer = "LightGreen"
					end if
					
					if timerThis > mellem1 AND timerThis < mellem2 then
					bgTimer = "#cccccc"
					end if
					
					if timerThis => hoj1 then
					bgTimer = "PeachPuff"
					end if
					
					
					
					
					lastMonth = newMonth
					
					dato = nday
					
					
					select case periodeSel 
					case 1
					chkSELdg = ""
					chkSELdgW = ""
					
					strDageChkbox = "<input type=checkbox name=FM_sel_dage id=FM_sel_dage value="&dato&" CHECKED>" & dato
					
					stThisDato = dato
					txtwidth = 18
					
					
					case 12
					
					stThisDato = dato
					txtwidth = 35
					
					end select
					
					
					
					if timerThis = 0 then
					showTimerThis = ""
					else
					showTimerThis = timerThis
					end if
					%>
					<td bgcolor="<%=bgTimer%>" valign=top align=center width="<%=txtwidth+4%>" style="padding-right:3px;">
					<%if medarbSel <> 99999999999999 then%>
					<input type="hidden" name="FM_jobid" id="FM_jobid" value="<%=useJobid(x)%>">
					<input type="hidden" name="FM_medarbid" id="FM_medarbid" value="<%=xmid(x)%>">
					<input type="hidden" name="FM_dato" id="FM_dato" value="<%=stThisDato%>">
					
					<%if PeriodeSel <> 1 then%>
					<%=showTimerThis%><br><input type="text" name="FM_timer" id="FM_timer" style="width:<%=txtwidth%>px; font-family:arial; font-size:9px;" value="">
					<%else%>
					<input type="text" name="FM_timer" id="FM_timer" style="width:<%=txtwidth%>px; font-size:8px; font-family:arial;" value="<%=showTimerThis%>">
					<%end if
					
					'*** hidden array felt. *** Bruges så der kan skelneshvis man bruger komma.%>
					<input type="hidden" name="FM_timer" id="FM_timer" style="width:<%=txtwidth%>px; font-size:9px;" value="#">
					
					
					<%else%>
					<a href="#" onClick="showtildeltimer('<%=xmid(x)%>','<%=strDageChkbox%>','<%=lastJid%>', '<%=periodeSel%>')" class=vmenulille><%=timerThis%></a>
					<%end if
					
					csvTxt = csvTxt & vbcrlf & mnavn(x) &" ("& mnr(x) &")"
					csvTxt = csvTxt & ";"& minit(x) 
					csvTxt = csvTxt & ";"& formatnumber(ntimPer, 2) 
					csvTxt = csvTxt & ";"& useJobnavn(x) & ";" & useJobnr(x) 
					csvTxt = csvTxt & ";"& knavn(x) & ";" & knr(x)
					csvTxt = csvTxt & ";"& csvDatoer(y) &";" & timerThis
					'csvTxt = csvTxt & vbcrlf & ";;;;;;;"& csvDatoer(y) &";" & timerThis%>
					</td>
					
					<%
					medarbTotalTimer(y) = medarbTotalTimer(y) + timerThis
					timerTotalTotal(y) = timerTotalTotal(y) + timerThis
					next%>
				
				
		<%
		'csvTxt = csvTxt & vbcrlf
		end if '** Lastjobid
		
		lastxmid = xmid(x)
		%>
		</tr>
		<%	
		'csvTxt = csvTxt & vbcrlf	
		next
		
		
		
		'call medarbTotal
		
		if medarbSel <> 0 then
		usemrn = medarbSel
		%>
		
		<%call medarbTotal%>
		</table>
		<br>
		<table cellspacing=0 cellborder=0 border=0>
		<tr><td><img src="../ill/blank.gif" width="750" height="1" alt="" border="0"><input type="submit" value="Indlæs forecast!"></td></tr>
		</form>
		</table>
		
		<br><br>
		<table border=0 cellspacing=1 cellpadding=0 bgcolor="#8caae6">
		
		<tr><td colspan=36 bgcolor="#ffffff" style="padding:10px;">
		<img src="../ill/ac0052-24.gif" width="24" height="24" alt="Se Guiden dine aktive job (Skjulte job)" border="0">&nbsp;<a href="javascript:popUp('guiden_2006.asp?mid=<%=usemrn%>&lc=res','600','500','150','120');" target="_self"; class=vmenu>Guiden dine aktive job</a>
		<br><b>Dine aktive job. (Uden tildelte ressourcetimer.)</b></td></tr>
		
		<%
		skrivCsvFil = 0
		call datoeroverskrift
		
		
		if medarbSel <> 0 then
			call hentbgrppamedarb(medarbSel)
		else
			strPgrpSQLkri = ""
		end if
		
		
		
		'******************************************************
		'**** Henter job fra Guiden Dine aktive job ***********
		varUseJob = "("
		
		strSQL4 = "SELECT id, medarb, jobid FROM timereg_usejob WHERE "& varAktivejob &" medarb = "& usemrn &""
		oRec3.open strSQL4, oConn, 3
		
		while not oRec3.EOF 
		varUseJob = varUseJob & " j.id = "& oRec3("jobid") & " OR "
		oRec3.movenext
		wend 
		
		oRec3.close 
		
		if varUseJob = "(" then
		varUseJob = " j.id = 0 AND "
		else
		varUseJob_len = len(varUseJob)
		varUseJob_left = varUseJob_len - 3
		varUseJob_use = left(varUseJob, varUseJob_left) 
		varUseJob = varUseJob_use & ") AND "
		end if
		
		
		
		'************** Main SQL Call (Ikke angivne ressource timer) ******************************************
		strSQL = "SELECT j.id AS id, jobnr, jobnavn, jobknr, kkundenavn, kkundenr, count(a.id) AS antalakt,"_
		&" j.budgettimer, j.ikkebudgettimer, "_
		&" kid, jobans1, jobans2 FROM job j, kunder k"_
		&" LEFT JOIN aktiviteter a ON (a.job = j.id)"
		
		strSQL = strSQL &" WHERE ("& varUseJob &" j.jobstatus = 1 AND j.fakturerbart = 1 AND k.Kid = j.jobknr "& jobNavnKri &") "& strPgrpSQLkri & " GROUP BY j.id, a.job ORDER BY k.kkundenavn, j.jobnavn, j.id, j.jobnavn" 
		
		'Response.write strSQL
		
		oRec.open strSQL, oConn, 3 
		j = 0
		while not oRec.EOF 
		
		
		if instr(strjIdWrt, "#"&oRec("id")&"#") = 0 then%>
		<tr>
		<td bgcolor="#ffffff" style="padding-left:5px; padding-right:5px;" width=200><b><%=oRec("jobnavn")%> (<%=oRec("jobnr")%>) </b><br>
		<font class=megetlillesilver><%=oRec("kkundenavn")%> ( <%=oRec("kkundenr")%>)</font></td>
					
					
		<td valign=top bgcolor="#eff3ff" align=right class=lille style="padding-top:1px; padding-right:5px; border-left:2px #8caae6 solid;">
		Forecast:<br>
		Realiseret:<br>
		Afvigelse:</td>
					
					<%
					for s = 0 to 2
					
					select case s
					case 0
					statDatoUse = statdatoM3
					bdleft = 0
					bdright = 0
					case 1
					statDatoUse = statdatoM2
					bdleft = 0
					bdright = 0
					case 2
					statDatoUse = statdatoM1 
					bdleft = 0
					bdright = 2
					end select
					%>
					<td valign=top bgcolor="#eff3ff" align=right class=lille style="padding-top:1px; padding-right:5px; border-right:<%=bdright%> #8caae6 solid; border-left:<%=bdleft%> #8caae6 solid;">
					
					<%call tildelttimerPer(usemrn, oRec("id"), statDatoUse, s)%>
					<%=formatnumber(tildtimPer, 1)%><br>
					
					<%
					sumTimer = 0
					
					
					'**** Realiseret ****
					strSQLtimer = "SELECT sum(timer) AS sumtimerFaktisk FROM timer WHERE tmnr = "& usemrn &" AND tjobnr = "& oRec("id") &" AND (YEAR(tdato) = "& year(statDatoUse) &" AND MONTH(tdato) = "& month(statDatoUse) &")"
					'Response.write strSQLtimer &"<br>"
					'Response.flush 
					oRec3.open strSQLtimer, oConn, 3
					while not oRec3.EOF
					
					sumTimer = oRec3("sumtimerFaktisk")
					
					oRec3.movenext
					wend
					
					oRec3.close
					
					
					if len(trim(sumTimer)) <> 0 then
					sumTimer = sumTimer 
					else
					sumTimer = 0
					end if%>
					
					<%=formatnumber(sumTimer, 1)%><br>
					
					<%
					'*** Afvigelse ***
					afv = 0
					if sumTimer <> 0 then
					
						if tildtimPer = 0 then
						afv = sumTimer
						else
						afv = 100 -((tildtimPer/sumTimer) * 100)
						end if
					
						'if afv >= 100 then
						'afv = afv - 100
						'else
						'afv = 100 - afv
						'end if
						
					else
							if tildtimPer = 0 then
							afv = 0
							else
							afv = tildtimPer '100
							end if
					end if '(tildtimPer - (sumTimer))
					
					Response.write formatnumber(afv, 0) &"%"
					
					
					select case s
					case 0
					totFaktisk3 = totFaktisk3 + sumTimer
					case 1
					totFaktisk2 = totFaktisk2 + sumTimer
					case 2
					totFaktisk1 = totFaktisk1 + sumTimer
					end select
					
					%>
					
					</td>
					<%next%>
					
					
					
					<%
					select case periodeSel 
					case 1
					txtwidth = 18
					case 12
					txtwidth = 35
					end select 
					
					for y = 0 TO numoffdaysorweeksinperiode
							
							call antaliperiode(periodeSel, y, startdato, monthUse)
							
					
					
					
					lastMonth = newMonth
					%>	
							
					<td bgcolor="#ffffff" width="<%=txtwidth+4%>" valign=top align=right style="padding-right:3px;">
					<input type="hidden" name="FM_jobid" id="FM_jobid" value="<%=oRec("id")%>">
					<input type="hidden" name="FM_medarbid" id="FM_medarbid" value="<%=medarbSel%>">
					<input type="hidden" name="FM_dato" id="FM_dato" value="<%=nday%>">
					&nbsp;<br>
					
					<%if PeriodeSel <> 1 then%>
					<input type="text" name="FM_timer" id="FM_timer" style="width:<%=txtwidth%>px; font-family:arial; font-size:9px;" value="">
					<%else%>
					<input type="text" name="FM_timer" id="FM_timer" style="width:<%=txtwidth%>px; font-size:8px; font-family:arial;" value="">
					<%end if
					%>
					
					<input type="hidden" name="FM_timer" id="FM_timer" value="#">
					</td>
					<%next%>
		
			</tr>
		
		<%
		j = j + 1
		end if
		
		
		oRec.movenext
		wend
		oRec.close
		
		
		'call medarbTotal
		
		
		
		if j = 0 then
		%>
		<tr>
		<td bgcolor="#ffffff" style="padding-left:5px; padding-right:5px;" width=200>&nbsp;
		</td>
		<td bgcolor="#ffffff" height=50 colspan=<%=cspan%>>&nbsp;(Ingen Job)</td>
		</tr>
		<%
		end if
		else
		%>
		<tr>
			<td height=20 bgcolor="#ffffff" align=right style="padding-right:15px;"><b>Total:</b> (Alle job / medarbejdere)</td>
			<td colspan=4 bgcolor="#ffffff">&nbsp;</td>
				<%
				for y = 0 TO numoffdaysorweeksinperiode
					
					'if maanedY(y) = 1 AND y > 0 then  
					'Response.write "<td bgcolor='#ffffe1'>&nbsp;</td>"
					'end if%>
					
				<td bgcolor="#ffffff" align=center><b><%=timerTotalTotal(y)%></b></td>
				<%
				next
				%>
		</tr>
		<%
		end if
		%>
		
		
		
		
		
		
		
		</table>
		<br>
		<table cellspacing=0 cellborder=0 border=0>
		<tr><td><img src="../ill/blank.gif" width="750" height="1" alt="" border="0"><input type="submit" value="Indlæs forecast!"></td></tr>
		</form>
		</table>
		
		
		<form action="ressourcer_eksport.asp" method="post" name=theForm onsubmit="BreakItUp()"> <!--  -->
			<input type="hidden" size=200 name="txt1" id="txt1" value="<%=csvTxtTop%>">
			<input type="hidden" name="BigTextArea" id="BigTextArea" value="<%=csvTxt%>">
			<!--<br><textarea cols="60" rows="10" name="sdf"><=csvTxt%></textarea>-->
			<input type="hidden" name="txt20" id="txt20" value="<%=csvTxtTotal%>">
			
			<input type="submit" value="Eksporter Forecast."><br>
			(kun job med tildelte timer i den valgte periode bliver eksporteret.)
			
			</form>
			
			
			<div style="position:relative; background-color:#eff3ff; visibility:visible; border:1px red dashed; border-right:2px red dashed; border-bottom:2px red dashed; padding:15px; width:450px;"><img src="../ill/ac0005-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Side note(r):<br></b>
	- Tildel timer ved at vælge den ønskede medarbejder.<br>
	- 1 måneds oversigt:<u> Timer angives antal timer pr. dag.</u><br>
	- 12 måneders oversigt:<u> Timer angives som samlet antal timer pr. måned.</u><br>
	- Slet forecast ved at indtaste et 0 (nul).
	</div>
	<br><br>&nbsp;
		
</div>






	
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
	
	
