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
	
	
	print = request("print")
	periodeSel = 12
	
	dim timerTotalTotal, maanedY, medarbTotalTimer
	Redim timerTotalTotal(0), maanedY(0), medarbTotalTimer(0)
	
	
	periodeShow = "12 Måneder frem (måned/måned)"
	Redim preserve timerTotalTotal(12)
	Redim preserve medarbTotalTimer(12)
	Redim preserve maanedY(12)
	
	
	
	
	
	
	
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
							
							
												
												monththis = datepart("m", datoer(y))
												yearthis = datepart("yyyy", datoer(y))
												
												
													'*** Opretter nye records ***
												
													if len(trim(tTimertildelt(y))) <> 0 then
													
													
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
																
															
															else
																		
																'**** Sletter evt. gamle registreringer (Hvis 0 er angivet) ***	
																'*** Renser ud i tabellen 
																strSQLdel = "DELETE FROM ressourcer_md WHERE jobid = "& tjobid(y) &" AND medid = "& tmedid(y) &""_
																&" AND md = "& monththis &" AND aar =  "& yearthis 
															
																oConn.execute strSQLdel
														
															end if
													end if
												
							
												
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
	
	
	
	jobSlutTEMP = dateadd("m", 11, "1/"& monththis &"/"& yearthis) '** år ?
	
	monthUse = datepart("m", jobSlutTEMP)
	yearUse = datepart("yyyy", jobSlutTEMP)
	
	call dageimd(monthUse,yearUse)	
	mthDaysUse = mthDays
	
	jobStartKri = yearthis&"/"&monththis&"/"&"1"
	jobSlutKri = yearUse&"/"&monthUse&"/"&mthDaysUse
	
	
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
	
	
	startdatoMD = month(jobStartKri)
	startdatoYY = year(jobStartKri)
	slutdatoMD =  month(jobSlutKri)
	slutdatoYY = year(jobSlutKri)
	
	
	
	if print <> "j" then%>
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
	
	else
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%
	
	topdiv = 20
	leftdiv = 20
	
	end if
	
	%>
	
	<!-------------------------------Sideindhold------------------------------------->
	
	<div id="overblik" style="position:absolute; left:<%=leftdiv%>; top:<%=topdiv%>; visibility:visible;">
	<h3><img src="../ill/ac0020-24.gif" width="24" height="24" alt="" border="0">&nbsp;Ressourcetimer (Forecast)</h3>
	
	<%
	
	if print <> "j" then
	toppxThisDiv = 170
	leftpxThisDiv = 460
	else
	toppxThisDiv = 60
	leftpxThisDiv = 110
	end if
	
	lav1 = 0
	lav2 = 40
	mellem1 = 40
	mellem2 = 80
	hoj1 = 80 
	
	
	
	%>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv-90%>px; background-color:LightGreen; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv-75%>px; padding:2px;">&nbsp; <%=lav1%> - <%=lav2%> time tildelt</div>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv+30%>px; background-color:#cccccc; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv+50%>px; padding:2px;">&nbsp;<%=mellem1%> - <%=mellem2%> tildelt</div>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv+125%>px; background-color:PeachPuff; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv+145%>px; padding:2px; width:300px;">&nbsp;> <%=hoj1%> timer tildelt.</div>		

	
	
	
	<%if print <> "j" then%>
	<div style="position:absolute; top:10px; left:750px;">
	<a href="javascript:popUp('ressource_belaeg_jbpla.asp?medarbSel=<%=medarbSel%>&jobselect=<%=jobnavn%>&print=j&mdselect=<%=monththis%>&aarselect=<%=yearThis%>','950','600','150','60');" target="_self" class=vmenu>Print venlig version&nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a>
	</div>
	
	
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
	<option value="0">Alle (loadtid optil 1 min.)</option>
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
	<!--<option value="1">1 Måned (dag/dag)</option>-->
	<!--<option value="3">3 Måneder frem (uge/uge)</option>
	<option value="4">4 Måneder frem (måned/måned)</option>-->
	<option value="12">12 Måneder frem (måned/måned)</option>
	<!--<option value="6">6 Måneder frem</option>-->
	</select></td>
	
	<td align=right style="padding-left:10px; padding-top:4px;"><input type="image" src="../ill/pilstorxp.gif"></td>
	</tr>
	
	<input type="hidden" name="selectminus" id="selectminus" value="0">
	<input type="hidden" name="selectplus" id="selectplus" value="0">
	
	<%if showonejob <> 1 then
	knapperTop = topdiv + 30
	else
	knapperTop = topdiv + 10
	end if
	
	if cint(medarbSel) = 0 Then
	leftKnap = 508
	else
	leftKnap = 708
	end if
	%>
	
	<div style="position:absolute; left:<%=leftKnap%>; top:<%=knapperTop-110%>; z-index:60000; width:150px;">
	<input type="submit" name="submitplus" id="submitplus" value="<- <%=monthname(month(dateadd("m", -1, "1/"&monththis&"/"&yearthis)))%>" style="font-size:9px; border:1px #003399 solid; background-color:#ffffff; width:75px;" onClick="selminus()">
	</div>
	<div style="position:absolute; left:<%=leftKnap+100%>; top:<%=knapperTop-110%>; z-index:60000; width:150px;">
	<input type="submit" name="submitplus" id="submitplus" value="<%=monthname(month(dateadd("m", +1, "1/"&monththis&"/"&yearthis)))%> ->" style="font-size:9px; border:1px #003399 solid; background-color:#ffffff; width:75px;" onClick="seladd()">
	</div>
	</form>
	</table>
	</div>
	<%end if 'Print%>
	
	
	</div>
	
	
	<%
	if print <> "j" then
	dtop = topdiv + 60
	else
	dtop = topdiv
	end if
	%>
	<div id="resbelaeg" style="position:absolute; left:20; top:<%=dtop+60%>; visibility:visible; z-index:10000;">
	<br><br><br><br><br><br>
	
	
	<%
	skrivCsvFil = 1 
	public csvTxtTop, csvTxt
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
	strSQL = "SELECT jobnavn, mnavn, mnr, j.id AS jobid, j.jobnr, medarbejdertype, m.mid, "
	
	'if periodeSel = 1 then
	'strSQL = strSQL & "r.timer AS restimer, r.dato AS resdato, r.id AS resid,"
	'else
	strSQL = strSQL & "rmd.timer AS restimer, rmd.md AS resmd, rmd.aar AS resaar, rmd.id AS resid,"
	'end if
	
	strSQL = strSQL & " k.kkundenavn, k.kkundenr, m.init, m.forecaststamp FROM medarbejdere m "
	
	'if periodeSel = 1 then
	'strSQL = strSQL & " LEFT JOIN ressourcer r ON (r.mid = m.mid AND r.dato BETWEEN '"& startdato &"' AND '"& slutdato &"')"
	'strSQL = strSQL & " LEFT JOIN job j ON  (j.id = r.jobid "& jobNavnKri &")"
	'else
	strSQL = strSQL & " LEFT JOIN ressourcer_md rmd ON (rmd.medid = m.mid AND "_
	&" ((rmd.md >= "& startdatoMD &" AND rmd.aar = "& startdatoYY &") OR (rmd.md <= "& slutdatoMD &" AND rmd.aar = "& slutdatoYY &")))"
	strSQL = strSQL &" LEFT JOIN job j ON (j.id = rmd.jobid "& jobNavnKri &")"
	'end if 
	
	strSQL = strSQL &" LEFT JOIN kunder k ON (kid = j.jobknr)"
	
	'if periodeSel = 1 then
	'strSQL = strSQL &" WHERE "& strMedidSQL &" AND mansat <> 2 ORDER BY m.mid, j.id, r.dato"
	'else
	strSQL = strSQL &" WHERE "& strMedidSQL &" AND m.mansat <> 2 ORDER BY m.mid, j.id, rmd.aar, rmd.md"
	'end if
	
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
	
	
	showresdato(x) = "1/"&oRec("resmd")&"/"&oRec("resaar")
	'showresdato(x) = oRec("resdato")
	
	showrestimer(x) = oRec("restimer")
	knavn(x) = oRec("kkundenavn")
	knr(x) = oRec("kkundenr")
	mnr(x) = oRec("mnr")
	minit(x) = oRec("init")
	forecaststamp(x) = oRec("forecaststamp")
	
	
	
	
			if len(oRec("resid")) <> 0 AND len(oRec("jobid")) <> 0 then
		
				if cdbl(lastMid) = oRec("mid") AND cdbl(lastJid) = oRec("jobid") AND cint(lastMonth) = datepart("m", showresdato(x)) AND datepart("yyyy", showresdato(x)) = lastYear then
				
					midjiddato(n, 0) = midjiddato(n, 0) + showrestimer(x)
					 
				else
					
					n = n + 1
					midjiddato(n, 0) = showrestimer(x) 
					midjiddato(n, 1) = useJobid(x) 
					midjiddato(n, 2) = xmid(x) 
					midjiddato(n, 3) = showresdato(x) 
					
				end if
				
			lastYear = datepart("yyyy", showresdato(x)) 
			lastMonth = datepart("m", showresdato(x), 2 ,3) 
			lastMid = oRec("mid")
			lastJid = oRec("jobid")
			
			end if
	
		
	
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
<input type="hidden" name="inklweekend_2" id="inklweekend_2" value="0">
<input type="hidden" name="func" id="func" value="redalle" checked><!-- <b>Opdatér felter enkeltvis.</b>--> 
<table border=0 cellspacing=1 cellpadding=0 bgcolor="#8caae6">

<%

MedIdSQLKri = ""
MedTotSQLKri = ""	
'MedJobFcastSQLKri = ""
numoffdaysorweeksinperiode = datediff("m", startdato, slutdato)
y = 0 
nday = startdato
weekrest = 0		
antalxx = x
xx = 0
foundone = 0
daynowrpl = replace(daynow, "/", "-")
lastMonth = monththis



call datoeroverskrift

'public MedTotSQLKri
strjIdWrt = "0#," '** Bruges kun når der er valgt en enkelt medarbejder


'**** Medarbejdere ****
for x = 0 to x - 1

	
		
	
	if lastxmid <> xmid(x) then	 
	

	tdcolspan = 5
	strDageChkboxOne = "<input type=text size=10 name=FM_sel_dage id=FM_sel_dage value="&daynowrpl&">&nbsp;dd-mm-aaaa"
	
	if x > 0 then				
		call medarbtotal
	end if%>
	
	
	<tr>
		<td bgcolor="#ffffe1" height=20 style="padding-left:5px; padding-top:5px;  padding-bottom:5px;">
		<font class=blaa><b><%=mnavn(x)%> (<%=mnr(x)%>)</b></font>&nbsp;
		<!--<a href="#" onClick="showtildeltimer('<=xmid(x)%>','<=strDageChkboxOne%>','<=id%>','<=periodeSel%>')" class=vmenuglobal>Ny</a>-->
		<%if len(trim(minit(x))) <> 0 then
		Response.write " - " & minit(x)
		end if
		%><br>
		<font class=megetlillesilver>Sidst opd: <%=forecaststamp(x)%></font>
		</td>
		
		<%if cint(medarbSel) <> 0 then%>
		<td colspan=4 bgcolor="#eff3ff">&nbsp;</td>
		<%end if%>
		
		<%
		cspan = 12
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
	
	MedIdSQLKri = xmid(x)
	csvTmnavn = mnavn(x)
	csvTmnr = mnr(x)
	csvTinit = minit(x)
	csvTnormuge = ntimPer
	end if
	
		
		
		
		
		if lastJid <> useJobid(x) OR lastxmid <> xmid(x) AND len(useJobid(x)) <> 0 then%>
		<tr>
		<td bgcolor="#ffffff" style="padding-left:5px; padding-right:5px;" width=200><%=useJobnavn(x)%> (<%=useJobnr(x)%>)<br>
		<font class=megetlillesilver><%=knavn(x)%> (<%=knr(x)%>)</font></td>
		
		

		<%if cint(medarbSel) <> 0 then%>	
				<td valign=top bgcolor="#eff3ff" align=right class=lille style="padding-top:1px; padding-right:5px; border-left:2px #8caae6 solid;">
				Forecast:<br>
				Realiseret:<br>
				Afvigelse:
				
				</td>
		<%else
		
		MedTotSQLKri = MedTotSQLKri & " OR tjobnr = "&useJobnr(x)
		end if
						
						
						
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
								
								
								
								
												
												
												
												if cint(medarbSel) <> 0 then
															
															'** Timer forecastet i periode ***
															call tildelttimerPer(xmid(x), useJobid(x), statDatoUse, s)
															%>	
												
															<td valign=top bgcolor="#eff3ff" align=right class=lille style="padding-top:1px; padding-right:5px;  border-right:<%=bdright%> #8caae6 solid; border-left:<%=bdleft%> #8caae6 solid;">
															<%=formatnumber(tildtimPer, 1)%><br>
															<%
															
															
															sumTimer = 0
															
															strSQLtimer = "SELECT sum(timer) AS sumtimerFaktisk FROM timer WHERE tmnr = "& xmid(x) &" AND tjobnr = "&useJobnr(x)&" AND tfaktim <> 5 AND (YEAR(tdato) = "& year(statDatoUse) &" AND MONTH(tdato) = "& month(statDatoUse) &")"
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
															
															
															else
																	if tildtimPer = 0 then
																	afv = 0
																	else
																	afv = tildtimPer '100
																	end if
															end if '(tildtimPer - (sumTimer))
															
															Response.write formatnumber(afv, 0) &"%"
												
												
												
												end if
												
							
							if cint(medarbSel) <> 0 then
							csvTxt = csvTxt & vbcrlf & mnavn(x) &" ("& mnr(x) &")"
							csvTxt = csvTxt & ";"& minit(x) 
							csvTxt = csvTxt & ";"& formatnumber(ntimPer, 2) 
							csvTxt = csvTxt & ";"& useJobnavn(x) & ";" & useJobnr(x) 
							csvTxt = csvTxt & ";"& knavn(x) & ";" & knr(x)
							
							csvTxt = csvTxt &";"& csvDatoer(thisy) 
							csvTxt = csvTxt & ";" & formatnumber(tildtimPer, 2)
							csvTxt = csvTxt & ";" & formatnumber(sumTimer, 2)
							csvTxt = csvTxt & ";" & formatnumber(afv, 0)
							end if
							
							
							
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
						<%next
		
		
		
		
		
		'***** Udkskriver forecast timer pr. md. **********************
		strjIdWrt = strjIdWrt & "#"&useJobid(x)&"#," 
		lastJid = useJobid(x)
	
					
					
					lastMonth = monththis
					y = 0
					for y = 0 TO numoffdaysorweeksinperiode
					
						call antaliperiode(periodeSel, y, startdato, monthUse)
					
								
									timerThis = 0	
									For n = 0 to UBOUND(midjiddato, 1)
										if (datepart("m", midjiddato(n, 3)) = datepart("m", nday)) AND midjiddato(n, 1) = useJobid(x) AND midjiddato(n, 2) = xmid(x) AND datepart("yyyy", midjiddato(n, 3)) = datepart("yyyy", nday) then
										timerThis = midjiddato(n, 0)
										foundone = 1
										end if
									next
									
								
								
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
					stThisDato = dato
					txtwidth = 35
					
					
					
					if timerThis = 0 then
					showTimerThis = ""
					else
					showTimerThis = timerThis
					end if
					%>
					<td bgcolor="<%=bgTimer%>" valign=top align=center width="<%=txtwidth+4%>" style="padding-right:3px;">
					
					<%if print <> "j" then%>
					<input type="hidden" name="FM_jobid" id="FM_jobid" value="<%=useJobid(x)%>">
					<input type="hidden" name="FM_medarbid" id="FM_medarbid" value="<%=xmid(x)%>">
					<input type="hidden" name="FM_dato" id="FM_dato" value="<%=stThisDato%>">
					<%=showTimerThis%><br>
					<input type="text" name="FM_timer" id="FM_timer" style="width:<%=txtwidth%>px; font-family:arial; font-size:9px;" value="">
					<%
					'*** hidden array felt. *** Bruges så der kan skelneshvis man bruger komma.%>
					<input type="hidden" name="FM_timer" id="FM_timer" style="width:<%=txtwidth%>px; font-size:9px;" value="#">
					<%else%>
					<%=showTimerThis%>
					<%end if 'Print%>
					
					<%
					
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
		end if '** Lastjobid
		
		lastxmid = xmid(x)
		%>
		</tr>
		<%	
		next
		
		
		
		'*** Hvis der er valgt medarbejder ***
		if medarbSel <> 0 then
				usemrn = medarbSel
				%>
				
				<%call medarbTotal%>
				</table>
				
				
				
				<%if print <> "j" then%>
				<br>
				<table cellspacing=0 cellborder=0 border=0>
				<tr><td><img src="../ill/blank.gif" width="750" height="1" alt="" border="0"><input type="submit" value="Indlæs forecast!"></td></tr>
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
								<td bgcolor="#ffffff" style="padding-left:5px; padding-right:5px;" width=200><%=oRec("jobnavn")%> (<%=oRec("jobnr")%>) <br>
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
											strSQLtimer = "SELECT sum(timer) AS sumtimerFaktisk FROM timer WHERE tmnr = "& usemrn &" AND tjobnr = "& oRec("id") &" AND tfaktim <> 5 AND (YEAR(tdato) = "& year(statDatoUse) &" AND MONTH(tdato) = "& month(statDatoUse) &")"
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
											txtwidth = 35
											
											
											for y = 0 TO numoffdaysorweeksinperiode
													
													call antaliperiode(periodeSel, y, startdato, monthUse)
													lastMonth = newMonth
											%>	
													
											<td bgcolor="#ffffff" width="<%=txtwidth+4%>" valign=top align=right style="padding-right:3px;">
											<input type="hidden" name="FM_jobid" id="FM_jobid" value="<%=oRec("id")%>">
											<input type="hidden" name="FM_medarbid" id="FM_medarbid" value="<%=medarbSel%>">
											<input type="hidden" name="FM_dato" id="FM_dato" value="<%=nday%>">
											&nbsp;<br>
											
											<input type="text" name="FM_timer" id="FM_timer" style="width:<%=txtwidth%>px; font-family:arial; font-size:9px;" value="">
											<input type="hidden" name="FM_timer" id="FM_timer" value="#">
											</td>
											<%next%>
								
									</tr>
								
								<%
								j = j + 1
								end if 'instr
								
								
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
				
				end if'Print
				
				else
				%>
				<tr>
					<td height=20 bgcolor="#ffffff" align=right style="padding-right:15px;"><b>Total:</b> (Alle job / medarbejdere)</td>
					
					<% if cint(medarbSel) <> 0 then%>
					<td colspan=4 bgcolor="#ffffff">&nbsp;</td>
					<%end if
					
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
		
		
		<%if print <> "j" then%>
		<br>
		<table cellspacing=0 cellborder=0 border=0>
		<tr><td>
		<%if cint(medarbSel) <> 0 then%>	
		<img src="../ill/blank.gif" width="750" height="1" alt="" border="0">
		<%else%>
		<img src="../ill/blank.gif" width="550" height="1" alt="" border="0">
		<%end if%><input type="submit" value="Indlæs forecast!"></td></tr>
		</form>
		</table>
		
		
		<form action="ressourcer_eksport.asp" method="post" name=theForm onsubmit="BreakItUp()"> <!--  -->
			<input type="hidden" size=200 name="txt1" id="txt1" value="<%=csvTxtTop%>">
			<input type="hidden" name="BigTextArea" id="BigTextArea" value="<%=csvTxt%>">
			<!--<br><textarea cols="60" rows="10" name="sdf"><=csvTxt%></textarea>-->
			<input type="hidden" name="txt20" id="txt20" value="<%=csvTxtTotal%>">
			
			<input type="submit" value="Eksporter (Pivot opti.)"><br>
			(kun job med tildelte timer i den valgte periode bliver eksporteret.)
			
			</form>
			
			
	<div style="position:relative; background-color:#eff3ff; visibility:visible; border:1px red dashed; border-right:2px red dashed; border-bottom:2px red dashed; padding:15px; width:450px;"><img src="../ill/ac0005-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Side note(r):<br></b>
	- Tildel timer ved at vælge den ønskede medarbejder.<br>
	- 1 måneds oversigt:<u> Timer angives antal timer pr. dag.</u><br>
	- 12 måneders oversigt:<u> Timer angives som samlet antal timer pr. måned.</u><br>
	- Slet forecast ved at indtaste et 0 (nul).
	</div>
	<%end if 'Print%>
	
	
	<br><br>&nbsp;
		
</div>






	
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
	
	
