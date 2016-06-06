<%response.buffer = true%>


<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/webblik_func.asp"-->
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
	
	'*** Akt. typer der tælles med i ugereg. ***'
	call akttyper2009(2)
	
	
	func = request("func")
	
	thisfile = "res_belaeg"
	'Response.write "func: " & func
	
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	media = request("media")
	'print = request("print")
	
	'*** Periode ***
	dim timerTotalTotal, maanedY, medarbTotalTimer, medarbTimerReal, realTotalTotal
	
	if len(request("periodeselect")) <> 0 then
	periodeSel = request("periodeselect")
	'periodeShow = ""& periodeSel &" Måneder frem"
	perBeregn = periodeSel - 1
	else
	
	        if len(request.cookies("resbel")("periodesel")) <> 0 then
	        periodeSel = request.cookies("resbel")("periodesel")
	        'periodeShow = ""& periodeSel &" Måneder frem"
	        perBeregn = periodeSel - 1
	        else
	        periodeSel = 6
	        'periodeShow = "6 Måneder frem"
	        perBeregn = 5
	        end if
	end if
	
	
	if len(request("periodeselect")) <> 0 then
	    if len(trim(request("FM_visdeak"))) <> 0 then
	    visDeak = 1
	    visDeakChk = "CHECKED"
	    else
	    visDeakChk = ""
	    visDeak = 0
	    end if
	else
	    if request.cookies("tsa")("fc_visdeak") = "1" then
	    visDeak = 1
	    visDeakChk = "CHECKED"
	    else
	    visDeakChk = ""
	    visDeak = 0
	    end if
	end if
	
	Response.Cookies("tsa")("fc_visdeak") = visDeak
	
	if cint(visDeak) = 0 then
	deakSQL = " AND m.mansat <> 2 AND m.mansat <> 3 "
	else
	deakSQL = " "
	end if
	
	select case periodeSel
	'case 1
	'leftwdt = 35
	'rdimnb = 31
	'periodeShow = periodeShow & " dage"
	case 3
	stDatoDag = 1 
	rdimnb = 14
	leftwdt = 350
	periodeShow = periodeShow & " 3 måneder frem (uge/uge)"
	case 6
	stDatoDag = 14 '** for at være sikker på at der ikke bliver skrvet en uge ind, der hører til forrige måned.
	rdimnb = 6
	leftwdt = 500
	periodeShow = "6 Måneder frem"
	case 12
	stDatoDag = 14
	rdimnb = 12
	leftwdt = 850
	periodeShow = "12 Måneder frem"
	end select
	
	Redim timerTotalTotal(rdimnb), maanedY(rdimnb), medarbTotalTimer(rdimnb), medarbTimerReal(rdimnb), realTotalTotal(rdimnb) 
	
	
	response.cookies("resbel")("periodesel") = periodeSel
	
	
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
	
	
	
	jobSlutTEMP = dateadd("m", perBeregn, "1/"& monththis &"/"& yearthis) '** år ?
	
	monthUse = datepart("m", jobSlutTEMP)
	yearUse = datepart("yyyy", jobSlutTEMP)
	
	call dageimd(monthUse,yearUse)	
	mthDaysUse = mthDays
	
	jobStartKri = yearthis&"/"&monththis&"/"&stDatoDag
	jobSlutKri = yearUse&"/"&monthUse&"/"&mthDaysUse
	
	datoInterval = datediff("d",jobStartKri,jobSlutKri,2,2)
	
	daynow = day(now)&"/"&month(now)&"/"&year(now)
	
	
	if len(trim(request("FM_jobsel"))) <> 0 then 		
	jobidsel = request("FM_jobsel")
	response.cookies("resbel")("jobidsel") = jobidsel
	else
	
	    if len(request.cookies("resbel")("jobidsel")) <> 0 then
		jobidsel = request.cookies("resbel")("jobidsel")
		else
		jobidsel = 0
		end if
	
	end if
	
	if cint(jobidsel) <> 0 then
	jobSQLkri = " j.id = " & jobidsel
	else
	jobSQLkri = " j.id <> 0 " 
	end if
		
	if len(trim(request("viskunjobiper"))) <> 0 then
	viskunjobiper = request("viskunjobiper")
	response.cookies("resbel")("viskunjobiper") = viskunjobiper
	else
	    
	    if len(request.cookies("resbel")("viskunjobiper")) <> 0 then
		viskunjobiper = request.cookies("resbel")("viskunjobiper")
		else
		viskunjobiper = 0
		end if
	
	end if 
	
	
	if viskunjobiper <> 0 then
	viskunjobiper1 = "CHECKED"
	viskunjobiper0 = ""
	jobDatoKri = " AND jobslutdato >= '"& jobStartKri &"' "
	else
	jobDatoKri = ""
	viskunjobiper1 = ""
	viskunjobiper0 = "CHECKED"
	end if
	
	'if len(request("showonejob")) <> 0 then
	'showonejob = request("showonejob")
	'else
	showonejob = 0
	'end if
	
	
	
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
	
	if (level <= 2 OR level = 6) then 
	whSQLkri = " m.mid <> 0 "
	multiUse = "multiple"
	
	call licensStartDato()
	
	if licensklienter < 20 then
	    if licensklienter < 5 then
	    szi = 5
	    else
	    szi = 10
	    end if
	else
	    if licensklienter > 36 then
	    szi = 20
	    else
	    szi = 15
	    end if
	end if
	
	else
	szi = 1
	multiUse = ""
	whSQLkri = " m.mid = " & medarbSel
	end if
	
	
	'Response.Write "her" & medarbSel & "<br>"
	
	if medarbSel = "0" Then
	strMedidSQL = " m.mid <> 0 "
	strMedidSel = 0
	else
	    strMedidSQL = " (m.mid = 0"
	    medarbSelArr = split(trim(medarbSel), ",")
	    for m = 0 to UBOUND(medarbSelArr)
	    strMedidSQL = strMedidSQL & " OR m.mid = " & trim(medarbSelArr(m))
	    strMedidSel = strMedidSel & "#"& trim(medarbSelArr(m)) &"#"
	    next
	    strMedidSQL = strMedidSQL & ")"
	end if
	
	
	
	response.cookies("resbel").expires = date + 30
	
	'Response.write strMedidSQL & "<br>"& strMedidSel
	'Response.flush
	
	
	
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
							
							
												
												monththis = datepart("m", datoer(y), 2,2)
												'monththis = datepart("m", datoer(y))
												yearthis = datepart("yyyy", datoer(y))
												weekThis = datepart("ww", datoer(y), 2,2)
												
												
													'*** Opretter nye records ***
												
													if len(trim(tTimertildelt(y))) <> 0 then
													
													
															if tTimertildelt(y) > 0 then
																
																'*** Renser ud i tabellen 
																select case periodeSel
																case 3
																strSQLperKri = " AND uge = "& weekThis
																case 6,12
														        strSQLperKri = " AND md = "& monththis
																end select
																
																strSQLdel = "DELETE FROM ressourcer_md WHERE jobid = "& tjobid(y) &" AND medid = "& tmedid(y) &""_
																&"  AND aar =  "& yearthis & strSQLperKri
															    
															    'AND md = "& monththis &"
															    'Response.Write "<br>periodeSel"& periodeSel &" strSQLdel:" & strSQLdel &"<br>"
																oConn.execute strSQLdel
															
																
																timerThisMD =  sqlbless(tTimertildelt(y))
																
																strSQL = "INSERT INTO ressourcer_md"_
																& " (jobid, medid, md, aar, timer, uge) VALUES ("_
																&" "& tjobid(y) &", "& tmedid(y) &", "& monththis &", "& yearthis &", "& timerThisMD &", "& weekThis &")" 
																
																
															    'Response.Write "<br>strSQL A"& strSQL & "<br>"
																'Response.flush
																
																oConn.execute strSQL
																
															
															else
																		
																'**** Sletter evt. gamle registreringer (Hvis 0 er angivet) ***	
																'*** Renser ud i tabellen 
																select case periodeSel
																case 3
																strSQLperKri = " AND uge = "& weekThis
																case 6,12
														        strSQLperKri = " AND md = "& monththis
																end select
																
																
																strSQLdel = "DELETE FROM ressourcer_md WHERE jobid = "& tjobid(y) &" AND medid = "& tmedid(y) &""_
																&" AND aar =  "& yearthis & strSQLperKri
															    'AND md = "& monththis &"
															    'Response.Write "<br>strSQldel B"& strSQldel & "<br>"
															    'Response.flush
															    
																oConn.execute strSQLdel
														
															end if
													end if
												
							
												
						next
	
	
	
	


	
	'Response.end
	response.Redirect "ressource_belaeg_jbpla.asp?medarbSel="&medarbSel&"&showonejob="&showonejob&"&id="&id&"&periodeselect="&periodeSel&"&jobselect="&jobnavn&"&mdselect="&request("mdselect")&"&aarselect="&request("aarselect")
	
	
	end if
	
	
	if func = "opduge" then
	
	    strSQL = "SELECT md, aar, id FROM ressourcer_md WHERE id <> 0"
	    oRec.open strSQL, oConn, 3
	    While not oRec.EOF
	            
	            thisdate = "1/"& oRec("md")&"/"&oRec("aar")
	            thisWeek = datepart("ww", thisdate, 2,2)
	            
	            
	            strSQLu = "UPDATE ressourcer_md SET uge = "& thisWeek & " WHERE id = "& oRec("id")
	            'Response.write "<b>"& thisdate &"</b>: "& strSQLu & "<br>"
	            'Response.flush
	            'oConn.execute strSQLu
	            
	    
	    oRec.movenext
	    wend 
	    oRec.close
	
	
	'Response.Write "ok"
	Response.end
	
	end if
	
	
	
	
	
	
	
	
	
	
	
	
		
	
	intMid = 0
	intJid = 0
	
	startdato = jobStartKri
	slutdato =  jobSlutKri
	
	
	startdatoMD = month(jobStartKri)
	startdatoYY = year(jobStartKri)
	slutdatoMD =  month(jobSlutKri)
	slutdatoYY = year(jobSlutKri)
	
	'** Forgående 3 md. (totaler) ***
	startdatoMDtot = month(dateadd("m", -3, jobStartKri))
	startdatoYYtot = year(dateadd("m", -3, jobStartKri))
	slutdatoMDtot =  month(dateadd("m", -1, jobStartKri))
	slutdatoYYtot = year(dateadd("m", -1, jobStartKri))
	
	timregStdato = year(dateadd("m", -3, jobStartKri))&"/"& month(dateadd("m", -3, jobStartKri)) &"/"& day(dateadd("m", -3, jobStartKri))
	timregSldato = year(dateadd("m", -0, jobStartKri))&"/"& month(dateadd("m", -0, jobStartKri)) &"/"& day(dateadd("m", -0, jobStartKri))
	
	
	if media <> "print" AND media <> "eksport" then%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%if showonejob <> 1 then%>
	
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	<%call tsamainmenu(4)%>
	</div>
	 <div id="Div1" style="position:absolute; left:15; top:82; visibility:visible;">
	        <%
	        call webbliktopmenu()
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
	
	
	'*************************************************************************************	
	
	%>
	
	
		
	
     
    <script src="inc/res_bl_jav.js"></script>
	
    <!-------------------------------Sideindhold------------------------------------->
	
	<%
	
	
	
	
	
    
	if media <> "eksport" then 
	ldTop = 400
	ldLft = 350
	else
	ldTop = 5
	ldLft = 10
	end if
	%>
	<div id="load" style="position:absolute; display:; visibility:visible; top:<%=ldTop%>px; left:<%=ldLft%>px; width:300px; background-color:#ffffff; border:1px #cccccc solid; padding:2px; z-index:100000000;">
    <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" /><br />
	&nbsp;&nbsp;&nbsp;&nbsp;Vent veligst. Der kan gå op til 1 min...
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	</td></tr></table>
	
	</div>
	<%
	
	Response.Flush 
	
	'end if %>
	
	<%'Response.Flush %>
	<%'Response.End %>
	
	
	
	
	<%
	
	if media <> "eksport" then
	
	%>
	<div id="overblik" style="position:absolute; left:<%=leftdiv%>; top:<%=topdiv%>; visibility:visible;">
	
	<%
	
	oimg = "ikon_ressourcer_48.png"
	oleft = 0
	otop = 0
	owdt = 800
	oskrift = "Ressource allokering (Forecast)"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	
	end if
	
	
	if media <> "print" then
	toppxThisDiv = 260
	leftpxThisDiv = 1110
	else
	toppxThisDiv = 20
	leftpxThisDiv = 910
	end if
	
	
	if periodeSel <> 3 then
	lav1 = 0
	lav2 = 40
	mellem1 = 40
	mellem2 = 80
	hoj1 = 80 
	else
	lav1 = 0
	lav2 = 15
	mellem1 = 15
	mellem2 = 30
	hoj1 = 30 
	end if
	
	
	if media <> "eksport" then
	
	
	%>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv-90%>px; background-color:LightGreen; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
	<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv-65%>px; padding:2px;">&nbsp; <%=lav1%> - <%=lav2%> time tildelt</div>
	<div style="position:absolute; top:<%=toppxThisDiv+30%>px; left:<%=leftpxThisDiv-90%>px; background-color:#cccccc; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
	<div style="position:absolute; top:<%=toppxThisDiv+30%>px; left:<%=leftpxThisDiv-65%>px; padding:2px;">&nbsp;<%=mellem1%> - <%=mellem2%> tildelt</div>
	<div style="position:absolute; top:<%=toppxThisDiv+60%>px; left:<%=leftpxThisDiv-90%>px; background-color:#FF0066; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
	<div style="position:absolute; top:<%=toppxThisDiv+60%>px; left:<%=leftpxThisDiv-65%>px; padding:2px; width:300px;">&nbsp;> <%=hoj1%> timer tildelt.</div>		

	
	
	
	<%
	end if
	
	
	
	
	if media <> "print" ANd media <> "eksport" then%>
	
	
	<%call filterheader(0,0,1004,pTxt) %>
	
	<table cellspacing="0" cellpadding="2" border="0" width=100%>
	<form action="ressource_belaeg_jbpla.asp?showonejob=<%=showonejob%>&id=<%=id%>" method="post" name="mdselect" id="mdselect">
	<tr>
	
	<td valign=top><b>Medarbejder:</b><br />
	
        <input id="FM_visdeak" name="FM_visdeak" value="1" type="checkbox" <%=visDeakChk %> /> Vis de-aktive og<br />
        passive medarbejdere</td><td>
	<%
	
	
	
	
	
	
	strSQL = "SELECT mnavn, mid, mnr FROM medarbejdere m WHERE "& whSQLKri &" "& deakSQL &" ORDER BY mnavn"
	'Response.write strSQL
	'Response.flush
	oRec.open strSQL, oConn, 3 
	%><select name="medarbSel" id="medarbSel" style="width:426px; font-size:9px;" <%=multiUse %> size="<%=szi %>">
	
	<%if level <= 2 OR level = 6 then 
	
	if medarbSel = "0" then
	nulMSel = "SELECTED"
	else
	nulMSel = ""
	end if
	
	%>
    <option value="0" <%=nulMSel %>>Alle (loadtid optil 20 sek.)</option>
	<%
	end if
	
	while not oRec.EOF 
	
	 if instr(strMedidSel, "#"&oRec("mid")&"#") <> 0 then
	 mSEL = "SELECTED"
	 else
	 mSEL = ""
	 end if %>
	 
	<option value="<%=oRec("mid")%>" <%=mSEL%>><%=oRec("mnavn")%>&nbsp;(<%=oRec("mnr")%>)</option>
	<%
	oRec.movenext
	wend
	oRec.close %>
	</select></td>
	
	<td valign=top><b>Periode fra:</b><br /> 
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
	<option value="2013">2013</option>
	<option value="2014">2014</option>
	<option value="2015">2015</option>
	<option value="2016">2016</option>
	</select> og 
	<select name="periodeselect" style="font-size:9px; width:190px;" onchange="submit()">
	<option value="<%=periodeSel%>" SELECTED><%=periodeShow%></option>
	<!--<option value="1">1 Måned (dag/dag)</option>-->
	<option value="3">3 Måneder frem (uge/uge)</option>
	<!--<option value="4">4 Måneder frem (måned/måned)</option>-->
	<option value="12">12 Måneder frem</option>
	<option value="6">6 Måneder frem</option>
    </select>
    
    <br />
    <input type="hidden" name="selectminus" id="selectminus" value="0">
	<input type="hidden" name="selectplus" id="selectplus" value="0">
	

	
	<%
	leftKnap = 640
	%>
	
	<div style="position:absolute; left:<%=leftKnap%>; top:90; z-index:60000; width:150px;">
	<input type="submit" name="submitplus" id="submitplus" value="<- <%=monthname(month(dateadd("m", -1, "1/"&monththis&"/"&yearthis)))%>" style="font-size:9px; border:1px #003399 solid; background-color:#ffffff; width:75px;" onClick="selminus()">
	</div>
	<div style="position:absolute; left:<%=leftKnap+80%>; top:90; z-index:60000; width:150px;">
	<input type="submit" name="submitplus" id="submitplus" value="<%=monthname(month(dateadd("m", +1, "1/"&monththis&"/"&yearthis)))%> ->" style="font-size:9px; border:1px #003399 solid; background-color:#ffffff; width:75px;" onClick="seladd()">
	</div>
    
    </td>
	
	</tr>
	<tr>
	
	<%if showonejob <> 1 then%>
	<td valign=top><b>Job:</b><br />
	Vis forecast på valgte job.</td>
	<td>
	<%
		strSQL = "SELECT jobnavn, jobnr, j.id, k.kkundenavn, k.kkundenr"_
		&" FROM job j "_
		&" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_
	    &" WHERE jobstatus = 1 "& jobDatoKri &" AND j.fakturerbart = 1 ORDER BY k.kkundenavn, j.jobnavn"
		
		'3 = tilbud
		'Response.write strSQL
		'Response.flush
		%>
        
		
		<select name="FM_jobsel" id="FM_jobsel" style="font-size : 12px; width:426px;" onChange="submit();">
		<option value="0">Alle</option>
		<%
				
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(jobidsel) = cint(oRec("id")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				
				
				%>
				<option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("kkundenavn")%>&nbsp;(<%=oRec("kkundenr")%>) ][ <%=oRec("jobnavn")%>&nbsp;(<%=oRec("jobnr")%>) <%=aftnavnognr %></option>
				<%
				oRec.movenext
				wend
				oRec.close
				
				%>
				
		</select>
		<br />
		
		       <input id="viskunjobiper0" name="viskunjobiper" value="0" type="radio" <%=viskunjobiper0 %> /> Vis alle aktive job.<br />
 
		       <input id="viskunjobiper1" name="viskunjobiper" value="1" type="radio" <%=viskunjobiper1 %> /> Vis kun job, hvis slutdato ligger efter den valgte periode start.
 
	
	</td>
	<%else%>
	<td>&nbsp;</td>
    <input type="hidden" name="FM_jobid" id="FM_jobid" value="<%=jobid%>"><b><%=jobid%></b>
	<%end if%>
	
	
	<td align=right style="padding-right:40px; padding-top:4px;">
        <input id="Submit1" type="submit" value=" Søg >> " /></td>
	</tr>
	
	
	
	</form>
	</table>
	</div>
	<!-- filter header --->
	</td>
	</tr>
	</table>
	</div>
	<%end if 'Print%>
	
	
	
	
	<%
	
	if media <> "eksport" then
	
	tTop = 50
	tLeft = 0
	tId = "sidediv"
	tVzb = "visible"
	tDsp = ""
	
	if media <> "print" then
	dtop = topdiv + 60
	tWdth = 1004
	call tableDivWid(tTop,tLeft,tWdth,tId,tVzb,tDsp)
	else
	tWdth = 1004
	dtop = topdiv
	call tableDiv(tTop,tLeft,tWdth)
	end if
	
	
	end if
	
	
	
	
	
	
	
	
	
	
	skrivCsvFil = 1 
	public csvTxtTop, csvTxt
	public csvDatoerMD(35), csvDatoerAR(35), csvDatoerUge(35)
	
	if skrivCsvFil = 1 then
		
		csvTxtTop = "Medarbejder;Initialer;Sidst opd.;Normeret uge;Job;Job nr.;Jobansvarlig;Init;Jobejer;Init;Kunde;Kunde nr." 
		
		'if lto <> "execon" then
		csvTxtTop = csvTxtTop & ";Uge"
		'end if
		
		csvTxtTop = csvTxtTop &";Måned;År;Forecast;Realiseret;Afvigelse %"
	end if
	
	mids = 150
	dt = 200
	ji = 4
	
	dim midjiddato 
	redim midjiddato(mids, dt, ji) 
	
	antalX = 4000
	xId = 16
	
	'Response.flush
	
	dim medarbKundeoplysX
	redim medarbKundeoplysX(antalX, xId) 
	
	'dim testarr
	'***  Md, Aar, Jobid, Mid **
	'select case lto
	'case "bminds", "execon"
	'redim testarr(53,12,44,900,110)
	'case else
	'redim testarr(53,12,500,40)
	'end select
	
	'*** Unique job id brugs i array, så hvis de når job id nr 901 fejler siden. **'
	
	lastYear = 0
	lastWeek = 0
	lastMid = 0
	lastJid = 0
	n = 0
	
	
	lastmedarbId = 0
	
	
	if cint(startdatoYY) = cint(slutDatoYY)  then
	orandval = " AND "
	else
	orandval = " OR "
	end if
	
	jobisWrt = " AND (t.tjobnr <> 0 "
	
	'************************************************************************
	'*** Main SQL ***
	'************************************************************************
	't = 0 alle med ress. timer i forecast
	't = 1 alle med realiserede timer og som ikke er med i forecast
	
	x = 0
	
	tStDato = year(jobStartKri) &"-"& month(jobStartKri) & "-1" 
	tSlDato = jobSlutKri
	
	for t = 0 to 1
	
	strSQL = "SELECT jobnavn, m.mnavn, m.mnr, j.id AS jobid, j.jobnr, m.medarbejdertype, m.mid, "
    
    if t = 0 then
    strSQL = strSQL & "rmd.timer AS restimer, rmd.md AS resmd, rmd.aar AS resaar, rmd.id AS resid, rmd.uge AS resuge, "
	else
	strSQL = strSQL & "SUM(t.timer) AS realtimer, "
	end if
	
	strSQL = strSQL & " k.kkundenavn, k.kkundenr, m.init, m.forecaststamp, jobans1, jobans2, "_
	&" m2.mnavn AS m2mnavn, m2.init AS m2init, m2.mnr AS m2mnr, m3.mnavn AS m3mnavn, m3.init AS m3init, m3.mnr AS m3mnr "_
    &" FROM medarbejdere m "
	
	if t = 0 then
	strSQL = strSQL & " LEFT JOIN ressourcer_md rmd ON (rmd.medid = m.mid AND "_
	&" ((rmd.md >= "& startdatoMD &" AND rmd.aar = "& startdatoYY &") "& orandval &" (rmd.md <= "& slutdatoMD &" AND rmd.aar = "& slutdatoYY &")))"_
	&" LEFT JOIN job j ON (j.id = rmd.jobid AND "& jobSQLkri &")"
	else
	strSQL = strSQL &" LEFT JOIN timer t ON (t.tmnr = m.mid "& jobisWrt &")"_
	&" LEFT JOIN job j ON (j.jobnr = t.tjobnr AND t.timer <> 0 AND "& jobSQLkri &" AND j.jobnavn IS NOT NULL)"
	end if
	
	strSQL = strSQL &" LEFT JOIN kunder k ON (kid = j.jobknr)"_
	&" LEFT JOIN medarbejdere m2 ON (m2.mid = j.jobans1)"_
	&" LEFT JOIN medarbejdere m3 ON (m3.mid = j.jobans2)"
	
	strSQL = strSQL &" WHERE "& strMedidSQL &" "& deakSQL 
	
	
	if t = 1 then
	strSQL = strSQL &" AND t.tdato BETWEEN '"& tStDato &"' AND '"& tSlDato &"'"
	strSQL = strSQL &" GROUP BY m.mid, j.id"
	end if
	
	strSQL = strSQL &" ORDER BY m.mid, j.id"
	
	if t = 0 then
	strSQL = strSQL &", rmd.aar, rmd.md"
	end if
	
	'AND (m.mansat <> 2 AND m.mansat <> 3)
	'Response.write strSQL &"<br>"
	'Response.flush
	
	oRec.open strSQL, oConn, 3 
	
	
	while not oRec.EOF 
	
	if t = 0 then
	
	    'if x = 0 then
	    'jobisWrt = " AND (t.tjobnr <> "& oRec("jobnr") 
	    'else
	    if oRec("jobnr") <> "" then
	    jobisWrt = jobisWrt & " AND t.tjobnr <> "& oRec("jobnr") 
	    end if
	    
	    'end if 
	
	end if
    
    if lastMid <> oRec("mid") then
    n = 0
    end if
    
    if t = 0 then
	showresdato = "1/"&oRec("resmd")&"/"&oRec("resaar")
	showresuge = oRec("resuge")
	else
	showresdato = "1/1/2002" '"&oRec("resmd")&"/"&oRec("resaar")
	showresuge = 1
	end if
	
	medarbKundeoplysX(x, 0) = oRec("jobid")
	medarbKundeoplysX(x, 1) = oRec("jobnr")
    medarbKundeoplysX(x, 2) = oRec("jobnavn")
    medarbKundeoplysX(x, 3) = oRec("mid")
    medarbKundeoplysX(x, 4) = oRec("mnavn")
    
	medarbKundeoplysX(x, 5) = oRec("kkundenavn")
	medarbKundeoplysX(x, 6) = oRec("kkundenr")
	medarbKundeoplysX(x, 7) = oRec("mnr")
	medarbKundeoplysX(x, 8) = oRec("init")
	
	if t = 0 then
	medarbKundeoplysX(x, 9) = oRec("forecaststamp")
	medarbKundeoplysX(x, 10) = oRec("restimer")
	resid = oRec("resid")
	else
	medarbKundeoplysX(x, 9) = "2002-01-01" 'oRec("forecaststamp")
	medarbKundeoplysX(x, 10) = 0
	resid = -1
	end if
	
	
	medarbKundeoplysX(x, 11) = oRec("jobans1")
	medarbKundeoplysX(x, 12) = oRec("jobans2")
	
	if len(trim(oRec("m2mnavn"))) <> 0 then
    medarbKundeoplysX(x, 13) = oRec("m2mnavn") &" ("&oRec("m2mnr")&")"
    medarbKundeoplysX(x, 14) = oRec("m2init") 
    else
    medarbKundeoplysX(x, 13) = ""
    medarbKundeoplysX(x, 14) = ""
    end if
    
    if len(trim(oRec("m3mnavn"))) <> 0 then
    medarbKundeoplysX(x, 15) = oRec("m3mnavn") &" ("&oRec("m3mnr")&")"
    medarbKundeoplysX(x, 16) = oRec("m3init")
    else
    medarbKundeoplysX(x, 15) = ""
    medarbKundeoplysX(x, 16) = ""
    end if 
	
	
	        
	        'Response.Write len(oRec("resid")) &" <> 0 AND "& len(oRec("jobid")) 
	
			if len(resid) <> 0 AND len(oRec("jobid")) <> 0 then
		         
		         
		         if periodeSel = 3 then   
		            
				        if cdbl(lastMid) = oRec("mid") AND cdbl(lastJid) = oRec("jobid") AND _
				        datepart("yyyy", showresdato) = lastYear AND showresuge = lastWeek then
        				
        				
	                        midjiddato(oRec("mid"),n, 0) = midjiddato(oRec("mid"),n, 0) + medarbKundeoplysX(x, 10) 'showrestimer
        					 
				        else
        					
					        'Response.Write "<br><br><br><br><br><br><br><br><br><br>x:"& x &" arr:"& oRec("resmd") &","& right(oRec("resaar"),2) &","& oRec("jobid") &","& oRec("mid") & "<br>"
	                        'Response.Flush
        					
					        'testarr(oRec("resuge"),right(oRec("resaar"),2),oRec("jobid"),oRec("mid")) = oRec("restimer")
        	                
        	                
        	                
					        n = n + 1
					        'Response.Write "N: " & n & "<br>"
					        'Response.flush
					        midjiddato(oRec("mid"),n, 0) = medarbKundeoplysX(x, 10) 'oRec("restimer") 'showrestimer(x) 
					        midjiddato(oRec("mid"),n, 1) = oRec("jobid") 'useJobid(x) 
					        midjiddato(oRec("mid"),n, 2) = oRec("mid") 'xmid(x) 
					        midjiddato(oRec("mid"),n, 3) = showresdato 
					        midjiddato(oRec("mid"),n, 4) = showresuge
        					
				        end if
				        
				   else
				   '*** MD / MD tjkker ikke om uge passer **'
				            
				            
				         if cdbl(lastMid) = oRec("mid") AND cdbl(lastJid) = oRec("jobid") AND _
				        cint(lastMonth) = datepart("m", showresdato) AND _
				        datepart("yyyy", showresdato) = lastYear then
        				
        				
	                        midjiddato(oRec("mid"),n, 0) = midjiddato(oRec("mid"),n, 0) + medarbKundeoplysX(x, 10) 'oRec("restimer") 'showrestimer
        					 
				        else
        					
					        'Response.Write "<br><br><br><br><br><br><br><br><br><br>x:"& x &" arr:"& oRec("resmd") &","& right(oRec("resaar"),2) &","& oRec("jobid") &","& oRec("mid") & "<br>"
	                        'Response.Flush
        					
					        'testarr(oRec("resuge"),right(oRec("resaar"),2),oRec("jobid"),oRec("mid")) = oRec("restimer")
        	                
        	                
        	                
					        n = n + 1
					        'Response.Write "N: " & n & "<br>"
					        'Response.flush
					        midjiddato(oRec("mid"),n, 0) = medarbKundeoplysX(x, 10) 'oRec("restimer") 'showrestimer(x) 
					        midjiddato(oRec("mid"),n, 1) = oRec("jobid") 'useJobid(x) 
					        midjiddato(oRec("mid"),n, 2) = oRec("mid") 'xmid(x) 
					        midjiddato(oRec("mid"),n, 3) = showresdato 
					        midjiddato(oRec("mid"),n, 4) = showresuge
        					
				        end if
				   
				   
				   end if
			
			
			lastWeek = showresuge	
			lastYear = datepart("yyyy", showresdato) 
			lastMonth = datepart("m", showresdato, 2,2) 
			lastMid = oRec("mid")
			lastJid = oRec("jobid")
			
			end if
	
		
	
	x = x + 1
	oRec.movenext
	wend
	oRec.close 
	
	if t = 0 then
	    if len(trim(jobisWrt)) <> 0 then
	    jobisWrt = jobisWrt & ")"
	    else
	    jobisWrt = ""
	    end if
	end if
	
	next	
	
	%>
	


<form action="ressource_belaeg_jbpla.asp?func=redalle&medarbSel=<%=medarbSel%>&showonejob=<%=showonejob%>&id=<%=id%>" method="post"><!--name="alledagemedarb" id="alledagemedarb"-->

<!--


<input type="hidden" name="inklweekend_2" id="inklweekend_2" value="0">
-->
<input type="hidden" name="mdselect" id="Hidden2" value="<%=monththis%>">
<input type="hidden" name="aarselect" id="aarselect" value="<%=yearthis%>">
<input type="hidden" name="periodeselect" id="periodeselect" value="<%=periodeSel%>">
<input type="hidden" name="jobselect" id="Hidden1" value="<%=jobnavn%>">

<%if media <> "print" AND media <> "eksport" then%>
<table cellspacing=0 cellpadding=0 border=0 width=100%>
		<tr><td align=right style="padding:10px 30px 5px 0px;">
		<input type="submit" value="Indlæs forecast >> "></td></tr>
</table>
<%end if %>


<%if media <> "eksport" then

%>
<table border=0 cellspacing=1 cellpadding=0 bgcolor="#8caae6" width=100%>

<%end if

MedIdSQLKri = ""
MedTotSQLKri = ""	
'MedJobFcastSQLKri = ""

select case periodeSel
case 1
numoffdaysorweeksinperiode = datediff("d", startdato, slutdato)
case 3
numoffdaysorweeksinperiode = datediff("ww", startdato, slutdato)
case 6, 12
numoffdaysorweeksinperiode = datediff("m", startdato, slutdato)
end select

y = 0 
nday = startdato
weekrest = 0		
antalxx = x
xx = 0
foundone = 0
daynowrpl = replace(daynow, "/", "-")
lastMonth = monththis

dagsdato = now

if media <> "eksport" then
call datoeroverskrift(1)
else
call datoeroverskriftCSV(1)
end if

'public MedTotSQLKri
'strjIdWrt = "0#," 


'**** Medarbejdere ****
for x = 0 to antalxx - 1 'UBOUND(medarbKundeoplysX, 1) 'medarbKundeoplysX(x, 0)

	
	if lastxmid <> medarbKundeoplysX(x, 3) then	'xmid(x) 
	
	tdcolspan = 5
	strDageChkboxOne = "<input type=text size=10 name=FM_sel_dage id=FM_sel_dage value="&daynowrpl&">&nbsp;dd-mm-aaaa"
	
	if x > 0 then				
		
		'if media <> "print" then
		 '   if cint(session("mid")) = cint(lastxmid) OR level = 1 OR session("mid") = lastJobans1 OR session("mid") = lastJobans2 then
		    'call tilfojNylinie(lastxmid)
		  '  end if
		'end if
		if media <> "eksport" then
		
		if cint(jobidsel) = 0 then 
	    call medarbtotal
	    end if
	    
	    end if
	    
	end if
	
	            
	            '*** Tjekker om job allerede er forecastet for denne medarbejder i valgete periode
                 strjIdWrt = "#," 
                 For q = 0 to UBOUND(midjiddato)
		         strjIdWrt = strjIdWrt &"#"& midjiddato(medarbKundeoplysX(x, 3),q, 1) & "#,"
		         next 
	
	%>
	
	
	<%if media <> "eksport" then %>
	<tr>
		<td bgcolor="#EFF3FF" style="padding:2px 0px 3px 5px; width:250px;">
		<b><%=medarbKundeoplysX(x, 4)%> (<%=medarbKundeoplysX(x, 7)%>)</b>
		
		
		
		<%if media <> "print" AND (instr(strjIdWrt, "#"&jobidSel&"#") = 0) then %>
		<a href="#" id="anl_<%=medarbKundeoplysX(x, 3)%>" class="vmenu">[Tilføj +]</a> &nbsp;
		<%end if%>
		<!--<a href="#" onClick="showtildeltimer('<=xmid(x)%>','<=strDageChkboxOne%>','<=id%>','<=periodeSel%>')" class=vmenuglobal>Ny</a>-->
		
		<%if len(trim(medarbKundeoplysX(x, 8))) <> 0 then
		Response.write " - " & medarbKundeoplysX(x, 8)
		end if
		%><br>
		Sidst opd: <%=medarbKundeoplysX(x, 9)%>
		
		<br />Norm timer pr. md ~ <%
		call normtimerPer(medarbKundeoplysX(x, 3), jobStartKri, datoInterval)
		Response.write "<b>"& formatnumber(ntimPer/periodesel, 1) & "</b>  timer" 
		%>
		
		</td>
		
		
		
		<%
		cspan = rdimnb  '12
		%>
		
		<td bgcolor="#FFFFFF" align=right style="padding-right:5px;" colspan=<%=cspan+2%>>
		
		
		
		
		</td>
		
	</tr>
	<%
	end if
	
	
		       if media <> "print" AND media <> "eksport" AND (instr(strjIdWrt, "#"&jobidSel&"#") = 0) then 
				    
				    'if (cint(session("mid")) = cint(lastxmid) OR level = 1 OR session("mid") = lastJobans1 OR session("mid") = lastJobans2) then
		                call tilfojNylinie(medarbKundeoplysX(x, 3))
				    'end if
				    
				end if 'Print
		
		
	
	
	
	lastJid = 0
	
	MedIdSQLKri = medarbKundeoplysX(x, 3)
	csvTmnavn = medarbKundeoplysX(x, 4)
	csvTmnr = medarbKundeoplysX(x, 7)
	csvTinit = medarbKundeoplysX(x, 8)
	csvTnormuge = ntimPer
	end if
	
		
		
		
		
		if lastJid <> medarbKundeoplysX(x, 0) OR lastxmid <> medarbKundeoplysX(x, 3) AND len(medarbKundeoplysX(x, 0)) <> 0 then
		
		if media <> "eksport" then
		%>
		
		<tr>
		<td bgcolor="#ffffff" style="padding-left:5px; padding-right:5px; width:250px;"><b><%=medarbKundeoplysX(x, 2)%> (<%=medarbKundeoplysX(x, 1)%>)</b><br />
		
		<font class=medlillesilver>
		<%if len(trim(medarbKundeoplysX(x, 13))) <> 0 then %>
	    <%=medarbKundeoplysX(x, 13) %>
		<%end if %>
		
		<%if len(trim(medarbKundeoplysX(x, 15))) <> 0 then %>
		, <%=medarbKundeoplysX(x, 15) %>
		<%end if %>
		</font>
		
		<%if len(trim(medarbKundeoplysX(x, 13))) <> 0 OR len(trim(medarbKundeoplysX(x, 15))) <> 0 then %>
		<br />
		<%end if %>
		
		<%=medarbKundeoplysX(x, 5)%> (<%=medarbKundeoplysX(x, 6)%>)</td>
		
		

		
		
	    <%end if 'media
		
        'Response.Write "numoffdaysorweeksinperiode" & numoffdaysorweeksinperiode & "<hr>"
		
		'**************************************************************
		'***** Udkskriver forecast timer pr. md. **********************
		'**************************************************************
		
		'strjIdWrt = strjIdWrt & "#"&medarbKundeoplysX(x, 0)&"#," 
		lastJid = medarbKundeoplysX(x, 0)
	    			
		if y = 0 then
	    forcastTimerJobtot = 0
        end if			
					lastMonth = monththis
					y = 0
					for y = 0 TO numoffdaysorweeksinperiode
					    
					   
					    'Response.Write periodeSel &","& y &","& startdato &","& monthUse & "<br>"
						
						call antaliperiode(periodeSel, y, startdato, monthUse)
					    'Response.Write "<br>Nday: "& nday 
								    
									timerThis = 0	
									For n = 0 to UBOUND(midjiddato, 1) AND foundone = 0
									
								    'Response.Write midjiddato(n, 4) & " = "& newWeek & "<br>"
									'***** Uge visning, tjkker ikke MD ****'
									if periodeSel = 3 then
									
								        if (midjiddato(medarbKundeoplysX(x, 3),n, 4) = newWeek) AND _
								          midjiddato(medarbKundeoplysX(x, 3),n, 1) = medarbKundeoplysX(x, 0) AND midjiddato(medarbKundeoplysX(x, 3),n, 2) = medarbKundeoplysX(x, 3) AND _
								           datepart("yyyy", midjiddato(medarbKundeoplysX(x, 3),n, 3)) = datepart("yyyy", nDay) then
    								        
								            timerThis = midjiddato(medarbKundeoplysX(x, 3),n, 0)
    										
										    foundone = 1
    										
    										
									    end if
								    
								    
								    else
								    '*** MD visning tjekker ikke om uge passer **'
								    
								    if (datepart("m", midjiddato(medarbKundeoplysX(x, 3),n, 3)) = datepart("m", nDay)) AND _
								        midjiddato(medarbKundeoplysX(x, 3),n, 1) = medarbKundeoplysX(x, 0) AND midjiddato(medarbKundeoplysX(x, 3),n, 2) = medarbKundeoplysX(x, 3) AND _
								        datepart("yyyy", midjiddato(medarbKundeoplysX(x, 3),n, 3)) = datepart("yyyy", nDay) then
								        
								        timerThis = midjiddato(medarbKundeoplysX(x, 3),n, 0)
										
										foundone = 1
										
										
									'Response.Write "<br>Nday: "& nday &" - "&  datepart("m", midjiddato(n, 3)) &" = " & datepart("m", nday) &" AND " _
								    '  & midjiddato(n, 4) & " = "& newWeek & " AND " _
								    '  & midjiddato(n, 1) &" = "& medarbKundeoplysX(x, 0) &" AND "& midjiddato(n, 2) &" ="& medarbKundeoplysX(x, 3) &" AND " _
								    '  & datepart("yyyy", midjiddato(n, 3)) &" = "& datepart("yyyy", nday)
								    end if
								    
								    
								    end if
									
									
									loops = loops + 1
									next
									'Response.Write "loops " & loops
									
									'if media <> "eksport" then 
									Response.flush
									'end if
									
									loops = 0
								    'timerThis = testarr(datepart("ww", nday, 2,2),right(datepart("yyyy", nday), 2),medarbKundeoplysX(x, 0),medarbKundeoplysX(x, 3)) 
									'timerthis = midjiddato(n, 0)
								
						foundone = 0
						xx = 0
					
					
					
					
					if media <> "eksport" then
					
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
					bgTimer = "#FF0066"
					end if
					
					
					end if
					
					lastMonth = newMonth
					dato = nday
					stThisDato = dato
					txtwidth = 60
					
					
				    if timerThis = 0 then
					showTimerThis = ""
					else
					showTimerThis = timerThis
					end if
					
					if media <> "eksport" then
					%>
					<td valign=top bgcolor="<%=bgTimer%>" style="padding:2px 2px 2px 2px; width:<%=txtwidth%>px;">
					<%
					end if
					
					'*** Realiserede timer /medarbejder på job pr. md
					
					sumTimer = 0
					
					if periodeSel <> 3 then
					mmwwSQLkri = " MONTH(tdato) = "& datepart("m", nday, 2, 2) 
					else
					mmwwSQLkri = " WEEK(tdato, 1) = "& datepart("ww", nday, 2, 2) 
					end if
															
				    strSQLtimer = "SELECT sum(timer) AS sumtimerFaktisk FROM timer WHERE tmnr = "& medarbKundeoplysX(x, 3) &" AND tjobnr = "&medarbKundeoplysX(x, 1)&" AND ("& aty_sql_realhours &") AND (YEAR(tdato) = "& datepart("yyyy", nday, 2, 2) &" AND " & mmwwSQLkri &")"
				    'Response.Write strSQLtimer
				    'Response.Flush
				    'Response.end
				    
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
				    end if
				    
				    if len(trim(showTimerThis)) <> 0 then
				    timerForecast = showTimerThis
				    else
				    timerForecast = 0
				    end if
				    
				    
				    '*** Afvigelse %
				    
				        if cint(timerForecast) <= cint(sumTimer) then
						    if sumTimer <> 0 then
                            afvThis = 100 - (timerForecast/sumTimer) * 100
                            else
                            afvThis = 0
                            end if
                        else
                            if timerForecast <> 0 then
                            afvThis = 100 - (sumTimer/timerForecast) * 100
                            'afvThis = replace(afvThis, "-", "")
                            else
                            afvThis = 0
                            end if
                        end if
				    
				   
				   'Response.Write "timerForecast:" & timerForecast & "<br>"
				   'Response.Write "sumTimer: " & sumTimer & "<br>"
				   
				   	if media <> "print" AND media <> "eksport" then
				   	
				   	         
            				   	
				   	            if (cint(session("mid")) = cint(medarbKundeoplysX(x, 3)) OR _
				   	            session("mid") = medarbKundeoplysX(x, 11) OR _
				   	            session("mid") = medarbKundeoplysX(x, 12)) AND _
				   	            ((month(cdate(dagsdato)) <= month(cdate(nday)) AND year(cdate(dagsdato)) = year(cdate(nday))) OR (year(cdate(dagsdato)) < year(cdate(nday)))) OR level = 1 then
				   	            iptype = "text"
				   	            else 
				   	            iptype = "hidden"
				   	            end if
				   	            %> 
                				
                				
    				            <input type="<%=iptype%>" name="FM_timer" id="FM_timer_<%=medarbKundeoplysX(x, 3) %>_<%=medarbKundeoplysX(x, 0) %>_<%=y%>" style="width:50px; font-size:10px;" value="<%=showTimerThis %>">
                                
                                <%if iptype = "text" then%>
                                <input id="Button1" type="button" value=">>" style="font-size:8px;" onclick="copyTimer('<%=medarbKundeoplysX(x, 3) %>','<%=medarbKundeoplysX(x, 0) %>', '<%=y %>')" />
    				            <%else %>
    				            <b><%=showTimerThis %></b>
    				            <%end if %>
                                
    				            <!--'*** hidden array felt. *** Bruges så der kan skelneshvis man bruger komma.-->
					            <input type="hidden" name="FM_timer" id="FM_timer" value="#">
					<%else 
					
					        if media <> "eksport" then%>
					        <b><%=showTimerThis %></b>
					        <%end if %>
					        
					<%end if 'Print / media%>
					
					
					<%if media <> "eksport" then %>
					
					<br />
					
					<%'if timerForecast <> 0 then %>
					 <font class=sort8px><%=formatnumber(sumTimer, 0)%> ~ <%=formatnumber(afvThis, 0)%>%</font>
					<%'end if %>
					
					<%
					
					end if 'media
					
					csvTxt = csvTxt & vbcrlf & medarbKundeoplysX(x, 4) &" ("& medarbKundeoplysX(x, 7) &")"
					csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 8) 
					csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 9) 
					csvTxt = csvTxt & ";"& formatnumber(ntimPer, 2) 
					csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 2) & ";" & medarbKundeoplysX(x, 1)
					
					csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 13) & ";" & medarbKundeoplysX(x, 14)
					csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 15) & ";" & medarbKundeoplysX(x, 16)
					
					csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 5) & ";" & medarbKundeoplysX(x, 6)
					
					'if lto <> "execon" then
					csvTxt = csvTxt & ";"& csvDatoerUge(y) 
					'end if
					
					csvTxt = csvTxt & ";"& csvDatoerMD(y) & ";"& csvDatoerAR(y) &";"& timerForecast & ";"& sumTimer &";"& formatnumber(afvThis, 0)
					'csvTxt = csvTxt & vbcrlf & ";;;;;;;"& csvDatoer(y) &";" & timerThis%>
					
					<%if media <> "eksport" then  %>
					
					<input type="hidden" name="FM_jobid" id="Hidden3" value="<%=medarbKundeoplysX(x, 0)%>">
					<input type="hidden" name="FM_medarbid" id="FM_medarbid" value="<%=medarbKundeoplysX(x, 3)%>">
					<input type="hidden" name="FM_dato" id="FM_dato" value="<%=stThisDato%>">
					
					
					</td>
					
					<%
					end if
					
					medarbTimerReal(y) = medarbTimerReal(y) + sumTimer
					medarbTotalTimer(y) = medarbTotalTimer(y) + timerForecast
					timerTotalTotal(y) = timerTotalTotal(y) + timerForecast
					realTotalTotal(y) = realTotalTotal(y) + sumTimer
					
					realTimerJobtot = realTimerJobtot + sumTimer 
					forcastTimerJobtot = forcastTimerJobtot + timerForecast 
					next
					
					
					if media <> "eksport" then %>
                    <td bgcolor="#ffffff" class=lille align=right style="padding:2px 2px 2px 2px;"><b><%=formatnumber(forcastTimerJobtot, 0) %></b></td>
			        <td bgcolor="#ffffff" class=lille align=right style="padding:2px 2px 2px 2px;"><%=formatnumber(realTimerJobtot, 0)%></td>	
		            <%
            		
		            end if
		realTimerJobtotGrand = realTimerJobtotGrand + realTimerJobtot 
		forecastTimerJobtotGrand = forecastTimerJobtotGrand + forcastTimerJobtot
		
		pageForecastGrandtot = pageForecastGrandtot + forcastTimerJobtot 
		pageRealGrandtot = pageRealGrandtot + realTimerJobtot
		
		realTimerJobtot = 0 
		forcastTimerJobtot = 0
		
		end if '** Lastjobid
		
		
		lastJobans1 = medarbKundeoplysX(x, 11)
		lastJobans2 = medarbKundeoplysX(x, 12)
		lastxmid = medarbKundeoplysX(x, 3)
		
		if media <> "eksport" then%>
		
		</tr>
		<% end if
		'y = y + 1	
		'Response.Flush
		next
		
		
		
		
		
		if media <> "eksport" then
		
		'*****************************************
		'*** Tildel timer  / Guiden aktive job****
		'*****************************************
		
	
		
				'*** Henter guiden dine aktive job
				
				
				
				if media <> "print" then
				    
				    if (cint(session("mid")) = cint(lastxmid) OR level = 1 OR session("mid") = lastJobans1 OR session("mid") = lastJobans2) then
		                'call tilfojNylinie(lastxmid)
				    end if
				    
				end if 'Print
				
				
				if cint(jobidsel) = 0 then 
				'**** Medarb. total ***'
				call medarbTotal
				end if
				
				
				if medarbSel = "0" then
				%>
				<!--- Grandtotal --->
				
				<tr>
					<td height=20 bgcolor="#ffffff" align=right style="padding:2px 5px 2px 2px;"><b>Total forecast:</b><font class=megetlillesort> (alle job / medarbejdere)<br />
					Real. ~ Afv.%</font></td>
					
					<%
					
						for y = 0 TO numoffdaysorweeksinperiode
						    
						'*** Afvigelse ***
			            tottotAfv = 0
    			
    				    
				        if timerTotalTotal(y) >= realTotalTotal(y) then
                            if timerTotalTotal(y) <> 0 then
                            tottotAfv = 100 - (realTotalTotal(y)/timerTotalTotal(y)) * 100
                            else
                            tottotAfv = 0
                            end if
                        else
                            if realTotalTotal(y) <> 0 then
                            tottotAfv = 100 - (timerTotalTotal(y)/realTotalTotal(y)) * 100
                            else
                            tottotAfv = 0
                            end if
                        end if
						
						%>
							
						<td bgcolor="#ffffff" class=lille style="padding:2px 5px 2px 2px;"><b><%=formatnumber(timerTotalTotal(y),0)%></b> <br /> <%=formatnumber(realTotalTotal(y), 0)%> ~ <%=formatnumber(tottotAfv,0) %>%</td>
						<%
						next
						%>
						<td bgcolor="#ffffff" style="padding:2px 5px 2px 2px;" align=right><b><%=formatnumber(pageForecastGrandtot,0) %></b></td>
						<td bgcolor="#ffffff" style="padding:2px 5px 2px 2px;" align=right><%=formatnumber(pageRealGrandtot,0) %></td>
				</tr>
				<%
				end if
				
				
		%>
		</table>
		
            		
		            <%if media <> "print" then%>
		                        <table cellspacing=0 cellpadding=0 border=0 width=100%>
		                        <tr><td align=right style="padding:10px 30px 5px 0px;">
		                        <input type="submit" value="Indlæs forecast >> "></td></tr>
                        </table>
                        <input id="ymax" name="ymax" value="<%=y-1 %>" type="hidden" />
                        </form> 
                                    
                        
                        
                        
                        <%end if %>
            
            
             <%end if '** eksport %>
        
</div>		

		
		<%
		if medarbSel = "111110" then
		
		  
				'*** Totaler foregående 3 md. uanset job. ***
				strSQL = "SELECT sum(rmd.timer) AS restimer FROM ressourcer_md rmd "_
				&" LEFT JOIN job j ON (j.id = jobid AND j.jobnavn IS NOT NULL)"_
				&" WHERE medid = "& medarbSel &" AND  "_
	            &" ((rmd.md >= "& startdatoMDtot &" AND rmd.aar = "& startdatoYYtot &") "_
	            &" AND (rmd.md <= "& slutdatoMDtot &" AND rmd.aar = "& slutdatoYYtot &")) AND j.jobnavn IS NOT NULL GROUP BY medid "
            	
	            
	            'Response.Write strSQL &"<br>"
            	
            	totResTimerForrige = 0
				oRec.open strSQL, oConn, 3
                if not oRec.EOF then
                
                    totResTimerForrige = oRec("restimer") 
                
                end if
                oRec.close
				
				strSQL = "SELECT sum(timer) AS realtimer FROM timer t"_
				&" WHERE tmnr = "& medarbSel &" AND  "_
	            &" tdato BETWEEN '"& timregStdato &"' AND '"& timregSldato &"' AND tfaktim <> 5 GROUP BY tmnr"

            	
	            'Response.Write strSQL &"<br>"
	            'Response.flush
            	
            	totRealTimerForrige = 0
				oRec.open strSQL, oConn, 3
                if not oRec.EOF then
                
                    totRealTimerForrige = oRec("realtimer")
                
                end if
                oRec.close			
				
				
				
				'*** Afvigelse ***
			    totAfvForrige = 0
    			
    				    
				        if totResTimerForrige >= totRealTimerForrige then
                            if totResTimerForrige <> 0 then
                            totAfvForrige = 100 - (totRealTimerForrige/totResTimerForrige) * 100
                            else
                            totAfvForrige = 0
                            end if
                        else
                            if totRealTimerForrige <> 0 then
                            totAfvForrige = 100 - (totResTimerForrige/totRealTimerForrige) * 100
                            else
                            totAfvForrige = 0
                            end if
                        end if
    				    
             
             
                Response.Write "<br><br><br><br><br><br><table cellspacing=0 cellpadding=5 border=0><tr><td bgcolor=#FFFFFF style='border:1px #8cAAe6 solid;'><b>Totaler forrige 3 md. (alle job uanset viste):</b><br>"
                 Response.write "Forecast: "& formatnumber(totResTimerForrige, 0) &" timer<br>"
    		 Response.write "Realiseret: "& formatnumber(totRealTimerForrige, 0) &" timer<br>"
            Response.write "Afvigelse: "& formatnumber(totAfvForrige, 0) &"%</td></tr></table>"

	 end if '** Medarb Sel **' %>
	 
	 
	 <%if media <> "print" AND media <> "eksport" then
     %>	 
	 <!--pagehelp-->

<%

itop = 56
ileft = 800
iwdt = 180
ihgt = 0
ibtop = 5 
ibleft = 0
ibwdt = 600
ibhgt = 450
iId = "pagehelp"
ibId = "pagehelp_bread"
call sideinfoId(itop,ileft,iwdt,ihgt,iId,phDsp,phVzb,ibtop,ibleft,ibwdt,ibhgt,ibId)
%>
	<b>Hvordan du bruger ressource forecast</b><br /><br /> 
	
	<ul>
	<li> Før der kan forecastes timer på en medarbejderen skal denne være være tildelt adgang til jobbet via dennes projektgrupper, og jobbet skal være aktivt i medarbejderens personlige aktiv liste.
	<li> Du kan tildele timer på dig selv, eller på andre medarbejdere hvis du er administrator eller er jobansvarlig på de pågældende job.		
	<li> Tildel timer ved at vælge det ønskede job i select-boksen under hver enkelt medarbejder.
	<li> Timer angives som samlet antal timer pr. måned eller uge. (afhængig af valgt periode)
	<li> Slet forecast ved at indtaste et 0 (nul).<br /><br />	
	<li> Normeret uge er standard normeret uge, uanset helligdage i periode. 	
	<li> Tal i sorte rammer under text-felter er <b>realiserede timer / afvigelse i %</b>
	<li> Når en måned bliver passeret kan der ikke længere ændres i forecast. Dette gælder også ved visning på uge basis, der kan der også ændres i forecast frem til måneds afslutning.
	Administratorer kan godt ændre i lukkede perioder.
	</ul>
	
	</td>
	</tr>
	</table>
	
	</div>
	
	<%end if 'Print / Media%>
	 
	 
	 <%if media <> "eksport" then %>
	 <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
	 
	 
	 <%end if
				
		
		
		'***************** Ekasport *************************'
        if media = "eksport" then
           
               ekspTxt = csvTxtTop & csvTxt & csvTxtTotal
               ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
               'kspTxt = replace(ekspTxt, ";", chr(34))
               'datointerval = request("datointerval")
                       
                       
                filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	            filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
            	
	            Set objFSO = server.createobject("Scripting.FileSystemObject")
            	
	            if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\ressource_belaeg_jbpla.asp" then
            							
		            Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
		            Set objNewFile = nothing
		            Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
            	
	            else
            		
		            Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, false)
		            Set objNewFile = nothing
		            Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8, -1)
            		
	            end if
            	
	            file = "ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
            	
            	
	            objF.WriteLine(ekspTxt)
            	
            	
	            objF.close
	            %>
	            <table border=0 cellspacing=1 cellpadding=0 width="200">
	            <tr><td valign=top bgcolor="#ffffff" style="padding:5px;">
	            <img src="../ill/outzource_logo_200.gif" />
	            </td>
	            </tr>
	            <tr>
	            <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
	            <a href="../inc/log/data/<%=file%>" class=vmenu target="_blank" onClick="Javascript:window.close()">Din CSV. fil er klar >></a>
	            </td></tr>
	            </table>
	            
	             <script>
	                 document.getElementById("load").style.visibility = "hidden";
	                 document.getElementById("load").style.display = "none";
			    </script>
	            
	            <%
                
                
                Response.end
	            Response.redirect "../inc/log/data/"& file &""	

         end if
                
                
                
                
                
                
                
                
           if media <> "print" AND media <> "eksport" then
            
           
            ptop = 72
            pleft = 1020
            pwdt = 180
            
           
            call eksportogprint(ptop, pleft, pwdt)
            %>
               
                                 
            <tr>
                <td align=center>
                <a href="#" onclick="Javascript:window.open('ressource_belaeg_jbpla.asp?media=eksport&medarbSel=<%=medarbSel%>&jobselect=<%=jobnavn%>&mdselect=<%=monththis%>&aarselect=<%=yearThis%>', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=rmenu><img src="../ill/export1.png" border=0></a>
                </td>
                
                </td><td>
                <a href="#" onclick="Javascript:window.open('ressource_belaeg_jbpla.asp?media=eksport&medarbSel=<%=medarbSel%>&jobselect=<%=jobnavn%>&mdselect=<%=monththis%>&aarselect=<%=yearThis%>', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=rmenu>.csv fil eksport (pivot)</a>
                </td>
                </tr>
                
                <tr>
                <td align=center><a href="ressource_belaeg_jbpla.asp?media=print&medarbSel=<%=medarbSel%>&jobselect=<%=jobnavn%>&mdselect=<%=monththis%>&aarselect=<%=yearThis%>" target="_blank"><img src="../ill/printer3.png" border="0"></a>
                                 
                             </td><td><a href="ressource_belaeg_jbpla.asp?media=print&medarbSel=<%=medarbSel%>&jobselect=<%=jobnavn%>&mdselect=<%=monththis%>&aarselect=<%=yearThis%>" target="_blank" class=rmenu>Print version</a></td>
               </tr>
               </table>
            </div>
            
            <%
            else
            Response.Write("<script language=""JavaScript"">window.print();</script>")
            end if%>
           
           
           </div>
			
			
			
		
	
	
	<br><br>&nbsp;
	
</div>




<br><br>&nbsp;<br><br>&nbsp;<br><br>&nbsp;<br><br>&nbsp;<br><br>&nbsp;<br><br>&nbsp;




	
<%end if%>

            <script>
			document.getElementById("load").style.visibility = "hidden";
			document.getElementById("load").style.display = "none";
			</script>

<!--#include file="../inc/regular/footer_inc.asp"-->

	
