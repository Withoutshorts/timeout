<%response.buffer = true%>


<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->

<!--#include file="../inc/regular/global_func.asp"-->


<!--#include file="inc/dato2.asp"-->

<!--#include file="inc/isint_func.asp"-->
<!--#include file="../inc/regular/job_func.asp"-->
<!--#include file="inc/webblik_joblisten_inc.asp"-->


<!--#include file="../inc/regular/topmenu_inc.asp"-->


<!--#include file="exch_meeting_webdav.asp"-->
<!--#include file="exch_meeting_webdav_del.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	public col_1
	function colors(one_or_two)
	
							select case one_or_two
							case 1, 47, 93
							colThis = "#66CC33"
							case 2, 48, 94
							colThis = "Tomato"
							case 3, 49, 95
							colThis = "#FF9933"
							case 4, 50, 96
							colThis = "#9999FF"
							case 5, 51, 97
							colThis = "#666699"
							case 6, 52, 98
							colThis = "#FFCC00"
							case 7, 53, 99
							colThis = "#34A4DE"
							case 8, 54, 100
							colThis = "#B4E2FF"
							case 9, 55, 101
							colThis = "#003399"
							case 10, 56, 102
							colThis = "#FFCCFF"
							case 11, 57, 103
							colThis = "#CCCCFF"
							case 12, 58, 104
							colThis = "#996600"
							case 13, 59, 105
							colThis = "#CC9900"
							case 14, 60, 106
							colThis = "#FFFF99"
							case 15, 61, 107
							colThis = "#333366"
							case 16, 62, 108
							colThis = "#FFFF00"
							case 17, 63, 109
							colThis = "#FFDB9D"
							case 18, 64, 110
							colThis = "#FFCC66"
							case 19, 65, 111
							colThis = "#0066CC"
							case 20, 66, 112
							colThis = "#FF794B"
							case 21, 67, 113
							colThis = "#FF3300"
							case 22, 68, 114
							colThis = "#990000"
							case 23, 69, 115
							colThis = "#0083D7"
							case 24, 70, 116
							colThis = "#6666CC"
							case 25, 71, 117
							colThis = "#9999CC"
							case 26, 72, 118
							colThis = "#0099FF"
							case 27, 73, 119
							colThis = "#006600"
							case 28, 74, 120
							colThis = "#006699"
							case 29, 75, 121
							colThis = "#DEFFFF"
							case 30, 76, 122
							colThis = "Silver"
							case 31, 77, 123
							colThis = "LightGrey"
							case 32, 78, 125
							colThis = "DarkGray"
							case 33, 79
							colThis = "Olive"
							case 34, 80
							colThis = "Dimgray"
							case 35, 81
							colThis = "#99FF66"
							case 36, 82
							colThis = "#CCFFCC"
							case 37, 83
							colThis = "RosyBrown"
							case 38, 84
							colThis = "Honeydew"
							case 39, 85
							colThis = "CornflowerBlue"
							case 40, 86
							colThis = "DarkMagenta"
							case 41, 78
							colThis = "skyblue"
							case 42, 88
							colThis = "Thistle"
							case 43, 89
							colThis = "#009900"
							case 44, 90
							colThis = "Gold"
							case 45, 91
							colThis = "yellow"
							case 46, 92
							colThis = "black"
							case else
							colThis = "#5582d2"
							end select
	col_1 = colThis
	end function
	
	public dheight, feltbase 
	
	sub dtosub
	
		select case dtop
		case 15
		feltbase = feltbase
		case 30
		feltbase = feltbase + 1
		case 45
		feltbase = feltbase + 2
		case 60
		feltbase = feltbase + 3
		case 75
		feltbase = feltbase + 4
		case 90
		feltbase = feltbase + 5
		case 105
		feltbase = feltbase + 6
		case 120
		feltbase = feltbase + 7
		case 135
		feltbase = feltbase + 8
		case 150
		feltbase = feltbase + 9
		case 165
		feltbase = feltbase + 10
		case 180
		feltbase = feltbase + 11
		case 195
		feltbase = feltbase + 12
		case 210
		feltbase = feltbase + 13
		case 225
		feltbase = feltbase + 14
		case 240
		feltbase = feltbase + 15
		case 255
		feltbase = feltbase + 16
		case else
		'**etc til kl 17.00
		end select
	
	end sub
	
	sub dhesub
	
	feltbase_start = feltbase
	
		select case dheight
		case 15
		feltbase = feltbase 
		case 30
		feltbase = feltbase + 1
		case 45
		feltbase = feltbase + 2
		case 60
		feltbase = feltbase + 3
		case 75
		feltbase = feltbase + 4
		case 90
		feltbase = feltbase + 5
		case 105
		feltbase = feltbase + 6
		case 120
		feltbase = feltbase + 7
		case 135
		feltbase = feltbase + 8
		case 150
		feltbase = feltbase + 9
		end select 
		
		
		
		
		%>
		<script>
		var fss = <%=feltbase_start%>;
		var fb = <%=feltbase%>;
		//var thisFeltValue = 0;
		for (i=fss;i<=fb;i++){
			if (i==fss) {
				if (parseInt(document.getElementById("felt_"+i).value) != 0) {
				thisFeltValue = parseInt(thisFeltValue) + 1
				document.getElementById("felt_"+i).value = thisFeltValue 
				} else {
				document.getElementById("felt_"+i).value = 1
				thisFeltValue = 1
				}
			} else {
			document.getElementById("felt_"+i).value = parseInt(thisFeltValue)  
			}
		}
		</script>
		<%
		'next
	end sub
	
	
	function slettetbokingEmail(rcNavn, rcEmail, dato, sttid, sltid)
	
	'*** Afsender ***
	if lto = "kringit" then
		seEmail = "connect@kringitsolutions.dk"
		seNavn = "KRING Timeout"
	else
		strSQL = "SELECT mid, mnavn, email FROM medarbejdere WHERE mid = "& session("mid")
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then
		seEmail = oRec("email")
		seNavn = oRec("mnavn")
		end if
		oRec.close
	end if
	
	'*** Email vedr. slet ***
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
					
	' Sætter Charsettet til ISO-8859-1 
	Mailer.CharSet = 2
	' Afsenderens navn 
	Mailer.FromName = ""& seNavn
	' Afsenderens e-mail 
	Mailer.FromAddress = ""& seEmail
	Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk"
	' Modtagerens navn og e-mail
	Mailer.AddRecipient rcNavn, rcEmail
	'Mailer.AddBCC "Support", "support@outzource.dk" 
	' Mailens emne
	Mailer.Subject = "SLETTET - Booking: d. " & formatdatetime(dato, 2) & " kl. " & sttid & " - "& sltid
	
	' Selve teksten
	Mailer.BodyText = "" & "Hej "& rcNavn & vbCrLf _ 
	& "Ovenstående booking er blevet slettet." & vbCrLf & vbCrLf _ 
	
	& "Link til timeregistrering: " & vbCrLf _
	& "https://outzource.dk/"& lto & "" & vbCrLf & vbCrLf _ 
	
	& "Med venlig hilsen"& vbCrLf & vbCrLf & seNavn & vbCrLf 
	
	'& "Link til kalender (rediger denne bookning): " & vbCrLf _
	'& "https://outzource.dk/"& lto & "" & vbCrLf & vbCrLf _ 
	
	Mailer.SendMail
	
	end function
	
	
	
	
	
	
	
	'*****************************************************************************
	'*****************************************************************************
	'**** Sideindhold *************************************************************
	
	select case func 
	
	case "slet"
	
	strSQL = "SELECT r.exch_appid, r.mid, m.exchkonto, m.mnavn, m.email, r.dato, r.starttp, r.sluttp FROM ressourcer r LEFT JOIN medarbejdere m ON (m.mid = r.mid) WHERE r.id = "& id
	
	'Response.write strSQL
	'Response.flush
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
		exch_appid = oRec("exch_appid")
		rcDomKonto = replace(oRec("exchkonto"), "#", "\") 
		rcEmail = oRec("email")
		rcNavn = oRec("mnavn")
		dato = oRec("dato")
		sttid = formatdatetime(oRec("starttp"), 3)
		sltid = formatdatetime(oRec("sluttp"), 3)
	end if
	oRec.close 
	
	'*** Sletter på Exchange **
	if len(exch_appid) <> 0 then
		call delExchange(rcDomKonto, exch_appid)
	end if
	
	'*** Email ***
	call slettetbokingEmail(rcNavn, rcEmail, dato, sttid, sltid)
	
	oConn.execute("DELETE FROM ressourcer WHERE id = "&id&"")
	activeuge = request("activeuge")
	yearthis = request("yearthis")
	monththis = request("monththis")
	
	
	Response.redirect "jbpla_w.asp?activeuge="&activeuge&"&mdselect="&monththis&"&aarselect="&yearthis
	
	
	
	case "sletserie"
	serieid = request("serieid")
	
	'*** Exchange ***
	strSQL = "SELECT r.exch_appid, r.mid, m.exchkonto, m.mnavn, m.email, r.dato, r.starttp, r.sluttp FROM ressourcer r "_
	&" LEFT JOIN medarbejdere m ON (m.mid = r.mid) WHERE r.id = "& serieid &" OR serieid = "& serieid 
	
	'Response.write strSQL
	'Response.flush
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
		exch_appid = oRec("exch_appid")
		rcDomKonto = oRec("exchkonto")
		rcEmail = oRec("email")
		rcNavn = oRec("mnavn")
		dato = oRec("dato")
		sttid = formatdatetime(oRec("starttp"), 3)
		sltid = formatdatetime(oRec("sluttp"), 3)
			
			'*** Sletter på Exchange **
			if len(exch_appid) <> 0 then
				call delExchange(rcDomKonto, exch_appid)
			end if
			
			'*** Email ***
			call slettetbokingEmail(rcNavn, rcEmail, dato, sttid, sltid)
			
	oRec.movenext
	wend 
	oRec.close 
	'*****************
	
	oConn.execute("DELETE FROM ressourcer WHERE id = "& serieid &" OR serieid = "& serieid &"")
	activeuge = request("activeuge")
	yearthis = request("yearthis")
	monththis = request("monththis")
	
	
	Response.redirect "jbpla_w.asp?activeuge="&activeuge&"&mdselect="&monththis&"&aarselect="&yearthis
	
	
	
	case "reddb"
	
	divid = request("divid")
	divxval = request("divxval")
	divyval = request("divyval")
	h = request("divheight")
	thisdato = request("thisdato")
	thisdatoSQL = year(thisdato)&"/"&month(thisdato)&"/"&day(thisdato)
				
				
				
				'** Betstemmer klokkeslet **
				sttid = thisdato&" 08:00:00"
				
				for yval = 1 to 36
				sttidthis = dateadd("n", (yval*15), sttid)
				bundval = 40 + (yval*15)
				topval = 55 + (yval*15)	
				
					if cint(divyval) > cint(bundval) AND cint(divyval) <= cint(topval) then
						thisstarttid = sttidthis 
					end if
					
				next
				
				
				'*** Bestemmmer sluttid ***
				'h = request("bookdiv_height_"& medid &"_"&u&"")
				instr_h = 0
				instr_h = instr(h, "p")
				if instr_h <> 0 then
				len_h = len(h)
				left_h = left(h, len_h - 2)
				else
				left_h = h
				end if
				
				thissluttid = dateadd("n", left_h, thisstarttid)
				'divid = request("bookdiv_id_"& medid &"_"&u&"")
				
				
				'*** Afsender ***
				if lto = "kringit" then
					seEmail = "connect@kringitsolutions.dk"
					seNavn = "KRING Connect"
				else
					strSQL = "SELECT mid, mnavn, email FROM medarbejdere WHERE mid = "& session("mid")
					oRec.open strSQL, oConn, 3 
					if not oRec.EOF then
					seEmail = oRec("email")
					seNavn = oRec("mnavn")
					end if
					oRec.close
				end if
				
				
				'*** Exchange info vedr. modtager ***
				strSQL = "SELECT r.exch_appid, r.mid, m.exchkonto, m.mnavn, m.email, r.jobid FROM ressourcer r "_
				&" LEFT JOIN medarbejdere m ON (m.mid = r.mid) WHERE r.id = "& divid   
				
				'Response.write strSQL
				'Response.flush
				oRec.open strSQL, oConn, 3 
				if not oRec.EOF then
					
					exch_appid = oRec("exch_appid")
					rcDomKonto = replace(oRec("exchkonto"), "#", "\")
					rcEmail = oRec("email")
					rcNavn = oRec("mnavn")
					jobid = oRec("jobid")
						
						'*** Sletter opgave på Exchange **
						if len(exch_appid) <> 0 then
							call delExchange(rcDomKonto, exch_appid)
						end if
						
						
				end if
				oRec.close 
				'*****************
				
				
				'**** Opdaterer db ressourcer ***
				strSQL = "UPDATE ressourcer SET dato = '"& thisdatoSQL &"', starttp = '"& formatdatetime(thisstarttid, 3) &"', sluttp = '"& formatdatetime(thissluttid, 3) &"' WHERE id = " & divid 
				oConn.execute(strSQL)
				
				
				'*** jobnavn ***
				strSQLJ = "SELECT jobnavn, jobnr, k.kkundenavn, k.kkundenr, "_
				&" k.adresse, k.postnr, k.city, j.jobknr, j.beskrivelse, j.kundekpers, "_
				&" kp.navn, kp.email, kp.dirtlf, kp.mobiltlf FROM job j "_
				&" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_
				&" LEFT JOIN kontaktpers kp ON (kp.id = j.kundekpers) WHERE j.id = "& jobid 
				'Response.write strSQLJ
				'Response.flush
				oRec.open strSQLJ, oConn, 3 
				if not oRec.EOF then
					
					jobnavnognr = oRec("jobnavn") & " ("& oRec("jobnr") &")"
					kundenavnognr = oRec("kkundenavn") & " ("& oRec("kkundenr") &")"
					jobbesk = oRec("beskrivelse")
					kundenavnogadr = oRec("kkundenavn") & ", " & oRec("adresse") &", "& oRec("postnr") &" "& oRec("city")
					kundeadr = oRec("adresse") &", "& oRec("postnr") &" "& oRec("city")
					
					kpers = oRec("navn")
					kpers_email = oRec("email")
					kpers_dirtlf = oRec("dirtlf")
					kpers_mobiltlf = oRec("mobiltlf")
					
				end if
				oRec.close
				
				
				
				exch_appid_new = "timeout_"& datepart("d", now) & "-" & datepart("m", now) & "-" & datepart("yyyy", now) & "_kl_" & datepart("h", now) & "_" & datepart("n", now) & "_" & datepart("s", now) 
	
				
               'Sender notifikations mail
		        if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\jbpla_w.asp" then

				'*** Exchange integration ***
				'**** Opdaterer mødet på Exchangeserver ****
				call oprExchange(rcDomKonto, exch_appid_new, thisdato, formatdatetime(thisstarttid, 3), formatdatetime(thissluttid, 3))
				'**********
				
			    	

				
				'*** Email ***
				Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
					
				' Sætter Charsettet til ISO-8859-1 
				Mailer.CharSet = 2
				' Afsenderens navn 
				Mailer.FromName = ""& seNavn
				' Afsenderens e-mail 
				Mailer.FromAddress = ""& seEmail
				Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk"
				' Modtagerens navn og e-mail
				Mailer.AddRecipient rcNavn, rcEmail
				'Mailer.AddBCC "Support", "support@outzource.dk" 
				' Mailens emne
				Mailer.Subject = "Booking (opdateret): d. " & formatdatetime(thisdato, 2) & " kl. " & formatdatetime(thisstarttid, 3) & " - "& formatdatetime(thissluttid, 3)
				
				' Selve teksten
				Mailer.BodyText = "" & "Hej "& rcNavn & vbCrLf _ 
				& "Du er blevet booket på følgende job." & vbCrLf & vbCrLf _ 
				& "Kunde: " & kundenavnognr &"" & vbCrLf _ 
				& "Adr: " & kundeadr &"" & vbCrLf _
				& "Kontakt: " & kpers &""& vbCrLf _  
				& "Email: " & kpers_email &""& vbCrLf _
				& "Dir. Tlf: "& kpers_dirtlf &""& vbCrLf _
				& "Mobil Tlf: "& kpers_mobiltlf &  vbCrLf & vbCrLf _ 
				& "Job: "& jobnavnognr &"" & vbCrLf _ 
				& "Beskrivelse: "& jobbesk &"" & vbCrLf & vbCrLf _ 
				& "Dato: d. " & formatdatetime(dato, 2) & " kl. " & sttid & " - "& sltid  & vbCrLf & vbCrLf _ 
				
				& "Link til timeregistrering: " & vbCrLf _
				& "https://outzource.dk/"& lto & "" & vbCrLf & vbCrLf _ 
				
				& "Med venlig hilsen"& vbCrLf & vbCrLf & seNavn & vbCrLf 
				
				'& "Link til kalender (rediger denne bookning): " & vbCrLf _
				'& "https://outzource.dk/"& lto & "" & vbCrLf & vbCrLf _ 
				
				Mailer.SendMail
				

                end if 'c.
	
	activeuge = request("activeuge")
	yearthis = request("yearthis")
	monththis = request("monththis")
	
	Response.redirect "jbpla_w.asp?menu=res&activeuge="&activeuge&"&mdselect="&monththis&"&aarselect="&yearthis
	
	
	case else 
	
	
	
	if len(request("mdselect")) <> 0 then
		monththis = request("mdselect")
		
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
			
			
	
	
	
	
	
	%>

	
	
    <!--#include file="../inc/regular/header_inc.asp"-->
	<SCRIPT language=javascript src="inc/jbpla_w_func.js"></script>

	
	
	        <%
	        'call webbliktopmenu() 
                call menu_2014()
	        %>
	       

		
	
	<%leftPos = 102 '190%>
	<%topPos = 102 '70%>
	<!-------------------------------Sideindhold------------------------------------->
	
	<div id="menu" style="position:absolute; left:<%=leftPos-20%>px; top:<%=topPos-20%>px; visibility:visible;">
	<h3>Ressource Kalender</h3>
	</div>
	
	<!-- Oversigts kalender -->
	<div id="overblik" style="position:absolute; left:<%=leftPos%>; top:<%=topPos+50%>; visibility:visible;">
	
	<div id="periode" name="periode" style="position:absolute; left:158px; top:-25px; visibility:visible; display:;">
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<form action="jbpla_w.asp?menu=res" method="post" name="mdformselect" id="mdformselect">
	<tr><td style="white-space:nowrap;"><b>Vælg periode:</b><br><select name="mdselect" id="mdselect">
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
	</select><select name="aarselect" id="aarselect">
	<option value="<%=yearthis%>" SELECTED><%=yearthis%></option>
	<option value="2002">2002</option>
	<option value="2003">2003</option>
	<option value="2004">2004</option>
	<option value="2005">2005</option>
	<option value="2005">2006</option>
	<option value="2005">2007</option>
	<option value="2005">2008</option>
	</select></td><td style="padding-top:15px;">&nbsp;<input type="submit" value=" >> " /></td></tr>
	</form>
	</table>
	</div>
	
	<%
	
	mthDaysUse = mthDays
	monthUse = monththis
	yearUse = yearthis
	
	daynow = day(now)&"/"&month(now)&"/"&year(now)
	denneuge = datepart("ww", day(now)&"/"&month(now)&"/"&year(now), 2, 2)
	lastweek = 0
	
		'** Finder aktive uge **
		if len(request("activeuge")) <> 0 then
		activeuge = request("activeuge")
		else
		
		for d = 0 to mthDaysUse - 1
		ugeday = d + 1
		ugeweekdato = ugeday&"/"&monththis&"/"&yearUse
		
			if d = 0 then
			activeuge = datepart("ww", ugeweekdato, 2, 2)
			end if 
			
			if denneuge = datepart("ww", ugeweekdato, 2, 2) then
			activeuge = datepart("ww", ugeweekdato, 2, 2)
			else
			activeuge = activeuge
			end if
			
		next
		end if
	
		
			
		firstdayofyear = "1/1/"&yearthis
		weekdayfirstday = dateadd("ww", activeuge, firstdayofyear)
		 
		select case weekday(weekdayfirstday, 2)
		case 1
		perStartKri = weekdayfirstday
		perSlutKri = dateadd("d", 6, weekdayfirstday)
		case 2	
		perStartKri = dateadd("d", -1, weekdayfirstday)
		perSlutKri = dateadd("d", 5, weekdayfirstday)
		case 3
		perStartKri = dateadd("d", -2, weekdayfirstday)
		perSlutKri = dateadd("d", 4, weekdayfirstday)
		case 4
		perStartKri = dateadd("d", -3, weekdayfirstday)
		perSlutKri = dateadd("d", 3, weekdayfirstday)
		case 5
		perStartKri = dateadd("d", -4, weekdayfirstday)
		perSlutKri = dateadd("d", 2, weekdayfirstday)
		case 6
		perStartKri = dateadd("d", -5, weekdayfirstday)
		perSlutKri = dateadd("d", 1, weekdayfirstday)
		case 7
		perStartKri = dateadd("d", -6, weekdayfirstday)
		perSlutKri = weekdayfirstday
		end select
	
	
	jobStartKri = year(perStartKri)&"/"&month(perStartKri)&"/"&day(perStartKri)
	jobSlutKri = year(perSlutKri)&"/"&month(perSlutKri)&"/"&day(perSlutKri)						
	
	'Response.write 	"perStartKri" & perStartKri
	
	
	'****************************************************************
	'*** Ugekalender ***
	'****************************************************************
	
	
	%>
	<form action="jbpla_w.asp?func=reddb&activeuge=<%=activeuge%>&mdselect=<%=monththis%>&aarselect=<%=yearthis%>" method="post" name="opdater" id="opdater"><!-- target="w" -->
	
	<%
	'** Medarbejdere ***%>
	<div id="medarbejdere" style="position:absolute; left:-20px; top:15px; visibility:visible; display:; width:170px; height:695px; overflow:auto; padding:2px;"><!-- background-color:#eff3ff; border:1px #003399 solid; -->
	<table cellspacing=2 cellpadding=0 border=0 width="100%">
	<tr>
		<td width=16><img src="../ill/icon_planner_tp_24.gif" width="24" height="24" alt="" border="0"></td>
		<td><b>Medarbejdere</b></td>
	</tr>
	<!--<tr>
	<td colspan=2 valign=top>-->
	<%
	dim medarbid, medarbnavn
	redim medarbid(0), medarbnavn(0)
	m = 1
			strSQLM = "SELECT mnavn, mid, mnr FROM medarbejdere WHERE mansat <> 2 ORDER BY mnavn"
			oRec.open strSQLM, oConn, 3 
			while not oRec.EOF 
			
				if right(m, 1) = 0 then%>
				<!--</td><td valign=top>-->
				<%end if
				
							
					call colors(oRec("mid"))
					medcol = col_1
				
					
					thisMnr = oRec("mnr")
					'thisfornavn = left(oRec("mnavn"), 1)
					'navn_instr = instr(oRec("mnavn"), " ")
					'navn_len = len(oRec("mnavn"))
					'navn_right = right(oRec("mnavn"), (navn_len - (navn_instr)))
					'if cint(navn_instr) > 0 then
					'thisefternavn = left(navn_right, 1)
					'else
					'thisefternavn = ""
					'end if
					
					'initthis = thisfornavn & thisefternavn
					%>
				
				
				<!--<table height=10 cellspacing=1 cellpadding=0 border=0>-->
				<tr>
				<td width=10 height=8 bgcolor="<%=medcol%>" valign=top style="border:1px #000000 solid; padding:1px;">
				<font class=megetlillesort><%=thisMnr%></td>
				<td valign=top bgcolor="#eff3ff" style="padding-left:2; border:1px #ffffff solid;" id="selmedarbtd_<%=oRec("mid")%>" class=lille><%=left(oRec("mnavn"), 18)%> (<%=oRec("mnr")%>)
				<!--<a href="#" onClick="showmidbook(<=oRec("mid")%>)" class=vmenulille></a>-->
				</td>
				</tr><!--</table> -->
				<%
				redim preserve medarbid(m) 
				redim preserve medarbnavn(m)
				medarbid(m) = oRec("mid")
				medarbnavn(m) = oRec("mnavn")
				m = m + 1
							
							
			oRec.movenext
			wend
			oRec.close 
		
	%><!--</td>
	</tr>--></table>
	</div>
	<input type="hidden" name="antalmedarb" id="antalmedarb" value="<%=m-1%>">
	
	
	
	
	
	<%'** Uger ***%>
	<div id="ugeoversigt" name="ugeoversigt" style="position:absolute; left:600px; top:-25px; visibility:visible; display:; padding:2px;">
	
		<table cellspacing=2 cellpadding=0 border=0>
		<tr>
			
		<%
		antaluger = 1
		for d = 0 to mthDaysUse - 1
		
				
				if d = 0 then%>
				<td style="padding:2px; white-space:nowrap;"><b>Vælg uge:</b></td>
				<%end if
		
		ugeday = d + 1
		ugeweekdato = ugeday&"/"&monththis&"/"&yearthis
		weeknumberThis = datepart("ww", ugeweekdato, 2, 2)
		
		if lastweek <> weeknumberThis then
		if cint(weeknumberThis) = cint(activeuge) then
		bgcol= "#5582d2"
		else
		bgcol= "#ffffff"
		end if%>
		<td style="border:1px <%=bgcol%> solid; padding:2px;" bgcolor="#eff3ff"><a href="jbpla_w.asp?menu=res&activeuge=<%=weeknumberThis%>&mdselect=<%=monththis%>&aarselect=<%=yearthis%>"><%=weeknumberThis%></a></td>
		<%
		end if	
		lastweek = weeknumberThis
		next	
		%>
		</tr></table>
	</div>
		
							
		<%
		'*** Ugens 7 dage ***
		weeksinmonth = datepart("ww", ugeweekdato, 2, 2)%>
		<div id="klokke_<%=weeksinmonth%>" name="klokke_<%=weeksinmonth%>" style="position:absolute; z-index:<%=weeksinmonth%>; left:158px; top:21px; width:600px; visibility:<%=vz%>; display:<%=disp%>;">
		<table cellspacing=2 cellpadding=0 border=0 width="100%">
		<tr>
		<td style="width:38px" align=center><font class=megetlillesort>Uge: <%=activeuge%>
		<input type="hidden" name="aktuge" id="aktuge" value="<%=activeuge%>"> 
		</td>
					
		<%'*** Datoer på de viste ugedage til brug for DB ***
		for d = 0 to 6
				
		ugeday = d 
		ugeweekdato = dateadd("d", d, perStartKri)
		
				select case weekday(ugeweekdato, 2)
				case 1
				dayn = "M"
				bgdato = "#8caae6"
				ugeweekdato1 = ugeweekdato
				%>
				<input type="hidden" name="datoweekday1" id="datoweekday1" value="<%=ugeweekdato%>">
				<%
				case 2
				dayn = "T"
				bgdato = "#8caae6"
				ugeweekdato2 = ugeweekdato%>
				<input type="hidden" name="datoweekday2" id="datoweekday2" value="<%=ugeweekdato%>">
				<%
				case 3
				dayn = "O"
				bgdato = "#8caae6"
				ugeweekdato3 = ugeweekdato%>
				<input type="hidden" name="datoweekday3" id="datoweekday3" value="<%=ugeweekdato%>">
				<%
				case 4
				dayn = "T"
				bgdato = "#8caae6"
				ugeweekdato4 = ugeweekdato%>
				<input type="hidden" name="datoweekday4" id="datoweekday4" value="<%=ugeweekdato%>">
				<%
				case 5
				dayn = "F"
				bgdato = "#8caae6"
				ugeweekdato5 = ugeweekdato%>
				<input type="hidden" name="datoweekday5" id="datoweekday5" value="<%=ugeweekdato%>">
				<%
				case 6
				dayn = "L"
				bgdato = "#cccccc"
				ugeweekdato6 = ugeweekdato%>
				<input type="hidden" name="datoweekday6" id="datoweekday6" value="<%=ugeweekdato%>">
				<%
				case 7
				dayn = "S"
				bgdato = "#cccccc"
				ugeweekdato7 = ugeweekdato%>
				<input type="hidden" name="datoweekday7" id="datoweekday7" value="<%=ugeweekdato%>">
				<%							
				end select
				
				if cdate(daynow) = cdate(ugeweekdato) then
					bgdato = "#FFCC00"
				else
					bgdato = bgdato
				end if
						
				%>
				<td bgcolor="<%=bgdato%>" align="center" style="border:1px #FFFFFF solid; width:77px;"><font class=megetlillesort><%=dayn%>&nbsp;<%=left(ugeweekdato, 5)%></td>
				<%
				
		next
		%>
		</tr>
		</table>
		</div>	
		
	
	<%'** Grid **%>
	<div id="grid" name="grid" style="position:absolute; left:158px; top:38px; visibility:visible; display:; border:1px #003399 solid; background-color:#cccccc; padding:2px; z-index:1;">
	<map name="grid">
	<%
	r = 0
	x1 = 40
	x2 = 115
	felt = 0
	for r = 1 to 7
		select case r
		case 1
		imgmapdato = ugeweekdato1
		case 2
		imgmapdato = ugeweekdato2
		case 3
		imgmapdato = ugeweekdato3
		case 4
		imgmapdato = ugeweekdato4
		case 5
		imgmapdato = ugeweekdato5
		case 6
		imgmapdato = ugeweekdato6
		case 7
		imgmapdato = ugeweekdato7
		end select
	
	if r > 1 then
	x1 = x1 + 80
	x2 = x2 + 80
	end if
	
		c = 1
		y1 = -15
		y2 = 0
		imgmapsttid = "07:45:00"
		for c = 0 to 38
		imgmapsttid = dateadd("n", 15, imgmapsttid)
		y1 = y1 + 15
		y2 = y2 + 15
		felt = felt + 1%>
		
		
		<area alt="" coords="<%=x1%>,<%=y1%>,<%=x2%>,<%=y2%>" href="javascript:popUp('jbpla_w_opr.asp?func=opr&sttid=<%=imgmapsttid%>&dato=<%=imgmapdato%>&datostkri=<%=jobStartKri%>&datoslkri=<%=jobSlutKri%>&id=0','600','500','250','120');" target="_self">
		<input type="hidden" name="felt_<%=felt%>" id="felt_<%=felt%>" value="0" style="width:20px; font-size:8; background-color:#ffffe1;">
		<%next%>
	<%next%>
	</map>
	<img src="../ill/jbpla_grid.gif" width="600" height="556" alt="" border="0" usemap="#grid">
	</div>	
	
        	<div style="position:absolute; top:-25px; width:100px; left:800px; background-color:#ffffff; padding:5px; border:1px yellowgreen solid;"><a href="jbpla_w_opr.asp?func=opr&sttid=<%=imgmapsttid%>&dato=<%=imgmapdato%>&datostkri=<%=jobStartKri%>&datoslkri=<%=jobSlutKri%>&id=0" target="_blank">Opret ny +</a></div>
		
	<%	
	'*** Udksriver div felter i kalender *** 
	dim ugemid, ugedato, ugesttid, ugesluttid, ugejobid, ugemnavn, ugemnr, divid, bgcol, ugejobnavn, ugejobnr, ugeaid, ugeanavn, knavn, knr, kpers, kpers_email, kpers_dirtlf, kpers_mobiltlf, serieId 
	Redim ugemid(0), ugedato(0), ugesttid(0), ugesluttid(0), ugejobid(0), ugemnavn(0), ugemnr(0), divid(0), bgcol(0),ugejobnavn(0), ugejobnr(0), ugeaid(0), ugeanavn(0), knavn(0), knr(0), kpers(0), kpers_email(0), kpers_dirtlf(0), kpers_mobiltlf(0), serieId(0) 
	
	
	strSQL = "SELECT ressourcer.id AS divid, timer, starttp, sluttp, ressourcer.dato, ressourcer.mid AS mid, jobid, aktid, mnavn, "_
	&" mnr, jobnavn, jobnr, aktiviteter.id AS aid, aktiviteter.navn AS anavn, "_
	&" kkundenavn, kkundenr, kp.navn AS kpers, kp.email, kp.dirtlf, kp.mobiltlf, ressourcer.serieid FROM ressourcer "_
	&" LEFT JOIN job ON (job.id = jobid) "_
	&" LEFT JOIN medarbejdere ON (medarbejdere.mid = ressourcer.mid) "_
	&" LEFT JOIN aktiviteter ON (aktiviteter.id = aktid) "_
	&" LEFT JOIN kunder ON (kunder.kid = job.jobknr) "_
	&" LEFT JOIN kontaktpers kp ON (kp.id = job.kundekpers) "_
	&" WHERE ressourcer.dato BETWEEN '"& jobStartKri &"' AND '"& jobSlutKri &"' ORDER BY dato, starttp, sluttp DESC"
	
	'Response.write strSQL
	'Response.flush
	
	
	oRec.open strSQL, oConn, 3 
	u = 1
	lastmid = 0
	while not oRec.EOF 
	
	Redim preserve ugemid(u)
	Redim preserve ugedato(u)
	Redim preserve ugesttid(u)
	Redim preserve ugesluttid(u)
	Redim preserve ugejobid(u)
	Redim preserve ugemnavn(u)
	Redim preserve ugemnr(u)
	Redim preserve divid(u)
	Redim preserve bgcol(u)
	Redim preserve ugejobnavn(u)
	Redim preserve ugejobnr(u)
	Redim preserve ugeaid(u)
	Redim preserve ugeanavn(u)
	Redim preserve knavn(u), knr(u) 
	Redim preserve kpers(u), kpers_email(u), kpers_dirtlf(u), kpers_mobiltlf(u), serieId(u) 
	ugemid(u) = oRec("mid")
	ugedato(u) = oRec("dato")
	ugesttid(u) = oRec("starttp")
	ugesluttid(u) = oRec("sluttp")
	ugejobid(u) = oRec("jobid")
	ugejobnavn(u) = oRec("jobnavn")
	ugejobnr(u) = oRec("jobnr")
	ugemnavn(u) = oRec("mnavn")
	ugemnr(u) = oRec("mnr")
	divid(u) = oRec("divid")
	ugeaid(u) = oRec("aid")
	ugeanavn(u) = oRec("anavn")
	knavn(u) = oRec("kkundenavn")
	knr(u) = oRec("kkundenr")
	kpers(u) = oRec("kpers")
	kpers_email(u) = oRec("email")
	kpers_dirtlf(u) = oRec("dirtlf")
	kpers_mobiltlf(u) = oRec("mobiltlf")
	
	serieId(u) = oRec("serieid")
	
	'** Hvis dette er basis bookningen, tjekkes det om denne er en del af en serie her.
	erDenneBasisEnSerie = 0
	if oRec("serieid") = 0 then
		strSQL6 = "SELECT serieid FROM ressourcer WHERE serieid = "& oRec("divid")
		oRec3.open strSQL6, oConn, 3 
		if not oRec3.EOF then 
		erDenneBasisEnSerie = 1
		end if
		oRec3.close 
	end if
	
	if erDenneBasisEnSerie = 1 then
	serieId(u) = oRec("divid")
	else
	serieId(u) = serieId(u)
	end if
	'****
						
	call colors(ugemid(u))
	bgcol(u) = col_1
						
							
	
	%>
	<input type="hidden" name="u_this_<%=u%>" id="u_this_<%=u%>" value="<%=oRec("mid")%>">
	<%
	
	u = u + 1
	oRec.movenext
	wend
	oRec.close 
	
	
	
	
	for u = 1 to u - 1 
			
		dtop = datediff("n", "08:00:00 AM", formatdatetime(ugesttid(u), 3))
		dheight = datediff("n", ugesttid(u), ugesluttid(u))	
					
					
					'**** lægger div på den rigtige dag ***
					select case weekday(ugedato(u), 2)
					case 1
					leftc = 202 '53
					
					case 2
					leftc = 282 '98
					
					case 3
					leftc = 362
					
					case 4
					leftc = 442
					
					case 5
					leftc = 522
					
					case 6
					leftc = 602
					
					case 7
					leftc = 682
					
					end select
					
					
					
					'***** Fylder felter op således at ingen div ligger over hinanden ****
						select case leftc
						case 202
						feltbase = 1
						call dtosub
						feltnrorg = feltbase
						call dhesub
							
						case 282
						feltbase = 39
						call dtosub
						feltnrorg = feltbase
						call dhesub
							
						case 362
						feltbase = 77
						call dtosub
						feltnrorg = feltbase
						call dhesub
							
								
						case 442
						feltbase = 115
						call dtosub
						feltnrorg = feltbase
						call dhesub
						
						case 522
						feltbase = 153
						call dtosub
						feltnrorg = feltbase
						call dhesub
							
							
						case 602
						feltbase = 191
						call dtosub
						feltnrorg = feltbase
						call dhesub
							
						case 682
						feltbase = 229
						call dtosub
						feltnrorg = feltbase
						call dhesub
							
						end select
					
					
					
					
					
					
					
					
					'**** Personlige redigerbare div'er ****%>
					<!--<div name="bookdiv_<=ugemid(u)%>_<=u%>" id="bookdiv_<=ugemid(u)%>_<=u%>" style="position:absolute; display:none; visibility:hidden; left:<=leftc+antalux%>; top:<=dtop+162%>; height:<=dheight%>; background-color:<=bgcol(u)%>; width:35; z-index:4000; border:1px #000000 solid; padding:2px;" onMousedown="bookdivon()" onMousemove="bookdivmove()" onMouseup="bookdivoff('<=ugemid(u)%>','<=u%>')">
					<table cellspacing=0 cellpadding=0 border=0><tr>
					<td valign=top><a href="#" onClick="bookminus('<=ugemid(u)%>','<=u%>')"><img src="../ill/pilop2.gif" width="11" height="10" alt="Minimer 15 min." border="0"></a></td>
					<td style="padding-left:1px;"><a href="#" onClick="showinfo(<=u%>)"><img src="../ill/info2.gif" width="11" height="10" alt="Info" border="0"></a></td>
					<td valign=top style="padding-left:1px;"><a href="#" onClick="bookplus('<=ugemid(u)%>','<=u%>')"><img src="../ill/pilned2.gif" width="11" height="10" alt="Tilføj 15 min." border="0"></a></td>
					</tr>
					</table>
					</div>-->
					
					<input type="hidden" id="bookdiv_x_<%=ugemid(u)%>_<%=u%>" name="bookdiv_x_<%=ugemid(u)%>_<%=u%>" value="<%=(leftc)%>">
					<input type="hidden" id="bookdiv_y_<%=ugemid(u)%>_<%=u%>" name="bookdiv_y_<%=ugemid(u)%>_<%=u%>" value="<%=(dtop+42)%>"><!-- 162 -->
					<input type="hidden" id="bookdiv_id_<%=ugemid(u)%>_<%=u%>" name="bookdiv_id_<%=ugemid(u)%>_<%=u%>" value="<%=divid(u)%>">
					<input type="hidden" id="bookdiv_height_<%=ugemid(u)%>_<%=u%>" name="bookdiv_height_<%=ugemid(u)%>_<%=u%>" value="<%=dheight%>">
					<input type="hidden" id="bookdiv_feltbasenr_<%=ugemid(u)%>_<%=u%>" name="bookdiv_feltbasenr_<%=ugemid(u)%>_<%=u%>" value="<%=feltnrorg%>">
					
					
					<div name="info_thumb_<%=ugemid(u)%>_<%=u%>" id="info_thumb_<%=ugemid(u)%>_<%=u%>" style="position:absolute; top:<%=dtop+34%>; z-index:<%=1000+u+1%>;">
					<a href="#" onClick="showinfo('<%=u%>')"><img src="../ill/info_ikon.gif" width="11" height="14" alt="" border="0"></a>
					</div>
					
					<!--<div name="info_thumb2_<=ugemid(u)%>_<=u%>" id="info_thumb2_<=ugemid(u)%>_<=u%>" style="position:absolute; top:<=dtop+32%>; z-index:<=1000+u+1%>; height:10px; background-color:#ffffff; border:1px darkred solid; width:30px; padding-left:1px;">
					<a href="#" onClick="showinfo('<=u%>')" class=klog_lille2><=left(knavn(u),5)%></a>
					</div>-->
					
					<script>
					var faktorfelt = 0
					var faktorThis = 0
					faktorleftc = <%=leftc%>
					faktorfelt = faktorleftc + parseInt((document.getElementById("felt_"+<%=feltnrorg%>).value - 1) * 8)
					faktorThis = faktorfelt-faktorleftc
					document.write("<div name='bookdiv_<%=ugemid(u)%>_<%=u%>' id='bookdiv_<%=ugemid(u)%>_<%=u%>' style='position:absolute; display:<%=dspthis%>; visibility:<%=vzthis%>; top:<%=dtop+42%>; height:<%=dheight%>; width:12px; z-index:<%=1000+u%>; border:1px #000000 solid; background-color:<%=bgcol(u)%>; padding:0;' onMousedown=bookdivon() onMousemove=bookdivmove() onMouseup=bookdivoff('<%=ugemid(u)%>','<%=u%>')>") 
					document.getElementById("bookdiv_<%=ugemid(u)%>_<%=u%>").style.left = faktorfelt
					document.getElementById("info_thumb_<%=ugemid(u)%>_<%=u%>").style.left = faktorfelt + 2
					//document.getElementById("info_thumb2_<%=ugemid(u)%>_<%=u%>").style.left = faktorfelt + 15
					document.write("<input type=hidden id='faktor_<%=ugemid(u)%>_<%=u%>' name='faktor_<%=ugemid(u)%>_<%=u%>' value='"+ faktorThis +"'>")
					</script>
					<%
					'***** Oversigts div'er *****
					dspthis = ""
					vzthis = "visible"
					
					'thisfornavn = left(ugemnavn(u), 1)
					'navn_instr = instr(ugemnavn(u), " ")
					'navn_len = len(ugemnavn(u))
					'navn_right = right(ugemnavn(u), (navn_len - (navn_instr)))
					'if cint(navn_instr) > 0 then
					'thisefternavn = left(navn_right, 1)
					'else
					'thisefternavn = ""
					'end if
					'initthis = thisfornavn & thisefternavn
					
					thisMnr = ugemnr(u)
					Response.write "<font class=megetlillesort>"& thisMnr &"</font>"
					%>
					</div>
					
					<!--=(leftc)+10%>-->
					<div name="info_<%=u%>" id="info_<%=u%>" style="position:absolute; display:none; visibility:hidden; left:320; top:<%=(dtop+antalx+10) - 100%>; height:100; width:160; z-index:8000; border:1px darkred solid; background-color:#fffff1; padding:2px;">
					<table cellspacing=0 cellpadding=1 border=0 width=200>
					<tr>
						<td valign=top colspan=2><b><%=ugemnavn(u)%>  (<%=ugemnr(u)%>)</b></td>
					</tr>
					<tr>
						<td valign=top class="lille">Dato:</td><td valign=top class="lille"><%=ugedato(u)%></td>
					</tr>
					<tr>
						<td valign=top class="lille">Start:</td><td valign=top class="lille"><%=formatdatetime(ugesttid(u), 3)%></td>
					</tr>
					<tr>
						<td valign=top class="lille">Slut:</td><td valign=top class="lille"> <%=formatdatetime(ugesluttid(u), 3)%></td>
					</tr>
					<tr>
						<td valign=top class="lille">Job:</td><td valign=top class="lille">
						<%if len(ugejobnavn(u)) > 22 then
						Response.write left(ugejobnavn(u), 22) & ".."
						else
						Response.write ugejobnavn(u)
						end if%>
						(<%=ugejobnr(u)%>)</td>
					</tr>
					<!--<tr>	
						<td valign=top class="lille">Aktivitet:</td><td valign=top class="lille"><=ugeanavn(u)%></td>
					</tr>-->
					<tr>
						<td valign=top class="lille"><br>Kunde:</td><td valign=top><br>
						<b><%if len(knavn(u)) > 22 then
						Response.write left(knavn(u), 22) & ".."
						else
						Response.write knavn(u)
						end if%>
						(<%=knr(u)%>)</b></td>
					</tr>
					<tr>
						<td valign=top class="lille" colspan=2>Kontaktperson:<br><a href="mailto:<%=kpers_email(u)%>" class=vmenu> <%=kpers(u)%></a><br>
						Dir. Tlf: <%=kpers_dirtlf(u)%><br>
						Mobil Tlf: <%=kpers_mobiltlf(u)%> </td>
					</tr>
					<tr bgcolor="#ffffff">	
						<td colspan=2 style="border:1px #000000 solid;">&nbsp;<a href="#" onClick="hideinfo(<%=u%>)" class=vmenu>Luk</a>&nbsp;|&nbsp;
						<a href="jbpla_w.asp?func=slet&id=<%=divid(u)%>&activeuge=<%=activeuge%>&yearthis=<%=yearthis%>&monththis=<%=monththis%>" class=red>Slet</a>
						&nbsp;|&nbsp;
						<a href="javascript:popUp('jbpla_w_opr.asp?func=red&sttid=<%=formatdatetime(ugesttid(u), 3)%>&sltid=<%=formatdatetime(ugesluttid(u), 3)%>&dato=<%=ugedato(u)%>&datostkri=<%=jobStartKri%>&datoslkri=<%=jobSlutKri%>&jobid=<%=ugejobid(u)%>&medid=<%=ugemid(u)%>&aid=<%=ugeaid(u)%>&id=<%=divid(u)%>&step=1&serieid=<%=serieId(u)%>&stepialt=2','600','500','250','120');" target="_self" onClick="hideinfo(<%=u%>)"; class=vmenu>Rediger</a>
						<%if serieId(u) <> 0 then%>&nbsp;|
						<a href="jbpla_w.asp?func=sletserie&id=<%=divid(u)%>&activeuge=<%=activeuge%>&yearthis=<%=yearthis%>&monththis=<%=monththis%>&serieid=<%=serieId(u)%>" class=red>Slet Serie</a>
						<%end if%>
						</td>
					</tr>
					</table>
					</div>
					<%
			
			
	next
	%>
	<input type="hidden" name="antalu" id="antalu" value="<%=u-1%>">
	<input type="hidden" name="showlastmid" id="showlastmid" value="0">
	<div name="redsubmit" id="redsubmit" style="position:absolute; display:none; visibility:visible; left:280; top:10;">
	<input type="image" src="../ill/opdaterugepil.gif">
	</div>
	</form>
	
	

	<%
	end select 'func

end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
