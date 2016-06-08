<%
Public function whatDay(dayType, preDayType, preValue, dagsdato, startdag, clicked, usemnr)
		if startdag <> 0 then
		select case startdag
		case 1
		son1 = 1
		man1 = 2
		tir1 = 3
		ons1 = 4
		tor1 = 5
		fre1 = 6
		lor1 = 7
		countShowedDays = 7
		case 2
		son1 = 0
		man1 = 1
		tir1 = 2
		ons1 = 3
		tor1 = 4
		fre1 = 5
		lor1 = 6
		countShowedDays = 6
		case 3
		son1 = 0
		man1 = 0
		tir1 = 1
		ons1 = 2
		tor1 = 3
		fre1 = 4
		lor1 = 5
		countShowedDays = 5
		case 4
		son1 = 0
		man1 = 0
		tir1 = 0
		ons1 = 1
		tor1 = 2
		fre1 = 3
		lor1 = 4
		countShowedDays = 4
		case 5
		son1 = 0
		man1 = 0
		tir1 = 0
		ons1 = 0
		tor1 = 1
		fre1 = 2
		lor1 = 3
		countShowedDays = 3
		case 6
		son1 = 0
		man1 = 0
		tir1 = 0
		ons1 = 0
		tor1 = 0
		fre1 = 1
		lor1 = 2
		countShowedDays = 2
		case 7
		son1 = 0
		man1 = 0
		tir1 = 0
		ons1 = 0
		tor1 = 0
		fre1 = 0
		lor1 = 1
		countShowedDays = 1
		end select
		
		countDay = 0
		preValue = lor1
		
		
		z = 1
		stDayType = 0
		Redim tjekTodag(7)
		tjekTodag(1) = son1
		tjekTodag(2) = man1
		tjekTodag(3) = tir1
		tjekTodag(4) = ons1
		tjekTodag(5) = tor1
		tjekTodag(6) = fre1
		tjekTodag(7) = lor1
		
		for z = 1 to 7
		
		if tjekTodag(z) <> 0 then
		tjekTodag(z) = tjekTodag(z)
		stDayType = stDayType + 1
	    else
		tjekTodag(z) = "&nbsp;"
		end if
		next
		
		
		select case stDayType
		case "7"
		stDayType = "son1"
		case "6"
		stDayType = "man1"
		case "5"
		stDayType = "tir1"
		case "4"
		stDayType = "ons1"
		case "3"
		stDayType = "tor1"
		case "2"
		stDayType = "fre1"
		case "1"
		stDayType = "lor1"
		end select
		
		Select case Request("lastClicked")
		case "0"
			if andre = "show" then
			bgcolor = "#ffffff"
			else
			bgcolor = "#ffffe1"
			end if
		case ""
		bgcolor = "#ffffff"
		case else
		bgcolor = "#ffffff"
		end select
		
		
		
		
		%>
		<tr height="25">
		<td valign=middle align=center>
		<%
		if andre <> "show" then
		Response.write "<a href='timereg.asp?usrsel=d&lastClicked=0&menu=timereg&regdag=1&regdtype="&stDayType&"&regmrd="&m&"&regaar="&y&"&m="&m&"&m_forrige="&m_forrige&"&m_naeste="&m_naeste&"&y="&y&"&sort="&request("sort")&"&searchstring="&searchstring&"&FM_use_me="&usemrn&"'><img src='../ill/pil_kalender.gif' alt='Gå til denne uge' border='0' style='margin-left: 2px;'></a></td>"
		else
		Response.write "&nbsp;"
		end if

		
		z = 1
		for z = 1 to 7
		
			if day(now)&"/"&month(now)&"/"&year(now) =  tjekTodag(z)&"/"&m&"/"&y then
				strShowDayValue = "<font class='lille-kalender'><font color=red>"& tjekTodag(z) &"</font>"
				else
				strShowDayValue = tjekTodag(z)
				end if
			
			'** Henter job ***
			'strSQL3 = "SELECT jobnavn, jobnr FROM job WHERE jobslutdato = '"& y&"/"&m&"/"&tjekTodag(z) &"' ORDER BY jobnavn"
			'oRec3.open strSQL3, oConn, 0, 1
			'strAktnavn = "Job deadline idag:" &vbcrlf
			'akt_today = "n"
			
				'while not oRec3.EOF
				'strAktnavn = strAktnavn & oRec3("jobnr") & chr(032) & oRec3("jobnavn") &vbcrlf
				'akt_today = "y"
				'oRec3.movenext
				'wend
			'oRec3.close
			
			'strAktnavn = strAktnavn &vbcrlf
			
			'** Henter aktiviteter ***
			'strSQL3 = "SELECT navn, aktslutdato, jobnavn, jobnr FROM aktiviteter, job WHERE aktslutdato= '"& y&"/"&m&"/"&tjekTodag(z) &"' AND job.id = aktiviteter.job OR jobslutdato = '"& y&"/"&m&"/"&strDayValue &"' AND aktiviteter.job = job.id ORDER BY jobnavn"
			'oRec3.open strSQL3, oConn, 0, 1
			'strAktnavn = "Aktiviteter deadline idag:" &vbcrlf
			
				'while not oRec3.EOF
				'strAktnavn = strAktnavn & oRec3("jobnr") & chr(032) & oRec3("jobnavn") & chr(032) & oRec3("navn") &vbcrlf
				'akt_today = "y"
				'oRec3.movenext
				'wend
			'oRec3.close
			
			'if akt_today = "y" then
			'weight = "bold"
			'pointer = "default"
			'else
			weight = "normal"
			pointer = "default"
			strAktnavn = ""
			'end if
			
		if tjekTodag(z) <> "&nbsp;" then
		strSQL = "SELECT sum(timer) AS timer_indtastet FROM timer WHERE Tmnr = "& varMed_id &" AND Tdato='"&y&"/"&m&"/"&tjekTodag(z)&"' AND tfaktim <> 5"
		
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		intDayHoursVal = oRec("timer_indtastet")
		end if
		oRec.close
		else
		intDayHoursVal = 0
		end if
		%>
		<td BGCOLOR=<%=bgcolor%> valign=top align=center><font color="#000000" class='lille-kalender' title="<%=strAktnavn%>" style="font-weight : <%=weight%>; cursor:<%=pointer%>;"><%=strShowDayValue%></font>
		<br>
		<%if intDayHoursVal <> 0 then%>
		<span style="background-color:#D6DFF5; padding-left:2; padding-right:2;"><font class='lille-kalender' color=#003399><%=intDayHoursVal%></font></span>
		<%end if
		%>
		</td>
		<%
		next
		%>
		</tr>

		<%
		startdag = 0
		else
			strDayValue = (preValue + 1)
			
			if day(now)&"/"&month(now)&"/"&year(now) =  strDayValue&"/"&m&"/"&y then
			strShowDayValue = "<font class='lille-kalender'><font color=red>"& strDayValue &"</font>" 
			else
			strShowDayValue = strDayValue
			end if
			
			'** Henter job ***
			'strSQL3 = "SELECT jobnavn, jobnr FROM job WHERE jobslutdato = '"& y&"/"&m&"/"&strDayValue &"' ORDER BY jobnavn"
			'oRec3.open strSQL3, oConn, 0, 1
			'strAktnavn = "Job deadline idag:" &vbcrlf
			'akt_today = "n"
			
				'while not oRec3.EOF
				'strAktnavn = strAktnavn & oRec3("jobnr") & chr(032) & oRec3("jobnavn") &vbcrlf
				'akt_today = "y"
				'oRec3.movenext
				'wend
			'oRec3.close
			
			'strAktnavn = strAktnavn &vbcrlf
			
			'** Henter aktiviter ***
			'strSQL3 = "SELECT aktiviteter.navn, aktslutdato, jobnavn, jobnr FROM aktiviteter, job WHERE aktslutdato= '"& y&"/"&m&"/"&strDayValue &"' AND job.id = aktiviteter.job ORDER BY navn"
			'oRec3.open strSQL3, oConn, 0, 1
			'strAktnavn = strAktnavn & "Aktiviterer deadline idag:" &vbcrlf
			
				'while not oRec3.EOF
				'strAktnavn = strAktnavn & oRec3("jobnr") & chr(032) & oRec3("jobnavn") & chr(032) & oRec3("navn") &vbcrlf
				'akt_today = "y"
				'oRec3.movenext
				'wend
			'oRec3.close
			
			
			'if akt_today = "y" then
			'weight = "bold"
			'pointer = "default"
			'else
			weight = "normal"
			pointer = "default"
			strAktnavn = ""
			'end if
			
			
			if cint(Request("lastClicked")) > 0 then
				if cint(Request("lastClicked")) = cint(Clicked) then
				bgcolor = "#ffffe1"
				else
				bgcolor = "#ffffff"
				end if
			end if
			
			strSQL = "SELECT sum(timer) AS timer_indtastet FROM timer WHERE Tmnr = "& varMed_id &" AND Tdato='"&y&"/"&m&"/"&strDayValue&"' AND tfaktim <> 5"
			
			oRec.open strSQL, oConn, 1
			if not oRec.EOF then
			intDayHoursVal = oRec("timer_indtastet")
			end if
			oRec.close
			
			if len(intDayHoursVal) = 0 then
			intDayHoursVal = 0
			end if
			
			
			select case left(dayType, 3)
			case "lor"
			Response.write "<td valign=top  align=center BGCOLOR="&bgcolor&">"%>
			<font color="#000000" title="<%=strAktnavn%>" class='lille-kalender' style="font-weight : <%=weight%>; cursor:<%=pointer%>;"><%=strShowDayValue%></font>
			<br><span style="background-color:D6DFF5; padding-left:2; padding-right:2;"><font class='lille-kalender' color=#003399><%=intDayHoursVal%></font></span></td>
			<%
			clicked = clicked + 1
			case "son"
			'** Viser knapper ***
			if andre <> "show" then
				Response.write "<tr><td colspan='8' bgcolor='#D6DFF5'><img src='../ill/blank.gif' alt='' border='0'></td></tr>"
				Response.write "<tr height='25'><td valign=middle align=center><a href='timereg.asp?usrsel=d&lastClicked="&clicked&"&menu=timereg&regdag="&strDayValue&"&regdtype="&dayType&"&regmrd="&m&"&regaar="&y&"&m="&m&"&m_forrige="&m_forrige&"&m_naeste="&m_naeste&"&y="&y&"&sort="&request("sort")&"&searchstring="&searchstring&"&FM_use_me="&usemrn&"'><img src='../ill/pil_kalender.gif' alt='Gå til denne uge' border='0' style='margin-left: 2px;'></a></td>"
			else
				Response.write "<tr height=25><td>&nbsp;</td>"
			end if
			Response.write "<td valign=top align=center BGCOLOR="&bgcolor&">"
			%>
			
			<font color="#000000" title="<%=strAktnavn%>" class='lille-kalender' style="font-weight : <%=weight%>; cursor:<%=pointer%>;"><%=strShowDayValue%></font>
			<br><span style="background-color:D6DFF5; padding-left:2; padding-right:2;"><font class='lille-kalender' color=#003399><%=intDayHoursVal%></font></span></td>
			<%
			case else
			Response.write "<td valign=top align=center BGCOLOR="&bgcolor&">"%>
			<font color="#000000" title="<%=strAktnavn%>" class='lille-kalender' style="font-weight : <%=weight%>; cursor:<%=pointer%>;"><%=strShowDayValue%></font>
			<br><span style="background-color:D6DFF5; padding-left:2; padding-right:2;"><font class='lille-kalender' color=#003399><%=intDayHoursVal%></font></span>
			</td>
			<%
			end select
			preValue = strDayValue
		end if
		end function
%>
