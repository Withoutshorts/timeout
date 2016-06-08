<%
Public function whatDay_crm(dayType, preDayType, preValue, dagsdato, startdag, clicked)
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
		
		
		bgcolor = "#ffffff"
		
		%>
		<tr height="25">
		<td valign=middle align=center>&nbsp;
		<%
		z = 1
		for z = 1 to 7
		
			if day(now)&"/"&month(now)&"/"&year(now) =  tjekTodag(z)&"/"&m&"/"&y then
			strShowDayValue = "<font class='lille-kalender'><font color=red>"& tjekTodag(z) &"</font>"
			else
			strShowDayValue = tjekTodag(z)
			end if
			
			'** Henter crmaktioner ***
			strSQL3 = "SELECT DISTINCT(crmhistorik.id), kundeid, kkundenavn, crmdato, crmhistorik.navn, kontaktemne, crmemne.navn AS emnenavn FROM crmhistorik, crmemne, kunder, aktionsrelationer WHERE crmdato = '"& y&"/"&m&"/"&tjekTodag(z) &"' AND crmemne.id = kontaktemne AND kid = kundeid "& usemedarbKri &" ORDER BY kid, kontaktemne"
			oRec3.open strSQL3, oConn, 3
			strAktnavn = "Aktioner idag:" &vbcrlf
			akt_today = "n"
			
				while not oRec3.EOF
				strAktnavn = strAktnavn & oRec3("kkundenavn") & chr(032) & oRec3("emnenavn") & chr(032) & oRec3("navn") &vbcrlf
				akt_today = "y"
				oRec3.movenext
				wend
			oRec3.close
			
			if akt_today = "y" then
			weight = "bold"
			pointer = "pointer"
			else
			weight = "normal"
			pointer = "default"
			strAktnavn = ""
			end if
		%>
		<td BGCOLOR=<%=bgcolor%> valign=top align=center><a href="crmkalender.asp?selpkt=kal&menu=crm&dato=<%=y&"/"&m&"/"&tjekTodag(z)%>&status=<%=status%>&id=<%=id%>&emner=<%=emner%>&medarb=<%=medarb%>&regaar=<%=y%>&m=<%=m%>&usrsel=d&m_forrige=<%=m_forrige%>&m_naeste=<%=m_naeste%>&y=<%=y%>"><font class='lille-kalender' color="#000000" title="<%=strAktnavn%>" style="font-weight : <%=weight%>; cursor:<%=pointer%>;"><%=strShowDayValue%></font></a>
		<br><span style="background-color:#D6DFF5; padding-left:2; padding-right:2;"><font class='lille-kalender'></font></span>
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
			
			'** Henter crmaktioner ***
			strSQL3 = "SELECT DISTINCT(crmhistorik.id), kundeid, Kkundenavn, crmdato, crmhistorik.navn, kontaktemne, crmemne.navn AS emnenavn FROM crmhistorik, crmemne, kunder, aktionsrelationer WHERE crmdato = '"& y&"/"&m&"/"&strDayValue &"' AND crmemne.id = kontaktemne AND Kid = kundeid "& useidKri &""& useemneKri &""& usestatKri &""& usemedarbKri &" ORDER BY kid, kontaktemne"
			oRec3.open strSQL3, oConn, 3
			strAktnavn = "Aktioner idag:" &vbcrlf
			akt_today = "n"
			
				while not oRec3.EOF
				strAktnavn = strAktnavn & oRec3("Kkundenavn") & chr(032) & oRec3("emnenavn") & chr(032) & oRec3("navn") &vbcrlf
				akt_today = "y"
				oRec3.movenext
				wend
			oRec3.close
			
			if akt_today = "y" then
			weight = "bold"
			pointer = "pointer"
			else
			weight = "normal"
			pointer = "default"
			strAktnavn = ""
			end if
			
			bgcolor = "#ffffff"
				
			select case left(dayType, 3)
			case "lor"
			Response.write "<td valign=top  align=center BGCOLOR="&bgcolor&">"%>
			<a href="crmkalender.asp?selpkt=kal&menu=crm&dato=<%=y&"/"&m&"/"&strDayValue%>&status=<%=status%>&id=<%=id%>&emner=<%=emner%>&medarb=<%=medarb%>&regaar=<%=y%>&m=<%=m%>&usrsel=d&m_forrige=<%=m_forrige%>&m_naeste=<%=m_naeste%>&y=<%=y%>"><font color="#000000" title="<%=strAktnavn%>" class='lille-kalender' style="font-weight : <%=weight%>; cursor:<%=pointer%>;"><%=strShowDayValue%></font></a>
			<br><span style="background-color:D6DFF5; padding-left:2; padding-right:2;"><font size=3 color=#003399></font></span></td>
			<%
			clicked = clicked + 1
			case "son"
			'** Viser knapper ***
			Response.write "<tr height=25><td>&nbsp;</td>"
			Response.write "<td valign=top align=center BGCOLOR="&bgcolor&">"%>
			<a href="crmkalender.asp?selpkt=kal&menu=crm&dato=<%=y&"/"&m&"/"&strDayValue%>&status=<%=status%>&id=<%=id%>&emner=<%=emner%>&medarb=<%=medarb%>&regaar=<%=y%>&m=<%=m%>&usrsel=d&m_forrige=<%=m_forrige%>&m_naeste=<%=m_naeste%>&y=<%=y%>"><font color="#000000" title="<%=strAktnavn%>" class='lille-kalender' style="font-weight : <%=weight%>; cursor:<%=pointer%>;"><%=strShowDayValue%></font></a>
			<br><span style="background-color:D6DFF5; padding-left:2; padding-right:2;"><font size=1 color=#003399></font></span></td>
			<%
			case else
			Response.write "<td valign=top align=center BGCOLOR="&bgcolor&">"%>
			<a href="crmkalender.asp?selpkt=kal&menu=crm&dato=<%=y&"/"&m&"/"&strDayValue%>&status=<%=status%>&id=<%=id%>&emner=<%=emner%>&medarb=<%=medarb%>&regaar=<%=y%>&m=<%=m%>&usrsel=d&m_forrige=<%=m_forrige%>&m_naeste=<%=m_naeste%>&y=<%=y%>"><font color="#000000" title="<%=strAktnavn%>" class='lille-kalender' style="font-weight : <%=weight%>; cursor:<%=pointer%>;"><%=strShowDayValue%></font></a>
			<br><span style="background-color:D6DFF5; padding-left:2; padding-right:2;"><font size=1 color=#003399></font></span>
			</td>
			<%
			end select
			preValue = strDayValue
		end if
		end function
%>
