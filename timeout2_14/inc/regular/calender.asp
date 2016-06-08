<%
'*** Se timer for andre medarb ****
if request("FM_andre") = "show" then%>
<!--#include file="../connection/conn_db_inc.asp"-->
<LINK rel="stylesheet" type="text/css" href="../style/timeout_style.css">
<%
varMed_id = request("FM_medarb_id")
leftDiv = 30
topDiv = 77
Response.write "<br>Timeregistreringer for:<br><b>" & request("FM_Mnavn") & "</b><br>"
else
'*** fjernet 6/5-2004 ***%>
<!--include file="../connection/conn_db_inc.asp"-->
<!--include file="header_inc.asp"-->
<%
varMed_id = usemrn
leftDiv = 740
topDiv = 127
end if


'******************************************* dato funktioner ***********************************
		
		'*** Sætter lokal dato/kr format. *****
		Session.LCID = 1030
		
		If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
			if len(request("FM_start_dag")) <> 0 then 'Fra datovælger form
			strMrd = request("FM_start_mrd")
				select case strMrd
				case 2
					if request("FM_start_dag") > 28 then
					strDag = 28
					else
					strDag = request("FM_start_dag")
					end if
				case 1, 3, 5, 7, 8, 10, 12
					strDag = request("FM_start_dag")
				case else
					if request("FM_start_dag") > 30 then
					strDag = 30
					else
					strDag = request("FM_start_dag")
					end if
				end select
			strAar = request("FM_start_aar")
			else
			strDag = day(now)
			strMrd = month(now)
			strAar = year(now)
			end if
		else
			if len(request("strdag")) <> 0 then 'Fra dato link
			strMrd = request("strmrd")
				select case strMrd
				case 2
					if request("strdag") > 28 then
					strDag = 28
					else
					strDag = request("strdag")
					end if
				case 1, 3, 5, 7, 8, 10, 12
					strDag = request("strdag")
				case else
					if request("strdag") > 30 then
					strDag = 30
					else
					strDag = request("strdag")
					end if
				end select
			strAar = request("straar")
			else
			strDag = day(now)
			strMrd = month(now)
			strAar = year(now)
			end if
		end if
		
		daynow = formatdatetime(day(now) & "/" & month(now) & "/" & year(now), 0)
		useDate = formatdatetime(strDag & "/" & strMrd & "/" & strAar, 0)
		firstDayOfMonth = formatdatetime(1 & "/" & strMrd & "/" & strAar, 0)
		
		Select case strMrd
		case 1, 3, 5, 7, 8, 10, 12
		lastDayOfMonth = formatdatetime("31/" & strMrd & "/" & strAar, 0)
		numberofdaysinmonth = 31
		case 2
			select case strAar
			case 2004, 2008, 2012, 2016, 2020
			lastDayOfMonth = formatdatetime("29/" & strMrd & "/" & strAar, 0)
			numberofdaysinmonth = 29
			case else
			lastDayOfMonth = formatdatetime("28/" & strMrd & "/" & strAar, 0)
			numberofdaysinmonth = 28
			end select
		case else
		lastDayOfMonth = formatdatetime("30/" & strMrd & "/" & strAar, 0)
		numberofdaysinmonth = 30
		end select
		
		firstWeekday = Weekday(firstDayOfMonth, 2) 
		lastWeekday = Weekday(lastDayOfMonth, 2) 
		
		prevMonth = datePart("m", DateAdd("m",-1, useDate))
		nextMonth = datePart("m", DateAdd("m",1, useDate))
		
		thisMonthName = monthname(strMrd)
		prevMonthName = left(monthname(prevMonth), 3)
		nextMonthName = left(monthname(nextMonth), 3)
		
		select case prevMonth 
		case 12
		prevYear = datePart("yyyy", DateAdd("yyyy",-1, useDate))
		case else
		prevYear = strAar
		end select
		
		select case nextMonth 
		case 1
		nextYear = datePart("yyyy", DateAdd("yyyy",1, useDate))
		case else
		nextYear = strAar
		end select
		
		countDay = 1
		andre = request("FM_andre")
		
		if menu = "crm" then
		seldocument = "crmkalender"
		kalenderlink = "shownumofdays="&shownumofdays&"&ketype=e&selpkt=kal&status="&status&"&id="&id&"&emner="&emner&"&medarb="&medarb&"&sort="&request("sort")&""
		showother = "notused"
		else
			if andre = "show" then 
			seldocument = "calender"
			showother = "show"
			kalenderlink = "FM_medarb_id="&request("FM_medarb_id")&"&FM_Mnavn="&request("FM_Mnavn")&"&searchstring="&searchstring&""
			else
			seldocument = "timereg"
			showother = "dontshow"
			kalenderlink = "searchstring="&searchstring&""
			end if
		'kalenderlink = ""
		end if
		
	if andre = "show" then 
	illpath = "../../"
	else
	illpath = "../"
	end if

	'*****************************************************************************************
%>
<div id="rmenu1" style="position:absolute; left:<%=leftDiv%>; top:<%=topDiv%>; visibility:visible; z-index=1000;">
<%
kalenderMnavn = ""

if request("FM_andre") <> "show" AND (thisfile = "timereg" OR thisfile = "timereg_db") then
strSQL = "SELECT mnavn, mnr FROM medarbejdere WHERE mid = "& usemrn 'medarbedjer id
oRec.open strSQL, oConn, 3 
if not oRec.EOF then 
kalenderMnavn = "<font class=megetlillehvid>"&oRec("mnavn")&"&nbsp;("&oRec("mnr")&")</font>"
end if
oRec.close 
end if%>

<table cellpadding="0" cellspacing="0" border="0" width="200">
<form action="<%=seldocument%>.asp?menu=<%=menu%>&FM_andre=<%=showother%>&<%=kalenderlink%>" method="POST" name="pickdate" id="pickdate">
<tr bgcolor="#5582D2">
	<td width="3" valign="top"><img src="<%=illpath%>ill/venstre_hjorne.gif" alt="" border="0"></td>
	<td width="194" colspan="2" style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid; padding-right:0; padding-bottom:3; padding-left:5; padding-top:2;">
	<font class="stor-hvid">Kalender</font> <%=kalenderMnavn%></td>
	<td width="3" valign="top" align="right"><img src="<%=illpath%>ill/hojre_hjorne.gif" alt="" border="0"></td>
</tr>
<tr bgcolor="#5582D2">
	<td valign="top"><img src="<%=illpath%>ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
	<td align="left" style="padding-bottom:3; padding-left:5;">
	<!--#include file="../../timereg/inc/weekselector_cal.asp"--></td>
	<td valign="top" align="right"><img src="<%=illpath%>ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td></tr>
	<tr><td colspan=4 bgcolor="#003399" height=1><img src="<%=illpath%>ill/blank.gif" width="1" height="1" alt="" border="0"></td></tr>
</form></table>

<table cellspacing="0" cellpadding="0" border="0" width="200" bgcolor="#FFFFFF">
	<tr bgcolor="ffffff" height="25">
		<td>
		<img src="<%=illpath%>ill/pile_tilbage.gif" alt="" vspace="1" hspace="2" border="0"><a class="vmenu" href="<%=seldocument%>.asp?menu=<%=menu%>&FM_andre=<%=showother%>&strdag=28&strmrd=<%=prevMonth%>&straar=<%=prevYear%>&<%=kalenderlink%>"><%=prevMonthName%></a>
		</td>
		<td align="center"><b><%=thisMonthName%>&nbsp;<%=strAar%></b></td>
		<td align="right">
		<a class="vmenu" href="<%=seldocument%>.asp?menu=<%=menu%>&FM_andre=<%=showother%>&strdag=1&strmrd=<%=nextMonth%>&straar=<%=nextYear%>&<%=kalenderlink%>"><%=nextMonthName%></a><img src="<%=illpath%>ill/pile_selected.gif" alt="" vspace="1" hspace="2" border="0">
		</td>
		</tr>
</table>
<table cellspacing="0" cellpadding="0" border="0" width="200" bgcolor="#FFFFFF">
<tr bgcolor="#ffffff">
	<td align=center width=15></td>
	<td align=center width=30>M</td>
    <td align=center width=30>T</td>
    <td align=center width=30>O</td>
    <td align=center width=30>T</td>
    <td align=center width=30>F</td>
    <td align=center width=30>L</td>
	<td align=center width=30>S</td>
</tr>
<tr>
	<td colspan="8" bgcolor="#000000"><img src="<%=illpath%>ill/blank.gif" alt="" border="0"></td>
</tr>

<%
'**************************************************************************************************'
'** Udskriver dage og timeforbrug i kalender.
'**************************************************************************************************'
WeekdayfirstW = firstWeekday 

fwdno = 1
acls = "class=vmenu"
acls2 = "class=vmenu"
'** ugenr **%>
<tr><td height=25 class=lillegray valign=top>&nbsp;&nbsp;<%=datepart("ww", firstDayOfMonth,2,2)%></td>
<%
'** Mellemrum før første dag i første uge
for fwdno = 1 to firstWeekday - 1 %>
<td>&nbsp;</td>
<%
next

daysinFirstWeek = 1

'*** Første uge datoer og timer/crm aktioner
for WeekdayfirstW = firstWeekday To 7 

if menu = "timereg" OR menu = "" then
	'** Henter timer ***
		strSQL = "SELECT sum(timer) AS timer_indtastet FROM timer WHERE Tmnr = "& varMed_id &" AND Tdato='"& strAar &"/" & strMrd & "/" & daysinFirstWeek &"' AND tfaktim <> 5"
		
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		intDayHoursVal = oRec("timer_indtastet")
		end if
		strAktnavn = ""
		oRec.close
else
	'** Henter crmaktioner ***
			strSQL3 = "SELECT crmhistorik.id, aktionsid, medarbid, kkundenavn, crmhistorik.navn, crmemne.navn AS emnenavn FROM crmhistorik, aktionsrelationer LEFT JOIN crmemne ON(crmemne.id = kontaktemne) LEFT JOIN kunder ON (kunder.kid = kundeid) WHERE ((crmdato = '"& strAar &"/" & strMrd & "/" & daysinFirstWeek &"') OR (crmdato <= '"& strAar &"/" & strMrd & "/" & daysinFirstWeek &"' AND crmdato_slut >= '"& strAar &"/" & strMrd & "/" & daysinFirstWeek &"' AND crmdato <> crmdato_slut)) "& usemedarbKri &" ORDER BY crmhistorik.id"
			oRec3.open strSQL3, oConn, 3
			akt_today = "n"
			strAktnavn = "Aktioner idag:" &vbcrlf
				while not oRec3.EOF
				strAktnavn = strAktnavn & oRec3("Kkundenavn") & chr(032) & oRec3("emnenavn") & chr(032) & oRec3("navn") &vbcrlf
				akt_today = "y"
				oRec3.movenext
				wend
			oRec3.close
			
end if		

Response.write "<td align=center valign=top>"
if cdate(day(now)&"/" & month(now)&"/" & year(now)) = cdate(daysinFirstWeek&"/" & strMrd &"/" & strAar) then
acls = "class=red"
else
	if menu = "crm" AND akt_today = "y" then 'er der aktioner idag
	acls = "class=vmenu"
	else
	acls = "class=kalblue"
	end if
end if%>
<a <%=acls%> href="<%=seldocument%>.asp?menu=<%=menu%>&FM_andre=<%=showother%>&strdag=<%=daysinFirstWeek%>&strmrd=<%=strMrd%>&straar=<%=strAar%>&<%=kalenderlink%>"><font title="<%=strAktnavn%>"><%=daysinFirstWeek%></font></a>
<%
if menu = "timereg" OR menu = "" then
	Response.write "<br>"
	if intDayHoursVal <> 0 then%>
	<span style="background-color:#D6DFF5; padding-left:2; padding-right:2;"><font class='lille-kalender' color=#003399><%=intDayHoursVal%></font></span>
	<%end if
end if%>
</td>
<%		
lastDayinfirstWeek = daysinFirstWeek
daysinFirstWeek = daysinFirstWeek + 1


next
%>
</tr>



<%
'*** De næste uger datoer og timer/ crm aktioner
startsecondWeek = lastDayinfirstWeek + 1
for dayCount = startsecondWeek to numberofdaysinmonth
	
	if menu = "timereg" OR menu = "" then
	strSQL = "SELECT sum(timer) AS timer_indtastet FROM timer WHERE Tmnr = "& varMed_id &" AND Tdato='"& strAar &"/" & strMrd & "/" & dayCount  &"' AND tfaktim <> 5"
		
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		intDayHoursVal = oRec("timer_indtastet")
		end if
		strAktnavn = ""
		oRec.close
	else
	
		'** Henter crmaktioner ***
			strSQL3 = "SELECT crmhistorik.id, aktionsid, medarbid, kkundenavn, crmhistorik.navn, crmemne.navn AS emnenavn FROM crmhistorik, aktionsrelationer LEFT JOIN crmemne ON(crmemne.id = kontaktemne) LEFT JOIN kunder ON (kunder.kid = kundeid) WHERE ((crmdato ='"& strAar &"/" & strMrd & "/" & dayCount &"') OR (crmdato <= '"& strAar &"/" & strMrd & "/" & dayCount &"' AND crmdato_slut >= '"& strAar &"/" & strMrd & "/" & dayCount &"' AND crmdato <> crmdato_slut)) "& usemedarbKri &" ORDER BY crmhistorik.id"
			oRec3.open strSQL3, oConn, 3
			akt_today = "n"
			strAktnavn = "Aktioner idag:" &vbcrlf
				while not oRec3.EOF
				strAktnavn = strAktnavn & oRec3("Kkundenavn") & chr(032) & oRec3("emnenavn") & chr(032) & oRec3("navn") &vbcrlf
				akt_today = "y"
				oRec3.movenext
				wend
			oRec3.close
			
	end if
	
	if Weekday(formatdatetime(dayCount &"/" & strMrd & "/" & strAar, 0), 2) = 1 then 
		if dayCount <> startsecondWeek then%>
		</tr>
		<%end if%>
	<tr>
		<td colspan="8" bgcolor="#D6DFF5"><img src="<%=illpath%>ill/blank.gif" alt="" border="0"></td>
	</tr>
	<tr>
		<td height=25 valign=top class=lillegray>&nbsp;<!--
		Bør ændres pga US og DK første ugedag ikke er identisk. 
		Så ugenr er ikke det samme.-->
		<%=datepart("ww", dayCount &"/" & strMrd & "/" & strAar,2,2)%>
</td>
	<%end if%>
<td align=center valign=top>
<%
if cdate(day(now)&"/" & month(now)&"/" & year(now)) = cdate(dayCount &"/" & strMrd &"/" & strAar) then
acls2 = "class=red"
else
	if menu = "crm" AND akt_today = "y" then 'er der aktioner idag
	acls2 = "class=vmenu"
	else
	acls2 = "class=kalblue"
	end if
end if%>
<a <%=acls2%> href="<%=seldocument%>.asp?menu=<%=menu%>&FM_andre=<%=showother%>&strdag=<%=dayCount%>&strmrd=<%=strMrd%>&straar=<%=strAar%>&<%=kalenderlink%>"><font title="<%=strAktnavn%>"><%=dayCount%></font></a>
<br>
<%
if menu = "timereg" OR menu = "" then
	if intDayHoursVal <> 0 then%>
	<span style="background-color:#D6DFF5; padding-left:2; padding-right:2;"><font class='lille-kalender' color=#003399><%=intDayHoursVal%></font></span>
	<%end if
end if%>
</td>
<%
next 
'**************************************************************************************************'
%>
</tr>	
</table>
<br><br>
<script>
	function expand2(div) {
		if (document.all(div).style.display == "none"){
			document.all(div).style.display = "";
		}else{
			document.all(div).style.display = "none";
		}
	}
</script>
<%

if instr(Request.ServerVariables("HTTP_USER_AGENT"), "Firefox") <> 0 then
wdt = 145
else
wdt = 195
end if


'Response.write "thisfile" & thisfile
if request("FM_andre") <> "show" AND thisfile <> "crmkalender" then%>
<!--include file="todo.asp"-->
<%end if%>
<br>
<%'********* Modul A4 se indtastninger for andre medarbejdere ***********
if strA4 = "on" then
	if level <= 1 AND andre <> "show" OR level = 6 AND andre <> "show" then%>
	<%if request("menu") <> "crm" then%>
	
	<DIV ID="vistimerformedarb" name="vistimerformedarb" style="position: absolute; display:; visibility:visible; top:282; z-index:500; width:<%=wdt%>px; padding:5px; border:2px #cccccc solid; background-color:#ffffff;">
	<table cellpadding="0" cellspacing="0" border="0">
	<tr>
		<td><img src="../ill/ac0020-24.gif" width="24" height="24" alt="" border="0"><b>Se time-registreringer for andre medarbejdere:</b><br><br></td>
	</tr>
	<%
		strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe FROM medarbejdere WHERE mansat = 1 OR mansat = 'yes' ORDER BY Mnavn"
		oRec.open strSQL, oConn, 3
		
		while not oRec.EOF 
		%>
		<tr>
			<td><a href="javascript:NewWin_cal('../inc/regular/calender.asp?FM_andre=show&FM_medarb_id=<%=oRec("Mid")%>&FM_Mnavn=<%=oRec("Mnavn")%>')" class="rmenu"><%=oRec("Mnavn")%>&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"></a></td>
		</tr>
		<%
		oRec.movenext
		wend
		oRec.close
		%>
		</table>
	</div>
	<%
	end if
	end if
end if
%>
</div>
	
	


