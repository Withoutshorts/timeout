
<!--#include file="../xml/calender_xml_inc.asp"-->

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

<%
varMed_id = usemrn
leftDiv = 0
topDiv = 10
end if


'******************************************* dato funktioner ***********************************
		
		'*** Sætter lokal dato/kr format. *****
		'Session.LCID = 1030
		
		If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		
		'cval = "her"
		
			if len(request("FM_start_dag")) <> 0 then 'Fra datovælger form
			strMrd = request("FM_start_mrd")
            strAar = request("FM_start_aar")
				select case strMrd
				case 2
                    
                    select case right(strAar, 2)
			        case "04", "08", "12", "16", "20", "24", "28", "32", "36", "40", "44"

					if request("FM_start_dag") > 29 then
					strDag = 29
					else
					strDag = request("FM_start_dag")
					end if

                    case else

                    if request("FM_start_dag") > 28 then
					strDag = 28
					else
					strDag = request("FM_start_dag")
					end if

                    end select


					'if request("FM_start_dag") > 28 then
					'strDag = 28
					'else
					'strDag = request("FM_start_dag")
					'end if


				case 1, 3, 5, 7, 8, 10, 12
					strDag = request("FM_start_dag")
				case else
					if request("FM_start_dag") > 30 then
					strDag = 30
					else
					strDag = request("FM_start_dag")
					end if
				end select
			
			else
				
				
				
				if len(request.cookies("timereg_2006")("dato")) <> 0 AND session("forste") = "n" then
					
					strDag = day(request.cookies("timereg_2006")("dato"))
					strMrd = month(request.cookies("timereg_2006")("dato"))
					strAar = year(request.cookies("timereg_2006")("dato"))
					
					'cval = "cook"
					
				else
					
					strDag = day(now)
					strMrd = month(now)
					strAar = year(now)
					
					'cval = "dagsdato"
				
				end if
			
			end if
		else
			
            if len(request("strdag")) <> 0 then 'Fra dato link

            strAar = request("straar")
			strMrd = request("strmrd")

				select case strMrd
				case 2
                    select case right(strAar, 2)
			        case "04", "08", "12", "16", "20", "24", "28", "32", "36", "40", "44"

					if request("strdag") > 29 then
					strDag = 29
					else
					strDag = request("strdag")
					end if

                    case else

                    if request("strdag") > 28 then
					strDag = 28
					else
					strDag = request("strdag")
					end if

                    end select

				case 1, 3, 5, 7, 8, 10, 12
					strDag = request("strdag")
				case else
					if request("strdag") > 30 then
					strDag = 30
					else
					strDag = request("strdag")
					end if
				end select
			
			else
				
				if len(request.cookies("timereg_2006")("dato")) <> 0 AND session("forste") = "n" then
					
					strDag = day(request.cookies("timereg_2006")("dato"))
					strMrd = month(request.cookies("timereg_2006")("dato"))
					strAar = year(request.cookies("timereg_2006")("dato"))
					
					'cval = "cook"
					
				else
					
					strDag = day(now)
					strMrd = month(now)
					strAar = year(now)
					
					'cval = "dagsdato"
				
				end if
			
			end if
		end if
		
		daynow = formatdatetime(day(now) & "/" & month(now) & "/" & year(now), 0)
		useDate = formatdatetime(strDag & "/" & strMrd & "/" & strAar, 0)
		firstDayOfMonth = formatdatetime(1 & "/" & strMrd & "/" & strAar, 0)
		
		useDatePrevWeek = dateadd("d", -7, useDate)
		useDateNextWeek = dateadd("d", 7, useDate)
		
		Select case strMrd
		case 1, 3, 5, 7, 8, 10, 12
		lastDayOfMonth = formatdatetime("31/" & strMrd & "/" & strAar, 0)
		numberofdaysinmonth = 31
		case 2
			select case right(strAar, 2)
			case "04", "08", "12", "16", "20", "24", "28", "32", "36", "40", "44"
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
		
		
			seldocument = "timereg_akt_2006"
			showother = "dontshow"
			kalenderlink = "searchstring="&searchstring&""
			illpath = "../"
			
		
	response.cookies("timereg_2006")("dato") = strDag&"/"&strMrd&"/"&strAar
	
'*****************************************************************************************

kalenderMnavn = ""


if request("FM_andre") <> "show" AND (thisfile = "timereg_akt_2006") then


call meStamdata(usemrn)
kalenderMnavn = "<font class=megetlillehvid>"& meTxt &"</font>"

end if


if print <> "j" then
%>
<table cellpadding="0" cellspacing="0" border="0" width="100%">
<form action="timereg_akt_2006.asp?jobid=<%=jobid%>&usemrn=<%=usemrn%>&showakt=1&fromsdsk=<%=fromsdsk%>" method="POST" name="pickdate" id="pickdate">
<tr bgcolor="#5C75AA">
	<td width="194" colspan="4" style="padding:2px 3px 0px 5px;">
	<font class="stor-hvid"><%=calender_txt_107%></font> <%=kalenderMnavn%></td>
</tr>
<tr bgcolor="#5C75AA">
	<td valign="top"><img src="<%=illpath%>ill/blank.gif" width="1" height="25" alt="" border="0"></td>
	<td align="left" style="padding-bottom:3px; padding-left:5px;">
	<!--#include file="../../timereg/inc/weekselector_cal.asp"--></td>
	<td valign="top" align="right"><img src="<%=illpath%>ill/blank.gif" width="1" height="25" alt="" border="0"></td></tr>
	<tr><td colspan=4 bgcolor="#003399" height=1><img src="<%=illpath%>ill/blank.gif" width="1" height="1" alt="" border="0"></td></tr>
</form></table>

<table cellspacing="0" cellpadding="0" border="0" width="100%" bgcolor="#FFFFFF">
	<tr bgcolor="ffffff" height="25">
		<td>
		<img src="<%=illpath%>ill/pile_tilbage.gif" alt="" vspace="1" hspace="2" border="0">
		<a class="vmenu" href="timereg_akt_2006.asp?jobid=<%=jobid%>&usemrn=<%=usemrn%>&showakt=<%=showakt%>&strdag=28&strmrd=<%=prevMonth%>&straar=<%=prevYear%>&<%=kalenderlink%>&fromsdsk=<%=fromsdsk%>"><%=prevMonthName%></a></td>
		<td align="center"><b><%=thisMonthName%>&nbsp;<%=strAar%></b></td>
		<td align="right"><a class="vmenu" href="timereg_akt_2006.asp?jobid=<%=jobid%>&usemrn=<%=usemrn%>&showakt=<%=showakt%>&strdag=1&strmrd=<%=nextMonth%>&straar=<%=nextYear%>&<%=kalenderlink%>&fromsdsk=<%=fromsdsk%>"><%=nextMonthName%></a><img src="<%=illpath%>ill/pile_selected.gif" alt="" vspace="1" hspace="2" border="0"></td>
		</tr>
</table>
<table cellspacing="0" cellpadding="0" border="0" width="100%" bgcolor="#FFFFFF">
<tr bgcolor="#ffffff">
	<td align=center width=15></td>
	<td align=center width=30><%=calender_txt_108 %></td>
    <td align=center width=30><%=calender_txt_109 %></td>
    <td align=center width=30><%=calender_txt_110 %></td>
    <td align=center width=30><%=calender_txt_111 %></td>
    <td align=center width=30><%=calender_txt_112 %></td>
    <td align=center width=30><%=calender_txt_113 %></td>
	<td align=center width=30><%=calender_txt_114 %></td>
	<td align=center width=30>Tot.</td>
</tr>
<tr>
	<td colspan="9" bgcolor="#000000"><img src="<%=illpath%>ill/blank.gif" alt="" border="0"></td>
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
<tr><td height=25 class=lillegray valign=top align=right style="padding:0px 3px 0px 3px;">
<%=datepart("ww", firstDayOfMonth,2,2)%></td>

<%
'** Mellemrum før første dag i første uge
for fwdno = 1 to firstWeekday - 1 %>
<td>&nbsp;</td>
<%
next

daysinFirstWeek = 1
cweekTot = 0
mTot = 0
'** Finder de typer der er med i det daglige timeregnskab ***'
call akttyper2009(2)

'*** Første uge datoer og timer/crm aktioner
for WeekdayfirstW = firstWeekday To 7 

if menu = "timereg" OR menu = "" then
	'** Henter timer ***
    '** Alle der tælles med i daglig timerengskab + 
		strSQL = "SELECT timer AS timer_indtastet, tfaktim FROM timer WHERE Tmnr = "& varMed_id &" AND Tdato='"& strAar &"/" & strMrd & "/" & daysinFirstWeek &"'"
    
        select case lto
        case "tec", "xintranet - local", "esn"
        strSQL = strSQL &" AND (tfaktim <> 0)"
        case else
        strSQL = strSQL & "  AND (("& aty_sql_realhours &")"_
		& " OR (tfaktim = 30 OR tfaktim = 31 OR tfaktim = 7 OR tfaktim = 11))"
        end select 
    
  
		
        

        '30 = Overarbejde / afspadsering optjent 
        '31 = Afspadsering
        '7 Fleks afholdt
        '11 planlagt ferie
        
        intDayHoursVal = 0
		'Response.Write strSQL
		cal_cls = "cal"
        focls = 0
		oRec.open strSQL, oConn, 3
		while not oRec.EOF 
        
       
        intDayHoursVal = intDayHoursVal + (oRec("timer_indtastet")/1)
        
        'select case lto
        'case "wwf"
        '        if oRec("tfaktim") <> 30 then
        '        cweekTot = cweekTot + (oRec("timer_indtastet")/1)
        '        end if
        'case else
            '** Kan se timer på dagen med skal de regnes med ugetotal, fleks mv.?
            call akttyper2009Prop(oRec("tfaktim"))
            if cint(aty_real) = 1 then
            cweekTot = cweekTot + (oRec("timer_indtastet")/1)
            else
            cweekTot = cweekTot
            end if
        'end select


            if focls = 0 then
                select case oRec("tfaktim")
                case 7,8,11,13,14,18,19,20,21,22,23,24,25,26,31,23,115,120,121,122
                cal_cls = "calalt"
                focls = 1
                case else
                cal_cls = "cal"
                end select

                'if oRec("tfaktim") = 11 OR oRec("tfaktim") = 13 OR oRec("tfaktim") = 14 _
                'OR oRec("tfaktim") = 7 OR oRec("tfaktim") = 18 _
                'OR oRec("tfaktim") = 19 OR oRec("tfaktim") = 20 OR oRec("tfaktim") = 21 OR oRec("tfaktim") = 22 OR oRec("tfaktim") = 23 OR oRec("tfaktim") = 24 OR oRec("tfaktim") = 25 _
                'OR oRec("tfaktim") = 31 OR oRec("tfaktim") = 8 OR oRec("tfaktim") = 23 OR oRec("tfaktim") = 115 then 'Ferie, Afspad, Sygdom m.fl.
                'cal_cls = "calalt"
                'focls = 1
                'else
                'cal_cls = "cal"
                'end if 
            end if

	    oRec.movenext
        wend
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
%>
<td align=center valign=top>
<%
    
call helligdage(daysinFirstWeek &"/" & strMrd &"/" & strAar, 0, lto)
if erHellig = 1 then
 
acls = "class=kalsilver"

else

	if cdate(day(now)&"/" & month(now)&"/" & year(now)) = cdate(daysinFirstWeek&"/" & strMrd &"/" & strAar) then
	acls = "class=kalred"
	else
	acls = "class=kalblue"
	end if

end if
%>
<a <%=acls%> href="timereg_akt_2006.asp?jobid=<%=jobid%>&usemrn=<%=usemrn%>&showakt=1&strdag=<%=daysinFirstWeek%>&strmrd=<%=strMrd%>&straar=<%=strAar%>&<%=kalenderlink%>&fromsdsk=<%=fromsdsk%>"><%=daysinFirstWeek%></a>
<%

	Response.write "<br>"
	if intDayHoursVal <> 0 then
    
    if intDayHoursVal > 100 then
    %>
    <span class="<%=cal_cls%>"><%=formatnumber(intDayHoursVal, 0)%></span>
    <%else%>
	<span class="<%=cal_cls%>"><%=formatnumber(intDayHoursVal, 2)%></span>
	
	<%end if
	
	end if%>
</td>
<%		
lastDayinfirstWeek = daysinFirstWeek
daysinFirstWeek = daysinFirstWeek + 1


next
%>
<td valign=bottom align=right><span class="caltot"><%=formatnumber(cweekTot, 2) %></span>
<%mTot = mTot +  cweekTot%>
		<%cweekTot = 0 %></td>
</tr>



<%
'*** De næste uger datoer og timer/ crm aktioner
startsecondWeek = lastDayinfirstWeek + 1
cweekTot = 0
for dayCount = startsecondWeek to numberofdaysinmonth
	
	
	
	
	if Weekday(formatdatetime(dayCount &"/" & strMrd & "/" & strAar, 0), 2) = 1 then 
		
		if dayCount <> startsecondWeek then%>
		<td valign=bottom align=right><span class=caltot><%=formatnumber(cweekTot, 2) %></span>
		<%mTot = mTot +  cweekTot%>
		<%cweekTot = 0 %></td>
		</tr>
		<%end if%>
	<tr>
		<td colspan="9" bgcolor="#D6DFF5"><img src="<%=illpath%>ill/blank.gif" alt="" border="0"></td>
	</tr>
	<tr>
		<td height=25 class=lillegray valign=top align=right style="padding:0px 3px 0px 3px;">
		<%=datepart("ww", dayCount &"/" & strMrd & "/" & strAar,2,2)%>
</td>
	<%end if%>
	
	<%
	
	    strSQL = "SELECT timer AS timer_indtastet, tfaktim FROM timer WHERE Tmnr = "& varMed_id &" AND Tdato='"& strAar &"/" & strMrd & "/" & dayCount &"'"
		
        '31 = Afspadsering SKAL altid vises med orange og altid med uanset om den ikke tælles med i dagligt timeregnskab.
		'(tfaktim = 1 OR tfaktim = 2 OR tfaktim = 6 OR tfaktim = 13 OR tfaktim = 14 OR tfaktim = 20 OR tfaktim = 21)"
		

        select case lto
        case "tec", "xintranet - local", "esn"
        strSQL = strSQL &" AND (tfaktim <> 0)"
        case else
        strSQL = strSQL & "  AND (("& aty_sql_realhours &")"_
		& " OR (tfaktim = 30 OR tfaktim = 31 OR tfaktim = 7 OR tfaktim = 11))"
        end select 

        intDayHoursVal = 0
		'Response.Write strSQL
		cal_cls = "cal"
        focls = 0
		
        oRec.open strSQL, oConn, 3
		while not oRec.EOF 
		intDayHoursVal = intDayHoursVal + (oRec("timer_indtastet")/1)
            
            '** Kan se timer på dagen med skal de regnes med ugetotal, fleks mv.?
            call akttyper2009Prop(oRec("tfaktim"))
            if cint(aty_real) = 1 then
            cweekTot = cweekTot + (oRec("timer_indtastet")/1)
            else
            cweekTot = cweekTot
            end if

            if focls = 0 then
                select case oRec("tfaktim")
                case 7,8,11,13,14,18,19,20,21,22,23,24,25,26,31,23,115,120,121,122
                cal_cls = "calalt"
                focls = 1
                case else
                cal_cls = "cal"
                end select
            end if

	    oRec.movenext
        wend
		strAktnavn = ""
		oRec.close
		
	
	 %>
	
<td align=center valign=top>
<%


call helligdage(dayCount &"/" & strMrd &"/" & strAar, 0, lto)
if erHellig = 1 then
 
acls2 = "class=kalsilver"

else

if cdate(day(now)&"/" & month(now)&"/" & year(now)) = cdate(dayCount &"/" & strMrd &"/" & strAar) then
acls2 = "class=kalred"
else
acls2 = "class=kalblue"
end if

end if%>


<a <%=acls2%> href="timereg_akt_2006.asp?jobid=<%=jobid%>&usemrn=<%=usemrn%>&showakt=1&strdag=<%=dayCount%>&strmrd=<%=strMrd%>&straar=<%=strAar%>&<%=kalenderlink%>&fromsdsk=<%=fromsdsk%>"><%=dayCount%></a>
<br>
<%

    if intDayHoursVal <> 0 then

    if intDayHoursVal > 100 then
    %>
    <span class="<%=cal_cls%>"><%=formatnumber(intDayHoursVal, 0)%></span>
    <%else%>
	<span class="<%=cal_cls%>"><%=formatnumber(intDayHoursVal, 2)%></span>
	
	<%end if

	
	end if%>
</td>
<%

lastDay = dayCount
   
next 
'**************************************************************************************************'
ld = Weekday(formatdatetime(lastDay &"/" & strMrd & "/" & strAar, 0), 2)

for d = ld + 1 to 7 %>
<td>&nbsp</td>
<%next %>

<td valign=bottom align=right>
<span class=caltot><%=formatnumber(cweekTot, 2) %>
<%mTot = mTot +  cweekTot%></span> <br />
<span class=caltot><b><%=formatnumber(mTot, 2) %></b></span></td>
</tr>
<tr>
    <td colspan=20 style="font-size:9px; padding:3px 5px 2px 2px;" align=right>
<span class="calalt" style="border:1px #999999 solid; font-size:8px; width:15px; height:10px;">&nbsp;</span>&nbsp;<%=calender_txt_115 %>&nbsp;&nbsp;
<span class="cal" style="border:1px #999999 solid; font-size:8px; width:15px; height:10px;">&nbsp;</span>&nbsp;<%=calender_txt_116 %>
</td>
</tr>

</table>



<%cweekTot = 0 
mTot = 0%>
<%end if %>


