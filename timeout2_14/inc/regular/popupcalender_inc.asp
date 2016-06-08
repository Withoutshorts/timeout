<!--#include file="../connection/conn_db_inc.asp"-->
<script type="text/javascript">
 function updato(dag, md, aar, hv){
 var thisdag = dag;
 var thismd = md;
 var thisaar = aar; 
 var thishv = hv;

  if (thishv == 1) {
  window.opener.document.all["FM_start_dag"].value = thisdag;
  window.opener.document.all["FM_start_mrd"].value = thismd;
  window.opener.document.all["FM_start_aar"].value = thisaar;
  window.close();
  }
		if (thishv == 2){
		window.opener.document.all["FM_start_dag_2"].value = thisdag;
	  	window.opener.document.all["FM_start_mrd_2"].value = thismd;
	  	window.opener.document.all["FM_start_aar_2"].value = thisaar;
		window.close();
		}
		
			if (thishv == 3){
			window.opener.document.all["FM_gentag_dag"].value = thisdag;
		  	window.opener.document.all["FM_gentag_mrd"].value = thismd;
		  	window.opener.document.all["FM_gentag_aar"].value = thisaar;
			window.close();
			}
			
			if (thishv == 4){
			window.opener.document.all["FM_slut_dag"].value = thisdag;
		  	window.opener.document.all["FM_slut_mrd"].value = thismd;
		  	window.opener.document.all["FM_slut_aar"].value = thisaar;
			window.close();
            }
            if (thishv == 99) {
                window.opener.document.all["FM_slut_dag_2"].value = thisdag;
                window.opener.document.all["FM_slut_mrd_2"].value = thismd;
                window.opener.document.all["FM_slut_aar_2"].value = thisaar;
                window.close();
            }
			
			if (thishv == 5){
			window.opener.document.all["FM_slut_dag"].value = thisdag;
		  	window.opener.document.all["FM_slut_mrd"].value = thismd;
			thisaar = thisaar + "";
			antalchar = thisaar.length 
			var year = eval(thisaar.substring(2, 4));
			if (year < 10) {
			year = "0"+year
			}
		  	window.opener.document.all["FM_slut_aar"].value = year
			window.close();
			}
			
			if (thishv == 6){
			window.opener.document.all["FM_start_dag"].value = thisdag;
		  	window.opener.document.all["FM_start_mrd"].value = thismd;
			thisaar = thisaar + "";
			antalchar = thisaar.length 
			var year = eval(thisaar.substring(2, 4));
			if (year < 10) {
			year = "0"+year
			}
		  	window.opener.document.all["FM_start_aar"].value = year
			window.close();
            }

            if (thishv == 7) {
                alert(aar)
                window.opener.document.getElementById("FM_start_dag").value = thisdag;
                window.opener.document.getElementById("FM_start_mrd").value = thismd;
                thisaar = thisaar + "";
                antalchar = thisaar.length
                var year = eval(thisaar.substring(2, 4));
                if (year < 10) {
                    year = "0" + year
                }
                window.opener.document.getElementById("FM_start_aar").value = year;
                window.close();
            }
			
		}
</script>
<LINK rel="stylesheet" type="text/css" href="../style/timeout_style.css">
<%
'******************************************* dato funktioner ***********************************
		
		'*** Sætter lokal dato/kr format. *****
		Session.LCID = 1030
		
		select case request("use") 
		case 1
		useFMdato = 1
		case 2
		useFMdato = 2
		case 3
		useFMdato = 3
		case 4
		useFMdato = 4
		case 5
		useFMdato = 5
		case 6
		useFMdato = 6
		case 7
		useFMdato = 7
		case 99
		useFMdato = 99
		end select
		
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
		
		firstWeekday = Weekday(firstDayOfMonth ,2) 
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
		
		
		seldocument = "popupcalender_inc"
		kalenderlink = "shownumofdays="&shownumofdays&"&ketype=e&selpkt=kal&status="&status&"&id="&id&"&emner="&emner&"&medarb="&medarb&"&sort="&request("sort")&""
		showother = "notused"
		
		illpath = "../../"
	
'*****************************************************************************************
%>
<div id="popupcal" style="position:absolute; left:20; top:10; visibility:visible; z-index=1000;">
<table cellpadding="0" cellspacing="0" border="0" width="200">
<form action="<%=seldocument%>.asp?use=<%=useFMdato%>" method="POST" name="pickdate" id="pickdate">
<tr bgcolor="#5582D2">
	<td width="3" valign="top"><img src="<%=illpath%>ill/venstre_hjorne.gif" alt="" border="0"></td>
	<td width="194" colspan="2" style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid; padding-right:0; padding-bottom:3; padding-left:5; padding-top:2;">
	<font class="stor-hvid">Kalender</font></td>
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
		<img src="<%=illpath%>ill/pile_tilbage.gif" alt="" vspace="1" hspace="2" border="0"><a class="vmenu" href="<%=seldocument%>.asp?strdag=28&strmrd=<%=prevMonth%>&straar=<%=prevYear%>&use=<%=useFMdato%>"><%=prevMonthName%></a>
		</td>
		<td align="center"><b><%=thisMonthName%>&nbsp;<%=strAar%></b></td>
		<td align="right">
		<a class="vmenu" href="<%=seldocument%>.asp?strdag=1&strmrd=<%=nextMonth%>&straar=<%=nextYear%>&use=<%=useFMdato%>"><%=nextMonthName%></a><img src="<%=illpath%>ill/pile_selected.gif" alt="" vspace="1" hspace="2" border="0">
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
'** Udskriver dage i kalender.
'**************************************************************************************************'
WeekdayfirstW = firstWeekday 
fwdno = 1
acls = "class=vmenu"
acls2 = "class=vmenu"
'** ugenr **%>
<tr><td height=25 class=lillegray valign=top>&nbsp;<%=datepart("ww", firstDayOfMonth,2,2)%></td>
<%
'** Mellemrum før første dag i første uge
for fwdno = 1 to firstWeekday - 1%>
<td>&nbsp;</td>
<%
next

daysinFirstWeek = 1

'*** Første uge datoer og timer/crm aktioner
for WeekdayfirstW = firstWeekday To 7 

Response.write "<td align=center valign=top>"
if cdate(day(now)&"/" & month(now)&"/" & year(now)) = cdate(daysinFirstWeek&"/" & strMrd &"/" & strAar) then
acls = "class=red"
else
acls = "class=kalblue"
end if%>
<a <%=acls%> href="#" onClick="updato('<%=daysinFirstWeek%>','<%=strMrd%>','<%=strAar%>','<%=useFMdato%>')"><%=daysinFirstWeek%></a>
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
	
	if Weekday(formatdatetime(dayCount &"/" & strMrd & "/" & strAar, 0), 2) = 1 then 
		if dayCount <> startsecondWeek then%>
		</tr>
		<%end if%>
	<tr>
		<td colspan="8" bgcolor="#D6DFF5"><img src="<%=illpath%>ill/blank.gif" alt="" border="0"></td>
	</tr>
	<tr>
		<td height=25 valign=top class=lillegray>&nbsp;<%=datepart("ww", formatdatetime(dayCount &"/" & strMrd & "/" & strAar, 0), 2, 2)%></td>
	<%end if%>
<td align=center valign=top>
<%

if cdate(day(now)&"/" & month(now)&"/" & year(now)) = cdate(dayCount &"/" & strMrd &"/" & strAar) then
acls2 = "class=red"
else
acls2 = "class=kalblue"
end if%>
<a <%=acls2%> href="#" onClick="updato('<%=dayCount%>','<%=strMrd%>','<%=strAar%>','<%=useFMdato%>')"><%=dayCount%></a>
</td>
<%
next 
'**************************************************************************************************'
%>
</tr>	
</table>
</div>
	
	


