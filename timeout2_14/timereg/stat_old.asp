<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	
	<div id="sindhold" style="position:absolute; left:170; top:80; width:60%; height:600; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
    <td valign="top"><font class="pageheader">Statistik</font><br><br>
		
<!-------------------------------Sideindhold------------------------------------->


<%
strSedatoDag = day(now())  
strSedatoMrd = month(now())
strSedatoAar = year(now())
'response.write strSedatoDag &"-"
'response.write strSedatoMrd &"-"
'response.write strSedatoAar 

if strSedatoDag < 15 then
strDatodag = strSedatoMrd&"-"&"0"&strSedatoDag&"-"&strSedatoAar 
strDatoBeregn = (strSedatoDag + 31)-15
strMonthBeregn = strSedatoMrd - 1
strDatokri = StrMonthBeregn&"-"&strDatoBeregn&"-"&strSedatoAar
else
strDatodag = strSedatoMrd&"-"&strSedatoDag&"-"&strSedatoAar 
strDatoBeregn = day(now()) - 15
strMrdJust = month(now()) - 1
	if strMrdJust < 10 then
	strMrdJust = strMrdJust + 1
	strDatokri = "0"&strMrdJust&"-"&strDatoBeregn&"-"&strSedatoAar 
	else
	strDatokri = strMrdJust&"-"&strDatoBeregn&"-"&strSedatoAar
	end if
end if
strDatokriWrite = date - 15
strDatodagWrite = date

%>
Indtastninger for: <b><%= Session("user") %></b><br>
<br><br><br>


    <%
	oRec.Open "SELECT Tid, Tdato, Timer, Tfaktim, Tjobnr FROM timer ORDER BY Tdato", oConn, 3
	
	intFakbareTimer = 0
	totTimer = 0
	antalarbDage = 0
	lastDato = "1/1/2002"
	
	antalarbDageJanA = 0
	antalarbDageJanB = 0
	
	intFakbareTimerJanA = 0
	intFakbareTimerJanB = 0
	
	totTimerJanA = 0
	totTimerJanB = 0
	
	omsJanA = 0
	omsJanB = 0
	
	antalarbDageFebA = 0
	antalarbDageFebB = 0
	
	intFakbareTimerFebA = 0
	intFakbareTimerFebB = 0
	
	totTimerFebA = 0
	totTimerFebB = 0
	
	omsFebA = 0
	omsFebB = 0
	
	while not oRec.EOF 
		
		
		if oRec("Tfaktim") = 1 then
		intFakbareTimer = intFakbareTimer +oRec("timer")
		else
		intFakbareTimer = intFakbareTimer
		end if
		
		totTimer = totTimer +oRec("timer")
		'Slut Årstotal
		
	StrTdato = oRec("Tdato")
	'Response.write StrTdato & "<br>"
	
	' Hvis datoformat er Long Date e.lign.
	justerLangDatoT = instr(StrTdato, "20")
	if justerLangDatoT <> 0 then
	StrTdato_end = right(StrTdato, 2)
	StrTdato_start = left(StrTdato, (justerLangDatoT - 1)) 
	StrTdato = StrTdato_start & StrTdato_end
	end if
	
	if len(StrTdato) > 7 then
		strMrd = left(StrTdato, 2)
		strDag  = mid(StrTdato,4,2)
		strAar = Right(StrTdato,2)
	else
		if len(StrTdato) = 6 then
		strMrd = left(StrTdato, 1)
		strDag  = mid(StrTdato,3,1)
		strAar = Right(StrTdato,2)
		else
			varStregMrd = instr(StrTdato, "/")
			'Response.write varStregMrd
			if varStregMrd < 3 then 
			strMrd = left(StrTdato, 1)
			strDag  = mid(StrTdato,3,2)
			strAar = Right(StrTdato,2)
			'Response.write "H"
			else
			strMrd = left(StrTdato, 2)
			strDag  = mid(StrTdato,4,1)
			strAar = Right(StrTdato,2)
			end if
		end if
	end if
	
	
	
	Select case strMrd
	case 1
		
		if strDag < 15 then
		
		 	if oRec("Tdato") <> lastDato then
			lastDato =oRec("Tdato")
			antalarbDageJanA = antalarbDageJanA + 1
			else
			antalarbDageJanA = antalarbDageJanA
			end if
			
			totTimerJanA = totTimerJanA +oRec("timer")
			
				if oRec("Tfaktim") = 1 then
					intFakbareTimerJanA = intFakbareTimerJanA +oRec("timer")
					
					strSQL = "SELECT jobnr, jobTpris FROM job WHERE jobnr = " &oRec("Tjobnr")
					oRec2.open strSQL, oConn, 3
					
					omsJanA = (omsJanA + (550 *oRec("timer")))
					'Response.write omsJanA &", " & oRec2("jobTpris") &", " &oRec("timer") &"<br>" 
					timerTjTotJanA = timerTjTotJanA +oRec("timer")
					oRec2.close
				else
					intFakbareTimerJanA = intFakbareTimerJanA
				end if
			
		else
		
			if oRec("Tdato") <> lastDato then
			lastDato =oRec("Tdato")
			antalarbDageJanB = antalarbDageJanB + 1
			else
			antalarbDageJanB = antalarbDageJanB 
			end if
			totTimerJanB = totTimerJanB +oRec("timer")
			
				if oRec("Tfaktim") = 1 then
				
					intFakbareTimerJanB = intFakbareTimerJanB +oRec("timer")
					
					strSQL = "SELECT jobnr, jobTpris FROM job WHERE jobnr = " &oRec("Tjobnr") 
					oRec2.open strSQL, oConn, 3
					
					omsJanB = (omsJanB + (550 *oRec("timer")))
					'Response.write omsJanB &", " & oRec2("jobTpris") &", " &oRec("timer") &"<br>" 
					timerTjTotJanB = timerTjTotJanB +oRec("timer")
					oRec2.close
				else
					intFakbareTimerJanB = intFakbareTimerJanB
				end if
		end if
	case 2
	
		if strDag < 15 then
		
		 	if oRec("Tdato") <> lastDato then
			lastDato =oRec("Tdato")
			antalarbDageFebA = antalarbDageFebA + 1
			else
			antalarbDageFebA = antalarbDageFebA
			end if
			
			totTimerFebA = totTimerFebA +oRec("timer")
			
				if oRec("Tfaktim") = 1 then
					intFakbareTimerFebA = intFakbareTimerFebA +oRec("timer")
					
					strSQL = "SELECT jobnr, jobTpris FROM job WHERE jobnr = " &oRec("Tjobnr")
					oRec2.open strSQL, oConn, 3
					
					omsFebA = (omsFebA + (550 *oRec("timer")))
					'Response.write omsFebA &", " & oRec2("jobTpris") &", " &oRec("timer") &"<br>" 
					timerTjTotFebA = timerTjTotFebA +oRec("timer")
					oRec2.close
				else
					intFakbareTimerFebA = intFakbareTimerFebA
				end if
			
		else
		
			if oRec("Tdato") <> lastDato then
			lastDato =oRec("Tdato")
			antalarbDageFebB = antalarbDageFebB + 1
			else
			antalarbDageFebB = antalarbDageFebB 
			end if
			totTimerFebB = totTimerFebB +oRec("timer")
			
				if oRec("Tfaktim") = 1 then
				
					intFakbareTimerFebB = intFakbareTimerFebB +oRec("timer")
					
					strSQL = "SELECT jobnr, jobTpris FROM job WHERE jobnr = " &oRec("Tjobnr") 
					oRec2.open strSQL, oConn, 3
					
					omsFebB = (omsFebB + (550 *oRec("timer")))
					'Response.write omsFebB &", " & oRec2("jobTpris") &", " &oRec("timer") &"<br>" 
					timerTjTotFebB = timerTjTotFebB +oRec("timer")
					oRec2.close
				else
					intFakbareTimerFebB = intFakbareTimerFebB
				end if
		end if
		
		
	case 3
	
	case 4
	
	case 5
	
	case 6
	
	case 7
	
	case 8
	
	case 9
	
	case 10
	
	case 11
	
	case 12
	
	end select
	
	oRec.movenext
	wend
	
	
	'Her adderes de forskelle Data til det endelige Output
	'Jan
	'Response.write "antalarbDageJanA: " & antalarbDageJanA
	if antalarbDageJanA <> 0 then
	avrTimerJanA = (totTimerJanA/antalarbDageJanA)
	avrFaktimerJanA = (intFakbareTimerJanA/antalarbDageJanA) 
	else
	avrTimerJanA = 0.00
	avrFaktimerJanA = 0.00
	end if
	
	'Response.write ", antalarbDageJanB: " & antalarbDageJanB
	if antalarbDageJanB <> 0 then
	avrTimerJanB = (totTimerJanB/antalarbDageJanB)
	avrFaktimerJanB = (intFakbareTimerJanB/antalarbDageJanB) 
	else
	avrTimerJanB = 0.00
	avrFaktimerJanB = 0.00
	end if
	
	omsJan = (omsJanA + omsJanB)
	' Slut Addering Januar
	
	
	
	'Her adderes de forskelle Data til det endelige Output
	'Feb
	if antalarbDageFebA <> 0 then
	avrTimerFebA = (totTimerFebA/antalarbDageFebA)
	avrFaktimerFebA = (intFakbareTimerFebA/antalarbDageFebA) 
	else
	avrTimerFebA = 0.00
	avrFaktimerFebA = 0.00
	end if
	
	
	if antalarbDageFebB <> 0 then
	avrTimerFebB = (totTimerFebB/antalarbDageFebB)
	avrFaktimerFebB = (intFakbareTimerFebB/antalarbDageFebB) 
	else
	avrTimerFebB = 0.00
	avrFaktimerFebB = 0.00
	end if
	
	
	omsFeb = (omsFebA + omsFebB)
	' Slut Addering Februar
	
	antalarbDage = (antalarbDageJanA + antalarbDageJanB + antalarbDageFebA + antalarbDageFebB)
	avrTimer = (totTimer/antalarbDage)
	avrFaktimer = (intFakbareTimer/antalarbDage) 
	
	%>
<table cellspacing="0" cellpadding="2" border="1" width="600" bordercolorDark="#C1E1d4">
<tr>
	<td><FONT FACE="verdana" SIZE="1" COLOR="#A90014">Arb. Dage: <b><%=antalarbDage%></b></td>
    <td><FONT FACE="verdana" SIZE="1" COLOR="#A90014">Fak Job</td>
	<td><FONT FACE="verdana" SIZE="1" COLOR="#A90014">Timer Total: <b><%=totTimer%></b></td>
	<td><FONT FACE="verdana" SIZE="1" COLOR="#A90014">Fak Timer: <b><%=intFakbareTimer%></b></td>
	
</tr>
<tr>
    <td><FONT FACE="verdana" SIZE="1" COLOR="#A90014">&nbsp;</td>
    <td><FONT FACE="verdana" SIZE="1" COLOR="#A90014">Fordelt på job</td>
	<td><FONT FACE="verdana" SIZE="1" COLOR="#A90014">Avr. timer/dag: <b><%=left(avrTimer,4)%></b></td>
	<td><FONT FACE="verdana" SIZE="1" COLOR="#A90014">Avr. Fak timer/dag: <b><%=left(avrFaktimer,4)%></b></td>
	
</tr>
</table>
<br>
<br>
&nbsp;&nbsp;Avr. arbejdstimer pr. dag:<br>
<table cellspacing="0" cellpadding="0" border="1" bordercolorDark="darkRed" width="600">
<tr>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jan</td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jan</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Feb</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Feb</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Mar</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Mar</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Apr</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Apr</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Maj</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Maj</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jun</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jun</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jul</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jul</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Aug</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Aug</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Sep</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Sep</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Okt</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Okt</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Nov</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Nov</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Dec</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Dec</td>
</tr>
<tr>
    <td width="20"><img src="ill/scala.gif" width="20" height="100" alt="" border="0"></td>
    <td valign="bottom" width="20"><img src="ill/scalag1.gif" width="25" height="<%=10*(left(avrTimerJanA,4))%>" alt="" border="0"></td>
    <td valign="bottom" width="20"><img src="ill/scalag1.gif" width="25" height="<%=10*(left(avrTimerJanB,4))%>" alt="" border="0"></td>
	<td valign="bottom" width="20"><img src="ill/scalag1.gif" width="25" height="<%=10*(left(avrTimerFebA,4))%>" alt="" border="0"></td>
    <td valign="bottom" width="20"><img src="ill/scalag1.gif" width="25" height="<%=10*(left(avrTimerFebB,4))%>" alt="" border="0"></td> 
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
</tr>
<tr bgcolor="#ffffff">
    <td width="20">&nbsp;</td>
    <td valign="bottom" width="25">&nbsp;<%=left(avrTimerJanA,4)%></td>
    <td valign="bottom" width="25">&nbsp;<%=left(avrTimerJanB,4)%></td>
	<td valign="bottom" width="25">&nbsp;<%=left(avrTimerFebA,4)%></td>
    <td valign="bottom" width="25">&nbsp;<%=left(avrTimerFebB,4)%></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
</tr>
</table>
<br>
<br>
&nbsp;&nbsp;Avr. Fakturérbare arbejdstimer pr. dag:<br>
<table cellspacing="0" cellpadding="0" border="1"  bordercolorDark="darkRed" width="600">
<tr>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jan</td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jan</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Feb</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Feb</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Mar</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Mar</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Apr</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Apr</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Maj</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Maj</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jun</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jun</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jul</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jul</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Aug</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Aug</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Sep</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Sep</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Okt</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Okt</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Nov</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Nov</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Dec</td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Dec</td>
</tr>
<tr>
    <td width="20"><img src="ill/scala.gif" width="20" height="100" alt="" border="0"></td>
    <td valign="bottom" width="20"><img src="ill/scalag1.gif" width="25" height="<%=10*(left(avrFaktimerJanA,4))%>" alt="" border="0"></td>
    <td valign="bottom" width="20"><img src="ill/scalag1.gif" width="25" height="<%=10*(left(avrFaktimerJanB,4))%>" alt="" border="0"></td>
	<td valign="bottom" width="20"><img src="ill/scalag1.gif" width="25" height="<%=10*(left(avrFaktimerFebA,4))%>" alt="" border="0"></td>
    <td valign="bottom" width="20"><img src="ill/scalag1.gif" width="25" height="<%=10*(left(avrFaktimerFebB,4))%>" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
</tr>
<tr bgcolor="#ffffff">
    <td width="20">&nbsp;</td>
    <td valign="bottom" width="25">&nbsp;<%=left(avrFakTimerJanA,4)%></td>
    <td valign="bottom" width="25">&nbsp;<%=left(avrFakTimerJanB,4)%></td>
	<td valign="bottom" width="25">&nbsp;<%=left(avrFakTimerFebA,4)%></td>
    <td valign="bottom" width="25">&nbsp;<%=left(avrFakTimerFebB,4)%></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
</tr>
</table>
<br>
<br>
<%timTotTjJan = (timerTjTotJanA + timerTjTotJanB)%>
Omsætning Januar:&nbsp;<b><%=timTotTjJan &" timer = " & omsJan%>,- kr. (excl. moms)</b>
<br>
<%timTotTjFeb = (timerTjTotFebA + timerTjTotFebB)%>
Omsætning Febuar:&nbsp;<b><%=timTotTjFeb &" timer = " & omsFeb%>,- kr. (excl. moms)</b>
<br><br>
<br>
<a href="Javascript:history.back()">Tilbage</a>
<br>
<br>
</TD>
</TR>
</TABLE>
</div>

<%
 'now close and clean up
	oRec.Close
	oConn.close
	Set oRec = Nothing
	set oConn = Nothing 

end if 


%>
<!--#include file="../inc/regular/footer_inc.asp"-->
