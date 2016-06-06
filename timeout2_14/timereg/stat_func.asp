<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include file="inc/conn_dbAdm.asp"-->
<!--#include file="inc/functions_inc.asp"-->
<%
if session("user") = "" then
	%>
	<!--#include file="inc/loggetaf.asp"-->
	<%
	else
	%>
	<!--#include file="inc/topmenu.asp"-->
	<table cellspacing="0" cellpadding="0" border="0">
	<tr>
   	 	<td width="5" rowspan="2" BGCOLOR="#C1E1d4"><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td BGCOLOR="#C1E1d4"><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
    	<td width="100" height="540" BGCOLOR="#C1E1d4" valign="top"><FONT FACE="verdana,helvetica,arial" SIZE="2" COLOR="#A90014"><br>
		<!--#include file="../inc/regular/vmenu.asp"--></td>
      	<td width="20"><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	 	<td valign="top"><FONT FACE="verdana,helvetica,arial" SIZE="2" COLOR="#A90014">
<!-------------------------------Sideindhold------------------------------------->
<br>
<br>
<%
Dim antalarbDageA()
Dim totTimerA()
Dim intFakbareTimerA()
Dim omsA()
Dim timerTjTotA()

Dim avrTimerA() 
Dim	avrFaktimerA()

Dim antalarbDageB()
Dim totTimerB()
Dim intFakbareTimerB()
Dim omsB()
Dim timerTjTotB()

Dim avrTimerB() 
Dim	avrFaktimerB()

Public varStrMrdA
Public varStrMrdB

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
Fra den: <FONT FACE="verdana,helvetica,arial" SIZE="2" COLOR="#5881ac">1/14/02<FONT FACE="verdana,helvetica,arial" SIZE="2" COLOR="#A90014"> og til idag den <FONT FACE="verdana,helvetica,arial" SIZE="2" COLOR="#5881ac"><%=strDatodagWrite %>
<br>
<a href="Javascript:history.back()">Tilbage</a>
<br>
<br>

    <%
	Set connection = Server.CreateObject("ADODB.Connection")
	Set objRec = Server.CreateObject ("ADODB.Recordset")
	Set oRec2 = server.createobject("ADODB.Recordset")
		
	connection.Open strConnect
	
	objRec.Open "SELECT Tid, Tdato, Timer, Tfaktim, Tjobnr FROM timer ORDER BY Tdato", connection, 3
	
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
	
	while not objRec.EOF 
		
		
		if objRec("Tfaktim") = "Yes" then
		intFakbareTimer = intFakbareTimer + objRec("timer")
		else
		intFakbareTimer = intFakbareTimer
		end if
		
		totTimer = totTimer + objRec("timer")
		'Slut Årstotal
		
	StrTdato = objRec("Tdato")
	
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
		Call beregnTimerDage(strMrd)
		
	case 2
	
		if strDag < 15 then
		
		 	if objRec("Tdato") <> lastDato then
			lastDato = objRec("Tdato")
			antalarbDageFebA = antalarbDageFebA + 1
			else
			antalarbDageFebA = antalarbDageFebA
			end if
			
			totTimerFebA = totTimerFebA + objRec("timer")
			
				if objRec("Tfaktim") = "Yes" then
					intFakbareTimerFebA = intFakbareTimerFebA + objRec("timer")
					
					strSQL = "SELECT jobnr, jobTpris FROM job WHERE jobnr = '" & objRec("Tjobnr") & "'"
					oRec2.open strSQL, connection, 3
					
					omsFebA = (omsFebA + (oRec2("jobTpris") * objRec("timer")))
					'Response.write omsFebA &", " & oRec2("jobTpris") &", " & objRec("timer") &"<br>" 
					timerTjTotFebA = timerTjTotFebA + objRec("timer")
					oRec2.close
				else
					intFakbareTimerFebA = intFakbareTimerFebA
				end if
			
		else
		
			if objRec("Tdato") <> lastDato then
			lastDato = objRec("Tdato")
			antalarbDageFebB = antalarbDageFebB + 1
			else
			antalarbDageFebB = antalarbDageFebB 
			end if
			totTimerFebB = totTimerFebB + objRec("timer")
			
				if objRec("Tfaktim") = "Yes" then
				
					intFakbareTimerFebB = intFakbareTimerFebB + objRec("timer")
					
					strSQL = "SELECT jobnr, jobTpris FROM job WHERE jobnr = '"& objRec("Tjobnr") &"'"
					oRec2.open strSQL, connection, 3
					
					omsFebB = (omsFebB + (oRec2("jobTpris") * objRec("timer")))
					'Response.write omsFebB &", " & oRec2("jobTpris") &", " & objRec("timer") &"<br>" 
					timerTjTotFebB = timerTjTotFebB + objRec("timer")
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
	
	objRec.movenext
	wend
	
	
	
	
	
	
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
    <td valign="bottom" width="20"><img src="ill/scalag1.gif" width="25" height="<%=10*(left(avrTimerA(varStrMrdA),4))%>" alt="" border="0"></td>
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
<%
 'now close and clean up
	objRec.Close
	connection.close
	Set objRec = Nothing
	set connection = Nothing 

end if 


%>
</body>
</html>
