	<!--#include file="../inc/connection/conn_db_inc.asp"-->
	<!--#include file="../inc/errors/error_inc.asp"-->
	<%
	if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	if len(request("searchstring")) = "0" then
	searchstring = 0
	else
	searchstring = cint(request("searchstring"))
	end if
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="inc/dato.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/calender.asp"-->
	<!--#include file="inc/convertDate.asp"-->

<script> 
<!-- 

function setTimerTot(nummer, dagtype) {
	var varValue;
	var varValue_total;
	
	oprVerdi = (document.all["FM_"+ dagtype +"_opr_" + nummer + ""].value / 1);
	varValue = (document.all["Timer_"+ dagtype +"_" + nummer + ""].value / 1);
	varValue_total = (document.all["FM_"+ dagtype +"_total"].value / 1); 
 	
	if (varValue > 24) {
	alert("En time-indtastning må ikke overstige 24 timer. \n Der er angivet: " + varValue + " timer.")
	document.all["Timer_"+ dagtype +"_" + nummer + ""].value = oprVerdi;
	}
	else {
		varTillaeg = (varValue - oprVerdi);
		varTotal_dag_beg = (varValue_total + varTillaeg);
			
		if (varTotal_dag_beg > 24) {
		alert("Et døgn indeholder kun 24 timer!!")
		document.all["Timer_"+ dagtype +"_" + nummer + ""].value = oprVerdi;
		}
		
		if (varTotal_dag_beg <= 24){
		document.all["FM_"+ dagtype +"_total"].value = varTotal_dag_beg; 
		document.all["FM_"+ dagtype +"_opr_" + nummer + ""].value = varValue;
		sonValue = document.all["FM_son_total"].value/1;
		manValue = document.all["FM_man_total"].value/1;
		tirValue = document.all["FM_tir_total"].value/1;
		onsValue = document.all["FM_ons_total"].value/1;
		torValue = document.all["FM_tor_total"].value/1;
		freValue = document.all["FM_fre_total"].value/1;
		lorValue = document.all["FM_lor_total"].value/1;
		document.all["FM_week_total"].value = (sonValue + manValue + tirValue + onsValue + torValue + freValue + lorValue) 
		}
	}
} 

function isNum(passedVal){
	invalidChars = " /:,;<>abcdefghijklmnopqrstuvwxyzæøå"
	if (passedVal == ""){
	return false
	}
	
	for (i = 0; i < invalidChars.length; i++) {
	badChar = invalidChars.charAt(i)
		if (passedVal.indexOf(badChar, 0) != -1){
		return false
		}
	}
	
		for (i = 0; i < passedVal.length; i++) {
			if (passedVal.charAt(i) == ".") {
			return true
			}
			else {
				if (passedVal.charAt(i) < "0") {
				return false
				}
					if (passedVal.charAt(i) > "9") {
					return false
					}
				}
			return true
			} 
			
}
	
function validZip(inZip){
	if (inZip == "") {
	return true
	}
	if (isNum(inZip)){
	return true
	}
	return false
}

function tjektimer(dagtype, nummer){
	if (!validZip(document.all["Timer_"+ dagtype +"_" + nummer + ""].value)){
	alert("Der er angivet et ugyldigt tegn.")
	document.all["Timer_"+ dagtype +"_" + nummer + ""].value = oprVerdi;
	document.all["Timer_"+ dagtype +"_" + nummer + ""].focus()
	document.all["Timer_"+ dagtype +"_" + nummer + ""].select()
	return false
	}
return true
}

--> 
</script>

<%

function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
end function

'***** Funktion til at vise eksterne timer *** Kaldes i linje ?? (langt nede)
function eksterneTimer()

					
					varTjDatoUS = cdate(convertDate(tjekdag(1)&"/"&strRegaar))
					varTjDatoUS_son = convertDateYMD(varTjDatoUS)
					
					varTjDatoUS = cdate(convertDate(tjekdag(2)&"/"&strRegaar))
					varTjDatoUS_man = convertDateYMD(varTjDatoUS)
					
					varTjDatoUS = cdate(convertDate(tjekdag(3)&"/"&strRegaar))
					varTjDatoUS_tir = convertDateYMD(varTjDatoUS)
					
					varTjDatoUS = cdate(convertDate(tjekdag(4)&"/"&strRegaar))
					varTjDatoUS_ons = convertDateYMD(varTjDatoUS) 
					
					varTjDatoUS = cdate(convertDate(tjekdag(5)&"/"&strRegaar))
					varTjDatoUS_tor = convertDateYMD(varTjDatoUS) 
					
					varTjDatoUS = cdate(convertDate(tjekdag(6)&"/"&strRegaar)) 
					varTjDatoUS_fre = convertDateYMD(varTjDatoUS)
					
					varTjDatoUS = cdate(convertDate(tjekdag(7)&"/"&strRegaar)) 
					varTjDatoUS_lor = convertDateYMD(varTjDatoUS)
					
					'** Hvis der findes en faktura på det pågældende job ***** 
					strSQLtimer = "SELECT timer.timer, timer.tdato, timer.tjobnr, timer.tidspunkt AS ttidspkt FROM timer "_
						&" WHERE "_
						&" TAktivitetId = "& oRec("akt") &" AND Tmnavn = '"&Session("user")&"' AND Tdato = '"& varTjDatoUS_son &"'"_
						&" OR "_
						&" TAktivitetId = "& oRec("akt") &" AND Tmnavn = '"&Session("user")&"' AND Tdato = '"& varTjDatoUS_man &"'"_
						&" OR "_
						&" TAktivitetId = "& oRec("akt") &" AND Tmnavn = '"&Session("user")&"' AND Tdato = '"& varTjDatoUS_tir &"'"_
						&" OR "_
						&" TAktivitetId = "& oRec("akt") &" AND Tmnavn = '"&Session("user")&"' AND Tdato = '"& varTjDatoUS_ons &"'"_
						&" OR "_
						&" TAktivitetId = "& oRec("akt") &" AND Tmnavn = '"&Session("user")&"' AND Tdato = '"& varTjDatoUS_tor &"'"_
						&" OR "_
						&" TAktivitetId = "& oRec("akt") &" AND Tmnavn = '"&Session("user")&"' AND Tdato = '"& varTjDatoUS_fre &"'"_
						&" OR "_
						&" TAktivitetId = "& oRec("akt") &" AND Tmnavn = '"&Session("user")&"' AND Tdato = '"& varTjDatoUS_lor &"'"
					
					oRec2.Open strSQLtimer, oConn, 0, 1  
					while not oRec2.EOF 
					dtimeTtidspkt = oRec2("ttidspkt")
					
						select case DatePart("w", oRec2("tdato")) 
						case 1	
						sonTimerVal = SQLBless(oRec2("timer"))
						case 2
						manTimerVal = SQLBless(oRec2("timer"))
						case 3
						tirTimerVal = SQLBless(oRec2("timer"))
						case 4
						onsTimerVal = SQLBless(oRec2("timer"))
						case 5
						torTimerVal = SQLBless(oRec2("timer"))
						case 6
						freTimerVal = SQLBless(oRec2("timer"))
						case 7
						lorTimerVal = SQLBless(oRec2("timer"))
						end select
						
						
					oRec2.movenext
					wend
					oRec2.close
					
					if len(dtimeTtidspkt) <> 0 then
					dtimeTtidspkt = dtimeTtidspkt
					else
					dtimeTtidspkt = "00:00:00"
					end if
					
					if len(sonTimerVal) <> "0" then
					sonTimerTot = sonTimerTot + sonTimerVal
					end if
					
					if len(manTimerVal) <> "0" then
					manTimerTot = manTimerTot + manTimerVal
					end if
					
					if len(tirTimerVal) <> "0" then
					tirTimerTot = tirTimerTot + tirTimerVal
					end if
					
					if len(onsTimerVal) <> "0" then
					onsTimerTot = onsTimerTot + onsTimerVal
					end if
					
					if len(torTimerVal) <> "0" then
					torTimerTot = torTimerTot + torTimerVal
					end if
					
					if len(freTimerVal) <> "0" then
					freTimerTot = freTimerTot + freTimerVal
					end if
					
					if len(lorTimerVal) <> "0" then
					lorTimerTot = lorTimerTot + lorTimerVal
					end if
					
end function

%>
	<div id="sindhold" style="position:absolute; left:190; top:50; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="550">
	<tr>
    <td valign="top">
			 <!-------------------------------Sideindhold------------------------------------->
			<br>
	<%
	'*** Følgende bruges også i rmenu (kalenderen) ****
	%>
	
<img src="../ill/header_reg.gif" alt="" border="0"><hr align="left" width="750" size="1" color="#000000" noshade>

		
		<%
		 
		f = 0
		Redim brugergruppeId(f)
		
		strSQL = "SELECT Mnavn, Mid, type, ProjektgruppeId, MedarbejderId FROM medarbejdere, medarbejdertyper, progrupperelationer WHERE medarbejdere.Mid="& session("Mid")&" AND medarbejdertyper.id = medarbejdere.Medarbejdertype AND MedarbejderId = Mid" 
			oRec.Open strSQL, oConn, 0, 1 
			
			While Not oRec.EOF
			Redim preserve brugergruppeId(f) 
				strMnavn = oRec("Mnavn")
				strMnr = oRec("Mid")
				strMType = oRec("type")
				brugergruppeId(f) = oRec("ProjektgruppeId")
				f = f + 1
				oRec.MoveNext
			Wend
		
		oRec.Close

%>
	</td>
</tr>
</table>
<br>



<table cellspacing="0" cellpadding="0" border="0" width="530">
<tr bgcolor="#5582D2">
	<td width="3"><img src="../ill/venstre_hjorne.gif" alt="" border="0"></td>
	<td width="524" style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid;"><font class="stor-hvid">&nbsp;Eksterne job</font></td>
	<td  width="3" align="right"><img src="../ill/hojre_hjorne.gif" alt="" border="0"></td>
</tr>
</table>

<table cellspacing="0" cellpadding="0" border="0" width="530">
<form action="timereg_db.asp" method="POST"> 
<input type="Hidden" name="FM_mnr" value="<%=strMnr%>">
<%
sonTimerTot = 0
manTimerTot = 0
tirTimerTot = 0
onsTimerTot = 0
torTimerTot = 0
freTimerTot = 0
lorTimerTot = 0
areJobPrinted = 0

'** Sætter værdier på de variable der skal bruges for at finde de datoer der skal vises
strDtype = left(request("regdtype"), 3)
if len(strDtype) <> 0 then

select case request("regmrd")
case "1"
strRegAar = request("regaar")
strRegMrd = request("regmrd")
strRegMrdKor = "13"
case "12"
strRegAar = request("regaar")
strRegMrd = request("regmrd")
strRegMrdKor = "12"
case else 
strRegAar = request("regaar")
strRegMrd = request("regmrd")
strRegMrdKor = request("regmrd")
end select

strWeek =  datepart("ww", strRegMrd &"/" & request("regdag")&"/" & strRegAar)

select case strDtype
	case "son"
	sondag = request("regdag") &"/"&strRegMrd
	mandag = request("regdag") + 1 &"/"&strRegMrd
	tirsdag = request("regdag") + 2 &"/"&strRegMrd
	onsdag = request("regdag") + 3 &"/"&strRegMrd
	torsdag = request("regdag") + 4 &"/"&strRegMrd
	fredag = request("regdag") + 5 &"/"&strRegMrd
	lordag = request("regdag") + 6 &"/"&strRegMrd
	case "man"
	mandag = request("regdag")&"/"&strRegMrd
	sondag = request("regdag") - 1 &"/"&strRegMrd
	tirsdag = request("regdag") + 1 &"/"&strRegMrd
	onsdag = request("regdag") + 2 &"/"&strRegMrd
	torsdag = request("regdag") + 3 &"/"&strRegMrd
	fredag = request("regdag") + 4 &"/"&strRegMrd
	lordag = request("regdag") + 5 &"/"&strRegMrd
	case "tir"
	tirsdag = request("regdag") &"/"&strRegMrd
	sondag = request("regdag") - 2 &"/"&strRegMrd
	mandag = request("regdag") - 1 &"/"&strRegMrd
	onsdag = request("regdag") + 1 &"/"&strRegMrd
	torsdag = request("regdag") + 2 &"/"&strRegMrd
	fredag = request("regdag") + 3 &"/"&strRegMrd
	lordag = request("regdag") + 4 &"/"&strRegMrd
	case "ons"
	onsdag = request("regdag") &"/"&strRegMrd
	sondag = request("regdag") - 3 &"/"&strRegMrd
	mandag = request("regdag") - 2 &"/"&strRegMrd
	tirsdag = request("regdag") - 1 &"/"&strRegMrd
	torsdag = request("regdag") + 1 &"/"&strRegMrd
	fredag = request("regdag") + 2 &"/"&strRegMrd
	lordag = request("regdag") + 3 &"/"&strRegMrd
	case "tor"
	torsdag = request("regdag") &"/"&strRegMrd
	sondag = request("regdag") - 4 &"/"&strRegMrd 
	mandag = request("regdag") - 3 &"/"&strRegMrd
	tirsdag = request("regdag") - 2 &"/"&strRegMrd
	onsdag = request("regdag") - 1 &"/"&strRegMrd
	fredag = request("regdag") + 1 &"/"&strRegMrd
	lordag = request("regdag") + 2 &"/"&strRegMrd
	case "fre"
	fredag = request("regdag") &"/"&strRegMrd
	sondag = request("regdag") - 5 &"/"&strRegMrd
	mandag = request("regdag") - 4 &"/"&strRegMrd
	tirsdag = request("regdag") - 3 &"/"&strRegMrd
	onsdag = request("regdag") - 2 &"/"&strRegMrd
	torsdag = request("regdag") - 1 &"/"&strRegMrd
	lordag = request("regdag") + 1 &"/"&strRegMrd
	case "lor"
	lordag = request("regdag") &"/"&strRegMrd
	sondag = request("regdag") - 6 &"/"&strRegMrd
	mandag = request("regdag") - 5 &"/"&strRegMrd  
	tirsdag = request("regdag") - 4 &"/"&strRegMrd
	onsdag = request("regdag") - 3 &"/"&strRegMrd
	torsdag = request("regdag") - 2 &"/"&strRegMrd
	fredag = request("regdag") - 1 &"/"&strRegMrd
	end select
else
	strRegAar = year(now)
	strRegMrd = month(now)
	strRegMrdKor = month(now) 
	
	strWeek =  datepart("ww", strRegMrd &"/" & day(now)&"/" & strRegAar)
	strDtypeToday = weekday(""&month(now)&"/"&day(now)&"/"&year(now)&"")
	
	select case strDtypeToday
	case 1
	sondag = day(now)&"/"&strRegMrd 
	mandag = day(now) + 1 &"/"&strRegMrd
	tirsdag = day(now) + 2 &"/"&strRegMrd
	onsdag = day(now) + 3 &"/"&strRegMrd
	torsdag = day(now) + 4 &"/"&strRegMrd
	fredag = day(now) + 5 &"/"&strRegMrd
	lordag = day(now) + 6 &"/"&strRegMrd
	case 2
	mandag = day(now)&"/"&strRegMrd
	sondag = day(now) - 1 &"/"&strRegMrd
	tirsdag = day(now) + 1 &"/"&strRegMrd
	onsdag = day(now) + 2 &"/"&strRegMrd
	torsdag = day(now) + 3 &"/"&strRegMrd
	fredag = day(now) + 4 &"/"&strRegMrd
	lordag = day(now) + 5 &"/"&strRegMrd
	case 3
	tirsdag = day(now) &"/"&strRegMrd
	sondag = day(now) - 2 &"/"&strRegMrd
	mandag = day(now) - 1 &"/"&strRegMrd
	onsdag = day(now) + 1 &"/"&strRegMrd
	torsdag = day(now) + 2 &"/"&strRegMrd
	fredag = day(now) + 3 &"/"&strRegMrd
	lordag = day(now) + 4 &"/"&strRegMrd
	case 4
	onsdag = day(now) &"/"&strRegMrd
	sondag = day(now) - 3 &"/"&strRegMrd
	mandag = day(now) - 2 &"/"&strRegMrd
	tirsdag = day(now) - 1 &"/"&strRegMrd
	torsdag = day(now) + 1 &"/"&strRegMrd
	fredag = day(now) + 2 &"/"&strRegMrd
	lordag = day(now) + 3 &"/"&strRegMrd
	case 5
	torsdag = day(now) &"/"&strRegMrd
	sondag = day(now) - 4 &"/"&strRegMrd 
	mandag = day(now) - 3 &"/"&strRegMrd
	tirsdag = day(now) - 2 &"/"&strRegMrd
	onsdag = day(now) - 1 &"/"&strRegMrd
	fredag = day(now) + 1 &"/"&strRegMrd
	lordag = day(now) + 2 &"/"&strRegMrd
	case 6
	fredag = day(now) &"/"&strRegMrd
	sondag = day(now) - 5 &"/"&strRegMrd
	mandag = day(now) - 4 &"/"&strRegMrd
	tirsdag = day(now) - 3 &"/"&strRegMrd
	onsdag = day(now) - 2 &"/"&strRegMrd
	torsdag = day(now) - 1 &"/"&strRegMrd
	lordag = day(now) + 1 &"/"&strRegMrd
	case 7
	lordag = day(now) &"/"&strRegMrd
	sondag = day(now) - 6 &"/"&strRegMrd
	mandag = day(now) - 5 &"/"&strRegMrd  
	tirsdag = day(now) - 4 &"/"&strRegMrd
	onsdag = day(now) - 3 &"/"&strRegMrd
	torsdag = day(now) - 2 &"/"&strRegMrd
	fredag = day(now) - 1 &"/"&strRegMrd
	end select
	
end if
	
	Redim tjekdag(7)
	tjekdag(1) = sondag
	tjekdag(2) = mandag
	tjekdag(3) = tirsdag
	tjekdag(4) = onsdag
	tjekdag(5) = torsdag
	tjekdag(6) = fredag
	tjekdag(7) = lordag
	
	Redim writedagnavn(7)
	writedagnavn(1) = "Søn"
	writedagnavn(2) = "Man"
	writedagnavn(3) = "Tir"
	writedagnavn(4) = "Ons"
	writedagnavn(5) = "Tor"
	writedagnavn(6) = "Fre"
	writedagnavn(7) = "Lør"
	
	t = 1
	
%>
<tr>
<td align="center" style="!background-color : #8DA7E5; border-left: 1px solid #003399; border-right: 1px solid #003399; width : 138px;">
<font style="color:#ffffff;font-family:arial,sans-serif;font-size:11px;"><b>Uge: 
<%
Response.write strWeek & "&nbsp;"
if len(request("regaar")) = 0 then
		Response.write year(now)
		else
		Response.write request("regaar")
		end if%>
</b></td>
<%
'*** Gennemløber de 7 dage for at finde deres dato. ***
	for t = 1 to 7
		intStreg = instr(tjekdag(t), "/")
		select case intStreg
		case 2
				if day(now)&"/"&month(now)&"/"&year(now) =  tjekdag(t)&"/" &year(now)  then
				borderColor = "#D6DFF5"
				border = 2
				else
				borderColor = "#FFFFE1"
				border = 1
				end if
		case 3 
				if day(now)&"/"&month(now)&"/"&year(now) =  tjekdag(t)&"/" &year(now)  then
				borderColor = "#D6DFF5"
				border = 2
				else
				borderColor = "#FFFFE1"
				border = 1
				end if
			
		if cint(left(tjekdag(t), 2)) > cint(daysinMonth) then
			if lastTjekdagKor = "y" then
			tjekdag(t) = left(lastTjekDag, 1) + 1& "/"&(strRegMrd + 1)
			else
			tjekdag(t) = "1/"&(strRegMrd + 1)
			end if
			lastTjekdagKor = "y"
		else
		tjekdag(t) = tjekdag(t)
		end if
		end select
		lastTjekDag = tjekdag(t)
		
		'*** Hvis dag er fra forrige mrd
		intStreg = instr(tjekdag(t), "/")
		select case intStreg
		case 2
		varNegativ = left(tjekdag(t), 1)
		case 3
		varNegativ = left(tjekdag(t), 2)
		end select
		
		select case varNegativ
		case 0
		tjekdag(t) = daysinMonthForrige&"/"&(strRegMrdKor -1)
		case "-1"
		tjekdag(t) = (daysinMonthForrige -1)&"/"&(strRegMrdKor -1)
		case "-2"
		tjekdag(t) = (daysinMonthForrige -2)&"/"&(strRegMrdKor -1)
		case "-3"
		tjekdag(t) = (daysinMonthForrige -3)&"/"&(strRegMrdKor -1)
		case "-4"
		tjekdag(t) = (daysinMonthForrige -4)&"/"&(strRegMrdKor -1)
		case "-5"
		tjekdag(t) = (daysinMonthForrige -5)&"/"&(strRegMrdKor -1)
		case "-6"
		tjekdag(t) = (daysinMonthForrige -6)&"/"&(strRegMrdKor -1)
		case "-7"
		tjekdag(t) = (daysinMonthForrige -7)&"/"&(strRegMrdKor -1)
		case else
		tjekdag(t) = tjekdag(t)
		end select
		
		%>
<td height="25" align="center" style="background-color: <%=borderColor%>; border-right: 1px solid #003399; width : 56px;">
<font style="color:#000000; font-family:arial,sans-serif; font-size:11px;"><%=writedagnavn(t)%>&nbsp;<%=tjekdag(t)%></td>
		<%
	next
	'*** Gennenløb slut ****%>
	</tr>
	</table>
	
<script language=javascript>
 if (document.images){
		plus = new Image(200, 200);
		plus.src = "ill/plus.gif";
		minus = new Image(200, 200);
		minus.src = "ill/minus2.gif";
		}

		function expand (de){
			if (document.all("Menu" + de)){
				if (document.all("Menu" + de).style.display == "none"){
				document.all("Menu" + de).style.display = "";
				document.images["Menub" + de].src = minus.src;
			}else{
				document.all("Menu" + de).style.display = "none";
				document.images["Menub" + de].src = plus.src;
			}
		}
	}
</script>
 <script language=javascript>
 function antalakt(jid){
 newval = document.all["FM_akt_" + jid].value;
 document.all["FM_antal_akt_" + jid].value = newval;
 }
 </script>
	
<table cellspacing="0" cellpadding="0" border="0" width="530">
<input type="hidden" name="year" value="<%=strRegAar%>">
<input type="hidden" name="datoSon" value="<%=tjekdag(1)%>">
<input type="hidden" name="datoMan" value="<%=tjekdag(2)%>">
<input type="hidden" name="datoTir" value="<%=tjekdag(3)%>">
<input type="hidden" name="datoOns" value="<%=tjekdag(4)%>">
<input type="hidden" name="datoTor" value="<%=tjekdag(5)%>">
<input type="hidden" name="datoFre" value="<%=tjekdag(6)%>">
<input type="hidden" name="datoLor" value="<%=tjekdag(7)%>">

<% 
 '** Sætter jobnr på det sidst udskrevne job. ***
  areJobPrinted = 0
  strLastjobnr = 0
  jobPrinted = "0#"
  aktPrinted = "0#"
  
  '** Array nr. ***
  de = 1

 if searchstring = 0 then
 varSelectedJob = ""
 else
 varSelectedJob = " jobknr = "& searchstring &" AND " 
 end if	
  
  '**************************** Rettigheds tjeck **************************
  	for intcounter = 0 to f - 1  
  
  	strSQLkri = strSQLkri &" "& varSelectedJob &" jobstatus = 1 AND fakturerbart = 1 AND kunder.Kid = job.jobknr AND job.projektgruppe1 = "&brugergruppeId(intcounter)&""_
	&" OR "_
	&" "& varSelectedJob &" jobstatus = 1 AND fakturerbart = 1 AND kunder.Kid = job.jobknr AND job.projektgruppe2 = "&brugergruppeId(intcounter)&""_
	&" OR "_
	&" "& varSelectedJob &" jobstatus = 1 AND fakturerbart = 1 AND kunder.Kid = job.jobknr AND job.projektgruppe3 = "&brugergruppeId(intcounter)&""_
	&" OR "_
	&" "& varSelectedJob &" jobstatus = 1 AND fakturerbart = 1 AND kunder.Kid = job.jobknr AND job.projektgruppe4 = "&brugergruppeId(intcounter)&""_
	&" OR "_
  	&" "& varSelectedJob &" jobstatus = 1 AND fakturerbart = 1 AND kunder.Kid = job.jobknr AND job.projektgruppe5 = "&brugergruppeId(intcounter)&""_
	&" OR "
	
	next
  	
	'** Trimmer de 2 sql states ***
	strSQLkri_len = len(strSQLkri)
	strSQLkri_left = strSQLkri_len - 3
	strSQLkri_use = left(strSQLkri, strSQLkri_left) 
  	strSQLkri = strSQLkri_use
  	
	'**************************** Rettigheds tjeck aktiviteter **************************
  	for intcounter = 0 to f - 1  
  
	strSQLkri3 = strSQLkri3 &" aktiviteter.job = job.id AND aktiviteter.projektgruppe1 = "& brugergruppeId(intcounter) &""_
	&" OR aktiviteter.job = job.id AND aktiviteter.projektgruppe2 = "& brugergruppeId(intcounter) &""_
	&" OR aktiviteter.job = job.id AND aktiviteter.projektgruppe3 = "& brugergruppeId(intcounter) &""_
	&" OR aktiviteter.job = job.id AND aktiviteter.projektgruppe4 = "& brugergruppeId(intcounter) &""_
	&" OR aktiviteter.job = job.id AND aktiviteter.projektgruppe5 = "& brugergruppeId(intcounter) &" OR "
	
	next
  	
	'** Trimmer sql states ***
	strSQLkri3_len = len(strSQLkri3)
	strSQLkri3_left = strSQLkri3_len - 3
	strSQLkri3_use = left(strSQLkri3, strSQLkri3_left) 
  	strSQLkri3 = strSQLkri3_use	
	
	varTjDatoUS_start = cdate(convertDate(tjekdag(1)&"/"&strRegaar))
	use_varTjDatoUS_start = convertDateYMD(varTjDatoUS_start)
	'Response.write use_varTjDatoUS_start
	
	'********* Adgang til job *****
	strSQL = "SELECT job.id AS id, jobnr, jobnavn, jobknr, kkundenavn, aktiviteter.id AS akt, aktiviteter.navn AS navn, aktiviteter.job, COUNT(aktiviteter.id) AS antal, fakdato, f.tidspunkt AS ftidspkt FROM job, kunder LEFT JOIN aktiviteter ON ("& strSQLkri3 &") LEFT JOIN fakturaer f ON (f.jobid = job.id AND fakdato >= '"& use_varTjDatoUS_start &"') WHERE " & strSQLkri & " GROUP BY jobnavn, navn" 
	oRec.cachesize = 100
	oRec.Open strSQL, oConn, 3 
	
	lastJobid = 0
	
	While Not oRec.EOF
	lastfakdato = oRec("fakdato")
	
	sonTimerVal = ""
	manTimerVal = ""
	tirTimerVal = ""
	onsTimerVal = ""
	torTimerVal = ""
	freTimerVal = ""
	lorTimerVal = ""
	
	if lastJobid <> oRec("id") AND lastJobid > 0 then				
		Response.write "</table></div></td></tr>"
	end if	
					
				
					if lastJobid <> oRec("id") then
					intAkt = oRec("antal")
					
					Response.write "<tr><td style='!width : 30px; background-color: #5582D2; padding-left : 4px; border-left: 1px solid #003399'><font class='lillehvid'><b>"_
					& oRec("Jobnr") &"</b></td><td height=20 style='!width : 500px;  background-color: #5582D2; padding-left : 4px;padding-right : 4px; border-right: 1px solid #003399'>"
					'Response.write "<font class='storhvid'><b>("&intAkt&")&nbsp;</font><b>"
					Response.write "<font class='storhvid'>"
					'** Er der nogen aktiviteter **
					if intAkt > "0" then%>
					<a href="javascript:expand('<%=de%>');"><img src="ill/plus.gif" width="9" height="9" border="0" name="Menub<%=de%>"></a>&nbsp;
					<a href="javascript:expand('<%=de%>');" class='alt'>
					<%
					Response.write oRec("Jobnavn")&"</a></b>&nbsp;&nbsp;&nbsp;<font class='lilleblaa'>("&oRec("kkundenavn")&")</font></td></tr>"
					%>
					<tr style="border-left: 1px solid #003399; border-right: 1px solid #003399">
					<td colspan="2">
					<DIV ID="Menu<%=de%>" Style="position: relative; display: none;">
					<table cellspacing="0" cellpadding="0" border="0" width="530" bgcolor="#ffffff">
					<%else
					Response.write "<img src='ill/blank.gif' width='18' height='0' border='0' alt=''>&nbsp;<font class='storhvid'><b>"&oRec("Jobnavn")&"</b>&nbsp;&nbsp;&nbsp;<font class='lilleblaa'>("&oRec("kkundenavn")&")</font></td></tr>"
					%>
					<tr style=" border-left: 1px solid #003399; border-right: 1px solid #003399">
					<td colspan="2"><DIV ID="Menu_<%=oRec("jobnavn")%>" Style="position: relative; display: none;">
					<table cellspacing="0" cellpadding="0" border="0" width="530">
					<%
					end if
				end if
						if intAkt > "0" then		
						%>
							<input type="hidden" name="FM_jobnr_<%=de%>" value="<%=oRec("jobnr")%>">
							<input type="hidden" name="FM_Aid_<%=de%>" value="<%=oRec("akt")%>">
						<%
						call eksterneTimer()
						
						'** Sætter farve på indtastningfelt efter om der er udskrevet en faktura eller ej ***
						if len(lastfakdato) <> 0 AND lastfakdato => varTjDatoUS_start then
								
									select case DatePart("w", lastfakdato)
									case 1
									if formatdatetime(dtimeTtidspkt, 3) < formatdatetime(oRec("ftidspkt"), 3) AND len(sonTimerVal) <> 0 then
									fakbgcol_son = "limegreen"
									maxl_son = 0
									fmbg_son = "#cccccc"
									else
									fakbgcol_son = "#cd853f"
									maxl_son = 4
									fmbg_son = "#FFDFDF"
									end if 
									
									
									fakbgcol_man = "#7F9DB9"	
									fakbgcol_tir = "#7F9DB9"	
									fakbgcol_ons = "#7F9DB9"	
									fakbgcol_tor = "#7F9DB9"	
									fakbgcol_fre = "#7F9DB9"	
									fakbgcol_lor = "#cd853f"
						
									maxl_man = 4
									maxl_tir = 4
									maxl_ons = 4
									maxl_tor = 4
									maxl_fre = 4
									maxl_lor = 4
									
									fmbg_man = "#FFFFFF" 
									fmbg_tir = "#FFFFFF" 
									fmbg_ons = "#FFFFFF" 
									fmbg_tor = "#FFFFFF" 
									fmbg_fre = "#FFFFFF" 
									fmbg_lor = "#FFDFDF"  
									
									
									
									case 2
									if formatdatetime(dtimeTtidspkt, 3) < formatdatetime(oRec("ftidspkt"), 3) AND len(manTimerVal) <> 0 then
									fakbgcol_man = "limegreen"
									maxl_man = 0
									fmbg_man = "#cccccc"
									else
									fakbgcol_man = "#7F9DB9"
									maxl_man = 4
									fmbg_man = "#FFFFFF"
									end if 
									
									
									fakbgcol_son = "limegreen"	
									fakbgcol_tir = "#7F9DB9"	
									fakbgcol_ons = "#7F9DB9"	
									fakbgcol_tor = "#7F9DB9"	
									fakbgcol_fre = "#7F9DB9"	
									fakbgcol_lor = "#cd853f"
						
									maxl_son = 0
									maxl_tir = 4
									maxl_ons = 4
									maxl_tor = 4
									maxl_fre = 4
									maxl_lor = 4
									
									fmbg_son = "#cccccc" 
									fmbg_tir = "#FFFFFF" 
									fmbg_ons = "#FFFFFF" 
									fmbg_tor = "#FFFFFF" 
									fmbg_fre = "#FFFFFF" 
									fmbg_lor = "#FFDFDF"  
									 
									case 3
									if formatdatetime(dtimeTtidspkt, 3) < formatdatetime(oRec("ftidspkt"), 3) AND len(tirTimerVal) <> 0 then
									fakbgcol_tir = "limegreen"
									maxl_tir = 0
									fmbg_tir = "#cccccc"
									else
									fakbgcol_tir = "#7F9DB9"
									maxl_tir = 4
									fmbg_tir = "#FFFFFF"
									end if 
									
									
									fakbgcol_son = "limegreen"	
									fakbgcol_man = "limegreen"	
									fakbgcol_ons = "#7F9DB9"	
									fakbgcol_tor = "#7F9DB9"	
									fakbgcol_fre = "#7F9DB9"	
									fakbgcol_lor = "#cd853f"
						
									maxl_son = 0
									maxl_man = 0
									maxl_ons = 4
									maxl_tor = 4
									maxl_fre = 4
									maxl_lor = 4
									
									fmbg_son = "#cccccc" 
									fmbg_man = "#cccccc" 
									fmbg_ons = "#FFFFFF" 
									fmbg_tor = "#FFFFFF" 
									fmbg_fre = "#FFFFFF" 
									fmbg_lor = "#FFDFDF"  
									 
									case 4
									if formatdatetime(dtimeTtidspkt, 3) < formatdatetime(oRec("ftidspkt"), 3) AND len(onsTimerVal) <> 0 then
									fakbgcol_ons = "limegreen"
									maxl_ons = 0
									fmbg_ons = "#cccccc"
									else
									fakbgcol_ons = "#7F9DB9"
									maxl_ons = 4
									fmbg_ons = "#FFFFFF"
									end if 
									
									
									fakbgcol_son = "limegreen"	
									fakbgcol_man = "limegreen"	
									fakbgcol_tir = "limegreen"	
									fakbgcol_tor = "#7F9DB9"	
									fakbgcol_fre = "#7F9DB9"	
									fakbgcol_lor = "#cd853f"
						
									maxl_son = 0
									maxl_man = 0
									maxl_tir = 0
									maxl_tor = 4
									maxl_fre = 4
									maxl_lor = 4
									
									fmbg_son = "#cccccc" 
									fmbg_man = "#cccccc" 
									fmbg_tir = "#cccccc" 
									fmbg_tor = "#FFFFFF" 
									fmbg_fre = "#FFFFFF" 
									fmbg_lor = "#FFDFDF"  
									
									 
									case 5
									Response.write "tval:"& torTimerVal
									if formatdatetime(dtimeTtidspkt, 3) < formatdatetime(oRec("ftidspkt"), 3) AND len(torTimerVal) <> 0 then
									fakbgcol_tor = "limegreen"
									maxl_tor = 0
									fmbg_tor = "#cccccc"
									else
									fakbgcol_tor = "#7F9DB9"
									maxl_tor = 4
									fmbg_tor = "#FFFFFF"
									end if
									
									
									fakbgcol_son = "limegreen"
									fakbgcol_man = "limegreen"	
									fakbgcol_tir = "limegreen"	
									fakbgcol_ons = "limegreen"	
									fakbgcol_fre = "#7F9DB9"	
									fakbgcol_lor = "#cd853f"
						
									maxl_son = 0
									maxl_man = 0
									maxl_tir = 0
									maxl_ons = 0
									maxl_fre = 4
									maxl_lor = 4
									
									fmbg_son = "#cccccc"
									fmbg_tor = "#cccccc"  
									fmbg_tir = "#cccccc" 
									fmbg_ons = "#cccccc" 
									fmbg_fre = "#FFFFFF" 
									fmbg_lor = "#FFDFDF"   
									case 6
									if formatdatetime(dtimeTtidspkt, 3) < formatdatetime(oRec("ftidspkt"), 3) AND len(freTimerVal) <> 0 then
									fakbgcol_fre = "limegreen"
									maxl_fre = 0
									fmbg_fre = "#cccccc"
									else
									fakbgcol_fre = "#7F9DB9"
									maxl_fre = 4
									fmbg_fre = "#FFFFFF"
									end if
									
									
									fakbgcol_son = "limegreen"
									fakbgcol_man = "limegreen"	
									fakbgcol_tir = "limegreen"	
									fakbgcol_ons = "limegreen"	
									fakbgcol_tor = "limegreen"	
									fakbgcol_lor = "#cd853f"
						
									maxl_son = 0
									maxl_man = 0
									maxl_tir = 0
									maxl_ons = 0
									maxl_tor = 0
									maxl_lor = 4
									
									fmbg_son = "#cccccc"
									fmbg_man = "#cccccc"  
									fmbg_tir = "#cccccc" 
									fmbg_ons = "#cccccc" 
									fmbg_tor = "#cccccc" 
									fmbg_lor = "#FFDFDF" 
									
									
									case 7
									if formatdatetime(dtimeTtidspkt, 3) < formatdatetime(oRec("ftidspkt"), 3) AND len(lorTimerVal) <> 0 then
									fakbgcol_lor = "limegreen"
									maxl_lor = 0
									fmbg_lor = "#cccccc"
									else
									fakbgcol_lor = "#7F9DB9"
									maxl_lor = 4
									fmbg_lor = "#FFFFFF"
									end if
								
									
									fakbgcol_son = "limegreen"
									fakbgcol_man = "limegreen"	
									fakbgcol_tir = "limegreen"	
									fakbgcol_ons = "limegreen"	
									fakbgcol_tor = "limegreen"	
									fakbgcol_fre = "limegreen"
						
									maxl_son = 0
									maxl_man = 0
									maxl_tir = 0
									maxl_ons = 0
									maxl_tor = 0
									maxl_fre = 0
									
									fmbg_son = "#cccccc"
									fmbg_man = "#cccccc"  
									fmbg_tir = "#cccccc" 
									fmbg_ons = "#cccccc" 
									fmbg_tor = "#cccccc" 
									fmbg_fre = "#cccccc"  
									end select
								'end if
							else
									fakbgcol_son = "#cd853f"
									fakbgcol_man = "#7F9DB9"	
									fakbgcol_tir = "#7F9DB9"	
									fakbgcol_ons = "#7F9DB9"	
									fakbgcol_tor = "#7F9DB9"	
									fakbgcol_fre = "#7F9DB9"	
									fakbgcol_lor = "#cd853f"
									
									maxl_son = 4
									maxl_man = 4
									maxl_tir = 4
									maxl_ons = 4
									maxl_tor = 4
									maxl_fre = 4
									maxl_lor = 4
									
									fmbg_son = "#FFDFDF" 
									fmbg_man = "#FFFFFF" 
									fmbg_tir = "#FFFFFF" 
									fmbg_ons = "#FFFFFF" 
									fmbg_tor = "#FFFFFF" 
									fmbg_fre = "#FFFFFF" 
									fmbg_lor = "#FFDFDF"   
							end if
						
						%>
						<tr style='background-color: #D6DFF5;'>
							<td style='!width : 138px; padding-left : 0px; padding-right : 4px; border-left: 1px solid #003399;'>&nbsp;<%=left(oRec("navn"), 20)%></td>	
							<td align="center" style="width : 56px;"><input type="hidden" name="FM_son_opr_<%=de%>" value="<%=sonTimerVal%>"><input type="Text" name="Timer_son_<%=de%>" maxlength="<%=maxl_son%>" value="<%=sonTimerVal%>"  onkeyup="setTimerTot(<%=de%>, 'son'), tjektimer('son',<%=de%>)"; style="!border: 1px; background-color: <%=fmbg_son%>; border-color: <%=fakbgcol_son%>; border-style: solid; width : 56px;"></td>
							<td align="center" style="width : 56px;"><input type="hidden" name="FM_man_opr_<%=de%>" value="<%=manTimerVal%>"><input type="Text" name="Timer_man_<%=de%>" maxlength="<%=maxl_man%>" value="<%=manTimerVal%>"  onkeyup="setTimerTot(<%=de%>, 'man'), tjektimer('man',<%=de%>)"; style="!border: 1px; background-color: <%=fmbg_man%>; border-color: <%=fakbgcol_man%>; border-style: solid; width : 56px;"></td>
							<td align="center" style="width : 56px;"><input type="hidden" name="FM_tir_opr_<%=de%>" value="<%=tirTimerVal%>"><input type="Text" name="Timer_tir_<%=de%>" maxlength="<%=maxl_tir%>" value="<%=tirTimerVal%>"  onkeyup="setTimerTot(<%=de%>, 'tir'), tjektimer('tir',<%=de%>)"; style="!border: 1px; background-color: <%=fmbg_tir%>; border-color: <%=fakbgcol_tir%>; border-style: solid; width : 56px;"></td>
							<td align="center" style="width : 56px;"><input type="hidden" name="FM_ons_opr_<%=de%>" value="<%=onsTimerVal%>"><input type="Text" name="Timer_ons_<%=de%>" maxlength="<%=maxl_ons%>" value="<%=onsTimerVal%>"  onkeyup="setTimerTot(<%=de%>, 'ons'), tjektimer('ons',<%=de%>)"; style="!border: 1px; background-color: <%=fmbg_ons%>; border-color: <%=fakbgcol_ons%>; border-style: solid; width : 56px;"></td>
							<td align="center" style="width : 56px;"><input type="hidden" name="FM_tor_opr_<%=de%>" value="<%=torTimerVal%>"><input type="Text" name="Timer_tor_<%=de%>" maxlength="<%=maxl_tor%>" value="<%=torTimerVal%>"  onkeyup="setTimerTot(<%=de%>, 'tor'), tjektimer('tor',<%=de%>)"; style="!border: 1px; background-color: <%=fmbg_tor%>; border-color: <%=fakbgcol_tor%>; border-style: solid; width : 56px;"></td>
							<td align="center" style="width : 56px;"><input type="hidden" name="FM_fre_opr_<%=de%>" value="<%=freTimerVal%>"><input type="Text" name="Timer_fre_<%=de%>" maxlength="<%=maxl_fre%>" value="<%=freTimerVal%>"  onkeyup="setTimerTot(<%=de%>, 'fre'), tjektimer('fre',<%=de%>)"; style="!border: 1px; background-color: <%=fmbg_fre%>; border-color: <%=fakbgcol_fre%>; border-style: solid; width : 56px;"></td>
							<td align="center" style="width : 56px; border-right: 1px solid #003399;"><input type="hidden" name="FM_lor_opr_<%=de%>" value="<%=lorTimerVal%>"><input type="Text" name="Timer_lor_<%=de%>" maxlength="<%=maxl_lor%>" value="<%=lorTimerVal%>" size="6" onkeyup="setTimerTot(<%=de%>, 'lor'), tjektimer('lor',<%=de%>)"; style="!border: 1px; background-color: <%=fmbg_lor%>; border-color: <%=fakbgcol_lor%>; border-style: solid; width : 56px;"></td>
						</tr>
						<%
						de = de + 1
					end if
	
	'intExt = 0
	'antAkt = 0
	lastJobid = oRec("id")
	oRec.movenext
	wend
	oRec.close
	
	
	
	if de > 1 then
	%>
		</table></div></td></tr>
		</table>
	<%
	end if
	
	
	
	antal_de = de - 1
	
	%><input type="hidden" name="FM_de" value="<%=antal_de%>">
	
	<br><br>
	

<table cellspacing="0" cellpadding="0" border="0" width="530">
<tr bgcolor="#5582D2">
<td width="3"><img src="../ill/venstre_hjorne.gif" alt="" border="0"></td>
<td width="524" style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid;"><font class="stor-hvid">&nbsp;Interne job</font></td>
<td  width="3" align="right"><img src="../ill/hojre_hjorne.gif" alt="" border="0"></td>
</tr>
</table>

<table width='530' border='0' cellspacing='0' cellpadding='0'>
	<%
	'** Udskriver Interne Jobnr **
	strLastjobnr = ""
	strSQL = "SELECT job.id, jobnr, jobnavn, kkundenavn FROM job, kunder WHERE jobstatus = 1 AND fakturerbart = 0 AND kunder.Kid = jobknr ORDER BY Jobnr DESC"
	oRec.cachesize = 10
	oRec.Open strSQL, oConn, 0, 1
		
		di = antal_de + 1
		
		While Not oRec.EOF
		
			
		if showWeekInt <> 1 then
		%>
		<td colspan="2" align="center" style="!background-color : #8DA7E5; border-left: 1px solid #003399; border-right: 1px solid #003399; width : 138px;">
		<font style="color:#ffffff;font-family:arial,sans-serif;font-size:11px;"><b>Uge: <%
		Response.write strWeek & "&nbsp;&nbsp;"
		
		if len(request("regaar")) = 0 then
		Response.write year(now)
		else
		Response.write request("regaar")
		end if%></td>
		<%
		'*** Gennemløber de 7 dage (igen) for at finde deres dato. ***
			for t = 1 to 7
				intStreg = instr(tjekdag(t), "/")
				select case intStreg
				case 2
						if day(now)&"/"&month(now)&"/"&year(now) =  tjekdag(t)&"/" &year(now)  then
						borderColor = "#D6DFF5"
						else
						borderColor = "#FFFFE1"
						end if
				case 3 
						if day(now)&"/"&month(now)&"/"&year(now) =  tjekdag(t)&"/" &year(now)  then
						borderColor = "#D6DFF5"
						else
						borderColor = "#FFFFE1"
						end if
				end select
			
			select case t
			case 1
			dagnavn = "Søn"
			case 2
			dagnavn = "Man"
			case 3
			dagnavn = "Tir"
			case 4
			dagnavn = "Ons"
			case 5
			dagnavn = "Tor"
			case 6
			dagnavn = "Fre"
			case 7
			dagnavn = "Lør"
			end select
			
			%>
			<td height="25" align="center" style="background-color:<%=borderColor%>; border-right: 1px solid #003399; width : 56px;"><font style="color:#000000; font-family:arial,sans-serif; font-size:11px;"><%=dagnavn%>&nbsp;<%=tjekdag(t)%></td>
			<%
			next
		%>		
		</tr>
		<%
		end if
			
					'*** FInder timer på interne job. ****
					sonTimerVal = ""
					manTimerVal = ""
					tirTimerVal = ""
					onsTimerVal = ""
					torTimerVal = ""
					freTimerVal = ""
					lorTimerVal = ""
		
					varTjDatoUS = cdate(convertDate(tjekdag(1)&"/"&strRegaar))
					varTjDatoUS_son = convertDateYMD(varTjDatoUS)
					
					varTjDatoUS = cdate(convertDate(tjekdag(2)&"/"&strRegaar))
					varTjDatoUS_man = convertDateYMD(varTjDatoUS)
					
					varTjDatoUS = cdate(convertDate(tjekdag(3)&"/"&strRegaar))
					varTjDatoUS_tir = convertDateYMD(varTjDatoUS)
					
					varTjDatoUS = cdate(convertDate(tjekdag(4)&"/"&strRegaar))
					varTjDatoUS_ons = convertDateYMD(varTjDatoUS) 
					
					varTjDatoUS = cdate(convertDate(tjekdag(5)&"/"&strRegaar))
					varTjDatoUS_tor = convertDateYMD(varTjDatoUS) 
					
					varTjDatoUS = cdate(convertDate(tjekdag(6)&"/"&strRegaar)) 
					varTjDatoUS_fre = convertDateYMD(varTjDatoUS)
					
					varTjDatoUS = cdate(convertDate(tjekdag(7)&"/"&strRegaar)) 
					varTjDatoUS_lor = convertDateYMD(varTjDatoUS)
					
					
					strSQLtimer = "SELECT timer, tdato FROM timer WHERE Tjobnr = "&oRec("Jobnr")&" AND Tmnavn = '"&Session("user")&"' AND Tdato = '"& varTjDatoUS_son &"'"_
					&" OR "_
					&" Tjobnr = "&oRec("Jobnr")&" AND Tmnavn = '"&Session("user")&"' AND Tdato = '"& varTjDatoUS_man &"'"_
					&" OR "_
					&" Tjobnr = "&oRec("Jobnr")&" AND Tmnavn = '"&Session("user")&"' AND Tdato = '"& varTjDatoUS_tir &"'"_
					&" OR "_
					&" Tjobnr = "&oRec("Jobnr")&" AND Tmnavn = '"&Session("user")&"' AND Tdato = '"& varTjDatoUS_ons &"'"_
					&" OR "_
					&" Tjobnr = "&oRec("Jobnr")&" AND Tmnavn = '"&Session("user")&"' AND Tdato = '"& varTjDatoUS_tor &"'"_
					&" OR "_
					&" Tjobnr = "&oRec("Jobnr")&" AND Tmnavn = '"&Session("user")&"' AND Tdato = '"& varTjDatoUS_fre &"'"_
					&" OR "_
					&" Tjobnr = "&oRec("Jobnr")&" AND Tmnavn = '"&Session("user")&"' AND Tdato = '"& varTjDatoUS_lor &"'"
					
					
					oRec2.Open strSQLtimer, oConn, 0, 1  
					while not oRec2.EOF 
					
						select case DatePart("w", oRec2("tdato")) 
						case 1	
						sonTimerVal = SQLBless(oRec2("timer"))
						case 2
						manTimerVal = SQLBless(oRec2("timer"))
						case 3
						tirTimerVal = SQLBless(oRec2("timer"))
						case 4
						onsTimerVal = SQLBless(oRec2("timer"))
						case 5
						torTimerVal = SQLBless(oRec2("timer"))
						case 6
						freTimerVal = SQLBless(oRec2("timer"))
						case 7
						lorTimerVal = SQLBless(oRec2("timer"))
						end select
						
					oRec2.movenext
					wend
					oRec2.close
					'***************************************************
					
					
		Response.write "<tr><td style='!background-color: #ffffff; padding-left : 4px; padding-right : 4px; border-left: 1px solid #003399;'><font class='lillesort'>"_
		& oRec("Jobnr") &"</td><td style='!border: 0px; background-color: #ffffff; padding-left : 4px; padding-right : 4px;'>"& left(oRec("Jobnavn"), 15)&"</td>"
		%><input type="hidden" name="FM_jobnr_<%=di%>" value="<%=oRec("jobnr")%>">
		<td align="center" style="width : 56px;"><input type="hidden" name="FM_son_opr_<%=di%>" value="<%=sonTimerVal%>"><input type="Text" name="Timer_son_<%=di%>" maxlength="4" value="<%=sonTimerVal%>" onkeyup="setTimerTot(<%=di%>, 'son'), tjektimer('son',<%=di%>)"; style="!border: 1px; background-color: #FFDFDF; border-color: #cd853f; border-style: solid; width : 56px;"></td>
		<td align="center" style="width : 56px;"><input type="hidden" name="FM_man_opr_<%=di%>" value="<%=manTimerVal%>"><input type="Text" name="Timer_man_<%=di%>" maxlength="4" value="<%=manTimerVal%>" onkeyup="setTimerTot(<%=di%>, 'man'), tjektimer('man',<%=di%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #7F9DB9; border-style: solid; width : 56px;"></td>
		<td align="center" style="width : 56px;"><input type="hidden" name="FM_tir_opr_<%=di%>" value="<%=tirTimerVal%>"><input type="Text" name="Timer_tir_<%=di%>" maxlength="4" value="<%=tirTimerVal%>" onkeyup="setTimerTot(<%=di%>, 'tir'), tjektimer('tir',<%=di%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #7F9DB9; border-style: solid; width : 56px;"></td>
		<td align="center" style="width : 56px;"><input type="hidden" name="FM_ons_opr_<%=di%>" value="<%=onsTimerVal%>"><input type="Text" name="Timer_ons_<%=di%>" maxlength="4" value="<%=onsTimerVal%>" onkeyup="setTimerTot(<%=di%>, 'ons'), tjektimer('ons',<%=di%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #7F9DB9; border-style: solid; width : 56px;"></td>
		<td align="center" style="width : 56px;"><input type="hidden" name="FM_tor_opr_<%=di%>" value="<%=torTimerVal%>"><input type="Text" name="Timer_tor_<%=di%>" maxlength="4" value="<%=torTimerVal%>" onkeyup="setTimerTot(<%=di%>, 'tor'), tjektimer('tor',<%=di%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #7F9DB9; border-style: solid; width : 56px;"></td>
		<td align="center" style="width : 56px;"><input type="hidden" name="FM_fre_opr_<%=di%>" value="<%=freTimerVal%>"><input type="Text" name="Timer_fre_<%=di%>" maxlength="4" value="<%=freTimerVal%>" onkeyup="setTimerTot(<%=di%>, 'fre'), tjektimer('fre',<%=di%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #7F9DB9; border-style: solid; width : 56px;"></td>
		<td align="center" style="border-right: 1px solid #003399;"><input type="hidden" name="FM_lor_opr_<%=di%>" value="<%=lorTimerVal%>"><input type="Text" name="Timer_lor_<%=di%>" maxlength="4" value="<%=lorTimerVal%>" size="6" onkeyup="setTimerTot(<%=di%>, 'lor'), tjektimer('lor',<%=di%>)"; style="!border: 1px; background-color: #FFDFDF; border-color: #cd853f; border-style: solid;"></td>
		</tr>
		<%
		if len(sonTimerVal) <> "0" then
		sonTimerTot = sonTimerTot + sonTimerVal
		end if
		
		if len(manTimerVal) <> "0" then
		manTimerTot = manTimerTot + manTimerVal
		end if
		
		if len(tirTimerVal) <> "0" then
		tirTimerTot = tirTimerTot + tirTimerVal
		end if
		
		if len(onsTimerVal) <> "0" then
		onsTimerTot = onsTimerTot + onsTimerVal
		end if
		
		if len(torTimerVal) <> "0" then
		torTimerTot = torTimerTot + torTimerVal
		end if
		
		if len(freTimerVal) <> "0" then
		freTimerTot = freTimerTot + freTimerVal
		end if
		
		if len(lorTimerVal) <> "0" then
		lorTimerTot = lorTimerTot + lorTimerVal
		end if
		
		showWeekInt = 1
		di = di + 1
		strLastjobnr = oRec("Jobnr")
	oRec.MoveNext
	Wend 
	
	oRec.Close
	
	antal_di = di - 1%>
<input type="hidden" name="FM_di" value="<%=antal_di%>">
<tr>
    <td colspan="9" align="right"><br><input type="image" src="../ill/frem.gif"></td>
</tr>
</table>

</td>
</tr>
</table>
<br>

<!--Viser timer pr. i toppen af siden -->
<%
'Response.write antal_de &" -- "& antal_di 
TimerTot = sonTimerTot + manTimerTot + tirTimerTot + onsTimerTot + torTimerTot + freTimerTot + lorTimerTot 
%>
	
	<table cellspacing="0" cellpadding="0" border="0" width="530">
	<tr bgcolor="#5582D2">
	<td width="3"><img src="../ill/venstre_hjorne.gif" alt="" border="0"></td>
	<td width="524" style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid;"><font class="stor-hvid">&nbsp;Oversigt for denne uge</font></td>
	<td  width="3" align="right"><img src="../ill/hojre_hjorne.gif" alt="" border="0"></td>
	</tr>
	</table>
	<table cellpadding="2" cellspacing="1" border="0" width="530" bgcolor="#D6DFF5">
	<tr bgcolor="#ffffff">
	<td style="width : 138px;"></td>
	<td style="width : 60px;" align="center"><font style="color:#000000; font-family:arial,sans-serif; font-size:11px;">Søn <%=tjekdag(1)%></td>
	<td style="width : 60px;" align="center"><font style="color:#000000; font-family:arial,sans-serif; font-size:11px;">Man <%=tjekdag(2)%></td>
	<td style="width : 60px;" align="center"><font style="color:#000000; font-family:arial,sans-serif; font-size:11px;">Tir <%=tjekdag(3)%></td>
	<td style="width : 60px;" align="center"><font style="color:#000000; font-family:arial,sans-serif; font-size:11px;">Ons <%=tjekdag(4)%></td>
	<td style="width : 60px;" align="center"><font style="color:#000000; font-family:arial,sans-serif; font-size:11px;">Tor <%=tjekdag(5)%></td>
	<td style="width : 60px;" align="center"><font style="color:#000000; font-family:arial,sans-serif; font-size:11px;">Fre <%=tjekdag(6)%></td>
	<td style="width : 60px;" align="center"><font style="color:#000000; font-family:arial,sans-serif; font-size:11px;">Lør <%=tjekdag(7)%></td>
	</tr>
	<tr bgcolor="#ffffff">
	<td style="width : 138px;">Antal timer:</td>
	<td style="width : 56px;" align="center"><input type="text" name="FM_son_total" value="<%=SQLBless(sonTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:36px;'></td>
	<td style="width : 56px;" align="center"><input type="text" name="FM_man_total" value="<%=SQLBless(manTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:36px;'></td>
	<td style="width : 56px;" align="center"><input type="text" name="FM_tir_total" value="<%=SQLBless(tirTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:36px;'></td>
	<td style="width : 56px;" align="center"><input type="text" name="FM_ons_total" value="<%=SQLBless(onsTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:36px;'></td>
	<td style="width : 56px;" align="center"><input type="text" name="FM_tor_total" value="<%=SQLBless(torTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:36px;'></td>
	<td style="width : 56px;" align="center"><input type="text" name="FM_fre_total" value="<%=SQLBless(freTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:36px;'></td>
	<td style="width : 56px;" align="center"><input type="text" name="FM_lor_total" value="<%=SQLBless(lorTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:36px;'></td>
	</tr>
	<tr bgcolor="#ffffff">
	<td colspan="6">Antal timer denne uge:</td>
	<td colspan="2" align="center"><input type="text" name="FM_week_total" value="<%=SQLBless(TimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:45px;' maxlenght=4> timer</td>
	</tr>
	</table>
	<br><br><br>

</div>
</form>
<%
end if
%>



<!--#include file="../inc/regular/footer_inc.asp"-->
 




