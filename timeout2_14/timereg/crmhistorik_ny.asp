<%Response.Buffer = true %>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/google_conn.asp"-->

<!--#include file="inc/isint_func.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	
	
	strOprHid = request.Form("FM_oprethidden")
	func = request("func")
	
	'hvis der er blevet valgt en kunde i øverste dropdown
	if func = "dbopr" and Not strOprHid = "j" then
	func = "opret"
	end if
	
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	if Instr(id, ",") > 0 then
	tid = split(id,",")
	id = tid(0)
	end if
	end if
	
	
	
	
	'*********************************************
	'**************Start  på dropdown funktioner
	'*******************************************
Function DropDownMinute(FormNavn, KlokkesletField)
if func = "red" then
strSQL = "SELECT " & KlokkesletField & " FROM crmhistorik Where id = " & request.QueryString("crmaktion")
			oRec.open strSQL, oConn, 3
strSelVal = Minute(FormatDateTime(oRec(KlokkesletField)))
oRec.Close
else
        strSelVal = Request.Form(FormNavn)
end if
        if Not strSelVal <> "" then
        strSelVal = "00"
        end if

        Dim arrMinutes
        arrMinutes = split("00,15,30,45", ",")
        for i = 0 to 3
        if arrMinutes(i) = strSelVal then
        strSelOpt = " selected"
        else
        strSelOpt = Null
        end if
        %>
        <option value="<%=arrMinutes(i)%>" <%=strSelOpt%>><%=arrMinutes(i)%></option>
        <%Next
End Function



Function DropDownHour(FormNavn, KlokkesletField)
if func = "red" then
strSQL = "SELECT " & KlokkesletField & " FROM crmhistorik Where id = " & request.QueryString("crmaktion")
			oRec.open strSQL, oConn, 3
strSelVal = Hour(FormatDateTime(oRec(KlokkesletField)))
oRec.close
else
        strSelVal = Request.Form(FormNavn)
end if
        if Not strSelVal <> "" then
		 strSelVal = Hour(Now)
		 end if
        Response.Write "strSelVal="&strSelVal
		 for i = 0 to 23
		 if i = cint(strSelVal) then
		 strSelOpt = " selected"
		 else
		 strSelOpt = Null
		 end if%>
		<option value="<%=i%>" <%=strSelOpt%>><%=i%></option>
		<%Next
End Function




Function DropDownDay(FormNavn, DatoField)
if func = "red" then
strSQL = "SELECT " & DatoField & " FROM crmhistorik Where id = " & request.QueryString("crmaktion")
			oRec.open strSQL, oConn, 3
strSplitDato = split(oRec(DatoField), "-")
oRec.Close
strSelVal = strSplitDato(0)
else
        strSelVal = Request.Form(FormNavn)
end if
        if Not strSelVal <> "" then
		 strSelVal = Day(Now)
		end if

        for i = 1 to 31
		if cint(strSelVal) = i then
		strSelOpt = " selected"
		else
		strSelOpt = Null
		end if %>
		<option value="<%=i%>" <%=strSelOpt%>><%=i%></option>
        <%Next
End Function


Function DropDownMonth(FormNavn, DatoField)
if func = "red" then
strSQL = "SELECT " & DatoField & " FROM crmhistorik Where id = " & request.QueryString("crmaktion")
			oRec.open strSQL, oConn, 3
strSplitDato = split(oRec(DatoField), "-")
oRec.Close
strSelVal = strSplitDato(1)
else
		strSelVal = Request.Form(FormNavn)
end if
		if Not strSelVal <> "" then
		strSelVal = Month(Now())
		end if

		dim ArrMonth
		ArrMonth = split("0, jan, feb, mar, apr, maj, jun, jul, aug, sep, okt, nov, dec",",")
		for i = 1 to 12
		if cint(strSelVal) = i then
		strSelOpt = " selected"
		else
		strSelOpt = Null
		end if
		%>
		<option value="<%=i%>"<%=strSelOpt%>><%=ArrMonth(i)%></option>
<%Next 
End Function
Function DropDownYear(FormNavn, DatoField)
if func = "red" then
strSQL = "SELECT " & DatoField & " FROM crmhistorik Where id = " & request.QueryString("crmaktion")
			oRec.open strSQL, oConn, 3
strSplitDato = split(oRec(DatoField), "-")
oRec.Close
strSelVal = strSplitDato(2)
else
        strSelVal = Request.Form(FormNavn)
end if
		 if Not strSelVal <> "" then
		 strSelVal = Year(Now())
		 end if

		 for i = Year(Now())-5 to Year(Now())+5
		 if i = cint(strSelVal) then
		 strSelOpt = " selected"
		 else
		 strSelOpt = Null
		 end if
		 %>
		<option value="<%=i%>"<%=strSelOpt%>><%=i%></option>
		<%Next
End Function




	'***************************************                            *******'
	'******             Slut på funktioner              **********************'
	'*************************************                              *********'

	
	kl = request("kl")
	crmaktion = request("crmaktion")
	thisfile = "crmhistorik"
	
	strNavn = request("navn")
	select case func
	case "slet"
	serial = request("serial")
	'*** Her spørges om det er ok at der slettes en aktion ***
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:120; visibility:visible;">
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td><img src="../ill/alert.gif" width="44" height="45" alt="" border="0">
		<%if cint(serial) <> 0 then%>
		Du er ved at <b>slette</b> en aktion.<br>
		<img src="../ill/blank.gif" width="46" height="1" alt="" border="0"><b>Denne aktion er en del af en serie!.</b><br>
		<img src="../ill/blank.gif" width="46" height="1" alt="" border="0">Derfor har du følgende valgmuligheder:<br>
		<%
		else%>
		&nbsp;Du er ved at <b>slette</b> en aktion. Er dette korrekt?
		<%end if%></td>
	</tr>
	<tr>
	   <td><%if cint(serial) <> 0 then%>
		<br><img src="../ill/blank.gif" width="44" height="1" alt="" border="0"><a href="crmhistorik.asp?menu=crm&func=sletok&id=<%=id%>&crmaktion=<%=crmaktion%>&ketype=e&emner=<%=request("emner")%>&status=<%=request("status")%>&medarb=<%=request("medarb")%>&selpkt=<%=request("selpkt")%>&serial=<%=serial%>"><font color="LimeGreen"><i><b>VV</b></i></font>&nbsp;Ja, slet hele serien!</a><br>
		<%end if%>
	  	<br><img src="../ill/blank.gif" width="44" height="1" alt="" border="0"><a href="crmhistorik.asp?menu=crm&func=sletok&id=<%=id%>&crmaktion=<%=crmaktion%>&ketype=e&emner=<%=request("emner")%>&status=<%=request("status")%>&medarb=<%=request("medarb")%>&selpkt=<%=request("selpkt")%>"><font color="LimeGreen"><i><b>V</b></i></font>&nbsp;Ja, slet denne aktion</a><br><br>
		<img src="../ill/blank.gif" width="48" height="1" alt="" border="0"><a href="javascript:history.back()"><font color="Darkred"><b>!</b></font>&nbsp;&nbsp;&nbsp;Nej, dette er en fejl. Vend tilbage...</a></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	case "sletok"
	serial = request("serial")
	strGoogleCase = "delete"
	GoogleStringAktionsId = crmaktion
	%>
	<!--#include file="googlepost.asp"-->
	<img src="../ill/info.gif" width="42" height="38" alt="" border="0" style="position: absolute; top: -1000px;" />
	<%
	GoogleStringAktionsId = Null
	strGoogleCase = Null
	'*** Her slettes en crmaktion ***
	if cint(serial) <> 0 then
	oConn.execute("DELETE FROM crmhistorik WHERE serialnb = "& serial &"")
	else
	oConn.execute("DELETE FROM crmhistorik WHERE id = "& crmaktion &"")
	end if
	
	'if request("selpkt") <> "kal" then
	'Response.redirect "crmhistorik.asp?menu=crm&shokselector=1&ketype=e&id="&id&"&func=hist&selpkt=hist&emner="&request("emner")&"&status="&request("status")&"&medarb="&request("medarb")&""
	'else
	'Response.redirect "crmkalender.asp?menu=crm&shokselector=1&id="&id&"&medarb="&request("medarb")&"&selpkt=kal"
	'end if
	
	'**********************************************************************************************
	'** Her indsættes en ny aktion i db 													   ****
	'**********************************************************************************************
	case "dbopr", "dbred"
	personal = request("personal") 
	opretopf = request("FM_opf")
	
	'** Disse bruges ved redigering af serie / gentagelse af aktion
	redserie = request("FM_opd_serie")
	thisSerialnb = request("FM_serialnb")
	
	'**********************************************************************
	'*** Validering starter 											***
	'**********************************************************************
	
	
	'** Validering af start og slut dato. ***
	if personal <> "y" then
	valstrCrmdato = request("FM_start_dag")&"/"&request("FM_start_mrd")&"/"&request("FM_start_aar")
	else
	valstrCrmdato = "1/1/2001"
	end if
	
	if personal <> "y" then
		if cint(opretopf) <> 1 then
		valstrCrmdato_2 = "12/12/2014"
		else
		valstrCrmdato_2 = request("FM_start_dag_2")&"/"&request("FM_start_mrd_2")&"/"&request("FM_start_aar_2")
		end if
	else
	valstrCrmdato_2 = "12/12/2014"
	end if
	'***** Validering *******
	if cint(opretopf) = 1 AND cdate(valstrCrmdato) > cdate(valstrCrmdato_2) then%>	
		<!--#include file="../inc/regular/header_inc.asp"-->
		<!--#include file="../inc/regular/topmenu_inc.asp"-->
		<%
		errortype = 37
		call showError(errortype)	
				
	else
		'****************************************
		'**** sætter CRM-klokkeslet på aktion ***
		'****************************************
		if len(request("klok_timer")) = 0 then
		strKloktimer = "12"
		else
		strKloktimer = request("klok_timer")
		end if
		
		if len(request("klok_min")) = 0 then
		strKlokmin = "00"
		else
		strKlokmin = request("klok_min")
		end if
		
		if len(request("klok_timer_slut")) = 0 then
		strKloktimer_slut = "12"
		else
		strKloktimer_slut = request("klok_timer_slut")
		end if
		
		if len(request("klok_min_slut")) = 0 then
		strKlokmin_slut = "00"
		else
		strKlokmin_slut = request("klok_min_slut")
		end if
		
		'**** Sætter CRM-klokkeslet på aktion *****
		strKlokkeslet = strKloktimer &":"& strKlokmin &":"&"01"
		strKlokkeslet_slut = strKloktimer_slut &":"& strKlokmin_slut &":"&"01"
		
		
		'**** Sætter CRM-klokkeslet på opfølgende aktion *****
		strKlokkeslet_2 = request("klok_timer_2") &":"& request("klok_min_2") &":"&"01"
		strKlokkeslet_slut_2 = request("klok_timer_slut_2") &":"& request("klok_min_slut_2") &":"&"01"
		' *** '
		
		
				'***** Validering *******
				if cint(opretopf) = 1 AND (cdate(valstrCrmdato) = cdate(valstrCrmdato_2)) AND (strKlokkeslet > strKlokkeslet_2 OR strKlokkeslet_slut > strKlokkeslet_2 OR strKlokkeslet_slut > strKlokkeslet_slut_2) then%>	
				<!--#include file="../inc/regular/header_inc.asp"-->
				<!--#include file="../inc/regular/topmenu_inc.asp"-->
				<%
				errortype = 38
				call showError(errortype)	
							
				else
	
					'***** Validering *******
					if (cint(opretopf) = 2) AND (cdate(valstrCrmdato) >= cdate(valstrCrmdato_2)) then%>
						<!--#include file="../inc/regular/header_inc.asp"-->
						<!--#include file="../inc/regular/topmenu_inc.asp"-->
						<%
						errortype = 36
						call showError(errortype)
					else
		
					
					'*** Validering ***
					if (cint(opretopf) = 3 AND (request("FM_dageldato") = "dato" AND len(trim(request("FM_gentag_dato"))) = 0)) OR (cint(opretopf) = 3 AND (request("FM_dageldato") = "dag" AND len(trim(request("FM_gentag_ugedag"))) = 0)) OR (cint(opretopf) = 3 AND len(trim(request("FM_gentag_md"))) = 0) then
						%>
						<!--#include file="../inc/regular/header_inc.asp"-->
						<%
						errortype = 41
						call showError(errortype)
					else
						
		'*****************************
		'* Validering godkendt    ****
		'* Starter oprettelse 	 *****
		'*****************************
		%>
		<!--#include file="inc/convertDate.asp"-->
		<%
		
		
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		
		'*****************************
		'** Henter medarbejder id ****
		'*****************************
		strSQL = "SELECT mid FROM medarbejdere WHERE mnavn = '"& session("user") &"'"
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		intMid = oRec("mid")
		End if
		oRec.close
		''**''
		
		'****************************************************************************
		'*** Sætter værdier alt efter om det er en personlig note eller aktion. 	*
		'****************************************************************************
		if personal <> "y" then
			strNavn = SQLBless(request("FM_navn"))
			strEmne = request("FM_emne")
			tid = split(request("id"),",")
	        intKundeId = tid(0)
			strBesk = SQLBless(request("FM_besk"))
			strEditor = session("user")
			strDato = year(now)&"/"&month(now)&"/"&day(now)
			strCrmdato = request("FM_start_aar")&"/"&request("FM_start_mrd")&"/"&request("FM_start_dag")
			strStatus = request("FM_status")
			strKontaktform = request("FM_kform")
			strKpers = request("FM_kpers")
			'strKlokkeslet sættes under validering
			'strKlokkeslet_slut sættes under validering
			if opretopf = 2 then
			strCrmdato_slut = request("FM_slut_aar")&"/"&request("FM_slut_mrd")&"/"&request("FM_slut_dag")
			else
			strCrmdato_slut = strCrmdato
			end if
			
			'** Hvis slut klokkeslet er før starttid klokkeslet
			if strKlokkeslet > strKlokkeslet_slut then
			strKlokkeslet_slut = strKlokkeslet 
			end if
		
		else '*** personal note ****
			strNavn = ""
			strEmne = 0
			intKundeId = -1
			strBesk = SQLBless(request("FM_note_personal"))
			strEditor = session("user")
			strDato = year(now)&"/"&month(now)&"/"&day(now)
			strCrmdato = request("FM_start_aar_per")&"/"&request("FM_start_mrd_per")&"/"&request("FM_start_dag_per")
			strStatus = 0
			strKontaktform = 0
			strKpers = ""
			strKlokkeslet = request("FM_klslet_personal")&":"&"01"
			strKlokkeslet_slut = dateadd("h", 1, strKlokkeslet)
			strCrmdato_slut = strCrmdato
		end if
		' *** '
		
		'***************************************************************
		'** Sætter værdier på Getagelser, hvis dette punkt er valgt. ***
		'***************************************************************
		if cint(opretopf) = 3 then
		
		'*** finder næste serienummer ***
		strSQL = "SELECT serialnb FROM crmhistorik ORDER BY serialnb DESC"
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then 
		nextSerialnb = oRec("serialnb") + 1
		else
		nextSerialnb = 0
		end if
		oRec.close
		
		Call updinsertAktion(strCrmdato, strCrmdato_slut, nextSerialnb)
		
		g = 0
		i = 0
		a = 0
		num = 0
		dim useThis_g
		Redim useThis_g(i)
		dim intDatoer
		Redim intDatoer(g)
		
		dageldato = request("FM_dageldato")
		slutgentag = request("FM_gentag_dag")&"/"&request("FM_gentag_mrd")&"/"&request("FM_gentag_aar")
		
		if len(request("FM_gentag_ugedag")) <> 0 then
		g_ugedag = request("FM_gentag_ugedag")
		else
		g_ugedag = "1"
		end if
		
		if len(request("FM_gentag_dato")) <> 0 then
		g_dato = request("FM_gentag_dato")
		else
		g_dato = "1"
		end if
		
		if len(request("FM_gentag_md")) <> 0 then
		g_md = request("FM_gentag_md")
		else
		g_md = "1"
		end if
		
		diff = (datediff("yyyy", strCrmdato, slutgentag) + 1)
		weekdiff = datediff("ww", strCrmdato, slutgentag)
				
			if dageldato = "dato" then
					
				For a = 0 to diff - 1
					
					if a = 0 then
					ya = year(strCrmdato)
					else
					ya = ya + 1
					end if
					
					intDatoer = Split(g_dato, ", ")
					For g = 0 to Ubound(intDatoer)
						
						intMD = Split(g_md, ", ")
						For i = 0 to Ubound(intMD)
							Redim preserve useThis_g(i)
							useThis_g(i) = intMD(i)&"/"&intDatoer(g)&"/"&ya 
							
							if (isdate(useThis_g(i)) AND isdate(slutgentag) AND isdate(strCrmdato)) then
								if (cdate(useThis_g(i)) <= cdate(slutgentag)) AND cdate(useThis_g(i)) >= cdate(strCrmdato) then
								Call updinsertAktion(convertDateYMD(useThis_g(i)), convertDateYMD(useThis_g(i)), nextSerialnb)
								thisdateisused = thisdateisused & formatdatetime(useThis_g(i), 1) &"<br>"
								end if
							else
								somerrors = "y"
							end if
						Next
					Next
				Next
				
			else 'Uge dage er valgt
				
				Dim  usethisDato
				Redim  usethisDato(g)
				usedDatoer = "#"
				crmDatoweekday = datepart("w", strCrmdato)
				crmDatoday = day(strCrmdato)
				
				'*** Tjekker hvor mange uger der er imellem crmdato og slutdato ***
				if weekdiff > 0 then
				useWeekdif = weekdiff - 1
				else
				useWeekdif = 1
				end if
				
				
				For a = 0 to useWeekdif
						
					intDatoer = Split(g_ugedag, ", ")
					For g = 0 to Ubound(intDatoer)
					
					'** tjekker om det er første gennemløb og finder dato på de dage der er valgt. **
					isused = instr(usedDatoer, intDatoer(g))
					if cint(isused) = 0 then
						dayadd = crmDatoday + (intDatoer(g) - crmDatoweekday)  
						Redim preserve usethisDato(g)
						usethisDato(g) = dayadd &"/"& month(strCrmdato) &"/"& year(strCrmdato)
						'usethisDato(g) = year(strCrmdato)&"/"&month(strCrmdato)&"/"&dayadd 
						
						if cint(intDatoer(g)) > cint(crmDatoweekday) then
						usethisDato(g) = usethisDato(g)
						else
						usethisDato(g) = dateadd("ww", 1, usethisDato(g))
						end if
						
						usedDatoer = usedDatoer & "#"&intDatoer(g)&"#,"
					else
						usethisDato(g) = dateadd("ww", 1, usethisDato(g))
					end if 
						
						'** tjekker om der er valgt måneder **
						intMD = Split(g_md, ", ")
						For i = 0 to Ubound(intMD)
							if cint(month(usethisDato(g))) = cint(intMD(i)) then 
								
								if (cdate(usethisDato(g)) <= cdate(slutgentag)) AND cdate(usethisDato(g)) >= cdate(strCrmdato) then
									Redim preserve useThis_g(i)
									useThis_g(i) = usethisDato(g)
									thisdateisused = thisdateisused & formatdatetime(useThis_g(i), 1) &"<br>"
									Call updinsertAktion(convertDateYMD(useThis_g(i)), convertDateYMD(useThis_g(i)), nextSerialnb)
									'Call updinsertAktion(useThis_g(i), useThis_g(i), nextSerialnb)
								
								end if
							end if
						Next
					Next
				Next
			
			
			end if
			
			
		
		else
		intserialnb = 0
		Call updinsertAktion(strCrmdato, strCrmdato_slut, intserialnb)
		end if
		' *** '
	
		
		
		'**********************************************************
		'**** Opretter / Redigerer aktion(er) 				*******
		'**********************************************************
		Public aktionsid
		function updinsertAktion(usefuncdato, usefuncslutdato, useserial)
		
		thisfuncdato = usefuncdato
		thisfuncslutdato = usefuncslutdato
		thisserial = useserial
		if func = "dbopr" and strOprHid = "j" then
		oConn.execute("INSERT INTO crmhistorik (editor, dato, crmdato, kontaktemne, "_
		&" kontaktpers, status, komm, navn, kundeid, kontaktform, editorid, "_
		&" crmklokkeslet, crmKlokkeslet_slut, crmdato_slut, serialnb) "_
		&" VALUES ('"& strEditor &"', '"& strDato &"', '"& thisfuncdato &"', "& strEmne &", "_
		&" '"& strKpers &"', '"& strStatus &"', '"& strBesk &"', '"& strNavn &"',"_
		&" "& intKundeId &", "& strKontaktform &", "& intMid &", '"& strKlokkeslet &"', "_
		&" '"& strKlokkeslet_slut &"', '"& thisfuncslutdato &"', "& thisserial &")")

			strSQL = "SELECT id FROM crmhistorik ORDER BY id DESC"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
				aktionsid = oRec("id")
			end if
			oRec.close
			
		elseif strOprHid = "j" then
		
			if cint(redserie) = 1 then 'rediger serie
			useSQL = "UPDATE crmhistorik SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', kontaktemne = "& strEmne &", kontaktpers = '"& strKpers &"', status = '"& strStatus &"', komm = '"& strBesk &"', navn = '"& strNavn &"', kundeid = "& intKundeId &", kontaktform = "& strKontaktform &", editorid = "& intMid &", crmklokkeslet = '"& strKlokkeslet &"', crmklokkeslet_slut = '"& strKlokkeslet_slut &"' WHERE serialnb = "& thisSerialnb &""
			else
			useSQL = "UPDATE crmhistorik SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', crmdato = '"& thisfuncdato &"', kontaktemne = "& strEmne &", kontaktpers = '"& strKpers &"', status = '"& strStatus &"', komm = '"& strBesk &"', navn = '"& strNavn &"', kundeid = "& intKundeId &", kontaktform = "& strKontaktform &", editorid = "& intMid &", crmklokkeslet = '"& strKlokkeslet &"', crmklokkeslet_slut = '"& strKlokkeslet_slut &"', crmdato_slut = '"& thisfuncslutdato &"', serialnb = 0 WHERE id = "& crmaktion &""
			end if
			
			oConn.execute(useSQL)
			aktionsid = crmaktion
		end if
		
		call relationer()
		
		end function
		' *** '
		
		
		
		'*******************************************************
		'**** Opdaterer aktions relationer til medarbejdere ****
		'*******************************************************
		function relationer()
		oConn.execute("DELETE FROM aktionsrelationer WHERE aktionsid = "& aktionsid &"")
			
		if len(request("FM_medarbrel")) <> 0 then
			Dim medarbRelationer
			Dim b
			b = 0
			medarbRelationer = Split(request("FM_medarbrel"), ", ")
				For b = 0 to Ubound(medarbRelationer)
				oConn.execute("INSERT INTO aktionsrelationer (aktionsid, medarbid) VALUES ("& aktionsid &", "& medarbRelationer(b) &") ")
				next
		else
		oConn.execute("INSERT INTO aktionsrelationer (aktionsid, medarbid) VALUES ("& aktionsid &", "& intMid &") ")
		end if
		end function
		' *** '
		
		
		'*************************************************
		'*** Opretter opfølgning on then fly ... 		 *
		'*************************************************
		if cint(opretopf) = 1 AND func = "dbopr" then
			
			strNavn_2 = SQLBless(request("FM_navn_2"))
			strEmne_2 = request("FM_emne")
			tid = split(request("id"),",")
	        intKundeId_2 = tid(0)
			strBesk_2 = SQLBless(request("FM_besk_2"))
			strEditor_2 = session("user")
			strDato_2 = year(now)&"/"&month(now)&"/"&day(now)
			strCrmdato_2 = request("FM_slut_aar_2")&"/"&request("FM_slut_mrd_2")&"/"&request("FM_slut_dag_2")
			strStatus_2 = request("FM_status_2")
			strKontaktform_2 = request("FM_kform_2")
			strKpers_2 = request("FM_kpers_2")
			strCrmdato_slut = strCrmdato_2
			'strKlokkeslet_2 sættes under validering
			'strKlokkeslet_slut_2 sættes under validering
			
			'** Hvis slut klokkeslet er før starttid
			if strKlokkeslet_2 > strKlokkeslet_slut_2 then
			strKlokkeslet_slut_2 = strKlokkeslet_2 
			end if

			oConn.execute("INSERT INTO crmhistorik (editor, dato, crmdato, kontaktemne, kontaktpers, status, komm, navn, kundeid, kontaktform, editorid, crmklokkeslet, crmKlokkeslet_slut, crmdato_slut, serialnb) VALUES ('"& strEditor_2 &"', '"& strDato_2 &"', '"& strCrmdato_2 &"', "& strEmne_2 &", '"& strKpers_2 &"', '"& strStatus_2 &"', '"& strBesk_2 &"', '"& strNavn_2 &"', "& intKundeId_2 &", "& strKontaktform_2 &", "& intMid &", '"& strKlokkeslet_2 &"', '"& strKlokkeslet_slut_2 &"', '"& strCrmdato_slut &"', 0)")
			
			strSQL = "SELECT id FROM crmhistorik ORDER BY id DESC"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
				aktionsid = oRec("id")
			end if
			oRec.close
		
		call relationer()
		end if
		' *** '
		
		
		
		'*** Skærm ***'
		'************************************************************************'
		'** Oprettelse /redigering slut. Viser ok side til bruge eller			*'
		'** Retunerer til crmkalender hvis det eren personlig note. 			*'
		'************************************************************************'
			if personal <> "y" then
			%>
			<!--#include file="../inc/regular/google_conn.asp"-->
			<body topmargin="0" leftmargin="10" class="regular">
            
            <table cellspacing="0" cellpadding="0" border="0">
			<tr>
		    <td valign="top"><br>
			<img src="../ill/header_crmosigt.gif" alt="" border="0"><hr align="left" width="360" size="1" color="#000000" noshade>
			</td></tr></table>
			<br>
			<table><tr><td>
			<img src="../ill/info.gif" width="42" height="38" alt="" border="0"><b>Aktionen er oprettet!</b><br><br>
			<%if cint(opretopf) = 3 AND len(thisdateisused) > 0 then
				Response.write "<br>Følgende andre aktioner er indsat i denne serie:<br>"
				
				response.write thisdateisused
				
				if somerrors = "y" then
				Response.write "<br>" & "Nogle af de valgte datoer er ikke oprettet.<br> Dette skyldes enten at de ligger uden for det valgte gentagelses interval, eller fordi de valgte datoer ikke eksisterer."
				end if
				
			Response.write "<br><br>"	
			%>
			<a onClick="alt()" href="Javascript:window.close()">Luk dette vindue</a><br><br>
			NB) Husk at opdatere [F5] den side du kom fra. <br>
			så du kan se den nye aktion.
			</td></tr></table>
			
			<%
			
			Response.Write "her 2"
			Response.end
			
			else
			
			Response.Write "her 3"

			if func = "dbopr" and strOprHid = "j" then
			strGoogleCase = "create"
			elseif strOprHid = "j" then
			strGoogleCase = "update"
			end if
			GoogleStringAktionsid = aktionsid
			%>
			<!--#include file="googlepost.asp"-->
<%
            
			end if
			
			
			 
			else
				if len(request("kalviskunde")) <> 0 then
				kalviskunde = request("kalviskunde")
				else
				kalviskunde = 0
				end if
			Response.redirect "crmkalender.asp?menu=crm&shokselector=1&strdag="&request("FM_start_dag_per")&"&strmrd="&request("FM_start_mrd_per")&"&straar="&request("FM_start_aar_per")&"&medarb="&request("kalmedarb")&"&id="&kalviskunde&"&ketype=e&selpkt=kal"
			end if
		' *** '
		
		end if '*validering
		end if '*validering
	end if '*validering
	end if '*validering
	
	
	
	
	
	
	
	
	
	case "opret", "red"
	
	'************************************************************************************************'
	'*** Her indlæses form til rediger/oprettelse af ny aktion.										*'
	'************************************************************************************************'
	'** local kr/dato/nummer format
	session.LCID = 1030
	
	if func = "opret" then
	strTimepris = ""
	trearsize = "40"
	'varSubVal = "opretpil" 
	'varbroedkrumme = "Opret ny aktion"
		
		'** Hvis der først skal vælges kunde ***
		dbfunc = "dbopr"
		varSubVal = "opretpil"
		varbroedkrumme = "Opret ny aktion"
		'*** ****
	
	thisserialnb = -1
	intKontaktpers = 0	
	else
	
	strSQL = "SELECT editor, dato, crmdato, kontaktemne, kontaktpers, status, komm, navn, kundeid, kontaktform, crmklokkeslet, crmklokkeslet_slut, crmDato_slut, serialnb FROM crmhistorik WHERE id=" & crmaktion 
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	crmDato = oRec("crmDato")
	StrTdato = crmDato
	intEmne = oRec("kontaktemne")
	intStatus = oRec("status")
	strBesk = oRec("komm")
	intKontaktform = oRec("kontaktform")
	strKlokkeslet = oRec("crmklokkeslet")
	strKlokkeslet_slut = oRec("crmklokkeslet_slut")
	crmDato_slut = oRec("crmDato_slut")
	thisserialnb = oRec("serialnb")
	intKontaktpers = oRec("kontaktpers")
	end if
	oRec.close
	
	
	trearsize = "40"
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	
	if request("showinwin") <> "j" then%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<%else%>
	<html>
		<head>
		<title>timeOut 2.1</title>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print.css">
		<script LANGUAGE="javascript">
		function NewWin_popupcal(url)    {
		window.open(url, 'Calpick', 'width=250,height=250,scrollbars=no,toolbar=no');
}
		</script>
		</head>
		<% if func <> "opret" and func <> "red" then %>
		<body topmargin="0" leftmargin="0" class="regular">
		<% else%>
		<body topmargin="0" leftmargin="0">
		<%end if%>
	<%end if%>
<script language="javascript">
    var ns, ns6, ie, newlayer;

    ns4 = (document.layers) ? true : false;
    ie4 = (document.all) ? true : false
    ie5 = (document.getElementById) ? true : false
    ns6 = (document.getElementById && !document.all) ? true : false;

    function getLayerStyle(lyr) {
        if (ns4) {
            return document.layers[lyr];
        } else if (ie4) {
            return document.all[lyr].style;
        } else if (ie5) {
            return document.all[lyr].style;
        } else if (ns6) {
            return document.getElementById(lyr).style;
        }
    }

    function ShowHide(layer) {
        newlayer = getLayerStyle(layer)

        var styleObj = (ns4) ? document.layers[layer] : (ie4) ? document.all[layer].style : document.getElementById(layer).style;

        if (newlayer.visibility == "hidden") {
            newlayer.visibility = "visible";
            styleObj.display = ""
        }
        else if (newlayer.visibility == "visible") {
            newlayer.visibility = "hidden";
            styleObj.display = "none"
        }
    }

</script>


	<script LANGUAGE="javascript">
	    function ValidateDate() {
	        var SDateD = document.getElementById("FM_start_dag").value;
	        var SDateM = document.getElementById("FM_start_mrd").value;
	        var SDateY = document.getElementById("FM_start_aar").value;
	        var SDateH = document.getElementById("klok_timer").value;
	        var SDateMm = document.getElementById("klok_min").value;

	        var EDateD = document.getElementById("FM_slut_dag").value;
	        var EDateM = document.getElementById("FM_slut_mrd").value;
	        var EDateY = document.getElementById("FM_slut_aar").value;
	        var EdateH = document.getElementById("klok_timer_slut").value;
	        var EdateMm = document.getElementById("klok_min_slut").value;

	        var alertReason1 = "Slut Dato skal være efter eller den samme som startdatoen.";

	        var endDate = new Date(+EDateY, +EDateM - 1, +EDateD, +EdateH, +EdateMm);
	        var startDate = new Date(+SDateY, +SDateM - 1, +SDateD, +SDateH, +SDateMm);

	        if (startDate.getTime() > endDate.getTime()) {
	            alert(alertReason1);
	            document.getElementById("FM_slut_dag").value = document.getElementById("FM_start_dag").value;
	            document.getElementById("FM_slut_mrd").value = document.getElementById("FM_start_mrd").value;
	            document.getElementById("FM_slut_aar").value = document.getElementById("FM_start_aar").value;
	            document.getElementById("klok_timer_slut").value = document.getElementById("klok_timer").value;
	            document.getElementById("klok_min_slut").value = document.getElementById("klok_min").value;
	            return false;
	        }
	    }
function ValidateDateOpf() {
	        var SDateD = document.getElementById(FM_start_dag_2).value;
	        var SDateM = document.getElementById(FM_start_mrd_2).value;
	        var SDateY = document.getElementById(FM_start_aar_2).value;
	        var EDateD = document.getElementById(FM_slut_dag_2).value;
	        var EDateM = document.getElementById(FM_slut_mrd_2).value;
	        var EDateY = document.getElementById(FM_slut_aar_2).value;

	        var alertReason1 = 'Slut Dato skal være efter eller den samme som startdatoen.'

	        var endDate = new Date(+EDateY, +EDateM - 1, +EDateD, +EdateH, +EdateMm);
	        var startDate = new Date(+SDateY, +SDateM - 1, +SDateD, +SDateH, +SDateMm);

	        if (startDate.getTime() > endDate.getTime()) {
	            alert(alertReason1);
	            document.getElementById("FM_slut_dag_2").value = document.getElementById("FM_start_dag_2").value;
	            document.getElementById("FM_slut_mrd_2").value = document.getElementById("FM_start_mrd_2").value;
	            document.getElementById("FM_slut_aar_2").value = document.getElementById("FM_start_aar_2").value;
	            return false;
	        }
	    }
	function insvaloskrift(){
	document.all["FM_navn_2"].value = document.all["FM_navn"].value;
	document.all["FM_navn_2"].focus()
}
function opret_hidden() {
    document.all["FM_oprethidden"].value = 'j';
}
	function insvalkpers(){
	document.all["FM_kpers_2"].value = document.all["FM_kpers"].value;
	document.all["FM_kpers_2"].focus()
	}
	function insvalstat(){
	document.all["FM_status_2"].value = document.all["FM_status"].value;
	document.all["FM_status_2"].focus()
	}
	function insvalkform(){
	document.all["FM_kform_2"].value = document.all["FM_kform"].value;
	document.all["FM_kform_2"].focus()
	}
	
	function showakt1(){
	document.all["dato2"].style.visibility = "hidden";
	
	
	if (document.all["FM_opf2"].checked == true) {
	document.all["sd"].style.visibility = "hidden";
	document.all["div1B"].style.display = "none";
	document.all["div1B"].style.visibility = "hidden";
	}
	
	document.all["divak1"].style.top = "137";
	document.all["divak2"].style.top = "136";
	document.all["divak3"].style.top = "136";
	
	//document.all["divak1"].style.border = "1px red solid"
	
	document.all["div0A"].style.display = "";
	document.all["div0A"].style.visibility = "visible";
	document.all["div1A"].style.display = "";
	document.all["div1A"].style.visibility = "visible";
	document.all["div2A"].style.display = "";
	document.all["div2A"].style.visibility = "visible";
	
	document.all["div0B"].style.display = "none";
	document.all["div0B"].style.visibility = "hidden";
	document.all["div2B"].style.display = "none";
	document.all["div2B"].style.visibility = "hidden";
	
	document.all["div1C"].style.display = "none";
	document.all["div1C"].style.visibility = "hidden";
	}
	
	function showakt2(){
	    document.all["dato2"].style.visibility = "visible";
	    document.all["dato2"].style.display = "";
	
	document.all["div0A"].style.display = "none";
	document.all["div0A"].style.visibility = "hidden";
	document.all["div1A"].style.display = "none";
	document.all["div1A"].style.visibility = "hidden";
	document.all["div2A"].style.display = "none";
	document.all["div2A"].style.visibility = "hidden";
	
	document.all["div0B"].style.display = "";
	document.all["div0B"].style.visibility = "visible";
	document.all["div1B"].style.display = "";
	document.all["div1B"].style.visibility = "visible";
	document.all["div2B"].style.display = "";
	document.all["div2B"].style.visibility = "visible";
	
	document.all["divak3"].style.visibility = "hidden";
	document.all["div1C"].style.display = "none";
	document.all["div1C"].style.visibility = "hidden";
	}
	
	function showopf(){
	document.all["dato2"].style.visibility = "visible";
	document.all["sd"].style.visibility = "hidden";
	
	document.all["divak2"].style.visibility = "visible";
	document.all["divak1"].style.top = "136";
	document.all["divak2"].style.top = "137";
	
	document.all["div0A"].style.display = "none";
	document.all["div0A"].style.visibility = "hidden";
	document.all["div1A"].style.display = "none";
	document.all["div1A"].style.visibility = "hidden";
	document.all["div2A"].style.display = "none";
	document.all["div2A"].style.visibility = "hidden";
	
	document.all["div0B"].style.display = "";
	document.all["div0B"].style.visibility = "visible";
	document.all["div1B"].style.display = "";
	document.all["div1B"].style.visibility = "visible";
	document.all["div2B"].style.display = "";
	document.all["div2B"].style.visibility = "visible";
	
	document.all["divak3"].style.visibility = "hidden";
	document.all["div1C"].style.display = "none";
	document.all["div1C"].style.visibility = "hidden";
	}
	
	function hideopf(){
	document.all["dato2"].style.visibility = "hidden";
	document.all["sd"].style.visibility = "hidden";
	
	document.all["divak2"].style.visibility = "hidden";
	document.all["divak1"].style.top = "137";
	
	document.all["div0A"].style.display = "";
	document.all["div0A"].style.visibility = "visible";
	document.all["div1A"].style.display = "";
	document.all["div1A"].style.visibility = "visible";
	document.all["div2A"].style.display = "";
	document.all["div2A"].style.visibility = "visible";
	
	document.all["div0B"].style.display = "none";
	document.all["div0B"].style.visibility = "hidden";
	document.all["div1B"].style.display = "none";
	document.all["div1B"].style.visibility = "hidden";
	document.all["div2B"].style.display = "none";
	document.all["div2B"].style.visibility = "hidden";
	
	document.all["divak3"].style.visibility = "hidden";
	document.all["div1C"].style.display = "none";
	document.all["div1C"].style.visibility = "hidden";
	}
	
	function visakt(){
	document.all["andreakt"].style.display = "";
	document.all["andreakt"].style.visibility = "visible";
	}
	
	function hideakt(){
	document.all["andreakt"].style.display = "none";
	document.all["andreakt"].style.visibility = "hidden";
	}
	
	function showendd(){
	document.all["dato2"].style.visibility = "hidden";
	document.all["sd"].style.visibility = "visible";
	
	document.all["divak2"].style.visibility = "hidden";
	document.all["divak1"].style.top = "137";
	
	
	document.all["div0A"].style.display = "";
	document.all["div0A"].style.visibility = "visible";
	document.all["div1A"].style.display = "";
	document.all["div1A"].style.visibility = "visible";
	document.all["div2A"].style.display = "";
	document.all["div2A"].style.visibility = "visible";
	
	document.all["div0B"].style.display = "none";
	document.all["div0B"].style.visibility = "hidden";
	document.all["div1B"].style.display = "";
	document.all["div1B"].style.visibility = "visible";
	document.all["div2B"].style.display = "none";
	document.all["div2B"].style.visibility = "hidden";
	
	document.all["divak3"].style.visibility = "hidden";
	document.all["div1C"].style.display = "none";
	document.all["div1C"].style.visibility = "hidden";
	}
	
	function showreocc(){
	document.all["dato2"].style.visibility = "hidden";
	document.all["sd"].style.visibility = "hidden";
	
	document.all["divak2"].style.visibility = "hidden";
	document.all["divak3"].style.visibility = "visible";
	document.all["divak1"].style.top = "136";
	document.all["divak3"].style.top = "137";
	
	document.all["div0A"].style.display = "none";
	document.all["div0A"].style.visibility = "hidden";
	document.all["div1A"].style.display = "none";
	document.all["div1A"].style.visibility = "hidden";
	document.all["div2A"].style.display = "none";
	document.all["div2A"].style.visibility = "hidden";
	
	document.all["div0B"].style.display = "none";
	document.all["div0B"].style.visibility = "hidden";
	document.all["div1B"].style.display = "none";
	document.all["div1B"].style.visibility = "hidden";
	document.all["div2B"].style.display = "none";
	document.all["div2B"].style.visibility = "hidden";
	
	document.all["div1C"].style.display = "";
	document.all["div1C"].style.visibility = "visible";
	}
	
	function NewWin_popupcal(url)    {
		window.open(url, 'Calpick', 'width=250,height=250,scrollbars=no,toolbar=no');    
		}
	</script>
	<%
	showinwin = request("showinwin")
	
	if showinwin <> "j" then%>
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	<%end if
	
	if func = "red" then
	'Response.write strKlokkeslet
	
	formatstrKlokkeslet = FormatDateTime(strKlokkeslet,4)
	klok_timer = left(formatstrKlokkeslet,2)
	klok_min = right(formatstrKlokkeslet,2)
	
	formatstrKlokkeslet_slut = FormatDateTime(strKlokkeslet_slut,4)
	klok_timer_slut = left(formatstrKlokkeslet_slut,2)
	klok_min_slut = right(formatstrKlokkeslet_slut,2)
	%>
	<!--#include file="inc/dato2.asp"-->
	<% 
	else
	
	if len(request("strdag")) <> 0 then
	strDag = request("strdag")
	strMrd = request("strmrd")
	strAar = request("straar")
	else
	strDag = day(now)
	strMrd = month(now)
	strAar = year(now)
	end if
	%>
	<!--#include file="inc/dato2_b.asp"-->
	<%
	
	
	klok_timer = left(kl, 2)
	if len(klok_timer) <> 0 then
	klok_timer = klok_timer
	else
	klok_timer = left(FormatDateTime(time,4),2)
	end if
	klok_timer_slut = klok_timer
	end if
	
	if request("selpkt") = "kal" then
	selpkt = "kal"
	kalvalues = "&kaldato="&request("kaldato")&"&kalstatus="&request("kalstatus")&"&kalemner="&request("kalemner")&"&kalmedarb="&request("kalmedarb")&"&kalviskunde="&request("kalviskunde")&""
	else
	selpkt = "hist"
	kalvalues = " "
	end if
	
	if showinwin <> "j" then
	divleft = 190
	divtop = 50
	else
	divleft = 10
	divtop = 20
	end if
	%>
	<!-------------------------------Sideindhold------------------------------------->
	<% if func = "opret" or func = "dbopr" or func = "red" then
	strBG = " background-color:#FFFFFF; "
	end if %>
	<div id="sindhold"  style="<%=strBG%>absolute; left:<%=divleft%>; top:<%=divtop%>; visibility:visible;">
	<%if showinwin <> "j" then%>
	<table cellspacing="0" cellpadding="0" border="0">
	<tr>
    <td valign="top"><br>
	<img src="../ill/header_crmosigt.gif" alt="" border="0"><hr align="left" width="360" size="1" color="#000000" noshade>
	</td></tr></table>
	<br>
	<%end if%>
	<table cellspacing="0" cellpadding="0" border="0" width="360" bgcolor="#FFFFFF">
	<form action="crmhistorik.asp?menu=crm&func=<%=dbfunc%>&ketype=e&selpkt=<%=selpkt&kalvalues%>" name="aktionsform" method="post">
	<input type="hidden" name="crmaktion" value="<%=crmaktion%>">
	<input type="hidden" name="showinwin" value="<%=showinwin%>">
	<tr bgcolor="#FFFFFF">
		<td width="8" rowspan="2" valign=top>&nbsp;</td>
		<td colspan=3 valign="top"></td>
		<td align=right rowspan="2" valign=top></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="3" style="height:20;">
		<b>CRM aktion -- &nbsp;<%=varbroedkrumme%></b></td>
	</tr>
	<%if dbfunc = "dbred" then%>
	<tr>
		<td valign="top"></td>
		<td colspan="3"><br>Sidst opdateret den <b><%=formatdatetime(strDato, 1)%></b> af <b><%=strEditor%></b></td>
		<td valign="top" align=right></td>
	</tr>
    <%elseif func = "opret" then%>
    <tr style="padding-bottom:3px;">
    <%thiskri = Request.Form("FM_soeg") %>
        <td colspan="2">Navn, kundenr. el. tlf. nr.</td><td colspan="3"><input type="text" name="FM_soeg" id="FM_soeg" value="<%=thiskri%>" size="30" style="font-family: arial,helvetica,sans-serif; font-size: 10px;"></td></tr>
        <tr style="padding-bottom:3px;">
        <td colspan="2">Kundeans.</td><td colspan="3"><select name="medarb" size="1" style="font-family: arial,helvetica,sans-serif; font-size: 10px; width:180;">
				<option value="0">Alle</option>
				<%
						strSQL = "SELECT mnavn, mid FROM medarbejdere ORDER BY mnavn"
						oRec.open strSQL, oConn, 3
						while not oRec.EOF
						
							if (cint(medarb) = cint(oRec("mid")) And Request.Form("medarb") = "") or (cint(oRec("mid")) = cint(Request.Form("medarb"))) then
							isSelected = "SELECTED"
							else
							isSelected = ""
							end if
						
						%>
						<option value="<%=oRec("mid")%>" <%=isSelected%>><%=oRec("mnavn")%></option>
						<%
						oRec.movenext
						wend
						oRec.close
						%>
				</select></td></tr>
				<tr style="padding-bottom:3px;" >
        <td colspan="2">Type:</td><td colspan="3"> <select name="ktype" size="1" style="font-family: arial,helvetica,sans-serif; font-size: 10px; width:100;">
				<option value="0">Alle</option>
				<%
						strSQL = "SELECT navn, id FROM kundetyper ORDER BY navn"
						oRec.open strSQL, oConn, 3
						while not oRec.EOF
						
							if (cint(ktype) = cint(oRec("id")) And Request.Form("ktype") = "") or (cint(oRec("id")) = cint(Request.Form("ktype"))) then
							isSelected = "SELECTED"
							else
							isSelected = ""
							end if
						
						%>
						<option value="<%=oRec("id")%>" <%=isSelected%>><%=left(oRec("navn"), 15)%></option>
						<%
						oRec.movenext
						wend
						oRec.close
						%>
				</select></td>
    </tr><tr>
    <td colspan="5"><center><input id="SoegSubmit" type="submit" value="Find kontakter>>" /></center></td></tr>
	<%end if
	
	table_topgif = "343"
	
	%>
	<tr>
	<% if func = "red" then
	strDisabled = "disabled"
	else 
	strDisabled = null
	end if %>
	<td style="padding-top:5; padding-left:30px;" colspan="5"><br>Kunde:&nbsp;&nbsp;<select name="id" onChange="document.aktionsform.submit();" style="width:200px;" <%=strDisabled %>>
		<%
		useMedarb = request.Form("medarb")
		if useMedarb <> "" then
		useMedarbKri = " AND (kunder.kundeans1 = "& useMedarb &" OR kunder.kundeans2 = "& useMedarb &")"
		end if
		ktype = Request.Form("ktype")
		if ktype > 0 then
		typeKri = " AND ktype = "&ktype&" "
		else
		ktype = null
		end if
		if thiskri <> "" then
		sqlsearchKri = " (Kkundenavn LIKE '"& thiskri &"%' OR (Kkundenr = '"& thiskri &"' AND kkundenr <> '0') OR telefon = '"& thiskri &"')"
		else
		sqlsearchKri = "ketype <> 'k'"
		end if
		strSQL = "SELECT Kkundenavn, adresse, postnr, city, land, Kkundenr, email, Kid FROM kunder WHERE "& sqlsearchKri & typeKri & useMedarbKri & " ORDER BY Kkundenavn"
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		strKid = oRec("Kid")
		if strKid = cint(id) then
		strSelVal = "selected"
		else
		strSelVal = Null
		end if
		%>
		<option value="<%=strKid%>" <%=strSelVal%>><%=left(oRec("Kkundenavn"), 30)%></option>
		<%
		if strSelVal = "selected" then
		GoogleStringKunde = oRec("Kkundenavn")
		GoogleStringEmail = oRec("email")
		
	    if oRec("adresse") <> "" then
	    GoogleStringAdresse = GooglestringAdresse & oRec("adresse") & ", "
        end if
        if oRec("postnr") <> "" then
	    GoogleStringAdresse = GooglestringAdresse & oRec("postnr") & ", "
        end if   
        if oRec("city") <> "" then
	    GoogleStringAdresse = GooglestringAdresse & oRec("city") & ", "
        end if  
        if oRec("land") <> "" then
	    GoogleStringAdresse = GooglestringAdresse & oRec("land") 
        end if  
        end if
		oRec.movenext
		wend
		oRec.close
		%>
	</select>
	<br><br>&nbsp;</td>
	<% if func = "red" then
	strSQL = "SELECT crmklokkeslet, crmklokkeslet_slut, crmdato_slut, crmdato, crmklokkeslet_slut FROM crmhistorik Where id = " & crmaktion
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
Function toGT(klokkeslet, dato)
if Instr(klokkeslet," ") then
splitklokken = split(klokkeslet," ")
klokkeslet = splitklokken(1)
end if
    DatoSplit = split(dato,"-")
    Gdate = DatoSplit(2) & "-" & DatoSplit(1) & "-" & DatoSplit(0)
    Gtime = klokkeslet & ".000+01.00"
    Response.Write Gdate & "T" & Gtime
end function
	 %>
	<input type="hidden" name="GoogleStringOrgTime" value="<%call toGT(oRec("crmklokkeslet"),oRec("crmdato"))%>" />
	<input type="hidden" name="GoogleStringOrgTimeCalc" value="<%call toGT(oRec("crmklokkeslet_slut"),oRec("crmdato_slut"))%>" />
	<%end if
	oRec.close
	end if %>
	<input type="hidden" name="GoogleStringEmne" value="" />
	<input type="hidden" name="GoogleStringAdresse" value="<%=GoogleStringAdresse%>" />
	<input type="hidden" name="GoogleStringEmail" value="<%=GoogleStringEmail%>" />
	<input type="hidden" name="GoogleStringKunde" value="<%=GoogleStringKunde%>" />
	<input type="hidden" name="kl" value="<%=kl%>">
	<input type="hidden" name="strdag" value="<%=strDag%>">
	<input type="hidden" name="strmrd" value="<%=strMrd%>">
	<input type="hidden" name="straar" value="<%=strAar%>">
	</tr>
	</table>
	<%
	if id <> "" then
	
			'** Finder kunde ****
			strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE kid = " & id
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			 strKundenavn = oRec("Kkundenavn")
			end if
			oRec.close
	end if		
%>
			<input type="hidden" name="id" value="<%=id%>">
	</table>
	<%if func = "red" then
	vizi = "hidden" 
	else
	vizi = "visible"
	end if%>
	<div id=div0A name=div0A style="position:relative; visibility:visible; display: z-index:20;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#FFFFFF" width="360">
	<tr>
		<td width="8" rowspan="2" valign=top></td>
		<td colspan=3 valign="top"></td>
		<td align=right rowspan="2" valign=top></td>
	</tr>
	<tr >
		<td colspan="3" style="height:20; padding-left:10;" class=alt><b><%=strKundenavn%>&nbsp;aktion:</b>&nbsp;</td>
	</tr>
	</table>
	</div>
	
	<div id=andreakt name=andreakt style="position:absolute; left:250; top:240; background-color:#ffffff; width:100; z-index:100000; visibility:hidden; display:none; border:1px #000000 dashed; padding-left:5;">
	<a href="#" onClick="hideakt()" class='vmenu'>Luk vindue [X]</a><br><br>
				<%
				strSQL = "SELECT crmdato FROM crmhistorik WHERE serialnb = " & thisserialnb
				oRec.open strSQL, oConn, 3
				While not oRec.EOF
				Response.write oRec("crmdato") &"<br>"
				oRec.movenext
				Wend 
				oRec.close
				%>
	</div>
	
	
	<div id=div0B name=div0B style="position:relative; visibility:hidden; display:none; z-index:20;">
	<table cellspacing="0" cellpadding="0" border="0"  width="360">
	<tr bgcolor="#FFFFFF">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=3 valign="top"><img src="../ill/tabel_top.gif" width="<%=table_topgif%>" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="3" style="height:20; padding-left:10;" class=alt><b><%=strKundenavn%>&nbsp;2. aktion: (opfølgende aktion)</b>&nbsp;</td>
	</tr>
	</table>
	</div>
	
	<%if func = "red" then
	vizi1B = "visible"
	show1B = "" 
	else
	vizi1B = "hidden"
	show1B = "none"
	end if%>
	
	
	
	<div id=div1B name=div1B style="position:relative; visibility:<%=vizi1B%>; display:<%=show1B%>; z-index:20;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#FFFFFF" width="360">
	<tr>
		<td valign="top"></td>
		<td style="padding-left:10;" colspan="2" valign=top><br>
		<%
		'** Opfølgning / slut dato ***
		if func = "opret" OR func = "red" then
			
			if func = "opret" then 
			viz = "hidden"
			else
			viz = "visible"
			end if
			
			'if func = "red" AND cint(thisserialnb) = 0 then>
			'Aktionen løber over flere dage:
			'	<%
			'	if len(crmDato_slut) <> 0 then
			'		if cdate(crmDato_slut) <> cdate(crmDato) then 
			'		chk = "checked"
			'		else
			'		chk = ""
			'		end if
			'	else
			'	chk = ""
			'	end if
			'><b>Ja:</b><input type="checkbox" <%=chk> name="FM_opf" value="2"><br>
			'Slut dato:
			'<%
			'else
				
				if func = "red" AND cint(thisserialnb) <> 0 then
				%>
				<font color="darkred"><b>!</b></font>&nbsp;Denne aktion er del af en serie. Skal alle aktioner i denne serie opdateres?
				<b>Ja:</b><input type="checkbox" checked name="FM_opd_serie" value="1">
				<input type="hidden" name="FM_serialnb" value="<%=thisserialnb%>">&nbsp;&nbsp;<a href="#" class='vmenu' onclick=visakt()>Aktioner i denne serie...</a>
				<br><font color="#999999" size=1>Hvis ja kan dato ikke ændres!<br>
				Hvis nej vil denne aktion ikke længere være en del af en serie.</font><br>
				<%end if
		end if%>
		</td>
		<td>&nbsp;
		</td>
			<td valign="top" align="right"></td>
		</tr>
		</table>
		</div>
	
	<div id=div2A name=div2A style="position:relative; visibility:visible; display: z-index:20;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#FFFFFF" width="360">
	<tr>
	<td valign="top"></td><td style="padding-left:10;">Startdato:</td></td>
    <td colspan="2">

				<select id="FM_start_dag" name="FM_start_dag" Onchange="ValidateDate()">
				<%Call DropDownDay("FM_start_dag","crmdato") %></select>
				
				<select id="FM_start_mrd" name="FM_start_mrd" Onchange="ValidateDate()">
				<%Call DropDownMonth("FM_start_mrd","crmdato") %>
				</select>
				
				<select id="FM_start_aar" name="FM_start_aar" Onchange="ValidateDate()">
                <%Call DropDownYear("FM_start_aar","crmdato") %>
                </select>&nbsp;&nbsp;
				<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=1')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>
    </td>
	<td></td></tr>
	<tr>
		<td valign="top"></td>
		<td style="padding-left:10; width:100px;">Start tidspunkt:
		</td><td><select id="klok_timer" name="klok_timer" Onchange="ValidateDate()">
		<%Call DropDownHour("klok_timer","crmklokkeslet") %>
	  	</select>&nbsp;:
		
		<select id="klok_min" name="klok_min" Onchange="ValidateDate()">
        <%Call DropDownMinute("klok_min","crmklokkeslet") %>
		</select>
		</td>
		<td>&nbsp;</td>
		<td valign="top" align="right"></td>
		</tr>
		<tr valign="top"><td></td><td style="padding-left:10;">Slutdato:</td>
		<td colspan="2">				<select id="FM_slut_dag" name="FM_slut_dag" Onchange="ValidateDate()">
				<%Call DropDownDay("FM_slut_dag","crmdato_slut") %></select>
				
				<select id="FM_slut_mrd" name="FM_slut_mrd" Onchange="ValidateDate()">
				<%Call DropDownMonth("FM_slut_mrd","crmdato_slut") %>
				</select>
				
				<select id="FM_slut_aar" name="FM_slut_aar" Onchange="ValidateDate()">
                <%Call DropDownYear("FM_slut_aar","crmdato_slut") %>
                </select>&nbsp;&nbsp;
				<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=4')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a></td>
		<td></td></tr>
		<tr>
		<td valign="top"></td>
		<td style="padding-left:10;">Slut tidspunkt:
		</td><td>
		<select id="klok_timer_slut" name="klok_timer_slut" Onchange="ValidateDate()">
		<%Call DropDownHour("klok_timer_slut","crmklokkeslet_slut") %>
	  	</select>&nbsp;:
		
		<select id="klok_min_slut" name="klok_min_slut" Onchange="ValidateDate()">
		<%Call DropDownMinute("klok_min_slut","crmklokkeslet_slut") %>
		</select>
		</td>
		<td>&nbsp; 
		</td>
			<td valign="top" align="right"></td>
		</tr>
		
		<tr>
			<td valign="top"></td>
			<td style="padding-left:10;">Emne:</td>
			<td><select name="FM_emne" onchange="this.form.GoogleStringEmne.value=this.options[this.selectedIndex].text;">
			<%strSQL = "SELECT id, navn FROM crmemne ORDER BY navn"
			oRec.open strSQL, oConn, 3
			While not oRec.EOF %>
				<%if (intEmne = oRec("id") and Request.Form("FM_emne") = "") or (cint(Request.Form("FM_emne")) = cint(oRec("id")))  then
				selected = "selected"
                GoogleStringOrgEmne = oRec("navn")
				else
				selected = ""
				end if
				%> 
			<option value="<%=oRec("id")%>" <%=selected%>><%=oRec("navn")%></option>
			<%
			oRec.movenext
			wend
			oRec.close%>
		</select>
		
		</td>
		<td>&nbsp;</td>
		<td valign="top" align="right"></td>
		</tr>
		<tr>
		<td valign="top"></td>
		<td style="padding-left:10;">Hvad:</td>
		<% if Request.Form("FM_navn") <> "" then
		strNavn = Request.Form("FM_navn")
		end if %>
		<td>
		<input type="text" name="FM_navn" value="<%=strNavn%>" size="25">
		</td>
		<td>&nbsp;</td>
		<% if strNavn <> "" then
		strGoogleStringOrgName = strNavn & " - " & GoogleStringOrgEmne
		else
		strGoogleStringOrgName = GoogleStringOrgEmne
		end if %>
		<input type="hidden" name="GoogleStringOrgName" value="<%=strGoogleStringOrgName%>" />
		<td valign="top" align="right"></td>
	</tr>
	<tr>
		<td valign="top"></td>
		<td style="padding-left:10;">Kontakt person:
		<%if func = "red" then%>
		<!--<br>
		<font size=1 color=#999999>Hvis den rigtige person<br> ikke er forvalgt skyldes <br>det en struktur ændring<br> i TimeOut d. 25/1-2004.</font>-->
		<%end if
		if id = 0 then
		strSelDisabled = "disabled"
		else
		strSelDisabled = ""
		end if
		%>
		</td>
		<td valign="top"><select name="FM_kpers" <%=strSelDisabled%>>
		<%
		call erDetInt(intKontaktpers)
		if isInt > 0 then
		intKontaktpers = 0
		else
		intKontaktpers = intKontaktpers
		end if
		isInt = 0
		
		strSQL = "SELECT id, navn FROM kontaktpers WHERE kundeid = "& id
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		if (cint(intKontaktpers) = oRec("id") and Request.Form("FM_kpers") = "") or (cint(request.Form("FM_kpers")) = oRec("id")) then
		kpersSel = "SELECTED"
		else
		kpersSel = ""
		end if
		%>
		<option value="<%=oRec("id")%>" <%=kpersSel%>><%=oRec("navn")%></option>
		<%oRec.movenext 
		wend
		oRec.close%>
		<option value="0">Ingen</option>
	</select></td>
	<td>&nbsp;</td>
	<td valign="top" align="right"></td>
	</tr>
	<tr>
		<td valign="top"></td>
		<td valign=top colspan="3" style="padding-left:10;">Beskrivelse/ Log:<br>
		<% if request.Form("FM_besk") <> "" then
		strBesk = request.Form("FM_besk")
		end if %>
		<textarea cols="<%=trearsize%>" rows="8" name="FM_besk"><%=strBesk%></textarea>
		&nbsp;</td>
		<td valign="top" align="right"></td>
	</tr>
	<tr>
		<td valign="top"></td>
		<td style="padding-left:10;">Status:</td>
		<td><select name="FM_status">
		<%strSQL = "SELECT id, navn FROM crmstatus ORDER BY navn"
		oRec.open strSQL, oConn, 3
		While not oRec.EOF
		if (intStatus = oRec("id") AND request.Form("FM_status") = "") or cint(oRec("id")) = cint(request.Form("FM_status")) then
		selected = "selected"
		else
		selected = ""
		end if%> 
		<option value="<%=oRec("id")%>" <%=selected%>><%=oRec("navn")%></option>
		<%oRec.movenext
		wend
		oRec.close%>
		</select></td>
		<td>&nbsp;</td>
		<td valign="top" align="right"></td>
		</tr>
	<tr>
		<td valign="top"></td>
		<td style="padding-left:10;">Kontaktform:</td>
		<td><select name="FM_kform">
			<%strSQL = "SELECT id, navn FROM crmkontaktform ORDER BY navn"
			oRec.open strSQL, oConn, 3
			While not oRec.EOF %>
			<%if (intKontaktform = oRec("id") AND Request.Form("FM_kform") = "") or cint(oRec("id")) = cint(Request.Form("FM_kform")) then
			selected = "selected"
			else
			selected = ""
			end if%> 
			<option value="<%=oRec("id")%>" <%=selected%>><%=oRec("navn")%></option>
			<%oRec.movenext
			wend
			oRec.close%>
		</select></td>
		<td>&nbsp;</td>
		<td valign="top" align="right"></td>
	</tr>
	<tr>
		<td valign="top"></td>
		<td colspan="3" style="padding-left:10;"><br>Tilføj medarbejdere:</td>
		<td valign="top" align="right"></td>
	</tr>
	<%
		strSQL = "SELECT mid, mnavn FROM medarbejdere WHERE (Brugergruppe = 1 OR Brugergruppe = 3 OR Brugergruppe = 6 OR Brugergruppe = 8) AND mansat <> '2' ORDER BY mnavn"
		oRec.open strSQL, oConn, 0, 1
		while not oRec.EOF
		
			if func = "red" then
				strSQL2 = "SELECT medarbid FROM aktionsrelationer WHERE aktionsid = "& crmaktion &" AND medarbid = "& oRec("mid")
				oRec2.open strSQL2, oConn, 0, 1
				if not oRec2.EOF then
				strcheckmedarb = "CHECKED"
				else
				strcheckmedarb = " "
				end if
				oRec2.close
			else
			if instr(request.Form("FM_medarbrel"),",") then
			CheckMid = split(Request.Form("FM_medarbrel"),",")
			for i = 0 to Ubound(CheckMid)
			if cint(CheckMid(i)) = cint(oRec("mid")) then
			strCheckit = "y"
			end if
			Next
			elseif cint(oRec("mid")) = cint(Request.Form("FM_medarbrel")) then
			strCheckit = "y"
			end if
				if (oRec("mnavn") = session("user") AND request.Form("FM_medarbrel") = "") or strCheckit = "y" then
				strcheckmedarb = "CHECKED"
				else
				strcheckmedarb = " "
				end if
				strCheckit = Null
				CheckMid = Null
		    end if
		%> 
		<tr>
			<td colspan="2" style="padding-left:10;"><%=oRec("mnavn")%></td><td colspan="3" style="padding-left:10;"><input type="checkbox" name="FM_medarbrel" value="<%=oRec("mid")%>" <%=strcheckmedarb%> style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
			</td>
		</tr>
		<%
		x = x + 1
		oRec.movenext
		wend
		oRec.close
		%>
		</table>
		</div>
		
	<table cellspacing="0" cellpadding="0" border="0" width="360">
	<tr>
		<td valign="top"></td>
		<td colspan=3 valign="bottom"></td>
		<td valign="top" align="right"></td>
	</tr>
	<tr>
		<td colspan="5"><br><br><img src="ill/blank.gif" width="120" height="1" alt="" border="0"><input type="image" name="subm" onclick="opret_hidden();" src="../ill/<%=varSubVal%>.gif"></td>
	</tr>
	<input type="hidden" name="FM_oprethidden" value="" />
	</form>
	</table>
	<br>
	</div>
	
	
	
	
	
	<%case else
	emner = request("emner")
	status = request("status")
	medarb = request("medarb")
	
	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
	end function
	
	'** Hvis der søges på telefon nummer **
	if len(request("sogekri")) <> 0 then
	sogekri = trim(SQLBless(request("sogekri")))
	kundeKritlf = " Kkundenavn LIKE '%"& sogekri &"%' OR telefon LIKE '" & sogekri &"%'"_
	&" OR telefonmobil LIKE '"& sogekri &"%'"_
	&" OR telefonalt LIKE '"& sogekri & "%' OR kunder.email LIKE '%"& sogekri &"%'"
	
	kundeKritlf2 = "kontaktpers.email LIKE '%"& sogekri & "%' OR kontaktpers.navn LIKE '%"& sogekri & "%'"_
	&" OR kontaktpers.dirtlf LIKE '%"& sogekri & "%' OR kontaktpers.mobiltlf LIKE '%"& sogekri & "%'"_
	&" OR kontaktpers.mobiltlf LIKE '%"& sogekri & "%'"
	
	else
	sogekri = ""
	kundeKritlf = " Kid =" & id &""
	end if
	%>
	
	
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	
	
	
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call crmmainmenu(3)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	
	</div>
	
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:127;">
	<h3>Aktions historik</h3>
	
	
	
	<%call filterheader(0,0,720,pTxt)%>
	<table cellspacing="1" cellpadding="2" border="0" width="100%">
	<form action="crmhistorik.asp?menu=crm&func=hist&selpkt=hist" method="post">
   	
	<tr>
		<td width=100><b>Kontakter:</b>&nbsp; </td><td><select name="id" size="1" style="font-size : 9px; width:200 px;" onchange="submit()";>
		<option value="0">Alle</option>
		
		<%
		ketypeKri = " ketype <> 'k' "
		
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE "& ketypeKri &" ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(request("id")) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=left(oRec("Kkundenavn"), 18)%></option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
		</select>
	
	
	
</td></tr>

<%
 if len(request("FM_hotnot")) <> 0 then
            hotnot = request("FM_hotnot")
            response.cookies("tsa")("hotnot") = hotnot
            response.cookies("tsa").expires = date + 5
            else
                
                if request.cookies("tsa")("hotnot") <> "" then
                hotnot = request.cookies("tsa")("hotnot") 
                else
                hotnot = 99
                end if
            
            end if 

select case cint(hotnot) 
case -2
hotnotSEL1 = "SELECTED"
hotnotSEL2 = ""
hotnotSEL3 = ""
hotnotSEL4 = ""
hotnotSEL5 = ""
hotnotSEL0 = ""
case -1
hotnotSEL2 = "SELECTED"
hotnotSEL1 = ""
hotnotSEL3 = ""
hotnotSEL4 = ""
hotnotSEL5 = ""
hotnotSEL0 = ""
case 0
hotnotSEL3 = "SELECTED"
hotnotSEL2 = ""
hotnotSEL1 = ""
hotnotSEL4 = ""
hotnotSEL5 = ""
hotnotSEL0 = ""
case 1
hotnotSEL4 = "SELECTED"
hotnotSEL2 = ""
hotnotSEL3 = ""
hotnotSEL1 = ""
hotnotSEL5 = ""
hotnotSEL0 = ""
case 2
hotnotSEL5 = "SELECTED"
hotnotSEL2 = ""
hotnotSEL3 = ""
hotnotSEL4 = ""
hotnotSEL1 = ""
hotnotSEL0 = ""
case else
hotnotSEL5 = ""
hotnotSEL2 = ""
hotnotSEL3 = ""
hotnotSEL4 = ""
hotnotSEL1 = ""
hotnotSEL0 = "SELECTED"
end select

%>

<tr>
		<td width=100><b>Kontakt interesse niveau:</b>&nbsp; </td><td>
		<select name="FM_hotnot" size="1" style="font-size : 9px; width:200 px;" onchange="submit()";>
		<option value="99" <%=hotnotSEL0  %>>Alle</option>
		<option value="-2" <%=hotnotSEL1  %>>-2 (Not)</option>
		<option value="-1" <%=hotnotSEL2  %>>-1</option>
		<option value="0" <%=hotnotSEL3  %>>0</option>
		<option value="1" <%=hotnotSEL4  %>>+1</option>
		<option value="2" <%=hotnotSEL5  %>>+2 (Hot)</option>
				
		</select>
	
	
	
</td></tr>


<%

if level <= 2 then
%>

<tr>
<td><b>Medarbejder:</b>&nbsp;</td><td>
<select name="medarb" size="1" style="font-size : 11px; width:200 px;" onchange="submit()";>
<option value="0">Alle</option>
<%
		strSQL = "SELECT mnavn, mid FROM medarbejdere WHERE mansat <> '2' ORDER BY mnavn"
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		
			if cint(medarb) = cint(oRec("mid")) then
			isSelected = "SELECTED"
			else
			isSelected = ""
			end if
		
		%>
		<option value="<%=oRec("mid")%>" <%=isSelected%>><%=left(oRec("mnavn"), 15)%></option>
		<%
		oRec.movenext
		wend
		oRec.close
		%>
</select></td>
</tr>


<%end if%>



	<tr>
	<td><b>Emne:</b>&nbsp;</td><td>
	<select name="emner" size="1" style="font-size : 11px; width:200 px;" onchange="submit()";>
	<option value="0">Alle</option>
	<%
			strSQL = "SELECT navn, id FROM crmemne ORDER BY navn"
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
			
			if cint(request("emner")) = cint(oRec("id")) then
			isSelected = "SELECTED"
			else
			isSelected = ""
			end if
			%>
			<option value="<%=oRec("id")%>" <%=isSelected%>><%=left(oRec("navn"), 15)%></option>
			<%
			oRec.movenext
			wend
			oRec.close
			%>
	</select></td>
	
	</tr>
	


	<tr>
	<td><b>Status:</b> &nbsp;</td><td>
	<select name="status" size="1" style="font-size : 11px; width:200 px;" onchange="submit()";>
	<option value="0">Alle</option>
	<%
			strSQL = "SELECT navn, id FROM crmstatus ORDER BY navn"
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
			
			if cint(request("status")) = cint(oRec("id")) then
			isSelected = "SELECTED"
			else
			isSelected = ""
			end if
			%>
			<option value="<%=oRec("id")%>" <%=isSelected%>><%=left(oRec("navn"), 15)%></option>
			<%
			oRec.movenext
			wend
			oRec.close
			%>
	</select></td>
	
	</tr>
	

<tr>
			<!-- Brug altid datointerval, FM_usedatokri = 1 -->
			<input type="hidden" name="FM_usedatokri" id="FM_usedatokri" value="1">
			<td colspan=2 style="padding-top:5px; padding-left:5px;"><b>Periode:</b><br>
			<!--#include file="inc/weekselector_s.asp"--> <!-- b -->
			</td>
	</tr>
	<tr>
	<td colspan=2 align=right>
        <input id="Submit1" type="submit" value="Vis aktions historik >>" /></td>
	</tr>
	
	</form>
	</table>
	
	<!-- filter div -->
		</td></tr></table>
	</div>
	
	<%
	
	oleftpx = 0
	otoppx = 20
	owdtpx = 130
	url = "javascript:NewWin_popupaktion(""crmhistorik.asp?menu=crm&func=opret&id="&id&"&ketype=e&showinwin=j"")"
	call opretNy(url, "Opret ny aktion", otoppx, oleftpx, owdtpx) 
	%>
	
	<br /><br /><br />
	
	<%stDato = strAar &"/"& strMrd &"/"& strDag %>
	<%slDato = strAar_slut &"/"& strMrd_slut &"/"& strDag_slut %>
	
	
	<%
	
	
	
	
	x = 0
	
	'********************************************************************************************
	' Udskriver aktions historik liste
	'********************************************************************************************
	
    
    tTop = 20
	tLeft = 0
	tWdth = 1004
	
	
	call tableDiv(tTop,tLeft,tWdth)
	%>
	
	
	<table cellspacing="0" cellpadding="0" border="0" width="1004" bgcolor="#EFF3FF">
	<% 
	'************************************************************************************ 
	'Table header 
	'************************************************************************************
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030%>
	
	<tr bgcolor="#5582d2">
	    <td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	    <td colspan=8 class=alt ><b>Aktions Historik</b></td>
	    <td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/blank.gif" width="1" height="30" alt="" border="0"></td>
		<td height=20>&nbsp;&nbsp;<b>Firma</b></td>
		<td align=right style="padding-right:15px;"><b>Dato</b></td>
		<td colspan=2 width=200>Emne&nbsp;(¤ = serie)<br />
		<b>Navn</b><br />
		Kommentar</td>
		
		<td><b>Status</b></td>
		<td>&nbsp;</td>
		<td><b>Oprettet af</b></td>
		<td>&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/blank.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	select case level 
	case "1", "2"
			if medarb = 0 OR len(medarb) = 0 then
			usemedarbKri = "aktionsid = ch.id AND medarbid <> 0"
			else
			usemedarbKri = "aktionsid = ch.id AND medarbid = '" & medarb & "'"
			end if
	case else
	usemedarbKri = "aktionsid = ch.id AND medarbid = '" & session("Mid") & "'"
	end select
	
	if id = 0 OR len(id) = 0 then
	useidKri = " kundeid <> " & id & ""
	else
	useidKri = " kundeid = " & id 
	end if
	
	if emner = 0 OR len(emner) = 0 then
	useemneKri = " kontaktemne <> -1"
	else
	useemneKri = " kontaktemne = " & emner 
	end if
	
	if status = 0 OR len(status) = 0 then
	usestatKri = " AND status <> -1 "
	else
	usestatKri = " AND status = " & status 
	end if
	
	if cint(hotnot) = 99 then
	hotnotKri = " AND hot <> 99 "
	else
	hotnotKri = " AND hot = "& hotnot
	end if
	
	
	function kpers_her()
	
	if len(oRec("kontaktpers")) <> 0 then
	kpersId = oRec("kontaktpers")
	else
	kpersId = 0
	end if
	
				strSQL2 = "SELECT id, navn FROM kontaktpers WHERE id = "& kpersId
				
			    
			    'Response.Write strSQL2
	            'Response.flush
	            
	            oRec2.open strSQL2, oConn, 3
			    
				if not oRec2.EOF then
					if len(oRec2("navn")) > 25 then
					Response.write left(oRec2("navn"), 25) &".. "
					else
					Response.write oRec2("navn") 
					end if
				else
				Response.write "Ingen"
				end if
				oRec2.close
	
	end function
	
	
	
	
	
	strSQL = "SELECT DISTINCT(ch.id), kid, kkundenavn, ch.editor, crmdato, kontaktemne, ch.navn, crmemne.navn AS enmenavn, "_
	&" crmstatus.navn AS statusnavn, crmkontaktform.navn AS kontaktform, "_
	&" ikon, ch.kontaktpers, ch.serialnb, telefon, ch.komm "_
	&" FROM kunder k "_
	&" LEFT JOIN crmhistorik ch ON (ch.crmdato BETWEEN '"& stDato &"' AND '"& slDato &"' AND ch.kundeid = k.kid)"_
	& "LEFT JOIN aktionsrelationer ON ("& usemedarbKri &")"_
	&" LEFT JOIN crmemne ON (crmemne.id = kontaktemne)"_
	&" LEFT JOIN crmstatus ON (crmstatus.id = status)"_
	&" LEFT JOIN crmkontaktform ON (crmkontaktform.id = kontaktform)"_
	&" WHERE ("& useidKri &" "& hotnotKri &") "& usestatKri &""_
	&" AND "& useemneKri &" "_
	&" ORDER BY ch.crmdato DESC"
	
	
	c = 0
	
	'Response.Write strSQL
	'Response.flush
	'Response.end
	
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	'************************************************************************************ 
	'Table Rows indhold 
	'************************************************************************************
	
	select case right(c, 1)
	case 2,4,6,8,0
	bgcol = "#ffffff"
	case else
	bgcol = ""
	end select
	
	%>
	<tr>
		<td bgcolor="#cccccc" colspan="10"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bgcol %>">
		<td valign="top"><img src="../ill/blank.gif" width="1" height="42" alt="" border="0"></td>
		<td valign="top" height=20 style="padding-left:5px;"><a href="kunder.asp?menu=crm&func=red&id=<%=oRec("Kid")%>"><%
		if len(oRec("kkundenavn")) > 20 then
		Response.write left(oRec("kkundenavn"), 20)&".."
		else
		Response.write oRec("kkundenavn")
		end if%></a>
		<br><font size='1' color='#000000'>tlf: <%=oRec("telefon")%><font size='1' color='#999999'><br>
		<%
		'**************************************************
		'* finder kontaktpers afhængig af om aktionen er
		'* fra før der belv oprettet kontaktpers tabel.
		'**************************************************
		'if oRec("crmdato") < cdate("1-1-2004") then
			call erDetInt(oRec("kontaktpers"))
			if isInt > 0 then
				if len(oRec("kontaktpers")) > 25 then
				Response.write left(oRec("kontaktpers"), 25) &".." 
				else
				Response.write oRec("kontaktpers") 
				end if
			else
			
				call kpers_her
			end if
			isInt = 0
			
		'else
		
			'call kpers_her
		'end if
		
		if IsNull(trim(oRec("enmenavn"))) then
		strEmne = "ej angivet"
		else
		strEmne = oRec("enmenavn")
		end if
		if isNull(trim(oRec("statusnavn"))) then
		strStatus = "ej angivet"
		else
		strStatus = left(oRec("statusnavn"), 16)
		end if
		%></td>
		
		<td valign="top" align=right style="padding-right:15px;"><%=formatdatetime(oRec("crmdato"), 1)%></td>
		<td valign="top" colspan=2>
		<a href="javascript:NewWin_popupaktion('crmhistorik.asp?selpkt=hist&menu=crm&func=red&id=<%=oRec("Kid")%>&crmaktion=<%=oRec("id")%>&showinwin=j')" class='vmenu'><%=strEmne%></a>
		&nbsp;
		<%if oRec("serialnb") <> 0 then%>
		¤&nbsp;
		<%end if%>
		<br />
		<b><%=oRec("navn")%></b>
		<br />
		<%=oRec("komm") %>&nbsp;&nbsp;</td>
		<td valign="top"><font size="1"><%=strStatus%></td>
		<td valign="top">&nbsp;&nbsp;<img src="../ill/<%=oRec("ikon")%>" border="0" alt="<%=oRec("kontaktform")%>"></td>
		<td valign="top"><font size="1" color="#999999"><%=left(oRec("editor"), 16)%></td>
		<td><a href="crmhistorik.asp?selpkt=hist&menu=crm&func=slet&id=<%=oRec("Kid")%>&crmaktion=<%=oRec("id")%>&ketype=e&emner=<%=emner%>&status=<%=status%>&medarb=<%=medarb%>"><img src="../ill/slet_16.gif" alt="Slet aktion" border="0"></a></td>
		<td valign="top" align="right"><img src="../ill/blank.gif" width="1" height="42" alt="" border="0"></td>
	</tr>
	<%
	c = c + 1
	x = 0
	oRec.movenext
	wend
	%>
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="8" valign="bottom"><img src="../ill/tabel_top.gif" width="721" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>	
	</table>
	
	<!-- table div slut -->
		</div>
	
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	
<%end select%><%end if%><!--#include file="../inc/regular/footer_inc.asp"-->