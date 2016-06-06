<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/isint_func.asp"-->

<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	kl = request("kl")
	crmaktion = request("crmaktion")
	thisfile = "crmhistorik"
	
	rdir = request("rdir")
	
	select case func
	case "opdaterliste" 


            uid = split(request("FM_id"), ",")
            ustatus = split(request("FM_status"), ",")
            medarb = request("medarb")

           

            for u = 0 to UBOUND(uid)
	           
            			
                strSQLupd = "UPDATE crmhistorik SET status = "& ustatus(u) &" "_
	            &" WHERE id = " & uid(u)
	            oConn.execute(strSQLupd)
            	
	            'Response.write strSQLupd & "<br>"
            	
	            
            	
            next

            'Response.Write "her"
            'Response.flush
            Response.redirect "crmhistorik.asp?menu=crm&ketype=e&func=hist&selpkt=hist&medarb="&medarb&"rdir="&rdir
            'Response.end
      
	
	
	
	
	
	
	case "slet"
	serial = request("serial")
	'*** Her spørges om det er ok at der slettes en aktion ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	
	<%
	
	slttxtalt = ""
	slturlalt = ""
	
	slttxt = "Du er ved at <b>slette</b> en aktion. Er dette korrekt?"
    slturl = "crmhistorik.asp?menu=crm&func=sletok&selpkt="&request("selpkt")&"&crmaktion="&crmaktion
    	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,210,100)
	
	
	case "sletok"
	'serial = request("serial")
	
	'*** Her slettes en crmaktion ***
	'if cint(serial) <> 0 then
	'oConn.execute("DELETE FROM crmhistorik WHERE serialnb = "& serial &"")
	'else
	oConn.execute("DELETE FROM crmhistorik WHERE id = "& crmaktion &"")
	'end if
	
	if request("selpkt") <> "kal" then
	Response.redirect "crmhistorik.asp?menu=crm&ketype=e&func=hist&selpkt=hist"
	'emner="&request("emner")&"&status="&request("status")&"&medarb="&request("medarb")&""
	else
	Response.redirect "crmkalender.asp?menu=crm&id="&id&"&medarb="&request("medarb")&"&selpkt=kal"
	end if
	
	
	
	
	
	'**********************************************************************************************'
	'** Her indsættes en ny aktion i db 													   ****'
	'**********************************************************************************************'
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
			intKundeId = request("id")
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
			strCrmdato_slut = request("FM_start_aar_2")&"/"&request("FM_start_mrd_2")&"/"&request("FM_start_dag_2")
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
		if func = "dbopr" then
		oConn.execute("INSERT INTO crmhistorik (editor, dato, crmdato, kontaktemne, kontaktpers, status, komm, navn, kundeid, kontaktform, editorid, crmklokkeslet, crmKlokkeslet_slut, crmdato_slut, serialnb) VALUES ('"& strEditor &"', '"& strDato &"', '"& thisfuncdato &"', "& strEmne &", '"& strKpers &"', '"& strStatus &"', '"& strBesk &"', '"& strNavn &"', "& intKundeId &", "& strKontaktform &", "& intMid &", '"& strKlokkeslet &"', '"& strKlokkeslet_slut &"', '"& thisfuncslutdato &"', "& thisserial &")")
			
			strSQL = "SELECT id FROM crmhistorik ORDER BY id DESC"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
				aktionsid = oRec("id")
			end if
			oRec.close
		else
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
			intKundeId_2 = request("id")
			strBesk_2 = SQLBless(request("FM_besk_2"))
			strEditor_2 = session("user")
			strDato_2 = year(now)&"/"&month(now)&"/"&day(now)
			strCrmdato_2 = request("FM_start_aar_2")&"/"&request("FM_start_mrd_2")&"/"&request("FM_start_dag_2")
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
		
		
		
		'*** Skærm ***
		'************************************************************************************************
		'** Oprettelse /redigering slut. Viser ok side til bruge eller									*
		'** Retunerer til crmkalender hvis det eren personlig note. 									*
		'************************************************************************************************
			if personal <> "y" then
			%>
			<html>
			<head>
			<title>TimeOut - Tid & Overblik</title>
			<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print.css">
			</head>
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
			
			Response.end
			
			else
			
			
			'Select case rdir
			'case "crmhist"
			Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
            'case else
			'Response.Write("<script language=""JavaScript"">window.opener.location.href('kunder.asp?menu=crm&shokselector=1&ketype=e&selpkt=osigt');</script>")
            'end select
			                
			Response.Write("<script language=""JavaScript"">window.close();</script>")
			
			
			
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
	
	'************************************************************************************************
	'*** Her indlæses form til rediger/oprettelse af ny aktion.										*
	'************************************************************************************************
	'** local kr/dato/nummer format
	session.LCID = 1030
	
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	trearsize = "40"
	'varSubVal = "opretpil" 
	'varbroedkrumme = "Opret ny aktion"
		
		'** Hvis der først skal vælges kunde ***
		if func = "opret" AND id = 0 then
		dbfunc = "opret"
		varSubVal = "step2pil" 
		varbroedkrumme = "Opret ny aktion step 1/2"
		else
		dbfunc = "dbopr"
		varSubVal = "opretpil"
		varbroedkrumme = "Opret ny aktion"
		end if
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
		<body topmargin="0" leftmargin="0" class="regular">
	<%end if%>
	
	<script LANGUAGE="javascript">
	function insvaloskrift(){
	document.all["FM_navn_2"].value = document.all["FM_navn"].value;
	document.all["FM_navn_2"].focus()
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
	document.all["sd"].style.visibility = "hidden";
	
	document.all["divak1"].style.top = "136";
	document.all["divak2"].style.top = "137";
	
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
	id = id
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
	<div id="sindhold" style="position:absolute; left:<%=divleft%>; top:<%=divtop%>; visibility:visible;">
	<%if showinwin <> "j" then%>
	<table cellspacing="0" cellpadding="0" border="0">
	<tr>
    <td valign="top"><br>
	<img src="../ill/header_crmosigt.gif" alt="" border="0"><hr align="left" width="360" size="1" color="#000000" noshade>
	</td></tr></table>
	<br>
	<%end if%>
	<table cellspacing="0" cellpadding="0" border="0" width="360" bgcolor="#EFF3FF">
	<form action="crmhistorik.asp?menu=crm&func=<%=dbfunc%>&ketype=e&selpkt=<%=selpkt&kalvalues%>&rdir=<%=rdir%>" method="post">
	<input type="hidden" name="crmaktion" value="<%=crmaktion%>">
	<input type="hidden" name="showinwin" value="<%=showinwin%>">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=3 valign="top"><img src="../ill/tabel_top.gif" width="343" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan="3" style="height:20;" class=alt>
		<b>CRM aktion -- &nbsp;<%=varbroedkrumme%></b></td>
	</tr>
	<%if dbfunc = "dbred" then%>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td colspan="3"><br>Sidst opdateret den <b><%=formatdatetime(strDato, 1)%></b> af <b><%=strEditor%></b></td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%end if
	
	table_topgif = "343"
	
	if func = "opret" AND id = 0 then
	%>
	<tr>
	<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="70" alt="" border="0"></td>
	<td style="padding-top:5;" colspan="3"><br>Kunde:&nbsp;&nbsp;<select name="id">
		<%
		strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'k' ORDER BY Kkundenavn"
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		%>
		<option value="<%=oRec("Kid")%>"><%=left(oRec("Kkundenavn"), 30)%></option>
		<%
		oRec.movenext
		wend
		oRec.close
		%>
	</select>
	<br><br>&nbsp;</td>
	<input type="hidden" name="kl" value="<%=kl%>">
	<input type="hidden" name="strdag" value="<%=strDag%>">
	<input type="hidden" name="strmrd" value="<%=strMrd%>">
	<input type="hidden" name="straar" value="<%=strAar%>">
	<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="70" alt="" border="0"></td>
	</tr>
	</table>
	<%
	else
	
			'** Finder kunde ****
			strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE kid = " & id
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			 strKundenavn = oRec("Kkundenavn")
			end if
			oRec.close
			
			if func = "opret" then%>
			<tr>
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="90" alt="" border="0"></td>
			<td colspan="3" style="padding-top:5;">
			<b>A)</b> Opret enkelt aktion: (1 Aktion) <input type="radio" checked name="FM_opf" id="FM_opf1" value="0" onclick="hideopf();">
			<br><b>B) </b>Tilføj <b>Opfølgning</b> til denne aktion?: (2 Aktion) <input type="radio" name="FM_opf" id="FM_opf2" value="1" onclick="showopf();">
			<!--<br><b>C)</b> 1 Aktion løber <b>over flere dage</b>:<input type="radio" name="FM_opf" id="FM_opf3" value="2" onclick="showendd();">-->
			<br><b>C)</b> Gentag 1 Aktion i et <b>bestemt interval</b>:<input type="radio" name="FM_opf" id="FM_opf4" value="3" onclick="showreocc();">
			</td>
			<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="90" alt="" border="0"></td>
			</tr>
			<%end if%>
			<input type="hidden" name="id" value="<%=id%>">
			<tr bgcolor="#EFF3FF">
				<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
				<td colspan=3 valign="bottom"><img src="../ill/tabel_top.gif" width="346" height="1" alt="" border="0"></td>
				<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
			</tr>
	</table>
	<br>
	
	<!-- dato /slut dato -->
	<div id=dato2 name=dato2 style="position:absolute; top:190; left:10; visibility:hidden; z-index:20000;">
	Dato:</div>
	<div id=sd name=sd style="position:absolute; top:230; left:10; width:100; height:19; visibility:hidden; z-index:20000;">
	Slut dato:</div>
	
	<%if func = "red" then
	vizi = "hidden" 
	else
	vizi = "visible"
	end if%>
	
	<!-- 1 og 2 aktion knapper -->
	<div id=divak1 name=divak1 style="position:absolute; top:136; left:15; visibility:<%=vizi%>; z-index:800;">
	<table cellspacing="0" cellpadding="0" border="0">
	<tr bgcolor="5582D2">
			<td width="8" rowspan="2" valign=top><img src="../ill/tt_tabel_top_left.gif" width="8" height="15" alt="" border="0"></td>
			<td valign="top"><img src="../ill/tabel_top.gif" width="64" height="1" alt="" border="0"></td>
			<td align=right rowspan="2" valign=top><img src="../ill/tt_tabel_top_right.gif" width="8" height="15" alt="" border="0"></td>
	</tr>
	<tr bgcolor="5582D2">
			<td><a href="#" onClick="showakt1()" class='alt'>1. Aktion</a></td>
	</tr>
	</table>
	</div>
	<div id=divak2 name=divak2 style="position:absolute; top:136; left:100; visibility:hidden; z-index:800;">
	<table cellspacing="0" cellpadding="0" border="0">
	<tr bgcolor="5582D2">
			<td width="8" rowspan="2" valign=top><img src="../ill/tt_tabel_top_left.gif" width="8" height="15" alt="" border="0"></td>
			<td valign="top"><img src="../ill/tabel_top.gif" width="64" height="1" alt="" border="0"></td>
			<td align=right rowspan="2" valign=top><img src="../ill/tt_tabel_top_right.gif" width="8" height="15" alt="" border="0"></td>
	</tr>
	<tr bgcolor="5582D2">
			<td><a href="#" onClick="showakt2()" class='alt'>2. Aktion</a></td>
	</tr>
	</table>
	</div>
	<div id=divak3 name=divak3 style="position:absolute; top:136; left:100; visibility:hidden; z-index:200;">
	<table cellspacing="0" cellpadding="0" border="0">
	<tr bgcolor="5582D2">
			<td width="8" rowspan="2" valign=top><img src="../ill/tt_tabel_top_left.gif" width="8" height="15" alt="" border="0"></td>
			<td valign="top"><img src="../ill/tabel_top.gif" width="84" height="1" alt="" border="0"></td>
			<td align=right rowspan="2" valign=top><img src="../ill/tt_tabel_top_right.gif" width="8" height="15" alt="" border="0"></td>
	</tr>
	<tr bgcolor="5582D2">
			<td><a href="#" onClick="showreocc()" class='alt'>Gentag aktion</a></td>
	</tr>
	</table>
	</div>
	<!--slut -->
	
	<br>
	<div id=div0A name=div0A style="position:relative; visibility:visible; display:; z-index:20;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="360">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=3 valign="top"><img src="../ill/tabel_top.gif" width="<%=table_topgif%>" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
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
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="360">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=3 valign="top"><img src="../ill/tabel_top.gif" width="<%=table_topgif%>" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan="3" style="height:20; padding-left:10;" class=alt><b><%=strKundenavn%>&nbsp;2. aktion: (opfølgende aktion)</b>&nbsp;</td>
	</tr>
	</table>
	</div>
	
	<div id=div1A name=div1A style="position:relative; visibility:visible; display:; z-index:20;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="360">
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="40" alt="" border="0"></td>
		<td style="padding-left:10;"><br>Dato:&nbsp;</td>
		<td><br>
		<select name="FM_start_dag">
		<option value="<%=strDag%>"><%=strDag%></option> 
		<option value="1">1</option>
	   	<option value="2">2</option>
	   	<option value="3">3</option>
	   	<option value="4">4</option>
	   	<option value="5">5</option>
	   	<option value="6">6</option>
	   	<option value="7">7</option>
	   	<option value="8">8</option>
	   	<option value="9">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	   	<option value="22">22</option>
	   	<option value="23">23</option>
	   	<option value="24">24</option>
	   	<option value="25">25</option>
	   	<option value="26">26</option>
	   	<option value="27">27</option>
	   	<option value="28">28</option>
	   	<option value="29">29</option>
	   	<option value="30">30</option>
		<option value="31">31</option></select>
		
		<select name="FM_start_mrd">
		<option value="<%=strMrd%>"><%=strMrdNavn%></option>
		<option value="1">jan</option>
	   	<option value="2">feb</option>
	   	<option value="3">mar</option>
	   	<option value="4">apr</option>
	   	<option value="5">maj</option>
	   	<option value="6">jun</option>
	   	<option value="7">jul</option>
	   	<option value="8">aug</option>
	   	<option value="9">sep</option>
	   	<option value="10">okt</option>
	   	<option value="11">nov</option>
	   	<option value="12">dec</option></select>
		
		
		<select name="FM_start_aar">
		<option value="<%=strAar%>">
		<%if id <> 0 AND func <> "opret" then%>
		20<%=strAar%>
		<%else%>
		<%=strAar%>
		<%end if%></option>
		
		<%for x = -5 to 10
		useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
		<option value="<%=useY%>"><%=right(useY, 2)%></option>
		<%next %>
		
		</select>&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=1')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>
		</td>
		<td>&nbsp;</td>
			<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="40" alt="" border="0"></td>
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
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="360">
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="80" alt="" border="0"></td>
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
			
			
			'end if
			
			if func = "opret" then 'OR (func = "red" AND cint(thisserialnb) = 0)
			
				if func = "opret" then
				opfdato = strDag&"/"&strMrd&"/"&strAar
				useopfdato = dateadd("d", 1, opfdato)
				else
					if len(crmDato_slut) <> 0 then
					useopfdato = crmDato_slut
					else
					useopfdato = crmDato
					end if
				end if
				
				useopfdag = day(useopfdato)
				useopfmrd = month(useopfdato)
				useopfmrdNavn = left(monthname(useopfmrd), 3)
				useopfaar = right(year(useopfdato), 2) 
				
				%>
				<img src="../ill/blank.gif" width="50" height="1" alt="" border="0">
				<select name="FM_start_dag_2">
				<option value="<%=useopfdag%>"><%=useopfdag%></option> 
				<option value="1">1</option>
			   	<option value="2">2</option>
			   	<option value="3">3</option>
			   	<option value="4">4</option>
			   	<option value="5">5</option>
			   	<option value="6">6</option>
			   	<option value="7">7</option>
			   	<option value="8">8</option>
			   	<option value="9">9</option>
			   	<option value="10">10</option>
			   	<option value="11">11</option>
			   	<option value="12">12</option>
			   	<option value="13">13</option>
			   	<option value="14">14</option>
			   	<option value="15">15</option>
			   	<option value="16">16</option>
			   	<option value="17">17</option>
			   	<option value="18">18</option>
			   	<option value="19">19</option>
			   	<option value="20">20</option>
			   	<option value="21">21</option>
			   	<option value="22">22</option>
			   	<option value="23">23</option>
			   	<option value="24">24</option>
			   	<option value="25">25</option>
			   	<option value="26">26</option>
			   	<option value="27">27</option>
			   	<option value="28">28</option>
			   	<option value="29">29</option>
			   	<option value="30">30</option>
				<option value="31">31</option></select>
				
				<select name="FM_start_mrd_2">
				<option value="<%=useopfmrd%>"><%=useopfmrdNavn%></option>
				<option value="1">jan</option>
			   	<option value="2">feb</option>
			   	<option value="3">mar</option>
			   	<option value="4">apr</option>
			   	<option value="5">maj</option>
			   	<option value="6">jun</option>
			   	<option value="7">jul</option>
			   	<option value="8">aug</option>
			   	<option value="9">sep</option>
			   	<option value="10">okt</option>
			   	<option value="11">nov</option>
			   	<option value="12">dec</option></select>
				
				<select name="FM_start_aar_2">
				<option value="<%=useopfaar%>">20<%=useopfaar%></option>
				<%for x = -5 to 10
		useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
		<option value="<%=useY%>"><%=right(useY, 2)%></option>
		<%next %></select>&nbsp;&nbsp;
				<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=2')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>
			<%end if
		end if%>
		</td>
		<td>&nbsp;
		</td>
			<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="80" alt="" border="0"></td>
		</tr>
		</table>
		</div>
		
		
		
		
	
	<div id=div2A name=div2A style="position:relative; visibility:visible; display:; z-index:20;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="360">
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		<td style="padding-left:10;">Start tidspunkt:
		</td><td><select name="klok_timer">
		<option value="<%=klok_timer%>" SELECTED><%=klok_timer%></option>
		<option value="07">7</option>
	   	<option value="08">8</option>
	   	<option value="09">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	  	</select>&nbsp;:
		
		<select name="klok_min">
		<%if func <> "opret" then%>
		<option value="<%=klok_min%>" selected><%=klok_min%></option>
		<%end if%>
		<option value="00">00</option>
		<option value="15">15</option>
	   	<option value="30">30</option>
	   	<option value="45">45</option>
		</select>
		</td>
		<td>&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		</tr>
		
		
		<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		<td style="padding-left:10;">Slut tidspunkt:
		</td><td>
		<select name="klok_timer_slut">
		<option value="<%=klok_timer_slut%>" SELECTED><%=klok_timer_slut%></option>
		<option value="07">7</option>
	   	<option value="08">8</option>
	   	<option value="09">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	  	</select>&nbsp;:
		
		<select name="klok_min_slut">
		<%if func <> "opret" then%>
		<option value="<%=klok_min_slut%>" selected><%=klok_min_slut%></option>
		<%else%>
		<option value="30" SELECTED>30</option>
		<%end if%>
		<option value="00">00</option>
		<option value="15">15</option>
	   	<option value="30">30</option>
	   	<option value="45">45</option>
		</select>
		</td>
		<td>&nbsp;
		</td>
			<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		</tr>
		
		<tr>
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
			<td style="padding-left:10;">Emne:</td>
			<td><select name="FM_emne">
			<%strSQL = "SELECT id, navn FROM crmemne ORDER BY navn"
			oRec.open strSQL, oConn, 3
			While not oRec.EOF %>
				<%if intEmne = oRec("id") then
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
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		</tr>
		<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		<td style="padding-left:10;">Overskrift:</td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" size="25" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		</td>
		<td>&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="85" alt="" border="0"></td>
		<td style="padding-left:10;">Kontakt person:
		<%if func = "red" then%>
		<!--<br>
		<font size=1 color=#999999>Hvis den rigtige person<br> ikke er forvalgt skyldes <br>det en struktur ændring<br> i TimeOut d. 25/1-2004.</font>-->
		<%end if%>
		</td>
		<td valign="top"><select name="FM_kpers">
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
		if cint(intKontaktpers) = oRec("id") then
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
	<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="85" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="165" alt="" border="0"></td>
		<td valign=top colspan="3" style="padding-left:10;">Beskrivelse/ Log:<br>
		<textarea cols="<%=trearsize%>" rows="8" name="FM_besk" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><%=strBesk%></textarea>
		&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="165" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		<td style="padding-left:10;">Status:</td>
		<td><select name="FM_status">
		<%strSQL = "SELECT id, navn FROM crmstatus ORDER BY navn"
		oRec.open strSQL, oConn, 3
		While not oRec.EOF %>
		<%if intStatus = oRec("id") then
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
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		<td style="padding-left:10;">Kontaktform:</td>
		<td><select name="FM_kform">
			<%strSQL = "SELECT id, navn FROM crmkontaktform ORDER BY navn"
			oRec.open strSQL, oConn, 3
			While not oRec.EOF %>
			<%if intKontaktform = oRec("id") then
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
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td colspan="3" style="padding-left:10;"><br>Tilføj medarbejdere:</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
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
				if oRec("mnavn") = session("user") then
				strcheckmedarb = "CHECKED"
				else
				strcheckmedarb = " "
				end if
			end if
		
		%> 
		<tr>
			<td colspan="2" style="padding-left:10; border-left: 1px solid #003399;"><%=oRec("mnavn")%></td><td colspan="3" style="padding-left:10; border-right: 1px solid #003399;"><input type="checkbox" name="FM_medarbrel" value="<%=oRec("mid")%>" <%=strcheckmedarb%> style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
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
		
	
	<div id=div2B name=div2B style="position:relative; visibility:hidden; display:none; z-index:20;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="360">
	<!--#include file="inc/crmhistorik_opf_inc.asp"-->
	</table>
	</div>
	
	
	<div id=div1C name=div1C style="position:relative; visibility:hidden; display:none; z-index:20;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="360">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=3 valign="top"><img src="../ill/tabel_top.gif" width="<%=table_topgif%>" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan="3" style="height:20; padding-left:10;" class=alt><b>Gentag aktion:</b>&nbsp;</td>
	</tr>
	<tr>
		<td valign="top" rowspan="3"><img src="../ill/tabel_top.gif" width="1" height="360" alt="" border="0"></td>
		<td style="padding-left:10;" valign=top colspan="3"><br>
		<b>Hvor ofte skal aktionen gentages:</b>
		</td>
		<td valign="top" align="right" rowspan="3"><img src="../ill/tabel_top.gif" width="1" height="360" alt="" border="0"></td>
	</tr>
	<tr>
		<td style="padding-left:10;" valign=top>Vælg <br><input type="radio" name="FM_dageldato" value="dag" checked><b>Ugedage</b><br>
		<input type="checkbox" name="FM_gentag_ugedag" value="1">&nbsp;Søndag<br>
		<input type="checkbox" name="FM_gentag_ugedag" value="2">&nbsp;Mandag<br>
		<input type="checkbox" name="FM_gentag_ugedag" value="3">&nbsp;Tirsdag<br>
		<input type="checkbox" name="FM_gentag_ugedag" value="4">&nbsp;Onsdag<br>
		<input type="checkbox" name="FM_gentag_ugedag" value="5">&nbsp;Torsdag<br>
		<input type="checkbox" name="FM_gentag_ugedag" value="6">&nbsp;Fredag<br>
		<input type="checkbox" name="FM_gentag_ugedag" value="7">&nbsp;Lørdag<br>
		<br><br>
		
		</td>
		<td valign=top width="190">eller <br><input type="radio" name="FM_dageldato" value="dato"><b>Datoer</b><br>
		<input type="checkbox" name="FM_gentag_dato" value="1">&nbsp;1
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="2">&nbsp;2
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="3">&nbsp;3<br>
		<input type="checkbox" name="FM_gentag_dato" value="4">&nbsp;4
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="5">&nbsp;5
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="6">&nbsp;6<br>
		<input type="checkbox" name="FM_gentag_dato" value="7">&nbsp;7
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="8">&nbsp;8
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="9">&nbsp;9<br>
		<input type="checkbox" name="FM_gentag_dato" value="10">&nbsp;10
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="11">&nbsp;11
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="12">&nbsp;12<br>
		<input type="checkbox" name="FM_gentag_dato" value="13">&nbsp;13
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="14">&nbsp;14
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="15">&nbsp;15<br>
		
		<input type="checkbox" name="FM_gentag_dato" value="16">&nbsp;16
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="17">&nbsp;17
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="18">&nbsp;18<br>
		
		<input type="checkbox" name="FM_gentag_dato" value="19">&nbsp;19
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="20">&nbsp;20
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="21">&nbsp;21<br>
		
		<input type="checkbox" name="FM_gentag_dato" value="22">&nbsp;22
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="23">&nbsp;23
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="24">&nbsp;24<br>
		
		<input type="checkbox" name="FM_gentag_dato" value="25">&nbsp;25
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="26">&nbsp;26
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="27">&nbsp;27<br>
		
		<input type="checkbox" name="FM_gentag_dato" value="28">&nbsp;28
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="29">&nbsp;29
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="30">&nbsp;30<br>
		
		<input type="checkbox" name="FM_gentag_dato" value="31">&nbsp;31<br></td>
		<td valign=top>og <br><img src="ill/blank.gif" width="1" height="4" alt="" border="0"><br><b>Måneder:</b><br>
		<input type="checkbox" name="FM_gentag_md" value="1">&nbsp;Jan<br>
		<input type="checkbox" name="FM_gentag_md" value="2">&nbsp;Feb<br>
		<input type="checkbox" name="FM_gentag_md" value="3">&nbsp;Mar<br>
		<input type="checkbox" name="FM_gentag_md" value="4">&nbsp;Apr<br>
		<input type="checkbox" name="FM_gentag_md" value="5">&nbsp;Maj<br>
		<input type="checkbox" name="FM_gentag_md" value="6">&nbsp;Jun<br>
		<input type="checkbox" name="FM_gentag_md" value="7">&nbsp;Jul<br>
		<input type="checkbox" name="FM_gentag_md" value="8">&nbsp;Aug<br>
		<input type="checkbox" name="FM_gentag_md" value="9">&nbsp;Sep<br>
		<input type="checkbox" name="FM_gentag_md" value="10">&nbsp;Okt<br>
		<input type="checkbox" name="FM_gentag_md" value="11">&nbsp;Nov<br>
		<input type="checkbox" name="FM_gentag_md" value="12">&nbsp;Dec<br></td>
	</tr>
	<tr>
		<td colspan="3" align=center><br><b>Slut gentagelser den:</b><br>
		<select name="FM_gentag_dag">
		<option value="<%=strDag%>"><%=strDag%></option> 
		<option value="01">1</option>
	   	<option value="02">2</option>
	   	<option value="03">3</option>
	   	<option value="04">4</option>
	   	<option value="05">5</option>
	   	<option value="06">6</option>
	   	<option value="07">7</option>
	   	<option value="08">8</option>
	   	<option value="09">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	   	<option value="22">22</option>
	   	<option value="23">23</option>
	   	<option value="24">24</option>
	   	<option value="25">25</option>
	   	<option value="26">26</option>
	   	<option value="27">27</option>
	   	<option value="28">28</option>
	   	<option value="29">29</option>
	   	<option value="30">30</option>
		<option value="31">31</option></select>
		
		<select name="FM_gentag_mrd">
		<option value="<%=strMrd%>"><%=strMrdNavn%></option>
		<option value="1">jan</option>
	   	<option value="2">feb</option>
	   	<option value="3">mar</option>
	   	<option value="4">apr</option>
	   	<option value="5">maj</option>
	   	<option value="6">jun</option>
	   	<option value="7">jul</option>
	   	<option value="8">aug</option>
	   	<option value="9">sep</option>
	   	<option value="10">okt</option>
	   	<option value="11">nov</option>
	   	<option value="12">dec</option></select>
		
		
		<select name="FM_gentag_aar">
		<option value="<%=strAar%>">
		<%if id <> 0 AND func <> "opret" then%>
		20<%=strAar%>
		<%else%>
		<%=strAar%>
		<%end if%></option>
		<option value="2002">2002</option>
		<option value="2003">2003</option>
	   	<option value="2004">2004</option>
	   	<option value="2005">2005</option>
		<option value="2006">2006</option>
		<option value="2007">2007</option></select>&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=3')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>
		</td>
	</tr>
	</table>
	</div>
	
	
	
	<%end if 'vælg kunde%>
	<table cellspacing="0" cellpadding="0" border="0" width="360" bgcolor="#EFF3FF">
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan=3 valign="bottom"><img src="../ill/tabel_top.gif" width="<%=table_topgif%>" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#d6dff5">
		<td colspan="5"><br><br><img src="ill/blank.gif" width="120" height="1" alt="" border="0"><input type="image" name="subm" src="../ill/<%=varSubVal%>.gif"></td>
	</tr>
	</form>
	</table>
	<br>
	</div>
	<%case else
	
	
	emner = request("emner")
	status = request("status")
	
	if len(request("medarb")) <> 0 then
	medarb = request("medarb")
	else
	medarb = session("mid")
	end if
	
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
		
		kid = 0
		
		if len(trim(request("id"))) <> 0 then
		kid = request("id")
		else
		    if request.cookies("crm")("kid") <> "" then
		    kid = request.cookies("crm")("kid")
		    else
		    kid = 0
		    end if
		end if
		
		Response.Cookies("crm")("kid") = kid
		
		ketypeKri = " ketype <> 'k' "
		
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE "& ketypeKri &" ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(kid) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn") & " ("&oRec("kkundenr") &") " %></option>
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
            response.cookies("crm")("hotnot") = hotnot
            
            else
                
                if request.cookies("crm")("hotnot") <> "" then
                hotnot = request.cookies("crm")("hotnot") 
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
		strSQL = "SELECT mnavn, mid FROM medarbejdere WHERE mansat <> '2' AND mansat <> '3' ORDER BY mnavn"
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		
			if cint(medarb) = cint(oRec("mid")) then
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
</select></td>
</tr>


<%end if



     if len(trim(request("emner"))) <> 0 then
		emneid = request("emner")
		else
		    if request.cookies("crm")("emneid") <> "" then
		    emneid = request.cookies("crm")("emneid")
		    else
		    emneid = 0
		    end if
		end if
		
		Response.Cookies("crm")("emneid") = emneid

%>



	<tr>
	<td><b>Emne:</b>&nbsp;</td><td>
	<select name="emner" size="1" style="font-size : 11px; width:200 px;" onchange="submit()";>
	<option value="0">Alle</option>
	<%
			strSQL = "SELECT navn, id FROM crmemne ORDER BY navn"
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
			
			if cint(emneid) = cint(oRec("id")) then
			isSelected = "SELECTED"
			else
			isSelected = ""
			end if
			%>
			<option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("navn")%></option>
			<%
			oRec.movenext
			wend
			oRec.close
			%>
	</select></td>
	
	</tr>
	
	<%
	
	
	 if len(trim(request("status"))) <> 0 then
		statusid = request("status")
		else
		    if request.cookies("crm")("statusid") <> "" then
		    statusid = request.cookies("crm")("statusid")
		    else
		    statusid = 0
		    end if
		end if
		
		Response.Cookies("crm")("statusid") = statusid
	
	%>


	<tr>
	<td><b>Status:</b> &nbsp;</td><td>
	<select name="status" size="1" style="font-size : 11px; width:200 px;" onchange="submit()";>
	<option value="0">Alle</option>
	<%
			strSQL = "SELECT navn, id FROM crmstatus ORDER BY navn"
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
			
			if cint(statusid) = cint(oRec("id")) then
			isSelected = "SELECTED"
			else
			isSelected = ""
			end if
			%>
			<option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("navn")%></option>
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
	
	response.cookies("crm").expires = date + 40
	
	
	
	oleftpx = 0
	otoppx = 20
	owdtpx = 130
	url = "javascript:NewWin_popupaktion(""crmhistorik.asp?menu=crm&func=opret&id="&id&"&ketype=e&showinwin=j&rdir=crmhist"")"
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
	<form method="post" action="crmhistorik.asp?func=opdaterliste&medarb=<%=medarb%>">
	
	 <%if print <> "j" then %>
	<tr bgcolor="#ffffff">
	<td colspan=10 align=right style="padding:20px 10px 10px 0px;"><input type="submit" value="Opdater liste"></td>
    </tr>
	<%end if %>
	
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
		<td colspan=2 style="width:300px;">Emne&nbsp;(¤ = serie)<br />
		<b>Navn</b><br />
		Kommentar</td>
		
		<td style="width:170px;"><b>Status</b></td>
		<td><b>Kontaktform</b></td>
		<td><b>Oprettet af</b></td>
		<td>&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/blank.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	select case level 
	case "1", "2", "6"
			if cint(medarb) = 0 OR len(medarb) = 0 then
			usemedarbKri = "aktionsid = ch.id AND medarbid <> 0"
			else
			usemedarbKri = "aktionsid = ch.id AND medarbid = '" & medarb & "'"
			end if
	case else
	usemedarbKri = "aktionsid = ch.id AND medarbid = '" & session("Mid") & "'"
	end select
	
	if kid = 0 OR len(kid) = 0 then
	useidKri = " kundeid <> 0"
	else
	useidKri = " kundeid = " & kid 
	end if
	
	if emneid = 0 OR len(emneid) = 0 then
	useemneKri = " kontaktemne <> -1"
	else
	useemneKri = " kontaktemne = " & emner 
	end if
	
	if statusid = 0 OR len(statusid) = 0 then
	usestatKri = " AND status <> -1 "
	else
	usestatKri = " AND status = " & statusid 
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
	
				strSQL2 = "SELECT id, navn, email, dirtlf, mobiltlf FROM kontaktpers WHERE id = "& kpersId
				
			    
			    'Response.Write strSQL2
	            'Response.flush
	            
	            oRec2.open strSQL2, oConn, 3
			    
				if not oRec2.EOF then
					if len(oRec2("navn")) > 25 then
					Response.write left(oRec2("navn"), 25) &".. "
					else
					Response.write oRec2("navn") 
					end if
					
					if len(trim(oRec2("email"))) <> 0 then
					Response.Write "<br><a href='mailto:"&oRec2("email")&"' class=kal_g>" & oRec2("email") & "</a>" 
					end if
					
					if len(trim(oRec2("dirtlf"))) <> 0 then
				    Response.Write "<br>Tel.(dir): " & oRec2("dirtlf") 
				    end if
				    
				    if len(trim(oRec2("mobiltlf"))) <> 0 then 
				    Response.Write "<br>Mobil: "& oRec2("mobiltlf") 
				    end if
					
				else
				Response.write ""
				end if
				
				
				
				oRec2.close
	
	end function
	
	
	
	
	
	strSQL = "SELECT DISTINCT(ch.id), ch.status, kid, kkundenavn, ch.editor, crmdato, kontaktemne, ch.navn, crmemne.navn AS enmenavn, "_
	&" crmstatus.navn AS statusnavn, crmkontaktform.navn AS kontaktform, "_
	&" ikon, ch.kontaktpers, ch.serialnb, telefon, ch.komm "_
	&" FROM kunder k "_
	&" LEFT JOIN crmhistorik ch ON (ch.crmdato BETWEEN '"& stDato &"' AND '"& slDato &"' AND ch.kundeid = k.kid)"_
	&" LEFT JOIN aktionsrelationer ON ("& usemedarbKri &")"_
	&" LEFT JOIN crmemne ON (crmemne.id = kontaktemne)"_
	&" LEFT JOIN crmstatus ON (crmstatus.id = ch.status)"_
	&" LEFT JOIN crmkontaktform ON (crmkontaktform.id = kontaktform)"_
	&" WHERE ("& useidKri &" "& hotnotKri &") "& usestatKri &""_
	&" AND "& useemneKri &" AND "& usemedarbKri &""_
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
		<td colspan="10" style="height:1px; border-top:1px #cccccc solid;"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bgcol %>">
        <input id="Hidden1" name="FM_id" value="<%=oRec("id") %>" type="hidden" />
		<td valign="top"><img src="../ill/blank.gif" width="1" height="42" alt="" border="0"></td>
		<td valign="top" height=20 style="padding:5px 4px 2px 0px;"><a href="kunder.asp?menu=crm&func=red&id=<%=oRec("Kid")%>"><%
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
			
				call kpers_her()
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
		
		<td valign="top" align=right style="padding:5px 15px 2px 0px;"><%=formatdatetime(oRec("crmdato"), 1)%></td>
		<td valign="top" colspan=2 style="width:400px; padding:5px 4px 2px 0px;">
		<a href="javascript:NewWin_popupaktion('crmhistorik.asp?selpkt=hist&menu=crm&func=red&id=<%=oRec("Kid")%>&crmaktion=<%=oRec("id")%>&showinwin=j&rdir=crmhist')" class='vmenu'><%=strEmne%></a>
		&nbsp;
		<%if oRec("serialnb") <> 0 then%>
		¤&nbsp;
		<%end if%>
		<br />
		<b><%=oRec("navn")%></b>
		<br />
		<%=oRec("komm") %>&nbsp;&nbsp;</td>
		<td valign="top" style="padding:5px 10px 2px 0px;">
		<select name="FM_status" size="1" style="font-size : 9px; width:100 px;">
	           
	            <%
			            strSQLs = "SELECT navn, id FROM crmstatus ORDER BY navn"
			            oRec2.open strSQLs, oConn, 3
			            while not oRec2.EOF
            			
			            if cint(oRec("status")) = cint(oRec2("id")) then
			            isSelected = "SELECTED"
			            else
			            isSelected = ""
			            end if
			            %>
			            <option value="<%=oRec2("id")%>" <%=isSelected%>><%=oRec2("navn")%></option>
			            <%
			            oRec2.movenext
			            wend
			            oRec2.close
			            %>
	     </select>
            &nbsp;
            
            
            
            <a href="javascript:NewWin_popupaktion('crmhistorik.asp?menu=crm&func=opret&id=<%=oRec("Kid")%>&ketype=e&showinwin=j')"><img src="../ill/add2.png" border=0 alt="Tilføj aktion" /></a>
            
		</td>
		<td valign="top" style="padding:5px 4px 2px 0px;"><%=oRec("kontaktform")%></td>
		<td valign="top" style="padding:5px 4px 2px 0px;"><i><%=left(oRec("editor"), 16)%></i></td>
		<td valign="top" style="padding:5px 4px 2px 0px;"><a href="crmhistorik.asp?selpkt=hist&menu=crm&func=slet&id=<%=oRec("Kid")%>&crmaktion=<%=oRec("id")%>&ketype=e&emner=<%=emner%>&status=<%=status%>&medarb=<%=medarb%>"><img src="../ill/slet_16.gif" alt="Slet aktion" border="0"></a></td>
		<td valign="top" align="right"><img src="../ill/blank.gif" width="1" height="42" alt="" border="0"></td>
	</tr>
	<%
	c = c + 1
	x = 0
	oRec.movenext
	wend
	%>
	
	 <%if print <> "j" then %>
	<tr bgcolor="#ffffff">
	<td colspan=10 align=right style="padding:20px 10px 10px 0px;"><input type="submit" value="Opdater liste"></td>
    </tr>
	<%end if %>
	
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/blank.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="8" valign="bottom"><img src="../ill/blank.gif" width="721" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/blank.gif" width="8" height="10" alt="" border="0"></td>
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
	
<%end select%>
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
