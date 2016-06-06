
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/job_func.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

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
	
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<% 
	
	
	slttxt = "<b>Slet Stam-aktivitets gruppe?</b><br />"_
	&"Du er ved at <b>slette</b> en Stam-aktivitets gruppe. Er dette korrekt?"
	slturl = "akt_gruppe.asp?menu=job&func=sletok&id="&id
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,210,90)
	
	
	case "sletok"
	'*** Her slettes en akt gruppe ***
    

    '*** Indsætter i delete historik ****' 
        
        '** gruppen **'
        strSQL = "SELECT id, navn FROM akt_gruppe WHERE id = "& id &"" 
        oRec.open strSQL, oConn, 3
	    if not oRec.EOF then
        
        call insertDelhist("aktstamgrp", id, 0, oRec("navn"), session("mid"), session("user"))
        
        end if
        oRec.close
        
        
        '** Aktivitteterne **'    
        strSQL = "SELECT id, navn FROM aktiviteter WHERE aktfavorit = "& id &" AND job = 0"

        oRec.open strSQL, oConn, 3
	    while not oRec.EOF
       
	    call insertDelhist("aktstam", id, 0, oRec("navn"), session("mid"), session("user"))

        oRec.movenext
        wend
        oRec.close


    oConn.execute("DELETE FROM aktiviteter WHERE aktfavorit = "& id &" AND job = 0")

	oConn.execute("DELETE FROM akt_gruppe WHERE id = "& id &"")


	Response.redirect "akt_gruppe.asp?menu=job&shokselector=1"
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
	
		<%
		errortype = 8
		call showError(errortype)
		
		else
		
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		forvalgt = request("FM_forvalgt")
        skabelonType = request("FM_type")
		
        '** Nulstiller forvalgt **'
        if cint(forvalgt) = 1 then
        oConn.execute("UPDATE akt_gruppe SET forvalgt = 0 WHERE id <> 0 AND skabelonType = "& skabelonType)
        end if
		

		if func = "dbopr" then
		oConn.execute("INSERT INTO akt_gruppe (navn, editor, dato, forvalgt, skabelontype) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& forvalgt &", "& skabelontype &")")
		else
		oConn.execute("UPDATE akt_gruppe SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', forvalgt = "& forvalgt &", skabelontype = "& skabelontype &" WHERE id = "&id&"")
		end if
		
		Response.redirect "akt_gruppe.asp?menu=job&shokselector=1"
		end if
	
	
	case "kopier"

            %>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
            <%

    call menu_2014() %>

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; visibility:visible;">
	<h4>Kopier Stam-aktivitetsgruppe</h4>
	
	<%
	
	
	
	tTop = 0
	tLeft = 0
	tWdth = 800
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	
	
	strSQL = "SELECT navn FROM akt_gruppe WHERE id =" & id
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	grpNavn = oRec("navn")
	
	oRec.movenext
	wend
	
	oRec.close
	%>
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr><td style="padding:20px;" valign=top>
	
	<table cellspacing="1" cellpadding="2" border="0" width="100%">
	<form method=post action="akt_gruppe.asp?func=kopierja&id=<%=id%>">
	<tr><td colspan=2><h4>Stam-aktivitetsgruppe:</h4></td></tr>
	<tr>
		<td align=right><b>Gruppenavn:</b></td><td><input type="text" name="FM_grpnavn" id="FM_grpnavn" style="width:400px;" value="Kopi af: <%=grpNavn%>"></td>
	</tr>
	<tr><td colspan=2><br><h4>Aktiviteterne:</h4></td></tr>
	<tr>
		<td align=right valign=top style="padding-top:6px;"><b>Aktivitetsnavn:</b></td>
		<td style="padding-top:2px;"><input type="checkbox" name="FM_udskift_navn" id="FM_udskift_navn" value="1">Ja, omdøb navn på aktiviteter.<br>
		Udskift denne del af aktivitets navnet:<br> 
		<input type="text" name="FM_aktnavn_for" id="FM_aktnavn_for" value=""><br>
		med dette: <br>
		<input type="text" name="FM_aktnavn_efter" id="FM_aktnavn_efter" value=""></td>
	</tr>
	<tr>
		<td align=right valign=top style="padding-top:6px;"><b>Fase:</b></td>
		<td style="padding-top:2px;"><input type="checkbox" name="FM_udskift_fase" id="Checkbox1" value="1">Ja, omdøb fase på aktiviteter.<br>
		Udskift denne del af fase navnet:<br> 
		<input type="text" name="FM_fase_for" id="Text1" value=""><br>
		med dette: <br>
		<input type="text" name="FM_fase_efter" id="Text2" value=""></td>
	</tr>
	<tr>
		<td align=right valign=top style="padding-top:20px;"><b>Faktor:</b></td>
		<td style="padding-top:16px;"><input type="checkbox" name="FM_udskift_faktor" id="FM_udskift_faktor" value="1">Ja, udskift faktor.<br>
		<input type="text" name="FM_faktor" id="FM_faktor" value="" size=2></td>
	</tr>
	<tr>
		<td align=right valign=top style="padding-top:20px;"><b>Forretningsområder:</b></td>
			<td style="padding-top:16px;"><input type="checkbox" name="FM_udskift_fomr" id="FM_udskift_fomr" value="1">Ja, udskift forretningsområder med:<br>
			<select name="FM_fomr" multible="multible" size="4" style="width:400px;">
		<option value="0">Ingen</option>
		<%
		strSQL = "SELECT id, navn FROM fomr ORDER BY navn"
		oRec.open strSQL, oConn, 3 
		while not oRec.EOF 
		%>
		<option value="<%=oRec("id")%>"><%=oRec("navn")%></option>
		<%
		oRec.movenext
		wend
		oRec.close 
		%>
		</select>
		</td>
	</tr>
	<tr>
		<td align=right valign=top style="padding-top:20px;"><b>Tidslås:</b></td>
		<td style="padding-top:16px;"><input type="checkbox" name="FM_udskift_tidslaas" id="FM_udskift_tidslaas" value="1">Ja, udskift tidslås.<br>
		<input type="checkbox" name="FM_tidslaas" id="FM_tidslaas" value="1">Tidslås aktiv. Der skal <b>kun</b> kunne registreres timer mellem:<br> 
		kl. <input type="text" name="FM_tidslaas_start" id="FM_tidslaas_start" size="3" value="07:00"> og kl.
		<input type="text" name="FM_tidslaas_slut" id="FM_tidslaas_slut" size="3" value="23:30"> (Format: tt:mm)
		
		<br />På følgende dage:<br />
                         Man 
                        <input id="FM_tidslaas_man" name="FM_tidslaas_man" value="1" CHECKED type="checkbox"/>&nbsp;&nbsp;
                         Tir
                        <input id="FM_tidslaas_tir" name="FM_tidslaas_tir" value="1" CHECKED type="checkbox"/>&nbsp;&nbsp;
                         Ons 
                        <input id="FM_tidslaas_ons" name="FM_tidslaas_ons" value="1" CHECKED type="checkbox"/>&nbsp;&nbsp;
                         Tor 
                        <input id="FM_tidslaas_tor" name="FM_tidslaas_tor" value="1" CHECKED type="checkbox"/>&nbsp;&nbsp;
                         Fre 
                        <input id="FM_tidslaas_fre" name="FM_tidslaas_fre" value="1" CHECKED type="checkbox"/>&nbsp;&nbsp;
                         Lør 
                        <input id="FM_tidslaas_lor" name="FM_tidslaas_lor" value="1" type="checkbox"/>&nbsp;&nbsp;
                         Søn 
                        <input id="FM_tidslaas_son" name="FM_tidslaas_son" value="1" type="checkbox"/>&nbsp;&nbsp;
     
		
		</td>
	</tr>
	<tr>
		<td colspan=2 align=center><br><br><input type="submit" value="Kopier gruppe"></td>
	</tr>
	</form>
	</table>
	
	<!-- tableDiv -->
	</td></tr></table>
	</div>
	<br /><br /><br /><br /><br /><br /><br />&nbsp;
	</div>
	
	<br /><br /><br /><br /><br /><br /><br />&nbsp;
	
	<%
	case "kopierja"
	
	'Response.write "så er vi her"
	
			
			'** Faktor **
			if len(request("FM_udskift_faktor")) <> 0 then
			udskiftFaktor = 1
			else
			udskiftFaktor = 0
			end if
			aktFaktor = replace(request("FM_faktor"), ",", ".")
				
			%>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(dblFaktor)
			if isInt > 0 then
			%>
			<!--#include file="../inc/regular/header_inc.asp"-->
			
			<%
			errortype = 51
			call showError(errortype)
			isInt = 0
			response.end
			end if 
				
					
				
				'*** tidslaas validering ***
				'** Tidslaas **
				udskiftTidslaas = request("FM_udskift_tidslaas")
				
				if len(request("FM_tidslaas")) <> 0 then
				tidslaas = 1
				else
				tidslaas = 0
				end if
				
				tidslaas_st = request("FM_tidslaas_start") & ":00"
				tidslaas_sl = request("FM_tidslaas_slut") & ":00"
				
				if len(trim(request("FM_tidslaas_man"))) <> 0 then
				tidslaas_man = 1
				else 
				tidslaas_man = 0
				end if
				
				if len(trim(request("FM_tidslaas_tir"))) <> 0 then
				tidslaas_tir = 1
				else 
				tidslaas_tir = 0
				end if
				
				if len(trim(request("FM_tidslaas_ons"))) <> 0 then
				tidslaas_ons = 1
				else 
				tidslaas_ons = 0
				end if
				
				if len(trim(request("FM_tidslaas_tor"))) <> 0 then
				tidslaas_tor = 1
				else 
				tidslaas_tor = 0
				end if
				
				if len(trim(request("FM_tidslaas_fre"))) <> 0 then
				tidslaas_fre = 1
				else 
				tidslaas_fre = 0
				end if
				
				if len(trim(request("FM_tidslaas_lor"))) <> 0 then
				tidslaas_lor = 1
				else 
				tidslaas_lor = 0
				end if
				
				if len(trim(request("FM_tidslaas_son"))) <> 0 then
				tidslaas_son = 1
				else 
				tidslaas_son = 0
				end if
					
				function SQLBless3(s3)
				dim tmp3
				tmp3 = s3
				tmp3 = replace(tmp3, ":", "")
				SQLBless3 = tmp3
				end function
				
				
				idagErrTjek = day(now)&"/"&month(now)&"/"&year(now)
				
				for t = 1 to 2
				
				select case t
				case 1
				tdsl = tidslaas_st
				case 2
				tdsl = tidslaas_sl
				end select
				
				'Response.write SQLBless3(trim(tSttid(y))) & "<br>"
				'Response.flush
				
				call erDetInt(SQLBless3(trim(tdsl)))
				if isInt > 0 then
					antalErr = 1
					errortype = 63
					useleftdiv = "t"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end
				end if	
				isInt = 0
				
				'*** Punktum  i angivelse ved registrering af klokkeslet
				if instr(trim(tdsl), ".") <> 0 OR instr(trim(tdsl), ",") <> 0 then
					antalErr = 1
					errortype = 66
					useleftdiv = "t"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end
				end if	
				
				if len(trim(tdsl)) <> 0 then
				
				'Response.write idagErrTjek &" "& tSttid(y)&":00" &" - "& isdate(idagErrTjek &" "& tSttid(y)&":00") &"<br>"
					if isdate(idagErrTjek &" "& tdsl) = false then
						antalErr = 1
						errortype = 64
						useleftdiv = "t"
						%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
						call showError(errortype)
						response.end
					end if
				end if
				
			
			next
				
				
			'*** Validering ok! *** 		
			
			
			'** Gruppe **
			intAktfavgp = id
			grpNavn = SQLBless(request("FM_grpnavn"))
			strEditor = session("user")
			strDato = day(now) &"-" & month(now) & "-" & year(now)
			
			'** Opretter gruppe ***
			strSQLopr = "INSERT INTO akt_gruppe (navn, editor, dato) VALUES ('"& grpNavn &"', '"& strEditor &"', '"& strDato &"')"
			oConn.execute strSQLopr
			
			
			'*** Henter gruppe ID ***
			strSQL = "SELECT id FROM akt_gruppe WHERE id <> 0 ORDER BY id DESC"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
				
				newaktgrpId = oRec("id")
			
			end if
			oRec.close   
			
			'** Navn **
			if len(request("FM_udskift_navn")) <> 0 then
			udskiftNavn = 1
			else
			udskiftNavn = 0
			end if
			aktNavn_for = SQLBless(request("FM_aktnavn_for")) 
			aktNavn_efter = SQLBless(request("FM_aktnavn_efter"))
			
			
			'** Fase **'
			if len(request("FM_udskift_fase")) <> 0 then
			udskiftfase = 1
			else
			udskiftfase = 0
			end if
			fase_for = SQLBless(request("FM_fase_for")) 
			fase_efter = SQLBless(request("FM_fase_efter"))
			
			function SQLBless4(s, glnavn, nytnavn)
			tmpthis = s
			tmpthis = replace(lcase(tmpthis), ""&lcase(glnavn)&"", ""&nytnavn&"")
			SQLBless4 = tmpthis
			end function
			
			
			'** Forretningsområde ***
			if len(request("FM_udskift_fomr")) <> 0 then
			udskiftFomr = 1
			else
			udskiftFomr = 0
			end if

			aktFomr = 0 'request("FM_fomr") 


                  '************************************'
                  '***** Forrretningsområder **********'
                  '************************************'

                  if cint(udskiftFomr) = 1 then

                    fomrArr = split(request("FM_fomr"), ",")

                    for_faktor = 0
                    for afor = 0 to UBOUND(fomrArr)
                    for_faktor = for_faktor + 1 
                    next

                    if for_faktor <> 0 then
                    for_faktor = for_faktor
                    else
                    for_faktor = 1
                    end if

                    for_faktor = formatnumber(100/for_faktor, 2)
                    for_faktor = replace(for_faktor, ",", ".")
                  
                  
                  end if

                  '**************************************'
			
			
			
				
				
							'**** Henter aktiviteter fra oprindelig gruppe ***
							strSQL2 = "select id, navn, fakturerbar, budgettimer, fomr, faktor, "_
							&" aktbudget, aktstatus, tidslaas, tidslaas_st, tidslaas_sl, antalstk, "_
							&" tidslaas_man, tidslaas_tir, tidslaas_ons, "_
	                        &" tidslaas_tor, tidslaas_fre, tidslaas_lor, tidslaas_son, fase, sortorder, bgr, aktbudgetsum, easyreg, brug_fasttp, brug_fastkp, fasttp, fasttp_val, fastkp, fastkp_val, avarenr, kostpristarif, aktkonto "_
							&" FROM aktiviteter WHERE aktFavorit = "& intAktfavgp 
								
								    oRec2.open strSQL2, oConn, 3
									while not oRec2.EOF
									
									aktNavn = replace(oRec2("navn"), "'", "''")
									
									if cint(udskiftNavn) = 1 then
									aktNavn = SQLBless4(trim(aktNavn),aktNavn_for,trim(aktNavn_efter))
									end if
									
									aktFakbar = oRec2("fakturerbar")
									
									'if cint(udskiftFomr) = 0 then
									'aktFomr = oRec2("fomr")
									'else
									'aktFomr = aktFomr
									'end if
									
									if cint(udskiftFaktor) = 0 then
									aktFaktor = replace(oRec2("faktor"), ",", ".")
									else
									aktFaktor = aktFaktor
									end if
									
									if len(aktFaktor) <> 0 then
									aktFaktor = aktFaktor
									else
									aktFaktor = 0
									end if
									
									aktBudget = replace(oRec2("aktbudget"), ",", ".")
									aktbudgetsum = replace(oRec2("aktbudgetsum"), ",", ".")
									
									aktstatus = oRec2("aktstatus")
									
									if len(trim(oRec2("fase"))) <> 0 then
									strFase = replace(oRec2("fase"), "'", "''")
									else
									strFase = ""
									end if
									
									if cint(udskiftfase) = 1 then
									strFase = SQLBless4(trim(strFase),fase_for,trim(fase_efter))
									end if
									
									bgr = oRec2("bgr")
									
									sortorder = oRec2("sortorder")
									
									
									if cint(tidslaas) = 0 then 
									
										tidslaas = oRec2("tidslaas")
										tidslaas_st = oRec2("tidslaas_st")
										tidslaas_sl = oRec2("tidslaas_sl")
										
										
										if len(tidslaas) <> 0 then
										tidslaas = tidslaas
										else
										tidslaas = 0
										end if
										
										
										if len(trim(tidslaas_st)) <> 0 AND len(trim(tidslaas_sl)) <> 0 then
										tidslaas_st = left(formatdatetime(tidslaas_st, 3), 7)
										tidslaas_sl = left(formatdatetime(tidslaas_sl, 3), 7)
										else
										tidslaas_st = "07:00:00"
										tidslaas_sl = "23:30:00"
										end if
										
										
				                        tidslaas_man = oRec2("tidslaas_man")
				                        tidslaas_tir = oRec2("tidslaas_tir")
				                        tidslaas_ons = oRec2("tidslaas_ons")
				                        tidslaas_tor = oRec2("tidslaas_tor")
				                        tidslaas_fre = oRec2("tidslaas_fre")
				                        tidslaas_lor = oRec2("tidslaas_lor")
				                        tidslaas_son = oRec2("tidslaas_son")


				                       
									
									end if
									
									antalstk = replace(oRec2("antalstk"), ",", ".")
									aktBudgettimer = replace(oRec2("budgettimer"), ",", ".")
									
									easyreg = oRec2("easyreg")

                                    brug_fasttp = oRec2("brug_fasttp")
                                    brug_fastkp = oRec2("brug_fastkp")
                                    fasttp = replace(oRec2("fasttp"), ",", ".")
                                    fasttp_val = oRec2("fasttp_val")
                                    fastkp = replace(oRec2("fastkp"), ",", ".")
                                    fastkp_val = oRec2("fastkp_val")
                         
                                    avarenr = oRec2("avarenr")

                                    kostpristarif = oRec2("kostpristarif")

                                   
                                    aktkonto = oRec2("aktkonto")
									
									
									strSQLins = "INSERT INTO aktiviteter "_
									&" (navn, dato, editor, job, fakturerbar, "_
									&" projektgruppe1, projektgruppe2, projektgruppe3, "_
									&" projektgruppe4, projektgruppe5, projektgruppe6, "_
									&" projektgruppe7, projektgruppe8, projektgruppe9, "_
									&" projektgruppe10, aktstartdato, aktslutdato, "_
									&" budgettimer, fomr, faktor, aktBudget, aktstatus, tidslaas, tidslaas_st, "_
									&" tidslaas_sl, antalstk, aktfavorit,"_
									&" tidslaas_man, tidslaas_tir, tidslaas_ons, "_
						            &" tidslaas_tor, tidslaas_fre, tidslaas_lor, tidslaas_son, sortorder, bgr, aktbudgetsum,"_
                                    &" easyreg, brug_fasttp, brug_fastkp, fasttp, fasttp_val, fastkp, fastkp_val, avarenr, kostpristarif, aktkonto"
						            
						            if len(trim(strFase)) <> 0 then
						            strSQLins = strSQLins & ", fase"
						            end if
						             
						             
						             
									 strSQLins = strSQLins &") VALUES "_
									&"('"& aktNavn &"', "_
									&"'"& strDato &"', "_ 
									&"'"& strEditor &"', "_
									&"0, "_ 
									&""& aktFakbar &", "_
									&"10,1,1,1,1,1,1,1,1,1, "_ 
									&"'2001/1/1', "_ 
									&"'2044/1/1', "_
									&""& aktBudgettimer & ", "& aktFomr &", "_
									&""& aktFaktor &", "& aktBudget &", "& aktstatus &", "_
									&""&tidslaas&", '"&tidslaas_st&"', '"&tidslaas_sl&"', "& antalstk &", "& newaktgrpId &", "_
									&" "& tidslaas_man &", "& tidslaas_tir &", "& tidslaas_ons &", "_
						            &" "& tidslaas_tor &", "& tidslaas_fre &", "& tidslaas_lor &", "& tidslaas_son & ", " & sortorder & ", "& bgr &", "_
                                    &" "& aktbudgetsum & ", "& easyreg &", "& brug_fasttp &","& brug_fastkp &","& fasttp &","& fasttp_val &","& fastkp &","& fastkp_val &","_
                                    &" '"& avarenr &"', '"& kostpristarif &"', "& aktkonto &""
						            
						            if len(trim(strFase)) <> 0 then
						            strSQLins = strSQLins & ", '"& strFase &"'"
						            end if
						            
						            strSQLins = strSQLins &")"
									
									
									'Response.write strSQLins
									'Response.flush
									oConn.execute(strSQLins)
									
									
									'*** Henter det netop oprettede akt-id ***
									strSQLid = "SELECT id FROM aktiviteter ORDER BY id DESC"
									oRec3.open strSQLid, oConn, 3
									if not oRec3.EOF then
									useAktid = oRec3("id")
									end if
									oRec3.close
									
									if len(useAktid) <> 0 then
									useAktid = useAktid
									else
									useAktid = 0
									end if
									
									
									'*** Timepriser ***
									strSQLtp1 = "SELECT jobid, aktid, medarbid, timeprisalt, 6timepris FROM timepriser WHERE aktid = " & oRec2("id")
									oRec3.open strSQLtp1, oConn, 3
									while not oRec3.EOF
									 	
										strjnr = oRec3("jobid")
										medarbid = oRec3("medarbid")
										timeprisalt = replace(oRec3("timeprisalt"), ",",".")
										this6timepris = replace(oRec3("6timepris"), ",",".")
										
										strSQLtp2 = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris) VALUES ("& strjnr &", "& useAktid &", "& medarbid &", "&timeprisalt&", "& this6timepris &")"
										oConn.execute (strSQLtp2)
										
										
									oRec3.movenext
									wend
									oRec3.close
									
									
                                                    
                                                    '********************************'
                                                    '***** Forretningsområder ******'
                                                    '********************************'
                    
                    
                                                    '**** Hvis udskift = 1 (ja), ellers kopier ***'
                                                    if cint(udskiftFomr) = 1 then

                                                            for afor = 0 to UBOUND(fomrArr)

                                                                    'Response.Write "her2" & afor & "<br>"
                                                                    'Response.Flush

                                                                    if fomrArr(afor) <> 0 then

                                                                    strSQLfomri = "INSERT INTO fomr_rel "_
                                                                    &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                                                                    &" VALUES ("& fomrArr(afor) &", "& strjnr &", "& useAktid &", "& for_faktor &")"

                                                                    oConn.execute(strSQLfomri)

                                                                   end if


                                                            next

                                                    else 'kopier
                                                            
                                                            strSQLafor = "SELECT for_fomr, for_jobid, for_aktid, for_faktor FROM fomr_rel WHERE for_aktid = "& oRec2("id")

                                                            oRec3.open strSQLafor, oConn, 3
                                                            while not oRec3.EOF

                                                                    strSQLfomri = "INSERT INTO fomr_rel "_
                                                                    &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                                                                    &" VALUES ("& oRec3("for_fomr") &", "& oRec3("for_fomr") &", "& useAktid &", "& oRec3("for_faktor") &")"

                                                                    oConn.execute(strSQLfomri)

                                                            oRec3.movenext
                                                            wend
                                                            oRec3.close


                                                    end if

                                                    '********************************' 

									
							
							oRec2.movenext
							wend
							oRec2.close
					
						
		
	Response.redirect "akt_gruppe.asp"
	
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	intForvalgt = 0
    skabelontype = 0
	
	else
	strSQL = "SELECT navn, forvalgt, editor, dato, skabelontype FROM akt_gruppe WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	intForvalgt = oRec("forvalgt")
    skabelontype = oRec("skabelontype")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if


    if cint(skabelontype) = 0 then
    skabelontypeSEL0 = "SELECTED"
    skabelontypeSEL1 = ""
    else
    skabelontypeSEL1 = "SELECTED"
    skabelontypeSEL0 = ""
    end if

	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
  

    <%call menu_2014() %>

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:92px; visibility:visible;">
	<h4>Stam-aktivitetsgrupper

        <%if dbfunc = "dbred" then%>
        <br /><span style="font-size:9px; font-weight:lighter;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></span>
        <%end if%>

	</h4>
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr><form action="akt_gruppe.asp?menu=crm&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
    	<td valign="top" colspan="2">&nbsp;</td>
	</tr>
	
	<tr>
		<td><b>Stam-aktivitetsgruppe:</b> (Navn)</td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" style="width:300px;"></td>
	</tr>
	<tr>
		<td valign=top><br /><b>Forvalgt:</b><br />
		<span style="color:#999999">Angiv om denne gruppe skal stå forvalgt ved joboprettelse.<br /> Der kan kun være 1 forvalgt stam-aktivitetsgruppe.</span></td>
		<td valign=top><br />
        <%if cint(intForvalgt) <> 0 then
        fvCHK1 = "CHECKED"
        fvCHK0 = ""
        else
        fvCHK0 = "CHECKED"
        fvCHK1 = ""
        end if %>

            <input id="Radio1" type="radio" name="FM_forvalgt" value="0" <%=fvCHK0 %> /> Nej
            <br />
            <input id="Radio2" type="radio" name="FM_forvalgt" value="1" <%=fvCHK1 %> /> Ja

       
		</td>
	</tr>
        <tr><td><br /><b>Type:</b><br />
            <span style="color:#999999">Angiv om denne gruppe er en Projekt- eller budget -skabelon</span>
            </td><td>

            <select name="FM_type">
                <option value="0" <%=skabelontypeSEL0 %>>Projekt</option>
                <option value="1" <%=skabelontypeSEL1 %>>Budget</option>
            </select>


            </td></tr>
	<tr>
		<td colspan="2" align="right"><br><br>
		<input type="submit" value=" Opdater >> " /></td>
	</tr>
	</form>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%case else%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
 

    <%call menu_2014() %>
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:82px; visibility:visible;">
	
	<%
    'oimg = "aktstam_48.png"
	
    
   
    
    
     tTop = 20
	tLeft = 0
	tWdth = 804
	
	
	call tableDiv(tTop,tLeft,tWdth)


        oleft = 0
	otop = 0
	owdt = 600
	oskrift = "Stam-aktivitetsgrupper (skabeloner)"
	
	call sideoverskrift_2014(oleft, otop, owdt, oskrift)
    
    %><br /><br />

       <div style="position:absolute; left:597px; top:75px;"><% 
                nWdt = 210
                nTxt = "Opret ny stam-aktivitetsgruppe"
                nLnk = "akt_gruppe.asp?menu=job&func=opret"
                nTgt = ""
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) %> 
	        
                </div> 
	
	
	<table cellspacing="0" cellpadding="2" border="0" width="100%" bgcolor="#EFF3FF">
 	
	<tr bgcolor="#5582D2">
	    <td style="height:30px;">&nbsp;</td>
		<td class='alt'><b>Id</b></td>
		<td class='alt'><b>Stam-aktivitets gruppe</b></td>
		<td class='alt'><b>Forvalgt</b><br />
		(ved joboprettelse)</td>
		<td class='alt'><b>Se gruppe </b>(antal akt.)</td>
		<td class='alt'><b>Slet gruppe?</b></td>
		<td class='alt'><b>Kopier gruppe?</b></td>
		<td>&nbsp;</td>
	</tr>
	<%
	strSQL = "SELECT id, navn, forvalgt, skabelontype FROM akt_gruppe WHERE id <> 2 ORDER BY navn" 'gl. kørsels gruppe
	
    c = 0
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
    select case right(c, 1)
    case 0,2,4,6,8
    bgt = "#EFF3FF"
    case else
    bgt = "#FFFFFF"
    end select

	%>

	<tr bgcolor="<%=bgt %>">
		<td valign="top" style="border-bottom:1px #CCCCCC solid;"><img src="../ill/blank.gif" width="1" height="30" alt="" border="0"></td>
		<td style="border-bottom:1px #CCCCCC solid;" class=lille><%=oRec("id")%></td>
					
					<%
					'** Henter aktiviteter i den aktuelle gruppe ****
					strSQL2 = "SELECT count(id) AS antal FROM aktiviteter WHERE aktfavorit = "&oRec("id")&""
					oRec2.open strSQL2, oConn, 3
					if not oRec2.EOF then
					intAntal = oRec2("antal")
					end if
					oRec2.close 
					%>
		
		<%if oRec("id") = 2 then%>
		<td height="20" style="border-bottom:1px #CCCCCC solid;"><%=oRec("navn")%></td>
		<%else%>
		<td height="20" style="border-bottom:1px #CCCCCC solid;"><a href="akt_gruppe.asp?menu=job&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
		<%end if%>
		<td style="border-bottom:1px #CCCCCC solid;"><%=oRec("forvalgt") %>

            <%if cint(oRec("skabelontype")) = 1 then 'budget
                %>
            - Budgetskabelon
            <%
            end if %>

		</td>
		
		<td style="border-bottom:1px #CCCCCC solid;"><a href='aktiv.asp?menu=job&func=favorit&id=<%=oRec("id")%>&stamakgrpnavn=<%=oRec("navn")%>' class='vmenuglobal'>Se / Rediger Stam-aktiviteter i grp.&nbsp;</a>(<%=intAntal%>)</td>
		<%if oRec("id") <= 2 then%>
		<td style="border-bottom:1px #CCCCCC solid;">&nbsp;</td>
		<%else%>
			<td style="border-bottom:1px #CCCCCC solid;" align=center><a href="akt_gruppe.asp?menu=job&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet_16.gif" alt="Slet Stamaktivitetes-gruppe" border="0"></a></td>
			
			
		<%end if%>
		
		<td style="border-bottom:1px #CCCCCC solid;"><a href="akt_gruppe.asp?func=kopier&id=<%=oRec("id")%>" class=rmenu><img src="../ill/aktstam_kopier_24.gif" alt="Kopier gruppe" border="0"></a></td>
		<td style="border-bottom:1px #CCCCCC solid;" valign="top" align="right"><img src="../ill/blank.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	x = 0
    c = c + 1
	oRec.movenext
	wend
	%>	
	
	</table>
	</div>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	
	&nbsp;
	
	
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
