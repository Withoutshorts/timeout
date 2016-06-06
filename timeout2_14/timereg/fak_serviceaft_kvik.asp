<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"--> 
<!--#include file="inc/functions_inc.asp"-->  

<%



if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
		if request("FM_fak_alle") <> "1" OR len(trim(request("FM_aft"))) = 0 then
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
		<%
		errortype = 52
		call showError(errortype)
		else
		%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="inc/isint_func.asp"--> 
	<%
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	func = request("func")
	
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	
	'*** Opbygger aftale kriterie ***
	if len(request("FM_aft")) <> 0 then
	aftaler = split(request("FM_aft"), ",")
	
	fakDato = Request("FM_fak_aar") & "/" & Request("FM_fak_mrd") & "/" & Request("FM_fak_dag") '& time 
	
	strEditor = session("user")
	strDato = session("dato")
	strjobid = 0
	
	intFakbetalt = request("FM_betalt")
	if intFakbetalt <> 1 then
	intFakbetalt = 0
	else
	intFakbetalt = intFakbetalt 
	end if
	
	'aftid = Request("aftid")
	jobid = strjobid
	
	intEnhedsang = request("FM_enheds_ang")
	dtb_dato = year(now) & "/"& month(now) & "/" & day(now)
	
	
	useDebKre = request("FM_debkre")
	'Response.write useDebKre 
	
	varKonto = request("FM_kundekonto")
	varModkonto = request("FM_modkonto")
	
	
	
				
	'*** Opretter Fakturaer ***
				
				i = 0
	
				for i = 0 to Ubound(aftaler)
					
					'Response.write aftaler(i) &"<br>"
					
					
					'*** Henter Kunde og Aftale oplysninger ***
					strSQL2 = "SELECT s.kundeid, kid, s.enheder, s.pris, s.navn, s.besk, s.varenr FROM serviceaft s "_
					&" LEFT JOIN kunder k ON (kid = s.kundeid) WHERE s.id = " & aftaler(i)
					oRec2.open strSQL2, oConn, 3 
					if not oRec2.EOF then
					
					intfakadr = oRec2("kid")
					intEnheder = replace(oRec2("enheder"), ",",".")
					intPris = replace(oRec2("pris"), ",",".")
					strBesk = left(oRec2("navn"), 20) 
					strVarenr = oRec2("varenr")
							
							
							'*** Beregner Moms, Konto og Modkonto *******
				
							call beregnmoms(useDebKre, intPris, varKonto, varModkonto)
		
					
					
					end if
					oRec2.close 
					
					
						'*** Hent faknr ***
						lastFaknr = 0 	
						
						strSQL = "SELECT Fid, faknr FROM fakturaer WHERE faktype = 0 ORDER BY fid DESC" 
						oRec.open strSQL, oConn, 3
						
						while not oRec.EOF 
						
						call erDetInt(oRec("faknr"))
							if isInt = 0 then
								if cint(oRec("faknr")) >= lastFaknr then
								lastFaknr = cint(oRec("faknr")) + 1
								else
								lastFaknr = lastFaknr
								end if
							end if
						isInt = 0
						oRec.movenext
						wend
						oRec.close
					
					 
					lastFaknr = lastFaknr '+ i
					
					
					'** Opretter faktrura ***
					strSQL = "INSERT INTO fakturaer ("_
					&" dato, editor, tidspunkt, faknr, fakdato, jobid, timer, beloeb, "_
					&" kommentar, betalt, b_dato,  "_
					&" fakadr, att, faktype, "_
					&" konto, modkonto, "_
					&" parentfak, moms, "_
					&" enhedsang, varenr, aftaleid "_
					&") VALUES ( "_
					&" '"& strDato &"',"_
					&" '"& strEditor &"', '23:59:59', '"& lastFaknr &"', "_
					&" '"& fakDato &"', "& jobid &", "& intEnheder &", "& intPris &", '"& strBesk &"', "_
					&" "& intFakbetalt &", '"& dtb_dato &"', "& intfakadr &", 0, 0, "_
					&" "& varKonto &", "& varModkonto &", 0, "_
					&" "& showMoms &", "& intEnhedsang &", '"& strVarenr &"', "& aftaler(i) &")"
					
					'Response.write strSQL & "<br><br>"
					'Response.flush
					
					oConn.execute(strSQL)
					
					''*** Opdaterer seraft, så aftalen ER faktureret ***
					'select case intFakbetalt
					'case 1
					'erfak = 1 'godkendt
					'case 0
					'erfak = 2 'kladde
					'end select
					
					'oConn.execute("UPDATE serviceaft SET erfaktureret = "& erfak &" WHERE id = "& aftaler(i) &"")
					
					
				
				
				
				'*** Opretter posteringer på valgte konti ***	
				
				intFaknum = lastFaknr
				
				'*** Henter fak id ***
							
							strSQL3 = "SELECT Fid FROM fakturaer"
							oRec3.open strSQL3, oConn, 3
							oRec3.movelast
							if not oRec3.EOF then
							thisfakid = cint(oRec3("Fid")) 
							end if 
							oRec3.close
						
						
				'** Opretter tilhørende posteringer hvis det ikke er kladde. **		
				if intFakbetalt <> 0 then 
						
						posteringsTxt = strBesk
						call opretpos()
						
				end if '** kladde / endelig
					
					
									
				next
				
				
				
				'if len(request("aftaleid")) <> 0 then
				'aftid = request("aftaleid")
				'else
				'aftid = 0
				'end if
				
				if len(request("FM_aftnr")) <> 0 then
				sogaftnr = request("FM_aftnr")
				else
				sogaftnr = 0
				end if
				
				if len(request("thiskid")) <> 0 then
				thiskid = request("thiskid")
				else
				thiskid = 0
				end if
	
	
			Response.redirect "fak_serviceaft_osigt.asp?menu=stat_fak&FM_aftnr="&sogaftnr&"&id="&thiskid
				
				
	'***
	end if
	
				
	
	
			
			
end if 'FM_fak_alle	
end if '5


%>

<!--#include file="../inc/regular/footer_inc.asp"-->


