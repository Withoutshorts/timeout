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
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:290; top:155; width:60%; height:600; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="400" bgcolor="#ffffe1">
	<tr>
	    <td colspan=2 style="padding-top:5; padding-left:30; border-left:1px darkred solid; border-right:1px darkred solid; border-top:1px darkred solid;"><img src="../ill/alert.gif" width="44" height="45" alt="" border="0">Du er ved at <b>slette</b> en faktura-skrivelse.<br>
		Er dette korrekt?<br><br>&nbsp;</td>
	</tr>
	<tr>
	   <td width=250 style="padding:5; border-left:1px darkred solid; border-bottom:1px darkred solid;"><a href="javascript:history.back()"><img src="../ill/pile_tilbage.gif" width="10" height="6" alt="" border="0">&nbsp;Nej</a></td>
	   <td align=right style="padding:5; border-right:1px darkred solid; border-bottom:1px darkred solid;"><a href="fak_org.asp?menu=stat_fak&func=sletok&id=<%=id%>&FM_job=<%=Request("FM_job")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>">Ja&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"></a></td>
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
	'*** Her slettes en medarbejder ***
	'strSQL = "SELECT fid, oprid FROM fakturaer WHERE Fid = "& id 
	'oRec.open strSQL, oConn, 3 
	'if not oRec.EOF then 
	'	thisoprid = oRec("oprid")
	'end if
	'oRec.close 
	
	oConn.execute("DELETE FROM fakturaer WHERE Fid = "& id &"")
	oConn.execute("DELETE FROM faktura_det WHERE fakid = "& id &"")
	oConn.execute("DELETE FROM fak_med_spec WHERE fakid = "& id &"")
	
	'if thisoprid <> 0 then
	'oConn.execute("DELETE FROM posteringer WHERE oprid = "& thisoprid &"")
	'end if
	
	if request("FM_usedatointerval") = "1" then
	usedval = 1
	else
	usedval = 0
	end if
	
	Response.redirect "fak_osigt.asp?menu=stat_fak&FM_job=" & Request("FM_job") &"&FM_usedatointerval="&usedval&"&FM_start_dag="&request("FM_start_dag")&"&FM_start_mrd="&request("FM_start_mrd")&"&FM_start_aar="&request("FM_start_aar")&"&FM_slut_dag="&request("FM_slut_dag")&"&FM_slut_mrd="&request("FM_slut_mrd")&"&FM_slut_aar="&request("FM_slut_aar")&"&FM_fakint="&request("FM_fakint")
	
	case "dbred", "dbopr" 
	
	
	'Public function kommaFunc(x)
	'if len(x) <> 0 then
	'instr_komma = instr(x, ",")
	'	
	'	if instr_komma > 0 then
	'	left_intX = left(x, instr_komma - 1)
	'	rstr = mid(x, instr_komma, 3)
	'	left_intX = left_intX & rstr
	'	else
	'	left_intX = x
	'	end if
	'else
	'left_intX = 0
	'end if
	
	'Response.write left_intX 
	'end function
	
	
	function SQLBlessDOT(s)
	dim tmpdot
	tmpdot = s
	tmpdot = replace(tmpdot, ".", "")
	SQLBlessDOT = tmpdot
	end function
	
	function SQLBless(s)
	dim tmp
	tmp = s
	tmp = replace(tmp, ",", ".")
	SQLBless = tmp
	end function
	
	function SQLBlessPunkt(s)
	dim tmpPunkt
	tmpPunkt = s
	tmpPunkt = replace(tmpPunkt, ".", ",")
	SQLBlessPunkt = tmpPunkt
	end function
	
	intBeloeb = SQLBless(Request("FM_Beloeb"))
	intTimer = SQLBless(Request("FM_Timer"))
	intFaknum = SQLBless(Request("FM_faknr"))
	
	'Response.write "intFaknum: " & intFaknum & " / opr: " &  SQLBless(Request("FM_faknr_opr"))
	
	'*** Her tjekkes om alle required felter er udfyldt. ***
	if len(Request("FM_faknr")) = 0 OR len(Request("FM_timer")) = 0 OR len(Request("FM_beloeb")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<%
	errortype = 26
	call showError(errortype)
	
	else%>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(intBeloeb)
			if isInt > 0 then
			%>
			<!--#include file="../inc/regular/header_inc.asp"-->
			<!--#include file="../inc/regular/topmenu_inc.asp"-->
			<%
			errortype = 27
			call showError(errortype)
			
			isInt = 0
					else
					
					call erDetInt(intTimer)
					if isInt > 0 then
					%>
					<!--#include file="../inc/regular/header_inc.asp"-->
					<!--#include file="../inc/regular/topmenu_inc.asp"-->
					<%
					errortype = 28
					call showError(errortype)
					isInt = 0
							
							else
							strSQL2 = "SELECT faknr FROM fakturaer WHERE faknr = '"& intFaknum &"' AND Fid <> "& id &""
							oRec2.open strSQL2, oConn, 3
							
							if not oRec2.EOF then
							%>
							<!--#include file="../inc/regular/header_inc.asp"-->
							<!--#include file="../inc/regular/topmenu_inc.asp"-->
							<%
							errortype = 29
							call showError(errortype)
							oRec2.close
							
							else
							
							
							'*** Hvis alle required er udfyldt ***
							
									fakDato = Request("FM_start_aar") & "/" & Request("FM_start_mrd") & "/" & Request("FM_start_dag") '& time 
									showfakDato = Request("FM_start_dag") & "/" & Request("FM_start_mrd") & "/" & Request("FM_start_aar")
									strEditor = session("user")
									strDato = session("dato")
									strjobid = Request("jobid")
									intFakbetalt = request("FM_betalt")
									intfakadr = request("FM_Kid")
									
									if request("FM_att_vis") = "on" then
									strAtt = request("FM_att")
									else
									strAtt = "Ingen"
									end if
									
									if len(request("FM_faktype")) <> 0 then
									intFaktype = request("FM_faktype")
									else
									intFaktype = 0
									end if
									
									if intFakbetalt <> 1 then
									intFakbetalt = 0
									else
									intFakbetalt = intFakbetalt 
									end if
									
									'** Database format dato **
									dtb_dato = Request("FM_b_start_aar") & "/" & Request("FM_b_start_mrd") & "/" & Request("FM_b_start_dag")
									
									
									function SQLBlessK(s)
									dim tmp
									tmp = s
									tmp = replace(tmp, "'", "''")
									SQLBlessK = tmp
									end function
									
									'************************
									'*** findes allerede ****
									'************************
									function SQLBless2(s)
									dim tmp2
									tmp2 = s
									tmp2 = replace(tmp2, ",", ".")
									SQLBless2 = tmp2
									end function
									
									function SQLBless3(s)
									dim tmp3
									tmp3 = s
									tmp3 = replace(tmp3, ".", ",")
									SQLBless3 = tmp3
									end function
									
									
									varKonto = request("FM_kundekonto")
									varModkonto = request("FM_modkonto")
									strKomm = SQLBlessK(Request("FM_Komm"))
									antalAkt = Request("antal_x")
									
									subtotal = 0
									
									
									'******************************* Beregner moms ********************************
									intTotal = intBeloeb
									
									
									if request("FM_debkre") = "d" then
									intTotal_konto = intTotal
									intTotal_modkonto = 0 - SQLBless3(intTotal)
									intTotal_modkonto = SQLBless2(intTotal_modkonto)
									else
									intTotal_konto = 0 - SQLBless3(intTotal)
									intTotal_konto = SQLBless2(intTotal_konto)
									intTotal_modkonto = intTotal
									end if
									
									intMoms = 0
									intNetto = 0
									intModMoms = 0
									intModNetto = 0
									showmoms = 0
									
									
									'** Beregner moms og netto **
									'*** Konto ***
									strSQL = "SELECT momskode, kvotient FROM kontoplan LEFT JOIN momskoder ON (momskoder.id = momskode) WHERE kontonr = "& varKonto
									oRec.open strSQL, oConn, 3 
									if not oRec.EOF then
										if oRec("kvotient") <> 1 then
											
												if oRec("kvotient") = 5 then
												intMoms = (SQLBless3(intTotal_konto) * 0.25)  'Alm
												else
												intMoms = (SQLBless3(intTotal_konto) * 0.05)  'Rep
												end if
											
												if intMoms < 0 then
												showmoms = -(intMoms)
												else
												showmoms = intMoms
												end if
											
											intMoms = SQLBless2(intMoms) 
											
										else
										intMoms = 0
										end if
									intNetto = intTotal_konto 
									end if
									oRec.close
									
									
									'** Beregner moms og netto **
									'*** Modkonto ***
									strSQL = "SELECT momskode, kvotient FROM kontoplan LEFT JOIN momskoder ON (momskoder.id = momskode) WHERE kontonr = "& varModkonto
									oRec.open strSQL, oConn, 3 
									if not oRec.EOF then
										if oRec("kvotient") <> 1 then
												if oRec("kvotient") = 5 then
												intModMoms = (SQLBless3(intTotal_modkonto) * 0.25)  'Alm
												else
												intModMoms = (SQLBless3(intTotal_modkonto) * 0.05)  'Rep
												end if
												
										intModMoms = SQLBless2(intModMoms) 
										
										else
										intModMoms = 0
										end if
									intModNetto = intTotal_modkonto 'SQLBless2((intTotal_modkonto - (formatnumber(intModMoms, 2))))
									end if
									oRec.close
									'***********************************************************************************
									
									
									
									
							'************************************************************************************
							'*** Opdaterer / Redigerer faktura 													*
							'************************************************************************************
											if func = "dbred" then
											oConn.execute("UPDATE fakturaer SET"_
											&" faknr = '"& intFaknum &"', "_
											&" fakdato = '"& fakDato &"', "_
											&" timer = "& intTimer  &", "_
											&" beloeb = "& intBeloeb &", "_
											&" kommentar = '"& strKomm &"', "_
											&" editor = '"& strEditor &"', "_
											&" dato = '"& strDato &"', "_
											&" tidspunkt = '23:59:59', "_
											&" betalt = "& intFakbetalt &", "_
											&" b_dato = '"& dtb_dato &"', "_
											&" fakadr = "& intfakadr &", "_
											&" att = '"& strAtt &"', "_
											&" konto = "& varKonto &", "_
											&" modkonto = "& varModkonto &", "_
											&" faktype = "& intFaktype &""_
											&" WHERE Fid = "& id &"")
											
											strSQL = "SELECT oprid FROM fakturaer WHERE Fid="& id
											oRec.open strSQL, oConn, 3 
											if not oRec.EOF then
											oprid = oRec("oprid")
											end if
											oRec.close 
											
											
											if oprid <> 0 then '0 hvis fak er fra før 1/8-2004
											oconn.execute("DELETE FROM posteringer WHERE oprid ="&oprid&"")
											end if
											
													if intFakbetalt <> 0 then '** Opretter tilhørende posteringer hvis det ikke er kladde. **	
													'***************** Redigerer posteringer på de valgte konti **********
													oConn.execute("INSERT INTO posteringer (editor, dato, modkontonr, kontonr, bilagsnr, beloeb, nettobeloeb, moms, tekst, posteringsdato, status, att, oprid)"_
													&" VALUES ('"& strEditor &"', '"& strDato &"', '"& varModkonto &"', '"& varKonto &"', '"& intFaknum &"', "_
													&" "& intTotal_konto &", "&intNetto&", "&intMoms&", '"& request("FM_jobnavn") &"', "_
													&" '"& fakDato &"', "_
													&" 1, "& session("Mid") &", "& oprid &")")
													
													
													'**** Modkonto ***
													oConn.execute("INSERT INTO posteringer (editor, dato, modkontonr, kontonr, bilagsnr, beloeb, nettobeloeb, moms, tekst, posteringsdato, status, att, oprid)"_
													&" VALUES ('"& strEditor &"', '"& strDato &"', '"& varKonto &"', '"& varModkonto &"', '"& intFaknum &"', "_
													&" "& intTotal_modkonto &", "&intModNetto&", "&intModMoms&", '"& request("FM_jobnavn") &"', "_
													&" '"& fakDato &"', "_
													&" 1, "& session("Mid") &", "& oprid &")")
													'*********************************************************************
													end if
											end if
											
											
											
							'************************************************************************************
							'*** Opretter faktura *********														*
							'************************************************************************************
											if func = "dbopr" then
												
												if intFaktype <> 0 then '*** kreditnota eller rykker
												parentfak = request("id")
												else
												parentfak = 0
												end if
											
													strSQL = ("INSERT INTO fakturaer"_
													&" (faknr, fakdato, jobid, timer, beloeb, kommentar, dato, editor, tidspunkt, betalt, b_dato, fakadr, att, faktype, konto, modkonto, parentfak) VALUES ("_
													&" '"& intFaknum &"',"_
													&" '"& fakDato &"',"_
													&" "& strjobid &","_
													&" "& intTimer &","_
													&" "& intBeloeb &","_
													&" '"& strKomm &"',"_
													&" '"& strDato &"',"_
													&" '"& strEditor &"', '23:59:59', "_
													&" "& intFakbetalt &", '"& dtb_dato &"', "& intfakadr &", '"& strAtt &"', "& intFaktype &", '"& varKonto &"', '"& varModkonto &"', "& parentfak &")")
													
													'Response.write strSQL
													oConn.execute(strSQL)
													
													
														'** Henter fak id ***
														strSQL3 = "SELECT Fid FROM fakturaer"
														oRec3.open strSQL3, oConn, 3
														oRec3.movelast
														if not oRec3.EOF then
														thisfakid = cint(oRec3("Fid")) 
														end if 
														oRec3.close
													
													if intFakbetalt <> 0 then '** Opretter tilhørende posteringer hvis det ikke er kladde. **
													
														
														'***************** Opretter posteringer på de valgte konti **********
														strSQLK = "INSERT INTO posteringer (editor, dato, modkontonr, kontonr, bilagsnr, beloeb, nettobeloeb, moms, tekst, posteringsdato, status, att)"_
														&" VALUES ('"& strEditor &"', '"& strDato &"', '"& varModkonto &"', '"& varKonto &"', '"& intFaknum &"', "_
														&" "& intTotal_konto &", "&intNetto&", "&intMoms&", '"& request("FM_jobnavn") &"', "_
														&" '"& fakDato &"', "_
														&" 1, "& session("Mid") &")"
														
														oConn.execute(strSQLK)
														
														'*** Finder det netop opr. id ***
														strSQL = "SELECT id FROM posteringer ORDER BY id DESC"
														oRec.Open strSQL, oConn, 3 
														if not oRec.EOF then
														oprid = oRec("id")
														end if
														oRec.close
														
														oConn.execute("UPDATE posteringer SET oprid = "& oprid &" WHERE id = "& oprid &"")
														oConn.execute("UPDATE fakturaer SET oprid = "& oprid &" WHERE Fid = "& thisfakid &"")
														
														'**** Modkonto ***
														strSQLMK = "INSERT INTO posteringer (editor, dato, modkontonr, kontonr, bilagsnr, beloeb, nettobeloeb, moms, tekst, posteringsdato, status, att, oprid)"_
														&" VALUES ('"& strEditor &"', '"& strDato &"', '"& varKonto &"', '"& varModkonto &"', '"& intFaknum &"', "_
														&" "& intTotal_modkonto &", "&intModNetto&", "&intModMoms&", '"& request("FM_jobnavn") &"', "_
														&" '"& fakDato &"', "_
														&" 1, "& session("Mid") &", "& oprid &")"
														
														oConn.execute(strSQLMK)
														
																'*********************************************************************
																'*** Hvis Rykker oprettes krediteres tidligere faktura ***************
																'********************************************************************* 
																if intFaktype = 2 then
																strSQL = "SELECT oprid FROM fakturaer WHERE Fid="& request("id")
																'Response.write strSQL & "<br><br>"
																oRec.open strSQL, oConn, 3 
																if not oRec.EOF then
																oprid = oRec("oprid")
																end if
																oRec.close 
																
																
																	if oprid <> 0 then '** undgå tidligere fakturaer fra før økonomi del.
																	
																	strSQL = "SELECT editor, dato, modkontonr, "_
																	&" kontonr, bilagsnr, beloeb, nettobeloeb, moms, "_
																	&" tekst, posteringsdato, status, att "_
																	&" FROM posteringer WHERE oprid=" & oprid
																	
																	oRec.open strSQL, oConn, 3 
																	if not oRec.EOF then
																		
																		useModkonto = oRec("modkontonr")
																		useKonto = oRec("kontonr")
																		useTotal_konto = oRec("beloeb")
																		useNetto = oRec("nettobeloeb")
																		useMoms = oRec("moms")
																		
																	end if
																	oRec.close 
																		
																		if trim(len(request("bilagsnrkredit"))) <> 0 then
																		useFaknum = request("bilagsnrkredit")
																		else
																		useFaknum = 0
																		end if
																	
																	strSQLK = "INSERT INTO posteringer (editor, dato, modkontonr, kontonr, bilagsnr, beloeb, nettobeloeb, moms, tekst, posteringsdato, status, att, oprid)"_
																	&" VALUES ('"& strEditor &"', '"& strDato &"', '"& useModkonto &"', '"& useKonto &"', '"& useFaknum &"', "_
																	&" "& -(SQLBless2(useTotal_konto)) &", "&-(SQLBless2(useNetto))&", "&-(SQLBless2(useMoms))&", '"& request("FM_jobnavn") &"', "_
																	&" '"& fakDato &"', "_
																	&" 1, "& session("Mid") &", "& oprid &")"
																	
																	oConn.execute(strSQLK)
																	'Response.write strSQLK
																	
																	'**** Modkonto ***
																	strSQLMK = "INSERT INTO posteringer (editor, dato, modkontonr, kontonr, bilagsnr, beloeb, nettobeloeb, moms, tekst, posteringsdato, status, att, oprid)"_
																	&" VALUES ('"& strEditor &"', '"& strDato &"', '"& useKonto &"', '"& useModkonto &"', '"& useFaknum &"', "_
																	&" "& SQLBless2(useTotal_konto) &", "&SQLBless2(useNetto)&", "&SQLBless2(useMoms)&", '"& request("FM_jobnavn") &"', "_
																	&" '"& fakDato &"', "_
																	&" 1, "& session("Mid") &", "& oprid &")"
																	
																	oConn.execute(strSQLMK)
																	end if
														end if '** kladde / endelig 
												end if
												'*********************************************************************
													
												
										end if	
							'**************************
							'** Slut opret / rediger db
							'**************************
														
												
							'***********************************************************************************
							'*** Hvis faktura allerede er oprrettet en gang i denne session ***
							'***********************************************************************************
							if session("erOprettet") <> "y" then
								'*** h styrer height på den tabel der skal printes ud ***
								h = 0
								 
								'************************************************************** 
								'*** Indsætter akt i fak_det 
								'**************************************************************
								if func = "dbred" then
								oConn.execute("DELETE FROM faktura_det WHERE fakid = "& id &"")
								thisfakid = id
								end if
								
								'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
								for intcounter = 0 to antalAkt - 1
								
									if len(request("FM_show_akt_"& intcounter &"")) = 1 then
										
										timerThis = request("FM_timerthis_"& intcounter &"")
										if len(timerThis) <> 0 then
										timerThis = SQLBless2(timerThis)
										else 
										timerThis = 0
										end if
										
										if len(request("FM_beloebthis_"& intcounter &"")) <> 0 then
										beloebThis = SQLBless2(request("FM_beloebthis_"& intcounter &""))
										else
										beloebThis = 0
										end if
										
										if len(request("FM_enhedspris_"& intcounter &"")) <> 0 then
										enhpris = SQLBless2(request("FM_enhedspris_"& intcounter &""))
										else
										enhpris = 0
										end if
										
										'response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
										'Response.write "timerThis" & timerThis & "<br>"
										'response.write "beloebThis" & beloebThis & "<br>"
										'Response.write "enhpris"& enhpris & "<br>"
										
										oConn.execute("INSERT INTO faktura_det "_
										&" (antal, beskrivelse, aktpris, fakid, enhedspris, aktid) "_
										&" VALUES ("& timerThis &", "_
										&"'" & SQLBlessK(request("FM_aktbesk_"& intcounter &"")) &"', "_
										&""  & beloebThis &", "_
										&""  & thisfakid &", "& enhpris &", "& request("FM_aktid_"&intcounter&"") &" )")
										
										
									varHojde = len(request("FM_aktbesk_"& intcounter &""))
									if varHojde < 50 then
									useh = 40
									end if
									if varHojde > 50 AND varHojde < 150 then
									useh = 60
									end if
									if varHojde > 150 then
									useh = 80
									end if
									
									
									'**** opbygning af faktura timer til skærm ***	
									strFakdet = strFakdet &"<tr><td valign='top'>&nbsp;</td>"_
									&"<td><img src='../ill/blank.gif' width='20' height='1' alt='' border='0'></td>"_
									&"<td valign=top align=right style='padding-right:2;'>"& formatnumber(SQLBlessPunkt(timerThis), 2) &"</td>"_
									&"<td colspan=2 valign=top style='padding-left:5;'>" & request("FM_aktbesk_"& intcounter &"") &"</td>"_
									&"<td valign=top align=right>" & formatnumber(request("FM_enhedspris_"& intcounter &""), 2) &"&nbsp;dkr.</td>"_
									&"<td valign='top' align=right colspan=2>"& formatnumber(SQLBlessPunkt(request("FM_beloebthis_"& intcounter &"")),2) &"&nbsp;dkr.<img src='../ill/blank.gif' width='20' height='1' alt='' border='0'></td>"_
									&"<td valign='top' align='right'>&nbsp;</td>"_
									&"</tr>"
									
									subtotal = subtotal + request("FM_beloebthis_"& intcounter &"")
									
												
											'***** Medarbejder udspecificering ************
											antalmedspec = request("antal_n_"&intcounter&"") 'medarbejdere
											
											for intcounter2 = 0 to antalmedspec - 1
											
												if request("FM_show_medspec_"&intcounter2&"_"&intcounter&"") = "show" then
												
													'**** opbygning af medarbejder udspecificering til skærm ***	
													strFakdet = strFakdet &"<tr><td valign='top'>&nbsp;</td>"_
													&"<td><img src='../ill/blank.gif' width='20' height='1' alt='' border='0'></td>"_
													&"<td valign=top align=right style='padding-right:2;'><font class=megetlillesort>"& formatnumber(SQLBlessPunkt(request("FM_m_fak_"&intcounter2&"_"&intcounter&"")), 2) &"</td>"_
													&"<td colspan=2 valign=top style='padding-left:5;'><font class=megetlillesort>" & request("FM_m_tekst_"& intcounter2&"_"&intcounter&"") &"</td>"_
													&"<td valign=top align=right><font class=megetlillesort>" & formatnumber(request("FM_mtimepris_"& intcounter2&"_"&intcounter&""), 2) &"&nbsp;dkr.</td>"_
													&"<td valign='top' align=right colspan=2><font class=megetlillesort>"& formatnumber(SQLBlessPunkt(request("FM_mbelob_"& intcounter2 &"_"&intcounter&"")),2) &"&nbsp;dkr.<img src='../ill/blank.gif' width='20' height='1' alt='' border='0'></td>"_
													&"<td valign='top' align='right'>&nbsp;</td>"_
													&"</tr>"
												
												end if
												
												if len(request("FM_mid_"&intcounter2&"_"&intcounter&"")) <> 0 then
												thisMid = request("FM_mid_"&intcounter2&"_"&intcounter&"")
												else
												thisMid = 0
												end if 
												
												oConn.execute("DELETE FROM fak_med_spec WHERE fakid = "& thisfakid &" AND aktid = "&request("FM_aktid_"&intcounter&"")&" AND mid = "&thisMid&"")
											 	
												
												if len(request("FM_mbelob_"& intcounter2 &"_"&intcounter&"")) <> 0 then
												useBeloeb = SQLBless2(request("FM_mbelob_"& intcounter2 &"_"&intcounter&""))
												else
												useBeloeb = 0
												end if
												
												
												if len(request("FM_m_vent_"&intcounter2&"_"&intcounter&"")) <> 0 then
												useVenter = SQLBless2(request("FM_m_vent_"&intcounter2&"_"&intcounter&""))
												else
												useVenter = 0
												end if
												
												
												oConn.execute("UPDATE fak_med_spec SET venter = 0 WHERE aktid = "&request("FM_aktid_"&intcounter&"")&" AND mid = "&thisMid&"")
											 	
												
												if len(request("FM_m_fak_"&intcounter2&"_"&intcounter&"")) <> 0 then
												useFak = SQLBless2(request("FM_m_fak_"&intcounter2&"_"&intcounter&""))
												else
												useFak = 0
												end if
												
												if len(request("FM_mtimepris_"& intcounter2&"_"&intcounter&"")) <> 0 then
												usemedTpris = SQLBless2(request("FM_mtimepris_"& intcounter2&"_"&intcounter&""))
												else
												usemedTpris = 0
												end if 
												
												if request("FM_show_medspec_"&intcounter2&"_"&intcounter&"") = "show" then
												showonfak = 1
												else
												showonfak = 0
												end if
												
												if thisMid <> 0 then
												strSQL = ("INSERT INTO fak_med_spec (fakid, aktid, mid, fak, venter, tekst, enhedspris, beloeb, showonfak) "_
												&" VALUES ("& thisfakid &", "&request("FM_aktid_"&intcounter&"")&", "_
												&"" &request("FM_mid_"&intcounter2&"_"&intcounter&"")&", "_
												&"" &useFak&", "_
												&"" &useVenter&", "_
												&"'"&SQLBlessK(request("FM_m_tekst_"& intcounter2&"_"&intcounter&""))&"', "_
												&"" &usemedTpris&", "_
												&"" &useBeloeb&", "& showonfak &")")
												oConn.execute(strSQL)
												end if
											next
												
									
									h = h + useh
									end if
								next
								
							session("erOprettet") = "y"
							else
							strFakdet = "<tr><td valign='top'><img src='../ill/tabel_top.gif' width='1' height='65' alt='' border='0'></td>"_
							&"<td><img src='../ill/blank.gif' width='20' height='1' alt='' border='0'></td>"_
							&"<td colspan='7'><font class='error'><br><br>Denne side er forældet!</font><br>&nbsp;"_
							&"<td valign='top' align='right'><img src='../ill/tabel_top.gif' width='1' height='65' alt='' border='0'></td>"_
							&"</tr>"
							end if
							'**********************************************************************************************
							
							
												
												
											'*** Viser den oprettede faktura til print *****%>
											<!-------------------------------Sideindhold------------------------------------->
											<html>
											<head>
												<title>TimeOut 2.1</title>
												<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
											</head>
											<body topmargin="0" leftmargin="0" class="regular">
											<div id="sindhold" style="position:absolute; left:10; top:5; visibility:visible;">
											<br><br>
											<table cellspacing="0" cellpadding="0" border="0" width="600">
												<tr><td>
												
												<%select case intFaktype
												case 0
												strtopimg = "faktura_top"
												strTypeThis = "faktura"
												strTypeThis2 = "Faktura"  
												case 1
												strtopimg = "kreditnota_top"
												strTypeThis = "kreditnota"
												strTypeThis2 = "Kreditnota"   
												case 2
												strtopimg = "rykker_top"
												strTypeThis = "rykker"
												strTypeThis2 = "Rykker"  
												case else
												strtopimg = "blank"
												end select%>
												
												<img src="../ill/<%=strtopimg%>.gif" width="271" height="25" alt="" border="0">
												
												
												</td></tr>
											</table>
											<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#FFFFFF">
											<tr>
												<td valign="top" rowspan=2 width="10">&nbsp;</td>
												<td colspan="5" height="10">&nbsp;</td>
												<td valign="top" align=right rowspan=2>&nbsp;</td>
											</tr>
											<tr>
												<td style="padding-left : 4px;" valign="top" colspan="2">
												<%
												'* Modtager **
												intKnr = request("FM_knr")
												strKnavn = request("FM_knavn")
												strKadr = request("FM_adr")
												strKpostnr = request("FM_postnr")
												strBy = request("FM_by")
												
												strLand = request("FM_land")
												strAtt = request("FM_att")
												
												'**** finder att *****
												select case strAtt
												case 991
												showAtt = "Den økonomi ansvarlige"
												case 992
												showAtt = "Regnskabs afd."
												case 993
												showAtt = "Administrationen"
												case else
												
												strSQL3 = "SELECT id, navn FROM kontaktpers WHERE id="& strAtt
												oRec3.open strSQL3, oConn, 3
												if not oRec3.EOF then
												showAtt = oRec3("navn")
												else
												showAtt = ""
												end if
												oRec3.close
												
												end select
												
												intTlf = request("FM_tlf")
												intCvr = request("FM_cvr")
												
												'*** Afsender ***
												yourKnavn = request("FM_yknavn")
												yourbank = request("FM_bank")
												yourRegnr = request("FM_regnr")
												yourKontonr = request("FM_kontonr")
												yourCVR = request("FM_ycvr")
												yourNavn = request("FM_kkundenavn")
												yourAdr = request("FM_yadr")
												yourPostnr = request("FM_ypostnr")
												yourCity = request("FM_ycity")
												if request("FM_yland_vis") = "on" then
												yourLand = request("FM_yland")
												else
												yourLand = "19229332"
												end if
												if request("FM_yemail_vis") = "on" then
												yourEmail = request("FM_yemail")
												else
												yourEmail = ""
												end if
												if request("FM_ytlf_vis") = "on" then
												yourTlf = request("FM_ytlf")
												else
												yourTlf = ""
												end if
												yourSwift = request("FM_swift")
												yourIban = request("FM_iban")
												
												
													%>
													<table cellspacing="0" cellpadding="0" border="0" bgcolor="#F5f5f5">
													<tr>
														<td valign="top"><img src="../ill/tabel_top_left_wsmoke.gif" width="8" height="23" alt="" border="0"></td>
														<td valign="top"><img src="../ill/tabel_top.gif" width="300" height="1" alt="" border="0"></td>
														<td valign="top"><img src="../ill/tabel_right_left_wsmoke.gif" width="8" height="23" alt="" border="0"></td>
													</tr>
													<tr bgcolor="#f5f5f5">
														<td width="1"><img src="../ill/tabel_top.gif" width="1" height="180" alt="" border="0"></td>
														<td valign="top"><font size=1 color="#999999">Afsender:&nbsp;<%=yourKnavn%>&nbsp;&nbsp;<%=yourAdr%>&nbsp;&nbsp;<%=yourPostnr%>&nbsp;<%=yourCity%><br></font>
														<%=strKnavn%>
						
														<%if request("FM_cvr_vis") = "on" then
														Response.write "<br>CVR:&nbsp;" & intCvr 
														end if%>
														<br>
														<%=strKadr%><br>
														<%=strKpostnr%>&nbsp;&nbsp;<%=strBy%><br>
														
														<%if request("FM_land_vis") = "on" then
														Response.write strLand &"<br>"
														end if
														
														if request("FM_att_vis") = "on" then
														Response.write "Att: "& showAtt &"<br>"
														end if 
														%>
														
														<%if request("FM_tlf_vis") = "on" then
														Response.write "Tlf: "& intTlf
														end if%>
														
														<br><br>
														<font class="megetlillesort">Kundenr:&nbsp;<%=intKnr%></font>
														<br>
														<%=formatdatetime(showfakDato, 1)%><br>
														<%=strTypeThis2%> nr:&nbsp;<b><%=intFaknum%></b>
														<td width="8" align="right"><img src="../ill/tabel_top.gif" width="1" height="180" alt="" border="0"></td>
													</tr>
													<tr>
														<td valign="bottom"><img src="../ill/tabel_bund_left_wsmoke.gif" width="8" height="23" alt="" border="0"></td>
														<td valign="bottom"><img src="../ill/tabel_top.gif" width="300" height="1" alt="" border="0"></td>
														<td valign="bottom"><img src="../ill/tabel_bund_right_wsmoke.gif" width="8" height="23" alt="" border="0"></td>
													</tr>
													</table>
													</td>
												<td width="150"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
												<td style="padding-right : 25px;" valign="top" colspan="2">
												<!-- Afsender -->
													<table cellspacing="0" cellpadding="0" border="0" bgcolor="#F5f5f5">
													<tr>
														<td valign="top"><img src="../ill/tabel_top_left_wsmoke.gif" width="8" height="23" alt="" border="0"></td>
														<td valign="top"><img src="../ill/tabel_top.gif" width="200" height="1" alt="" border="0"></td>
														<td valign="top"><img src="../ill/tabel_right_left_wsmoke.gif" width="8" height="23" alt="" border="0"></td>
													</tr>
													<tr bgcolor="#f5f5f5">
														<td width="1"><img src="../ill/tabel_top.gif" width="1" height="180" alt="" border="0"></td>
														<td valign="top"><%=yourKnavn%><br>
														<%=yourAdr%><br>
														<%=yourPostnr%>&nbsp;<%=yourCity%><br>
														<%
														if yourLand <> "19229332" then
														Response.write yourLand & "<br>"
														end if%>
														<%=yourTlf%>, <%=yourEmail%><br><br>
														Bank:&nbsp;<%=yourBank%><br>
														Regnr:&nbsp; <%=yourRegnr%><br>
														Kontonr:&nbsp;<%=yourKontonr%><br>
														<%if request("FM_yswift_vis") = "on" then%>
														Swift:&nbsp;<%=yourSwift%><br>
														<%end if%>
														<%if request("FM_yiban_vis") = "on" then%>
														Iban:&nbsp;<%=yourIban%><br>
														<%end if%>
														CVR:&nbsp;&nbsp;&nbsp;<%=yourCVR%></td>
														<td width="8" align="right"><img src="../ill/tabel_top.gif" width="1" height="180" alt="" border="0"></td>
													</tr>
													<tr>
														<td valign="bottom"><img src="../ill/tabel_bund_left_wsmoke.gif" width="8" height="23" alt="" border="0"></td>
														<td valign="bottom"><img src="../ill/tabel_top.gif" width="200" height="1" alt="" border="0"></td>
														<td valign="bottom"><img src="../ill/tabel_bund_right_wsmoke.gif" width="8" height="23" alt="" border="0"></td>
													</tr>
													</table>
												</td>
											</tr>
											</table>
											
											<%if request("FM_viskomfor") = 0 then%>
											<br><br>
											<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#FFFFFF">
											<tr>
												<td style="padding-left:10;"><b>Kommentar og betalings betingeleser vedr. denne <%=strTypeThis%>.</b><br><%=request("FM_komm")%><br>&nbsp;</td>
											</tr>
											</table>
											<%end if%>
											
											<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#FFFFFF">
											<tr>
												<td valign='top' height='40'>&nbsp;</td>
												<td colspan=7>&nbsp;&nbsp;&nbsp;Vedrørende <b><%=request("FM_jobnavn")%></b><br></td>
												<td valign='top' align=right>&nbsp;</td>
											</tr>
											<tr>
												<td valign='top'>&nbsp;</td>
												<td width="20"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
												<td width=40 align=right style="padding-right:2;"><b>Antal</b></td>
												<td colspan="2" width="430" style="padding-left:5;"><b>Beskrivelse</b></td>
												<td align=right width="100"><b>Timepris</b>&nbsp;&nbsp;&nbsp;&nbsp;</td>
												<td align=right width="100"><b>Pris</b>&nbsp;&nbsp;&nbsp;&nbsp;</td>
												<td width="40"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
												<td valign='top' align=right>&nbsp;</td>
											</tr>
											<%=strFakdet%>
											<tr>
											</table>
											<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#FFFFFF">
											<tr>
												<td valign='top'>&nbsp;</td>
												<td width="20"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
												<td valign='top' width=300 style="padding-top:30;">Total antal:&nbsp;<b><%=formatnumber(SQLBlessPunkt(request("FM_timer")), 2)%></b>&nbsp;</td>
												<td valign='top' align=right style="padding-top:30;">Subtotal:&nbsp;&nbsp;</td>
												<td valign='top' align=right style="padding-top:30;">&nbsp;&nbsp;<b><%=formatnumber(SQLBlessPunkt(subtotal))%></b>&nbsp;dkr.</td>
												<td width="20"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
												<td valign='top' align=right>&nbsp;</td>
											</tr>
											<tr>
												<td valign='top'>&nbsp;</td>
												<td width="20"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0">
												<td valign='top'>&nbsp;</td>
												<td valign='top' align=right>Moms:&nbsp;&nbsp;</td>
												<td valign='top' align=right>&nbsp;&nbsp;<b>
												<%
												Response.write formatnumber(SQLBless3(showmoms),2)%></b>&nbsp;dkr.</td>
												<td width="20"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
												<td valign='top' align=right>&nbsp;</td>
											</tr>
											<tr>
												<td valign='top'>&nbsp;</td>
												<td width="20"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0">
												<td valign='top'>&nbsp;</td>
												<td valign='top' align=right><br>Total:&nbsp;&nbsp;</td>
												<td valign='top' align=right>&nbsp;&nbsp;____________<br>
												<%
												totalinclmoms = (SQLBless3(showmoms) + subtotal)
												%>
												&nbsp;&nbsp;<b><%=formatnumber(totalinclmoms,2)%></b>&nbsp;dkr.<br>
												&nbsp;&nbsp;____________</td>
												<td width="20"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
												<td valign='top' align=right>&nbsp;</td>
											</tr>
											</table>
											</div>
											
											
											<!--- footer table -->
											<div id="fakfooter" style="position:absolute; left:10; top:800; visibility:visible;">
											<%if request("FM_viskomfor") = 1 then%>
											<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#FFFFFF">
											<tr>
												<td valign='top'>&nbsp;</td>
												<td width="20" valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0">
												<td colspan="4"><b>Kommentar og betalings betingeleser vedr. denne <%=strTypeThis%>.</b><br><%=request("FM_komm")%><br>&nbsp;</td>
												<td valign='top' align=right>&nbsp;</td>
											</tr>
											</table>
											<%end if%>
											<table cellspacing="0" cellpadding="0" border="0" width="600">
											<tr>
												<td colspan="7" valign="bottom"><img src="../ill/fakfooter.gif" width="600" height="42" alt="" border="0"></td>
											</tr>
											</table>
											<br><br>&nbsp;
											</div>
											
											
											
											<div id="link" style="position:absolute; left:720; top:60; visibility:visible;">
											<table cellspacing="0" cellpadding="0" border="0" width="200">
											<tr>
												<td colspan="2" valign="bottom">Print denne side ud, <br>og sørg for selv at gemme en kopi.<br><br>
												<a href="Javascript:window.print()">Print&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a><br><br>
												<a href="fak_osigt.asp?menu=stat_fak&FM_job=<%=Request("FM_job")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag_ival")%>&FM_start_mrd=<%=request("FM_start_mrd_ival")%>&FM_start_aar=<%=request("FM_start_aar_ival")%>&FM_slut_dag=<%=request("FM_slut_dag_ival")%>&FM_slut_mrd=<%=request("FM_slut_mrd_ival")%>&FM_slut_aar=<%=request("FM_slut_aar_ival")%>&FM_fakint=<%=request("FM_fakint_ival")%>">Tilbage til faktura-oversigt&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
											<br><br><br>
											nb)<br> Indstil din browser til ikke at printe side- hoved og fod ud. Det gøres ved at klikke på menu punktet [Filer] (øverst til ventre) --> [Side opsætning].<br><br> Fjern teksten i sidehoved og sidefod.<br><br>
											For at spare på farven i din printer patron, bør du også sørge for at printeren ikke er indstillet til at udskrive baggrundsfarver.</td>
											</tr>
											</table>
											</div>
											<%
											
									end if
							end if
					end if
			end if
	case else
	'**************************** Opret/rediger Faktura Data ************************************
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	
	
	function setTimerTot(nummer) {
	if (window.event.keyCode == 37){ 
	} else {
		if (window.event.keyCode == 39){
		} else {
		
		var oprTimerThis;
		var nyTimerThis;
		var oprVerdi_timertot;
		var oprBeloebThis;
		var timeprisThis;
		var antalx;
		var nyTimertot = 0;
		var nyBeloebtot = 0;
		var varmoms = 0;
		var addmoms = 0;
		var tmpbel = 0;
		var tmp = 0;
		
		antalx = (document.all["antal_x"].value); 
		timeprisThis = (document.all["FM_hidden_timepristhis_" + nummer + ""].value.replace(",",".") / 1);
		oprTimerThis = (document.all["FM_hidden_timerthis_" + nummer + ""].value.replace(",",".") / 1);
		oprVerdi_timertot = (document.all["FM_Timer"].value / 1);
		nyTimerThis = (document.all["FM_timerthis_" + nummer + ""].value.replace(",",".") / 1);
		oprBeloebThis = (document.all["FM_beloebthis_" + nummer + ""].value.replace(",",".")/ 1);
		nyBeloeb = (nyTimerThis * timeprisThis);
		
		document.all["FM_beloebthis_" + nummer + ""].value = nyBeloeb; 
		document.all["FM_beloebthis_" + nummer + ""].value = document.all["FM_beloebthis_" + nummer + ""].value.replace(".",",");
		document.all["FM_hidden_timerthis_" + nummer + ""].value = nyTimerThis; 
		document.all["FM_timerthis_" + nummer + ""].value = document.all["FM_timerthis_" + nummer + ""].value.replace(".",",")
				
				//afrunder decimaler
				offSet_this = String(document.all["FM_beloebthis_" + nummer + ""].value);
				offSetL_this = String(document.all["FM_beloebthis_" + nummer + ""].length);
				pkt_this = offSet_this.indexOf(",");
				if (pkt_this > 0){
					document.all["FM_beloebthis_" + nummer + ""].value = offSet_this.substring(0, pkt_this + 3)
				}	
					
				//afrunder decimaler
				offSet_this = String(document.all["FM_hidden_timerthis_" + nummer + ""].value);
				offSetL_this = String(document.all["FM_hidden_timerthis_" + nummer + ""].length);
				pkt_this = offSet_this.indexOf(",");
				if (pkt_this > 0){
					document.all["FM_hidden_timerthis_" + nummer + ""].value = offSet_this.substring(0, pkt_this + 3)
				}
				
				//afrunder decimaler
				offSet_this = String(document.all["FM_timerthis_" + nummer + ""].value);
				offSetL_this = String(document.all["FM_timerthis_" + nummer + ""].length);
				pkt_this = offSet_this.indexOf(",");
				if (pkt_this > 0){
					document.all["FM_timerthis_" + nummer + ""].value = offSet_this.substring(0, pkt_this + 3)
				}
		
		
		for (i=0;i<antalx;i++){
		tmp = document.all["FM_hidden_timerthis_" + i + ""].value.replace(",",".");
		document.all["FM_hidden_timerthis_" + i + ""].value = document.all["FM_hidden_timerthis_" + i + ""].value.replace(".",",")
		nyTimertot = (nyTimertot + parseFloat(""+tmp+""));
			if (document.all["FM_beloebthis_" + i + ""].value !=""){
			tmpbel = document.all["FM_beloebthis_" + i + ""].value.replace(",",".");	
			document.all["FM_beloebthis_" + i + ""].value = document.all["FM_beloebthis_" + i + ""].value.replace(".",",")
			nyBeloebtot = (nyBeloebtot + parseFloat(""+tmpbel+""));	
			}
		}
		document.all["FM_Timer"].value = nyTimertot;
		document.all["FM_Timer"].value = document.all["FM_Timer"].value.replace(".",",")
		
				//afrunder decimaler
				offSet_this = String(document.all["FM_Timer"].value);
				offSetL_this = String(document.all["FM_Timer"].length);
				pkt_this = offSet_this.indexOf(",");
				if (pkt_this > 0){
					document.all["FM_Timer"].value = offSet_this.substring(0, pkt_this + 3)
				}
		
		
		document.all["FM_beloeb"].value = nyBeloebtot;
		document.all["FM_beloeb"].value = document.all["FM_beloeb"].value.replace(".",",")
		}
		}
		}
		
		
		
	function enhedspris(nummer){
	if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
				document.all["FM_hidden_timepristhis_" + nummer + ""].value = document.all["FM_enhedspris_" + nummer + ""].value.replace(",",".");
				document.all["FM_enhedspris_" + nummer + ""].value = document.all["FM_enhedspris_" + nummer + ""].value.replace(".",",")
							
							//afrunder decimaler
							offSet_this = String(document.all["FM_hidden_timepristhis_" + nummer + ""].value);
							offSetL_this = String(document.all["FM_hidden_timepristhis_" + nummer + ""].length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.all["FM_hidden_timepristhis_" + nummer + ""].value = offSet_this.substring(0, pkt_this + 3)
							}
							
							//afrunder decimaler
							offSet_this = String(document.all["FM_enhedspris_" + nummer + ""].value);
							offSetL_this = String(document.all["FM_enhedspris_" + nummer + ""].length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.all["FM_enhedspris_" + nummer + ""].value = offSet_this.substring(0, pkt_this + 3)
				}
			}
		}		
	}
	
	
	function enhedsprismedarb(x, n){
	if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
				document.getElementById("FM_mtimepris_" + n + "_" + x +"").value = document.getElementById("FM_mtimepris_" + n + "_" + x +"").value.replace(".",",")
							
							//afrunder decimaler
							offSet_this = String(document.getElementById("FM_mtimepris_" + n + "_" + x +"").value);
							offSetL_this = String(document.getElementById("FM_mtimepris_" + n + "_" + x +"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_mtimepris_" + n + "_" + x +"").value = offSet_this.substring(0, pkt_this + 3)
							}
							
				}
		}		
	}
	
	
	function updantaltimerprakt(x, n){
	if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
				var fak_opr = 0;
				var fak = 0;
				var mtimepris = 0;
				var vent = 0;
				
				fak_opr = document.getElementById("FM_m_fak_opr_" + n + "_" + x +"").value;
				fak = document.getElementById("FM_m_fak_" + n + "_" + x +"").value.replace(",",".") / 1; 
				vent = (parseFloat(fak_opr)) - (parseFloat(fak))
				
				
				if (vent < 0) {
				document.getElementById("FM_m_vent_" + n + "_" + x +"").value = 0
				}else{
				document.getElementById("FM_m_vent_" + n + "_" + x +"").value = vent
				document.getElementById("FM_m_vent_" + n + "_" + x +"").value = document.getElementById("FM_m_vent_" + n + "_" + x +"").value.replace(".",",")
				}
				
				mtimepris = document.getElementById("FM_mtimepris_" + n + "_" + x +"").value.replace(",",".")
				document.getElementById("FM_mbelob_" + n + "_" + x +"").value = parseFloat(mtimepris) * parseFloat(fak)
				
				var nyBeloebtot = 0;
				var tmpbel = 0;
				var antalmedarbtimertot = 0;
				var tmptim = 0;
				
				//antaln = document.getElementById("antal_n").value
				antal_n_x = document.getElementById("antal_n_"+x+"").value
				
					for (i=0;i<antal_n_x;i++){
					//alert(antal_n_x +" / "+i)
					tmpbel = document.getElementById("FM_mbelob_" + i + "_" + x +"").value.replace(",",".")
					nyBeloebtot = nyBeloebtot + parseFloat(""+tmpbel+"");
					tmptim = document.getElementById("FM_m_fak_" + i + "_" + x +"").value.replace(",",".") 
					antalmedarbtimertot = antalmedarbtimertot + parseFloat(tmptim)
					}
					
					document.getElementById("FM_mbelob_tot_" + x +"").value = parseFloat(nyBeloebtot)
					document.getElementById("FM_m_fak_tot_" + x +"").value = antalmedarbtimertot
					
					//replace . med ,
					document.getElementById("FM_mbelob_" + n + "_" + x +"").value = document.getElementById("FM_mbelob_" + n + "_" + x +"").value.replace(".",",")
					document.getElementById("FM_mbelob_tot_" + x +"").value = document.getElementById("FM_mbelob_tot_" + x +"").value.replace(".",",")
					document.getElementById("FM_m_fak_tot_" + x +"").value = document.getElementById("FM_m_fak_tot_" + x +"").value.replace(".",",")
					document.getElementById("FM_m_fak_" + n + "_" + x +"").value = document.getElementById("FM_m_fak_" + n + "_" + x +"").value.replace(".",",")
					
					
					//afrunder decimaler
							offSet_this = String(document.getElementById("FM_mbelob_" + n + "_" + x +"").value);
							offSetL_this = String(document.getElementById("FM_mbelob_" + n + "_" + x +"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_mbelob_" + n + "_" + x +"").value = offSet_this.substring(0, pkt_this + 3)
							}
							
							offSet_this = String(document.getElementById("FM_mbelob_tot_" + x +"").value);
							offSetL_this = String(document.getElementById("FM_mbelob_tot_" + x +"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_mbelob_tot_" + x +"").value = offSet_this.substring(0, pkt_this + 3)
							}
							
							offSet_this = String(document.getElementById("FM_m_fak_tot_" + x +"").value);
							offSetL_this = String(document.getElementById("FM_m_fak_tot_" + x +"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_m_fak_tot_" + x +"").value = offSet_this.substring(0, pkt_this + 3)
							}
							
							offSet_this = String(document.getElementById("FM_m_fak_" + n + "_" + x +"").value);
							offSetL_this = String(document.getElementById("FM_m_fak_" + n + "_" + x +"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_m_fak_" + n + "_" + x +"").value = offSet_this.substring(0, pkt_this + 3)
							}
			
			}
		}
	}
	
	
	function overfortilakt(x){
	document.getElementById("FM_timerthis_" + x +"").value = document.getElementById("FM_m_fak_tot_" + x +"").value.replace(",",".")
	
	var nyenhedspris = 0;
				nyenhedspris = parseFloat(document.getElementById("FM_mbelob_tot_" + x +"").value.replace(",",".") / document.getElementById("FM_timerthis_" + x +"").value)
				document.all["FM_hidden_timepristhis_" + x + ""].value = parseFloat(nyenhedspris)
				document.all["FM_enhedspris_" + x + ""].value = parseFloat(nyenhedspris)
				document.all["FM_enhedspris_" + x + ""].value = document.all["FM_enhedspris_" + x + ""].value.replace(".",",")
				
				//Afrunding
				offSet_this = String(document.getElementById("FM_enhedspris_" + x + "").value);
							offSetL_this = String(document.getElementById("FM_enhedspris_" + x + "").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_enhedspris_" + x + "").value = offSet_this.substring(0, pkt_this + 3)
							}
	setTimerTot(x)
	}
	
	
	function setBeloebTot(nummer) {
			if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
				
				var antalx;
				var nyBeloebtot = 0;
				var tmpbel = 0;
				
				antalx = (document.all["antal_x"].value); 
				
				for (i=0;i<antalx;i++){
					if (document.all["FM_beloebthis_" + i + ""].value !=""){
					tmpbel = document.all["FM_beloebthis_" + i + ""].value.replace(",",".");	
					nyBeloebtot = nyBeloebtot + parseFloat(""+tmpbel+"");	
					document.all["FM_beloebthis_" + i + ""].value = document.all["FM_beloebthis_" + i + ""].value.replace(".",",");
					
						//afrunder decimaler
						offSet_this = String(document.all["FM_beloebthis_" + i + ""].value);
						offSetL_this = String(document.all["FM_beloebthis_" + i + ""].length);
						pkt_this = offSet_this.indexOf(",");
						if (pkt_this > 0){
							document.all["FM_beloebthis_" + i + ""].value = offSet_this.substring(0, pkt_this + 3)
						}
					}
				}
				document.all["FM_beloeb"].value = nyBeloebtot;
				document.all["FM_beloeb"].value = document.all["FM_beloeb"].value.replace(".",",")
				
				//afrunder decimaler
				offSet = String(document.all["FM_beloeb"].value);
				offSetL = String(document.all["FM_beloeb"].value.length);
				pkt = offSet.indexOf(",");
					if (pkt > 0){
						document.all["FM_beloeb"].value = offSet.substring(0, pkt + 3)
					}
				}
			}	
	}
		
	
	function isNum(passedVal){
	invalidChars = " /:;<>abcdefghijklmnopqrstuvwxyzæøå"
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
	
	function tjektimer(nummer){
	var oprTimerThisTj = 0;
	oprTimerThisTj = (document.all["FM_hidden_timerthis_" + nummer + ""].value / 1);
		if (!validZip(document.all["FM_timerthis_" + nummer + ""].value)){
		alert("Der er angivet et ugyldigt tegn.")
		document.all["FM_timerthis_" + nummer + ""].value = oprTimerThisTj;
		document.all["FM_timerthis_" + nummer + ""].focus()
		document.all["FM_timerthis_" + nummer + ""].select()
		return false
		}
	return true
	}
	
	function tjekBeloeb(nummer){
		if (!validZip(document.all["FM_beloebthis_" + nummer + ""].value)){
		alert("Der er angivet et ugyldigt tegn.")
		setTimerTot(nummer)
		return false
		}
	return true
	}
	//-->
	</script>
	
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<%
	public strDag
	public strMrd
	public strMrdNavn
	public function dag() 
		if func <> "red" then
				
			Select case strMrd
			case "4", "6", "9", "11"
				if strDag > 30 then
				strDag = 30
				else
				strDag = strDag
				end if
			case 2
				if strDag > 28 then
				strDag = 28
				else
				strDag = strDag
				end if
			case else
			strDag = strDag
			end select
			
		end if
			
				%>
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
				<option value="31">31</option></select>&nbsp;
				<%
	end function
	
	function mrd()
		
		select case strMrd
		case "1"
		strMrdNavn = "jan" 
		case "2"
		strMrdNavn = "feb"
		case "3"
		strMrdNavn = "mar"
		case "4"
		strMrdNavn = "apr"
		case "5"
		strMrdNavn = "maj"
		case "6"
		strMrdNavn = "jun"
		case "7"
		strMrdNavn = "jul"
		case "8"
		strMrdNavn = "aug"
		case "9"
		strMrdNavn = "sep"
		case "10"
		strMrdNavn = "okt"
		case "11"
		strMrdNavn = "nov"
		case "12"
		strMrdNavn = "dec"
		end select
		%>
		<option value="<%=strMrd%>" selected><%=strMrdNavn%></option>
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
	<%
	end function
	
	
	function Aar()
	%>
	<option value="<%=strAar%>">
		<%if id <> 0 then%>
		20<%=strAar%>
		<%else%>
		<%=strAar%>
		<%end if%></option>
		<option value="02">2002</option>
		<option value="03">2003</option>
	   	<option value="04">2004</option>
	   	<!--<option value="05">2005</option>
		<option value="06">2006</option>
		<option value="07">2007</option>--></select>
	<%
	end function
	
	
	sub top
	%>
	<div id="sidetop" style="position:absolute; border:2; left:<%=pleft+0%>; top:<%=ptop%>; visibility:visible; z-index:100;">
	<br><img src="../ill/header_fak.gif" alt="" border="0" width="279" height="34">
	<br>
	<%
	end sub
	%>
	
	
	<%
	jobnr = Request("jobnr")
	jobid = request("jobid")
	timeforbrug = 0
	faktureretBeloeb = 0
	
	pleft = 200
	ptop = 50
	
	Redim thisaktid(0)
	Redim thisaktnavn(0)
	Redim thisAktTimer(0)
	Redim thisAktBeloeb(0)
	Redim preserve thisTimePris(0)
	
	if func = "red" then
	'*********************************************************************
	' Rediger faktura 
	'*********************************************************************
	
	
						strSQL = "SELECT Fid, faknr, fakdato, jobid, timer, beloeb, kommentar, dato, editor, betalt, b_dato, fakadr, att, faktype, konto, modkonto FROM fakturaer WHERE Fid = "& id 
						
						oRec.open strSQL, oConn, 3
						
						if not oRec.EOF then
						strEditor = oRec("editor")
						strFaknr = oRec("faknr")
						strTdato = formatdatetime(oRec("fakdato"), 2)
								'*** Hvis faktype er <> for faktura skiftes dbfunc til dbopr ***
								if request("faktype") <> 0 AND request("rykkreopr") = "j" then
									dbfunc = "dbopr"
									strNote = "<font color=darkred>Husk at ændre nr. i forhold til det oprindelige faktura nummer!</font>"
									strNote1 = "Nedenstående er en kopi af den oprindelige faktura."
									rykelnotopr = 1
										Select case request("faktype")
										case 1
										strKom = "Kontakt os venligst vedrørende overførsel / tilbagebetaling."
										case 2
										strKom = "Hvis jeres indbetaling har krydset dette brev, så se venligst bort fra denne skrivelse."
										End select
								else
									rykelnotopr = 0
									strNote = ""
									dbfunc = "dbred"
									strNote1 = ""
									strKom = oRec("kommentar")
								end if
								
						strTimer = oRec("timer")
						strBeloeb = oRec("beloeb")
						StrUdato = "12/12/2007"
						strDato = oRec("dato")
						intBetalt = oRec("betalt")
						strB_dato = oRec("b_dato")
						intFakid = oRec("Fid")
						intfakadr = oRec("fakadr")
						strAtt = oRec("att")
						intFaktype = oRec("faktype")
						intKonto = oRec("konto")
						intModKonto = oRec("modkonto")
						end if
						oRec.close
						
						
						if len(intModKonto) <> 0 then
						intModKonto = intModKonto
						else
						intModKonto = 0
						end if
						
						if intBetalt = 1 then
						betaltch = "checked"
						else
						betaltch = ""
						end if 
						
						strSQL = "SELECT beloeb, timer, fakdato, tidspunkt FROM fakturaer WHERE jobid = " & jobid
						oRec.open strSQL, oConn, 3
						while not oRec.EOF 
							faktureretBeloeb = faktureretBeloeb + oRec("beloeb")
							timeforbrug = timeforbrug + oRec("timer")
							lastFakdato = convertDateYMD(oRec("fakdato"))
							strfaktidspkt = formatdatetime(oRec("tidspunkt"), 3) 
							oRec.movenext
						Wend
						oRec.close
						
						strSQL = "SELECT jobTpris, budgettimer, fastpris, jobknr, jobnr, jobnavn FROM job WHERE id = " & jobid
						oRec.open strSQL, oConn, 3
						if not oRec.EOF then
							intBudgettimer = oRec("budgettimer")
							intJobTpris = oRec("jobTpris")
							fastpris = oRec("fastpris")
							if intfakadr <> 0 then
							intjobknr = intfakadr
							else
							intjobknr = oRec("jobknr")
							end if
							intjobnr = oRec("jobnr")
							strJobnavn = oRec("jobnavn")
						end if
						oRec.close
		'*******************************************************************************************
	
	
		call top
		'** Mulighed for at slette og se sidst redigeret dato ****
		%><img src="../ill/blank.gif" width="10" height="325" alt="" border="0">
		<table cellspacing="0" cellpadding="5" border="0" width="370" bgcolor="#EFF3FF">
		<%
		if rykelnotopr <> 1 then%>
			<tr>
				<td valign="bottom" style="border:1px #003399 solid;" height="25">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b>
				<!--&nbsp;&nbsp;<a href="fak.asp?menu=stat_fak&func=slet&id=<=id%>&FM_job=<=Request("FM_job")%>"><img src="../ill/slet_eks.gif" width="20" height="20" alt="Slet denne faktura" border="0"></a>-->
				</td>
			</tr>
		<%end if%>
		</table>
		</div>
		<%
	else
	call top
	%>
	</div>
	
	
	<%
		'************************************************************************'
		' Opret faktura 
		'************************************************************************
		
			strSQL = "SELECT Fid, faknr FROM fakturaer ORDER BY Fid DESC" 
			oRec.open strSQL, oConn, 3
			
			lastFaknrFundet = 0
			while not oRec.EOF AND lastFaknrFundet <> 1
			call erDetInt(oRec("faknr"))
				if isInt = 0 then
				lastFaknr = oRec("faknr") + 1
				lastFaknrFundet = 1
				end if
			oRec.movenext
			wend
			oRec.close
		
		
		strSQL = "SELECT beloeb, timer, fakdato, tidspunkt FROM fakturaer WHERE jobid = " & jobid &" ORDER BY fakdato"
		oRec.open strSQL, oConn, 3
		while not oRec.EOF 
			faktureretBeloeb = faktureretBeloeb + oRec("beloeb")
			timeforbrug = timeforbrug + oRec("timer")
			lastFakdato = convertDateYMD(oRec("fakdato")) 
			strfaktidspkt = oRec("tidspunkt")
		oRec.movenext
		Wend
		oRec.close
		
		strSQL = "SELECT jobTpris, budgettimer, fastpris, jobknr, jobnr, jobnavn FROM job WHERE id = " & jobid
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
			intBudgettimer = oRec("budgettimer")
			intJobTpris = oRec("jobTpris")
			fastpris = oRec("fastpris")
			'*** Hvis der er skiftet kunde ****
			if len(request("newkundeid")) <> 0 then
			intjobknr = request("newkundeid")
			else
			intjobknr = oRec("jobknr")
			end if
			intjobnr = oRec("jobnr")
			strJobnavn = oRec("jobnavn")
		end if
		oRec.close
		
		intKonto = 0
		intModKonto = 0
		
		strFaknr = lastFaknr
		strEditor = ""
		strTdato = "" 'month(now)&"/"&day(now)&"/"& year(now)
		strKom = ""
		dbfunc = "dbopr"
		varSubVal = "Opretpil" 
		'strTimer = Request("ttf")
		'strBeloeb = Request("ktf")
		'*******************************************************************************************
	end if
	%>
	
	
	<div id="vaelgmodtager" style="position:absolute; left:<%=pleft%>; top:<%=ptop+310%>; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="300">
	<%if func <> "red" then%>
	<tr><form action="fak_org.asp?menu=stat_fak" method="post" name="skiftkunde">
		<td><br>Hvis ovenstående firmaoplysninger ikke passer,<br>kan der skiftes kunde her: <br>
		
			<!-- hvis der skiftes kunde mister man det valgte datointerval i fakkorsel. -->
			<input type="hidden" name="FM_jobnavn" value="<%=strJobnavn%>">
			<input type="hidden" name="FM_job" value="<%=Request("FM_job")%>">
			<input type="hidden" name="id" value="<%=id%>">
			<input type="hidden" name="jobnr" value="<%=jobnr%>">
			<input type="hidden" name="jobid" value="<%=jobid%>">
			<input type="hidden" name="ttf" value="<%=strTimer%>">
			<input type="hidden" name="ktf" value="<%=strBeloeb%>">
			<input type="hidden" name="faktype" value="<%=request("faktype")%>">
			<input type="hidden" name="faknr" value="<%=request("faknr")%>">
			<input type="hidden" name="rykkreopr" value="<%=request("rykkreopr")%>">
		<select name="newkundeid" style="font-family: arial,helvetica,sans-serif; font-size: 10px;">
		<%
		strSQL = "SELECT Kid, kkundenavn, kkundenr FROM kunder WHERE ketype <> 'e'" 		
		oRec.open strSQL, oConn, 3
		While not oRec.EOF
		if cint(intjobknr) = cint(oRec("Kid")) then
		ksel = "SELECTED"
		else
		ksel = ""
		end if
		%><option value="<%=oRec("Kid")%>" <%=ksel%>>(<%=oRec("kkundenr")%>)&nbsp;<%=oRec("kkundenavn")%></option><% 
		oRec.movenext
		wend
		oRec.close
		%>
		</select>
		<input type="image" src="../ill/pillillexp_tp.gif" name="skiftkunde"><br>
		<font class=megetlillesort>Hvis der skiftes kunde nulstilles et evt. valgt datointerval, og datoen for sidste faktura bruges som interval startdao.</font></form>
		</td>
	</tr>
	<%end if%>
	</table>
	</div>
	
	
	
	<%
	if func <> "red" then
	id = 0 'Sættes så dato2 virker!
	else
	id = id
	end if
	%>	
	<!--#include file="inc/dato2.asp"-->
	<form action="fak_org.asp?menu=stat_fak&func=<%=dbfunc%>" method="post">
			<%
			if request("FM_usedatointerval") = "1" then
			usedt_ival = 1
			else
			usedt_ival = 0
			end if
			%>
			<input type="hidden" name="FM_usedatointerval" value="<%=usedt_ival%>">
			<input type="hidden" name="FM_start_dag_ival" value="<%=request("FM_start_dag")%>">
			<input type="hidden" name="FM_start_mrd_ival" value="<%=request("FM_start_mrd")%>">
			<input type="hidden" name="FM_start_aar_ival" value="<%=request("FM_start_aar")%>">
			<input type="hidden" name="FM_slut_dag_ival" value="<%=request("FM_slut_dag")%>">
			<input type="hidden" name="FM_slut_mrd_ival" value="<%=request("FM_slut_mrd")%>">
			<input type="hidden" name="FM_slut_aar_ival" value="<%=request("FM_slut_aar")%>">
			
			<input type="hidden" name="FM_fakint_ival" value="<%=trim(request("FM_fakint"))%>">
			
	
	<input type="hidden" name="FM_faktype" value="<%=request("faktype")%>">
	<input type="hidden" name="FM_jobnavn" value="<%=strJobnavn%>">
	<input type="hidden" name="FM_job" value="<%=Request("FM_job")%>">
	<input type="hidden" name="id" value="<%=id%>">
	<input type="hidden" name="jobid" value="<%=jobid%>">
	
	<div id="modtager" style="position:absolute; left:<%=pleft%>; top:<%=ptop+50%>; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="370" bgcolor="#EFF3FF">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="31" alt="" border="0"></td>
		<td valign="top"><img src="../ill/tabel_top.gif" width="354" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="31" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class=alt><b>Modtager:</b></td>
	</tr>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td valign=top>
		<b>Type:</b>&nbsp;&nbsp;
		<%
		select case request("faktype")
		case "0"
		strFaktypeNavn = "Faktura"
		case "1"
		strFaktypeNavn = "Kreditnota"
		case "2"
		strFaktypeNavn = "Rykker"
		end select
		Response.write strFaktypeNavn
		%>
		<br>
		<%
		strSQL = "SELECT Kid, kkundenavn, kkundenr, adresse, postnr, city, land, telefon, cvr FROM kunder WHERE Kid =" & intjobknr  		
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then 
			intKnr = oRec("kkundenr")
			strKnavn = oRec("kkundenavn")
			strKadr = oRec("adresse")
			strKpostnr = oRec("postnr")
			strBy = oRec("city")
			strLand = oRec("land")
			intKid = oRec("Kid")
			intCVR = oRec("cvr")
			intTlf = oRec("telefon")
			'attA = oRec("kontaktpers1")
			'attB = oRec("kontaktpers2")
			'attC = oRec("kontaktpers3")
			'attD = oRec("kontaktpers4")
			'attE = oRec("kontaktpers5")
		end if
		
		oRec.close
		%>
		<input type="hidden" name="FM_Kid" value="<%=intKid%>">
		<input type="text" name="FM_knr" value="<%=intKnr%>" size=10 style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;&nbsp;
		<input type="text" name="FM_knavn" value="<%=strKnavn%>" size=40 style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><br>
		<input type="text" name="FM_adr" value="<%=strKadr%>" size=20 style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><br>
		<input type="text" name="FM_postnr" value="<%=strKpostnr%>" size=4 style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;&nbsp;
		<input type="text" name="FM_by" value="<%=strBy%>" size=20 style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><br>
		<input type="text" name="FM_land" value="<%=strLand%>" size=18 style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		&nbsp;&nbsp;&nbsp;Vis land på faktura:&nbsp;<input type="checkbox" name="FM_land_vis" value="on"><br>
		Att:&nbsp;<select name="FM_att" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		<%
		strSQLkpersCount = "SELECT id, navn FROM kontaktpers WHERE kundeid = "& intKid &" ORDER BY id"
		oRec2.open strSQLkpersCount, oConn, 3
			
			while not oRec2.EOF
				if strAtt = oRec2("id") then
				attSel = "SELECTED"
				else
				attSel = ""
				end if%>
				<option value="<%=oRec2("id")%>" <%=attSel%>><%=oRec2("navn")%></option>
				<%
			oRec2.movenext
			wend
			oRec2.close
			%>
		<%if strAtt = "991" then
		selth1 = "SELECTED"
		else
		selth1 = ""
		end if%>	
		<option value="991" <%=selth1%>>Den økonomi ansvarlige</option>
		<%if strAtt = "992" then
		selth2 = "SELECTED"
		else
		selth2 = ""
		end if%>	
		<option value="992" <%=selth2%>>Regnskabs afd.</option>
		<%if strAtt = "993" then
		selth3 = "SELECTED"
		else
		selth3 = ""
		end if%>	
		<option value="993" <%=selth3%>>Administrationen</option>
		</select>&nbsp;&nbsp;Vis att.&nbsp;<input type="checkbox" name="FM_att_vis" value="on" checked><br>
		<input type="text" name="FM_tlf" value="<%=intTlf%>" size=8 style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;&nbsp;&nbsp;Vis tlf. på faktura:&nbsp;<input type="checkbox" name="FM_tlf_vis" value="on"><br>
		<input type="text" name="FM_cvr" value="<%=intCVR%>" size=8 style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;&nbsp;&nbsp;Vis cvr på faktura:&nbsp;<input type="checkbox" name="FM_cvr_vis" value="on">
		<br>
		<b>Kundekonto:</b>
		<select name="FM_kundekonto">
		<option value="0">(0)&nbsp;&nbsp;Ingen konto</option>
		<%
			strSQL = "SELECT kontonr, navn, id, kid FROM kontoplan ORDER BY kontonr, navn"
			oRec.open strSQL, oConn, 3 
			while not oRec.EOF
			
			if func = "red" then
				if intKonto = oRec("kontonr") then
				selkon = "SELECTED"
				else
				selkon = ""
				end if
			else
				if intKid = oRec("kid") then
				selkon = "SELECTED"
				else
				selkon = ""
				end if
			end if
			
			%>
			<option value="<%=oRec("kontonr")%>" <%=selkon%>>(<%=oRec("kontonr")%>)&nbsp;&nbsp;<%=oRec("navn")%></option>
			<%
			oRec.movenext
			Wend 
			oRec.close
		%>
		</select>&nbsp;
		<select name="FM_debkre">
				<%
					if request("faktype") <> 1 then
					selK = ""
					selD = "SELECTED"
					else
					selK = "SELECTED"
					selD = ""
					end if
				
				%>
		<option value="k" <%=selK%>>Krediter</option>
		<option value="d" <%=selD%>>Debiter</option>
		</select>
		</td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<tr>
		<td colspan=3 style="border-right:1px #003399 solid; border-bottom:1px #003399 solid; border-left:1px #003399 solid;">&nbsp;</td>
	</tr>
	</table>
	</div>
	
	<!--- afsender af faktura --->
	<div id="afsender" style="position:absolute; left:<%=pleft+400%>; top:<%=ptop+50%>;">
	<table cellspacing="0" cellpadding="0" border="0" width="300" bgcolor="#FFFFFF">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="31" alt="" border="0"></td>
		<td valign="top"><img src="../ill/tabel_top.gif" width="284" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="31" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class=alt><b>Afsender:</b></td>
	</tr>
	<tr><td width="30" style="border-left:1px #003399 solid;">&nbsp;</td>
	<td>
		<%
		strSQL = "SELECT adresse, postnr, city, land, telefon, email, regnr, kkundenavn, kontonr, cvr, bank, swift, iban, kid FROM kunder WHERE useasfak = 1"
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then 
			yourbank = oRec("bank")
			yourRegnr = oRec("regnr")
			yourKontonr = oRec("kontonr")
			yourCVR = oRec("cvr")
			yourNavn = oRec("kkundenavn")
			yourAdr = oRec("adresse")
			yourPostnr = oRec("postnr")
			yourCity = oRec("city")
			yourLand = oRec("land")
			yourEmail = oRec("email")
			yourTlf = oRec("telefon")
			yourSwift = oRec("swift")
			yourIban = oRec("iban")
			afskid = oRec("kid")
		end if
		oRec.close
		%><br>
		<input type="text" name="FM_yknavn" value="<%=yourNavn%>" size=25 style="!border: 1px; background-color: whitesmoke; border-color: #86B5E4; border-style: solid;"><br>
		<input type="text" name="FM_yadr" value="<%=yourAdr%>" size=25 style="!border: 1px; background-color: whitesmoke; border-color: #86B5E4; border-style: solid;">
		<br><input type="text" name="FM_ypostnr" value="<%=yourPostnr%>" size=4 style="!border: 1px; background-color: whitesmoke; border-color: #86B5E4; border-style: solid;">&nbsp;&nbsp;
		<input type="text" name="FM_ycity" value="<%=yourCity%>" size=25 style="!border: 1px; background-color: whitesmoke; border-color: #86B5E4; border-style: solid;">
		<!--<input type="text" name="FM_yland" value="<%=yourLand%>" size=25 style="!border: 1px; background-color: whitesmoke; border-color: #86B5E4; border-style: solid;">
		<br>Vis land på faktura:&nbsp;<input type="checkbox" name="FM_yland_vis" value="on">-->
		<input type="text" name="FM_ytlf" value="<%=yourTlf%>" size=25 style="!border: 1px; background-color: whitesmoke; border-color: #86B5E4; border-style: solid;">
		<br>Vis telefon nr. på faktura:&nbsp;<input type="checkbox" name="FM_ytlf_vis" value="on" CHECKED>
		<input type="text" name="FM_yemail" value="<%=yourEmail%>" size=25 style="!border: 1px; background-color: whitesmoke; border-color: #86B5E4; border-style: solid;">
		<br>Vis email på faktura:&nbsp;<input type="checkbox" name="FM_yemail_vis" value="on" CHECKED>
		<br><br>
		Bank:&nbsp;&nbsp;<input type="text" name="FM_bank" value="<%=yourBank%>" size=25 style="!border: 1px; background-color: whitesmoke; border-color: #86B5E4; border-style: solid;">
		Regnr:&nbsp;<input type="text" name="FM_regnr" value="<%=yourRegnr%>" size=4 maxlength="4" style="!border: 1px; background-color: whitesmoke; border-color: #86B5E4; border-style: solid;">&nbsp;&nbsp;
		Kontonr:&nbsp;<input type="text" name="FM_kontonr" value="<%=yourKontonr%>" size=10 maxlength="12" style="!border: 1px; background-color: whitesmoke; border-color: #86B5E4; border-style: solid;"><br>
		Swift:&nbsp;<input type="text" name="FM_swift" value="<%=yourSwift%>" size=6 maxlength="6" style="!border: 1px; background-color: whitesmoke; border-color: #86B5E4; border-style: solid;">&nbsp;Vis:<input type="checkbox" name="FM_yswift_vis" value="on"><br>
		Iban:&nbsp;<input type="text" name="FM_iban" value="<%=yourIban%>" size=24 maxlength="30" style="!border: 1px; background-color: whitesmoke; border-color: #86B5E4; border-style: solid;">&nbsp;Vis:<input type="checkbox" name="FM_yiban_vis" value="on"><br>
		CVR:&nbsp;<input type="text" name="FM_ycvr" value="<%=yourCVR%>" size=8 maxlength="8" style="!border: 1px; background-color: whitesmoke; border-color: #86B5E4; border-style: solid;"><br>
		</td><td width="30" style="border-right:1px #003399 solid;">&nbsp;</td>
		</tr>
		<tr bgcolor="#2c962d"><td width="30" style="border-left:1px #003399 solid;">&nbsp;</td>
		<td class=alt><b>Modkonto:</b>&nbsp;
		<select name="FM_modkonto" style="width:200;">
		<option value="0">(0)&nbsp;&nbsp;Ingen konto</option>
		<%
				strSQL = "SELECT kontonr, navn, id, kid FROM kontoplan WHERE kid = "& afskid &" ORDER BY kontonr, navn"
				
				oRec.open strSQL, oConn, 3 
				while not oRec.EOF 
				if intModKonto = oRec("kontonr") then
				selkon = "SELECTED"
				else
				selkon = ""
				end if
				%>
				<option value="<%=oRec("kontonr")%>" <%=selkon%>>(<%=oRec("kontonr")%>)&nbsp;&nbsp;<%=oRec("navn")%></option>
				<%
				oRec.movenext
				Wend 
				oRec.close
		%>
		</select>
		</td><td width="30" style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<tr>
		<td colspan=3 style="border-right:1px #003399 solid; border-bottom:1px #003399 solid; border-left:1px #003399 solid;">&nbsp;</td>
	</tr>
	</table>
	</div>
	<!-- afsender slut -->
	
	
	
	
	
	
	
	
	<div id="aktiviteter" style="position:absolute; left:<%=pleft%>; top:<%=ptop+420%>; visibility:visible; z-index:100;">
	
	<!-- De enkelte aktiviteter --->
	<table cellspacing="0" cellpadding="0" border="0" width="700" bgcolor="#FFFFFF">
	<tr>
		<td colspan=2 style="padding-left:10; padding:3; border-left:1px #003399 solid; border-top:1px #003399 solid; border-bottom:1px #003399 solid;">(<%=jobnr%>)&nbsp;&nbsp;<b><%=strJobnavn%></b>
		&nbsp;&nbsp;&nbsp;&nbsp;<%=strFaktypeNavn%> nr:&nbsp;
		<input type="text" name="FM_faknr" value="<%=strFaknr%>" size="14" maxlength="10" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		<input type="hidden" name="FM_faknr_opr" value="<%=strFaknr%>">
		<%if request("faktype") = "2" then%>
		<br>
		Bilags nr. på kreditering af den oprindelige faktura:&nbsp;<input type="text" name="bilagsnrkredit" value="0" size="14" maxlength="10" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		<%end if%>
		<br><%=strNote%></td>
		<td style="padding-left:10; padding:3; border-top:1px #003399 solid; border-bottom:1px #003399 solid; border-right:1px #003399 solid;">
		<%=strFaktypeNavn%> dato:&nbsp;&nbsp;
		
		<%'** Bruger afg. dato fra interval som faktura dato-
		if cint(usedt_ival) = 1 and func <> "red" then
		strDag = request("FM_slut_dag")
		strMrd = request("FM_slut_mrd")
		strAar = request("FM_slut_aar")
		end if
		%>
		
		<select name="FM_start_dag">
		<%
		call dag()
		%>
		<select name="FM_start_mrd">
		<%
		call mrd()
		%>
		<select name="FM_start_aar">
		<%
		call Aar()
		%></td>
	</tr>
	</table>
	<br>
	<table cellspacing="0" cellpadding="0" border="0" width="700" bgcolor="#EFF3FF">
	<tr>
		<td valign="top" rowspan=3 width="10"><img src="../ill/tabel_top.gif" width="1" height="180" alt="" border="0"></td>
		<td height="10" colspan="4" style="border-top:1px #003399 solid;">&nbsp;</td>
		<td valign="top" align=right rowspan=3><img src="../ill/tabel_top.gif" width="1" height="180" alt="" border="0"></td>
	</tr>
	<tr>
		<td colspan=3 style="padding-left:10;" valign=top><b>Udspecificering.</b><br>Hvis en aktivitet ikke skal medtages på skrivelsen,
		<br>skal <b>"Antal"</b> rubrikken herunder være tomt.<br>
		Alle beløb er vejledende. Afrundering kan forekomme, kontroller derfor altid om afrundeing passer.
		<br><b><%=strNote1%></b>
		
		<%if request("faktype") = 0 then%>
			<br><b>Kladde / Endelig</b><br>
			<%if intBetalt <> 0 then%>
			Godkendt <input type="radio" name="FM_betalt" value="1" checked>
			<%else%>
			<div style="position:relative; width:300; border:1px darkred solid; padding-left:10px; background-color:#ffffe1;"><input type="radio" name="FM_betalt" value="0" checked>Opret som kladde <input type="radio" name="FM_betalt" value="1"> eller som godkend som endelig.</div>
			<font class="megetlillesort">Når en faktura godkendes som endelig, bliver der oprettet en tilhørende postering på de ovenstående valgte konti.</font>
			<%end if%>
		<%else%>
		<input type="hidden" name="FM_betalt" value="1">
		<%end if%>
		</td>
		<td>&nbsp;</td>
	</tr>
	<%
	if len(lastFakdato) <> 0 then
	lastFakdato = lastFakdato
	else
	lastFakdato = "2001/1/1"
	end if
	
	
	'*** Sætter datointerval eller bruger lastFakdato ****
	if cint(usedt_ival) = 1 then
	
		call datofindes(request("FM_start_dag"),request("FM_start_mrd"),request("FM_start_aar"))
		stdato = request("FM_start_aar")&"/"&request("FM_start_mrd")&"/"&dagparset
		showStDato = dagparset&"/"&request("FM_start_mrd")&"/"&request("FM_start_aar")
		
		call datofindes(request("FM_slut_dag"),request("FM_slut_mrd"),request("FM_slut_aar"))
		slutdato = request("FM_slut_aar")&"/"&request("FM_slut_mrd")&"/"&dagparset
		showSlutDato = dagparset&"/"&request("FM_slut_mrd")&"/"&request("FM_slut_aar") 
		useLastFakdato = 0
	
	else
		stdato_temp = dateadd("d", 1, lastFakdato)
		
		stdato = year(stdato_temp)&"/"&month(stdato_temp)&"/"&day(stdato_temp)
		slutdato = "2014/1/1"
		
		showStDato = stdato 'lastFakdato
		showSlutDato = slutdato 
		useLastFakdato = 1
	end if
	
	
	if cdate(stdato) > cdate(lastFakdato) AND cint(usedt_ival) = 1 then
		stdatoKri = stdato
		useLastFakdato = 0
	else
		stdatoKri_temp = dateadd("d", 1, lastFakdato)
		stdatoKri = year(stdatoKri_temp)&"/"&month(stdatoKri_temp)&"/"&day(stdatoKri_temp)
		showStDato = stdatoKri_temp 'lastFakdato
		useLastFakdato = 1
	
	end if
	'******************************************************
	
	if func <> "red" then%>
	<tr>
		<td colspan=3 style="padding-left:10;" valign=top><br><b>Periode afgrænsning:</b><br>
		Timer vist på denne faktura er alle indtastet i perioden: <b><%=formatdatetime(showStDato, 2)%></b>
		<%if useLastFakdato = 1 then%>
		&nbsp;(Dagen efter sidste fakturadato)&nbsp;
		<%end if%>
		til <b><%=formatdatetime(showSlutDato, 2)%></b></td>
	</tr>
	<%else%>
	<tr>
		<td colspan=3 style="padding-left:10;" valign=top>&nbsp;</td>
	</tr>
	<%end if%>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
		<td valign=bottom style="padding-left:5;"><b>Antal timer</b></td>
		<td valign=bottom>&nbsp;<b>Beskrivelse</b></td>
		<td valign=bottom><b>Timepris.</b><br>(anbefalet*)</td>
		<td valign=bottom width=82><b>Pris dkr.</b></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
	</tr>
	<%
	if func = "red" then
		
		'******************* Henter oprettede aktiviteter *******************************
		strSQL = "SELECT faktura_det.id, antal, faktura_det.beskrivelse, aktpris, job.fastpris, job.jobTpris, job.budgettimer, timer.timepris, faktura_det.enhedspris, faktura_det.aktid FROM faktura_det LEFT JOIN job ON (job.jobnr = "& jobnr &") LEFT JOIN timer ON (timer.Tjobnr = '"& jobnr &"') WHERE fakid = "& intFakid &" GROUP BY id" 
		'Response.write strSQL
		oRec.open strSQL, oConn, 3
		x = 0
		While not oRec.EOF 
		
			Redim preserve thisaktnavn(x)
			thisaktnavn(x) = oRec("beskrivelse") 
			Redim preserve thisAktTimer(x)
			thisAktTimer(x) = oRec("antal")
			Redim preserve thisAktBeloeb(x)
			thisAktBeloeb(x) = oRec("aktpris")
			Redim preserve thisTimePris(x)
			thisTimePris(x) = oRec("enhedspris")
			Redim preserve thisaktid(x)
			thisaktid(x) = oRec("aktid")	
		
		x = x + 1
		oRec.movenext
		wend 
		oRec.close
	
	else
			'**** Sidste faktura tidspunkt ****
			if len(strfaktidspkt) > 0 then
			LastFakTid = datepart("h", strfaktidspkt)&":"& datepart("n", strfaktidspkt)&":"& datepart("s", strfaktidspkt)
			else
			LastFakTid =  "00:00:02"
			end if
			
			'*** Henter aktiviter og finder registrede timer siden sidste faktura dato. *** 
			strSQL = "SELECT job.fastpris, job.jobtpris, job.budgettimer, aktstatus, aktiviteter.job, aktiviteter.id AS aid, aktiviteter.navn AS anavn FROM aktiviteter LEFT JOIN job ON (job.id = aktiviteter.job) WHERE aktiviteter.job = "& jobid &" AND aktiviteter.aktstatus = 1 ORDER BY aktiviteter.id"		
			oRec.open strSQL, oConn, 3
			
			x = 0
			lastAktId = 0
			timerThisT = 0
			timerThisD = 0
			strKom = "Netto Kontant 8 dage."
			totaltimer = 0
			lastaktid = 0
			
			
			while not oRec.EOF
				if cint(lastaktid) <> cint(oRec("aid")) then
						
						Redim preserve thisaktid(x)
						thisaktid(x) = oRec("aid")
						
						'Response.write lastaktid & " # "
						strSQL2 = "SELECT avg(timer.timepris) AS timepristhis FROM timer WHERE taktivitetid =" & oRec("aid") &" AND tfaktim = 1 AND (tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"') OR taktivitetid =" & oRec("aid") &" AND tfaktim = 1"  'AND (Tdato = '" & stdatoKri &"' AND timer.tidspunkt > '"& LastFakTid &"')"		
						oRec2.open strSQL2, oConn, 3
						
						if not oRec2.EOF then
							thisavrTimepris = oRec2("timepristhis")
						end if
						oRec2.close
							
							
							if len(thisavrTimepris) <> 0 then
							thisavrTimepris = SQLBlessDot(formatnumber(thisavrTimepris, 2))
							else
							thisavrTimepris = 0
							end if
							
							'response.write thisavrTimepris & "<br>"
							
						strSQL2 = "SELECT sum(timer.timer) AS timerthis FROM timer WHERE taktivitetid =" & oRec("aid") &" AND tfaktim = 1 AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')"   ' OR Tdato = '" & stdatoKri &"' AND timer.tidspunkt > '"& LastFakTid &"')"		
						'Response.write strSQL2 &"<br><br>"
						oRec2.open strSQL2, oConn, 3
						
						if not oRec2.EOF then
						
							if len(oRec2("timerthis")) <> 0 then
							timerThisD = oRec2("timerthis")
							else
							timerThisD = 0
							end if
							
							
							if cint(oRec("fastpris")) = 1 then
								if timerThisD <> 0 then
								thisBeloebD = SQLBlessDot(formatnumber(((oRec("jobtpris")/oRec("budgettimer")) * oRec2("timerthis")), 2))
								else
								thisBeloebD = 0
								end if
							else
								if len(oRec2("timerthis")) <> 0 then
								thisBeloebD = SQLBlessDot(formatnumber((oRec2("timerthis") * thisavrTimepris), 2))
								else
								thisBeloebD = 0
								end if
							end if
						
						end if
						oRec2.close	
						
						if timerThisD > 0 then
						timerThisD = timerThisD
						else
						timerThisD = 0
						end if
						
						'Response.write timerThisD & " - "
						
						if thisBeloebD > 0 then
						thisBeloebD = thisBeloebD
						else
						thisBeloebD = 0
						end if
					
					
					Redim preserve thisaktnavn(x)
					thisaktnavn(x) = oRec("anavn") 
					Redim preserve thisAktTimer(x)
					thisAktTimer(x) = timerThisD
					Redim preserve thisAktBeloeb(x)
					thisAktBeloeb(x) = thisBeloebD
					timerThisD = 0
					thisBeloebD = 0
					
					'** Finder anbefalet timepris ****
					if oRec("budgettimer") = 0 then
					usejobbudgettimer = 1
					else
					usejobbudgettimer = oRec("budgettimer")
					end if 
					
					Redim preserve thisTimePris(x)
					if cint(oRec("fastpris")) = 1 then 'fastprisjob
						if len(oRec("jobtpris")/usejobbudgettimer) <> 0 then
						thisTimePris(x) = formatnumber(oRec("jobtpris")/usejobbudgettimer, 2)
						thisTimePris(x) = SQLBlessDOT(thisTimePris(x))
						else
						thisTimePris(x) = 0
						end if
					else
						if len(thisavrTimepris) <> 0 then
						thisTimePris(x) = formatnumber(thisavrTimepris, 2) 'hvis der findes avg timepris på allerede índtastede timer.
						thisTimePris(x) = SQLBlessDOT(thisTimePris(x))
						else
							
							if len(oRec("jobtpris")/usejobbudgettimer) <> 0 then
							thisTimePris(x) = formatnumber(oRec("jobtpris")/usejobbudgettimer, 2) 'Ellers bruges faspris.
							thisTimePris(x) = SQLBlessDOT(thisTimePris(x))
							else
							thisTimePris(x) = 0
							end if
						end if
					end if
						
				x = x + 1		
				end if
			
			lastaktid = oRec("aid")		
			oRec.movenext
			wend
			oRec.close
	
	end if
	
	
	
	
	'********************* Udskriver aktiviterer *********************
	for x = 0 to x - 1 
	n = 0
	
					'*** Medarbejder udspecificering ***
					strSQL3 = "SELECT projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter WHERE id = "& thisaktid(x)
					oRec3.open strSQL3, oConn, 3
					if not oRec3.EOF then
					
						gp1 = oRec3("projektgruppe1")
						gp2 = oRec3("projektgruppe2")
						gp3 = oRec3("projektgruppe3")
						gp4 = oRec3("projektgruppe4")
						gp5 = oRec3("projektgruppe5")
						gp6 = oRec3("projektgruppe6")
						gp7 = oRec3("projektgruppe7")
						gp8 = oRec3("projektgruppe8")
						gp9 = oRec3("projektgruppe9")
						gp10 = oRec3("projektgruppe10")
					
					else
						
						gp1 = 1
						gp2 = 1
						gp3 = 1
						gp4 = 1
						gp5 = 1
						gp6 = 1
						gp7 = 1
						gp8 = 1
						gp9 = 1
						gp10 = 1
					
					end if
					oRec3.Close 
					
					if thisaktid(x) <> 0 then%>
					<tr>
						<td style="border-left:1px #003399 solid;">&nbsp;</td>
						<td><font class=megetlillesort>Vis på fak.</td>
						<td align=right style="padding-right:20;">(<%=thisaktnavn(x)%>)<font class=megetlillesort>&nbsp;&nbsp;&nbsp;Navn&nbsp;&nbsp;<img src="../ill/blank.gif" width="270" height="1" alt="" border="0">F<img src="../ill/blank.gif" width="30" height="1" alt="" border="0">V</td>
						<td><font class=megetlillesort>Timepris</td>
						<td><font class=megetlillesort>Beløb</td>
						<td style="border-right:1px #003399 solid;">&nbsp;</td>
					</tr>
					<%end if
					faktot = 0
					strSQL10 = "SELECT DISTINCT(medarbejderid), tmnr, mnavn, mid, timeprisalt, medarbejdertype FROM progrupperelationer, medarbejdere LEFT JOIN "_
					&" timer ON (Tmnr = medarbejderid AND taktivitetid ="& thisaktid(x) &" AND tfaktim = 1 AND Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')"_  
					&" LEFT JOIN timepriser ON (aktid = "& thisaktid(x) &" AND medarbid = medarbejderid) WHERE mid = medarbejderid AND (projektgruppeid = "& gp1 &" OR projektgruppeid = "& gp2 &" OR projektgruppeid = "& gp3 &" OR projektgruppeid = "& gp4 &" OR projektgruppeid = "& gp5 &" OR projektgruppeid = "& gp6 &" OR projektgruppeid = "& gp7 &" OR projektgruppeid = "& gp8 &" OR projektgruppeid = "& gp9 &" OR projektgruppeid = "& gp10 &") GROUP BY medarbejderid ORDER BY projektgruppeid"
					'OR Tdato = '" & stdatoKri &"' AND timer.tidspunkt > '"& LastFakTid &"'
					oRec3.open strSQL10, oConn, 3
					
					'Response.write strSQL10 & "<br><br>"
					while not oRec3.EOF
								
								
								if func = "red" then
									
									strSQL2 = "SELECT fak, venter, tekst, enhedspris, beloeb, mid, showonfak FROM fak_med_spec WHERE fakid = "& id &" AND aktid = "& thisaktid(x) &" AND mid = "&oRec3("mid")
									'Response.write strSQL2
									oRec2.open strSQL2, oConn, 3 
									if not oRec2.EOF then 
									
									fak = oRec2("fak")
									venter = oRec2("venter")
									medarbejderTimepris = oRec2("enhedspris")
									medarbejderBeloeb = oRec2("beloeb")
									txt = oRec2("tekst")
									medarbid = oRec2("mid")		
									showonfak = oRec2("showonfak")
									
									end if
									oRec2.close 
									
									
										if len(medarbejderTimepris) <> 0 then
										medarbejderTimepris = SQLBlessDOT(formatnumber(medarbejderTimepris, 2))
										else
										medarbejderTimepris = SQLBlessDot(formatnumber(0, 2))
										end if
								
								else
										
											
										'*** timer medarbejder ***
										strSQL2 = "SELECT sum(timer) AS timer FROM timer WHERE Tmnr = "&oRec3("medarbejderid")&" AND taktivitetid ="& thisaktid(x) &" AND tfaktim = 1 AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')" 'OR Tdato = '" & stdatoKri &"' AND timer.tidspunkt > '"& LastFakTid &"'
										oRec2.open strSQL2, oConn, 3 
										
										if not oRec2.EOF then 
											medarbejderTimer = oRec2("timer")
										end if
										oRec2.close
										
										if len(medarbejderTimer) <> 0 then
											medarbejderTimer = medarbejderTimer
										else
											medarbejderTimer = 0
										end if
										
										'*** timepris medarbejder ***
										strSQL2 = "SELECT avg(timepris) AS medarbejderTimepris FROM timer WHERE Tmnr = "&oRec3("medarbejderid")&" AND taktivitetid ="& thisaktid(x) &" AND tfaktim = 1 AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')" 'OR Tdato = '" & stdatoKri &"' AND timer.tidspunkt > '"& LastFakTid &"'
										oRec2.open strSQL2, oConn, 3 
										
										if not oRec2.EOF then 
											medarbejderTimepris = oRec2("medarbejderTimepris")
										end if
										oRec2.close
										
										
										
										if len(medarbejderTimepris) <> 0 then
										medarbejderTimepris = SQLBlessDOT(formatnumber(medarbejderTimepris, 2))
										else
											
											strSQL2 = "SELECT timeprisalt, 6timepris, medarbejdertype FROM timepriser LEFT JOIN medarbejdere ON (mid = medarbid) WHERE medarbid = "&oRec3("medarbejderid")&" AND aktid ="& thisaktid(x) &" OR medarbid = "&oRec3("medarbejderid")&" AND aktid = 0 AND jobid = "& jobid
											oRec2.open strSQL2, oConn, 3 
											if not oRec2.EOF then
											
												if oRec2("timeprisalt") <> 6 then
														
														strSQL4 = "SELECT timepris, timepris_a1, timepris_a2, timepris_a3, timepris_a4, timepris_a5 FROM medarbejdertyper WHERE id = "& oRec2("medarbejdertype")
														oRec4.open strSQL4, oConn, 3 
														if not oRec4.EOF then
															select case oRec2("timeprisalt")
															case 0
															medarbejderTimepris = oRec4("timepris")
															case 1
															medarbejderTimepris = oRec4("timepris_a1")
															case 2
															medarbejderTimepris = oRec4("timepris_a2")
															case 3
															medarbejderTimepris = oRec4("timepris_a3")
															case 4
															medarbejderTimepris = oRec4("timepris_a4")
															case 5
															medarbejderTimepris = oRec4("timepris_a5")
															end select
															
																	
														end if
														oRec4.close 
												else
												medarbejderTimepris = SQLBlessDot(formatnumber(oRec2("6timepris"), 2))
												end if
											end if
											oRec2.close
										end if
										
										if len(medarbejderTimepris) <> 0 then
										medarbejderTimepris = medarbejderTimepris
										else
										medarbejderTimepris = SQLBlessDot(formatnumber(0, 2))
										end if
										
										strSQL2 = "SELECT sum(venter) AS venter FROM fak_med_spec WHERE aktid = "& thisaktid(x) &" AND mid = "&oRec3("mid") &" ORDER BY aktid" 
										'Response.write strSQL2
										oRec2.open strSQL2, oConn, 3 
										if not oRec2.EOF then 
											venter = oRec2("venter")
										end if
										oRec2.close
										
									if len(venter) <> 0 then
									venter = venter 
									else
									venter = 0
									end if
									
									
									
									fak = medarbejderTimer + venter 'oRec3("timer")
									if len(fak) <> 0 then
									fak = fak
									else
									fak = 0
									end if
									
									showventer = venter
									if len(showventer) <> 0 then
									showventer = showventer
									else
									showventer = 0
									end if
									venter = 0
									
									
									medarbejderBeloeb = (fak * medarbejderTimepris) 'oRec3("timepris")
									'Response.write oRec3("timer") &" #  " & medarbejderTimepris &": "& (oRec3("timer") * medarbejderTimepris)
									txt = oRec3("mnavn")
									medarbid = oRec3("mid")
									showonfak = 0		
											
								end if
								
								
							%>	
							<tr>
								<td style="border-left:1px #003399 solid;">&nbsp;</td>
								<td><input type="hidden" name="FM_mid_<%=n%>_<%=x%>" value="<%=medarbid%>">
								<%
								if showonfak = 1 then
								chkshow = "CHECKED"
								else
								chkshow = ""
								end if
								%>
								<input type="checkbox" name="FM_show_medspec_<%=n%>_<%=x%>" <%=chkshow%> id="FM_show_medspec_<%=n%>_<%=x%>" value="show"></td>
								<td style="padding-right:10;" align=right>&nbsp;
								<input type="text" name="FM_m_tekst_<%=n%>_<%=x%>" value="<%=txt%>" size=65 style="!border: 1px; background-color: #FFFFFF; border-color: silver; border-style: solid; font-size:9px">
								<input type="text" name="FM_m_fak_<%=n%>_<%=x%>" id="FM_m_fak_<%=n%>_<%=x%>" onKeyup="updantaltimerprakt('<%=x%>','<%=n%>');" value="<%=fak%>" size=2 style="!border: 1px; background-color: #FFFFFF; border-color: limegreen; border-style: solid; font-size:9px">&nbsp;
								<%if func <> "red" then%>
								<font class=megetlillesort>(<%=showventer%>)</font>
								<%else%>
							    &nbsp;&nbsp;&nbsp;&nbsp;
								<%end if%><input type="hidden" name="FM_m_fak_opr_<%=n%>_<%=x%>" id="FM_m_fak_opr_<%=n%>_<%=x%>" value="<%=fak%>">
								<input type="text" name="FM_m_vent_<%=n%>_<%=x%>" id="FM_m_vent_<%=n%>_<%=x%>" value="<%=venter%>" size=1 style="!border: 1px; background-color: #FFFFFF; border-color: silver; border-style: solid; font-size:9px"></td>
								<input type="hidden" name="FM_m_vent_opr_<%=n%>_<%=x%>" id="FM_m_vent_opr_<%=n%>_<%=x%>" value="<%=fak%>">
								<td><input type="text" name="FM_mtimepris_<%=n%>_<%=x%>" id="FM_mtimepris_<%=n%>_<%=x%>" onKeyup="enhedsprismedarb('<%=x%>','<%=n%>'), updantaltimerprakt('<%=x%>','<%=n%>');" value="<%=medarbejderTimepris%>" size=6 style="!border: 1px; background-color: #FFFFFF; border-color: #999999; border-style: solid; font-size:9px">
								</td>
								<td>
								<%
								if len(medarbejderBeloeb) <> 0 then
								mbelob = medarbejderBeloeb
								else
								mbelob = 0
								end if
								%>
								<input type="text" name="FM_mbelob_<%=n%>_<%=x%>" id="FM_mbelob_<%=n%>_<%=x%>" value="<%=SQLBlessDot(formatnumber(mbelob, 2))%>" size=8 style="!border: 1px; background-color: #FFFFFF; border-color: #999999; border-style: solid; font-size:9px">
								</td>
								<td style="border-right:1px #003399 solid;">&nbsp;</td>
							</tr>
							
							<%
					
					faktot = faktot + fak
					mbelobtot = mbelobtot + mbelob
					
					fak = 0 
					venter = 0
					medarbejderBeloeb = 0
					n = n + 1
					oRec3.movenext
					wend
					oRec3.close
					
					'** Medarbejdere total på aktivitet%>
					<tr>
								<td style="border-left:1px #003399 solid;">&nbsp;</td>
								<td colspan=2><img src="../ill/blank.gif" width="138" height="1" alt="" border="0">
								<input type="button" name="FM_show_medtot_<%=x%>" id="FM_show_medtot_<%=x%>" onClick="overfortilakt(<%=x%>)"; value="Overfør timer og beløb til aktiviten." style="!border: 1px; background-color: #d6dff5; border-color: #999999; border-style: solid; font-size:9px">
								<img src="../ill/blank.gif" width="75" height="1" alt="" border="0">
								<input type="text" name="FM_m_fak_tot_<%=x%>" id="FM_m_fak_tot_<%=x%>" value="<%=faktot%>" size=3 style="!border: 1px; background-color: #FFFFFF; border-color: #000000; border-style: solid; font-size:9px">&nbsp;
								<td>&nbsp;</td>
								<td>
								<input type="text" name="FM_mbelob_tot_<%=x%>" id="FM_mbelob_tot_<%=x%>" value="<%=SQLBlessDot(formatnumber(mbelobtot, 2))%>" size=8 style="!border: 1px; background-color: #FFFFFF; border-color: #000000; border-style: solid; font-size:9px">
								</td>
								<td style="border-right:1px #003399 solid;">&nbsp;</td>
							</tr>
					
					<%
					faktot = 0
					mbelobtot = 0
	'********* Aktiviteten ***********%>	
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="72" alt="" border="0"></td>
		<td valign="top"><input type="hidden" name="FM_aktid_<%=x%>" value="<%=thisaktid(x)%>">
		<input type="hidden" name="FM_hidden_timepristhis_<%=x%>" value="<%=thisTimePris(x)%>">
		<%
		if len(thisAkttimer(x)) <> 0 then
		hiddentimer = thisAkttimer(x)
		else
		hiddentimer = 0
		end if
		%>
		<input type="hidden" name="FM_hidden_timerthis_<%=x%>" id="FM_hidden_timerthis_<%=x%>" value="<%=hiddentimer%>">
		<input type="checkbox" name="FM_show_akt_<%=x%>" id="FM_show_akt_<%=x%>" value="1" checked>&nbsp;
		<input type="text" name="FM_timerthis_<%=x%>" id="FM_timerthis_<%=x%>" value="<%
		if len(thisAkttimer(x)) <> 0 then
		Response.write thisAkttimer(x)
		end if
		%>" size="4" onkeyup="tjektimer(<%=x%>), setTimerTot(<%=x%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
		<td valign="top">&nbsp;&nbsp;<textarea cols="54" rows="3" name="FM_aktbesk_<%=x%>" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><%=thisaktnavn(x)%></textarea>&nbsp;&nbsp;</td>
		<td valign="top"><input type="text" name="FM_enhedspris_<%=x%>" id="FM_enhedspris_<%=x%>" onKeyup="tjektimer(<%=x%>), enhedspris(<%=x%>), setTimerTot(<%=x%>)"; value="<%=SQLBlessDOT(formatnumber(thisTimePris(x), 2))%>" size="7" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;</td>
		<%
		thisbel = 0
		'if len(thisAktBeloeb(x)) <> 0 then
		'thiskomma = instr(thisAktBeloeb(x), ",")
			'if thiskomma <> 0 then
			'left_thisbel = left(thisAktBeloeb(x), thiskomma + 2)
			'thisbel = left_thisbel
			'else
			thisbel = SQLBlessDot(formatnumber(thisAktBeloeb(x), 2))
			'end if
		'end if%>
		<td valign="top"><input type="text" name="FM_beloebthis_<%=x%>" value="<%=thisbel%>" onkeyup="tjekBeloeb(<%=x%>), setBeloebTot(<%=x%>)"; size="8" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="72" alt="" border="0"></td>
	</tr>
	<%
	intTimer = intTimer + thisAkttimer(x)
	totalbelob = cdbl(totalbelob) + cdbl(thisAktBeloeb(x))
	timeprisThisjob = thisTimePris(x)
	%>
	<input type="hidden" name="antal_n_<%=x%>" id="antal_n_<%=x%>" value="<%=n%>">
	<%
	next
	
	'*** antal felter ***
	antalxx = x
	antalnn = n%>
	<input type="hidden" name="antal_x" value="<%=antalxx%>">
	<input type="hidden" name="antal_n" id="antal_n" value="<%=antalnn%>">
	
	<tr>
	<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	<td colspan="4"><%
	'*********************************************************
	'Total timer og beløb
	'*********************************************************
	if intTimer <> 0 then
	intTimer = intTimer 
	else
	intTimer = 0
	end if
	
	%>&nbsp;&nbsp;&nbsp;<input type="text" name="FM_Timer" value="<%=intTimer%>" size="4" onkeyup="setTimerTot(<%=x-1%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
	&nbsp;&nbsp;Total antal.&nbsp;</td>
	<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		<td colspan="4" align="right" style="padding-right:16;"><b>Nettobeløb (subtotal)</b>&nbsp;&nbsp;
		<!-- strBeloeb -->
		<%thistotbel = 0
		if len(totalbelob) <> 0 then
		thiskomma = instr(totalbelob, ",")
			if thiskomma <> 0 then
			left_thisbel = left(totalbelob, thiskomma + 2)
			thistotbel = left_thisbel
			else
			thistotbel = totalbelob
			end if
		end if
		%>
		<input type="text" name="FM_beloeb" value="<%=thistotbel%>" size="8" onkeyup="setBeloebTot(<%=x%>);" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
	</tr>
	<%
	'***********************************************************
	%>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="220" alt="" border="0"></td>
		<td valign="top" colspan="4" style="padding-left:50;"><br><br>Betalings betingelser / Kommentar: 
		<br><textarea cols="70" rows="5" name="FM_komm" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><%=strKom%></textarea><br>
		Vis kommentar før <input type="radio" name="FM_viskomfor" value="0"> 
		eller efter <input type="radio" name="FM_viskomfor" value="1" checked> udspecificering.<br><br>
		<font size=1>*) Anbefalet timepris på en aktivitet er beregnet udfra:<br>
		A) <u>Angivet budget ell. fastpris på job / antal fakturerbare timer på job</u>.<br>
		B) <u>Gennemsnitlig timepris</u> på de allerede indtastede timer på aktiviteten. Hvis der ikke findes timer bruges A.<br><br></font></td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="220" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top" colspan="6"><img src="../ill/tabel_top.gif" width="700" height="1" alt="" border="0"></td>
	</tr>
	</table>
	<br><br>
	
	<!--
	<table cellpadding=0 border=0 cellspacing=0>
	<tr>
		<td valign="top" colspan="3"><img src="../ill/tabel_top.gif" width="700" height="1" alt="" border="0"></td>
	</tr>
	<tr bgcolor="whitesmoke">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="60" alt="" border="0"></td>
		<td valign="top" width="650">&nbsp;Luk dette job, efter denne faktura er oprettet?:
		&nbsp;<b>Ja:</b><input type="checkbox" name="FM_lukjob" value="1"><br><br>
		&nbsp;Faktura er betalt?&nbsp;<b>Ja:</b><input type="checkbox" name="FM_betalt" value="1" <=betaltch%>>&nbsp;&nbsp;
	
	<
	if func = "red" then
	strTdato = strB_dato
	else
	strTdato = formatdatetime(date, 0)
	end if 	
	%>
	<!--include file="inc/dato2.asp"-->
		<!--<select name="FM_b_start_dag">
		<
		call dag()
		%>
		<select name="FM_b_start_mrd">
		<
		call mrd()
		%>
		<select name="FM_b_start_aar">
		<
		call Aar()
		%>
		</td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="60" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top" colspan="3"><img src="../ill/tabel_top.gif" width="700" height="1" alt="" border="0"></td>
	</tr>
	</table>
	-->
	
	<table cellpadding=0 border=0 cellspacing=0>
	<tr>
		<td valign="top">&nbsp;</td>
		<td><br><br>
		<b>Der kan ikke foretages ændringer til denne faktura.</b>
		<!--<input type="image" src="../ill/opretpil_fak.gif">-->
		<td valign="top" align=right>&nbsp;</td>
	</tr>
	</form>
	</table>
	<br><br>
<br>
<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
<br>
<br>
</div>
<%end select%>
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
