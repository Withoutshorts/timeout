<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"--> 
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/functions_inc.asp"-->  
<%


a = 0
dim aval
redim aval(a)

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
	
	
	if len(request("aftaleid")) <> 0 then
	aftid = request("aftaleid")
	else
	aftid = 0
	end if
	
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
	   <td align=right style="padding:5; border-right:1px darkred solid; border-bottom:1px darkred solid;"><a href="fak.asp?menu=stat_fak&aftaleid=<%=aftid%>&FM_aftnr=<%=sogaftnr%>&thiskid=<%=thiskid%>&func=sletok&id=<%=id%>&FM_job=<%=Request("FM_job")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>">Ja&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"></a></td>
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
	
	
	oConn.execute("DELETE FROM fakturaer WHERE Fid = "& id &"")
	
	if aftid = 0 then '*** Hvis fakturaen tilhører et job **
	oConn.execute("DELETE FROM faktura_det WHERE fakid = "& id &"")
	oConn.execute("DELETE FROM fak_med_spec WHERE fakid = "& id &"")
	end if
	
	if request("FM_usedatointerval") = "1" then
	usedval = 1
	else
	usedval = 0
	end if
	
	if aftid = 0 then
	Response.redirect "fak_osigt.asp?menu=stat_fak&FM_job=" & Request("FM_job") &"&FM_usedatointerval="&usedval&"&FM_start_dag="&request("FM_start_dag")&"&FM_start_mrd="&request("FM_start_mrd")&"&FM_start_aar="&request("FM_start_aar")&"&FM_slut_dag="&request("FM_slut_dag")&"&FM_slut_mrd="&request("FM_slut_mrd")&"&FM_slut_aar="&request("FM_slut_aar")&"&FM_fakint="&request("FM_fakint")
	else
	Response.redirect "fak_serviceaft_osigt.asp?menu=stat_fak&aftaleid="&aftid&"&FM_aftnr="&sogaftnr&"&id="&thiskid
	end if
	
	
	case "fortryd"
	
	oprid = 0
	
	strSQL = "SELECT oprid FROM fakturaer WHERE fid =" & id
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	oprid = oRec("oprid")
	oRec.movenext
	wend
	oRec.close
	
	if oprid <> 0 then
	strSQL = "DELETE FROM posteringer WHERE oprid ="& oprid
	'Response.write strSQL
	oConn.execute(strSQL)
	end if
	
	
	strSQL = "UPDATE fakturaer SET betalt = 0 WHERE fid =" & id
	oConn.execute(strSQL)
	
	if request("FM_usedatointerval") = "1" then
	usedval = 1
	else
	usedval = 0
	end if
	
	'Response.redirect "stat_fak.asp?menu=stat_fak&FM_kunde=0&shokselector=1"
	Response.redirect "fak_osigt.asp?menu=stat_fak&FM_job="&Request("FM_job")&"&FM_usedatointerval="&usedval&"&FM_start_dag="&request("FM_start_dag")&"&FM_start_mrd="&request("FM_start_mrd")&"&FM_start_aar="&request("FM_start_aar")&"&FM_slut_dag="&request("FM_slut_dag")&"&FM_slut_mrd="&request("FM_slut_mrd")&"&FM_slut_aar="&request("FM_slut_aar")&"&FM_fakint="&request("FM_fakint")
		
	
	case "dbred", "dbopr" 
	
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
	
	if len(request("FM_faktype")) <> 0 then
	intFaktype = request("FM_faktype")
	else
	intFaktype = 0
	end if
	
	
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
								
								
									'** Indsætter cookie **
									Response.Cookies("clastfaknr") = "0"
									Response.Cookies("clastfaknr").Expires = date + 1
									
									strSQL2 = "SELECT faknr FROM fakturaer WHERE faknr = '"& intFaknum &"' AND Fid <> "& id &""
									oRec2.open strSQL2, oConn, 3
									fakfindeallerede = 0
									if not oRec2.EOF then 'faknr findes allerede
									fakfindeallerede = 1
									oRec2.close
									
											if fakfindeallerede = 1 then%>
											<!--#include file="../inc/regular/header_inc.asp"-->
											<!--#include file="../inc/regular/topmenu_inc.asp"-->
											<%
											errortype = 29
											call showError(errortype)
											
											end if
									
									
									else
							
							
							
							'*** Hvis alle required er udfyldt ***
							
							'Response.write "1<br>"
							'Response.flush
							
									fakDato = Request("FM_start_aar") & "/" & Request("FM_start_mrd") & "/" & Request("FM_start_dag") '& time 
									showfakDato = Request("FM_start_dag") & "/" & Request("FM_start_mrd") & "/" & Request("FM_start_aar")
									strEditor = session("user")
									strDato = session("dato")
									strjobid = Request("jobid")
									intFakbetalt = request("FM_betalt")
									intfakadr = request("FM_Kid")
									
									aftid = Request("aftid")
									jobid = strjobid
									
									kundeid = request("FM_kundeid")
									
									if request("FM_att_vis") = "on" then
									strAtt = request("FM_att")
									else
									strAtt = "Ingen"
									end if
									
									if intFakbetalt <> 1 then
									intFakbetalt = 0
									else
									intFakbetalt = intFakbetalt 
									end if
									
									if len(request("FM_varenr")) <> 0 then
									strVarenr = request("FM_varenr")
									else
									strVarenr = 0
									end if
									
									intEnhedsang = request("FM_enheds_ang")
																		
									'** Database format dato **
									'dtb_dato = Request("FM_b_start_aar") & "/" & Request("FM_b_start_mrd") & "/" & Request("FM_b_start_dag")
									dtb_dato = year(now) & "/"& month(now) & "/" & day(now)
									
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
									
									useDebKre = request("FM_debkre")
									
									
									call beregnmoms(useDebKre, intBeloeb, varKonto, varModkonto)
									
									
									
									'Response.write "2<br>"
									'Response.flush
							
									
							
							if request("FM_land_vis") = "on" then
							intModland = 1
							else
							intModland = 0
							end if
							
							if request("FM_att_vis") = "on" then
							intModAtt = 1
							else
							intModAtt = 0
							end if
							
							if request("FM_tlf_vis") = "on" then
							intModTlf = 1
							else
							intModTlf = 0
							end if
							
							if request("FM_cvr_vis") = "on" then
							intModCvr = 1
							else
							intModCvr = 0
							end if
							
							if request("FM_yswift_vis") = "on" then
							intAfsSwift = 1
							else
							intAfsSwift = 0
							end if
							
							
							if request("FM_yiban_vis") = "on" then
							intAfsIban = 1
							else
							intAfsIban = 0
							end if
							
							if request("FM_ycvr_vis") = "on" then
							intAfsICVR = 1
							else
							intAfsICVR = 0
							end if
								
							if request("FM_yemail_vis") = "on" then
							intAfsEmail = 1
							else
							intAfsEmail = 0
							end if
							
							if request("FM_ytlf_vis") = "on" then
							intAfsTlf = 1
							else
							intAfsTlf = 0
							end if		
							
							
							'*** txt på postering ***
							if aftid <> 0 then
							posteringsTxt = left(request("FM_komm"), 20)
							else
							posteringsTxt = left(request("FM_jobnavn"), 20)
							end if
							
							
							'*** Jobbesk *****
							if len(request("FM_visjobbesk")) <> 0 then
							jobBesk = SQLBlessK(request("FM_jobbesk"))
							else
							jobBesk = ""
							end if
							
							
									
							'************************************************************************************
							'*** Opdaterer / Redigerer faktura 													*
							'************************************************************************************
							if func = "dbred" then
							
											strSQLupd = "UPDATE fakturaer SET"_
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
											&" faktype = "& intFaktype &", "_
											&" vismodland = "& intModland &", "_
											&" vismodatt = "& intModAtt &", "_
											&" vismodtlf = "& intModTlf &", "_
											&" vismodcvr = "& intModCvr &", "_
											&" visafstlf = "& intAfsTlf &", "_
											&" visafsemail = "& intAfsEmail &", "_
											&" visafsswift = "& intAfsSwift &", "_
											&" visafsiban = "& intAfsIban &", "_
											&" visafscvr = "& intAfsICVR &", "_
											&" moms = "& showmoms &", enhedsang = "& intEnhedsang &", varenr = '"& strVarenr &"', jobbesk = '"& jobBesk &"'"_
											&" WHERE Fid = "& id 
											
											'Response.write (strSQLupd)
											oConn.execute(strSQLupd)
											
											
							thisfakid = id
							end if
							'***
											
											
											
							'************************************************************************************
							'*** Opretter faktura *********														*
							'************************************************************************************
							if func = "dbopr" then
												
												if intFaktype <> 0 then '*** kreditnota eller rykker
												parentfak = request("id")
												else
												parentfak = 0
												end if
													
													'***tidspunkt = 23:23:59 pga luk for indtastning muligheden på timeregsiden
													
													strSQL = ("INSERT INTO fakturaer"_
													&" (faknr, fakdato, jobid, timer, beloeb, kommentar, dato, editor, tidspunkt, "_
													&" betalt, b_dato, fakadr, att, faktype, konto, modkonto, parentfak, "_
													&" vismodland, vismodatt, vismodtlf, vismodcvr, visafstlf, visafsemail, "_
													&" visafsswift, visafsiban, visafscvr, moms, enhedsang, aftaleid, varenr, jobbesk) VALUES ("_
													&" '"& intFaknum &"',"_
													&" '"& fakDato &"',"_
													&" "& strjobid &","_
													&" "& intTimer &","_
													&" "& intBeloeb &","_
													&" '"& strKomm &"',"_
													&" '"& strDato &"',"_
													&" '"& strEditor &"', '23:59:59', "_
													&" "& intFakbetalt &", '"& dtb_dato &"', "& intfakadr &", "_
													&" '"& strAtt &"', "& intFaktype &", '"& varKonto &"', "_
													&" '"& varModkonto &"', "& parentfak &", "_
													&" "& intModland &", "& intModAtt &", "& intModTlf &", "_
													&" "& intModCvr &", "& intAfsTlf &", "& intAfsEmail &", "& intAfsSwift &", "_
													&" "& intAfsIban &", "& intAfsICVR &", "& showmoms &", "& intEnhedsang &", "& aftid &", '"& strVarenr &"', '"& jobBesk &"')")
													
													'Response.write strSQL
													oConn.execute(strSQL)
													
										
													
													''** Henter fak id ***
													strSQL3 = "SELECT Fid FROM fakturaer"
													oRec3.open strSQL3, oConn, 3
													oRec3.movelast
													if not oRec3.EOF then
													thisfakid = cint(oRec3("Fid")) 
													end if 
													oRec3.close	
									
									
									end if	'** Opret
									
									
									'Response.write "3<br>"
									'Response.flush
									
										
											'** Posteringer ***		
											'** Opretter tilhørende posteringer hvis faktua godkendes som endelig **		
											if intFakbetalt <> 0 then 
														
														call opretpos() '*** /inc/functions_inc.asp
														
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
																	 	
																	end if 'Tidligere
														end if  'Rykker
											end if '** kladde / endelig
											
							
							'***
							
														
							'Response.write "4<br>"
							'Response.flush
							
							
							
							
							if jobid <> 0 then '** Bruges kun hvis der oprettes faktura på job ***					
							
							
							'***********************************************************************************
							'*** Hvis faktura allerede er oprrettet en gang i denne session ***
							'***********************************************************************************
								
								 
								'************************************************************** 
								'*** Indsætter akt i fak_det 
								'**************************************************************
								if func = "dbred" then
								oConn.execute("DELETE FROM faktura_det WHERE fakid = "& id &"")
								thisfakid = id
								end if
								
								'***** n starter på 1 **************
								'***** Aktivitets udspecificering ************
								for intcounter = 0 to antalAkt - 1
								
								thisAktId = request("aktId_n_"&intcounter&"")
									
									'** antal aktiviter udskrevet pga. forskellige timepriser 
									'antalsumaktprakt = request("antal_subtotal_akt_"&intcounter&"") 
									antalsumaktprakt = request("antal_n_"&intcounter&"") 
									
									for intcounter3 = -(antalsumaktprakt) to antalsumaktprakt
									
												
												'*** Enhedsprisen på denn akt.
												if len(request("FM_enhedspris_"& intcounter &"_"&intcounter3&"")) <> 0 then
												enhpris = SQLBless2(request("FM_enhedspris_"& intcounter &"_"&intcounter3&""))
												else
												enhpris = 0
												end if
									
										'*** Vis sum aktivitet på print (og DB)
										if len(request("FM_show_akt_"& intcounter &"_"&intcounter3&"")) = 1 then
												
												timerThis = request("FM_timerthis_"& intcounter &"_"&intcounter3&"")
												if len(timerThis) <> 0 then
												timerThis = SQLBless2(timerThis)
												else 
												timerThis = 0
												end if
												
												if len(request("FM_beloebthis_"& intcounter &"_"&intcounter3&"")) <> 0 then
												beloebThis = SQLBless2(request("FM_beloebthis_"& intcounter &"_"&intcounter3&""))
												else
												beloebThis = 0
												end if
												
												
												strSQL_sumakt = ("INSERT INTO faktura_det "_
												&" (antal, beskrivelse, aktpris, fakid, enhedspris, aktid, showonfak) "_
												&" VALUES ("& timerThis &", "_
												&"'" & SQLBlessK(request("FM_aktbesk_"& intcounter &"_"&intcounter3&"")) &"', "_
												&""  & beloebThis &", "_
												&""  & thisfakid &", "& enhpris &", "& thisAktId &", 1)")
												
												'Response.write strSQL_sumakt  & "<br>"
												oConn.execute(strSQL_sumakt)
												
										end if 'show sumaktivitet
													
										
										'**************************************************			
										'***** Medarbejder udspecificering i db ***********
										'**************************************************
										
										
										'Response.write "5<br>"
										'Response.flush
											
										antalmedspec2 = request("antal_n_"&intcounter&"") 'medarbejdere
										for intcounter2 = 0 to antalmedspec2 - 1
												
												'*** Passer timeprisen på denne akt og medarbejder **
												thismedarbtpris = request("FM_mtimepris_"& intcounter2&"_"&intcounter&"")
												thismedarbtpris = formatnumber(thismedarbtpris, 2)
												thismedarbtpris = replace(thismedarbtpris, ".", "")
												thismedarbtpris = SQLBless2(thismedarbtpris) 
												
												thisenhpris = request("FM_enhedspris_"& intcounter &"_"&intcounter3&"")
												thisenhpris = formatnumber(thisenhpris, 2)
												thisenhpris = replace(thisenhpris, ".", "")
												enhpris = SQLBless2(thisenhpris)
												
												
												'** registrerer timer (fak og vent) på alle medarbejdere 
												if len(request("FM_mid_"&intcounter2&"_"&intcounter&"")) <> 0 then
												thisMid = request("FM_mid_"&intcounter2&"_"&intcounter&"")
												else
												thisMid = 0
												end if 
												
												'if thisMid <> 0 then
												'Response.write "<hr>" & request("FM_aktbesk_"& intcounter &"_"&intcounter3&"") & "<br>"
												'Response.write "if" & thismedarbtpris &" = "& enhpris  &" then<br>"
												'Response.write "timer:" & request("FM_m_fak_"&intcounter2&"_"&intcounter&"")
												'Response.write "<br>true sum aktr:" & request("FM_show_akt_"&intcounter&"_"&intcounter3&"")
												'end if
												if cdbl(thismedarbtpris) = cdbl(enhpris) AND thisMid <> 0 then
												'Response.write "<br>ok!<br>"
											
													
														'* Nulstiller evt. tidligere indtastninger på denne faktura ***
														oConn.execute("DELETE FROM fak_med_spec WHERE fakid = "& thisfakid &" AND aktid = "&thisAktId&" AND mid = "&thisMid&"")
													 	
														'* Beløb
														if len(request("FM_mbelob_"& intcounter2 &"_"&intcounter&"")) <> 0 then
														useBeloeb = SQLBless2(request("FM_mbelob_"& intcounter2 &"_"&intcounter&""))
														else
														useBeloeb = 0
														end if
														
														'* Venter
														if len(request("FM_m_vent_"&intcounter2&"_"&intcounter&"")) <> 0 then
														useVenter = SQLBless2(request("FM_m_vent_"&intcounter2&"_"&intcounter&""))
														else
														useVenter = 0
														end if
														
														
														'** Nulstiler altid vente timer inden der tildeles nye vente timer for denne medarbejder på denne aktivitet 
														'** (Uanset hvilken fak akt. hører til)
														
														oConn.execute("UPDATE fak_med_spec SET venter = 0 WHERE aktid = "&thisAktId&" AND mid = "&thisMid&"")
													 	
														
														if len(request("FM_m_fak_"&intcounter2&"_"&intcounter&"")) <> 0 then
															'*** Hvis show sum-akt ikke er true skal faktimer altid være = 0
															'if request("FM_show_akt_"&intcounter&"_"&intcounter3&"") = "1" then 
															useFak = SQLBless2(request("FM_m_fak_"&intcounter2&"_"&intcounter&""))
															'else
															'useFak = 0
															'end if
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
														
														'*** Indsætter i db ****
														strSQL = ("INSERT INTO fak_med_spec (fakid, aktid, mid, fak, venter, tekst, enhedspris, beloeb, showonfak) "_
														&" VALUES ("& thisfakid &", "&thisAktId&", "_
														&"" &thisMid&", "_
														&"" &useFak&", "_
														&"" &useVenter&", "_
														&"'"&SQLBlessK(request("FM_m_tekst_"& intcounter2&"_"&intcounter&""))&"', "_
														&"" &usemedTpris&", "_
														&"" &useBeloeb&", "& showonfak &")")
														oConn.execute(strSQL)
														
														'Response.write strSQL & "<br><br>"
														'Response.flush
													
													'else
													'Response.write "Stemmer ikke <br>"	
													
													end if 'thismedarbtpris
													
											next 'Medarbejdere
										next 'Sumaktiviteter
								next 'Antal aktiviteter
							
							
							'Response.write "6<br>"
							'Response.flush
							
							else	
							
							'select case intFakbetalt
							'case 1
							'erfak = 1 'godkendt
							'case 0
							'erfak = 2 'kladde
							'end select
							'*** Opdaterer seraft, så aftalen ER faktureret ***
							'oConn.execute("UPDATE serviceaft SET erfaktureret ="& erfak &" WHERE id = "& aftid &"")
									
							end if '*** Kun job (opr. aktiviterer og fak_med_spec)	
							
							'**********************************************************************************************
							
							'**** Opdaterer faste betalingsbetingelser på kunde ***
							if len(request("FM_gembetbet")) <> 0 then
							strSQLupd = "UPDATE kunder SET betbet = '"& strKomm &"' WHERE kid = " & kundeid
							oConn.execute(strSQLupd)
							end if
							
							'** Gemmer på alle kunder **
							if len(request("FM_gembetbetalle")) <> 0 then
							strSQLupd = "UPDATE kunder SET betbet = '"& strKomm &"' WHERE kid <> 0"
							oConn.execute(strSQLupd)
							end if
							
							
												
												
											'*** Viser den oprettede faktura til print *****%>
											<!-------------------------------Sideindhold------------------------------------->
											<%
											
											Response.redirect "fak_godkendt.asp?jobid="&jobid&"&aftid="&aftid&"&id="&thisfakid&"&menu=stat_fak&FM_job="&Request("FM_job")&"&FM_usedatointerval="&request("FM_usedatointerval")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")&"&FM_fakint_ival="&request("FM_fakint_ival")&"&FM_aftnr="&sogaftnr&"&thiskid="&thiskid
											
			
											
									end if
							end if
					end if
			end if
	
	
	
	
	
	
	
	
	
	
	
	case else
	'**************************** Opret/rediger Faktura Data ************************************
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<SCRIPT language=javascript src="inc/fak_func.js"></script>
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="inc/fak_inc_subs.asp" -->
	<!-------------------------------Sideindhold------------------------------------->
	<%
	'public strDag
	'public strMrd
	'public strMrdNavn
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
	'*** Fakturér job ***
	
	if len(Request("jobnr")) <> 0 then
	jobnr = Request("jobnr")
	else
	jobnr = 0
	end if
	
	if len(request("jobid")) <> 0 then
	jobid = request("jobid")
	else
	jobid = 0
	end if
	
	timeforbrug = 0
	faktureretBeloeb = 0
	
	
	
	'*** Fakturér Aft. ***
	if len(request("aftaleid")) <> 0 then 
	aftid = request("aftaleid")
	else
	aftid = 0
	end if
	
	
	
	pleft = 200
	ptop = 50
	
	Redim thisaktid(0)
	Redim thisaktnavn(0)
	Redim thisAktTimer(0)
	Redim thisAktBeloeb(0)
	Redim usefastpris(0)
	Redim thisTimePris(0)
	
	
	if func = "red" then
	'*********************************************************************
	' Rediger faktura 
	'*********************************************************************
	
	
						strSQL = "SELECT Fid, faknr, fakdato, jobid, timer, beloeb, kommentar, dato, editor, betalt, b_dato, fakadr, att, faktype, konto, modkonto, varenr, enhedsang, jobbesk FROM fakturaer WHERE Fid = "& id 
						
						oRec.open strSQL, oConn, 3
						
						if not oRec.EOF then
						strEditor = oRec("editor")
						strFaknr = oRec("faknr")
						strTdato = formatdatetime(oRec("fakdato"), 2)
								'*** Hvis faktype er <> for faktura skiftes dbfunc til dbopr ***
								if request("faktype") <> 0 AND request("rykkreopr") = "j" then
									dbfunc = "dbopr"
									strNote = "<font color=darkred>Husk at ændre nr. i forhold til det oprindelige faktura nummer!</font>"
									strNote1 = "Nedenstående udspecificering er en kopi af den oprindelige faktura."
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
						
						strVarenr = oRec("varenr")
						intEnhedsang = oRec("enhedsang")
						
						strJobBesk = oRec("jobbesk")
						
						strKom = oRec("kommentar")
						
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
						
						
						if jobid <> 0 then
							
							call faktureredetimerogbelob()
							
						
							strSQL = "SELECT jobTpris, budgettimer, fastpris, jobknr, jobnr, jobnavn, ikkebudgettimer, jobans1, jobans2 FROM job WHERE id = " & jobid
							oRec.open strSQL, oConn, 3
							if not oRec.EOF then
								
								
								
								intIkkeBtimer = oRec("ikkebudgettimer")
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
								
								jobans1 = oRec("jobans1")
								jobans2 = oRec("jobans2")
								
								
								
							end if
							oRec.close
							
						else
						
							intjobknr = intfakadr
							
							intEnheder = strTimer
							intPris = strBeloeb
							strBesk = strKom
						
						end if
		'*******************************************************************************************
	
	
		call top
		'** Mulighed for at slette og se sidst redigeret dato ****
		%><img src="../ill/blank.gif" width="10" height="318" alt="" border="0">
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
	'************************************************************************'
	' Opret faktura 
	'************************************************************************
	
	call top
	%>
	</div>
	<%
		
			'****************************************************************
			'*** Finder faknr på ny faktura ***
			'****************************************************************
			if request.Cookies("clastfaknr") = "0" OR len(request.Cookies("clastfaknr")) = 0 then
				lastFaknr = 0 	
				
				strSQL = "SELECT Fid, faknr FROM fakturaer WHERE faktype = 0 ORDER BY fid DESC" 
				oRec.open strSQL, oConn, 3
				
				while not oRec.EOF 
				
				call erDetInt(oRec("faknr"))
					if isInt = 0 then
						if cdbl(oRec("faknr")) >= lastFaknr then
						lastFaknr = cdbl(oRec("faknr")) + 1
						else
						lastFaknr = lastFaknr
						end if
					end if
				isInt = 0
				oRec.movenext
				wend
				oRec.close
				
				
				
			
				'*** Er faknr allerede under opr? ****
				strSQL = "SELECT id, faknr FROM fak_opr_faknr"
				oRec.open strSQL, oConn, 3
				if not oRec.EOF then
			
					if cdbl(lastFaknr) <= cdbl(oRec("faknr")) then
						lastFaknr = oRec("faknr") + 1
					end if
				
				end if
				oRec.close
				
				
				
				if lastFaknr <> 0 then 
				'*** Opdaterer lastfaknr ***
				strSQLF = "DELETE FROM fak_opr_faknr WHERE faknr <> 0"
				oConn.execute(strSQLF)
				strSQLF = "INSERT INTO fak_opr_faknr (faknr) VALUES ('"& lastFaknr &"')"
				oConn.execute(strSQLF)
				end if
			
				
				'** Indsætter cookie **
				Response.Cookies("clastfaknr") = lastFaknr
				Response.Cookies("clastfaknr").Expires = date + 65
				
			else
				'Response.Cookies("clastfaknr") = ""
				lastFaknr = request.Cookies("clastfaknr")
			end if
			
			
			'*******************************************************************************
		
		
		'*** Find kunde vedr opret faktura på JOB ***
		if jobid <> 0 then
		
		call faktureredetimerogbelob()
		
		strSQL = "SELECT jobTpris, budgettimer, fastpris, jobknr, jobnr, jobnavn, ikkebudgettimer, jobans1, jobans2, kundekpers, beskrivelse FROM job WHERE id = " & jobid
		oRec.open strSQL, oConn, 3
		
		if not oRec.EOF then
			intIkkeBtimer = oRec("ikkebudgettimer")
			intBudgettimer = oRec("budgettimer")
			intJobTpris = oRec("jobTpris")
			fastpris = oRec("fastpris")
			intjobknr = oRec("jobknr")
			intjobnr = oRec("jobnr")
			strJobnavn = oRec("jobnavn")
			
			jobans1 = oRec("jobans1")
			jobans2 = oRec("jobans2")
			
			intKundekpers = oRec("kundekpers")
			strJobBesk = oRec("beskrivelse")
			
			
		end if
		oRec.close
		
		else 
		'*** Find kunde vedr opret faktura på Aftale ***
		
		strSQL = "SELECT kundeid, besk, enheder, pris, varenr, navn, perafg, "_
		&" advitype, advihvor, stdato, sldato FROM serviceaft WHERE id = "& aftid
		oRec.open strSQL, oConn, 3 
		while not oRec.EOF 
		
		intjobknr = oRec("kundeid")
		strBesk = oRec("besk")
		intEnheder = oRec("enheder")
		intPris = oRec("pris")
		strVarenr = oRec("varenr")
		strNavn = oRec("navn")
		intPerafg = oRec("perafg")
		intAdvitype = oRec("advitype")
		intAdvihvor = oRec("advihvor")
		startdato = oRec("stdato") 
		slutdato = oRec("sldato") 
		
		oRec.movenext
		wend
		oRec.close 
		
		strBesk = strNavn&":&nbsp;"&strBesk
		
		end if
		
		strAtt = intKundekpers
		
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
	
	
	
	<div id=beregntpris style="position:absolute; left:200px; top:366px; height:150px;">
	<table cellspacing=0 cellpadding=0 border=0><form name=beregn>
	<tr bgcolor="#5582D2">
		<td width="8" height=20 rowspan="2" valign=top style="border-left:1px #003399 solid; border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=3 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="134" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top style="border-right:1px #003399 solid; border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class=alt colspan=3><b>Beregn en timepris:</b></td>
	</tr>
	<tr bgcolor="#eff3ff">
		<td align=right style="border-left:1px #003399 solid; border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td style="border-bottom:1px #003399 solid; padding-left:5px; padding-top:2px;">Beløb: <input type="text" name="beregn_belob" id="beregn_belob" value="0" size="5"> <b>/</b> </td>
		<td style="border-bottom:1px #003399 solid; padding-left:5px; padding-top:2px;">Timer: <input type="text" name="beregn_timer" id="beregn_timer" value="0" size="5"></td>
		<td width=160 style="border-bottom:1px #003399 solid; border-bottom:1px #003399 solid; padding-left:5px; padding-right:5px; padding-bottom:3px">
		<input type="button" name="beregn" id="beregn" value=" = " onClick="beregntimepris()" style="font-size:9px;"> <input type="text" name="beregn_tp" id="beregntp" value="0" style="width:99;"></td>
		<td align=right style="border-right:1px #003399 solid; border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr></form></table>								
	</div>
	
	
	
	
	
	<%
	if func <> "red" then
	id = 0 'Sættes så dato2 virker!
	else
	id = id
	end if
	%>	
	<!--#include file="inc/dato2.asp"-->
	<form name=main id=main action="fak.asp?menu=stat_fak&func=<%=dbfunc%>" method="post">
			<%
			'*** Bruger altid datointerval ****
			usedt_ival = 1
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
	
	<input type="hidden" name="aftid" value="<%=aftid%>">
	<input type="hidden" name="FM_aftnr" value="<%=sogaftnr%>">
	<input type="hidden" name="thiskid" value="<%=thiskid%>">
	
	
	
	
	
	
	
	
	<div id="modtager" style="position:absolute; left:<%=pleft%>; top:<%=ptop+50%>; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="370" bgcolor="#EFF3FF">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td valign="top"><img src="../ill/tabel_top.gif" width="354" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class=alt><b>Modtager:</b></td>
	</tr>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td valign=top>
		<br>
		<%
		strSQL = "SELECT Kid, kkundenavn, kkundenr, adresse, postnr, city, land, telefon, cvr, betbet FROM kunder WHERE Kid =" & intjobknr  		
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
			
			if func <> "red" then
			'*** Betalings betingelser ****
			strKom = oRec("betbet")
			end if
			
		end if
		
		oRec.close
		%>
		<input type="hidden" name="FM_Kid" value="<%=intKid%>">
		<table cellpadding=0 cellspacing=0 border=0 width=100%>
		<tr>
			<td colspan=2><b>Adresse:</b></td>
		</tr>
		<tr>
			<td colspan=2 class=lille>Kundenr: <%=intKnr%></td>
		</tr>
		<tr>
			<td colspan=2><%=strKnavn%></td>
		</tr>
		<tr>	
			<td colspan=2><%=strKadr%></td>
		</tr>
		<tr>	
			<td colspan=2><%=strKpostnr%>&nbsp;&nbsp;<%=strBy%></td>
		</tr>
		<tr>	
			<td width=200><%=strLand%></td><td ><input type="checkbox" name="FM_land_vis" value="on">Vis land på fak.
		</td>
		<tr>	
			<td>Att:&nbsp;<select name="FM_att" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid; width:200;">
				<%
				strSQLkpersCount = "SELECT id, navn FROM kontaktpers WHERE kundeid = "& intKid &" ORDER BY id"
				oRec2.open strSQLkpersCount, oConn, 3
					
					while not oRec2.EOF
						if cint(strAtt) = cint(oRec2("id")) then
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
				</select></td>
				<td ><input type="checkbox" name="FM_att_vis" value="on" checked>Vis att. på fak.</td>
		</tr>
		<tr><td><br>Tlf: <%=intTlf%></td><td><br><input type="checkbox" name="FM_tlf_vis" value="on">Vis tlf. på fak.<br></td></tr>
		<tr><td><b>Cvr/SE nr:</b> <%=intCVR%></td><td ><input type="checkbox" name="FM_cvr_vis" value="on">Vis cvr på fak.</td></tr>
		<tr><td colspan=2><b>Kundekonto:</b>
		<select name="FM_kundekonto" style="width:200;">
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
		</select>&nbsp;&nbsp;<select name="FM_debkre">
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
		</select></td></tr>
		</table>
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
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td valign="top"><img src="../ill/tabel_top.gif" width="284" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class=alt><b>Afsender:</b></td>
	</tr>
	<tr><td width="30" style="border-left:1px #003399 solid;">&nbsp;</td>
	<td>
		<%
		afskid = 0
		
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
		
		
		%>
		
		<table cellpadding=0 cellspacing=0 border=0 width=100%>
		<tr>
			<td colspan=2><br><b>Adresse:</b></td>
		</tr>
		
		<tr>
			<td colspan=2><%=yourNavn%></td>
		</tr>
		<tr>
			<td colspan=2><%=yourAdr%></td>
		</tr>
		<tr>
			<td colspan=2><%=yourPostnr%>&nbsp;&nbsp;<%=yourCity%></td>
		</tr>
		<tr>
			<td width=120 colspan=2><%=yourLand%></td>
		</tr>
		<tr>
			<td><br>Tlf: <%=yourTlf%></td><td><br><input type="checkbox" name="FM_ytlf_vis" value="on" CHECKED>Vis tlf. på fak.</td>
		</tr>
		<tr>
			<td>Email: <%=yourEmail%></td><td ><input type="checkbox" name="FM_yemail_vis" value="on">Vis email på fak.</td>
		</tr>
		<tr>
			<td colspan=2><br><b>Bank information:</b></td>
		</tr>
		<tr>
			<td colspan=2>Bank: <%=yourBank%></td>
		</tr>
		<tr>
			<td colspan=2>Reg. og kontonr: <%=yourRegnr%> - <%=yourKontonr%></td>
		</tr>
		<tr>
			<td>Swift: <%=yourSwift%></td><td ><input type="checkbox" name="FM_yswift_vis" value="on">&nbsp;Vis på fak.</td>
		</tr>
		<tr>
			<td>Iban: <%=yourIban%> </td><td ><input type="checkbox" name="FM_yiban_vis" value="on">&nbsp;Vis på fak.</td>
		</tr>
		<tr>
			<td><b>Cvr/SE nr:</b> <%=yourCVR%></td><td><input type="checkbox" name="FM_ycvr_vis" value="on" CHECKED>&nbsp;Vis på fak.</td>
		</tr>
		</table>
		
		</td><td width="30" style="border-right:1px #003399 solid;">&nbsp;</td>
		</tr>
		<tr><td width="30" style="border-left:1px #003399 solid;">&nbsp;</td>
		<td><b>Modkonto:</b>&nbsp;
		<select name="FM_modkonto" style="width:200px;">
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
	
	
	
	
	
	
	
	
	<div id="aktiviteter" style="position:absolute; left:<%=pleft%>; top:<%=ptop+420%>; visibility:visible; z-index:100; background-color:#d6dff5;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#ffffff" width="700">
	<tr bgcolor="#5582D2">
		<td width="8" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=4 style="border-top:1px #003399 solid;" class=alt><b>Faktura og job/aftale info.</b></td>
		<td align=right valign=top width="8"><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	
	<%
	'*** Jobinfo ***
	if jobid <> 0 then%>
	<tr bgcolor="#FFFFFF">
		<td style="border-left:1px #003399 solid;" rowspan="8">&nbsp;</td>
		<td colspan=4 height=5><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td style="border-right:1px #003399 solid;" rowspan="8">&nbsp;</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td colspan=2 valign=top style="padding:10px; border-bottom:1px silver dashed;">
		<b>(<%=jobnr%>)&nbsp;&nbsp;<%=strJobnavn%></b><br>
		<u>Forkalkuleret timer:</u><br>
		Fakturerbare: <%=intBudgettimer%> <br>
		Ikke fakturerbare: <%=intIkkeBtimer%> <br>
		<%if fastpris = 1 then
		jtype = "Fastpris"
		else
		jtype = "Løbende timer"
		end if%>
		Jobtype: <font color=red><%=jtype%> </font><br>
		Budget/Pris på job:	<%=formatcurrency(intJobTpris, 2)%><br>
		<u>Jobansvarlig(e):</u><br>
		<%
		
		strSQL2 = "SELECT mnavn, mnr, mid FROM medarbejdere WHERE  mid = "& jobans1
		
		'Response.write strSQL2
		'Response.flush
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
		jobans1txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
		end if
		oRec2.close
		
		strSQL2 = "SELECT mnavn, mnr, mid FROM medarbejdere WHERE  mid = "& jobans2
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
		jobans2txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
		end if
		oRec2.close
		%>
		
		<%=jobans1txt%><br>
		<%=jobans2txt%><br>&nbsp;
		</td><td colspan=2 valign=top style="padding:10px;  border-bottom:1px silver dashed;" class=lillegray><br>
		<b>Faktureret på dette job indtil dagsdato: </b><br>
		
		<%
		if len(faktureretBeloeb) <> 0 then
		faktureretBeloeb = faktureretBeloeb
		else
		faktureretBeloeb = 0
		end if
		%>
		<b>Beløb:</b> <%=formatcurrency(faktureretBeloeb, 2)%><br>
		<b>Timer/Antal/Stk.:</b> <%=timeforbrug%>
		
		<%if fastpris = 1 then%>
		<br><br>
		<b>Restbeløb til fakturering:</b><br>
		<%
		rastfak = (intJobTpris - faktureretBeloeb)
		%><font color=darkred>
		<b><%=formatcurrency(rastfak, 2)%></b>
		</font>
		<input type="hidden" name="restbelob" id="restbelob" value="<%=rastfak%>">
		<%else%>
		<input type="hidden" name="restbelob" id="restbelob" value="100000000">
		<%end if%>
		<br>&nbsp;
		</td>
	</tr>
	<%else%>
	<tr bgcolor="#FFFFFF">
		<td style="border-left:1px #003399 solid;" rowspan="8">&nbsp;</td>
		<td colspan=4 height=5><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td style="border-right:1px #003399 solid;" rowspan="8">&nbsp;</td>
	</tr>
	<%end if '*** Jobinfo%>
	
	
	
	<tr bgcolor="#FFFFFF">
		<td colspan=4 style="padding:10px; border-bottom:1px silver dashed;">
		<b>Type:</b>&nbsp;&nbsp;
		<%
		'** Faktype ***
		select case request("faktype")
		case "0"
		strFaktypeNavn = "Faktura"
		case "1"
		strFaktypeNavn = "Kreditnota"
		case "2"
		strFaktypeNavn = "Rykker"
		case else
		strFaktypeNavn = "Faktura"
		end select
		%>
		
		<%=strFaktypeNavn%>
		
		<br>
		<b><%=strFaktypeNavn%> nr:</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="text" name="FM_faknr" value="<%=strFaknr%>" size="14" maxlength="10" style="font-size:11px;">
		<input type="hidden" name="FM_faknr_opr" value="<%=strFaknr%>">
		
		<%if len(trim(strNote)) <> 0 then%>
		<br><%=strNote%><br>
		<%end if%>
		
		<%if request("faktype") = "2" then%>
		<br>
		<b>Bilags nr:</b><img src="ill/blank.gif" width="23" height="1" alt="" border="0">&nbsp;<input type="text" name="bilagsnrkredit" value="0" size="14" maxlength="10">
		<br>Når denne rykker oprettes, krediteres samtidig den oprindelige faktura.<br> Krediteringen bliver gemt med det angivne bilagsnummer.<br>
		<%end if%>
		
		<font color="darkred"><%=strNote1%></font>
		</td>
	</tr>
	
	
	
	
	
		<%if request("faktype") = 0 then%>
		<tr bgcolor="#FFFFFF">
			<td width=550 colspan=2 style="padding:10px; border-bottom:1px silver dashed;" valign=top>
			<b>Kladde / Endelig:</b><br>
				<%if intBetalt <> 0 then%>
				Godkendt <input type="radio" name="FM_betalt" value="1" checked>
				</td><td colspan=2 style="padding:10px; border-bottom:1px silver dashed;" valign=top>&nbsp;
				<%else%>
				<input type="radio" name="FM_betalt" value="0" checked> <b>Kladde</b><br>
				<input type="radio" name="FM_betalt" value="1"> <font color="limegreen"><b>Godkend</b></font> som endelig.
				</td><td colspan=2 style="padding:10; border-bottom:1px silver dashed;" valign=top class=lillegray><b>Note 1:</b>
				Når en faktura godkendes som endelig, bliver der oprettet en tilhørende postering på de ovenstående valgte konti.
				Når fakturaen er godkendt kan den <u>ikke</u> længere slettes eller opdateres.</font>
				<%end if%>
				</td>
				</tr>
		<%else%>
		<tr bgcolor="#FFFFFF">
		<td colspan=4>&nbsp;
		<input type="hidden" name="FM_betalt" value="1"></td>
		</tr>
		<%end if%>
		
	
	
	
	<%
	if len(lastFakdato) <> 0 then
	lastFakdato = lastFakdato
	else
	lastFakdato = "2001/1/1"
	end if
	
	
		'** Periodeinterval brugt.
		'** Start dato
		call datofindes(request("FM_start_dag"),request("FM_start_mrd"),request("FM_start_aar"))
		stdato = request("FM_start_aar")&"/"&request("FM_start_mrd")&"/"&dagparset
		showStDato = dagparset&"/"&request("FM_start_mrd")&"/"&request("FM_start_aar")
		
		'** Slut dato
		call datofindes(request("FM_slut_dag"),request("FM_slut_mrd"),request("FM_slut_aar"))
		slutdato = request("FM_slut_aar")&"/"&request("FM_slut_mrd")&"/"&dagparset
		showSlutDato = dagparset&"/"&request("FM_slut_mrd")&"/"&request("FM_slut_aar") 
		useLastFakdato = 0
	
	
	
	
	'**** Hvis startdato ligger før sidste fakdato sættes stdato = sidste fakdato ****
	if cdate(showStDato) > cdate(lastFakdato) then 'AND cint(usedt_ival) = 1 
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
	<tr bgcolor="#FFFFFF">
		<td colspan=2 style="padding:10px; border-bottom:1px silver dashed;" valign=top><b>Periode afgrænsning:</b><br>
		<%=formatdatetime(showStDato, 1)%> - <%=formatdatetime(showSlutDato, 1)%>
		<%if useLastFakdato = 1 then%>
		<br><font color="red">(Startdato korrigeret til dagen efter sidste faktura dato)</font>
		<%end if%>
		
		<%if datediff("d", lastFakdato, showStDato) > 1 then%>
		<br><font color="red">(Se note 2)</font>
		<%end if%>
		</td>
		<td colspan=2 style="padding:10px; border-bottom:1px silver dashed;" valign=top class=lillegray><b>Note 2:</b> Sidste faktura dato på dette job er: <b><%=formatdatetime(lastFakdato, 1)%></b><br>
		<%if cdate(lastFakdato) = cdate("1/1/2001") then%>
		(Eller denne er den første faktura!)&nbsp;
		<%end if%>
		For at få alle timer med vil korrekt startdato på periodeafgrænsning vil være dagen efter sidste faktura dato ell. job startdato.</td>
	</tr>
	<%else%>
	<tr bgcolor="#FFFFFF">
		<td colspan=4 style="padding-left:10;" valign=top>&nbsp;</td>
	</tr>
	<%end if%>
	
	<tr bgcolor="#FFFFFF">
		<td colspan=2 style="padding:10px; border-bottom:1px silver dashed;" valign=top>
		<b><%=strFaktypeNavn%> dato:</b>&nbsp;&nbsp;
		
		<%'** Bruger afg. dato fra interval som faktura dato-
		if cint(usedt_ival) = 1 and func <> "red" then
		strDag = request("FM_slut_dag")
		strMrd = request("FM_slut_mrd")
		strAar = request("FM_slut_aar")
		end if
		
		'*** findes datoen?? ***
		call dag()
		
		
		Response.write formatdatetime(strDag &"/"& strMrd &"/"& strAar, 1)
		%>
		<input type="hidden" name="FM_start_dag" id="FM_start_dag" value="<%=strDag%>">
		<input type="hidden" name="FM_start_mrd" id="FM_start_mrd" value="<%=strMrd%>">
		<input type="hidden" name="FM_start_aar" id="FM_start_aar" value="<%=strAar%>">
		<br>&nbsp;</td>
		<td colspan=2 style="padding:10px; border-bottom:1px silver dashed;" valign=top class=lillegray><b>Note 3:</b> 
		Fakturadatoen følger altid slutdatoen på det valgte datointerval.
		<br>&nbsp;
		</td>
	</tr>
	
	<%select case intEnhedsang
	case 1
	chke3 = ""
	chke2 = "CHECKED"
	chke1 = ""
	case 2
	chke3 = "CHECKED"
	chke2 = ""
	chke1 = ""
	case else
	chke3 = ""
	chke2 = ""
	chke1 = "CHECKED"
	end select%>
	
	<tr bgcolor="#FFFFFF">
		<td colspan=4 style="padding:10px;" valign=top><b>Enhed angives som:</b><br>
		<input type="radio" name="FM_enheds_ang" value="0" <%=chke1%>> Timepris&nbsp;&nbsp; 
		<input type="radio" name="FM_enheds_ang" value="1" <%=chke2%>> Stk. pris &nbsp;&nbsp;
		<input type="radio" name="FM_enheds_ang" value="2" <%=chke3%>> Enhedspris.
		</td>
	</tr>
	
	<%if jobid <> 0 then%>
	<tr><td bgcolor="#ffffff" colspan=4 style="padding:10px; border-top:1px silver dashed;">
	<b>Jobbeskrivelse:</b><br>
	
	<%if len(trim(strJobBesk)) <> 0 then
	visJb = "CHECKED"
	else
	visJb = ""
	end if
	%>
	<textarea cols="79" rows="6" name="FM_jobbesk" id="FM_jobbesk" id="FM_jobbesk"><%=strJobBesk%>
	</textarea>
	<br>
	<input type="checkbox" name="FM_visjobbesk" id="FM_visjobbesk" value="1" <%=visJb%>> Vis jobbeskrivelse på faktura.
	</td></tr>
	<%else%>
	<tr><td colspan=4>&nbsp;</td></tr>
	<%end if%>
	
	
	<tr bgcolor="#FFFFFF">
		<td valign="top" colspan=6  style="border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	</table>
	
	
	
	
	
	<%
	'*** Udspecificering af aktiviteter på job **
	if jobid <> 0 then '** jobid <> 0 (Vises kun ved fakturaer på job) %>
	<!--#include file="inc/fak_job_inc.asp"-->
	<%else
	'*** Fakturerer Aftale ***
	%>
	<!--#include file="inc/fak_aft_inc.asp"-->
	<%end if %>
	
<br>
<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
<br>
<br>
</div>




									
									<%if func <> "red" then%>
									<!-- joblog -->
									<div style="position:relative; visibility:visible; left:10; top:470; width:160; height:300; background-color:#ffffff; overflow:auto;">
									<table width=160 border=0 cellspacing=0 cellpadding=2>
									<tr bgcolor="limegreen">
										<td class=alt height=30 valign=bottom style="border:1px #003399 solid;"><b>Joblog i valgt periode:</b></td>
									</tr>
									<tr bgcolor="#ffffff">
									<td style="padding:5px;">Både fakturerbare og ikke-faktorerbare timer vises i jobloggen.<%
									strSQL = "SELECT tmnr, tmnavn, tdato, timer, timerkom, taktivitetnavn, taktivitetid, tfaktim FROM timer WHERE tjobnr = '"& jobnr &"' AND tfaktim <> 5 AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"') ORDER BY taktivitetid, taktivitetnavn, tdato DESC"
									'response.write strSQL
									oRec.open strSQL, oConn, 3 
									x = 0
									while not oRec.EOF 
										
										if oRec("taktivitetid") <> laktaktid then
											
											if oRec("tfaktim") = 1 then
											imgthis = "<img src='../ill/fakbarikon_14px.gif' width='14' height='14' alt='' border='0'>"
											else
											imgthis = "<img src='../ill/notfakbarikon_14px.gif' width='14' height='14' alt='' border='0'>"
											end if
											
										Response.write "<br><br>"& imgthis &"&nbsp;&nbsp;<b>"& oRec("taktivitetnavn") &"</b><hr style='color:#000000; height:1px;'>"
										end if
										
										Response.write "<u>"& oRec("tdato") &":</u> "& oRec("timer") &" "
										Response.write ""& left(oRec("tmnavn"), 12) &"<br>"
										
										if len(oRec("timerkom")) <> 0 then
										Response.write "<font class=megetlillesilver>"& oRec("timerkom") & "</font><br><br>"
										end if
									
									x = x + 1
									laktaktid = oRec("taktivitetid")  
									oRec.movenext
									wend
									oRec.close 
									
									if x = 0 then
									Response.write "<br><br><b><font class=roed>!</font> Ingen registreringer!</b>"
									end if
									%></td></tr>
									</table>
									</div>
									<%end if%>

<%end select%>
<%end if





%>
<!--#include file="../inc/regular/footer_inc.asp"-->


