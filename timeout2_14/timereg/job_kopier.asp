<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->




<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<% 
	errortype = 5
	call showError(errortype)
	else
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	func = request("func")
	id = request("id")
	filt = request("filt")
	'vmenukundefilt = request("fm_kunde")
	showaspopup = request("showaspopup")
	rdir = request("rdir")
	medarbejderid = request("usemrn")
	
	
	select case func
	case "dbkopier"
	
	                %>
	                <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			        <!--#include file="inc/isint_func.asp"-->
			        <%
			        call erDetInt(id) 
			        if isInt > 0 OR instr(id, ".") <> 0 OR len(trim(id)) = 0 OR id = 0 then
        				
				        errortype = 119
				        call showError(errortype)
        			
			        isInt = 0
			        
			        Response.end
			        end if
			
	
		
		
				
					'*** Finder det job der skal kopieres ***
					'*** Rest esitmat og stade_tim_proc kopieres ikke **'
					strSQL = "SELECT id, jobnavn, jobnr, kkundenavn, jobknr, "_
					&" jobTpris, jobstatus, jobstartdato, jobslutdato, projektgruppe1, projektgruppe2, "_
					&" projektgruppe3, projektgruppe4, projektgruppe5, job.dato, job.editor, "_
					&" fakturerbart, budgettimer, fastpris, kundeok, job.beskrivelse, "_
					&" ikkeBudgettimer, tilbudsnr, projektgruppe6, projektgruppe7, "_
					&" projektgruppe8, projektgruppe9, projektgruppe10, serviceaft, kundekpers, "_
					&" jobans1, jobans2, lukafmjob, valuta, "_
					&" jobfaktype, rekvnr, "_
	                &" jo_gnstpris, jo_gnsfaktor, jo_gnsbelob, jo_bruttofortj, jo_dbproc, udgifter, usejoborakt_tp, ski, job_internbesk, "_
                    &" jobans3, jobans4, jobans5, "_
                    &" diff_timer, diff_sum, jo_udgifter_intern, jo_udgifter_ulev, jo_bruttooms, "_
                    &" jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, "_
                    &" salgsans1, salgsans2, salgsans3, salgsans4, salgsans5, "_
                    &" salgsans1_proc, salgsans2_proc, salgsans3_proc, salgsans4_proc, salgsans5_proc, "_
                    &" virksomheds_proc, syncslutdato, altfakadr, fomr_konto, jfak_sprog, jfak_moms "_
					&" FROM job LEFT JOIN kunder ON (kunder.Kid = jobknr) WHERE id=" & id 
					oRec.open strSQL, oConn, 3
							
					
					if not oRec.EOF then	
							
						'strJobNavn = replace(oRec("jobnavn"),"'","''")
						strJobNavn = replace(request("FM_jobnavn"),"'","''")
                        'strjobnr = oRec("jobnr")
						strKnavn = oRec("kkundenavn")
						'strKundeId = oRec("jobknr")
						strBudget = replace(oRec("jobTpris"), ",", ".")
						
						if rdir <> "timereg" then
						strStatus = oRec("jobstatus")
						else
						strStatus = 1 'aktiv
						end if
						
						strTdato = oRec("jobstartdato")
						strUdato = oRec("jobslutdato")
						strProj_1 = oRec("projektgruppe1")
						strProj_2 = oRec("projektgruppe2")
						strProj_3 = oRec("projektgruppe3")
						strProj_4 = oRec("projektgruppe4")
						strProj_5 = oRec("projektgruppe5")
						strProj_6 = oRec("projektgruppe6")
						strProj_7 = oRec("projektgruppe7")
						strProj_8 = oRec("projektgruppe8")
						strProj_9 = oRec("projektgruppe9")
						strProj_10 = oRec("projektgruppe10")
						strDato = oRec("dato")
						strLastUptDato = oRec("dato")
						strEditor = oRec("editor")
						strFakturerbart = 1 'oRec("fakturerbart")
						strBudgettimer = oRec("budgettimer")
						strFastpris = oRec("fastpris")
						intkundeok = oRec("kundeok")
						strBesk = oRec("beskrivelse")
						'strtilbudsnr = oRec("tilbudsnr")
						
						'*** Tilføj job til aftale ***'
						if request("FM_aftaler") = 1 then
						intServiceaft = oRec("serviceaft")
						else
						intServiceaft = 0
						end if
						
						intkundekpers = oRec("kundekpers")
						
						jobans1 = oRec("jobans1")
						jobans2 = oRec("jobans2")
						
							if oRec("ikkeBudgettimer") > 0 then
							ikkeBudgettimer = oRec("ikkeBudgettimer")
							else
							ikkeBudgettimer = 0
							end if
							
							lukafmjob = oRec("lukafmjob")
							valuta = oRec("valuta")
							
							
						
						jobfaktype = oRec("jobfaktype")
						rekvnr = oRec("rekvnr")
	                    jo_gnstpris = replace(oRec("jo_gnstpris"), ",", ".")
	                    jo_gnsfaktor = replace(oRec("jo_gnsfaktor"), ",", ".")
	                    jo_gnsbelob = replace(oRec("jo_gnsbelob"), ",", ".")
	                    jo_bruttofortj = replace(oRec("jo_bruttofortj"), ",", ".")
	                    jo_dbproc = replace(oRec("jo_dbproc"), ",", ".")
	                    udgifter = replace(oRec("udgifter"), ",", ".")
						
						usejoborakt_tp = oRec("usejoborakt_tp")
						
						ski = oRec("ski")
						job_internbesk = oRec("job_internbesk")


                        jobans3 = oRec("jobans3")
                        jobans4 = oRec("jobans4")
                        jobans5 = oRec("jobans5")

                        diff_timer = replace(oRec("diff_timer"), ",", ".")
                        diff_sum = replace(oRec("diff_sum"), ",", ".")
                        jo_udgifter_intern = replace(oRec("jo_udgifter_intern"), ",", ".")
                        jo_udgifter_ulev = replace(oRec("jo_udgifter_ulev"), ",", ".")
                        jo_bruttooms = replace(oRec("jo_bruttooms"), ",", ".")

                        jobans_proc_1 = replace(oRec("jobans_proc_1"), ",", ".")
                        jobans_proc_2 = replace(oRec("jobans_proc_2"), ",", ".")
                        jobans_proc_3 = replace(oRec("jobans_proc_3"), ",", ".")
                        jobans_proc_4 = replace(oRec("jobans_proc_4"), ",", ".")
                        jobans_proc_5 = replace(oRec("jobans_proc_5"), ",", ".")
                        virksomheds_proc = replace(oRec("virksomheds_proc"), ",", ".")

                        salgsans1 = oRec("salgsans1")
                        salgsans2 = oRec("salgsans2")
                        salgsans3 = oRec("salgsans3")
                        salgsans4 = oRec("salgsans4")
                        salgsans5 = oRec("salgsans5")

                        salgsans1_proc = replace(oRec("salgsans1_proc"), ",", ".")
                        salgsans2_proc = replace(oRec("salgsans2_proc"), ",", ".")
                        salgsans3_proc = replace(oRec("salgsans3_proc"), ",", ".")
                        salgsans4_proc = replace(oRec("salgsans4_proc"), ",", ".")
                        salgsans5_proc = replace(oRec("salgsans5_proc"), ",", ".")

                          


                        syncslutdato = oRec("syncslutdato")

                        altfakadr = oRec("altfakadr")

                        fomr_konto = oRec("fomr_konto")
                        jfak_sprog = oRec("jfak_sprog") 
                        jfak_moms = oRec("jfak_moms")
						
					end if
						
					oRec.close
							
							
				
				'*** Finder ID på de kunder jobbet skal kopires til **'			
				if instr(request("FM_kunde"), ",") <> 0 then
				    
				kids = split(request("FM_kunde"), ",")
				for x = 1 to UBOUND(kids)
				        
					
		        'Response.Write kids(x)
		        'Response.flush
							
					
				'**********************************'
				'*** nyt jobnr / tilbudsnr      ***'
				'**********************************'
				strSQL = "SELECT jobnr, tilbudsnr FROM licens WHERE id = 1"
				oRec.open strSQL, oConn, 3
				if not oRec.EOF then 
				    
                    call erDetInt(oRec("jobnr"))
                    if cint(isInt) = 0 then
				    newjobnr = oRec("jobnr") + 1
                    newjobnr_maske = newjobnr 
                    else
                    newjobnr = oRec("jobnr") & "0" & right(now, 2)
                    newjobnr_maske = oRec("jobnr") + 1
                    end if
				    

                            '***tjekker om det findes i forvejen ****'
                            jobnrFindes = 0
                            strSQL2 = "SELECT jobnr FROM job WHERE jobnr = '"& newjobnr &"'"
				            oRec2.open strSQL2, oConn, 3
				            if not oRec2.EOF then 

                            jobnrFindes = 1

                            end if
                            oRec2.close

                            if cint(jobnrFindes) = 1 then
                            newjobnr = oRec("jobnr") &"0"& right(now, 2)
                            newjobnr_maske = oRec("jobnr") + 10
                            end if 

                    isInt = 0

				    if strStatus = 3 then
				    newTilbudsnr = oRec("tilbudsnr") + 1
				    else
				    newTilbudsnr = 0
				    end if
				    
				end if
				oRec.close
				
				'Response.Write "strStatus " & strStatus
				'Response.flush
				'Response.end
				
						
							
					'*** Dato/jobperiode ***
					oprPeriode = datediff("d", strTdato, strUdato)
					startDato = year(now) & "/" & month(now) & "/" & day(now)
					slutDato_temp = dateadd("d", oprPeriode, now)
					slutDato = year(slutDato_temp) & "/" & month(slutDato_temp) & "/" & day(slutDato_temp)
					
					strDato = session("dato")

                            

                   
							
							'now _kopi_"& formatdatetime(now, 2) &"
							strSQLjob = ("INSERT INTO job (jobnavn, jobnr, jobknr, jobTpris, jobstatus, jobstartdato,"_
							&" jobslutdato, editor, dato, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, "_
							&" projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, "_
							&" fakturerbart, budgettimer, fastpris, kundeok, beskrivelse, "_
							&" ikkeBudgettimer, tilbudsnr, jobans1, jobans2, serviceaft, kundekpers, lukafmjob, valuta, "_
							&" jobfaktype, rekvnr, "_
	                        &" jo_gnstpris, jo_gnsfaktor, jo_gnsbelob, jo_bruttofortj, "_
	                        &" jo_dbproc, udgifter, usejoborakt_tp, ski, job_internbesk, "_
                            &" jobans3, jobans4, jobans5, "_
                            &" diff_timer, diff_sum, jo_udgifter_intern, jo_udgifter_ulev, jo_bruttooms, "_
                            &" jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, virksomheds_proc, syncslutdato, altfakadr, "_
                            &" salgsans1, salgsans2, salgsans3, salgsans4, salgsans5, "_
                            &" salgsans1_proc, salgsans2_proc, salgsans3_proc, salgsans4_proc, salgsans5_proc, fomr_konto, jfak_sprog, jfak_moms "_
                            &") VALUES "_
							&"('"& strJobNavn &"', "_
							&"'"& newjobnr &"', "_ 
							&""& trim(kids(x)) &", "_  
							&""& strBudget &", "_ 
							&""& strStatus &", "_ 
							&"'"& startDato &"', "_ 
							&"'"& slutDato &"', "_
							&"'"& session("user") &"', "_
							&"'"& strDato &"', "_ 
							&""& strProj_1 &", "_ 
							&""& strProj_2 &", "_ 
							&""& strProj_3 &", "_ 
							&""& strProj_4 &", "_ 
							&""& strProj_5 &", "_
							&""& strProj_6 &", "_ 
							&""& strProj_7 &", "_ 
							&""& strProj_8 &", "_ 
							&""& strProj_9 &", "_ 
							&""& strProj_10 &", "_  
							&""& strFakturerbart &", "_ 
							&""& strBudgettimer  &", "_ 
							&"'"& strFastpris & "',"_
							&""& intkundeok &", "_
							&"'"& strBesk &"', "& ikkeBudgettimer &", "_
							&" "& newTilbudsnr &", "& jobans1 &", "& jobans2 &", "_
							&" "& intServiceaft &", "& intKundekpers &", "_
							&" "& lukafmjob &", "& valuta &", "_
							&" "& jobfaktype &", '"& rekvnr &"', "_
	                        &" "& jo_gnstpris &", "& jo_gnsfaktor &", "_
	                        &" "& jo_gnsbelob &", "& jo_bruttofortj &", "& jo_dbproc & ", "_
	                        &" "& udgifter &", "& usejoborakt_tp &", "& ski &", '"& job_internbesk &"',"_
                            &" "& jobans3 & ","_
                            &" "& jobans4 & ","_
                            &" "& jobans5 & ","_
                            &" "& diff_timer & ","_
                            &" "& diff_sum & ","_
                            &" "& jo_udgifter_intern & ","_
                            &" "& jo_udgifter_ulev & ","_
                            &" "& jo_bruttooms & ","_
                            &" "& jobans_proc_1 & ","_
                            &" "& jobans_proc_2 & ","_
                            &" "& jobans_proc_3 & ","_
                            &" "& jobans_proc_4 & ","_
                            &" "& jobans_proc_5 & ","_
                            &" "& virksomheds_proc & ","_
                            &" "& syncslutdato & ", "& altfakadr &", "_
                            &" "& salgsans1 &","& salgsans2 &","& salgsans3 &","& salgsans4 &","& salgsans5 &", "_
                            &" "& salgsans1_proc &","& salgsans2_proc &","& salgsans3_proc &","& salgsans4_proc &","& salgsans5_proc &", "& fomr_konto &", "& jfak_sprog &", "& jfak_moms &""_
                            &")")
							
                        'if session("mid") = 1 then
					'		Response.write strFakturerbart & "<br><br>"
					'		Response.write strSQLjob & "<hr>"
					'		'Response.flush
					'		Response.end
                     '   end if

							oConn.execute(strSQLjob)	
							
							'*** til søgeboks på jobside ***'
							jobsog = left(strNavn, 1)	
							
								
							'*** Finder id ****'
							strSQL = "SELECT id FROM job WHERE jobnr = '" & newjobnr & "'"
						    'Response.write strSQL & "<hr>"
							'Response.flush
                            
                            oRec2.open strSQL, oConn, 3
							if not oRec2.EOF then
								varJobIdThis = oRec2("id")
							end if
							oRec2.close
							
							'Response.write "<br>varJobIdThis" & varJobIdThis
                            'Response.end

							'*** tilføjer job i timereg_usejob (Vis guide) ****'
								strSQL = "SELECT DISTINCT(MedarbejderId) FROM progrupperelationer WHERE ("_
								&" ProjektgruppeId = "& strProj_1 &""_
								&" OR ProjektgruppeId =" & strProj_2 &""_
								&" OR ProjektgruppeId =" & strProj_3 &""_
								&" OR ProjektgruppeId =" & strProj_4 &""_
								&" OR ProjektgruppeId =" & strProj_5 &""_
								&" OR ProjektgruppeId =" & strProj_6 &""_
								&" OR ProjektgruppeId =" & strProj_7 &""_
								&" OR ProjektgruppeId =" & strProj_8 &""_
								&" OR ProjektgruppeId =" & strProj_9 &""_
								&" OR ProjektgruppeId =" & strProj_10 &""_
								&") GROUP BY MedarbejderId"
								oRec3.open strSQL, oConn, 3
								while not oRec3.EOF
									strSQL3 = "INSERT INTO timereg_usejob (medarb, jobid) VALUES ("& oRec3("MedarbejderId") &", "& varJobIdThis &")"
									oConn.execute(strSQL3)
									oRec3.movenext
								wend
								oRec3.close
							
							
							'*** Timepriser ***'
							strSQL = "SELECT aktid, medarbid, timeprisalt, 6timepris FROM timepriser WHERE jobid = "& id &" AND aktid = 0"
							oRec.open strSQL, oConn, 3 
							while not oRec.EOF 

                                    tp6 = replace(oRec("6timepris"), ",",".")
									tpalt = replace(oRec("timeprisalt"), ",", ".")

									strSQL2 = "INSERT INTO timepriser "_
									&" (jobid, aktid, medarbid, timeprisalt, 6timepris) "_
									&" VALUES ("& varJobIdThis &", 0, "& oRec("medarbid") &", "& tpalt &", "& tp6 &")"
									

                                    'Response.write strSQL2
                                    'Response.flush
									oConn.execute(strSQL2)
									
							oRec.movenext
							wend
							oRec.close 
							

                            '**** Forretingsområder job ***'
                            strSQLfomr = "SELECT * FROM fomr_rel WHERE for_jobid = "& id &" AND for_aktid = 0"
                            oRec3.open strSQLfomr, oCOnn, 3

                            while not oRec3.EOF

                            strSQLfomri = "INSERT INTO fomr_rel "_
                            &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                            &" VALUES ("& oRec3("for_fomr") &", "& varJobIdThis &", 0, "& replace(oRec3("for_faktor"), ",", ".") &")"

                            oConn.execute(strSQLfomri)

                            oRec3.movenext
                            wend
                            oRec3.close
							
							
							
									'*** Aktiviteter ***'
									'*** Kopierer ikke aktiviteter der er oprettet fra Incidents ***'
									if request("FM_aktiv") = 1 then
							
									strSQLselAkt = "SELECT id AS aktid, navn, beskrivelse, job, fakturerbar, projektgruppe1, "_
									&" projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, "_
									&" projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, "_
									&" aktstartdato, aktslutdato, editor, dato, budgettimer, aktbudget, aktfavorit, aktstatus, fomr, faktor, "_
									&" tidslaas, tidslaas_st, tidslaas_sl, antalstk, "_
									&" tidslaas_man, tidslaas_tir, tidslaas_ons, "_
	                                &" tidslaas_tor, tidslaas_fre, tidslaas_lor, tidslaas_son, fase, sortorder, bgr, aktbudgetsum, easyreg, fravalgt, brug_fasttp, brug_fastkp, fasttp, fasttp_val, fastkp, fastkp_val, avarenr "_
									&" FROM aktiviteter WHERE job =" & id & " AND incidentid = 0"
									oRec.open strSQLselAkt, oConn, 3
									
									while not oRec.EOF 
										
										strNavnAkt = replace(oRec("navn"), "'", "''")
										
										if len(trim(oRec("beskrivelse"))) <> 0 then
										strBeskrivelse = replace(oRec("beskrivelse"), "'", "''")
										else
										strBeskrivelse = oRec("beskrivelse")
										end if
										
										'strjobnr = oRec("job")
										strTdato = oRec("aktstartdato")
										strUdato = oRec("aktslutdato")
										strProj_1 = oRec("projektgruppe1")
										strProj_2 = oRec("projektgruppe2")
										strProj_3 = oRec("projektgruppe3")
										strProj_4 = oRec("projektgruppe4")
										strProj_5 = oRec("projektgruppe5")
										strProj_6 = oRec("projektgruppe6")
										strProj_7 = oRec("projektgruppe7")
										strProj_8 = oRec("projektgruppe8")
										strProj_9 = oRec("projektgruppe9")
										strProj_10 = oRec("projektgruppe10")
										strDato = oRec("dato")
										strLastUptDato = oRec("dato") 
										strEditor = oRec("editor")
										strFakturerbart = oRec("fakturerbar")
										intAktfavorit = oRec("aktfavorit")
										useBudgettimer = replace(oRec("budgettimer"), ",", ".")
										intaktstatus = oRec("aktstatus")
										intFomr = oRec("fomr")
										dblFaktor = oRec("faktor")
										
										tidslaas = oRec("tidslaas")
										tidslaas_st = oRec("tidslaas_st")
										tidslaas_sl = oRec("tidslaas_sl")
										
										
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
										
										tidslaas_man = oRec("tidslaas_man")
				                        tidslaas_tir = oRec("tidslaas_tir")
				                        tidslaas_ons = oRec("tidslaas_ons")
				                        tidslaas_tor = oRec("tidslaas_tor")
				                        tidslaas_fre = oRec("tidslaas_fre")
				                        tidslaas_lor = oRec("tidslaas_lor")
				                        tidslaas_son = oRec("tidslaas_son")
				                        
				                        antalstk = replace(oRec("antalstk"), ",", ".")
				                        aktbudget = replace(oRec("aktbudget"), ",", ".")
				                        aktbudgetsum = replace(oRec("aktbudgetsum"), ",", ".") 
				                        
				                        aktbgr = oRec("bgr")
													
										strfase = oRec("fase")	
										sortorder = oRec("sortorder")
										
										easyreg = oRec("easyreg")		
										fravalgt  = oRec("fravalgt")


                                            brug_fasttp = oRec("brug_fasttp")
                                            brug_fastkp = oRec("brug_fastkp")
                                            fasttp = replace(oRec("fasttp"), ",", ".")
                                            fasttp_val = oRec("fasttp_val")
                                            fastkp = replace(oRec("fastkp"), ",", ".")
                                            fastkp_val = oRec("fastkp_val")

                                        avarenr = oRec("avarenr")
                                        
                                        					
														strSQLinsakt = ("INSERT INTO aktiviteter (navn, beskrivelse, job, fakturerbar, aktfavorit, aktstartdato, aktslutdato, "_
														&" editor, dato, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, "_
														&" projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, "_
														&" budgettimer, aktbudget, aktstatus, fomr, faktor, tidslaas, tidslaas_st, "_
														&" tidslaas_sl, antalstk, "_
														&" tidslaas_man, tidslaas_tir, tidslaas_ons, "_
						                                &" tidslaas_tor, tidslaas_fre, tidslaas_lor, tidslaas_son, fase, sortorder, bgr, aktbudgetsum, easyreg, fravalgt, brug_fasttp, brug_fastkp, fasttp, fasttp_val, fastkp, fastkp_val, avarenr "_
														&") VALUES "_
														&" ('"& strNavnAkt &"', "_
														&"'"& strBeskrivelse &"', "_ 
														&""& varJobIdThis &", "_  
														&""& strFakturerbart &", "_ 
														&""& intAktfavorit &", "_ 
														&"'"& startDato &"', "_ 
														&"'"& slutDato &"', "_
														&"'"& strEditor &"', "_
														&"'"& strDato &"', "_ 
														&""& strProj_1 &", "_ 
														&""& strProj_2 &", "_
														&""& strProj_3 &", "_
														&""& strProj_4 &", "_
														&""& strProj_5 &", "_
														&""& strProj_6 &", "_ 
														&""& strProj_7 &", "_
														&""& strProj_8 &", "_
														&""& strProj_9 &", "_
														&""& strProj_10 &", "_       
														&""& useBudgettimer &", "_
														&""& aktbudget &", "_
														&""& intaktstatus &", "_
														&"" & intFomr &", "_
														&""& replace(dblFaktor,",",".") &", "& tidslaas &", '"& tidslaas_st &"', '"& tidslaas_sl &"', "& antalstk &",  "_
														&""& tidslaas_man &", "& tidslaas_tir &", "& tidslaas_ons &", "_
						                                &""& tidslaas_tor &", "& tidslaas_fre &", "& tidslaas_lor &", "& tidslaas_son &", "_
						                                &" '"& strfase &"', "& sortorder &", "& aktbgr &", "& aktbudgetsum &", "& easyreg &", "& fravalgt &", "& brug_fasttp &","& brug_fastkp &","& fasttp &","& fasttp_val &","& fastkp&","& fastkp_val &", '"& avarenr &"'"_
					                                    &")") 
														
														
														'Response.write strSQLinsakt & "<br><br>"
														'Response.flush
														oConn.execute(strSQLinsakt)
														
														
														varAktIdThis = 0
														'*** Finder id på netop oprettet aktivitet ****
														strSQL3 = "SELECT id FROM aktiviteter WHERE id <> 0 ORDER BY id DESC"
														oRec3.open strSQL3, oConn, 3
														if not oRec3.EOF then
															varAktIdThis = oRec3("id")
														end if
														oRec3.close	
														
														
														'*** Timepriser Aktiviteter ***
														strSQL3 = "SELECT medarbid, timeprisalt, 6timepris, 6valuta FROM timepriser WHERE jobid = "& id &" AND aktid = "& oRec("aktid")
														oRec3.open strSQL3, oConn, 3 
														while not oRec3.EOF 
																
																strSQL4 = "INSERT INTO timepriser "_
																&" (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) "_
																&" VALUES ("& varJobIdThis &", "& varAktIdThis &", "& oRec3("medarbid") &", "_
																&" "& oRec3("timeprisalt") &", "& replace(oRec3("6timepris"), ",", ".") &","& oRec3("6valuta") &")"
																
																oConn.execute(strSQL4)
																
														oRec3.movenext
														wend
														oRec3.close 
									
									                    

                                                        '***** Forrretningsomr Akt. *****'
                                                        strSQLfomr = "SELECT * FROM fomr_rel WHERE for_aktid = "& oRec("aktid")
                                                        oRec3.open strSQLfomr, oConn, 3

                                                        while not oRec3.EOF 

                                                        strSQLfomri = "INSERT INTO fomr_rel "_
                                                        &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                                                        &" VALUES ("& oRec3("for_fomr") &", "& varJobIdThis &", "& varAktIdThis &", "& replace(oRec3("for_faktor"), ",", ".") &")"

                                                        'Response.Write strSQLfomri
                                                        'Response.flush

                                                        oConn.execute(strSQLfomri)

                                                        oRec3.movenext
                                                        wend 
                                                        oRec3.close

									oRec.movenext
									wend
								    oRec.close
								

									end if 'FM_aktiv
								
							
							        'Response.end
							
								
							'**** Filer ****'
							if request("FM_filer") = 1 then
							
								strSQL = "SELECT id, navn, kundese, kundeid, jobid FROM foldere WHERE jobid = "& id
								oRec.open strSQL, oConn, 3 
								while not oRec.EOF 
									
										strSQL2 = "INSERT INTO foldere "_
										&" (navn, kundese, kundeid, jobid, editor, dato)"_
										&" VALUES ('"& oRec("navn") &"', "& oRec("kundese") &", "& oRec("kundeid") &", "& varJobIdThis &", '"& session("user") &"', '"& startDato &"')"	
										
										oConn.execute(strSQL2)
										
										lastFolderID = 0
										'*** Finder id på netop oprettet aktivitet ****
										strSQL3 = "SELECT id FROM foldere WHERE id <> 0 ORDER BY id DESC"
										oRec3.open strSQL3, oConn, 3
										if not oRec3.EOF then
											lastFolderID = oRec3("id")
										end if
										oRec3.close	
										
										
										strSQL3 = "SELECT filnavn, type, jobid, kundeid, oprses, folderid, adg_kunde, adg_admin, adg_alle FROM filer WHERE folderid = " & oRec("id")
										oRec3.open strSQL3, oConn, 3 
										while not oRec3.EOF 
											
											
											strSQL4 = "INSERT INTO filer (editor, dato, filnavn, type, jobid, kundeid, oprses, "_
											&" folderid, adg_kunde, adg_admin, adg_alle) VALUES "_
											&" ('"& session("user") &"', '"& date &"', '"& oRec3("filnavn") &"', "_
											&" '"& oRec3("type") &"', "& varJobIdThis &", "& oRec3("kundeid") &", '"& oRec3("oprses") &"', "_
											&" "& lastFolderID &", "& oRec3("adg_kunde") &", "& oRec3("adg_admin") &", "& oRec3("adg_alle") &")"
											
											oConn.execute(strSQL4)
										
										oRec3.movenext
										wend
										oRec3.close 
								
								
								oRec.movenext
								wend
								oRec.close 
							
							end if '** FM_filer
							
							
							
							
							
							'*** Kopierer under lev **'
							if request("FM_ulev") = 1 then
							
							
							
							strSQLulev1 = "SELECT ju_navn, ju_ipris, ju_faktor, ju_belob, ju_jobid FROM job_ulev_ju WHERE ju_jobid = " & id
							
							
							oRec.open strSQLulev1, oCOnn, 3
							while not oRec.EOF
							
							    strSQLulev2 = "INSERT INTO job_ulev_ju (ju_navn, ju_ipris, ju_faktor, ju_belob, ju_jobid) "_
							    &" VALUES ('"& replace(oRec("ju_navn"), "'", "''") &"', "_
							    &" "& replace(oRec("ju_ipris"), ",", ".") &", "_
							    &" "& replace(oRec("ju_faktor"), ",", ".") &", "_
							    &" "& replace(oRec("ju_belob"), ",", ".") &", "_
							    &" "& varJobIdThis &")"
							    
							     
							    oConn.execute(strSQLulev2)
							    
							
							oRec.movenext
							wend
							oRec.close
							
							
							end if
							
								
								t = 0
								if t = 9 then
								
									'** Viser jo på timereg siden **'
									strSQL = "SELECT DISTINCT(MedarbejderId) FROM progrupperelationer WHERE (ProjektgruppeId = "& strProjektgr1 &" OR ProjektgruppeId =" & strProjektgr2 &" OR ProjektgruppeId =" & strProjektgr3 &" OR ProjektgruppeId =" & strProjektgr4 &" OR ProjektgruppeId =" & strProjektgr5 &") GROUP BY MedarbejderId"
									oRec3.open strSQL, oConn, 3
									while not oRec3.EOF
										strSQL3 = "INSERT INTO timereg_usejob (medarb, jobid) VALUES ("& oRec3("MedarbejderId") &", "& varJobIdThis &")"
										oConn.execute(strSQL3)
										oRec3.movenext
									wend
									oRec3.close
								
								
								end if
								
								
								
								
				                
				                '*** Opdatere jobnrmaske ***'
				                strSQL = "UPDATE licens SET jobnr = "& newjobnr_maske &" WHERE id = 1"
				                oConn.execute(strSQL)
				                
				                
				                '*** Opdaterer tilbudnr maske **'
				                if strStatus = 3 then
				                strSQL = "UPDATE licens SET tilbudsnr = "&  newtilbudsnr &" WHERE id = 1"
				                oConn.execute(strSQL)
				                end if
				                
				                
				                
				 
				        
				    next '*** kids
				    
				   
				    
				else
			    
			    %>
	            <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	            <% 
	            errortype = 118
	            call showError(errortype)
			    Response.end
				end if
				
				'Response.Write kids
				'Response.end		          
								
				
				
				if showaspopup <> "y" then '*** Jobsiden eller res-kalenderopr. siden
				    
				   
				    Response.redirect "jobs.asp?menu=job&shokselector=1&id="&varJobIdThis&"&jobnr_sog="&jobsog&"&filt="&filt&"&fm_kunde="&vmenukundefilt
				   
				   
				else
				    select case rdir
				    case "sdsk"
				    Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
				    Response.Write("<script language=""JavaScript"">window.close();</script>")
				    case "timereg"
				    Response.Redirect "timereg_akt_2006.asp?usemrn="&medarbejderid&"&showakt=1&jobid="&varJobIdThis
				    case else
				    Response.redirect "jbpla_w_opr.asp?func=opr&id=0&sttid="&request("sttid")&"&dato="&request("dato")&"&datostkri="&request("datostkri")&"&datoslkri="&request("datoslkri")&"&jobkopieret=j"
				    end select
				    
				    
				    
				end if						
							
							
							
							
							
							
	
	case else
	
	
	
	if showaspopup <> "y" then%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	<%
	lft = 190
	tp = 70
	else%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	lft = 20
	tp = 20
	
	end if%>
	
	<div id="sindhold" style="position:absolute; left:<%=lft%>; top:<%=tp%>; visibility:visible; z-index:50;">
	
	
	
	
	<%
	oimg = "ikon_job_kopier_tilfoj_48.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = "Kopier job"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	
	tWdth = 600
	tTop = 10
	tLeft = 20
	call tableDiv(tTop,tLeft,tWdth) %>
	
	
	
	<%
	
	    if len(trim(id)) <> 0 then
		id = id
		strSQLjobKri = " id = "& id 
		else
		id = 0
		
		    if len(trim(request("jobnr"))) <> 0 then
		    jobnr = request("jobnr")
		    else
		    jobnr = 0
		    end if
		    
		    strSQLjobKri = " jobnr = "& jobnr 
		    
		end if
	
	%>
	
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<form action="job_kopier.asp?func=kopier" method="post">
	<input type="hidden" name="rdir" id="Hidden2" value="<%=rdir%>">
	<input type="hidden" name="usemrn" id="Hidden3" value="<%=medarbejderid%>">
	
	<input type="hidden" name="filt" id="Hidden4" value="<%=filt%>">
	<input type="hidden" name="showaspopup" id="Hidden5" value="<%=showaspopup%>">
	
	<input type="hidden" name="sttid" id="Hidden6" value="<%=request("sttid")%>">
	<input type="hidden" name="dato" id="Hidden7" value="<%=request("dato")%>">
	<input type="hidden" name="datostkri" id="Hidden8" value="<%=request("datostkri")%>">
	<input type="hidden" name="datoslkri" id="Hidden9" value="<%=request("datoslkri")%>">
	<tr>
		<td style="padding:10px;"><br />

        
		

		<b>Vælg Job, Indstillinger og Kontakter:</b><br />
		<br>
		
		
		
		
		
		<%
		strSQLjob = "SELECT id, jobnavn, jobnr FROM job WHERE" & strSQLjobKri 
		
		oRec.open strSQLjob, oCOnn, 3
	    if not oRec.EOF then
	    
	    strJobnavn = oRec("jobnavn")
	    strJobnr = oRec("jobnr")
	    id = oRec("id")
	    
	    
	    end if
	    oRec.close	
	
		 %>
		 
		 Jobnr på det job du ønsker at kopiere:
		  <br />
            <input id="Text1" name="jobnr" value="<%=strJobnr %>" type="text" />&nbsp;<input id="Submit2" type="submit" value=" Søg >> " />
		 
		 
		
            
		
		</td>
		</tr>
		</form>
		</table>
		
		
		<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<form action="job_kopier.asp?func=dbkopier" method="post">
	
	<input type="hidden" name="id" id="id" value="<%=id%>">
	<input type="hidden" name="rdir" id="rdir" value="<%=rdir%>">
	<input type="hidden" name="usemrn" id="usemrn" value="<%=medarbejderid%>">
	
	<input type="hidden" name="filt" id="filt" value="<%=filt%>">
	<input type="hidden" name="showaspopup" id="showaspopup" value="<%=showaspopup%>">
	
	<input type="hidden" name="sttid" id="sttid" value="<%=request("sttid")%>">
	<input type="hidden" name="dato" id="dato" value="<%=request("dato")%>">
	<input type="hidden" name="datostkri" id="datostkri" value="<%=request("datostkri")%>">
	<input type="hidden" name="datoslkri" id="datoslkri" value="<%=request("datoslkri")%>">
    <tr><td style="padding:10px;">
    
      <b> Jobnavn:</b> <input id="Text2" name="FM_jobnavn" value="<%=strJobnavn %>" style="width:450px;" type="text" />
		
    </td></tr>

		<tr>
		<td style="padding:10px;">
		<br />
		
		<b>Følgende oplsyninger vil blive kopieret:</b><br />
		Jobnavn og jobstamdata.<br />
		Adgangsrettigheder og timepriser på medarbjedere.<br />
		Jobstartdato vil blive sat til dagsdato og jobslutdato vil blive sat =
		dagsdato + den oprindelige jobperiode.<br />
	    Du vil selv komme til at stå som jobansvarlig.<br />
        Forretningsområder på job og aktiviteter.<br />
	    
	    <br /><br />
	    <b>Følgende oplsyninger vil <b>ikke</b> blive kopieret:</b><br />
		Milepæle udover start og slut dato vil <u>ikke</u> blive kopieret.<br />
		Aktiviteter tilknyttet Incidents fra ServiceDesk'en vil <u>ikke</u> blive kopieret.<br /><br />
		
		<input type="checkbox" name="FM_aktiv" id="FM_aktiv" value="1" checked> Kopier aktiviteter.<br>
		<input type="checkbox" name="FM_ulev" id="FM_ulev" value="1" checked> Kopier salgsomkostninger.<br>
		
		<input type="checkbox" name="FM_filer" id="FM_filer" value="1" checked> Kopier tilhørende filer.<br>
		
		
		<%select case lto 
		case "bowe"
		aftCHK = "CHECKED"
		case else
		aftCHK = ""
		end select %>
		
		<input type="checkbox" name="FM_aftaler" id="FM_aftaler" value="1" <%=aftCHK %>> Tilmeld det nye job de samme aftaler, som det oprindelige job er en del af.<br>&nbsp;</td>
	</tr>
	<tr>
	    <td colspan=2 style="padding:10px;"><b>Vælg de kontakter som jobbet skal kopieres til:</b><br /> 
	    
	        <input name="FM_kunde" id="Hidden1" value=0 type="hidden" />
			<select name="FM_kunde" id="Select1" size=10 multiple style="font-size:9px; width:500px;">
			<%
			
			
			strSQL = "SELECT Kkundenavn, Kkundenr, Kid, kundeans1, kundeans2 FROM kunder WHERE ketype <> 'e' ORDER BY Kkundenavn"
			oRec.open strSQL, oConn, 3
			kans1 = ""
			kans2 = ""
			while not oRec.EOF
				
				if cint(strKnr) = oRec("Kid") then
				kSel = "SELECTED"
				else
				kSel = ""
				end if
				
				strSQL2 = "SELECT mnavn, mnr FROM medarbejdere WHERE mid = "&oRec("kundeans1")
				oRec2.open strSQL2, oConn, 3 
				if not oRec2.EOF then
				kans1 = oRec2("mnavn")
				end if
				oRec2.close
				
				strSQL2 = "SELECT mnavn, mnr FROM medarbejdere WHERE mid = "&oRec("kundeans2")
				oRec2.open strSQL2, oConn, 3 
				if not oRec2.EOF then
				kans2 = " - &nbsp;&nbsp;" & oRec2("mnavn") 
				end if
				oRec2.close
				
				if len(kans1) <> 0 OR len(kans2) <> 0 then
				anstxt = "kontaktansv 1: "
				else
				anstxt = ""
				end if
				
			%>
			<option value="<%=oRec("Kid")%>" <%=kSel%>><%=oRec("Kkundenavn")%>&nbsp;&nbsp;(<%=oRec("Kkundenr")%>)&nbsp;&nbsp;-&nbsp;&nbsp;<%=anstxt%><%=kans1%>&nbsp;&nbsp;<%=kans2%></option>
			<%
			kans1 = ""
			kans2 = ""
			oRec.movenext
			wend
			oRec.close
			%>
		</select>
	    </td>
	</tr>
	<tr>
	    <td colspan=2 align="right" style="padding:10px;">
            <input id="Submit1" type="submit" value="Kopier job.." /></td>
	</tr>
	</table>
	
	</form>
	</div>
	<br><br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a><img src="../ill/blank.gif" width="410" height="1" alt="" border="0">
	
	<%
	
	end select
	
end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
