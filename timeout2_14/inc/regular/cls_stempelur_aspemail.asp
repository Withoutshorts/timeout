 
 <%
 public forvalgt_stempelur

 function fv_stempelur()
 forvalgt_stempelur = 1
    
    strSQLfv = "SELECT id FROM stempelur WHERE forvalgt = 1"
    'Response.write strSQLfv
    'Response.flush

    oRec5.open strSQLfv, oConn, 3
    if not oRec5.EOF then

    forvalgt_stempelur = oRec5("id")

    end if
    oRec5.close


end function



public use_ig_sltid, loginTid, manuelt_afsluttet, stpause, stpause2
function stpauser_ignorePer(lgi_dato,lgi_sttid_hh,lgi_sttid_mm)


    strSQL = "SELECT ignorertid_st, ignorertid_sl, stpause, stpause2 FROM licens WHERE id = 1"
	oRec.open strSQL, oConn, 3
    if not oRec.EOF then
    
    ig_sttid = oRec("ignorertid_st")
    ig_sltid = oRec("ignorertid_sl")
    
    stpause = oRec("stpause")
    stpause2 = oRec("stpause2")
   
    end if
    oRec.close
    
    if len(ig_sttid) <> 0 then
    ig_sttid = formatdatetime(ig_sttid, 3)
    end if
    
    if len(ig_sltid) <> 0 then
    ig_sltid = formatdatetime(ig_sltid, 3)
    end if
    
    
                    '*******************************************
					'*** Test for ignorer periode v. login *****
					testloginTid = year(lgi_dato) &"/"& month(lgi_dato)&"/"& day(lgi_dato) & " " & lgi_sttid_hh &":"& lgi_sttid_mm & ":00"
					testig_sttid = year(lgi_dato) &"/"& month(lgi_dato)&"/"& day(lgi_dato) & " " & ig_sttid
					testig_sltid = year(lgi_dato) &"/"& month(lgi_dato)&"/"& day(lgi_dato) & " " & ig_sltid
					
					'Response.Write cdate(testloginTid) &" > "& cdate(testig_sttid) &" AND "& cdate(testloginTid) &" < "& cdate(testig_sltid) &" then <br>"
					
					use_ig_sltid = 0
					
					if cdate(testloginTid) > cdate(testig_sttid) AND cdate(testloginTid) < cdate(testig_sltid) then
					use_ig_sltid = 1
					loginTid = year(lgi_dato) &"/"& month(lgi_dato)&"/"& day(lgi_dato) & " " & ig_sltid
					manuelt_afsluttet = 3
					else
					use_ig_sltid = 0
					manuelt_afsluttet = 0
					loginTid = year(lgi_dato) &"/"& month(lgi_dato)&"/"& day(lgi_dato) & " " & lgi_sttid_hh &":"& lgi_sttid_mm & ":00"
					end if
					'Response.Write loginTid & "<hr>"


end function

public errKonflikt, strLogindkonflikt
function stempelur_tidskonflikt(thisId, thisMid, kloginTid, klogudTid, kthisDato, io)
                
                

                errKonflikt = 0
                if io <> 1 then 
                '1 logind fra logind siden
                '2 logind fra terminal
                '3 Manuelt logind / eller logud via stempelur siden
                
                kloginTid = datepart("yyyy", kloginTid, 2, 2) & "-" & datepart("m", kloginTid, 2, 2) & "-" & datepart("d", kloginTid, 2, 2) & " " & datepart("h", kloginTid, 2, 2) & ":"& datepart("n", kloginTid, 2, 2) & ":00" 
                klogudTid = datepart("yyyy", klogudTid, 2, 2) & "-" & datepart("m", klogudTid, 2, 2) & "-" & datepart("d", klogudTid, 2, 2) & " " & datepart("h", klogudTid, 2, 2) & ":"& datepart("n", klogudTid, 2, 2) & ":00" 
                
                kthisDato = datepart("yyyy", kthisDato, 2, 2) & "-" & datepart("m", kthisDato, 2, 2) & "-" & datepart("d", kthisDato, 2, 2)
                
                strSQL = "SELECT dato, login, logud FROM login_historik WHERE mid = "& thisMid &" AND dato = '"& kthisDato &"'"_
		        &" AND stempelurindstilling <> -1 AND id <> "& thisId &" AND minutter <> 0 AND ((login <= '"& kloginTid &"' AND logud > '"& kloginTid &"') "_
		        &" OR (login <= '"& klogudTid &"' AND logud > '"& klogudTid &"') "_
		        &" OR (login >= '"& kloginTid &"' AND logud <= '"& klogudTid &"')) "
		        
		        'Response.Write "<br>KOnflikt SQL: "& strSQL & "<br>"
		        'Response.flush
		        
		        
		        
		        oRec4.open strSQL, oConn, 3
		        if not oRec4.EOF then
		        
		            if io = 3 then '1 fra TO ellers fra Terminal
		                strLogindkonflikt = oRec4("login") & " til " & oRec4("logud") 
		                errortype = 134
		                call showError(errortype)
		                Response.end
		            else
		                errKonflikt = 1
		            end if
		        

		        end if
		        oRec4.close
		        
		        end if
		        
		        'Response.Write "<br>errKn:" & errKonflikt & "<br>"


end function


public fo_logud
function logindStatus(strUsrId, intStempelur, io, tid)


errIndlaesTerminal = 0
'** io = indlæs / overfør 
if io = 1 then 'fra logind i Timeout

end if
              
if datepart("n", tid) < 10 then
nTid = "0"&datepart("n", tid)              
else
nTid = datepart("n", tid)
end if
                    
LoginDateTime = year(tid)&"/"& month(tid)&"/"&day(tid)&" "& datepart("h", tid) &":"& nTid &":00" 
LoginDato = year(tid)&"/"& month(tid)&"/"&day(tid)
			
                  

                    '*** Henter seneste login / logud ***
				    strSQLlog = "SELECT id, dato, login, logud, mid, stempelurindstilling FROM login_historik WHERE "_
				    &" mid = "& strUsrId &" AND stempelurindstilling > -1 AND logud IS NULL ORDER BY login DESC limit 0, 1" 
				    'id limit 0, 1"
                    '0, 10
    				
				    fo_logud = 0
				    
				    'Response.write strSQLlog
				    'Response.end
				    mailissentonce = 0
				    oRec.open strSQLlog, oConn, 3
                    while not oRec.EOF 
                                
                                
                                
                        
                       
                                '*** Autologud = Firma lukketid på dagen  ***'
                                '*** Lukker gammel logind **'
                                hoursDiff = 0
                                hoursDiff = datediff("h", oRec("login"), LoginDateTime, 2,2)
                                
                                'Response.Write "<br>id: "& oRec("id") &" logind: "& oRec("login") &" og logud: "& LoginDateTime &",  hoursDiff:" & hoursDiff & "<br>"
                               


                                'if cdate(LoginDato) <> cdate(oRec("dato")) then
                                '** Hvis der er mere end 20 timer mellem 20 logind, må det betragtes som at  **'
                                '** man har glemt at logge ind.                                              **'
                                '** TimeOut afslutter seneste logind med lukketid og opretter et nyt logind. **'
                                if cint(hoursDiff) > 20 then
                                logudDag = weekday(oRec("dato"), 2)
                                
                                'Response.Write logudDag &" dato:"& oRec("dato") &": wdn:"& weekdayname(weekday(oRec("dato"),1)) &"<br>"
                                
                                select case logudDag
                                case 1
                                dagkri = "normtid_sl_man"
                                case 2
                                dagkri = "normtid_sl_tir"
                                case 3
                                dagkri = "normtid_sl_ons"
                                case 4
                                dagkri = "normtid_sl_tor"
                                case 5
                                dagkri = "normtid_sl_fre"
                                case 6
                                dagkri = "normtid_sl_lor"
                                case 7
                                dagkri = "normtid_sl_son"
                                end select
                                
                                '*** Henter firmalukketid ****
                                lukketid = "23:59:00"
                                strSQL = "SELECT "& dagkri &" AS lukketid FROM licens WHERE id = 1 "
                                
                                'Response.Write strSQL
                                'Response.flush
                                oRec2.open strSQL, oConn, 3
                                if not oRec2.EOF then
                                    if len(trim(oRec2("lukketid"))) <> 0 then
                                    lukketid = oRec2("lukketid")
                                    else
                                    lukketid = "23:59:00"
                                    end if
                                End if
                                oRec2.close
                                
                                'Response.Write "<br>Lukketid: "& formatdatetime(lukketid, 3)
                                'Response.end
                                
                                LogudDateTime = year(oRec("dato"))&"/"& month(oRec("dato"))&"/"&day(oRec("dato"))&" "& formatdatetime(lukketid, 3)
                                
                                '**** Minutter beregning ***
                                loginTidAfr = formatdatetime(oRec("login"), 3)
                                logudTidAfr = formatdatetime(LogudDateTime, 3)
                               
                               
                                minThisDIFF = datediff("s", loginTidAfr, logudTidAfr)/60
                                minThisDIFF = replace(formatnumber(minThisDIFF, 0), ".", "")
	                            minThisDIFF = replace(formatnumber(minThisDIFF, 0), ",", ".")
                                minThisDIFF = abs(minThisDIFF)
	                            
	                            '*** Er der tidkonflikt med manulet oprettede loginds fra TO ???
                                call stempelur_tidskonflikt(oRec("id"), strUsrId, oRec("login"), LogudDateTime, oRec("dato"), io)
                                if cint(errKonflikt) <> 1 then 
	                            
	                                strSQLupd = "UPDATE login_historik SET "_
	                                &" logud = '"& LogudDateTime &"', minutter = "& minThisDIFF &", "_
	                                &" manuelt_afsluttet = 2, logud_first = '"& LogudDateTime &"' WHERE id = "& oRec("id")
    				                
				                    'Response.Write "<br>AUTO: "& strSQLupd & "<br>"
				                    'Response.end
    				                
				                    oConn.execute(strSQLupd)
    				                
				                    fo_logud = 2
				                    

                                    '***** Opretter AUTO pause ****'
				                    'Response.Write "her"

                                    loginDTp = year(LogudDateTime) &"/"& month(LogudDateTime)&"/"& day(LogudDateTime)
                                    loginTidp = loginDTp & " 00:00:00"
                                    logudTidp = loginDTp & " 00:00:00"
                                    medarbSel = session("mid")

                                    '** Standard pauser fra Licens ****
                                    psDt = now
                                    call stPauserFralicens(psDt)


                                    '*** Adgang for specielle projektgrupper ****'
                                    call stPauserProgrp(medarbSel, p1_grp, p2_grp)
				                    
                                    
                                    '***********************************************************               
	                                '*** Tilføj Pauser *****
                                    '***********************************************************

                                    '*** Tømmer pauser så der er altid kun er indlæst 2 pauser pr. dag pr. medarb.
	                                strSQLpDel = "DELETE FROM login_historik WHERE stempelurindstilling = -1  AND dato = '"& loginDTp &"' AND mid = "& medarbSel
	                                oConn.execute(strSQLpDel)
	                                
                                    'Response.Write strSQLpDel & "<br>"
                                    p1 = stPauseLic_1
	                                'Response.Write " p1on " & p1on & " p1: "& p1
	                        
	                                if p1on <> 0 then
	                            
	                                    call tilfojPauser(0,medarbSel,loginDTp,loginTidp,logudTidp,p1,p1_komm)
					        
					                end if
					        
					                'Response.Write " p2on " & p2on
					                
                                    p2 = stPauseLic_2
					                if p2on <> 0 then
					            
                                         call tilfojPauser(0,medarbSel,loginDTp,loginTidp,logudTidp,p2,p2_komm)
	                        
	                                end if
                                    
                                    
                                    'Response.end
                                                
				                                
				                    '***** Oprettter Mail object ***
			                        if cint(mailissentonce) = 0 AND (request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login.asp") _
			                        AND (request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\stempelur_terminal_ns.asp") then
                        			
                        			
                                    Set Mail = Server.CreateObject("Persits.MailSender") 
							
							            Mail.Host = "abizmail1.abusiness.dk" 'webmail.abusiness.dk;
                                        Mail.Port = 25 ' Optional. Port is 25 by default 
							            Mail.From = "timeout_no_reply@outzource.dk" ' Required
							            Mail.FromName = "TimeOut Stempelur Service" ' Optional 
							           
							           
													
							            'Mail.AddEmbeddedImage "f:\webs\dekaeresmaa.dk\wwwroot\ill\logo.gif", "dks"
													
							            '*** skal være slået fra. Ellers vises body indhold ikke.
							            'Mail.CharSet = 2
									

			                    
                                    'Mailens emne
			                        Mail.Subject = "Medarbejder "& session("user") &" har glemt at logge ud."  
			                       
			                        'Modtagerens navn og e-mail
			                        select case lto
			                        case "dencker" 
			                        Mail.AddAddress "dv@dencker.net", "Anders Dencker"
                                    Mail.AddBcc "sk@outzource.dk", "TimeOut Support"

                                    call meStamdata(session("mid"))
                                    if len(trim(meEmail)) <> 0 then
                                    Mail.AddCC ""& meEmail &","& meNavn
                                    end if
                                    
			                        case "outz"
			                        Mail.AddAddress "sk@outzource.dk", "OutZourCE"
                                    
                                    case "jttek"
                                    Mail.AddAddress "jt@jtteknik.dk", "JT-Teknik"
                                    
                                    call meStamdata(session("mid"))
                                    if len(trim(meEmail)) <> 0 then
                                    Mail.AddCC ""& meEmail &","& meNavn
                                    end if

                                    Mail.AddCC "sk@outzource.dk","TimeOut Support"

			                        case else
			                        'Mailer.AddRecipient "OutZourCE", "timeout_no_reply@outzource.dk"
                                    Mail.AddAddress "timeout_no_reply@outzource.dk", "OutZourCE"
			                        end select
                        			
			                       
                        			
			                                ' Selve teksten
					                        bodyHTML = "Medarbejder "& session("user") &" har glemt at logge ud "& weekdayname(weekday(oRec("dato"), 1)) &" d. "& oRec("dato") &"." & vbCrLf & vbCrLf _ 
					                        & "Der er oprettet en auto-logud tid der er sat til jeres firmas normale arbejdstid. (Sættes i kontrolpanelet)"  & vbCrLf & vbCrLf _ 
					                        & "Med venlig hilsen"& vbCrLf & "TimeOut Stempelur Service" & vbCrLf 
					                        
                                            Mail.Body = bodyHTML
							                'Mail.IsHTML = True
                                            '***** Queue = true ***'
                                            'Mail.Queue = True
                                            
                                            '** Afsender mail ***
							                'On Error Resume Next
							                Mail.Send


					                        'If Mailer.SendMail Then
                            				
					                        'Else
    					                    '    Response.Write "Fejl...<br>" & Mailer.Response
  					                        'End if	
                        				
                        			mailissentonce = 1
			                        end if ''** Mail
				                
				                
				                else
				                
				                '**ignorer Terminal logind pgs tidskonflikt 
                                fo_logud = 3
				                end if
				                                
				                                
				                
				                else
				                '** Der findes et igangværende uafsluttet logind ***'
				                '** Ikke hvis der er brugt Terminal, så skal alle logind overføres.
				                '** Kun fra TimeOut, hvis f.eks browser er genstartet, eller CPU gået ned.
				                '** Så skal der ikke oprettes et nyt logind. **'
				                
				                    if io = 1 then 
				                    '** 1 fra TO logind siden 
				                    '** 2 fra Terminal
				                    '** 3 fra Stempelurs siden manuelt eller logiud
				                    fo_logud = 3
				                    else
				                    
				                    
				                 
				                LogudDateTimeDDMMYYY_TIME = day(LoginDateTime)&"-"& month(LoginDateTime)&"-"&year(LoginDateTime)&" "& formatdatetime(LoginDateTime, 3)
                                LogudDateTime = LoginDateTime
                                
                                '**** Minutter beregning ***
                                loginTidAfr = oRec("login") 'formatdatetime(oRec("login"), 3)
                                logudTidAfr = LogudDateTimeDDMMYYY_TIME 'formatdatetime(LogudDateTimeDDMMYYY_TIME, 3)
                               
                                'Response.Write "<br>Min: "& oRec("login") & " - "& LogudDateTimeDDMMYYY_TIME
                               
                                minThisDIFF = datediff("s", loginTidAfr, logudTidAfr)/60
                                minThisDIFF = replace(formatnumber(minThisDIFF, 0), ".", "")
	                            minThisDIFF = replace(formatnumber(minThisDIFF, 0), ",", ".")
                                minThisDIFF = abs(minThisDIFF)
	                            
	                            
	                            '** Ignorer Terminal hvis der er oprettet et manuelt fra TO i samme peridode **'
	                            if cDate(LogudDateTimeDDMMYYY_TIME) >= cDate(oRec("login")) then
	                            
	                                '*** Er der tidkonflikt med manulet oprettede loginds fra TO ???
                                    call stempelur_tidskonflikt(oRec("id"), strUsrId, oRec("login"), LogudDateTime, oRec("dato"), io)
                                    
                                    if cint(errKonflikt) <> 1 then 
    	                            

                                    '*** Afslutter eksisterende logind ****' 
	                                strSQLupd = "UPDATE login_historik SET "_
	                                &" logud = '"& LogudDateTime &"', minutter = "& minThisDIFF &",  "_
	                                &" logud_first = '"& LogudDateTime &"' WHERE id = "& oRec("id")
    				                
				                    'Response.Write "<br>Alm. logud: "& strSQLupd & "<br>"
				                    'Response.flush
    				                
				                    oConn.execute(strSQLupd)

                                                


                                                    '*** Indlæser standardpauser fra Terminal fil
                                                    if thisfile = "stempelur_terminal_ns" then

                                                  

                                                    psloginTidp = LoginDato
                                                    pslogudTidp = psloginTidp



                                                    call stPauserFralicens(LoginDato)

                                                    '** tjekker om der skal tilføjes pause til projektgruppe ***' 
                                                    call stPauserProgrp(strUsrId, p1_grp, p2_grp)

                                                    
                                                    
                                                    '**** Tømmer pauser så der er altid kun er indlæst 2 pauser pr. dag pr. medarb. ****
                                                    strSQLpDel = "DELETE FROM login_historik WHERE stempelurindstilling = -1  AND dato = '"& LoginDato &"' AND mid = "& strUsrId
	                                                oConn.execute(strSQLpDel)


                                                    if p1_prg_on = 1 then
                                                    psKomm_1 = ""
                                                    psMin_1 = stPauseLic_1

                                                    
                                                    '** p1 **
                                                    call tilfojPauser(0,strUsrId,LoginDato,psloginTidp,pslogudTidp,psMin_1,psKomm_1)

                                                    end if


                                                    if p2_prg_on = 1 then
                                                    psKomm_2 = ""
                                                    psMin_2 = stPauseLic_2

                                                    '** p2 **
                                                    call tilfojPauser(0,strUsrId,LoginDato,psloginTidp,pslogudTidp,psMin_2,psKomm_2)

                                                    end if

                                                    end if



    				                
				                    end if
				                
				                end if
				                
				                fo_logud = 3   
				                    
				                    end if
				                    
				                end if
                         
                        
                        
                    oRec.movenext   
                    wend
                    oRec.close
				
				
				'Response.Write "<br>fo_logud" & fo_logud & "<br>"
				'Response.end
				
				
				'*** Hvis der ikke er et igangværende logind, oprettes et nyt. **'
				'** fo_logud = 0 (indlæs ny) **'
                if cint(fo_logud) <> 3 then
				
				
				'** Ignorer periode skal ikke kaldes, da man så 
				'** altid bliver tvunget til at angive en kommentar for ændret logind,
				'** ved logind i et interval der skal ignores. 
				
				lgi_sttid_hh = datepart("h", tid) 
				lgi_sttid_mm = datepart("n", tid)
				call stpauser_ignorePer(LoginDato,lgi_sttid_hh,lgi_sttid_mm)
			    
			    
			    if thisfile <> "stempelur_terminal_ns" then
			    ipn = request.servervariables("REMOTE_ADDR")
			    else
			    ipn = 0
			    end if
				
				
    				
				    strSQL = "INSERT INTO login_historik "_
				    &"(dato, login, mid, stempelurindstilling, login_first, ipn, manuelt_afsluttet) "_
				    &" VALUES ('"& LoginDato &"', '"& LoginTid &"', "& strUsrId &", "_
				    &" "& intStempelur &", '"& LoginDateTime &"', '"& ipn &"', "& manuelt_afsluttet &")"
				    'Response.write "<br>ins SQL: "& strSQL & "<br>"
				    'Response.flush
				    oConn.execute(strSQL)
				    
                  

                    
    				
				  
				
				end if

                'Response.end

end function



public lastMnavn, lastMnr, totalhours, totalmin
function stempelurlist(medarbSel, showtot, layout, sqlDatoStart, sqlDatoSlut, typ, d_end, lnk)



    %>
    



<script language="javascript">


    $(document).ready(function () {




        $(".loginhh").keyup(function () {



            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(12, idlngt)

            eVal = $("#FM_login_mm_" + idtrim).val()
            if (eVal.length == 0) {
                $("#FM_login_mm_" + idtrim).val('00')
            }

           
        });


        $(".logudhh").focus(function () {

            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(12, idlngt)

            eVal = $("#FM_logud_mm_" + idtrim).val()
            if (eVal.length == 0) {
                $("#FM_logud_mm_" + idtrim).val('00')
            }
            

        });






    });

</script>
    <%
	
    'sqlDatoStart = day(sqlDatoStart) &"-"& month(sqlDatoStart) &"-"& year(sqlDatoStart)

	'Response.Write "layout" & layout
	'** Fra stempels hsitorik ***'
	if layout <> 1 then
	medarbSel = 0
	            
	            for m = 0 to UBOUND(intMids)
			    
			     if m = 0 then
			     medSQLkri = " AND (l.mid = " & intMids(m)
			     else
			     medSQLkri = medSQLkri & " OR l.mid = " & intMids(m)
			     end if
			     
			    next
			    
			    medSQLkri = medSQLkri & ")"
			    
	else		    
	
	if medarbSel <> 0 then
	medSQLkri = " AND l.mid = " & medarbSel
	else
	medSQLkri = " AND l.mid <> 0 "
	end if
	
	end if
	
	'*** Finder afslutte uger på aktive medarbejdere ***'
	if cint(medarbSel) = 0 then
	strSQLmedarb = "SELECT mid FROM medarbejdere WHERE mansat <> '2' AND mansat <> '3'"
	oRec4.open strSQLmedarb, oConn, 3
	while not oRec4.EOF 
	    
	    call afsluger(oRec4("mid"), sqlDatoStart, sqlDatoSlut) 
	    
	oRec4.movenext
	wend
	oRec4.close
	
	else
	
	call afsluger(medarbSel, sqlDatoStart, sqlDatoSlut) 
	
	end if
	
	
	
	if showTot = 1 then
	csp = 3
	else
	csp = 10
	end if%>
	
	
    <%
    
   
    
    if layout = 1 then
    tblwdt = "100%"
    lnkTarget = "_blank"
    'tTop = 22
    'tLeft = 0
    'tWdth = "100%"
    hidemenu = "1"
    
    'nyTopPX = "690"
    'nyLeftPX = "250"

    else
    
    'nyTopPX = "0"
    'nyLeftPX = "0"
    
    
    tblwdt = "100%"
     tTop = 10
    tLeft = 0
    tWdth = 1004
    lnkTarget = "_self"
    hidemenu = ""
    call tableDiv(tTop,tLeft,tWdth)
    end if
    
    
    url = "stempelur.asp?menu=stat&func=oprloginhist&id=0&medarbSel="&medarbSel&"&showonlyone="&showonlyone&"&hidemenu="&hidemenu
    '&"&rdir="&rdir
    text = "Opret ny / pause "
    otoppx = 0
    oleftpx = 0
    owdtpx = 140
    java = "Javascript:window.open('"&url&"','','width=650,height=750,resizable=yes,scrollbars=yes')"
    urlhex = "#"
    
   

    %>

   
   
    
    <table cellspacing="0" cellpadding="0" border="0" width="<%=tblwdt %>">
     <%if media <> "print" AND layout = "1" then %>
    <tr><td colspan=<%=csp+2%> align=right style="padding:0px 20px 5px 0px;"><%call opretNyJava(urlhex, text, otoppx, oleftpx, owdtpx, java)%></td></tr>
    <%end if %>
    <form method="post" action="stempelur.asp?func=dbloginhist<%=lnk%>">
    
   
	<tr bgcolor="#5582D2">
		<td width="8" valign=top style="border-top:1px #8caae6 solid; border-left:1px #8caae6 solid;" rowspan=2><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan="<%=csp%>" valign="top" style="border-top:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right width="8" style="border-top:1px #8caae6 solid; border-right:1px #8caae6 solid;" valign=top rowspan=2><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class=alt><b>Dato</b></td>
		<%if showTot <> 1 then%>
		<td class=alt><b>Logget ind</b></td>
		<td class=alt><b>Logget ud</b></td>
		<td class=alt align=right style="padding-right:10px;"><b>IP</b></td>
		<%end if%>
		<td class=alt style="padding-left:0px;"><br /><b>Stempelur indstilling</b><br /> (Faktor / Minimum)</td>
		<td class=alt align=right><br /><b>Timer</b><br /> (hele min.)</td>
		<%if showTot <> 1 then %>
		<td class=alt style="padding-left:15px;"><b>Manuelt opr?</b></td>
		<td class=alt style="padding-left:5px;"><b>Logud ændret</b></td>
		<td class=alt style="padding-left:5px;"><b>Kommentar</b></td>
        <td>&nbsp</td>
		<%end if %>
	</tr>
	<%


    '****** Uge Loop ****'
    lastMid = 0
    g = 0
    medIDNavnWrt = "#0#"


    for d = 0 to d_end '6

    dtUse = dateAdd("d", d, sqlDatoStart)
    dtUseSQL = year(dtUse) &"/"& month(dtUse) &"/"& day(dtUse)
    'medIDNavnWrt = "#0#"


     if d = 0 then 'lastWeek <> datepart("ww", oRec("dato"), 2,2) OR%>
	
	<tr bgcolor="#cccccc">
		<td>&nbsp;</td>
		<td colspan="<%=csp%>" style="height:20px; padding:5px;"><b>Uge <%=datepart("ww", dtUse, 2,2) &" "& datepart("yyyy", dtUse, 2,2) %> </b></td>
		<td>&nbsp;</td>
	</tr>
	
	<%
    end if

    if cint(typ) <> 0 then
    strTypSQL = " AND l.stempelurindstilling = "& typ
    else
    strTypSQL = " AND l.stempelurindstilling <> 0 "
    end if
	
	if showTot = 1 then
	
	strSQL = "SELECT l.id AS lid, count(l.id) AS antal, l.mid AS lmid, m.mnavn AS mnavn, m.mnr, l.login, l.logud, m.init, "_
	&" s.navn AS stempelurnavn, s.faktor, s.minimum, l.stempelurindstilling, sum(l.minutter) AS minutter, l.dato FROM login_historik l"_
	&" LEFT JOIN medarbejdere m ON (m.mid = l.mid) "_
	&" LEFT JOIN stempelur s ON (s.id = l.stempelurindstilling) WHERE "_
    &" l.dato = '"& dtUseSQL &"' "& medSQLkri & strTypSQL &""_
	&" GROUP BY l.mid, l.stempelurindstilling ORDER BY mnavn, l.login DESC " 
	
	
	else
	
	strSQL = "SELECT l.id AS lid, l.mid AS lmid, m.mnavn AS mnavn, m.mnr, m.init, l.login, l.logud, "_
	&" s.navn AS stempelurnavn, s.faktor, s.minimum, "_
	&" l.stempelurindstilling, l.minutter, manuelt_afsluttet, manuelt_oprettet, l.dato, "_
	&" l.kommentar, logud_first, login_first, l.ipn FROM login_historik l"_
	&" LEFT JOIN medarbejdere m ON (m.mid = l.mid) "_
	&" LEFT JOIN stempelur s ON (s.id = l.stempelurindstilling) WHERE "_
    &" l.dato = '"& dtUseSQL &"' "& medSQLkri & strTypSQL &""_
	&" ORDER BY mnavn, l.login DESC" 
	
	end if
	
	'AND l.logud IS NOT NULL GROUP BY l.mid, l.dato, l.stempelurindstilling
	'Response.write medSQLkri &":"& medarbSel
	'Response.flush

    'if lto = "dencker" then
    'Response.write strSQL
    'end if
	

	x = 0
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
	
	'afslSTDato = "2009-1-1"
	'afslSLDato = "2009-12-31"
	'call afsluger(oRec("lmid"), afslSTDato, afslSLDato)
	
	thours = 0
	tmin = 0
	
	
	'*** Total ***
	if lastMid <> oRec("lmid") AND x > 0 then
        'Response.write "lastMId:"& lastMid &" lmid:"& oRec("lmid")

		call tottimer(lastMnavn, lastMnr, totalhours, totalmin, lastMid, sqlDatoStart, sqlDatoSlut, 1)
		totalhours = 0 
		totalmin = 0
	end if
	%>
	
	
	<%if lastMid <> oRec("lmid") then
    
  
    
    if instr(medIDNavnWrt, "#"&oRec("lmid")&"#") = 0 then

      meTxt = oRec("mnavn") & "("& oRec("mnr") &")"
      
      if len(trim(oRec("mnavn"))) <> 0 then
      minit = " - "& oRec("init")
      else
      minit = ""
      end if


      meTxt = meTxt & minit
    %>
	

	<tr bgcolor="#eff3ff">
		<td>&nbsp;</td>
		<td colspan="<%=csp%>" style="height:20px; padding:5px;"><b><%=meTxt%></b></td>
		<td>&nbsp;</td>
	</tr>
	
	<%
    medIDNavnWrt = medIDNavnWrt & ",#"& oRec("lmid") & "#" 
    end if
    
    end if%>

   
	
	<tr>
		<td bgcolor="#cccccc" colspan="<%=csp+2%>"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	
	<%
	if cdbl(lastID) = oRec("lid") then
	bgThis = "#ffff99"
	else
	    'if cint(oRec("stempelurindstilling")) = -1 then
	    'bgThis = "#c4c4c4"
	    'else
	    bgThis = "#ffffff"
	    'end if
	end if
	
	
	if len(trim(oRec("dato"))) <> 0 then
	loginDTShow = left(weekdayname(weekday(oRec("dato"))), 3) & "d. " & formatdatetime(oRec("dato"), 2)
	else
	loginDTShow = "--"
	end if%>
	
	<tr bgcolor="<%=bgThis%>">
        
          <%if cint(oRec("stempelurindstilling")) <> -1 then %>
        <input type="hidden" value="<%=oRec("lid") %>" name="id" />
        <input type="hidden" value="<%=oRec("lmid") %>" name="mid" />
        <input type="hidden" value="<%=formatdatetime(dtUse, 2) %>" name="logindato" />
        <%end if %>
        
		<td style="height:20px;">&nbsp;</td>
		<td style="white-space:nowrap;" align=right>
		<%if showTot = 1 then%>
			<%=oRec("antal")%> stk.&nbsp;
		<%else
		        '** er periode godkendt ***'
		        tjkDag = oRec("dato")
		        erugeafsluttet = instr(afslUgerMedab(oRec("lmid")), "#"&datepart("ww", tjkDag,2,2)&"_"& datepart("yyyy", tjkDag) &"#")
                
		        
		        'Response.Write "smilaktiv: "& smilaktiv & "<br>"
		        'Response.Write "ugeNrAfsluttet: "& ugeNrAfsluttet & "<br>"
		        'Response.Write "autolukvdatodato: "& autolukvdatodato & "<br>"
		        'Response.Write "tjkDag: "& tjkDag & "<br>"
		        'Response.Write "autolukvdato: "& autolukvdato & "<br>"
		        'Response.Write "erugeafsluttet:" & erugeafsluttet & "<br>"
		        
		        call lonKorsel_lukketPer(tjkDag)
		         
                if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", tjkDag) = year(now) AND DatePart("m", tjkDag) < month(now)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", tjkDag) < year(now) AND DatePart("m", tjkDag) = 12)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", tjkDag) < year(now) AND DatePart("m", tjkDag) <> 12) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", tjkDag) > 1))) OR cint(lonKorsel_lukketIO) = 1 then
              
                ugeerAfsl_og_autogk_smil = 1
                else
                ugeerAfsl_og_autogk_smil = 0
                end if 
                
                
                if (ugeerAfsl_og_autogk_smil = 0 OR (level = 1)) AND media <> "print" then%>
			    <a href="#" onclick="Javascript:window.open('stempelur.asp?id=<%=oRec("lid")%>&menu=stat&func=redloginhist&medarbSel=<%=medarbSel%>&showonlyone=<%=showonlyone%>&hidemenu=<%=hidemenu%>&rdir=popup','','width=650,height=650,resizable=yes,scrollbars=yes')" class="vmenu"><%=loginDTShow%></a>
			    <%else %>
		        <b><%=loginDTShow%></b>
		        <%end if %>
		
		<%end if%>
		</td>
		
		    <%if showTot <> 1 then%>
		    
		    <%if cint(oRec("stempelurindstilling")) <> -1 then %>
			<td align=right class=lille>
                    
                            <%if len(trim(oRec("login"))) <> 0 then
                    
                            loginTidThisHH = left(formatdatetime(oRec("login"), 3), 2)
                            loginTidThisMM = mid(formatdatetime(oRec("login"), 3),4,2)

                                if len(trim(loginTidThisHH)) <> 0 then
                                loginTidThisHH_opr = loginTidThisHH
                                loginTidThisMM_opr = loginTidThisMM
                                else
                                loginTidThisHH_opr = "00"
                                loginTidThisMM_opr = "00"
                                end if
                    
                            else

                            loginTidThisHH = ""
                            loginTidThisMM = ""
                             loginTidThisHH_opr = "00"
                            loginTidThisMM_opr = "00"

			                end if  
                    
                    if (layout = 1) AND (ugeerAfsl_og_autogk_smil = 0 OR (level = 1)) AND media <> "print"  then
                    %>
                    <input type="text" class="loginhh" name="FM_login_hh" id="FM_login_hh_<%=d %>" value="<%=loginTidThisHH%>" style="width:20px; font-size:9px;" />:
                    <input type="text" name="FM_login_mm" id="FM_login_mm_<%=d %>" value="<%=loginTidThisMM%>" style="width:20px; font-size:9px;" />
                    &nbsp;</td>  
			
                  <input type="hidden" name="oprFM_login_hh" id="Hidden5" value="<%=loginTidThisHH_opr%>">
		          <input type="hidden" name="oprFM_login_mm" id="Hidden6"  value="<%=loginTidThisMM_opr%>">
	    
	              <!-- bruges til arrray split -->
	              <input type="hidden" name="FM_login_hh" id="Hidden7" value="#">
		          <input type="hidden" name="FM_login_mm" id="Hidden8"  value="#"> 
                  <%else %>
                  <%=loginTidThisHH%>:<%=loginTidThisMM%>
                  
                  <%end if %> 


			<td align="right" class=lille>
            
            <%
			
                if oRec("logud") = oRec("login") then
                fColsl = "#cccccc"%>
                <%else 
                fColsl = "#000000"%>
                <%end if 
                
                
                if len(oRec("logud")) <> 0 then 
                logudTidThisHH = left(formatdatetime(oRec("logud"),3),2) 
                logudTidThisMM = mid(formatdatetime(oRec("logud"),3),4,2) 

                if len(logudTidThisHH) <> 0 then
                logudTidThisHH_opr = logudTidThisHH
                logudTidThisMM_opr = logudTidThisMM
                else
                logudTidThisHH_opr = "00"
                logudTidThisMM_opr = "00"
                end if

                else
                logudTidThisHH = ""
                logudTidThisMM = ""

                logudTidThisHH_opr = "00"
                logudTidThisMM_opr = "00"

                end if

                if (layout = 1) AND (ugeerAfsl_og_autogk_smil = 0 OR (level = 1)) AND media <> "print"  then
                    
                %>
                  <input type="text" class="logudhh" name="FM_logud_hh" id="FM_logud_hh_<%=d %>" value="<%=logudTidThisHH%>" style="width:20px; font-size:9px; color:<%=fColsl%>;" />:
                  <input type="text" name="FM_logud_mm" id="FM_logud_mm_<%=d %>" value="<%=logudTidThisMM%>" style="width:20px; font-size:9px; color:<%=fColsl%>;" />
                  <input type="hidden" name="oprFM_logud_hh" id="Hidden1" value="<%=logudTidThisHH_opr%>">
		          <input type="hidden" name="oprFM_logud_mm" id="Hidden2"  value="<%=logudTidThisMM_opr%>">
	    
	              <!-- bruges til arrray split -->
	              <input type="hidden" name="FM_logud_hh" id="Hidden3" value="#">
		          <input type="hidden" name="FM_logud_mm" id="Hidden4"  value="#">  
                 


                <%else %>

                <%=logudTidThisHH%>:<%=logudTidThisMM%>
                <%end if %>
                </td> 
       
                 

		    <td align=right style="padding-left:10px; padding-right:10px;" class=lille><%=oRec("ipn") %></td>
		    <%else %>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		    <%end if %>
		 
		<%end if%>
		
		<td class=lille>
        <%
        

        if cint(oRec("stempelurindstilling")) <> -1 then 

                if (layout = 1) AND (ugeerAfsl_og_autogk_smil = 0 OR (level = 1)) AND media <> "print"  then
        
              %>
                    <select name="FM_stur" style="width:140px; font-size:9px; font-family:arial;">
		            <%if lto <> "fk_bpm" AND lto <> "kejd_pb" AND lto <> "kejd_pb2" then %>
                    <option value="0">Ingen</option>
                    <%end if %>
		            <%
		            strSQL5 = "SELECT id, navn, faktor, minimum FROM stempelur ORDER BY navn"
		            oRec5.open strSQL5, oConn, 3 
		            while not oRec5.EOF 
		
		            if oRec("stempelurindstilling") = oRec5("id") then
		            sel = "SELECTED"
		            else
		            sel = ""
		            end if
		            %>
		            <option value="<%=oRec5("id")%>" <%=sel%>><%=oRec5("navn")%> (<%=oRec5("faktor")%> / <%=oRec5("minimum")%> min.)</option>
		
		            <%
		            oRec5.movenext
		            wend
		            oRec5.close 

        
		            %>
                    </select>
               
               
               <%else
                
            

            
		            strSQL5 = "SELECT id, navn, faktor, minimum FROM stempelur WHERE id = "& oRec("stempelurindstilling") &" ORDER BY navn"
		            oRec5.open strSQL5, oConn, 3 
		            if not oRec5.EOF then
		
		
		            %>
		          <%=oRec5("navn")%> (<%=oRec5("faktor")%> / <%=oRec5("minimum")%> min.)
		
		            <%
		    
		            end if
		            oRec5.close 

                end if

        else
        %>
        Pause
        <%
        end if %>
		</td>
		
		<td align=right class=lille>
		<%
		timerThis = 0
		timerThisDIFF = 0
		
		if len(oRec("login")) <> 0 AND len(oRec("logud")) <> 0 then
		    
		    if cint(oRec("stempelurindstilling")) = -1 then
		        
		        timerThisDIFF = oRec("minutter")
		        useFaktor = 1
		        
		    else
		    
		        timerThisDIFF = oRec("minutter")
    		
		        if timerThisDIFF < oRec("minimum") then
			        timerThisDIFF = oRec("minimum")
		        end if
    		    
		        if oRec("faktor") > 0 then
			    useFaktor = oRec("faktor")
			    else
			    useFaktor = 0
			    end if
			
			end if
		
		
		timerThis = (timerThisDIFF * useFaktor)
		'totaltimer = totaltimer + timerThis
		end if
		
		'Response.write timerThis
		'call timerogminutberegning(0,timerThis)
		
		timTemp = formatnumber(timerThis/60, 3)
		timTemp_komma = split(timTemp, ",")
		
		for x = 0 to UBOUND(timTemp_komma)
			
			if x = 0 then
			thours = timTemp_komma(x)
			end if
			
			if x = 1 then
			tmin = timerThis - (thours * 60)
			end if
			
		next
		
		if cint(oRec("stempelurindstilling")) = -1 then
		totalhours = totalhours - cint(thours)
		totalmin = totalmin - tmin
		else
		totalhours = totalhours + cint(thours)
		totalmin = totalmin + tmin
		end if
		
		if len(tmin) <> 0 then
		
			tmin = tmin
			
			if tmin = 0 then
			tmin = "00"
			end if
			
			if len(tmin) = 1 then
			tmin = "0"&tmin
			end if
			
		else
		tmin = "00"
		end if%>
		
		<%if cint(oRec("stempelurindstilling")) = -1 then %>
		-<%=thours%>:<%=left(tmin, 2)%>
		<%else %>
        <b>
		<%=thours%>:<%=left(tmin, 2)%></b>
		<%end if %>
		
		
		</td>
		<%if showTot <> 1 then %>
		
		
		<td align=center class=lille>
		<%if oRec("manuelt_oprettet") <> 0 then %>
		Ja
		<%else %>
		&nbsp;
		<%end if %>
		</td>
		
		<td class=lille>
        <span style="color:#999999;">
		<%select case cint(oRec("manuelt_afsluttet"))
		case 1 
		'** tider manuelt ændret **'%>
		
		Ja:
		
		    <%if len(trim(oRec("login_first"))) <> 0 then %>
			<%="("&left(formatdatetime(oRec("login_first"), 3), 5)%>
			<%end if %>
			
			<%if len(trim(oRec("logud_first"))) <> 0 then %>
			<%="-"&left(formatdatetime(oRec("logud_first"), 3), 5)&")"%>
			<%end if %>
			
		
		<%case 2 '** Glemt at logge ud ***'%>
		Glemt at logge ud!
		
		<%
		case 3 'Tidertilret pga stempelur indstillinger og tider manuelt ændret **'
		%>
		Ja:
		
		    <%if len(trim(oRec("login_first"))) <> 0 then %>
			<%="("&left(formatdatetime(oRec("login_first"), 3), 5)%>
			<%end if %>
			
			<%if len(trim(oRec("logud_first"))) <> 0 then %>
			<%="-"&left(formatdatetime(oRec("logud_first"), 3), 5)&")"%>
			<%end if %>
			
			<br />
			Logind tilpasset pga.<br /> stempelur indstilllinger.
		
		
		<%
		case else %>
		&nbsp;
		<%end select %>

        </span>
		</td>
		
		
		<%end if %>

        <%if showTot <> 1 then%>
        <td class=lille>
		
            <%
            if cint(oRec("stempelurindstilling")) <> -1 then 
            
                 if (layout = 1) AND (ugeerAfsl_og_autogk_smil = 0 OR (level = 1)) AND media <> "print"  then
                %>
		        <textarea id="FM_kommentar" name="FM_kommentar" style="font-size:9px; font-family:arial; width:150px; height:25px;"><%=oRec("kommentar") %></textarea>
                <input type="hidden" name="FM_kommentar" id="Hidden13"  value="#">  
                <%else %>
                <%=left(oRec("kommentar"), 100) %>
                <%end if %>
            &nbsp;
            <%else %>
            &nbsp;
		    <%end if %>
		</td>


        <td align=center>
		
		
		 <% if ((ugeerAfsl_og_autogk_smil = 0 AND showTot <> 1 AND len(oRec("logud")) <> 0) OR (level = 1 AND showTot <> 1)) AND media <> "print" then%>
		  <a href="#" onclick="Javascript:window.open('stempelur.asp?menu=<%=menu%>&func=sletlogind&id=<%=oRec("lid")%>&medarbSel=<%=medarbSel%>&showonlyone=<%=showonlyone%>&hidemenu=<%=hidemenu%>&rdir=<%=rdir%>', '', 'width=400,height=200,resizable=no')" class="vmenu">
              <img src="../ill/slet_16.gif" border="0" /></a>
		 <%else %>
		 &nbsp;
		 <%end if %>      
		</td>


        <%end if 'showtotal %>

        <td>&nbsp;</td>
	 </tr>
	<%
    lastWeek = datepart("ww", oRec("dato"), 2,2)
	lastMnavn = oRec("mnavn")
	lastMnr = oRec("mnr")
	lastMid = oRec("lmid")

    'Response.write "lastMid b:" & lastMid & "<br>"


	x = x + 1
	g = g + 1
	Response.Flush
	oRec.movenext
	wend
	oRec.close 



    '*** Kun fra egen logind historik ***'
    if layout = 1 then

    
    '*** Forvalgt stempelurr ****'
    call fv_stempelur() 
    
    


    if x = 0 then%>

                <%if instr(medIDNavnWrt, "#"&medarbsel&"#") = 0 then
    
    
                call meStamData(medarbsel)%>
	
	            <tr bgcolor="#eff3ff">
		            <td>&nbsp;</td>
		            <td colspan="<%=csp%>" style="height:20px; padding:5px;"><b><%=meTxt %></b></td>
		            <td>&nbsp;</td>
	            </tr>
	
	            <%
    
                medIDNavnWrt = medIDNavnWrt & ",#"& medarbsel & "#" 
                end if
    
    
                '** er periode godkendt ***'
		        tjkDag = dtUse
		        erugeafsluttet = instr(afslUgerMedab(medarbsel), "#"&datepart("ww", tjkDag,2,2)&"_"& datepart("yyyy", tjkDag) &"#")
                
		        
		        'Response.Write "smilaktiv: "& smilaktiv & "<br>"
		        'Response.Write "ugeNrAfsluttet: "& ugeNrAfsluttet & "<br>"
		        'Response.Write "autolukvdatodato: "& autolukvdatodato & "<br>"
		        'Response.Write "tjkDag: "& tjkDag & "<br>"
		        'Response.Write "autolukvdato: "& autolukvdato & "<br>"
		        'Response.Write "erugeafsluttet:" & erugeafsluttet & "<br>"
		        
		        call lonKorsel_lukketPer(tjkDag)
		         
                if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", tjkDag) = year(now) AND DatePart("m", tjkDag) < month(now)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", tjkDag) < year(now) AND DatePart("m", tjkDag) = 12)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", tjkDag) < year(now) AND DatePart("m", tjkDag) <> 12) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", tjkDag) > 1))) OR cint(lonKorsel_lukketIO) = 1 then
              
                ugeerAfsl_og_autogk_smil = 1
                else
                ugeerAfsl_og_autogk_smil = 0
                end if 
                
                
                if (ugeerAfsl_og_autogk_smil = 0 OR (level = 1)) AND media <> "print" then
    
    %>



    <tr>
		<td bgcolor="#cccccc" colspan="<%=csp+2%>"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="#FFFFFF">
        <input type="hidden" value="0" name="id" />
       <input type="hidden" value="<%=usemrn%>" name="mid" />
        <input type="hidden" value="<%=formatdatetime(dtUse, 2) %>" name="logindato" />
        <input type="hidden" value="<%=forvalgt_stempelur %>" name="FM_stur" />
		<td height=20>&nbsp;</td>
		<td align=right><b><%=left(weekdayname(weekday(dtUse)), 4) %>. <%=formatdatetime(dtUse, 2) %></b></td>
        <td align=right><input type="text" class="loginhh" name="FM_login_hh" id="FM_login_hh_<%=d %>" value="" style="width:20px; font-size:9px;" />:
        <input type="text" name="FM_login_mm" id="FM_login_mm_<%=d %>" value="" style="width:20px; font-size:9px;" />
        &nbsp;</td>
         <td align=right><input type="text" class="logudhh" name="FM_logud_hh" id="FM_logud_hh_<%=d %>" value="" style="width:20px; font-size:9px;" />:
         <input type="text" name="FM_logud_mm" id="FM_logud_mm_<%=d %>" value="" style="width:20px; font-size:9px;" /></td>
        <td>&nbsp;</td>
        <td><select name="FM_stur" style="width:140px; font-size:9px; font-family:arial;">
		<%if lto <> "fk_bpm" AND lto <> "kejd_pb" AND lto <> "kejd_pb2"  then %>
        <option value="0">Ingen</option>
        <%end if %>
		
        <%
		strSQL5 = "SELECT id, navn, faktor, minimum, forvalgt FROM stempelur ORDER BY navn"
		oRec5.open strSQL5, oConn, 3 
		while not oRec5.EOF 
		
		if oRec5("forvalgt") = 1 then
		sel = "SELECTED"
		else
		sel = ""
		end if
		%>
		<option value="<%=oRec5("id")%>" <%=sel%>><%=oRec5("navn")%> (<%=oRec5("faktor")%> / <%=oRec5("minimum")%> min.)</option>
		
		<%
		oRec5.movenext
		wend
		oRec5.close 

        %>
        </select>
</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td><textarea id="Textarea1" name="FM_kommentar" style="font-size:9px; font-family:arial; width:150px; height:25px;"></textarea>
        <input type="hidden" name="FM_kommentar" id="Hidden18"  value="#">  </td>
        <td>&nbsp;</td>
		<td>&nbsp;</td>
	 </tr>          

                  <input type="hidden" name="oprFM_login_hh" id="Hidden16" value="00">
		          <input type="hidden" name="oprFM_login_mm" id="Hidden17"  value="00">
                
                 <input type="hidden" name="FM_login_hh" id="Hidden14" value="#">
		          <input type="hidden" name="FM_login_mm" id="Hidden15"  value="#">  


                  <input type="hidden" name="oprFM_logud_hh" id="Hidden9" value="00">
		          <input type="hidden" name="oprFM_logud_mm" id="Hidden10"  value="00">
	    
	              <!-- bruges til arrray split -->
	              <input type="hidden" name="FM_logud_hh" id="Hidden11" value="#">
		          <input type="hidden" name="FM_logud_mm" id="Hidden12"  value="#">  

    <%

    end if '** ugeafsluttet

    end if '** x=0

    end if ' layout


    next




	
	if g = 0 AND layout <> 1 then%>
	<tr>
		<td bgcolor="#cccccc" colspan="<%=csp+2%>"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="#eff3ff">
		<td height=20 style="border-left:1px #8caae6 solid;">&nbsp;</td>
		<td colspan="<%=csp%>"><br><br>Der findes <b>ikke</b> nogen komme/gå tider for de(n) valgte medarbejder(e) i den valgte periode.<br><br>&nbsp;</td>
		<td style="border-right:1px #8caae6 solid;">&nbsp;</td>
	 </tr>
	<%
	else
		'Response.write "lastMnavn: " & lastMnavn & " lastMid:"& lastMid &" lastMnr. "&lastMnr &" totalhours "& totalhours  &" ¤"& totalmin &"datoer:"& sqlDatoStart &","& sqlDatoSlut
        'Response.end
		call tottimer(lastMnavn, lastMnr, totalhours, totalmin, lastMid, sqlDatoStart, sqlDatoSlut, 1)
		totalhours = 0 
		totalmin = 0
	end if%>
	
    <tr>
		<td bgcolor="#cccccc" colspan="<%=csp+2%>"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
    <!--
    <tr>
        <td colspan="<%=csp+2%>" style="padding:20px;" align=right><input type="submit" value=" Opdater >> " /></td>
    </tr>
    -->
	<tr>
		<td bgcolor="#cccccc" colspan="<%=csp+2%>"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="#5582d2">
		<td width="8" valign=top height=20 style="border-bottom:1px #8caae6 solid; border-left:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=<%=csp%> valign="top" style="border-bottom:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>

    <%if media <> "print" AND layout = 1 then %> 
    <tr >
		<td width="8"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=<%=csp%> align=right><br /><input type="submit" value="Opdater liste >>"</td>
		<td><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
    <%end if %>

    </form>
	</table>
	
	<%end function




public showkgtim, showkgpau, showkgtil, showkgtot, showkgnor, showkgsal, showkguds, showkgsaa
function stempelur_kolonne(lto)


    select case lcase(lto)
    case "fk", "fk_bpm", "kejd_pb", "kejd_pb2", "intranet - local"

    showkgtim = 1
    showkgpau = 0
    showkgtil = 0
    showkgtot = 0
    showkgnor = 1
    showkgsal = 1
    showkguds = 0
    showkgsaa = 1
    
    case "dencker", "jttek"

    showkgtim = 1
    showkgpau = 1
    showkgtil = 0
    showkgtot = 1
    showkgnor = 1
    showkgsal = 1
    showkguds = 0
    showkgsaa = 1

    case else

    showkgtim = 1
    showkgpau = 1
    showkgtil = 1
    showkgtot = 1
    showkgnor = 1
    showkgsal = 1
    showkguds = 1
    showkgsaa = 1

    end select


end function





public totaltimerPer, totalpausePer, totalTimerPer100
function fLonTimerPer(stDato, periode, visning, medid)

'Response.Write "////////her " & stDato & " Periode: " & periode & " visning: "& visning & " medid: "& medid & " rdir:"& rdir
'Response.end


totaltimerPer = 0
totalpausePer = 0
ugeIaltFraTilTimer = 0

slutDato = dateadd("d", periode, stDato)
'weekDiff = datediff("ww", stDato, slutDato, 2, 2)

         call meStamdata(medid)

         '** Er ansat dato efter statdato i interval ****'
         if cdate(meAnsatDato) <= cDate(stdato) then
                
                weekDiff = dateDiff("ww", stdato, slutDato, 2, 2)

         else

                weekDiff = dateDiff("ww", meAnsatDato, slutDato, 2, 2)
	
	     end if

        
            if len(weekDiff) <> 0 AND weekDiff <> 0 then
	        weekDiff = cint(weekDiff)
	        else
	        weekDiff = 1
	        end if





if visning = 0 then '** Stempelur faneblad på timereg siden
    call akttyper2009(2)
end if


'*** Finder navne på til/fra typer **'
		                
akttyperTFh = replace(aty_sql_tilwhours, "t.tfaktim = 0", "")
akttyperTF = split(akttyperTFh, " OR t.tfaktim = ")
for tf = 0 to UBOUND(akttyperTF)

'Response.Write akttyperTF(tf) & "<br>"

if len(trim(akttyperTF(tf))) <> 0 then
call akttyper(akttyperTF(tf),0)

if tf = 0 then
akttypenavnTil = akttypenavn
else
akttypenavnTil = akttypenavnTil &", "& akttypenavn
end if
'Response.Write akttyperTF(tf) & "<br>"
end if

next


		                
akttyperTFr = replace(aty_sql_frawhours, "t.tfaktim = 0", "")
akttyperTF = split(akttyperTFr, " OR t.tfaktim = ")
for tf = 0 to UBOUND(akttyperTF)

'Response.Write akttyperTF(tf) & "<br>"

if len(trim(akttyperTF(tf))) <> 0 then
call akttyper(akttyperTF(tf),0)

if tf = 0 then
akttypenavnFra = akttypenavn
else
akttypenavnFra = akttypenavnFra & ", "& akttypenavn
end if
'Response.Write akttyperTF(tf) & "<br>"
end if

next

'*******


'Response.Write stDato & ", "& periode & "meid: " & medid & " weekdiff: "& weekdiff

	'**** login historik (denne uge/ Periode) ****
	for intcounter = 0 to periode
	
					select case intcounter
					case 0
					useSQLd = stDato
					case else
					tDat = dateadd("d", 1, useSQLd)
					useSQLd = year(tDat) &"/"& month(tDat) & "/"& day(tDat)
					end select
					
					
					
					strSQL = "SELECT l.id AS lid, l.mid AS lmid, l.login, l.logud, l.minutter, "_
					&" s.navn AS stempelurnavn, s.faktor, s.minimum, stempelurindstilling FROM login_historik l"_
					&" LEFT JOIN stempelur s ON (s.id = l.stempelurindstilling) WHERE "_
					&" l.dato = '"& useSQLd &"' AND l.mid = " & medid &""_
					&" ORDER BY l.login" 
					
					'Response.Write strSQL & "<br><br>"
					
					
					f = 0
					oRec.open strSQL, oConn, 3 
					while not oRec.EOF 
					
						timerThis = 0
						timerThisDIFF = 0
						
						if len(oRec("login")) <> 0 AND len(oRec("logud")) <> 0 then
						
						'loginTidAfr = left(formatdatetime(oRec("login"), 3), 5)
						'logudTidAfr = left(formatdatetime(oRec("logud"), 3), 5)
						
						if cint(oRec("stempelurindstilling")) = -1 then
						    
						    timerThisDIFF = oRec("minutter")
						    useFaktor = 0
						    
						    timerThisPause = timerThisDIFF
						    timerThis = 0
						    
						else 
						    
						    timerThisDIFF = oRec("minutter") 'datediff("s", loginTidAfr, logudTidAfr)/60
						
						    if timerThisDIFF < oRec("minimum") then
							    timerThisDIFF = oRec("minimum")
						    end if
						
							
							if oRec("faktor") > 0 then
							useFaktor = oRec("faktor")
							else
							useFaktor = 0
							end if
							
							timerThisPause = 0
							timerThis = (timerThisDIFF * useFaktor)
							
						end if
						
						
						
						totaltimerPer = totaltimerPer + timerThis
						totalpausePer = totalpausePer + timerThisPause
						'Response.write oRec("lid") & ": " &  timerThis &" - "
						else
						
						totaltimerPer = totaltimerPer
						totalpausePer = totalpausePer 
						
						end if
						
						
						'timTemp = formatnumber(timerThis/60, 3)
						'timTemp_komma = split(timTemp, ",")
						
						'for f = 0 to UBOUND(timTemp_komma)
							
						'	if f = 0 then
						'	thours = timTemp_komma(f)
						'	end if
							
						'	if f = 1 then
						'	tmin = timerThis - (thours * 60)
						'	end if
							
						'next
						
						if visning = 0 OR visning = 20 then
					    
						select case intcounter
						case 0
						manMin = manMin + timerThis 
						manMinPause = manMinPause + timerThisPause 
						case 1
						tirMin = tirMin + timerThis 
						tirMinPause = tirMinPause + timerThisPause  
						case 2
						onsMin = onsMin + timerThis
						onsMinPause = onsMinPause + timerThisPause   
						case 3
						torMin = tormin + timerThis
						torMinPause = torminPause + timerThisPause   
						case 4
						freMin = freMin + timerThis
						freMinPause = freMinPause + timerThisPause   
						case 5
						lorMin = lorMin + timerThis
						lorMinPause = lorMinPause + timerThisPause  
						case 6
						sonMin = sonMin + timerThis
						sonMinPause = sonMinPause + timerThisPause   
						end select 
						
						
						end if
					    
					    'Response.write "tot:"& intcounter &" - "& totaltimer &":"& totalpause
						
						
					oRec.movenext
					wend
					oRec.close 
					
					f = 0
					
					
					
					    '*** Tillæg / Fradrag via Realtimer **'
					    if visning = 0 OR visning = 20 then
                	
		                
                		
		                tiltimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS tiltimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND ("& aty_sql_tilwhours &")) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                tiltimer = oRec2("tiltimer") * 60
		                end if
		                oRec2.close 
                		
		                if len(trim(tiltimer)) <> 0 then
		                tiltimer = tiltimer
		                else
		                tiltimer = 0
		                end if
                		
                		 
                		
                		
                		fradtimer = 0
		                strSQL2 = "SELECT sum(timer) AS fratimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND ("& aty_sql_frawhours &")) GROUP BY tmnr "
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                fradtimer = oRec2("fratimer") * 60
		                end if
		                oRec2.close 
                		
		                if len(trim(fradtimer)) <> 0 then
		                fradtimer = fradtimer
		                else
		                fradtimer = 0
		                end if
		                
		                
		                select case intcounter
						case 0
						manFraTimer = (tiltimer - fradtimer)
						case 1
						tirFraTimer = (tiltimer - fradtimer)
						case 2
						onsFraTimer = (tiltimer - fradtimer)
						case 3
						torFraTimer = (tiltimer - fradtimer)
						case 4
						freFraTimer = (tiltimer - fradtimer)
						case 5
						lorFraTimer = (tiltimer - fradtimer)
						case 6
						sonFraTimer = (tiltimer - fradtimer)
						end select 
						
		                
		                
		                '*** Fleks ***'
                		call akttyper2009prop(7)
                		aty_fleks_on = aty_on
                		aty_fleks_tf = aty_tfval
                		if cint(aty_fleks_on) = 1 then
		                '*** Fleks ****'   
                	    flekstimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS flekstimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 7) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                flekstimer = oRec2("flekstimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(flekstimer)) <> 0 then
		                flekstimer = flekstimer
		                else
		                flekstimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manflekstimer = (flekstimer)
						case 1
						tirflekstimer = (flekstimer)
						case 2
						onsflekstimer = (flekstimer)
						case 3
						torflekstimer = (flekstimer)
						case 4
						freflekstimer = (flekstimer)
						case 5
						lorflekstimer = (flekstimer)
						case 6
						sonflekstimer = (flekstimer)
						end select 
						
						
						end if
						
						
						'*** Ferie ***'
                		call akttyper2009prop(14)
                		aty_Ferie_on = aty_on
                		aty_Ferie_tf = aty_tfval
                		if cint(aty_Ferie_on) = 1 then
		                '*** Ferie ****'   
                	    Ferietimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS Ferietimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 14) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                Ferietimer = oRec2("Ferietimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(Ferietimer)) <> 0 then
		                Ferietimer = Ferietimer
		                else
		                Ferietimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manFerietimer = (Ferietimer)
						case 1
						tirFerietimer = (Ferietimer)
						case 2
						onsFerietimer = (Ferietimer)
						case 3
						torFerietimer = (Ferietimer)
						case 4
						freFerietimer = (Ferietimer)
						case 5
						lorFerietimer = (Ferietimer)
						case 6
						sonFerietimer = (Ferietimer)
						end select 
						
						
						end if
						
						
						
						'*** Syg ***'
                		call akttyper2009prop(20)
                		aty_Syg_on = aty_on
                		aty_Syg_tf = aty_tfval
                		if cint(aty_Syg_on) = 1 then
		                '*** Syg ****'   
                	    Sygtimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS Sygtimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 20) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                Sygtimer = oRec2("Sygtimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(Sygtimer)) <> 0 then
		                Sygtimer = Sygtimer
		                else
		                Sygtimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manSygtimer = (Sygtimer)
						case 1
						tirSygtimer = (Sygtimer)
						case 2
						onsSygtimer = (Sygtimer)
						case 3
						torSygtimer = (Sygtimer)
						case 4
						freSygtimer = (Sygtimer)
						case 5
						lorSygtimer = (Sygtimer)
						case 6
						sonSygtimer = (Sygtimer)
						end select 
						
						
						end if
						
						
						'*** BarnSyg ***'
                		call akttyper2009prop(21)
                		aty_BarnSyg_on = aty_on
                		aty_BarnSyg_tf = aty_tfval
                		if cint(aty_BarnSyg_on) = 1 then
		                '*** BarnSyg ****'   
                	    BarnSygtimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS BarnSygtimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 21) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                BarnSygtimer = oRec2("BarnSygtimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(BarnSygtimer)) <> 0 then
		                BarnSygtimer = BarnSygtimer
		                else
		                BarnSygtimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manBarnSygtimer = (BarnSygtimer)
						case 1
						tirBarnSygtimer = (BarnSygtimer)
						case 2
						onsBarnSygtimer = (BarnSygtimer)
						case 3
						torBarnSygtimer = (BarnSygtimer)
						case 4
						freBarnSygtimer = (BarnSygtimer)
						case 5
						lorBarnSygtimer = (BarnSygtimer)
						case 6
						sonBarnSygtimer = (BarnSygtimer)
						end select 
						
						
						end if
						
						
						'*** Lage ***'
                		call akttyper2009prop(81)
                		aty_Lage_on = aty_on
                		aty_Lage_tf = aty_tfval
                		if cint(aty_Lage_on) = 1 then
		                '*** Lage ****'   
                	    Lagetimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS Lagetimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 81) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                Lagetimer = oRec2("Lagetimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(Lagetimer)) <> 0 then
		                Lagetimer = Lagetimer
		                else
		                Lagetimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manLagetimer = (Lagetimer)
						case 1
						tirLagetimer = (Lagetimer)
						case 2
						onsLagetimer = (Lagetimer)
						case 3
						torLagetimer = (Lagetimer)
						case 4
						freLagetimer = (Lagetimer)
						case 5
						lorLagetimer = (Lagetimer)
						case 6
						sonLagetimer = (Lagetimer)
						end select 
						
						
						end if
						
						
						'*** Sund ***'
                		call akttyper2009prop(8)
                		aty_Sund_on = aty_on
                		aty_Sund_tf = aty_tfval
                		if cint(aty_Sund_on) = 1 then
		                '*** Sund ****'   
                	    Sundtimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS Sundtimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 8) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                Sundtimer = oRec2("Sundtimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(Sundtimer)) <> 0 then
		                Sundtimer = Sundtimer
		                else
		                Sundtimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manSundtimer = (Sundtimer)
						case 1
						tirSundtimer = (Sundtimer)
						case 2
						onsSundtimer = (Sundtimer)
						case 3
						torSundtimer = (Sundtimer)
						case 4
						freSundtimer = (Sundtimer)
						case 5
						lorSundtimer = (Sundtimer)
						case 6
						sonSundtimer = (Sundtimer)
						end select 
						
						
						end if
                	    
                	    
                	    '*** Frokost ***'
                		call akttyper2009prop(10)
                		aty_Frokost_on = aty_on
                		aty_Frokost_tf = aty_tfval
                		if cint(aty_Frokost_on) = 1 then
		                '*** Frokost ****'   
                	    Frokosttimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS Frokosttimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 10) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                Frokosttimer = oRec2("Frokosttimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(Frokosttimer)) <> 0 then
		                Frokosttimer = Frokosttimer
		                else
		                Frokosttimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manFrokosttimer = (Frokosttimer)
						case 1
						tirFrokosttimer = (Frokosttimer)
						case 2
						onsFrokosttimer = (Frokosttimer)
						case 3
						torFrokosttimer = (Frokosttimer)
						case 4
						freFrokosttimer = (Frokosttimer)
						case 5
						lorFrokosttimer = (Frokosttimer)
						case 6
						sonFrokosttimer = (Frokosttimer)
						end select 
						
						
						end if
	              
	                end if 'visning fradrag / tillæg
					
					
					
	
	next
	
	

	
	call stempelur_kolonne(lto)
	
	
	select case visning 
	case 0, 20%>


	<%if visning <> 20 then %>
	
	<h4><img src="../ill/ikon_stempelur_24.png" alt="" border="0">&nbsp; <%=tsa_txt_340 %>&nbsp;- 
		<%=tsa_txt_005 %>: <%=datepart("ww", stDato, 2, 2)%></h4>
		
		<!-- Denne uge, nuværende login -->
		<%if datepart("ww", stDato, 2, 2) =  datepart("ww", now, 2, 2) then%>
		<%=tsa_txt_134 %>: 
		<%
		
		sLoginTid = "00:00:00"
		
		strSQL = "SELECT l.id AS lid, l.login "_
		&" FROM login_historik l WHERE "_
		&" l.mid = " & medid &" AND stempelurindstilling <> -1"_
		&" ORDER BY l.id DESC" 
					
		'Response.write strSQL
		'Response.flush
		
		
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then 
        
        sLoginTid = oRec("login") 
        
        end if
        oRec.close
		
		%>
		<b><%=formatdatetime(sLoginTid, 3) %></b>
		
		<% 
		logindiffSidste = datediff("n", sLoginTid, now, 2, 2) 
		%>
		
		<br /><%=tsa_txt_135 %>
		<%call timerogminutberegning(logindiffSidste) %>
		<b><%=thoursTot &":"& left(tminTot, 2) %>&nbsp;t.</b>
		
		
		<%end if '*** Denne uge / nuværende login **' %>
		
	<table cellspacing=1 cellpadding=2 border=0 width=100% bgcolor="#c4c4c4">
    
    

	<tr bgcolor="#FFFFFF">
		<td style="width:100px;">
            &nbsp;</td>
		
		<td width=50 valign=bottom align=center class=lille ><b><%=tsa_txt_128 %> d. <%=formatdatetime(stDato, 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_129 %> d. <%=formatdatetime(dateadd("d", 1, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_130 %> d. <%=formatdatetime(dateadd("d", 2, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_131 %> d. <%=formatdatetime(dateadd("d", 3, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_132 %> d. <%=formatdatetime(dateadd("d", 4, stDato), 2) %></b></td>
	    <td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_133 %> d. <%=formatdatetime(dateadd("d", 5, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_127 %> d. <%=formatdatetime(dateadd("d", 6, stDato), 2) %></b></td>
		<td width=50 bgcolor="#ffdfdf" class=lille valign=bottom align=right><b><%=global_txt_167 %></b></td>
	</tr>
	<%
    cspsStur = 1
    else 
    cspsStur = 2%>

    <tr bgcolor="#FFFFFF"><td colspan=10 style="border-bottom:1px #CCCCCC solid;"><br /><br /><b><%=tsa_txt_340 %></b> (akkumuleret)</td></tr>

    <%end if 'visning %>

    <%if cint(showkgtim) = 1 then %>
    <tr bgcolor="#ffffff">
		<td colspan="<%=cspsStur %>"><%=tsa_txt_137 %>:</td>
		<td valign=top align=right><%call timerogminutberegning(manMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt + manMin/1%>
		<td valign=top align=right><%call timerogminutberegning(tirMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt + tirMin/1%>
		<td valign=top align=right><%call timerogminutberegning(onsMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt + onsMin/1%>
		<td valign=top align=right><%call timerogminutberegning(torMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt +  torMin/1%>
		<td valign=top align=right><%call timerogminutberegning(freMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt +  freMin/1%>
		<td valign=top align=right><%call timerogminutberegning(lorMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt +  lorMin/1%>
		<td valign=top align=right><%call timerogminutberegning(sonMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt +  sonMin/1%>
		<td valign=top align=right>= 
		<%call timerogminutberegning(ugeIalt)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
    <%end if %>


    <%if cint(showkgpau) = 1 then %>
	<tr bgcolor="lightgrey">
		<td colspan="<%=cspsStur %>"><%=tsa_txt_138 %>:</td>
		<td valign=top align=right><%call timerogminutberegning(-manMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause + manMinPause%>
		<td valign=top align=right><%call timerogminutberegning(-tirMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause + tirMinPause%>
		<td valign=top align=right><%call timerogminutberegning(-onsMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause + onsMinPause%>
		<td valign=top align=right><%call timerogminutberegning(-torMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause + torMinPause%>
		<td valign=top align=right><%call timerogminutberegning(-freMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause + freMinPause%>
		<td valign=top align=right><%call timerogminutberegning(-lorMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause + lorMinPause%>
		<td valign=top align=right><%call timerogminutberegning(-sonMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause + sonMinPause%>
		<td valign=top align=right>= <%call timerogminutberegning(-ugeIaltPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	<%end if %>

	<!-- Fradrag / Tillæg via Realtimer -->
	
    <%if cint(showkgtil) = 1 then %>
    <tr bgcolor="#ffffff">
		<td colspan="<%=cspsStur %>"><%=global_txt_168 %>:*</td>
		<td valign=top align=right><%call timerogminutberegning(manFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (manFraTimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (tirFraTimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (onsFraTimer)%>
		<td valign=top align=right><%call timerogminutberegning(torFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (torFraTimer)%>
		<td valign=top align=right><%call timerogminutberegning(freFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (freFraTimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (lorFraTimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (sonFraTimer)%>
		<td valign=top align=right>= <%call timerogminutberegning(ugeIaltFraTilTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	<%end if %>

    <%
	totMan = manMin - (manMinPause - (manFraTimer))
	totTir = tirMin - (tirMinPause - (tirFraTimer))
	totOns = onsMin - (onsMinPause - (onsFraTimer))
	totTor = torMin - (torMinPause - (torFraTimer))
	totFre = freMin - (freMinPause - (freFraTimer))
	totLor = lorMin - (lorMinPause - (lorFraTimer))
	totSon = sonMin - (sonMinPause - (sonFraTimer))
	%>


    <%if cint(showkgtot) = 1 then %>
	<!-- total -->
	
	
	 <tr bgcolor="#ffdfdf">
		<td colspan="<%=cspsStur %>"><b><%=global_txt_167%>:</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totMan)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totTir)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totOns)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totTor)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</b></td>
	    <td valign=top align=right><b><%call timerogminutberegning(totFre)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totLor)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totSon)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right>= <b><%call timerogminutberegning(ugeIalt - (ugeIaltPause - (ugeIaltFraTilTimer)))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<%loginTimerTot = ugeIalt - (ugeIaltPause - ugeIaltFraTilTimer)  %>
	</tr>
	<%end if %>
 
	<!-- Normtimer -->
	
    <%if cint(showkgnor) = 1 then %>
    <tr bgcolor="#ffffff">
		<td colspan="<%=cspsStur %>"><%=tsa_txt_259 %>:</td>
		
		<%call normtimerper(medid, varTjDatoUS_man, 6) %>
		
		<td valign=top align=right><%call timerogminutberegning(ntimMan*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		
		<td valign=top align=right><%call timerogminutberegning(ntimTir*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		
		<td valign=top align=right><%call timerogminutberegning(ntimOns*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		
		<td valign=top align=right><%call timerogminutberegning(ntimTor*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		
		<td valign=top align=right><%call timerogminutberegning(ntimFre*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		
		<td valign=top align=right><%call timerogminutberegning(ntimLor*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		
		
		<td valign=top align=right><%call timerogminutberegning(ntimSon*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		
		<%
		NormTimerWeekTot = 0
		NormTimerWeekTot = (ntimMan + ntimTir + ntimOns + ntimTor + ntimFre + ntimLor + ntimSon) * 60 %>
		
		<td valign=top align=right>= <%call timerogminutberegning(NormTimerWeekTot)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</td>
		</tr>

      <%end if %>
	    
	  <!-- Saldo -->
      
      <%if cint(showkgsal) = 1 then %>  
	  <tr bgcolor="#DCF5BD">
		<td colspan="<%=cspsStur %>" style="height:20px;"><b><%=global_txt_163 %>:</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totMan - (ntimMan*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totTir - (ntimTir*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totOns - (ntimOns*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totTor - (ntimTor*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</b></td>
	    <td valign=top align=right><b><%call timerogminutberegning(totFre - (ntimFre*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totLor - (ntimLor*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totSon - (ntimSon*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right>=
        <%
        call timerogminutberegning((ugeIalt - (ugeIaltPause - (ugeIaltFraTilTimer))) - NormTimerWeekTot)
        %>

        <%
        stDatoUS = year(stDato) & "/" & month(stDato) &"/"& day(stDato)  
        slDatoUS = year(slutDato) & "/" & month(slutDato) &"/"& day(slutDato)
        %>
        <a href="afstem_tot.asp?usemrn=<%=usemrn%>&show=5&varTjDatoUS_man=<%=stDatoUS%>&varTjDatoUS_son=<%=slDatoUS%>" class=vmenu><%=thoursTot &":"& left(tminTot, 2)%> t</a></td>
		<%loginTimerTot = ugeIalt - (ugeIaltPause - ugeIaltFraTilTimer)  %>
	</tr>
    <%end if %>


    <%if visning <> 20 then '** Ikke fra timereg. siden da det er for tungt  %>
        <%if cint(showkgsaa) = 1 then '*** IKKE aktiv - se aftamte tot istedet for
        
        
        %>


        <%end if %>
    <%end if %>


    <%if visning <> 20 then %>
	</table>
	
    

	<%if datepart("ww", stDato, 2, 2) =  datepart("ww", now, 2, 2) then%>
	
	<%=tsa_txt_139 %>: <% 
	call timerogminutberegning(logindiffSidste+(loginTimerTot))
	%>
	<b><%=thoursTot &":"& left(tminTot, 2) %>&nbsp;t.</b>
	
    <%end if ' *** Denne uge / Nuværende login **'%>
	

    
	<br /><br />
	<table><tr><td >
	
	</td></tr></table>

    <%
	if cint(showkguds) = 1 then

	Response.Write "*) <b> Tillægs typer:</b> " & akttypenavnTil
	Response.Write "<br><b>Fradrags typer:</b> " & akttypenavnFra
	
	%>

	
    <%if media <> "print" then %>    
	<br /><br /><a href="#" id="udspec" class="vmenu">+ Udspecificering</a> (fraværs typer)
    <%end if 
    
    
    end if%>
    
    
    <div id="udspecdiv" style="position:relative; width:600px; visibility:hidden; display:none;">
	<table cellspacing=1 cellpadding=2 border=0 width=100% bgcolor="#c4c4c4">    
	    <tr bgcolor="#FFFFFF">
	    <td colspan=9><br /><b>Udspecificering på fraværstyper</b> <br />
	    Ikke medregnet i saldo, med mindre de er en del af <%=global_txt_168 %> typerne*.</td>
	</tr>
	<tr bgcolor="#d6dff5">
		<td style="width:100px;">
            &nbsp;</td>
		
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_128 %> d. <%=formatdatetime(stDato, 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_129 %> d. <%=formatdatetime(dateadd("d", 1, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_130 %> d. <%=formatdatetime(dateadd("d", 2, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_131 %> d. <%=formatdatetime(dateadd("d", 3, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_132 %> d. <%=formatdatetime(dateadd("d", 4, stDato), 2) %></b></td>
	    <td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_133 %> d. <%=formatdatetime(dateadd("d", 5, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_127 %> d. <%=formatdatetime(dateadd("d", 6, stDato), 2) %></b></td>
		<td width=50 bgcolor="#ffdfdf" class=lille valign=bottom align=right><b><%=global_txt_167 %></b></td>
	</tr>
	
	    
	    <%if cint(aty_fleks_on) = 1 then %>
	<!-- Fleks Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_147 &" "& aty_fleks_tf%></td>
		<td valign=top align=right><%call timerogminutberegning(manFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (manFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (tirFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (onsFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(torFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (torFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(freFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (freFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (lorFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (sonFlekstimer)%>
		<td valign=top align=right>= <%call timerogminutberegning(ugeIaltFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>
	
	
	
	<%if cint(aty_Ferie_on) = 1 then %>
	<!-- Ferie Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_135 &" "& aty_Ferie_tf %></td>
		<td valign=top align=right><%call timerogminutberegning(manFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (manFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (tirFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (onsFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(torFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (torFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(freFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (freFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (lorFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (sonFerietimer)%>
		<td valign=top align=right>= <%call timerogminutberegning(ugeIaltFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>
	
	<%if cint(aty_Syg_on) = 1 then %>
	<!-- Syg Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_138 &" "& aty_Syg_tf %></td>
		<td valign=top align=right><%call timerogminutberegning(manSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (manSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (tirSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (onsSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(torSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (torSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(freSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (freSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (lorSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (sonSygtimer)%>
		<td valign=top align=right>= <%call timerogminutberegning(ugeIaltSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>
	
	
	<%if cint(aty_BarnSyg_on) = 1 then %>
	<!-- BarnSyg Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_139 &" "& aty_BarnSyg_tf %></td>
		<td valign=top align=right><%call timerogminutberegning(manBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (manBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (tirBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (onsBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(torBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (torBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(freBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (freBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (lorBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (sonBarnSygtimer)%>
		<td valign=top align=right>= <%call timerogminutberegning(ugeIaltBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>
	
	
	<%if cint(aty_Lage_on) = 1 then %>
	<!-- Lage Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_160 &" "& aty_Lage_tf %></td>
		<td valign=top align=right><%call timerogminutberegning(manLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (manLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (tirLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (onsLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(torLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (torLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(freLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (freLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (lorLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (sonLagetimer)%>
		<td valign=top align=right>= <%call timerogminutberegning(ugeIaltLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>
	
		<%if cint(aty_Sund_on) = 1 then %>
	<!-- Sund Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_148 &" "& aty_Sund_tf %></td>
		<td valign=top align=right><%call timerogminutberegning(manSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (manSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (tirSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (onsSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(torSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (torSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(freSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (freSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (lorSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (sonSundtimer)%>
		<td valign=top align=right>= <%call timerogminutberegning(ugeIaltSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>    
	
	<%if cint(aty_Frokost_on) = 1 then %>
	<!-- Frokost Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_133 &" "& aty_Frokost_tf %></td>
		<td valign=top align=right><%call timerogminutberegning(manFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (manFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (tirFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (onsFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(torFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (torFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(freFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (freFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (lorFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (sonFrokosttimer)%>
		<td valign=top align=right>= <%call timerogminutberegning(ugeIaltFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>    
	    
	    
	</table>

   
	</div>
	<%

    end if 'vis <> 3
	
	case 2
	
	totalTimerPer100 = 0
	totalTimerPer100 = (totaltimerPer-totalpausePer)
	
    call timerogminutberegning(totalTimerPer100) %>
	<!-- formatnumber(totalTimerPer100, 2) -->
	
	<%=thoursTot &":"& left(tminTot, 2) %>
	
	<%
	case 3
	
	totalTimerPer100 = 0
	totalTimerPer100 = (totaltimerPer-totalpausePer)
	
	case 4 ''** Avg. Week ***' 
	
	if weekDiff <> 0 then
	weekDiff = weekDiff
	else
	weekDiff = 1
	end if
	
	totalTimerPer100 = 0
	totalTimerPer100 = (totaltimerPer-totalpausePer)/weekDiff
	
	
	call timerogminutberegning(((totaltimerPer-totalpausePer)/weekDiff)) %>
	<!-- formatnumber(totalTimerPer100, 2) -->
	
	<%=thoursTot &":"& left(tminTot, 2) %>
	<%
	case else
	
	
	totalTimerPer100 = 0
	totalTimerPer100 = (totaltimerPer-totalpausePer)
	
	
	call timerogminutberegning(totaltimerPer-totalpausePer) %>
	<!-- formatnumber(totalTimerPer100, 2) -->
	
	<%=thoursTot &":"& left(tminTot, 2) %> 
	
	<%
	end select '** Visinng ***'%>
	
	

<%
end function

function tilfojPauser(io,psMid,psDato,psloginTidp,pslogudTidp,psMin,psKomm)

                            
                            
                            'Response.Write "psloginTidp: " & psloginTidp & "<br>"

                            '*** Indlæser ikke pauser på 0 min.
                            if psMin <> 0 then
                            
                            strSQLp = "INSERT INTO login_historik SET dato = '"& psDato &"', "_
	                        &" login = '"& psloginTidp &"', "_
	                        &" logud = '"& pslogudTidp &"', "_
					        &" stempelurindstilling = -1, minutter = "& psMin &", "_
					        &" manuelt_afsluttet = 0, kommentar = '"& psKomm &"', mid = "& psMid
					        
					        'Response.Write strSQLp & "<br>"
					        oConn.execute(strSQLp)

                            end if

end function


public stPauseLic_1, stPauseLic_2, p1on, p2on, p1_grp, p2_grp
function stPauserFralicens(psDt)


            dagidag = weekday(psDt, 2)
            p1on = 0
            p2on = 0
            
            select case dagidag
            case 1
            feltNavne = "p1_man AS p1on, p2_man AS p2on"
            case 2
            feltNavne = "p1_tir AS p1on, p2_tir AS p2on"
            case 3
            feltNavne = "p1_ons AS p1on, p2_ons AS p2on"
            case 4
            feltNavne = "p1_tor AS p1on, p2_tor AS p2on"
            case 5
            feltNavne = "p1_fre AS p1on, p2_fre AS p2on"
            case 6
            feltNavne = "p1_lor AS p1on, p2_lor AS p2on"
            case 7
            feltNavne = "p1_son AS p1on, p2_son AS p2on"
            end select
        
            '*** Henter forvalgte standard pauser ***

            p1_grp = 0 
            p2_grp = 0

            strSQLstp = "SELECT stpause, stpause2, "& feltNavne &", p1_grp, p2_grp"_
            &" FROM  "_
            &" licens WHERE id = 1" 
            
            oRec5.open strSQLstp, oConn, 3
            if not oRec5.EOF then
            
            stPauseLic_1 = oRec5("stpause")
            stPauseLic_2 = oRec5("stpause2")
            
            p1on = oRec5("p1on")
            p2on = oRec5("p2on")

            p1_grp = oRec5("p1_grp")
            p2_grp = oRec5("p2_grp")
             
            end if
            oRec5.close
        


end function


public p1_prg_on, p2_prg_on
function stPauserProgrp(useMid, p1_grp, p2_grp)

        p1_prg_on = 0
        p2_prg_on = 0

        strSQLmansat = " mansat <> 0"


        '*** Adgang til Pause 1 ****'
        if len(trim(p1_grp)) = 0 OR isNull(p1_grp) = true then
        p1_prg_on = 1
        else
            
            p1_grpArr = split(p1_grp, ",")
            for g = 0 to UBOUND(p1_grpArr)
                
                'Response.Write "p1_grpArr(g)"& p1_grpArr(g) & " useMid: "&  useMid&"<br>"
                call erdetint(p1_grpArr(g))
                if isInt = 0 then
                call medarbiprojgrp(p1_grpArr(g), useMid, 0, -1)
                end if
                isInt = 0
                
            next

                if instr(instrMedidProgrp, "#"& useMid &"#,") <> 0 then
                p1_prg_on = 1
                else
                p1_prg_on = 0
                end if

        end if

       
         '*** Adgang til Pause 2 ****'
        if len(trim(p2_grp)) = 0 OR isNull(p2_grp) = true then
        p2_prg_on = 1
        else
            
            p2_grpArr = split(p2_grp, ",")
            for g = 0 to UBOUND(p2_grpArr)
                
                'Response.Write "p1_grpArr(g)"& p1_grpArr(g) & " useMid: "&  useMid&"<br>"
                call erdetint(p2_grpArr(g))
                if isInt = 0 then
                call medarbiprojgrp(p2_grpArr(g), useMid, 0, -1)
                end if
                isInt = 0
            next

                if instr(instrMedidProgrp, "#"& useMid &"#,") <> 0 then
                p2_prg_on = 1
                else
                p2_prg_on = 0
                end if

        end if




end function

   
   %>