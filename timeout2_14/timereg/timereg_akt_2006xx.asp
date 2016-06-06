  

    <%
    
    'if request("usegl2006") = "0" then
    'Response.Cookies("tsa")("usegl2006") = "0"
    'else
    '    if request.Cookies("tsa")("usegl2006") = "1" AND lto <> "dencker" then
    '    Response.Redirect "../timereg2006/timereg_2006_fs.asp?usegl2006=1"
    '    else
    '    Response.Cookies("tsa")("usegl2006") = "0"
    '    end if
    'end if
    
    'Response.Write "request.Cookies(tsa)(usegl2006) " & request.Cookies("tsa")("usegl2006")
    %>
    
    
    <!--#include file="../inc/connection/conn_db_inc.asp"-->
    <!--#include file="../inc/xml/tsa_xml_inc.asp"-->
	<!--#include file="../inc/errors/error_inc.asp"-->
	<!--#include file="../inc/regular/global_func.asp"-->
	<!--#include file="inc/convertDate.asp"-->
	<!--#include file="inc/timereg_akt_2006_inc.asp"-->
	<!--#include file="inc/smiley_inc.asp"-->
	<!--#include file="inc/isint_func.asp"-->
    <!--#include file="../inc/regular/treg_func.asp"-->
	<!--#include file="../inc/regular/nocasche.asp"-->
	
	
	<%

    
	
	
	   '**** Søgekriterier AJAX **'
        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
        case "FN_fjeasyregakt"
                

               

                if len(trim(request.form("cust"))) <> 0 then
                 thisAktid = request.form("cust")
                 else
                 thisAktid = 0
                 end if

                strSQL = "UPDATE aktiviteter SET easyreg = 0 WHERE id = "& thisAktid
                oConn.execute(strSQL)

        case "FN_notify_jobluk"
                

                
                '*************************************************************************'
	            '*** Lukker job (afmelder arbejde på job er færdig) og afsender email ****'
	            '*************************************************************************'
	            if len(trim(request.form("jq_lukjob"))) <> 0 then
            	
		            'strSQL = "UPDATE job SET jobstatus = 0 WHERE id = "& jobid  'jobnr = "& intJobnr
		            'oConn.execute(strSQL)
            			
                        '***
                        medarbejderid = request("vlgtmedarb")
                        lukjob = request.form("jq_lukjob") 'split(request("FM_lukjob"), ",")

                        'for j = 0 to UBOUND(lukjob)
            				
                            'if lukjob(j) <> 0 then

				            '**** Finder jobansvarlige *****
				            strSQL = "SELECT job.id AS jid, jobnavn, jobnr, jobans1, jobans2, m1.mnavn AS m1mnavn, m1.email AS m1email,"_
				            &" m2.mnavn AS m2mnavn, m2.email AS m2email FROM job "_
				            &" LEFT JOIN medarbejdere m1 ON (m1.mid = jobans1)"_
				            &" LEFT JOIN medarbejdere m2 ON (m2.mid = jobans2)"_
				            &" WHERE job.id = "& lukjob

                            'Response.Write strSQL
                            'Response.flush

				            oRec.open strSQL, oConn, 3
				            x = 0
				            if not oRec.EOF then
            				
				            jobid = oRec("jid")
				            jobans1 = oRec("m1mnavn")
				            jobans2 = oRec("m2mnavn")
				            jobans1email = oRec("m1email")
				            jobans2email = oRec("m2email")
				            jobnavnThis = oRec("jobnavn")
                            intJobnr = oRec("jobnr")
            				
				            x = x + 1
				            end if
				            oRec.close
            				
            				
				            '*** Henter afsender **
				            strSQL = "SELECT mnavn, email FROM medarbejdere"_
				            &" WHERE mid = "& medarbejderid
				            oRec.open strSQL, oConn, 3
            				
				            if not oRec.EOF then
            				
				            afsNavn = oRec("mnavn")
				            afsEmail = oRec("email")
            				
				            end if
				            oRec.close
            				
            				
            					
            						
            					
				            if x <> 0 AND request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\timereg_akt_2006.asp" then
					  	            'Sender notifikations mail
		                            Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                            ' Sætter Charsettet til ISO-8859-1
		                            Mailer.CharSet = 2
		                            Mailer.FromName = "TimeOut | " & afsNavn 
		                            Mailer.FromAddress = afsEmail
		                            Mailer.RemoteHost = "webmail.abusiness.dk" '"pasmtp.tele.dk"
						            Mailer.AddRecipient "" & jobans1 & "", "" & jobans1email & ""
		                            Mailer.AddRecipient "" & jobans2 & "", "" & jobans2email & ""
            						
                                    if lto = "synergi1" then
                                    Mailer.AddRecipient "Bettina Schou", "bs@synergi1.dk"
                                    Mailer.AddRecipient "TimeOut Support", "support@outzource.dk"
                                    end if

                                    if lto = "outz" then
                                    Mailer.AddRecipient "TimeOut Support", "support@outzource.dk"
                                    end if

            						    
						            Mailer.Subject = jobnavnThis &" ("& intJobnr &") - Afsluttet."
		                            strBody = "Hej jobansvarlige." & vbCrLf & vbCrLf
                                    strBody = strBody &"Mit arbejde på: "& vbCrLf & vbCrLf
						            strBody = strBody &"Job: "& jobnavnThis &" ("& intJobnr &") er nu afsluttet.  "& vbCrLf & vbCrLf
		                            strBody = strBody &"Med venlig hilsen" & vbCrLf
		                            strBody = strBody & session("user") & vbCrLf & vbCrLf
            		                
            		
            		
		                            Mailer.BodyText = strBody
            		
		                            Mailer.sendmail()
		                            Set Mailer = Nothing

                                    Response.Write "mail afsendt - " & request.form("jq_lukjob")
				            
                            end if' x
            	
	                    'if request("FM_kopierjob") = "1" then
	                    
	                    ''Response.Write "kopier job"
	                    ''Response.end
	                    'Response.Redirect "job_kopier.asp?func=kopier&id="&jobid&"&fm_kunde=0&filt=aaben&showaspopup=y&rdir=timereg&usemrn="&medarbejderid&"&showakt=1"
	                    
	                 
                    	'end if '*** Kopier job **'
            	
                        'end if ' <> 0
                        'next
                
                
                


	            end if '** Luk job **'

              
                Response.end

        case "FN_fjernjob"
                  
                 jobid = request("jobid")   
                 medid = request("usemrn")

                 '** Nulstillerforvalgt job for denne medarbejder ****'
                 strSQLUpdFvOff = "UPDATE timereg_usejob SET forvalgt = 0 WHERE medarb = "& medid &" AND jobid = "& jobid 
                 oConn.Execute(strSQLUpdFvOff)
                    '******************************************************'

        case "FM_sortOrder"
               Call AjaxUpdateTreg("timereg_usejob","forvalgt_sortorder","")
        
        
        case "FN_updatejobkom"
        
                 jobid = request("jobid")   
                 jobkom = request("jobkom") 
                 jobkom = replace(jobkom, "''", "") 
                 jobkom = replace(jobkom, "'", "")
                 jobkom = replace(jobkom, "<", "")
                 jobkom = replace(jobkom, ">", "")

                 oldKom = ""

                 strSQLUpdj = "SELECT kommentar FROM job WHERE id = "& jobid
                 oRec3.open strSQLUpdj, oConn, 3
                 if not oRec3.EOF then 
                 oldKom = oRec3("kommentar")
                 end if
                 oRec3.close

                 session_user = request("session_user")
                 dt = request("dt")

                 jobkom = "<span style=""color:#999999;"">" & dt &", "& session_user &":</span> " & jobkom & "<br>"& oldKom
                 
                 
                 '** Job kommentar ****'
                 strSQLUpdj = "UPDATE job SET kommentar = '"& jobkom &"' WHERE id = "& jobid 
                 oConn.Execute(strSQLUpdj)
                 
                 
        case "FN_showakt"        
            
           
             '<script src="inc/timereg_2006_load_func.js"></script> 
           
          

            '*** Variable ****
          

            felt = ""
            
            level = session("rettigheder")

            jobid = request("jobid")
            'Response.Write jobid
            'Response.end
           
            intEasyreg = request("FM_easyreg")

            
            if instr(request("job_aktids"),",") <> 0 then
            aktidsSQL = " WHERE (a.id = "&  replace(request("job_aktids"),",", "  OR a.id = ")
            else
            aktidsSQL = " WHERE (a.id = 0)"
            end if
           
            

           

            job_iRowLoops = split(request("job_iRowLoops"),",")
            job_iRowLoops_cn = job_iRowLoops(1) 

            
          
            foundone = "n"
            usemrn = request("usemrn")
            
            'redim aktdata_jq(100, 30)

            stDato = request("stDato")
            slDato = request("slDato")

            'dim tjekDato
            redim tjekdag(7)
            tjekdag(1) = stDato
            
            for x = 2 to 7
            tjekdag(x) = dateAdd("d", x-1, stDato)
            next

            call ersmileyaktiv()
            
            '*** ÆØÅ **'

            'Response.Write "faser:" &  request("job_akt_faser")
            'Response.end

          

            'call jq_format(request("job_akt_faser"))
            'job_akt_faser = split(jq_formatTxt,"xx,xx ")

            'for r = 1 to UBOUND(job_akt_faser)
            'Response.Write job_akt_faser(r) & "<br>"
            'next

           
            acn = 0

            'call jq_format(request("job_akt_navne"))
            'job_akt_navne = split(jq_formatTxt,"#,#")


            'Response.Write request("job_akt_navne")
            'Response.end

            lastFakDato = request("lastFakDato")
            
            ignorertidslas = request("ingTlaas")
            
            visTimerElTid = request("FM_vistimereltid")

            ignJobogAktper = request("FM_ignJobogAktper")

            'Response.Write "ignJobogAktper" & ignJobogAktper
            'Response.end

            if ignJobogAktper = "1" then
	        useDateStSQL = year(stDato) &"/"& month(stDato) &"/"& day(stDato)
            useDateSlSQL = year(slDato) &"/"& month(slDato) &"/"& day(slDato)
	        strSQLAktDatoKri = " AND (a.aktstartdato <= '"& useDateSlSQL &"' AND a.aktslutdato >= '"& useDateStSQL &"')"
            else
            strSQLAktDatoKri = ""
	        end if

            '**** Table aktiviteter start ***
            'strAktiviteter = "<table><form action=""timereg_akt_2006.asp?func=db"" method=""post""><tr><td><input type=""text"" name=""FM_timer"" id=""FM_timer"" value=""2"">TImer<br><input type=submit></td></tr></form></table>"
            'Response.write strAktiviteter
            'Response.end


            strAktiviteter = "<table cellspacing='0' cellpadding='2' border='0' width=100% bgcolor='#8caae6' id='tablejid_"& jobid &"'>"
            strAktiviteter = strAktiviteter &"<tr><td colspan=9 style=""background-color:#D6dff5; background-image:url('../ill/stripe_10.png'); height:2px;"">"
            strAktiviteter = strAktiviteter &"<img src='../ill/blank.gif' width='1' height='1' /></td></tr>"
            'strAktiviteter = strAktiviteter & ""

             strAktiviteter = strAktiviteter & "<input type='hidden' name='FM_jobid_timerOn' id='FM_jobid_timerOn_"& jobid &"' value='0'>"
             
                                         
           
           '** Dato ooverskrifter ***'
           call dageDatoer(1)
           strAktiviteter = strAktiviteter & dageDatoerTxt

           fo = 0
           '**************************************************'
           '**** Aktiviteter loop ****************************'
           'for r = 1 to UBOUND(job_aktids)
           '***************************************************

           'if cint(intEasyreg) <> 1 then
           'thisFase = job_akt_faser(r)
           'else
           'thisFase = ""
           'end if


              
            

          
             

          
         
               

                
             
                                    '*** aktiviteter **'
                                    aktnavn = ".."
                                    job_fase = ""

                                    job_tidslass_1 = "0"
                                    job_tidslass_2 = "0"
                                    job_tidslass_3 = "0"
                                    job_tidslass_4 = "0"
                                    job_tidslass_5 = "0"
                                    job_tidslass_6 = "0"
                                    job_tidslass_7 = "0"
            
                                    job_tidslass = 0
                                    job_tidslassSt = "0"
                                    job_tidslassSl = "0"

                                    job_kundenavn = ""
                                    job_kundenr = "0"

                                    job_jobnavn = ""
                                    job_jobnr = "0"

                                    aktstartdato = "01-01-2011"
                                    aktslutdato = "01-01-2011"

                                    job_incidentid = 0
                                    job_fakturerbar = 0
                                    
                                    job_jobans1 = 0
                                    job_jobans2 = 0

                                    job_bgr = 0
                                    job_antalstk = 0
                                    job_aktbudgettimer = 0

                                    job_abesk = ""

                                    job_internt = 0

                                    strSQLakt = "SELECT a.id AS aid, j.id AS jid, navn AS aktnavn, a.fase, tidslaas, tidslaas_st, tidslaas_sl, tidslaas_man, tidslaas_tir, "_
                                    &"tidslaas_ons, tidslaas_tor, tidslaas_fre, tidslaas_lor, tidslaas_son, k.kkundenavn, k.kkundenr, k.kid, j.jobnavn, j.jobnr, j.risiko, "_
                                    &" a.aktstartdato, a.aktslutdato, a.incidentid, a.fakturerbar, a.bgr, j.jobans1, j.jobans2, a.budgettimer AS aktbudgettimer, a.antalstk, a.beskrivelse AS abesk "_
                                    &" FROM aktiviteter AS a "_
                                    &" LEFT JOIN job AS j ON (j.id = a.job) "_
                                    &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "
                                    
                                    
                                    'if intEasyreg = 1 then
                                    strSQLakt = strSQLakt & aktidsSQL &")" & strSQLAktDatoKri & " GROUP BY a.id ORDER BY a.fase, a.sortorder, a.navn"
                                    'else
                                    'strSQLakt = strSQLakt & " WHERE a.id = "& job_aid & "" & strSQLAktDatoKri
                                    'end if

                                    'Response.Write intEasyreg &" sql: "& strSQLakt
                                    'Response.end
                                   
                                    
                                            oRec6.open strSQLakt, oConn, 3
                                            while not oRec6.EOF
                                            fo = 1
                                            felt = ""
                                            
                                            aktnavn = oRec6("aktnavn")
                                            call jq_format(aktnavn)
                                            aktnavn = jq_formatTxt

                                            if isNull(oRec6("fase")) <> true then
	                                        call jq_format(oRec6("fase"))
                                            job_fase = jq_formatTxt
                                            else
	                                        job_fase = ""
                                            end if
                                            
                                            job_tidslass_1 = oRec6("tidslaas_man")
                                            job_tidslass_2 = oRec6("tidslaas_tir")
                                            job_tidslass_3 = oRec6("tidslaas_ons")
                                            job_tidslass_4 = oRec6("tidslaas_tor")
                                            job_tidslass_5 = oRec6("tidslaas_fre")
                                            job_tidslass_6 = oRec6("tidslaas_lor")
                                            job_tidslass_7 = oRec6("tidslaas_son")
            
                                            job_tidslass = oRec6("tidslaas")
                                            job_tidslassSt = oRec6("tidslaas_st")
                                            job_tidslassSl = oRec6("tidslaas_sl")
            
                                          
                                            call jq_format(oRec6("kkundenavn"))
                                            job_kundenavn = jq_formatTxt

                                            job_kundenr = oRec6("kkundenr")
                                            job_kid = oRec6("kid")

                                            job_jobnavn = oRec6("jobnavn")
                                            job_jobnr = oRec6("jobnr")

                                            aktstartdato = oRec6("aktstartdato")
                                            aktslutdato = oRec6("aktslutdato")

                                            job_incidentid = oRec6("incidentid")
                                            job_fakturerbar = oRec6("fakturerbar")
                                            tfaktimThis = oRec6("fakturerbar") 'oRec6("tfaktim")

                                            job_jobans1 = oRec6("jobans1")
                                            job_jobans2 = oRec6("jobans2")

                                            job_bgr = oRec6("bgr")

                                            job_aktbudgettimer = oRec6("aktbudgettimer")
	                                        job_antalstk = oRec6("antalstk")

                                            job_abesk = oRec6("abesk")
                                            call jq_format(job_abesk)
                                            job_abesk = jq_formatTxt

                                            job_jid = oRec6("jid")

                                            job_aid = oRec6("aid")

                                            job_internt = oRec6("risiko")

                                              '*** aktiviteter ****'
                                        'strAktiviteter = strAktiviteter & "<tr><td bgcolor='#FFFFFF' colspan='2' valign='top' style='padding:5px 5px 4px 2px; border:1px #cccccc solid; border-bottom:0px; border-left:0px;'>"& aktnavn &"</td>"
                                        'strAktiviteter = strAktiviteter & "</tr></table>"
                                          
                                          
                                           '**** Faser ***'
                                           '*** Tilføj VZB iflg Cookie ***'
                                           if cint(intEasyreg) <> 1 then

                                                   if lastFase <> lcase(job_fase) AND len(trim(job_fase)) <> 0 AND job_fase <> "nonexx99w" then 
           
				                                        strAktiviteter = strAktiviteter & "<tr><td bgcolor='#FFFFFF' colspan='9' valign='top' style='padding:0px 0px 0px 0px; height:10px; border-bottom:0px #cccccc solid;'>"_
                                                        &"<img src='../ill/blank.gif' width='1' height='1' border='0' alt='' /><br /></td></tr>"
                
            
			                                            strAktiviteter = strAktiviteter & "<tr><td bgcolor='#Eff3ff' colspan='2' valign='top' style='padding:5px 5px 4px 2px; border:1px #cccccc solid; border-bottom:0px; border-left:0px;'>"
			                                            strAktiviteter = strAktiviteter & "<a id='"&jobid&"_"&lcase(job_fase)&"' name='"&lcase(job_fase)&"' href='#"&lcase(job_fase)&"' class='afase'> + "& tsa_txt_329 &": "& replace(job_fase, "_", " ")&"</a></td>"
                

                                                        strAktiviteter = strAktiviteter & "<td bgcolor='#FFFFFF' colspan='7' valign='top' style='padding:5px 5px 4px 2px; border-bottom:1px #cccccc solid;'>&nbsp;</td></tr>"
			  
			                                            if request.Cookies("tsa_fase")(""&jobid&"_"&lcase(job_fase) &"") = "visible" then
                                                        tr_vzb = "visible"
			                                            tr_dsp = ""
                                                        else

                                                            'select case lto
                                                            'case "dencker", "intranet - local"
                                                            tr_vzb = "visible"
			                                                tr_dsp = ""
                                                            'case else
                                                            'tr_vzb = "hidden"
		                                                    'tr_dsp = "none"
                                                            'end select
                                                        
                                                        end if
                

                                                         '*** Faser ***'
			                                            strAktiviteter = strAktiviteter & "<input name=""faseshow"" id='faseshow_"&jobid&"_"&lcase(job_fase)&"' type=""hidden"" value='"& tr_vzb &"' />"
			                                            strAktiviteter = strAktiviteter & "<input name=""faseshowid"" type=""hidden"" value='"&jobid&"_"&lcase(job_fase) &"' />"
                                                        '** end faser **'
			
			
			                                        else 
			                                               if len(trim(tr_vzb)) = 0 then
			                                                tr_vzb = "visible"
			                                                tr_dsp = ""
			                                                else
			                                                tr_vzb = tr_vzb
			                                                tr_dsp = tr_dsp
			                                                end if
			
			                                        end if 
            
                                            else

                                            tr_vzb = "visible"
			                                tr_dsp = ""

                                            end if

                                            '**faser end
                                          
                                          
                                          
                                          
                                          
                                               
                                       

                                          strAktiviteter = strAktiviteter & "<tr class='td_"&jobid&"_"&lcase(job_fase)&"' id='"& job_aid &"' style='visibility:"&tr_vzb&"; display:"&tr_dsp&";'>"
                                          'strAktiviteter = strAktiviteter & "<tr>"
				                          strAktiviteter = strAktiviteter & "<td bgcolor='#ffffff' valign='top' height='30' style='width:250px; padding-right:3px; border-bottom:1px #cccccc solid;'>"
				                                        


                                        
                                          '**job og akt. arrays **'
                                          strAktiviteter = strAktiviteter & "<input type='hidden' name='FM_jobid' id='FM_jobid_"& job_iRowLoops_cn &"' value='"& job_jid &"'>"
		                                  strAktiviteter = strAktiviteter & "<input type='hidden' name='FM_aktivitetid' id='FM_aktivitetid' value='"& job_aid &"'>" 
                                         
                                             

                                             
                                                  

                                                            '** Kunde **' 
                                                            if cint(intEasyreg) = 1 then
		                                                    
                                                            strAktiviteter = strAktiviteter & "<span style='font-size:9px; color:#5582d2;'>"& left(job_kundenavn, 10) &"</span>"
		                                                   
		        
		                                                       '*** jobnavn ****'
                                                            
		                                                            strAktiviteter = strAktiviteter & " <span style='font-size:9px;'><b>"& left(job_jobnavn, 30) & "</b> ("& job_jobnr &")</span><br />"


                                                                      '** Aktnavn **'
				                                                      strAktiviteter = strAktiviteter & "<b>"& left(aktnavn, 25) &"</b>"

                                                                    
                                                                    '*** Fase ***'
                                                                    if len(trim(job_fase)) <> 0 AND trim(job_fase) <> "nonexx99w" then
                                                                    strAktiviteter = strAktiviteter & "&nbsp;<span style='font-size:9px; color:yellowgreen;'>"& left(job_fase, 20) &"</span>"
                                                                    end if 
                                                               
                                                                
                                                                    strAktiviteter = strAktiviteter & "&nbsp;<span href=""#"" class=""fjeasyregakt"" id='"& job_aid &"' style=""color:red; font-size:8px; cursor:hand;"">[X]</span>"

                                                               
                                                             



                                                               else

                                                               '** Aktnavn **'
				                                                strAktiviteter = strAktiviteter & "<b>"& replace(aktnavn, "(ej fakturerbar)", "") &"</b><br />"

                                                             '** Datoer **'
                                                             strAktiviteter = strAktiviteter & "<span style='font-size:8px; line-height:10px; color:#999999;'>"
                                                            strAktiviteter = strAktiviteter & datepart("d", aktstartdato, 2,2) &". "& left(monthname(datepart("m", aktstartdato, 2,2)), 3) &". "& datepart("yyyy", aktstartdato, 2,2) & " - " 
				                                            strAktiviteter = strAktiviteter & datepart("d", aktslutdato, 2,2) &". "& left(monthname(datepart("m", aktslutdato, 2,2)), 3) &". "& datepart("yyyy", aktslutdato, 2,2)
				                                            

                                                            
                                                            '** Incident ID **'
                                                            if job_incidentid <> 0 then
				                                            strAktiviteter = strAktiviteter & "&nbsp;<i>"& tsa_txt_036 &": "& job_incidentid &"</i>" 
				                                            end if
                                                            
                                                            strAktiviteter = strAktiviteter & "</span><br /><span style='font-family:Arial; line-height:10px; font-size:8px; color:#5582d2;'>"

				                                            '**** Typer **
				                                            call akttyper(job_fakturerbar, 1)
                                                            strAktiviteter = strAktiviteter & left(akttypenavn, 12) 
		     

		                                                      end if


                                                           
				
                

                                                           
				
                                                           

                                                           if cint(intEasyreg) <> 1 then
                
                                                                    '** Timepris **'
                                                                    'level = 100
				                                                    'if level <= 2 OR level = 6 then
				                                                     '** Admin eller jobansvarlig **'
                                                                    if level = 1 OR session("mid") = job_jobans1 OR session("mid") = job_jobans2 then
                                                                    'if level = 1000 then
				                                                            call akttyper2009Prop(job_fakturerbar)
				                                                            if cint(aty_fakbar) = 1 OR job_fakturerbar = 5 then '** Viser tpriser på fakturerbare eller kørsel
				
				                                                                'if cint(j_intjobtype) = 1 then 'fastpris (altid angivet i DKK (Indtild videre!))
				                                                                    ''if cint(j_usejoborakt_tp) = 0 then 'job ell. akt. grundlag
				        
				                                                                        ''if j_timerforkalk <> 0 then
				                                                                        ''j_timerforkalk = j_timerforkalk
				                                                                        ''else
				                                                                        ''j_timerforkalk = 1
				                                                                        ''end if
				        
				                                                                    'aktTpThis = formatnumber(j_budget/j_timerforkalk, 2) '&" DKK"
				                                                                    ''else
				                                                                    ''aktTpThis = formatnumber(aktdata(iRowLoop, 30), 2) '&" DKK"
				                                                                    ''end if
				    
				                                                                'else 'medarbtimepriser
				                                                                        '*** tjekker for alternativ timepris på aktivitet
				                                                                        call alttimepris(job_aid, jobid, usemrn, 0)  
				      
				                                                                        '** Er der alternativ timepris på jobbet
                                                                                        if foundone = "n" then
                                                                                        call alttimepris(0, jobid, usemrn, 0) 
                                                                                        end if
				       
				                                                                        aktTpThis = formatnumber(intTimepris, 2) '& " " & valutaISO
				                                                                        'aktTpThis = "..."

				                                                                'end if  
				
				                                                            strAktiviteter = strAktiviteter & " // " & aktTpThis & " DKK "  
				
				                                                            end if
				
				
				
				
				                                                   end if

                                                            select case job_bgr 
                                                            case 1 
                                                            job_bgr_Txt = "Fork.: "& formatnumber(job_aktbudgettimer, 2)  & " t."
                                                            case 2
                                                            job_bgr_Txt = "Ant.: "& formatnumber(job_antalstk, 2)  & " stk."
                                                            case else
                                                            job_bgr_Txt = ".."
                                                            end select
                                                                
                                                            if len(trim(job_bgr_Txt)) <> 0 then
                                                            strAktiviteter = strAktiviteter & "// " & job_bgr_Txt & " "
                                                            end if

                                                            strAktiviteter = strAktiviteter & "</span>"

                                                            '** Timer real + kost (lsået fra) **'
                                                            akt_timertot = 0
                                                            strSQLttot = "SELECT SUM(timer) AS timertot FROM timer WHERE taktivitetid = "& job_aid & " GROUP BY taktivitetid"
                                                            oRec2.open strSQLttot, oConn, 3
                                                            if not oRec2.EOF then  
                                                            akt_timertot = oRec2("timertot")
        
                                                                
                                                            end if
                                                            oRec2.close

                                                            '** Timer real kun valgte medarb **'
                                                            'aktdata(iRowLoop, 38) = 0
                                                            akt_timerTot_medarb = 0
                                                            strSQLtmnr = "SELECT SUM(timer) AS timermnr FROM timer WHERE taktivitetid = "& job_aid & " AND tmnr = "& usemrn &" GROUP BY taktivitetid"
                                                            oRec2.open strSQLtmnr, oConn, 3
                                                            if not oRec2.EOF then  
                                                            akt_timerTot_medarb = oRec2("timermnr")
                                                            end if
                                                            oRec2.close



                                                            if job_bgr = 1 then 'bg = timer
                                                                if akt_timertot > job_aktbudgettimer then
                                                                fcolreal = "red"
                                                                else
                                                                fcolreal = "#2c962d"
                                                                end if
                                                            else
                                                            fcolreal = "#999999"
                                                            end if
                                                                
                                                             
                                                             '***  Real  ***'
                                                            strAktiviteter = strAktiviteter & "<span style='font-family:Arial; font-size:8px; color:"&fcolreal&";'>// "& formatnumber(akt_timertot, 2) &"</span>"
				                                            strAktiviteter = strAktiviteter & "<span style='font-family:Arial; font-size:8px; color:#5582d2;'> ("& formatnumber(akt_timerTot_medarb, 2) &")</span>"
				
				                                          
				
				
				
				
				                                                         if len(trim(job_abesk)) <> 0 then
		                                                              
		                                                                strAktiviteter = strAktiviteter & "<div id='div_aktbesk_"& job_aid &"' style='position:relative; visibility:visible; display:; left:0px; top:0px; width:200px; height:14px; overflow:hidden; z-index:200000;'>"
		                                                                strAktiviteter = strAktiviteter & "<table bgcolor=#ffffff border='0' cellspacing='0' cellpadding='0' width='100%' ><tr><td style='padding:0px;' valign=top class=lille>"
		                                                                
		                                                                
		                                                                strAktiviteter = strAktiviteter & "<a class='a_showhideaktbesk' id='aktbesk_"& job_aid &"' href='#'>"& job_abesk &"</a></td></tr></table></div>"
		                                                                                    
		                                                                                        
		                                                                end if
				                                            



                                                            end if 'easyreg



                                                            '** gl. stopurs link **'
				                                            strAktiviteter = strAktiviteter & "</td><td bgcolor=#ffffff valign=top style='padding:8px 3px 3px 3px; border-bottom:1px #cccccc solid; border-right:1px #cccccc solid;'>&nbsp;</td>"
				                                          '** akt navn kolonne end **'
           
                                                   
                                                    
                                                      

                                for x = 1 to 7
				
				                    '*** Standard ****
                                    maxl = 5
                                    fmbgcol = "#ffffff"
                                    fmborcl = "1px #999999 solid"
				                            
                                            
                                            
                                            '*** Henter timer ***
                                            timerThis = 0
                                            timerKomThis = ""
                                            tfaktimThis = tfaktimThis '0
                                            bopalThis = 0
                                            godkendtstatus = 0
                                            offentlig = 0
                                            timer_sttid = ""
                                            timer_sltid = ""
                                           

                                            SQldatoX = year(tjekdag(x)) &"/"& month(tjekdag(x)) &"/"& day(tjekdag(x))    
				                            strSQLtim = "SELECT timer, timerKom, tfaktim, bopal, godkendtstatus, offentlig, sttid, sltid, destination FROM timer WHERE taktivitetid = "& job_aid &" AND tdato = '"& SQldatoX &"' AND tmnr = "& usemrn

                                            'Response.Write strSQLtim & "<br><br>"
                                            'Response.end
                                            oRec2.open strSQLtim, oConn, 3
                                            while not oRec2.EOF

                                         

                                            'tfaktimThis = oRec6("tfaktim")
                                            timerThis = timerThis + oRec2("timer")
                                            timerKomThis = timerKomThis & oRec2("timerkom") 
                                            offentlig = oRec2("offentlig")
                                            bopalThis = oRec2("bopal")
                                            destinationThis = oRec2("destination")
                                            godkendtstatus = oRec2("godkendtstatus")
                                          
                                                    if len(trim(oRec2("sttid"))) <> 0 then
		                                                if left(formatdatetime(oRec2("sttid"), 3), 5) <> "00:00" then
		                                                timer_sttid = left(formatdatetime(oRec2("sttid"), 3), 5)
		                                                else
		                                                timer_sttid = ""
		                                                end if
	                                                else
	                                                timer_sttid = ""
	                                                end if
	
	                                                if len(trim(oRec2("sltid"))) <> 0 then
		                                                if left(formatdatetime(oRec2("sltid"), 3), 5) <> "00:00" then
		                                                timer_sltid = left(formatdatetime(oRec2("sltid"), 3), 5)
		                                                else
		                                                timer_sltid = ""
		                                                end if
	                                                else
	                                                timer_sltid = ""
	                                                end if


                                            oRec2.movenext
                                            wend
                                            oRec2.close





                                            '*** ÆØÅ **'
                                            call jq_format(timerKomThis)
                                            timerKomThis = jq_formatTxt
                                      
                                            
                                         

						                    'if tjekdag(x) = aktdata(job_iRowLoops, 0) then%>
						                    <%
                                            'if cdbl(timerThis) <> -10 then			
								
								                    
                                                    '*** Tjekker om det er en helligdag **'
                                                    call helligdage(tjekdag(x), 0)
	                                                'helligdageCol(x) = erHellig

								
								                    if len(timerKomThis) <> 0 then
								                    kommtrue = "+"
								                    else
								                    kommtrue = "<font color='#999999'>+</font>"
								                    end if 
						
								
								                    if x = 6 OR x = 7 OR erHellig = 1 then
								                    bgColor = "gainsboro"
								                    else
								                    bgColor = "#FFFFFF"
								                    end if
								                    


                                                    call akttyper2009Prop(tfaktimThis) 

                                                    if cdbl(Timerthis) <> 0 then

                                                                
                                                                erIndlast = 1

                                                    else
                                                        
                                                                if aty_pre <> 0 AND (x < 6) AND erHellig <> 1 then
								                                preVal = aty_pre
								                                else
								                                preVal = ""
								                                end if

                                                                TimerThis = preVal
                                                                erIndlast = 0


                                                    end if





								                    
								                    '** tjek erIndlast = 0
                                                   
                                                      
                                                    if cint(intEasyreg) <> 1 then
								                    call fakfarver(lastfakdato, tjekdag(1), tjekdag(x), erIndlast, usemrn, tfaktimThis, job_tidslass_1, job_tidslass_2, job_tidslass_3, job_tidslass_4, job_tidslass_5, job_tidslass_6, job_tidslass_7, job_tidslass, job_tidslassSt, job_tidslassSl, ignorertidslas, godkendtstatus)
								                    
                                                    'Response.Write "her  lastfakdato " & lastfakdato
                                                    'Response.end
                                                   
                                                  
                                                    

								                    felt = felt & "<td valign='top' bgcolor='"&bgColor&"' style='border-bottom:1px #cccccc solid; border-right:1px #cccccc solid; padding-top:6px;' id='td_"&job_iRowLoops_cn&""&x&"'>"
								                     
                                                   
								                    
								
								                    if visTimerElTid <> 1 OR cint(tfaktimThis) = 5 OR cint(tfaktimThis) = 61 OR aty_pre <> 0 OR cint(intEasyreg) = 1 OR job_internt = -1 then
								                    timtype = "text"
								                    tidtype = "hidden"
								                    br = ""
								                    else
								                    timtype = "hidden"
								                    tidtype = "text"
								                    br = "<br>"
								                    end if
								
								                    felt = felt &"<input class='dcls_"&x&"' type='"& timtype &"' name='FM_timer' id='FM_timer_"&job_iRowLoops_cn&"_"&x&"' value='"& timerThis &"' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; border:"&fmborcl&"; height:16px; width:45px; font-family:arial; font-size:10px;"" onkeyup=""doKeyDown()"" onfocus=""markerfelt('"&job_iRowLoops_cn&""&x&"','"&x&"'), showKMdailog('"&job_fakturerbar&"', '"&job_iRowLoops_cn&""&x&"', '"& bopalThis &"', '"& job_iRowLoops_cn &"', '"& job_kid &"')"";>"
								                    felt = felt &"<input type='hidden' name='FM_timer_opr' id='FM_timer_opr_"&job_iRowLoops_cn&"_"&x&"' value='"& timerThis &"'>"
								                    
                                                    
                                                    
                                                    felt = felt &"<input type='"& tidtype &"' name='FM_sttid' id='FM_sttid' value='"& timer_sttid &"' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; height:16px; width:35px; font-family:arial; font-size:10px; border:"&fmborcl&";"">"_
								                    &""& br &"<input type='"& tidtype &"' name='FM_sltid' id='FM_sltid' value='"& timer_sltid &"' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; height:16px; width:35px; font-family:arial; font-size:10px; border:"&fmborcl&";"">"
								
								                    if maxl <> 0 then
								                    felt = felt &"&nbsp;<a class=""a_showkom"" href=""#anc_s0"" name='anc_"&job_iRowLoops_cn&"_"&x&"' onClick=""expandkomm('"&job_iRowLoops_cn&"', '"&x&"')"">"& kommtrue &"</a>"
								                    end if
								
								                    felt = felt &"<input type='hidden' name='FM_timer' id='FM_timer' value='xx'><input type='hidden' name='FM_datoer' id='FM_datoer' value='"& tjekdag(x) &"'>"_
								                    &"<input class='d_kom_cls_"&x&"' type='hidden' name='FM_kom_"&job_iRowLoops_cn&""&x&"' id='FM_kom_"&job_iRowLoops_cn&""&x&"' value='"&timerKomThis&"'>"_
								                    &"<input type='hidden' name='FM_feltnr' id='FM_feltnr' value='"&job_iRowLoops_cn&""&x&"'>"
								

								                    if len(offentlig) <> 0 then
								                    kommOnOff = offentlig 
								                    else
								                    kommOnOff = 0
								                    end if
								
								
								                    felt = felt &"<input type='hidden' name='FM_off_"&job_iRowLoops_cn&""&x&"' id='FM_off_"&job_iRowLoops_cn&""&x&"' value='"&kommOnOff&"'>"
								                    felt = felt &"<input id='FM_bopal_"&job_iRowLoops_cn&""&x&"' name='FM_bopal_"&job_iRowLoops_cn&""&x&"' type=""hidden"" value="""& bopalThis &"""/>"
								                    felt = felt &"<input id='FM_destination_"&job_iRowLoops_cn&""&x&"' name='FM_destination_"&job_iRowLoops_cn&""&x&"' type=""hidden"" value="""& destinationThis &"""/>"
								                    
                                                    felt = felt &"</td>"

                                                    else '** Easyreg
								                    felt = felt &"<td bgcolor=#FFFFFF valign=top style='border-bottom:1px #cccccc solid; border-right:1px #cccccc solid; padding-top:6px;'>"
                                                    
                                                    felt = felt &"<input class='dcls_"&x&"' type='"& timtype &"' name='FM_timer' id='FM_timer_"&job_iRowLoops_cn&"_"&x&"' value='"& timerThis &"' maxlength='5' style=""border:1px #999999 solid; height:16px; width:45px; font-family:arial; font-size:10px;"";>"
								                    
                                                    felt = felt &"<input type='hidden' name='FM_sttid' id='FM_sttid' value='"& timer_sttid &"'><input type='hidden' name='FM_sltid' id='FM_sltid' value='"& timer_sltid &"'>"
								

                                                    felt = felt &"<input type='hidden' name='FM_timer' id='FM_timer' value='xx'><input type='hidden' name='FM_datoer' id='FM_datoer' value='"& tjekdag(x) &"'>"_
                                                    &"<input class='d_kom_cls_"&x&"' type='hidden' name='FM_kom_"&job_iRowLoops_cn&""&x&"' id='FM_kom_"&job_iRowLoops_cn&""&x&"' value='"&timerKomThis&"'>"_
                                                    &"<input type='hidden' name='FM_feltnr' id='FM_feltnr' value='"&job_iRowLoops_cn&""&x&"'>"

                                                    felt = felt &"</td>"

                                                    end if
								
								
							    
							                      
						       
						                    %>
						                    <%'else
								              '          if instr(ugedagUdskrevet, weekday(tjekdag(x), 2)) = 0 then 
								
								
								
								
								                '                if x = 6 OR x = 7 OR helligdageCol(x) = 1 then
								                '                bgColor = "gainsboro"
								                '                else
								                '                bgColor = "#FFFFFF"
								                '                end if
							'	
								                                'call akttyper2009Prop(aktdata(job_iRowLoops_cn, 11)) 
								
								                 '               if aty_pre <> 0 AND (x < 6) ANd helligdageCol(x) <> 1 then
								                 '               preVal = aty_pre
								                  '              else
								                  '              preVal = ""
								                  '              end if
								
								                   '             erIndlast = 0
								
								
                                                                ''call fakfarver(lastfakdato, tjekdag(1), tjekdag(x), erIndlast, usemrn, aktdata(job_iRowLoops_cn, 11))
								
								                  '              felt = "<td valign=top bgcolor='"&bgColor&"' style=""border-bottom:1px #cccccc solid; border-right:1px #cccccc solid; padding-top:6px;"" id='td_"&job_iRowLoops_cn&""&x&"'>her"
								
								
								
								                   '             if visTimerElTid <> 1 OR aktdata(job_iRowLoops_cn, 11) = 5 OR aktdata(job_iRowLoops_cn, 11) = 61 OR aty_pre <> 0 OR cint(intEasyreg) = 1 then
								                   '             timtype = "text"
								                   '             tidtype = "hidden"
								                   '             br = ""
								                   '             else
								                   '             timtype = "hidden"
								                   '             tidtype = "text"
								                   '             br = "<br>"
								                   '             end if
								
								
								                                ''if lto <> "execon" then
								                   '             felt = felt &"<input class='dcls_"&x&"' type='"& timtype &"' name='FM_timer' id='FM_timer_"&job_iRowLoops_cn&"_"&x&"' value='"&preVal&"' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; border:"&fmborcl&"; height:16px; width:45px; font-family:arial; font-size:10px; line-height:12px;"" onkeyup=""doKeyDown()""  onfocus=""markerfelt('"&job_iRowLoops_cn&""&x&"','"&x&"'), showKMdailog('"&aktdata(job_iRowLoops_cn, 11)&"', '"&job_iRowLoops_cn&""&x&"', '0', '"& job_iRowLoops_cn &"', '"& topAdd &"')"";>"
								                   '             felt = felt &"<input type='hidden' name='FM_timer_opr' id='FM_timer_opr_"&job_iRowLoops_cn&"_"&x&"' value='"& aktdata(job_iRowLoops_cn, 1) &"'>"
								
								                                ''else
								                                ''felt = felt &"<input type='"& timtype &"' name='FM_timer' id='FM_timer' value='"&preVal&"' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; border:"&fmborcl&"; height:16px; width:45px; font-family:arial; font-size:10px; line-height:12px;"" onfocus=""markerfelt('"&job_iRowLoops_cn&""&x&"','"&x&"')"";>"
							                                    ''end if
								
								                   '             felt = felt &"<input type='"& tidtype &"' name='FM_sttid' id='FM_sttid' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; height:16px; width:35px; font-family:arial; font-size:10px; line-height:12px; border:"&fmborcl&";"">"_
								                   '             &""& br &"<input type='"& tidtype &"' name='FM_sltid' id='FM_sltid' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; height:16px; width:35px; font-family:arial; font-size:10px; line-height:12px; border:"&fmborcl&";"">"
								
								                   '             if maxl <> 0 then
								                   '             felt = felt &"&nbsp;<a class=""a_showkom"" href=""#anc_s0"" name='anc_"&job_iRowLoops_cn&"_"&x&"' onClick=""expandkomm('"&job_iRowLoops_cn&"', '"&x&"')""><font color='#999999'>+</font></a>"
								                   '             end if
								
								                   '             felt = felt &"<input type='hidden' name='FM_timer' id='FM_timer' value='#'><input type='hidden' name='FM_datoer' id='FM_datoer' value='"& tjekdag(x) &"'>"_
								                   '             &"<input type='hidden' class='d_kom_cls_"&x&"' name='FM_kom_"&job_iRowLoops_cn&""&x&"' id='FM_kom_"&job_iRowLoops_cn&""&x&"' value=''>"_
								                   '             &"<input type='hidden' name='FM_feltnr' id='FM_feltnr' value='"&job_iRowLoops_cn&""&x&"'>"
								
								
								                   '             if len(aktdata(job_iRowLoops_cn, 7) ) <> 0 then
								                   '             kommOnOff = aktdata(job_iRowLoops_cn, 7) 
								                   '             else
								                   '             kommOnOff = 0
								                   '             end if
								
								
								
								                   '             felt = felt &"<input type='hidden' name='FM_off_"&job_iRowLoops_cn&""&x&"' id='FM_off_"&job_iRowLoops_cn&""&x&"' value='"&kommOnOff&"'>"
							                       '             felt = felt &"<input id='FM_bopal_"&job_iRowLoops_cn&""&x&"' name='FM_bopal_"&job_iRowLoops_cn&""&x&"' type=""hidden"" value=""0"" />"
								                   '             felt = felt &"<input id='FM_destination_"&job_iRowLoops_cn&""&x&"' name='FM_destination_"&job_iRowLoops_cn&""&x&"' type=""hidden"" value=""""/>"
								                   '             felt = felt &"</td>"
							
								
								
								
								                    '    end if%>
						
						                    <%'end if%>
				                    <%
				                    ugedagUdskrevet = ugedagUdskrevet & "," & weekday(tjekdag(x), 2) & "#"
				                    'Response.write ugedagUdskrevet
				                    %>
				
				                    <%'end if%>
			                    <%next 'x 1 - 7

                                
								                    

                strAktiviteter = strAktiviteter & felt & "</tr>"
                'response.flush

                'Response.write strAktiviteter
                'Response.end 


            '*** akt end
                

            lastFase = lcase(job_fase)
            job_iRowLoops_cn = job_iRowLoops_cn + 1
            acn = acn + 1
            oRec6.movenext
            wend
            oRec6.close
            'next

         
        

            if fo = 0 OR acn = 0 then
            strAktiviteter = strAktiviteter & "<tr><td bgcolor='#FFFFFF' colspan='9' valign='top' style='padding:20px 0px 20px 20px; height:10px; border-bottom:0px #cccccc solid;'>"_
            &"Der blev ikke fundet nogen aktiviteter p&aring dette job. Tjek at du har rettigheder til aktiviteterne, og at de ligger i det korrekte tidsinterval.</td></tr>"
            else

            '*** ÆØÅ **'
            call jq_format(tsa_txt_085)
            tsa_txt_085jq = jq_formatTxt

            strAktiviteter = strAktiviteter & "<tr><td bgcolor='#FFFFFF' colspan='9' align='right' valign='top' style='padding:20px 0px 20px 20px; height:10px; border-bottom:0px #cccccc solid;'><input type=""submit"" value='"& tsa_txt_085jq &" >>' >"_
            &"<br>&nbsp;</td></tr>"
            end if


           strAktiviteter = strAktiviteter & "</table>"
           
           Response.Write strAktiviteter


           'call akttyper2009(2)
            

           Response.end
           
           
           
           


        

                
        case "FN_jobidenneuge"
         

         stDato = request("stDato")
         slDato = request("slDato")

         'stDato = year(stDato) & "/"& month(stDato) & "/"& day(stDato)
         'slDato = year(slDato) & "/"& month(slDato) & "/"& day(slDato)

         jobans = request("jobans")
         'medid = request("medid")

         lto = request("lto")
         level = session("rettigheder")

         if jobans = "1" then
         jobansSQL = " AND (jobans1 = "& session("mid") &" OR jobans2 = "& session("mid") &")"
         else
         jobansSQL = ""  
         end if

         ignrper = request("ignrper")
         limit = request("limit")
         
         select case ignrper
         case "1" 
         dTop = 14
         DbeginC = 14
         case "2"
         dTop = 21
         DbeginC = 14
         case "3"
         dTop = 28
         DbeginC = 14
         case "4"
         dTop = 35
         DbeginC = 14
         case else
         dTop = 14
         DbeginC = 14
         end select


         alfa = request("alfa")
         if alfa = "1" then
         orderbySQL = " jobnavn"
         else
         orderbySQL = " jobslutdato DESC"
         end if

          call akttyper2009(2)

         strJoiDUge = "<table cellspacing=0 cellpadding=0 border=0 width='100%'><tr><td>"
         
         if level <= 2 then
         strJoiDUge = strJoiDUge & "<a href='webblik_joblisten.asp?nomenu=1&rdir=treg&FM_kunde=0&FM_medarb_jobans="&session("mid")&"&st_sl_dato=2' target=""_blank"" class=""rmenu"">+ Rediger <br>igangv.job</a>"
         else
         strJoiDUge = strJoiDUge &"&nbsp;"
         end if
         
         strJoiDUge = strJoiDUge &"</td><td class=lille align=right><b>Fork.</b></td><td class=lille align=right><b>Real.</b></td><td class=lille align=right><b>Fak.</b></td>"



         for d = 0 to dTop
         datoWrt = 0

         dcounter = DbeginC-d



         stDatoSQL = dateAdd("d", -dcounter, slDato)
       
        

         stDatoSQL = year(stDatoSQL) & "/" & month(stDatoSQL) & "/"& day(stDatoSQL)
         jobdtSQL = " AND jobslutdato = '" & stDatoSQL & "'"


         strSQLfv = "select j.id AS jid, j.jobnavn, j.jobnr, j.jobslutdato, jobans1, jobans2, (budgettimer+ikkebudgettimer) AS forkalk, jo_bruttooms AS bruttooms, "_
         &" k.kkundenavn, k.kkundenr, k.kid, SUM(t.timer) AS realtimer, SUM(t.timer*t.timepris*(t.kurs/100)) AS realoms "_
         &" FROM job AS j "_
         &" LEFT JOIN kunder AS k on (k.kid = j.jobknr) "_
         &" LEFT JOIN timer AS t ON (t.tjobnr = j.jobnr AND ("& aty_sql_realhours &")) "_
         &" WHERE jobstatus = 1 AND risiko <> - 1 "& jobdtSQL &" "& jobansSQL &" GROUP BY j.id ORDER BY " & orderbySQL & " LIMIT 0,"& limit
         
         'Response.write strSQLfv & "<br><br>"
         'Response.Flush

         lastSlutDato = 0
         oRec3.open strSQLfv, oConn ,3
         while not oRec3.EOF 
         
         if alfa <> "1" AND datoWrt = 0 then
         
         'if lastSlutDato <> oRec3("jobslutdato") OR lastSlutDato = 0 then
         strJoiDUge = strJoiDUge &"<tr bgcolor=""#FFFFFF""><td class=lille valign=bottom style=""height:30px;"" colspan=4><span style=""color:#000000; font-size:11px;""><b>"&weekdayname(weekday(stDatoSQL)) &" d. "& formatdatetime(stDatoSQL, 1)&"</b></span></td></tr>"
         'end if 
         datoWrt = 1
         end if
       

         'jobs.asp?menu=job&func=red&id="& oRec3("id") &"&int=1&rdir=treg

         if (lto = "epi" OR lto = "intranet - local") OR (level = 1 OR (oRec3("jobans1") = session("mid") OR oRec3("jobans2") = session("mid"))) then

         strJoiDUge = strJoiDUge &"<tr><td class=lille style=""border-bottom:1px #CCCCCC solid;""><a href='webblik_joblisten.asp?nomenu=1&rdir=treg&FM_kunde="&oRec3("kid")&"&jobnr_sog="&oRec3("jobnr")&"&FM_medarb_jobans="&session("mid")&"&st_sl_dato=2' class=rmenu target=""_blank"">"& left(oRec3("jobnavn"), 10) &" ("& oRec3("jobnr") &")</a>"_
         &"&nbsp;<a href=""jobs.asp?menu=job&func=red&id="& oRec3("jid") &"&int=1&rdir=treg"" style=""font-size:9px; color:#6CAE1C;"">"& left(tsa_txt_251, 3) &".</a><br>"_
         &"<span style=""font-size:8px; color:#999999;"">"& left(oRec3("kkundenavn"), 14) &"</span></td>"_
         &"<td align=right class=lille style=""border-bottom:1px #CCCCCC solid; white-space:nowrap;"">"& formatnumber(oRec3("forkalk"),0) &" t.<br>"& formatnumber(oRec3("bruttooms")/1000,0) &" k.</td>"_
         &"<td align=right class=lille style=""border-bottom:1px #CCCCCC solid; white-space:nowrap;"">"& formatnumber(oRec3("realtimer"),0) &" t.<br>"& formatnumber(oRec3("realoms")/1000,0) &" k.</td>"


         fakbelob = 0
         strSQLfak = "SELECT if (f.faktype <> 1, SUM((f.beloeb * f.kurs)/100), SUM((f.beloeb * -1 * f.kurs)/100)) AS fakbelob FROM "_
         & " fakturaer AS f WHERE (f.jobid = "& oRec3("jid") &" AND medregnikkeioms = 0 AND shadowcopy = 0) GROUP BY f.jobid"

         'Response.Write "<br>"& strSQLfak
         'Response.flush

         oRec4.open strSQLfak, oConn, 3

         if not oRec4.EOF then
         fakbelob = oRec4("fakbelob")
         end if
         oRec4.close

         strJoiDUge = strJoiDUge &"<td align=right class=lille style=""border-bottom:1px #CCCCCC solid; white-space:nowrap;""><br>"& formatnumber(fakbelob/1000,0) &" k.</td></tr>"
       
         '<td class=lille align=right>"& formatnumber(oRec3("bruttooms")) &" DKK</td>

         else

          strJoiDUge = strJoiDUge &"<tr><td class=lille style=""border-bottom:1px #CCCCCC solid;"">"& left(oRec3("jobnavn"), 10) &" ("& oRec3("jobnr") &")</a><br>"_
         &"<span style=""font-size:8px; color:#999999;"">"& left(oRec3("kkundenavn"), 14) &"</span></td>"_
         &"<td align=right class=lille style=""border-bottom:1px #CCCCCC solid; white-space:nowrap;"">"& formatnumber(oRec3("forkalk"),0) &" t.</td>"_
         &"<td align=right class=lille style=""border-bottom:1px #CCCCCC solid; white-space:nowrap; "">"& formatnumber(oRec3("realtimer"),0) &" t.</td><td>&nbsp;</td></tr>"
      


         end if
        
         

         lastSlutDato = oRec3("jobslutdato")
         oRec3.movenext
         wend
         oRec3.close  




          '** Henter aktiviteter, der har slutdato der afviger fra job ***'
          
           if alfa <> "1" AND ignrper <> "1" AND lto = "dencker" then
             

                           'lastSlutDatoSQL = year(lastSlutDato) & "/"& month(lastSlutDato) & "/"& day(lastSlutDato) 
                        

                          strSQLfa = "SELECT a.navn, aktslutdato, j.jobnavn, j.jobnr FROM aktiviteter AS a LEFT JOIN job AS j ON (j.id = a.job) "_
                          &" WHERE aktstatus = 1 AND aktslutdato <> jobslutdato AND aktslutdato = '"& stDatoSQL &"' "& jobansSQL &" ORDER BY aktslutdato DESC"
         
                         'Response.write strSQLfv & "<br><br>"
                         'Response.Flush

                         'lastSlutDato = 0
                         aktfirst = 0
                         oRec4.open strSQLfa, oConn, 3
                         while not oRec4.EOF 
         
                              'if aktfirst = 0 then
                               'strJoiDUge = strJoiDUge &"<tr bgcolor=""#FFFFFF""><td colspan=3 valign=bottom style=""height:15px;"" class=lille><b>Aktiviteter:</b><br /><span style=""font-size:9px;"">Hvor slutdato afviger fra job</span></td></tr>"
                               'aktfirst = 1
                               'end if



                             if datoWrt = 0 then
         
                             'if lastSlutDato <> oRec3("jobslutdato") OR lastSlutDato = 0 then
                             strJoiDUge = strJoiDUge &"<tr bgcolor=""#FFFFFF""><td class=lille valign=bottom style=""height:30px;"" colspan=3><span style=""background-color:#ffC0CB;""><b>"&weekdayname(weekday(stDatoSQL)) &" d. "& formatdatetime(stDatoSQL, 1)&"</b></span></td></tr>"
                             'end if 
                             datoWrt = 1
                             end if

                         'if lastSlutDato <> oRec4("aktslutdato") OR lastSlutDato = 0 then
                         'strJoiDUge = strJoiDUge &"<tr bgcolor=""#DCF5BD""><td colspan=3 class=lille><b>"& oRec3("aktslutdato") &"</b></td></tr>"
                         'end if 

                         strJoiDUge = strJoiDUge & "<tr bgcolor=""#EFF3FF""><td class=lille colspan=3 style=""border-bottom:1px #CCCCCC solid;"">Akt. "& left(oRec4("navn"), 10) &"<br />"_
                         &"<span style=""font-size:9px; color:#999999;"">"& left(oRec4("jobnavn"), 10) &" ("& oRec4("jobnr") &")</span></td></tr>"


                         'lastSlutDato = oRec4("aktslutdato")
                         oRec4.movenext
                         wend
                         oRec4.close  

         end if


         next
       

         strJoiDUge = strJoiDUge &"</table>"       


          '*** ÆØÅ **'
          call jq_format(strJoiDUge)
          strJoiDUge = strJoiDUge
                

         Response.write strJoiDUge
         Response.end


        
        case "FN_getTildelJob_step_1"

        
              '<script src="inc/timereg_2006_load_func.js"></script> 
              

                if len(trim(request.form("prgrp"))) <> 0 then
                selPrgrp = request.form("prgrp")
                else
                selPrgrp = 0
                end if


              selmedarbid = request("vlgtmedarb")

              if len(trim(request("medCookie"))) <> 0 AND isNull(request("medCookie")) <> true AND request("medCookie") <> "null" then
              medarbCookie = request("medCookie") 
              else
              medarbCookie = 0
              end if 
               

               'Response.Write medarbCookie &": "& request("medCookie")
               'Response.end
                  
         '** Medarbejder i prgtrp ****'
         call  medarbiprojgrp(selPrgrp, selmedarbid)
            
           
           strTildelJob = "<select id=""tildel_sel_medarb"" style=""width:160px; font-size:9px;"">" 
            
         strSQLmigrp = "SELECT mid, mnavn, mansat FROM medarbejdere WHERE mid = 0 "& medarbgrpIdSQLkri &" ORDER BY mnavn " 
         'Response.Write "<option>"& strSQLmigrp & " - " & selmedarbid & "</option></select>"
         'Response.end
         
         oRec3.open strSQLmigrp, oConn, 3
         
         'strTildelJob = strTildelJob &"<option value='0'>Alle</option>"
         
         while not oRec3.EOF 
            
           if cint(medarbCookie) = cint(oRec3("mid")) then
           thisMSel = "SELECTED"
           else
           thisMSel = ""
           end if 

            if oRec3("mansat") = "3" then
            mansatTxt = " - Passiv"
            else
            mansatTxt = ""
            end if
          
         strTildelJob = strTildelJob &"<option value='"& oRec3("mid") &"' "& thisMSel &">"& oRec3("mnavn") &" "& mansatTxt &"</option>"
        
         oRec3.movenext
         wend
         oRec3.close
       
       
         Response.Write strTildelJob & "</select><input id=""tildel_sel_medarb_bt"" type=""button"" value="">>"" style=""font-size:9px;""><br><br>"
         Response.end


        case "FN_getTildelJob"

                if len(trim(request.form("prgrp"))) <> 0 then
                selPrgrp = request.form("prgrp")
                else
                selPrgrp = 0
                end if


              selmedarbid = request("vlgtmedarb")
             
                

         
         '<script src="inc/timereg_2006_load_func.js"></script> 
             

        strTildelJob = "<table width=""100%"" cellpadding=0 cellspacing=0 border=0>"
        strTildelJob = strTildelJob & "<input type=""hidden"" name=""tuid_mid"" value='"& selmedarbid &"'>"
        strTildelJob = strTildelJob & "<input type=""hidden"" name=""tuid_progrp"" value='"& selPrgrp &"'>"

        if selmedarbid <> 0 then
        selMedarbSQLkri = "medarb = "& selmedarbid
        else
        selMedarbSQLkri = "medarb <> 0"
        end if

         strSQLfv = "select jobid, medarb, j.jobnavn, j.jobnr, mnavn, mnr, tu.id AS tuid, forvalgt, kkundenavn, kkundenr FROM timereg_usejob AS tu"_
         &" LEFT JOIN job AS j on (j.id = tu.jobid) "_
         &" LEFT JOIN kunder AS k on (k.kid = j.jobknr) "_
         &" LEFT JOIN medarbejdere AS m ON (m.mid = tu.medarb) WHERE "& selMedarbSQLkri &" AND (jobstatus = 1 OR jobstatus = 3)"_
         &" GROUP BY j.id ORDER BY medarb, forvalgt DESC, forvalgt_sortorder, j.jobnavn"
         
        'Response.write strSQLfv
        'Response.end

         lastMid = 0
         oRec3.open strSQLfv, oConn ,3
         while not oRec3.EOF 
         

         call jq_format(oRec3("mnavn"))
         mNavn = jq_formatTxt
         
         if lastMid <> oRec3("medarb") OR lastMid = 0 then
         strTildelJob = strTildelJob & "<tr><td colspan=2 class=lille>Aktiv jobliste for:<br><b>"& left(mNavn, 30) &"</b></td></tr>"
         end if 

         call jq_format(oRec3("jobnavn"))
         jNavn = jq_formatTxt
        
         strTildelJob = strTildelJob & "<tr><td class=lille style=""border-bottom:1px #CCCCCC solid;"">"& left(jNavn, 20) &" ("& oRec3("jobnr") &")<br><span style=""font-size:8px; color:#999999;"">"& left(oRec3("kkundenavn"), 20) &" ("& oRec3("kkundenr") &")</span></td>"

         if cint(oRec3("forvalgt")) = 1 then
         fvlgtCHK = "CHECKED"
         else
         fvlgtCHK = ""
         end if

         strTildelJob = strTildelJob & "<td style=""border-bottom:1px #CCCCCC solid;""><input id='tuid_"& oRec3("tuid") &"' name=""tuid"" class=""aktivejobid"" type=""checkbox"" value='"& oRec3("tuid") &"' "& fvlgtCHK &" /></td></tr>"
         strTildelJob = strTildelJob & "<input type=""hidden"" name=""tuid"" value='XX'>"
         

        

         lastMid = oRec3("medarb")
         oRec3.movenext
         wend
         oRec3.close  

         strTildelJob = strTildelJob & "<tr><td class=lille><br><input type=""checkbox"" name=""tuid_progrp_alle"" id=""tuid_progrp_alle"" value=""1"" /> Opdater alle medarb. i valgte grp.</td></tr>"
         strTildelJob = strTildelJob & "<tr><td align=right><input type=""submit"" value="" Opdater >>"" style=""font-size:9px;""></td></tr></table>"
         

         call jq_format(strTildelJob)
         strTildelJob = jq_formatTxt

         Response.Write strTildelJob
         Response.end


        case "FN_opdtodo"
                
                if len(trim(request.form("cust"))) <> 0 then
                todoID = request.form("cust")
                else
                todoId = 0
                end if
                
                strSQL = "UPDATE todo_new SET afsluttet = 1 WHERE id = "& todoId & " OR parent = "& todoId
                oConn.execute(strSQL)
                
                'Response.redirect "timereg_akt_2006.asp" 
                
        
        
        
        
        
        
        
        case "FM_get_destinations"
        
            	if len(trim(request("kid"))) <> 0 then
				kid = request("kid") 
				else
				kid = 0
				end if
				
				if len(trim(request("xvalbegin"))) <> 0 then
				xvalbegin = (request("xvalbegin") + 1)/1
				else
				xvalbegin = 2
				end if
				
				    'visalle = request("visalle")
				    if len(trim(request("visalle"))) <> 0 then
				    visalle = request("visalle")
				    else
				    visalle = 0
				    end if
				    
				    if visalle <> 0 then
				        strSQLKundKri = " AND kid <> 0"
				    else
				        strSQLKundKri = " AND kid = " & kid
				    end if
				
				 more = request("more")
				
				'Response.Write "visalle: "& request("visalle")  & "<br>"
				
				
				if len(trim(request.form("cust"))) <> 0 then
				sogeKri = " (kp.navn LIKE '"& request.form("cust") &"%' OR k.kkundenavn LIKE '"& request.form("cust") &"%')"
				else
				sogeKri = " kkundenavn <> 'sdfsdfd556irp'"
				end if
				
				
				strSQL2 = "SELECT kp.id, kp.navn, kp.email, kp.dirtlf, kp.mobiltlf, "_
				&" kp.adresse As kpadr, kp.postnr As kpzip, kp.town As kptown, kp.land As kpland, kp.afdeling AS kpafd, "_
				&" k.kid, k.kkundenavn, k.kkundenr, k.adresse, k.postnr, k.city, k.telefon, k.land FROM kunder k "_ 
				&" LEFT JOIN kontaktpers kp ON (kp.kundeid = k.kid)"_
				&" WHERE "& sogeKri & strSQLKundKri & " ORDER BY k.kkundenavn, kp.navn LIMIT 10"
			   
			    'Response.Write strSQL2 & "<br><br>"
			    'Response.flush
			   
			    strFil_Kon_Txt = ""
			   
			    oRec2.open strSQL2, oConn, 3
                x = xvalbegin
                k = 0
                while not oRec2.EOF
                
                
                        if ((lastknavn <> oRec2("kkundenavn") AND instr(lcase(oRec2("kkundenavn")), lcase(request.form("cust"))) <> 0)) OR (cint(more) = 0 AND cint(k) = 0) then 'lastknavn <> oRec2("kkundenavn")) then
                            strFil_Kon_Txt = strFil_Kon_Txt &  "<br /><input id='ko"&x&"chk' type=checkbox />"_ 
                            &"<select id='ko"&x&"sel' style='font-size:9px;'>"_
                            &"<option value=0>-</option>"_
                            &"<option value=2>"&tsa_txt_363&"</option>"_
                            &"<option value=3>"&tsa_txt_364&"</option>"_
                            &"<option value=1>"&tsa_txt_362&"</option>"_
                            &"</select><input id='ko"&x&"kid' type=hidden value='k_"&oRec2("kid")&"' />" 
                            
                            '&"<b>&nbsp;"& left(oRec2("kkundenavn"), 30) &"</b> ("&oRec2("kkundenr")&") " & tsa_txt_370
                         
                           strFil_Kon_Txt = strFil_Kon_Txt &  "<br /><img src='../ill/blank.gif' width=1 height=5 /><br>"_
                           &"<div id='ko"&x&"' style=""padding:5px; border:1px #cccccc solid; top:5px; background-color:#DCF5BD;"">"
                        
                                if len(trim(oRec2("kkundenavn"))) <> 0 then
                                strFil_Kon_Txt = strFil_Kon_Txt &  oRec2("kkundenavn") &"<br />"_
                                & oRec2("adresse") &"<br />"_
                                & oRec2("postnr") &" "& oRec2("city")&"<br />"_
                                & oRec2("land") &"<br />"
                                end if
                        
                        strFil_Kon_Txt = strFil_Kon_Txt &  "</div>"
                        strFil_Kon_Txt = strFil_Kon_Txt &  "<br />"
                            
                        x = x + 1   
                        k = 1 
                        end if
                        
                
                                if len(trim(oRec2("id"))) <> 0 AND ((cint(more) = 0) OR (instr(lcase(oRec2("navn")), lcase(request.form("cust"))) <> 0) AND cint(more) = 1 ) then
                                
                                strFil_Kon_Txt = strFil_Kon_Txt &  "<br /><input id='ko"&x&"chk' type=checkbox />"_ 
                                &"<select id='ko"&x&"sel' style='font-size:9px;'>"_
                                &"<option value=0>-</option>"_
                                &"<option value=2>"&tsa_txt_363&"</option>"_
                                &"<option value=3>"&tsa_txt_364&"</option>"_
                                &"<option value=1>"&tsa_txt_362&"</option>"_
                                &"</select><input id='ko"&x&"kid' type=hidden value='kp_"&oRec2("id")&"' />"
                                
                                '&"<b>&nbsp;"& oRec2("navn")&"</b> " & tsa_txt_369 & ""_
                                '&"<br><span style=""color:#999999; size:9px;""><b>"& left(oRec2("kkundenavn"), 30)&"</b> ("&oRec2("kkundenr")&")</span>"
                                    
                                           
                                           
                                        strFil_Kon_Txt = strFil_Kon_Txt &  "<br /><img src=""../ill/blank.gif"" width=1 height=5 /><br>"_
                                        &"<div id='ko"&x&"' style=""padding:5px; border:1px #cccccc solid; top:5px; background-color:#EFF3FF;"">"
                                        if len(trim(oRec2("navn"))) <> 0 then
                                        strFil_Kon_Txt = strFil_Kon_Txt & oRec2("navn") &"<br />"
                                        strFil_Kon_Txt = strFil_Kon_Txt & "<span style=""color:#999999; font-size:9px;""><b>"& left(oRec2("kkundenavn"), 30)&"</b> ("&oRec2("kkundenr")&")</span><br>"
                                        
                                        if len(trim(oRec2("kpafd"))) <> 0 then
                                        strFil_Kon_Txt = strFil_Kon_Txt &  oRec2("kpafd") &"<br />"
                                        end if
                                        
                                        strFil_Kon_Txt = strFil_Kon_Txt &  oRec2("kpadr") &"<br />"_
                                        & oRec2("kpzip") &" "& oRec2("kptown") &"<br />"_
                                        & oRec2("kpland") &"<br />"
                                        end if
                                        
                                        strFil_Kon_Txt = strFil_Kon_Txt &  "</div>"
                                
                                'if len(trim(oRec2("email"))) <> 0 then
                                'strFil_Kon_Txt = strFil_Kon_Txt &  tsa_txt_025 &": <a href='mailto:"&oRec2("email")&"' class=rmenu>"& oRec2("email") &"</a><br>"
                                'end if
                                
                                'if len(trim(oRec2("dirtlf"))) <> 0 then
                                'strFil_Kon_Txt = strFil_Kon_Txt &  tsa_txt_026 &": "& oRec2("dirtlf") &"<br>"
                                'end if
                                
                                'if len(trim(oRec2("mobiltlf"))) <> 0 then
                                'strFil_Kon_Txt = strFil_Kon_Txt &  tsa_txt_027 &": "& oRec2("mobiltlf") &"<br><br />"
                                'end if
                                
                                x = x + 1
                                end if 'oRec2("id") <> 0
                
                lastknavn = oRec2("kkundenavn") 
                oRec2.movenext
                wend
                oRec2.close
                
                
                
                 '*** ÆØÅ **'
                call jq_format(strFil_Kon_Txt)
                strFil_Kon_Txt = jq_formatTxt
                
                Response.Write strFil_Kon_Txt & "<input id=""kperfil_fundet"" type=""hidden"" value='"&x-1&"'>" 
                '&"</div><input type=""text"" value='"&x&"'>"
                
                 
                 
                 
                 
                    
        '********* Henter valgte job i søge filter iforhold til søgning ****************
        case "FN_getCustDesc"
        
          
          %>
         <script src="inc/timereg_2006_load_func.js"></script>
           <%
           
     
        'Response.end
        
        'Response.Write "<select>"
        if len(trim(request("usemrn"))) <> 0 then
        usemrn = request("usemrn")
        else
        usemrn = session("mid")
        end if
        
        
        if len(trim(request("ignprg"))) <> 0 then
        ignProj = request("ignprg")
        else
        ignProj = 0
                
                '**** Henter projektgrupper ***
                '*** ER allerede hentet **'
                '** Side SKAL re-loades ved skift medarbejder ***' eller??
				'call hentbgrppamedarb(usemrn)
        
        end if
        
        
        if len(trim(request("visnyeste"))) <> 0 then
        visnyeste = request("visnyeste")
        else
        visnyeste = 0
        end if
        
        '*** Easyreg ***'
        if len(trim(request("viseasyreg"))) <> 0 then
        viseasyreg = request("viseasyreg")
            if cint(viseasyreg) = 1 then
            easyregFlt = " ,COUNT(a.id) AS antal_aeasy "
            easyWh = " AND tu.easyreg <> 0"
            else
            easyregFlt = ""
            easyWh = ""
            end if
        
        else
        viseasyreg = 0
        easyregFlt = ""
        easyWh = ""
        end if


        if viseasyreg <> 0 then
        Response.Write "Der er valgt Easyreg. <br>Alle Easyreg. aktiviteter p&aring; de valgte kunder vises."
        Response.end

        end if 
        
        '** Dato **'
        ugeStDato = trim(request("ugeStDato"))
        ugeStDato = year(ugeStDato) &"/"& month(ugeStDato) &"/"& day(ugeStDato)


        '*** Smartreg ****'
        if len(trim(request("vissmartreg"))) <> 0 then
        vissmartreg = request("vissmartreg")

            if cint(vissmartreg) = 1 then


                    '**** tilføjer dem til aktive joblisten (timere.g task listen) ****
                    '**** Denne kan slås fra igen og en uges tid Nov. 2011      ******'
                    visGuide = 0
                    strSQLuseguide = "SELECT visguide FROM medarbejdere WHERE mid = "& usemrn
                    
                    'Response.Write strSQLuseguide & "<br>"

                    oRec6.open strSQLuseguide, Oconn, 3
                    if not oRec6.EOF then

                    visGuide = oRec6("visguide")

                    end if
                    oRec6.close


            ugeStDatoSmart = dateAdd("d", -14, request("ugeStDato"))
            ugeStDatoSmart = year(ugeStDatoSmart) &"/"& month(ugeStDatoSmart) &"/"& day(ugeStDatoSmart)

            ugeSlDatoSmart = dateAdd("d", 6, request("ugeStDato"))
            ugeSlDatoSmart = year(ugeSlDatoSmart) &"/"& month(ugeSlDatoSmart) &"/"& day(ugeSlDatoSmart)


            jobnrSmartRegKri = " AND (j.jobnr = 0"
            strSQLjt = "SELECT tjobnr FROM timer WHERE tdato BETWEEN '"& ugeStDatoSmart &"' AND '"& ugeSlDatoSmart &"' AND tmnr = "& usemrn & " AND timer > 0 GROUP BY tjobnr" 
            
            'Response.Write strSQLjt
            'Response.end
            oRec3.open strSQLjt, oConn, 3
            while not oRec3.EOF 

            jobnrSmartRegKri = jobnrSmartRegKri & " OR j.jobnr = "& oRec3("tjobnr")


                    
                    'Response.Write visGuide
                    'Response.end

                    if cint(visGuide) = 11 then 
                    
                    strSQLuseguide = "SELECT id FROM job WHERE jobnr = "& oRec3("tjobnr")
                    'Response.Write strSQLuseguide

                    oRec6.open strSQLuseguide, oConn, 3
                    if not oRec6.EOF then

                    strSQLupdateTuseJob = "UPDATE timereg_usejob SET forvalgt = 1 WHERE jobid = "& oRec6("id") &" AND medarb = "& usemrn
                    'Response.write strSQLupdateTuseJob 
                    
                    oConn.execute(strSQLupdateTuseJob)

                    end if
                    oRec6.close
                    
                   


                    end if
           
            oRec3.movenext
            wend
            oRec3.close

            jobnrSmartRegKri = jobnrSmartRegKri & ")"

            end if


            

        else
        vissmartreg = 0
        end if

       
        
        if cint(visnyeste) = 1 then
        sqlOrderBy = "j.jobstartdato DESC, j.jobnavn"
        lmt = " LIMIT 0, 25"
        dtWhcls = " AND j.jobstartdato <= '"& year(now) &"/"& month(now) &"/"& day(now) &"' "
        
        
        else
            
     
        
        if cint(viseasyreg) = 1 then
        sqlOrderBy = ""
        lmt = ""
        dtWhcls = ""
        else
        sqlOrderBy = ""
        lmt = " LIMIT 0, 100"
        'lmt = ""
        dtWhcls = ""
        end if


        sortby = request("sortby")
        if sortby = 1 then
        sqlOrderBy = sqlOrderBy & "j.jobnavn "
        end if
        
        if sortby = 2 then
        sqlOrderBy = sqlOrderBy & "j.jobnr DESC "
        end if
        
        if sortby = 3 then
        sqlOrderBy = sqlOrderBy & "k.kkundenavn, j.jobnr DESC "
        end if
        
        if sortby = 0 then
        sqlOrderBy = sqlOrderBy & "k.kkundenavn, j.jobnavn "
        end if
      
        
        end if

        
        
        
        
        'Response.Write "usemrn "& usemrn & " ignProj " & ignProj & "easyreg " & viseasyreg
        'Response.end
        
        if len(trim(request("jobsog"))) <> 0 then
        jobsog = request("jobsog")
        else
        jobsog = ""
        end if
       
		if len(trim(jobsog)) <> 0 then 		
		strJobSogKri = " AND (jobnr LIKE '"& jobsog &"%' OR jobnavn LIKE '%"& jobsog &"%' OR kkundenavn LIKE '%"& jobsog &"%' OR kkundenr = '"& jobsog &"') "
		else
		strJobSogKri = ""
		end if				
        
        
        if cint(ignProj) = 0 then
        strSQLwh = " (j.jobstatus = 1 OR j.jobstatus = 3) "& strPgrpSQLkri & "" 
	    else
	    '** Henter alle aktive job + tilbud **'
	    strSQLwh = " (j.jobstatus = 1 OR j.jobstatus = 3) " 
	    end if
        
        if Request.Form("cust") <> 0 AND strJobSogKri = "" then
        strSQLwh = strSQLwh & " AND j.jobknr = "& Request.Form("cust")
        else
        strSQLwh = strSQLwh & " AND j.jobknr <> 0 " '"& Request.Form("cust")
        end if

        

        strSQLwh = strSQLwh & " AND (j.jobstartdato <= '"& ugeStDato &"')"




        
        if cint(ignProj) = 0 AND cint(vissmartreg) <> 1 then
        
            strSQL = "SELECT j.id, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr "& easyregFlt &" FROM timereg_usejob AS tu "_
            &"LEFT JOIN job AS j ON (j.id = tu.jobid) "
            
            if cint(viseasyreg) = 1 then
            strSQL = strSQL &" LEFT JOIN aktiviteter AS a ON (a.job = j.id AND a.easyreg = 1) "
            end if
            
            strSQL = strSQL &"LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
            &"WHERE tu.medarb = "& usemrn &" "& easyWh &" AND "_
            & strSQLwh &" "& strJobSogKri &" "& dtWhcls &" AND kkundenavn <> '' GROUP BY j.id ORDER BY "& sqlOrderBy & lmt
            
        else
            
            '** Timereg use jober ignoreret her. Både for job og easyreg akt.
            if cint(vissmartreg) = 1 then
            strSQLwh = strSQLwh & jobnrSmartRegKri
            end if
            
                
            strSQL = "SELECT j.id, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr "& easyregFlt &" FROM job AS j "_
            &"LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "
            
            if cint(viseasyreg) = 1 then
            strSQL = strSQL &" LEFT JOIN aktiviteter AS a ON (a.job = j.id AND a.easyreg = 1) "
            end if
            
            strSQL = strSQL & "WHERE "& strSQLwh &" "& strJobSogKri &" "& dtWhcls &" AND kkundenavn <> '' GROUP BY j.id ORDER BY " & sqlOrderBy & lmt
            
        end if




        
        
        if len(trim(request("seljobid"))) <> 0 then
        seljobidArr = request("seljobid")
             
            jobids = split(seljobidArr, ",") 
            for j = 0 to UBOUND(jobids)
            seljobid = seljobid & ",#"& jobids(j) &"#"
            next


        else
        seljobid = "#0#"
        end if
        
        'Response.Write strSQL & " seljobid: " & seljobid
        'Response.write "</option></select>"
        'Response.end
        
        lastknr = 0
        jobids_easyreg = 0
        j = 0


        Response.write "<input id=""jobid_nul"" name=""jobid"" type=""hidden"" value=""0"" />"
                 

        oRec.open strSQL, oConn, 3
        while not oRec.EOF
        
            if cint(viseasyreg) = 1 then
            antal_easy = oRec("antal_aeasy")
            else
            antal_easy = 0
            end if
            
            if cint(viseasyreg) = 0 OR (cint(viseasyreg) = 1 AND cint(antal_easy) > 0) then
            
                if instr(seljobid, "#"& oRec("id") &"#") <> 0 OR cint(visGuide) = 11 then
	            jSel = "CHECKED"
                else
	            jSel = ""
	            end if
        	    
        	    
	            delit = "..................................................................................................................................."
                'delit = "...................."
                
                if sortBy <> 2 AND sortBy <> 3 then
                jobnavnid = "<b>"& left(replace(oRec("jobnavn"), "'", ""), 45) & "</b> ("& oRec("jobnr") &")"
                else
                jobnavnid = "<b>"& replace(oRec("jobnr"), "'", "") & "</b> - "& left(oRec("jobnavn"), 45) 
                end if

                if oRec("jobstatus") = 3 then
                jobnavnid = jobnavnid & " - Tilbud"
                end if
                
                '*** ÆØÅ **'
                jq_format(jobnavnid)
                jobnavnid = jq_formatTxt
                
                
                
                len_jobnavnid = len(jobnavnid)
                if len(len_jobnavnid) < 25 then
                lenst = 25
                else
                lenst = len_jobnavnid + 2
                end if 
                
                if len(lenst) <> 0 AND len(len_jobnavnid) <> 0 AND (lenst-len_jobnavnid) > 0 then 
                delit_antal = left(delit, (lenst-len_jobnavnid)) 
                else
                delit_antal = "....."
                end if
                
                jq_format(left(oRec("kkundenavn"), 30))
                knavnOversk = jq_formatTxt
                

                jq_format(left(oRec("kkundenavn"), 15))
                knavn = jq_formatTxt
                
                
                'if (sortby = 0 OR sortby = 3) AND lastknr <> oRec("kkundenr") AND cint(visnyeste) <> 1 then
                'if j <> 0 then
                'Response.Write "<option value='"& oRec("id") &"' disabled></option>" 
                'end if
                'Response.Write "<option value='"& oRec("id") &"' disabled>"& knavn &" ("& oRec("kkundenr") &")</option>" 
                'end if
                
                'Response.Write "<option value='"& oRec("id") &"' "&jSel&" >"& jobnavnid &" "& delit_antal &" "& knavn &" ("& oRec("kkundenr") &")</option>" 
                
                'if right(j, 1) = 0 then
                'Response.flush
                'end if



                if (sortby = 0 OR sortby = 3) AND lastknr <> oRec("kkundenr") AND cint(visnyeste) <> 1 then
                if j <> 0 then
                Response.Write "<br><br>" 
                end if
                Response.Write "<br><span style=""color:#999999; font-size:14px;""><b>"& knavnOversk &" ("& oRec("kkundenr") &")</b></span><br>" 
                end if
                
                Response.Write "<input type=""checkbox"" value='"& oRec("id") &"' "&jSel&" class=""jobid"" id=""jobid_"& oRec("id") &""" name=""jobid"" > "& jobnavnid &" "& delit_antal &"<span style=""color:#999999;""> "& knavn &"..</span><br>" 
                
                if right(j, 1) = 0 then
                Response.flush
                end if



                
                lastknr = oRec("kkundenr")
                
                'jobids_easyreg = jobids_easyreg &", " & oRec("id")
                
                j = j + 1
            
            end if
        
        oRec.movenext
        wend
        oRec.close
        
        if j = 0 then
          'Response.Write "<option value='0'>Ingen aktive job matcher din s&oslash;gning!</option>"
          'Response.Write "<option value='0'>Tjek dine projektgrupper, din personlig aktivliste og job startdato.</option>"  
          Response.Write "Ingen aktive job matcher din s&oslash;gning!<br>"
          Response.Write "Tjek dine projektgrupper, din personlig aktivliste og job startdato."  
        end if
        
       
        
        'Response.Write "</select>"
        '<input id='jobids_easyreg' type='text' value='"& jobids_easyreg &"' />"
        'Response.Write  "<option value='0'>"& strSQL & "</option>"
        end select
        Response.end
        end if  
	
	


    'loadA =  now
    

	
	if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	if len(trim(request("fromsdsk"))) <> 0 then
	fromsdsk = request("fromsdsk")
	fromdeskKomm = request("sdskkomm")
	else
	fromsdsk = 0
	fromdeskKomm = ""
	end if
	
	print = request("print")
	
	thisfile = "timereg_akt_2006"
	func = request("func")
	
	if len(request("stopur")) <> 0 then
	stopur = request("stopur")
	else
	stopur = 0
	end if
	
	
	if len(trim(request("hideallbut_first"))) <> 0 then
    hideallbut_first = request("hideallbut_first")
    else
    hideallbut_first = 0
    end if


    if len(trim(request("sortby"))) <> 0 then
    sortBy = request("sortby")
    response.cookies("tsa")("sortby") = sortby
    else
        if request.cookies("tsa")("sortby") <> "" then
        sortBy = request.cookies("tsa")("sortby")
        else
        sortBy = "1"
        end if
    end if

   select case sortBy
   case 1
   sortBySQL = "tu.forvalgt_sortorder, j.jobnavn,"
   case 2
   sortBySQL = "j.jobnavn,"
   case 3
   sortBySQL = "j.jobslutdato DESC,"
   case 4
   sortBySQL = "k.kkundenavn, j.jobnavn,"
   case else
   sortBySQL = "tu.forvalgt_sortorder, j.jobnavn,"
   end select



	
	call ersmileyaktiv()
	
	'******************************************************
	'*************** Opdaterer timeregistreringer *********
	'******************************************************
	if func = "db" then
	
	if len(trim(request("jobid"))) <> 0 then
	jobid = request("jobid")
	else
	jobid = 0
	end if
	
	'if len(request("showjobinfo")) <> 0 then
	'showjobinfo = request("showjobinfo")
	'else
	'showjobinfo = 0
	'end if
	'Response.Cookies("showjobinfo") = showjobinfo
	
	if len(trim(request("tildelalle"))) <> 0 then
	multitildel = 1
	else
	multitildel = 0
	end if
	
	if len(trim(request("tildeliheledage"))) <> 0 AND request("tildeliheledage") <> 0 then
	tildeliheledage = 1
	else
	tildeliheledage = 0
	end if
	
	
	'*** Valgte medarbejdere ****'
    dim tildelselmedarb
    'redim tildelselmedarb(200)
	
    if len(trim(request("tildelselmedarb"))) <> 0 AND cint(multitildel) = 1 then
    tildelselmedarb = split(request("tildelselmedarb"), ", ")
    else
    tildelselmedarb = split(1, ", ")
    end if
	
	
	            '*** => Timer er indtastet på timereg. siden ***'
	            '*** Eller via stopur                        ***'
	            if stopur <> "1" then 
	            
	                
	                
	                '*** Sætter cookies på faser ****'
	                faseshowVal = split(request("faseshow"), ", ")
	                faserids = split(request("faseshowid"), ", ")
	                for f = 0 to UBOUND(faserids)
	                    'thisFaseVal = request("faseshow_"& faserids(f) &"")
	                    'Response.Write "#"& faserids(f) & "#" &faseshowVal(f) & "<br>"
	                    Response.Cookies("tsa_fase")(""& trim(faserids(f)) &"") = trim(faseshowVal(f))
	                
	                next
	                
	                
	                'Response.end
	                'Response.write "Opdaterer DB<br><br>"
                	
	                strdag = request("FM_start_dag")
	                strmrd = request("FM_start_mrd")
	                straar = request("FM_start_aar")
                	
	                strYear = Request("year")
                	
                	'Response.Write "request(FM_jobid): " & request("FM_jobid") & "<br>"
                	
                	
                	'Response.Write " St: " & request("FM_sttid") & "<br>" 
                	'Response.Write "lasttd:" & request("lasttd")
                	
                    'Response.Write "<br>Test FM "& request("test_fm") & "<br>"

	                jobids = split(request("FM_jobid"), ",")
	                
                    'Response.Write "jobids: "& request("FM_jobid")
                    'Response.end

	                aktids = split(request("FM_aktivitetid"), ",")
	                
	                medarbejderid = request("FM_mnr")
                	
	                tTimertildelt_temp = replace(request("FM_timer"), ", xx,", ";")
	                tTimertildelt_temp2 = replace(tTimertildelt_temp, ", xx", "")
	                tTimertildelt = split(tTimertildelt_temp2, ";") 
                	
                	'Response.Write "tTimertildelt: " & request("FM_timer") & "<br>"
                	'Response.end
                	
	                datoer = split(request("FM_datoer"), ",")
	                
	                'Response.Write "datoer:" & request("FM_datoer")
                	'Response.end
                	
	                tSttid = split(request("FM_sttid"), ",") 
	                tSltid = split(request("FM_sltid"), ",") 
                	
                	
                	'Response.Write tSttid &" til: "& tSltid 
                	'Response.end
                	
                	'Response.Write "Feltnr: "& request("FM_feltnr") & "<hr>"
                    'Response.end

	                feltnr = split(request("FM_feltnr"), ",")
                	for y = 0 to UBOUND(feltnr)
	                high_fetnr = feltnr(y)
	                z = y

	                next
                	
                	
                	'Response.Write  "<br>Antal felter:" & z & "<br>"

                    't = 0
                    'for j = 1 to UBOUND(jobids)
                    't = t + 1
                	'next

                    'Response.Write  "<br>Antal job felter:" & t & "<br>"

	                if len(request("FM_vistimereltid")) <> 0 then
	                visTimerelTid = request("FM_vistimereltid")
	                else
	                visTimerelTid = 0 '** Vis timer 
	                end if
	                
	                
	               
                	
	                 
	                
                                '**** ETS ****'
			                    '***** Kode til beregning af tidslås på timreg. med klokkeslet angivelse her ***'          
			                    '*** Skal array forlænges??? ***'
            				    
				    
				                 
				                 if lto = "ets-track" then 'lto = "intranet - local"
			                       
			                                call etsSub
				            
				                '**** Slut ETS ****'
	                             end if '*** ETS track kode slut
				               
	                
	
			
			
			'*** Validerer ***
			antalErr = 0
			for y = 0 to UBOUND(tTimertildelt)
			
				
				call erDetInt(SQLBless(trim(tTimertildelt(y))))
				if isInt > 0 then
					antalErr = 1
					errortype = 28
					useleftdiv = "t"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end
				end if	
				isInt = 0
				
			
			next
			
			
			
			idagErrTjek = day(now)&"/"&month(now)&"/"&year(now)
			
			for y = 0 to UBOUND(tSttid)	
				
				'Response.write SQLBless3(trim(tSttid(y))) & "<br>"
				'Response.flush
				
				tSttid(y) = replace(tSttid(y), ".", ":")
				tSttid(y) = replace(tSttid(y), ",", ":")
				
				if instr(tSttid(y), ":") = 0 AND len(trim(tSttid(y))) = 4 then
				  
				  left_tid = left(tSttid(y), 2)
				  right_tid = right(tSttid(y), 2)
				  nyTid = left_tid & ":"& right_tid
				  tSttid(y) = nyTid
				
				
				end if 
				  
				
				
				'Response.write "len(trim(tSttid(y))" & len(trim(tSttid(y))) & "<br>"
				if len(trim(tSttid(y))) <> 1 then ' == Slet
				
				
				
				call erDetInt(SQLBless3(trim(tSttid(y))))
				if isInt > 0 then
					antalErr = 1
					errortype = 63
					useleftdiv = "t"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end
				end if	
				isInt = 0
				
				'*** Punktum i angivelse ved registrering af klokkeslet
				'if instr(trim(tSttid(y)), ".") <> 0 OR instr(trim(tSttid(y)), ",") <> 0 then
				'	antalErr = 1
			    '	errortype = 66
				'	useleftdiv = "t"
				'	<include file="../inc/regular/header_lysblaa_inc.asp"--><
				'	call showError(errortype)
				'	response.end
				'end if	
				
				if len(trim(tSttid(y))) <> 0 then
				
				'Response.write idagErrTjek &" "& tSttid(y)&":00" &" - "& isdate(idagErrTjek &" "& tSttid(y)&":00") &"<br>"
					if isdate(idagErrTjek &" "& tSttid(y)&":00") = false OR instr(idagErrTjek &" "& tSttid(y),":") = 0 then
						errTid = tSttid(y)
						antalErr = 1
						errortype = 64
						useleftdiv = "t"
						%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
						call showError(errortype)
						response.end
					end if
				end if
				
				
				end if ' tSltid(y) <> 0 
				
			next
			
			
			
			
			for y = 0 to UBOUND(tSltid)	
				
				tSltid(y) = replace(tSltid(y), ".", ":")
				tSltid(y) = replace(tSltid(y), ",", ":")
				
				if instr(tSltid(y), ":") = 0 AND len(trim(tSltid(y))) = 4 then
				  
				  left_tid = left(tSltid(y), 2)
				  right_tid = right(tSltid(y), 2)
				  nyTid = left_tid & ":"& right_tid
				  tSltid(y) = nyTid
				
				end if 
				
				
				call erDetInt(SQLBless3(trim(tSltid(y))))
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
				'if instr(trim(tSltid(y)), ".") <> 0 OR instr(trim(tSltid(y)), ",") <> 0 then
				'	antalErr = 1
				'	errortype = 66
				'	useleftdiv = "t"
				'	include file="../inc/regular/header_lysblaa_inc.asp"--
				'	call showError(errortype)
				'	response.end
				'end if	
				
				
				if len(trim(tSltid(y))) <> 0 then
				'Response.Write "len(trim(tSltid(y))) " & len(trim(tSltid(y))) & "<br>"
					if isdate(idagErrTjek &" "& tSltid(y)&":00") = false OR instr(idagErrTjek &" "& tSltid(y),":") = 0  then
						errTid = tSltid(y)
						antalErr = 1
						errortype = 64
						useleftdiv = "t"
						%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
						call showError(errortype)
						response.end
					end if
				end if
				
			next
			
			
			
	
	else
	
	
	'**** Hvis indlæsnng sker fra Stopur  ***'
	
	
	
    antalErr = 0
	jobidSQLkri =  ""
	antalsids = split(request("sids"), ", ")
	for t = 1 to UBOUND(antalsids)
	
	if t = 1 then
	jobidSQLkri = jobidSQLkri & " w.id = "& antalsids(t)
	else
	jobidSQLkri = jobidSQLkri & " OR w.id = "& antalsids(t)
	end if  
	
	next
	 
	medid = request("medid")
	medarbejderid = medid
	sjobids = ""
	saktids = ""
	timerthis = ""
	datoerthis = ""
	sids = ""
	
	'*** Henter jobids og aktids ***'
	strSQLstop = "SELECT w.id AS wid, w.jobid, w.aktid, w.sttid, w.sltid, w.kommentar, w.incident, "_
	&" s.dato, s.tidspunkt FROM stopur w "_
	&" LEFT JOIN sdsk s ON (s.id = w.incident) WHERE ("& jobidSQLkri &") AND timereg_overfort = 0"
	
	'Response.write strSQLstop &"<br>"
	'Response.flush 
	
	oRec2.open strSQLstop, oConn, 3
	while not oRec2.EOF
	
	sids = sids & oRec2("wid") &","
	sjobids = sjobids & oRec2("jobid") &","
	saktids = saktids & oRec2("aktid") &","
	
	    if isDate(oRec2("sttid")) AND isDate(oRec2("sltid")) then 
	    tidThis = datediff("n", oRec2("sttid"),oRec2("sltid"), 2,2)
	    else
	    tidThis = 0 
	    end if
	    
	    
	    call timerogminutberegning(tidThis)
	    timerthis = timerthis & thoursTot&"."& tminProcent &","
	    
	   
	datoerthis = datoerthis & oRec2("sttid") &","   
	datoogklokkeslet = left(weekdayname(weekday(oRec2("dato"))), 3) &". "&formatdatetime(oRec2("dato"), 2) &" "& left(formatdatetime(oRec2("tidspunkt"), 3), 5)
    kommentars = kommentars & "<br><br> == Incident Id "& oRec2("incident") &" == <br>Opr. "& datoogklokkeslet &"<br>Udført ("& oRec2("wid") &") "& oRec2("sttid") &" - "& oRec2("sltid") &" <br> "& replace(oRec2("kommentar"), "'", "&#39;") &"#,#"
	
	oRec2.movenext
	Wend
	oRec2.close
	
	if len(kommentars) > 3 then
	kommentars_len = len(kommentars)
	kommentars_left = left(kommentars, kommentars_len - 3) 
	kommentars = kommentars_left
	end if
	
	
	if len(sjobids) > 1 then
	sjobids_len = len(sjobids)
	sjobids_left = left(sjobids, sjobids_len - 1) 
	sjobids = sjobids_left
	end if
	
	if len(saktids) > 1 then
	saktids_len = len(saktids)
	saktids_left = left(saktids, saktids_len - 1) 
	saktids = saktids_left
	end if
	
	if len(timerthis) > 1 then
	timerthis_len = len(timerthis)
	timerthis_left = left(timerthis, timerthis_len - 1) 
	timerthis = timerthis_left
	end if
	
	if len(datoerthis) > 1 then
	datoerthis_len = len(datoerthis)
	datoerthis_left = left(datoerthis, datoerthis_len - 1) 
	datoerthis = datoerthis_left
	end if
	
	if len(sids) > 1 then
	sids_len = len(sids)
	sids_left = left(sids, sids_len - 1) 
	sids = sids_left
	end if
	
	kommentarer = split(kommentars, "#,#")
	entrysids = split(sids, ",")
	jobids = split(sjobids, ",")
	aktids = split(saktids, ",")
	useTimer = split(timerthis, ",")
	useDatoer = split(datoerthis, ",")
	
	
	end if '***Fra  Stopur****'
	
	
	
	
	
	
	
	
	
	'*********************************************************************************
	'***** Indlæser timer array ****'
	'***** Indlæser i db ***
    '*********************************************************************************
	
	
	'Response.Write "<br><br>antalErr" & antalErr & "<br>"
	
	
	if antalErr = 0 then
	
	
	
	
	
	
	for m = 0 to UBOUND(tildelselmedarb)
	
	'Response.Write tildelselmedarb(m) & "<br>"
	
	if cint(multitildel) = 1 then
	medarbejderid = tildelselmedarb(m) 
	else
	medarbejderid = medarbejderid
	end if
	
	    
        
	
	    '**** Finder medarbejder timepris ************
		call mNavnogKostpris(medarbejderid)
	    
        'Response.Write "her" & jobids(0)

	    for j = 0 to UBOUND(jobids)
			    
			    '*****************************************************'
				'*** Henter Akt og Job oplysninger                  **'
				'*** Finder evt. alternativ medarbejder timepris    **'
				'*** på den valgte aktivitet                        **'
				'*****************************************************'
				
                 'foundone_0 = 0
                 'intTimepris_0 = 0
                 'intValuta_0 = 0
                
				
				strSQL = "SELECT job.id AS jobid, aktiviteter.navn, fakturerbar, jobnavn, jobnr, "_
				&" kkundenavn, jobknr, fastpris, job.serviceaft"_
				&" FROM aktiviteter, job, kunder "_
				&" WHERE aktiviteter.id = "& aktids(j) &" AND job.id = "& jobids(j) &" AND kunder.Kid = jobknr "
				
				
				
				'Response.write "<br><br>"& strSQL &"<br><br>"
				'Response.flush				
				
				oRec.Open strSQL, oConn, 3
				if Not oRec.EOF then
					
							 intjobid = oRec("jobid")
							 strJobnavn = oRec("Jobnavn")
					 		 strJobknavn = oRec("kkundenavn")
							 strJobknr = oRec("Jobknr") 	
							 strAktNavn = oRec("navn")
							 strFakturerbart = oRec("fakturerbar")
							 intServiceAft = oRec("serviceaft")
							 intJobnr = oRec("jobnr")
							 
							 '**** 2009 t.tfaktim er = a.fakturerbar (akt. type) også i timer tabellen. **'
							 tfaktimvalue = strFakturerbart
							 strFastpris = oRec("fastpris")
						
					        
					        '*** tjekker for alternativ timepris på aktivitet
							call alttimepris(aktids(j), jobids(j), medarbejderid, 1)
							
							'** Er der alternativ timepris på jobbet
							if foundone = "n" then
								'if cint(foundone_0) <> 1 then
                                call alttimepris(0, jobids(j), medarbejderid, 1)
                                'intTimepris_0 = intTimepris
                                'intValuta_0 = intValuta
                                'foundone_0 = 1
                                'else
                                'intTimepris = intTimepris_0 
                                'intValuta = intValuta_0
                                'end if
							end if
							
							'**************************************************************'
							'*** Hvis timepris ikke findes på job bruges Gen. timepris fra '
							'*** Fra medarbejdertype, og den oprettes på job **************'
							'**************************************************************'
							if foundone = "n" then
							intTimepris = replace(tprisGen, ",", ".") 
							intValuta = valutaGen
							
							strSQLtpris = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) "_
							&" VALUES ("& intjobid &", 0, "& medarbejderid &", 0, "& intTimepris &", "& intValuta &")"
							
							oConn.execute(strSQLtpris)
							
							end if
							
				end if
				oRec.Close
	
	        
	        
	        
	        
	        if stopur <> "1" then 
		    '****************************************************'
			'***** Dag 1 - 7 Indlæses i db fra Timereg **********'
			'****************************************************'
			
            '*** Dage 1-7 på hver aktlinje ***'
			if j = 0 then
			ystart = 0
			else
			ystart = (j * 7) 
			end if
			
			yslut = (ystart + 6)

            'Response.Write high_fetnr & "<br>"
			'Response.Write ystart & "<br>"
            'Response.Write yslut & "<br>"
            'Response.flush

			isAktMedidWrt = ""
			multiDagArrUse = 0
			
			'for y = 0 to UBOUND(tTimertildelt)
			for y = ystart to yslut 'AND y <= z  
			    
			    '*** tjekker om felt id liggen i den rigtige aktlinie *'
			    '**  DVS mellem 11 - 17, eller 21-27 etc..          ***'
				if y >= ystart AND y <= yslut then
				
                'Response.Write y & ": "& feltnr(y) &"<br>"
                'Response.Write "feltnr(y):"& feltnr(y)
                'Response.flush
				feltvalThis = cint(replace(feltnr(y), "_", ""))
				feltvalThis = cstr(feltvalThis)
				
				
				strKomm = replace(request("FM_kom_"&feltvalThis), "'", "&#39;")
				
			    '*** Tjekker at komm. er udfyldt ved pre-def aktiviteter ***'
				call akttyper2009prop(tfaktimvalue)
				if len(trim(tTimertildelt(y))) <> 0 AND aty_pre <> 0 AND tTimertildelt(y) > -9000 then
				        
				        if cdbl(replace(tTimertildelt(y), ".", ",")) <> cdbl(aty_pre) AND len(trim(strKomm)) = 0  then
				        'Response.Write cdbl(replace(tTimertildelt(y), ".", ",")) &" <> "& cdbl(aty_pre) & "<br>"
				        antalErr = 1
						errortype = 131
						useleftdiv = "t"
						%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
						call showError(errortype)
						response.end
						end if
				
				end if
				
				if len(request("FM_off_"&feltvalThis)) <> 0 then
				offentlig = request("FM_off_"&feltvalThis)
				else
				offentlig = 0
				end if
				
				if len(request("FM_bopal_"&feltvalThis)) <> 0 then
				bopal = request("FM_bopal_"&feltvalThis)
				else
				bopal = 0
				end if
				
				
				if len(request("FM_destination_"&feltvalThis)) <> 0 then
				destination = request("FM_destination_"&feltvalThis)
				else
				destination = ""
				end if
				
				if len(trim(tTimertildelt(y))) <> 0 then
				tTimertildelt(y) = SQLBless(tTimertildelt(y))
				else
					if visTimerelTid <> 0 then
					tTimertildelt(y) = -9002 '-2 Opdater
					else
					tTimertildelt(y) = -9001 '-1 Slet
					end if
				end if
				
				
				    'Response.Write "fltnt:"& feltnr(y) &"aktid: "& aktids(y) &" timer:" & tTimertildelt(y) & " Dato: " & datoer(y) & " tSttid(y): "& tSttid(y) &" medarbejderid: "& medarbejderid&"<br>"
                    'Response.flush
				
			    
				    '**** * Multitildel for flere dage ****'
				    if instr(tTimertildelt(y), "*") <> 0 then
    				    
				        lngt = len(tTimertildelt(y))
				        st = instr(tTimertildelt(y), "*")
				        multiDagArr = mid(tTimertildelt(y), st+1, lngt)
				        'multiDagArr = (multiDagArr/5) * 7
				            's = 0
				            'if s = 100 then
				            if instr(multiDagArr, ",") <> 0 OR instr(multiDagArr, ".") <> 0 then
				            antalErr = 1
						    errortype = 135
						    useleftdiv = "t"
						    %><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
						    call showError(errortype)
						    response.end
						    end if
    				    
    				    
				        findTimerForStar = left(tTimertildelt(y), (st-1))
				        multiDagArrUse = multiDagArr - 1
    				    
				        if left(trim(findTimerForStar), 1) = "." then
			            useTimer = "0"&trim(findTimerForStar)
			            else
			            useTimer = findTimerForStar
			            end if    
    				    
    				    
				    '*** tildeler timer på valgte dage, ved multiDage funktion tildeles ikke på hellgidage ***'
				    '*** Hvad med personlige fridage, der læses ike timer ind på dage ved indtasting i periode *''
				    '**  via * hvis medarbejderen ikke 
				    '*** ikke har normeret tid de pågældende dage **'
				    muDag = 0
				    for muDag = 0 to multiDagArrUse 
				    'Response.Write  "muDag " & muDag & "<br>"
    				'Response.write AktMedidWrt  & "<br>"
    				
				    useDato = dateAdd("d", muDag, datoer(y))
				    '**** Indlæser kun timer på dage hvor den valgte medarbejder har arbejdsdag, og ikke på helligdage **'
    				
				    ntimPer = 0
				    call normtimerPer(medarbejderid, useDato, 0)
    				
				        'Response.Write "timer:" & useTimer & " Dato: " & useDato & " ntimPer:"& ntimPer &" medarbejderid: "& medarbejderid&"<br>"
                        'Response.flush
    				
				        if cint(ntimPer) <> 0 then
    				    useTimerThis = useTimer
    				    else
    				    useTimerThis = 0
    				    end if
    				   
    				      
            				
				            '**** Indlæser timer i DB, alm timereg. i timer *****'
				            call opdaterimer(aktids(j), strAktNavn, tfaktimvalue, strFastpris, intJobnr, strJobnavn, strJobknr,_
				            strJobknavn, medarbejderid, strMnavn,_
				            useDato, useTimerThis, strKomm, intTimepris,_
				            dblkostpris, offentlig, intServiceAft, strYear,_
				            tSttid(y), tSltid(y), visTimerelTid, stopur, intValuta, bopal, destination)
            				
    				    
    				    
				        isAktMedidWrt = isAktMedidWrt & ",#"&aktids(j)&"_"&medarbejderid&"#"
				        'end if
    				    
				        findTimerForStar = ""
				        
    				 
				    next 
				    
    				
    				
    				
				    else
				    '*** else flere dage * ==> dvs. alm
    				        
    				        
				            if left(trim(tTimertildelt(y)), 1) = "." then
			                useTimer = "0"&trim(tTimertildelt(y))
			                else
			                useTimer = tTimertildelt(y)
			                end if    
    			            
    				        muDag = 0
				            useDato = datoer(y)
				            
				                 
				                 
				             '*** ETS HER ? ***  
				                 
				            
				            
				            'Response.write AktMedidWrt  & "<br>"
    				        
				            '*** Er akt/medarb. allerede opdateret via multiDatoer?****' 
				            '** Så skal de ikke opdateresd igen, da de så overskriver den allerede opdaterede værdi ***'
				            if instr(isAktMedidWrt, ",#"&aktids(j)&"_"&medarbejderid&"#") = 0 then
				            
				           
    	                    '**** Indlæser timer i DB, alm timereg. i timer *****'
				            
				            if aktids(j) <> 0 AND len(trim(aktids(j))) <> 0 then
				            call opdaterimer(aktids(j), strAktNavn, tfaktimvalue, strFastpris, intJobnr, strJobnavn, strJobknr,_
				            strJobknavn, medarbejderid, strMnavn,_
				            useDato, useTimer, strKomm, intTimepris,_
				            dblkostpris, offentlig, intServiceAft, strYear,_
				            tSttid(y), tSltid(y), visTimerelTid, stopur, intValuta, bopal, destination)
    				        
    				        
    				        isAktMedidDatoWrt = isAktMedidDatoWrt & ",#"&aktids(j)&"_"&medarbejderid&"_"&useDato&"#"
				            end if
				            
				            
				            end if
    				
    				
				    end if
			   
				        
				end if 'y <= ystart
			
			next 'y
			
			else ' = fra Stopur
			    
			        
			        tSttid = ""
	                tSltid = "" 
                	visTimerelTid = 0 '** Vis timer
                	offentlig = 0
                	strYear = year(useDatoer(j))
                    
                    
                    destination = ""
	                bopal = 0
               
               
                '****************************************************************'
                '*** Er uge alfsuttet af medarb, er smiley og autogk slået til **'
                '****************************************************************'
                regdato = useDatoer(j)
                firstWeekDay = weekday(regdato, 2)
                
                select case firstWeekDay
                case 1
                stTil = 0
                slTil = 6
                case 2
                stTil = 1
                slTil = 5
                case 3
                stTil = 2
                slTil = 4
                case 4
                stTil = 3
                slTil = 3
                case 5
                stTil = 4
                slTil = 2
                case 6
                stTil = 5
                slTil = 1
                case 7
                stTil = 6
                slTil = 0
                end select
                
                stwDayDato = dateadd("d", -stTil, regDato)
                slwDayDato = dateadd("d", slTil, regDato)
                
                regdatoStSQL = year(stwDayDato) &"/"& month(stwDayDato) &"/"& day(stwDayDato)
                regdatoSlSQL = year(slwDayDato) &"/"& month(slwDayDato) &"/"& day(slwDayDato)
                
                
                'Response.Write "firstWeekDay "& firstWeekDay
                
                call afsluger(medarbejderid, regdatoStSQL, regdatoSlSQL)
                
                
                erugeafsluttet = instr(afslUgerMedab(medarbejderid), "#"&datepart("ww", regdato,2,2)&"_"& datepart("yyyy", regdato) &"#")
                
                'Response.Write useDatoer(j) & "<br>"
                'Response.Write "erugeafsluttet --" & erugeafsluttet  &"<br>"
                'Response.Write "autogk --" & autogk  &"<br>"
                'Response.Write "smilaktiv  --" & smilaktiv   &"<br>" 
                'Response.flush
                'Response.end
                
                call lonKorsel_lukketPer(regdato)
              
                 if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", regdato) = year(now) AND DatePart("m", regdato) < month(now)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", regdato) < year(now) AND DatePart("m", regdato) = 12)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", regdato) < year(now) AND DatePart("m", regdato) <> 12) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", regdato) > 1))) OR cint(lonKorsel_lukketIO) = 1 then
              
                ugeerAfsl_og_autogk_smil = 1
                else
                ugeerAfsl_og_autogk_smil = 0
                end if 
		        
		                
		                
		        '*** Tjekker sidste fakdato ***'
		        'if len(trim(intjobid)) <> 0 then
		        'intjobid = intjobid
		        'else
		        'intjobid = 0
		        'end if
		        
		        
		        
		        lastFakdato = "01/01/2002"
		        strSQL = "SELECT fakdato FROM fakturaer WHERE jobid = "& intjobid &" AND faktype = 0 ORDER BY fakdato DESC LIMIT 0,1"
		        
		       
		        oRec.open strSQL, oConn, 3
		        if not oRec.EOF then
		        lastFakdato = oRec("fakdato")
		        end if
		        oRec.close
		        
		        if ugeerAfsl_og_autogk_smil = 1 OR cdate(lastFakdato) >= cdate(regdato) then
		        
		        %>
			    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			    <% 
			    
			    if lastFakdato = "01/01/2002" then
			    lastFakdato = "(ingen)"
			    end if
			    
			    useleftdiv = "c2"
			    errortype = 117
			    call showError(errortype)
		        
		        Response.end
		        end if
               
                strKomm = kommentarer(j)
			    
			    
			    'Response.Write "tSttid "& tSttid &",  tSltid "& tSltid &", visTimerelTid "& visTimerelTid &", stopur "& stopur
			    'Response.flush
			    
			    'aktids(j) ", strAktNavn, tfaktimvalue, strFastpris, intJobnr, strJobnavn, strJobknr,_
				'strJobknavn, medarbejderid, strMnavn,_
				'useDatoer(j), useTimer(j), strKomm, intTimepris,_
				'dblkostpris, offentlig, intServiceAft, strYear,
			    
			    '**** Indlæser timer i DB *****'
				call opdaterimer(aktids(j), strAktNavn, tfaktimvalue, strFastpris, intJobnr, strJobnavn, strJobknr,_
				strJobknavn, medarbejderid, strMnavn,_
				useDatoer(j), useTimer(j), strKomm, intTimepris,_
				dblkostpris, offentlig, intServiceAft, strYear, tSttid, tSltid, visTimerelTid, stopur, intValuta, bopal, destination)
 
			'*** Opdaterer stopur timer overført *****'
			strSQLstopU = "UPDATE stopur SET timereg_overfort = 1 WHERE id = "& entrysids(j) 
	        oConn.execute(strSQLstopU)
			       
			end if '**** Fra Stopur ***'
			
		
		
		next 'j
		
		next 'm (multideldel)
	
	    
	    '** Tjekker indlæs værdier **' 
	    '** Response.end
	
	
	
	
	
	           


           




	
	end if '** AntalErr **'
	
	
	 '** hvis der skal tjekkes values fra timeindlæsningen brus .end her ***'
	 '** 
     'Response.end
	
	    '**** Opdaterer stopurs siden / Redirect til Timregside ****'
	    if stopur = "1" then 
	    Response.Redirect "stopur_2008.asp?reload=1"
	    else
        select case lto
        case "jep"
        Response.Redirect "timereg_akt_2006.asp?showakt=1&hideallbut_first=2"   
        case else
	    Response.Redirect "timereg_akt_2006.asp?showakt=1" 
        end select   
	    end if
	
	
	
	end if '** func = db **'
	
	
	
	
	
	'******************************************************'
	'************ Opdater Smiley **************************'
	'******************************************************'
	 if func = "opdatersmiley" then
	 call opdaterSmiley
     
     
     
    
        
        '*** Vender tilbage til timereg ****'
        %>
        <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
        <div style="position:absolute; top:100px; left:200px; width:400px; padding:10px; border:10px #CCCCCC solid; background-color:#ffffff;">
        
        <table cellspacing="0" cellpadding="10" border="0" bgcolor="#ffffff" width=100%>
	    <tr>
	        <td bgcolor="#ffffff" style="padding-top:15px;">
	        <h4><%=tsa_txt_001%>:</h4></td>
	        <td valign="top" align=right bgcolor="#ffffff">
		    &nbsp;
		    </td>
	        </tr>
	        <tr>
	        <td colspan=2 valign=top>
            <%if instr(formatdatetime(cDateUge, 2), "1899") <> 0 then %>
            Der er opstået en fejl.<br />
            Du har glemt at vælge uge!<br /><br />
            <br />
            <a href="javascript:history.back()" class=vmenu><< Tilbage</a>
            <br /><br />&nbsp;

            <%else%>

	        <%=tsa_txt_002%>:<br /> <b><%=weekdayname(datepart("w",cDateUge,1)) %> d. 
	        <%=formatdatetime(cDateUge, 2) &" "& tsa_txt_003 %>. 
	        <%=formatdatetime(cDateUge, 3) %></b><br />
	        <br />
	        <%=tsa_txt_004 %>:<br /> <b><%=weekdayname(datepart("w",cDateAfs,1)) %> d. 
	        <%=formatdatetime(cDateAfs, 2) &" "& tsa_txt_003 &". "& formatdatetime(cDateAfs, 3) %></b>
	        <br /><br />
	        <b><%=tsa_txt_005 &" "& datepart("ww", cDateUgeTilAfslutning, 2, 2)%> <%=smileysttxt %></b>
	        
          
	        
	        </td>
	    </tr>
	<tr>
	<td>
	 <a href="timereg_akt_2006.asp?fromsdsk=<%=fromsdsk%>" class=vmenu><%=tsa_txt_006 %> >></a>
        
		</td>
	</tr>
      <%end if %>
</table>
</div>
        <%
        Response.end
	   
	    
	
	end if '*** Opdater smiley func **'
	
	
	
	

    if func = "opdaterjlist" then

    'Response.Write "her" & request("tuid")

    tuid_progrp = request("tuid_progrp")
    tuid_progrp_use_rdir = tuid_progrp
    'Response.Write tuid_progrp
    'Response.end

    tuid_progrp_alle = request("tuid_progrp_alle")
    tuid_mid = request("tuid_mid")

    if tuid_progrp_alle = "1" then ' hele gruppen

        call medarbiprojgrp(tuid_progrp, tuid_mid)

         
         strSQLUpTuidnulstil = "UPDATE timereg_usejob SET forvalgt = 0 WHERE medarb = 0  "& replace(medarbgrpIdSQLkri, "mid", "medarb")
         'Response.Write strSQLUpTuidnulstil
         'Response.end
         oConn.execute(strSQLUpTuidnulstil)
    
    else

         strSQLUpTuidnulstil = "UPDATE timereg_usejob SET forvalgt = 0 WHERE medarb = "& tuid_mid
         oConn.execute(strSQLUpTuidnulstil)

    end if

  
    dtNow = year(now) &"/"& month(now) &"/"& day(now)
    tuids = split(request("tuid"), "XX,")

    for t = 0 TO UBOUND(tuids) - 1

        if len(trim(tuids(t))) <> 0 then

        tuids(t) = replace(tuids(t), ",", "")

        strSQLUpTuid = "UPDATE timereg_usejob SET forvalgt = 1, forvalgt_dt = '"& dtNow &"', forvalgt_af = "& session("mid") &" WHERE id = "& tuids(t)
        oConn.execute(strSQLUpTuid)

            '*** Opdater helegruppen ***'
              if tuid_progrp_alle = "1" then ' hele gruppen
              strSQLUpTuidseljobid = "SELECT jobid FROM timereg_usejob WHERE id = "& tuids(t)
              oRec4.open strSQLUpTuidseljobid, oConn, 3
              if not oRec4.EOF then
                    
                      strSQLUpTuid_all = "UPDATE timereg_usejob SET forvalgt = 1, forvalgt_dt = '"& dtNow &"', forvalgt_af = "& session("mid") &" WHERE (jobid = "& oRec4("jobid") &") AND (medarb = 0  "& replace(medarbgrpIdSQLkri, "mid", "medarb") & ")"
                      oConn.execute(strSQLUpTuid_all)

              end if
              oRec4.close

              end if

        end if

	next

    
	

    'Response.end   
    'Response.write "timereg_akt_2006.asp?showakt=1&usemrn="&tuid_mid&"&tildel_sel_pgrp="&tuid_progrp_use_rdir
    'Response.end
    Response.redirect "timereg_akt_2006.asp?showakt=1&usemrn="&tuid_mid&"&tildel_sel_pgrp="&tuid_progrp_use_rdir
	
	Response.end   
	end if '* opdaterjlist
	
	   
	%>
	
	
	
	
	<%
	
	%><!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<SCRIPT language=javascript src="inc/timereg_2006_func.js"></script>
   
	
    <%call browsertype()
    
    if browstype_client <> "ip" then%>

	<!--#include file="../inc/regular/topmenu_inc.asp"-->
    <div id="sindhold" style="position:absolute; left:0; top:42; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	<%call tsamainmenu(1)%>
	</div>
	
	
	
	<div id="sekmenu" style="position:absolute; left:15; top:97; visibility:visible;">
	<%call tregsubmenu() %>
	</div>
	
	
    <%else

    end if
	
	
	
	
	'**** Aktiviteter på det valgte job ***********'
	'** Ved forste logon i sessionen må cookie ikke vælges
	'** Således at man altid får vist sig selv når men logger på selvom man har tastet timer inde
	'** for en anden sidst man var logget på ***'
	if session("forste") <> "n" then
	session("forste") = "j"
	end if
	'****** Findes nu i login.asp ****'
	
	
	'**** Jobid, showakt og usemrn sættes i sharepoint.asp link og hentes i timreg__2006_fs.asp **'
	if session("fromsharepoint") = "a6196f3646c2e60eddb95cd2134d457f" then
	
	showakt = 1
    jobid = session("shpjobid")
    usemrn = session("mid")      	
	
	
	else
	             
    

     '******************************************************************************
    '****************** Medarb *******************************************************
    '******************************************************************************


                '** Valgt mid ****
                if len(request("usemrn")) <> 0 then
                usemrn = request("usemrn")
                else
                    if len(request.Cookies("tsa")("usemrn")) <> 0 AND session("forste") = "n" then
                    usemrn = request.Cookies("tsa")("usemrn")
                    else
                    usemrn = session("mid")
                    end if
                end if
            	
            	response.cookies("timereg_2006")("usemrn") = usemrn

    
    '******************************************************************************
    '****************** Job *******************************************************
    '******************************************************************************


             

             
                '****** Sætter job aktive i timereg_usejob baseret på de valgte job i selelct boks + de åbne job på listen **'
                 '*** Det valgte job ****'
                'Response.Write "<br><br><br><br><br><br><br><br><br><br><br><br><br>jobid: " &request("jobid")
                'Response.write "<br>request(FM_viskunetjob); " & request("FM_viskunetjob")


                if (len(request("jobid")) <> 0 AND request("FM_ch_medarb") <> "1") then 'OR (len(request("jobid")) <> 0 AND request("FM_ch_medarb") = "1" AND request("FM_mtilfoj") = "1") then

                if len(trim(request("FM_viskunetjob"))) <> 1 then
                jobid = replace(request("seljobid"), "0,", "") 
                else
                jobid = "" '** Ryd liste, vis kun et job 
                end if
                
                '''jobid = replace(request("seljobid"), ",999999", "") 
                jobidsSeljids = split(request("jobid"), ",")
                for i = 0 TO UBOUND(jobidsSeljids)
                    
                    if instr(jobid, "#"& jobidsSeljids(i) & "#,") = 0 AND instr(jobid, ","& jobidsSeljids(i) & ",") = 0 then
                      jobid = jobid &",#"& trim(jobidsSeljids(i)) & "#" 
                    end if

                next

                ''jobid = replace(jobid, ",##", "")
                jobid = replace(jobid, " ", "")
                jobid = replace(jobid, "#", "")
                jobid = trim(jobid)
                jobid = "0,"&jobid &",999999,"

               

                    '** Nulstillerforvalgt ALLE job for denne medarbejder ****'
                    strSQLUpdFvOff = "UPDATE timereg_usejob SET forvalgt = 0 WHERE medarb = "& usemrn &"" 
                    oConn.Execute(strSQLUpdFvOff)
                    '******************************************************'
                   

                else	
                    
                        
                        '*** Ellers åbnes det job der sidst er reg. timer på (evt, 10) ***'
                        'strSQLlastjobid = "SELECT j.id AS jid FROM timer "_
                        '&" LEFT JOIN job j ON (j.jobnr = tjobnr) WHERE tmnr = "& session("mid") &" ORDER BY tid DESC LIMIT 1"
                        ''Response.Write "<br><br><br><br><br><br>"& strSQLlastjobid
                        ''Response.flush 
                        
                        'oRec.open strSQLlastjobid, oConn, 3
                        'if not oRec.EOF then
                        'jobid = oRec("jid")
                        'end if
                        'oRec.close
                        
                        if len(trim(usemrn)) <> 0 then
                        usemrn = usemrn
                        else
                        usemrn = session("mid")
                        end if
                           
                        jobid = "0"
                        strSQLfv = "SELECT jobid FROM timereg_usejob WHERE medarb = "& usemrn & " AND forvalgt = 1"

                        'Response.Write "<br><br><br><br><br>sdfsdfsdf: "& strSQLfv
                        'Response.flush

                        oRec5.open strSQLfv, oConn, 3
                        while not oRec5.EOF
            
                        jobid = jobid &","& oRec5("jobid") 

                        oRec5.movenext
                        wend
                        oRec5.close

                        jobid = jobid & ",999998,"
                        'end select
            		    
                    'end if
                end if





                '******************************************'
                '** Trimmer og analyser. jobid ********'
                '**************************************'
                '** Er der valgt et eller flere job ***'
                '** Sætter forvalgt = 1 i timereg_usejob **'
                '******************************************'

                'dim jobidss
                'redim jobidss(2)
                if instr(jobid, ",") <> 0 then

                seljobidSQL = " j.id = 0" 

                        jobids = split(jobid, ",") 
                        for j = 0 to UBOUND(jobids)
                        


                                if len(trim(jobids(j))) <> 0 then
                                thisJobId = jobids(j)
                                else
                                thisJobId = 0
                                end if

                                thisJobId = replace(thisJobId, "#", "")

                                if instr(jobisWrt,",#"& thisJobId &"#") = 0 then 

                                seljobidSQL = seljobidSQL & " OR j.id = "& thisJobId  
                                jobisWrt = jobisWrt &",#"& thisJobId &"#" 

                                end if

                                '** Sætter forvalgt job for denne medarbejder ****'
                                strSQLUpdFvOn = "UPDATE timereg_usejob SET forvalgt = 1 WHERE medarb = "& usemrn &" AND jobid = "& thisJobId
                                'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br>"& strSQLUpdFvOn
                                'Response.flush
                                oConn.Execute(strSQLUpdFvOn)
                                '*************************************************'

                        next

              
                else

                        seljobidSQL = seljobidSQL & " j.id = "& jobid &" "
                        Redim jobids(0) 
                        jobids(0) = jobid

                                '** Sætter forvalgt job for denne medarbejder ****'
                                strSQLUpdFvOn = "UPDATE timereg_usejob SET forvalgt = 1 WHERE medarb = "& usemrn &" AND jobid = "& jobid
                                oConn.Execute(strSQLUpdFvOn)
                                '*************************************************'


                end if


            	
            	
                    if len(trim(request("showakt"))) <> 0 then
                    showakt = request("showakt")
                   
                    else
                        if len(request.Cookies("showakt")) <> 0 then
                        showakt = request.Cookies("showakt")
                        else
                             'if lto = "demo" then
                             'showakt = 0
                             showakt = 1
                             'else
                             'showakt = 1
                             'Response.Write "her"
                             'end if
                        end if
                    end if
                    
               
    
    
	
	end if 'sharepointlink
	
	'showakt = 0
	
	if request.Cookies("showjobinfo") <> "" then
	showjobinfo = request.Cookies("showjobinfo")
	else  
	    select case lto
	    case "demo"
	    showjobinfo = 0
	    case else
	    showjobinfo = 1
	    end select
	end if
	
	if len(request("FM_vistimereltid")) <> 0 then
	visTimerElTid = request("FM_vistimereltid")
	else
	
		if len(request.Cookies("tsa")("timereltid")) <> 0 then
		visTimerElTid = request.Cookies("tsa")("timereltid")
		else
		visTimerElTid = 0
		end if
		
	end if
	
	'** tidslås ***
	if len(request("FM_chkfor_ignorertidslas")) <> 0 then
	    
	   
	    if len(request("FM_igtlaas")) <> 0 then
	    ignorertidslas = 1
	    else
	    ignorertidslas = 0
	    end if
	
	else
	    
	    if request.Cookies("tsa")("igtidslas") <> "" then
		ignorertidslas = request.Cookies("tsa")("igtidslas")
		else
		ignorertidslas = 0
		end if
	   
	end if
	
	
	'**** Job og Akt. periode ***'
	if len(request("FM_chkfor_ignJobogAktper")) <> 0 then
	    
	   
	    if len(request("FM_ignJobogAktper")) <> 0 then
	    ignJobogAktper = 1
	    else
	    ignJobogAktper = 0
	    end if
	
	else
	    
	    if request.Cookies("tsa")("ignJobogAktper") <> "" then
		ignJobogAktper = request.Cookies("tsa")("ignJobogAktper")
		else
		ignJobogAktper = 0
		end if
	   
	end if
	
    'ignJobogAktper = 0

	response.Cookies("tsa")("ignJobogAktper") = ignJobogAktper
	response.Cookies("tsa")("igtidslas") = ignorertidslas

    if len(trim(jobid)) <> 0 then
    jobid = jobid
    else
    jobid = 0
    end if
	
    'response.Cookies("tsa")("jobid") = jobid
	response.Cookies("showakt") = showakt
	response.Cookies("tsa")("usemrn") = usemrn
	response.Cookies("tsa")("timereltid") = visTimerElTid
	response.Cookies("tsa").expires = date + 32
	
	
	call erSDSKaktiv()
	
	%>
	
	
	
	
	
	
	<%
	level = session("rettigheder")
	timenow = formatdatetime(now, 3)
	%>
	
	
	<%
	
	'**** Divs ***'
	
	'Select case showakt
	'case 1
	dKalVzb = "visible"
	dKalDsp = ""
	dTimVzb = "visible"
    dTimDsp = ""
	
    dAfstemVzb = "hidden"
	dAfstemDsp = "none"
	dStempelVzb = "hidden"
	dStempelDsp = "none"
	sVzb = "visible"
	sDsp = ""
	urVzb = "visible"
	urDsp = ""
	phVzb = "visible"
	phDsp = ""
	
	fiVzb = "Visible"
	fiDsp = ""
	
	jinf_knap_vzb = "visible"
	jinf_knap_dsp = ""
	
	
	
	'case 2
	'dKalVzb = "hidden"
	'dKalDsp = "none"
	'dTimVzb = "hidden"
	'dTimDsp = "none"
	'dJobVzb = "hidden"
	'dJobDsp = "none"
	'dAfstemVzb = "visible"
	'dAfstemDsp = ""
	'dStempelVzb = "hidden"
	'dStempelDsp = "none"
	'sVzb = "hidden"
	'sDsp = "none"
	'urVzb = "hidden"
	'urDsp = "none"
	'phVzb = "hidden"
	'phDsp = "none"
	
    'fiVzb = "hidden"
	'fiDsp = "none"
	'jinftdBG = "#EFF3FF"
	'jinf_knap_vzb = "hidden"
	'jinf_knap_dsp = "none"
	
	'case 3
	'dKalVzb = "hidden"
	'dKalDsp = "none"
	'dTimVzb = "hidden"
	'dTimDsp = "none"
	'dJobVzb = "hidden"
	'dJobDsp = "none"
	'dAfstemVzb = "hidden"
	'dAfstemDsp = "none"
	'dStempelVzb = "visible"
	'dStempelDsp = ""
	'sVzb = "hidden"
	'sDsp = "none"
	'urVzb = "hidden"
	'urDsp = "none"
	'phVzb = "hidden"
	'phDsp = "none"
	
	'jinftdBG = "#EFF3FF"
    'fiVzb = "hidden"
	'fiDsp = "none"
	'jinf_knap_vzb = "hidden"
	'jinf_knap_dsp = "none"
	
	
	'case else
	
	'dKalVzb = "hidden"
	'dKalDsp = "none"
	'dTimVzb = "hidden"
	'dTimDsp = "none"
		
	'dAfstemVzb = "hidden"
	'dAfstemDsp = "none"
	'dStempelVzb = "hidden"
	'dStempelDsp = "none"
	
	'sVzb = "hidden"
	'sDsp = "none"
	'urVzb = "hidden"
	'urDsp = "none"
	'phVzb = "hidden"
	'phDsp = "none"
	
	
    'fiVzb = "visible"
	'fiDsp = ""
	
	'jinf_knap_vzb = "visible"
	'jinf_knap_dsp = ""
	
	'end select
    
    
    if browstype_client = "ip" then
    dKalVzb = "hidden"
	dKalDsp = "none"
    end if%>
	
	<% 
	if cint(fromsdsk) = 1 then
	addVal = 20
	else
	addVal = 100
	end if
	
	'**srctip **'
    
   
    
    %>
     <div id="kalender" style="position:absolute; left:780px; top:153px; width:210px; visibility:<%=dKalVzb%>; display:<%=dKalDsp%>;">
	<!--#include file="../inc/regular/calender_2006.asp"-->
	<!--#include file="inc/timereg_dage_2006_inc.asp"-->
	</div>
    
   <%
   varTjDatoUS_man = convertDateYMD(tjekdag(1))
	varTjDatoUS_tir = convertDateYMD(tjekdag(2))
	varTjDatoUS_ons = convertDateYMD(tjekdag(3)) 
	varTjDatoUS_tor = convertDateYMD(tjekdag(4)) 
	varTjDatoUS_fre = convertDateYMD(tjekdag(5))
	varTjDatoUS_lor = convertDateYMD(tjekdag(6))
	varTjDatoUS_son = convertDateYMD(tjekdag(7))
   
    '** redirect til afstem toto, hvis det er den er vist sidst. ***
    if showakt = "5" then
    Response.Redirect "afstem_tot.asp?usemrn="&usemrn&"&show=5&varTjDatoUS_man="&varTjDatoUS_man 
    Response.end
    end if

    
    call treg_3menu(thisfile)
    %>
    
    
    
   
    
   
	
    
    <div id="loadbar" style="position:absolute; display:; visibility:visible; top:260px; left:200px; width:300px; background-color:#ffffff; border:10px #9ACD32 solid; padding:10px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" /><br />&nbsp;
  
	</td></tr>
    <tr><td colspan=2>  <div id="load_cdown">Forventet loadtid: 4-7 sekunder...</div></td></tr>
    </table>

	</div>
	
	
	
  
	
	
	
	
    
    
	
	<%
	
	
	
	
	 
 '************************************************************************************************** 
  'Opbygger timereg SQL states
  '**************************************************************************************************
 	            
 	        
				
				
				if len(trim(request("FM_kontakt"))) <> 0 then
				show_sogjobnavn_nr = request("FM_sog_job_navn_nr")
				response.cookies("tsa")("jobsog") = show_sogjobnavn_nr
				else
				    if request.cookies("tsa")("jobsog") <> "" then
				    show_sogjobnavn_nr = request.cookies("tsa")("jobsog")
				    else
				    show_sogjobnavn_nr = ""
				    end if
				end if
				
				
				if len(request("FM_kontakt")) <> 0 then
				sogKontakter = request("FM_kontakt")
				
				
				response.cookies("tsa")("kontakt") = sogKontakter
	            response.cookies("tsa").expires = date + 65
				
				else
				    
				        if request.Cookies("tsa")("kontakt") <> "" _
				        AND request.Cookies("tsa")("kontakt") <> 0 _
				        then
				        sogKontakter = request.Cookies("tsa")("kontakt")
				        else
				        sogKontakter = 0
                        end if
				end if
				
				
				if len(request("FM_ignorer_projektgrupper")) <> 0 then
				ignorerProgrp = request("FM_ignorer_projektgrupper") '1
					if cint(ignorerProgrp) = 1 then
					selIgn = "CHECKED"
					else
					selIgn = ""
					end if
				else
				ignorerProgrp = 0
				selIgn = ""
				end if
				
				



                if len(request("FM_smartreg")) <> 0 then
				    intSmartreg = request("FM_smartreg") '1
					if cint(intSmartreg) = 1 then
					selSmart = "CHECKED"
					else
					selSmart = ""
					end if
					
				else
                    
                    '***** Vis guide forste gang medarb logger på efter ny timereg. side ***'
                    visGuide = 0
                    strSQLvisGuide = "SELECT visguide FROM medarbejdere WHERE mid = " & usemrn
                    oRec5.open strSQLvisGuide, oConn, 3
                    if not oRec5.EOF then

                    visGuide = oRec5("visguide")

                    end if
                    oRec5.close

                    if cint(visGuide) = 11 then
                    intSmartreg = 1
				    selSmart = "CHECKED"
                    sogKontakter = 0


                    '**** Opdaterer visguide ****
                    strSQLvisGuide = "UPDATE medarbejdere SET visguide = 12 WHERE mid = " & usemrn
                    oConn.execute(strSQLvisGuide)


				    else
				            if request.cookies("smartreg") = "1" then
				            intSmartreg = 1
				            selSmart = "CHECKED"
				            else
				            intSmartreg = 0
				            selSmart = ""
				            end if
				    end if


				   
				end if


                if cint(visGuide) <> 11 then '*** Tvungevisguide og sidste uges job


                if len(request("FM_nyeste")) <> 0 then
				    intNyeste = request("FM_nyeste") '1
					if cint(intNyeste) = 1 then
					selNye = "CHECKED"
					else
					selNye = ""
					end if
					
				else
				    if request.cookies("nyeste") = "1" then
				    intNyeste = 1
				    selNye = "CHECKED"
				    else
				    intNyeste = 0
				    selNye = ""
				    end if
				end if
				
				if len(request("FM_easyreg")) <> 0 then
				    intEasyreg = request("FM_easyreg") '1
					if cint(intEasyreg) = 1 then
					selEasy = "CHECKED"
					else
					selEasy = ""
					end if
					
				else
				    if request.cookies("easyreg") = "1" then
				    intEasyreg = 1
				    selEasy = "CHECKED"
				    else
				    intEasyreg = 0
				    selEasy = ""
				    end if
				end if

                if len(trim(request("tildeliheledageSet"))) <> 0 then
                tildeliheledageVal = 0
                else
                tildeliheledageVal = 1
                end if    

                 if len(request("FM_classicreg")) <> 0 then
				    intClassicreg = request("FM_classicreg") '1
					if cint(intClassicreg) = 1 then
					selClassic = "CHECKED"
					else
					selClassic = ""
					end if
					
				else
				    if request.cookies("classicreg") = "1" OR (intSmartreg = 0 AND intEasyreg = 0 AND intNyeste = 0) then
				    intClassicreg  = 1
				    selClassic = "CHECKED"
				    else
				    intClassicreg  = 0
				    selClassic = ""
				    end if
				end if


                end if


				


				if selNye = "CHECKED" then
				sotDivDSP = "none"
				sotDivVzb = "hidden"
				else
				sotDivDSP = ""
				sotDivVzb = "visible"
				end if
				
				
				if len(request("sort")) <> 0 then
				sortVal = request("sort")
				else
				    if request.cookies("treg_sort") <> "" then
				    sortVal = request.cookies("treg_sort")
				    else
				    select case lto
				    case "dencker"
				    sortVal = 3
				    case else 
				    sortVal = 0
				    end select
				    end if
				end if
				
				
				sort1CHK = ""
				sort2CHK = ""
				sort3CHK = ""
				sort0CHK = ""
				
				select case sortVal
				case 1
				sort1CHK = "CHECKED"
				case 2
				sort2CHK = "CHECKED"
				case 3
				sort3CHK = "CHECKED"
				case else
				sort0CHK = "CHECKED"
				end select
				
				
	call hentbgrppamedarb(usemrn)
				
	
	'*******************************************************************************'
	'*******************************************************************************'
    '************************* Filter *********************************************'
	'*******************************************************************************'
	
   
	call filterheaderid(80,20,745,pTxt,fiVzb,fiDsp,"filterTreg", "relative")%>

     <form action="timereg_akt_2006.asp" id="filterkri">
	<table cellspacing=0 width=100% cellpadding=1 border=0>
	
   
	
    <input type="hidden" name="FM_ch_medarb" id="FM_ch_medarb" value="0">
    <input type="hidden" name="tildeliheledageSet" id="Hidden3" value="<%=tildeliheledageVal%>">
    <input type="hidden" name="dtson" id="datoSonSQL" value="<%=now%>">
	<%if level = 1 then
	SQLkriEkstern = ""
    end if
	
	''*** Finder medarbejdere i de progrp hvor man er teamleder ***'
	call projgrp(-1,level,session("mid"),1)
	    
	     medarbgrpIdSQLkri = "AND (mid = "& session("mid")
    
	    
	    for p = 0 to prgAntal
	     
	     if prjGoptionsId(p) <> 0 then
	        call medarbiprojgrp(prjGoptionsId(p), session("mid"))
	     end if
	    
	    next 
	    
	     medarbgrpIdSQLkri = medarbgrpIdSQLkri & ")"
	    
	strSQLmids = medarbgrpIdSQLkri '" AND mid = "& usemrn
	%>
	<tr bgcolor="#ffffff">
	<td valign=top rowspan=2 style="padding-top:5px; padding-bottom:5px;"><b><%=tsa_txt_333 &" "& tsa_txt_077 %>:</b> (<%=tsa_txt_357 %>)
	<br />
				<%
					strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe FROM medarbejdere WHERE mansat <> 2 "& strSQLmids &" GROUP BY mid ORDER BY Mnavn"
					
					'Response.Write strsQL
					'Response.flush
					
					%>
					<select name="usemrn" id="usemrn" style="font-size: 9px; width:250px;"><!-- onchange="submit(); -->
					<%
					
					oRec.open strSQL, oConn, 3
					while not oRec.EOF 
					
					if cint(oRec("Mid")) = cint(usemrn) then
					rchk = "SELECTED"
					else
					rchk = ""
					end if%>
					<option value="<%=oRec("Mid")%>" <%=rchk%>><%=oRec("mnavn")%></option>
					<%
					
					
					oRec.movenext
					wend
					oRec.close
				%></select>
				
				 <!--
				 &nbsp;&nbsp;&nbsp;&nbsp;<a href="../timereg2006/timereg_2006_fs.asp" class=red>tilbage til den gamle timereg. side..</a>
				 -->
				
				<%
	'**** Finder kunder til dd *****'
	'**** Tag højde for ignorer projektgrupper? **'
	if cint(ignorerProgrp) = 0 then
	
	    
		
		strKundeKri = " AND ("
		strSQL3 = "SELECT j.id, medarb, jobid, easyreg, kid, jobknr FROM timereg_usejob "_
		&" LEFT JOIN job j ON (j.id = jobid) "_
		&" LEFT JOIN kunder ON (kid = j.jobknr) "_
		&" WHERE medarb = "& usemrn &" AND kid <> '' AND jobstatus = 1  GROUP BY kid ORDER BY kid"
	    'response.write strSQL3
		'Response.flush
		oRec3.open strSQL3, oConn, 3
		
		while not oRec3.EOF 
		strKundeKri = strKundeKri & " kid = "& oRec3("kid") & " OR "
		oRec3.movenext
		wend 
		
		oRec3.close 
    else
    strKundeKri = " AND (kid <> 0 OR "
    end if
	%>
	<br /><br /><b><%=tsa_txt_075 %>:</b><br />
	
	<%
	
	    if len(strKundeKri) <> 0 then
		strKundeKri = strKundeKri &" kid = 0)"
		else
		strKundeKri = strKundeKri &" AND (kid = 0)"
		end if
		
		ketypeKri = " ketype <> 'e'"
		
		
		if cint(ignorerProgrp) = 0 then
	    strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE "& ketypeKri &" "& strKundeKri &" GROUP BY kid ORDER BY Kkundenavn"
	    else
	    strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder, job WHERE "& ketypeKri &" "& strKundeKri &" AND jobstatus = 1 AND jobknr = kid GROUP BY kid ORDER BY Kkundenavn"
	    end if
	    'Response.write strSQL
		'Response.flush
	
	%>
	<select name="FM_kontakt" id="FM_kontakt" style="font-size:9px; width:250px;" size=8 onchange="renssog();"> <!-- onChange="renssog()" --> 
		
		<%if sogKontakter = 0 then
		allSel = "SELECTED"
		else
		allSel = ""
		end if %>
		
		<option value="0" <%=allSel %>><%=tsa_txt_076 %> (maks. 100 job)</option>
		<%
		
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(sogKontakter) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn") & " ("& oRec("Kkundenr")&")"%></option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
		
		</select>
       	
       	<br /><br />
       	<b><%=tsa_txt_078 %>:</b> (<%=tsa_txt_073 %>)<br />
		<input type="text" name="FM_sog_job_navn_nr" id="FM_sog_job_navn_nr" value="<%=show_sogjobnavn_nr%>" style="font-family:arial; font-size:10px; width:250px;"> 
		<br /><input type="checkbox" name="FM_ignorer_projektgrupper" id="FM_ignorer_projektgrupper" value="1" <%=selIgn%>> <%=tsa_txt_074%> 
		<br /><a href="javascript:popUp('guiden_2006.asp?mid=<%=usemrn%>','650','620','150','120');" target="_self" class=rmenu>
		(+ <%=lcase(tsa_txt_082)%>)
		</a>
		<br /><br>
        <%'*** Browsertype **' %>

            <div style="position:relative; padding:5px; border:1px #cccccc solid; width:237px; background-color:#eff3ff">
            <b>Visning:</b><br />
            <input id="FM_nyeste" name="FM_nyeste" type="radio" value="1" <%=selNye%> /> <%=tsa_txt_334 %>
            
            <%if (lto = "dencker" AND level = 1) OR lto = "outz" OR lto = "intranet - local" then %>
            <br /><input id="FM_easyreg" name="FM_easyreg" type="radio" value="1" <%=selEasy%> /> <%=tsa_txt_358 %> 
            <%end if %>
		
			
            <br /><input id="FM_smartreg" name="FM_smartreg" type="radio" value="1" <%=selSmart%> /> <%=tsa_txt_373 %>
            <br /><input id="FM_classicreg" name="FM_classicreg" type="radio" value="1" <%=selClassic%> /> <%=tsa_txt_374 %> 
         
				
            </div>

				</td>
				
				
	    <td valign=top style="height:200px; padding-top:5px;"><b><%=tsa_txt_336 %>:</b>
	    
        
	    <%if level <= 3 OR level = 6 then %>
	    &nbsp;<a href="jobs.asp?menu=job&func=opret&id=0&int=1&rdir=treg" class=rmenu><%=tsa_txt_335 %> >></a><br />
        <%end if %>
	    
	   
	    <div id="sorterdiv" style="display:<%=sotDivDsp%>; visibility:<%=sotDivVzb%>; padding:3px 3px 3px 0px; border:0px #cccccc solid;">
	        <b><%=tsa_txt_341 %>:</b> <input id="sort1" name="sort" type="radio" value=1 <%=sort1CHK%> /> <%=tsa_txt_342 %>&nbsp;
             <input id="sort2" name="sort" type="radio" value=2 <%=sort2CHK%> /> <%=tsa_txt_343 %>&nbsp;
              <input id="sort0" name="sort" type="radio" value=0 <%=sort0CHK%> /> <%=tsa_txt_344&", "&lcase(tsa_txt_342) %>&nbsp;
              <input id="sort3" name="sort" type="radio" value=3 <%=sort3CHK%> /> <%=tsa_txt_344&", "&lcase(tsa_txt_343) %>
              </div>
	   
	   
       <!--
	   <textarea id="fajl" cols="40" rows="40"></textarea>
       -->
       
       
       <div name="div_jobid" id="div_jobid" style="width:450px; height:250px; overflow:auto; padding:10px; border:1px #CCCCCC solid; font-size:11px;">
	    <!-- henter job fra jquery -->
	    </div>
       
        <span style="font-size:9px; color:#999999;">Der vises maks. 100 job i job-oversigten.</span>
        <br /><input type="checkbox" name="FM_viskunetjob" id="FM_viskunetjob" value="1" />Ryd din aktive jobliste
        
        <!--<br />
	    <input type="checkbox" name="FM_mtilfoj" id="FM_mtilfoj" value="1" />Tilføj valgte job til ny medarbejder, ved skift medarbejder
            -->
         <input id="jobid_nul" name="xjobid" type="hidden" value="0" /> <!-- skal vist altid være 0 for at undgå fejl, hvis der ikke er valgt noget -->

         <%'** trimmer jobid igen for dubletter **' 
         seljobidUse = "0"
         seljobidUseWrt = ""
         jobidstjk = split(jobid, ",")
         
         for i = 0 TO UBOUND(jobidstjk)

            if i = 0 then '** fjerner 0
            seljobidUse = ""
            end if
        
         if instr(seljobidUseWrt, ",#"& jobidstjk(i) &"#") = 0 then
         seljobidUse = seljobidUse &","& jobidstjk(i)
         seljobidUseWrt = seljobidUseWrt & ",#"& jobidstjk(i) &"#"
         end if

         next
         %>

         <input id="seljobid" name="seljobid" type="hidden" value="<%=seljobidUse %>,999997" />
         <input id="showakt" name="showakt" type="hidden" value="1" />
	    
	  
              
	    </td>
				
				
				
	</tr>
	
	
	
	    <tr><td align=right style="height:30px;">
	    <input type="submit" id="sogsubmit" value="<%=tsa_txt_314 &" "& lcase(tsa_txt_244) %> >>">
	    </td></tr>
	    </table>
	    </td>
	    
	</tr>
    </table>
        <!-- filter header -->
	</td>
    </tr>
    
    </table>
   </form>
     
 </div>
	
	
	
	
	<!--- tjekker adgang via projektgrupper --->
	
	<%
	'if ignorerProgrp = 1 then
	
	select case lto 
	case "acc"
	pgeditok = 1
	case else
	    if level <= 2 OR level = 6 then
	    pgeditok = 1
	    else
	    pgeditok = 0
	    end if
	end select
	
	
	strJobmedadgang = "#0#"
	
		strSQL = "SELECT j.id AS id FROM job j "_
		&" WHERE "& varUseJob &" (j.jobstatus = 1 OR j.jobstatus = 3) "& strPgrpSQLkri & "  ORDER BY j.id" 
		
        'AND (j.jobstartdato <= '"& varTjDatoUS_man &"')
        'Response.Write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"& strSQL
        'Response.flush

		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		
		strJobmedadgang = strJobmedadgang &",#"& oRec("id")&"#"
		
		oRec.movenext
		wend
		
		oRec.close
	
	
	
    
    
    
    if (instr(strJobmedadgang, "#"& jobids(0) &"#") <> 0) then
		projektgrpOK = 1
		
			strprojektgrpOK = "&nbsp;"
		
		else
		projektgrpOK = 0
			if pgeditok = 1 then
			strprojektgrpOK = "<b>"& tsa_txt_351 &"</b><br>"& tsa_txt_352 &" <a href=""javascript:popUp('tilknytprojektgrupper.asp?id="&jobid&"&medid="&usemrn&"','600','500','150','120');"" target=""_self""; class=vmenu>"& tsa_txt_353 &"..</a>"  
			'"<b>Du har ikke adgang til dette job.</b><br>Du kan tilmelde dig jobbet ved at tilføje en projektgruppe du er medlem af. Du kan tilføje en projektgruppe ved at klikke <a href=""javascript:popUp('tilknytprojektgrupper.asp?id="&jobid&"&medid="&usemrn&"','600','500','150','120');"" target=""_self""; class=vmenu>her..</a>"
			else
			
			strprojektgrpOK = "<b>"& tsa_txt_351 &"</b>" '"<b>Du har ikke adgang til dette job.</b>"
			end if
		end if
	
	
	'Response.Write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>projektgrpOK "& projektgrpOK
	
	'else
	'projektgrpOK = 1
	'end if
	







    '*************************************************************'
    '************ Højre mnubar på timereg. siden *****************'
    '**************** Aktive job, job til afv. +TODO LIST ********'
    '*************************************************************'
      if browstype_client <> "ip" then %>



     
       <div id="hmenu" style="position:absolute; left:779px; top:418px; width:212px; visibility:visible; display:; z-index:2;">
   <table cellpadding=0 cellspacing=0 border=0>
        <tr>
            <td align=center id="a_denneuge" style="white-space:nowrap; width:100px; padding:1px; background-color:#EFF3FF; border-right:1px #d6dff5 solid; border-bottom:0px;">
			<a href="#" id="A1" class=rmenu>+ Igangv.</a></td>
             <td align=center id="a_todos" style="white-space:nowrap; width:120px; padding:1px; background-color:#EFF3FF; border-right:1px #d6dff5 solid; border-bottom:0px;">
			<a href="#" id="A2" class=rmenu>+ ToDo's</a></td>
			</td>
            <%if level <= 2 then%>
			<td align=center id="a_tildel" style="white-space:nowrap; width:120px; padding:1px; background-color:#EFF3FF; border-right:1px #d6dff5 solid; border-bottom:0px;">
			<a href="#" id="A3" class=rmenu>+ Tildel job</a></td>
			<%end if%>
        </tr>
    </table>
    </div>

        <div id="div_tildel" style="position:absolute; left:779px; top:433px; width:210px; height:280px; visibility:hidden; display:none; z-index:100; border:1px #cccccc solid; padding:2px; background-color:#ffffff; overflow:auto;">
	     <form action="timereg_akt_2006.asp?func=opdaterjlist" method="post">
         <table cellpadding=2 cellspacing=0 border=0 width=100%>
         
         <tr bgcolor="#5C75AA"><td class=alt style="width:165px;" colspan=2><b>
         Tildel job til aktivjobliste:</b>
         </td>
       
         </tr>

         <tr><td class=lille colspan=2>
        
         <%
         if len(trim(request("tildel_sel_pgrp"))) <> 0 then
         tildel_sel_pgrpSEL = trim(request("tildel_sel_pgrp"))
         else
                if request.cookies("tsa")("td_prg_sel") <> "" then
                tildel_sel_pgrpSEL = request.cookies("tsa")("td_prg_sel")
                else
                tildel_sel_pgrpSEL = 0
                end if
         end if
         
         Response.cookies("tsa")("td_prg_sel") = tildel_sel_pgrpSEL

         strSQLtdgrp = "SELECT id, navn FROM projektgrupper WHERE id <> 0 ORDER BY navn " 
         oRec3.open strSQLtdgrp, oConn, 3
         %>
          Projektgrupper:<br /> <select id="tildel_sel_pgrp" name="tildel_sel_pgrp" style="width:180px; font-size:9px;">
         <%
         while not oRec3.EOF 
            
            if cint(tildel_sel_pgrpSEL) = cint(oRec3("id")) then
            selPrgrp = "SELECTED"
            else
            selPrgrp = ""
            end if
            
            %>
         <option value="<%=oRec3("id") %>" <%=selPrgrp %>><%=oRec3("navn") %></option>
         <%
         oRec3.movenext
         wend
         oRec3.close
         %>
       
         
         </select>

         
         <br /><br />
         Medarbejder(e):<br /> 
         <div id="div_tildeljoblisten_msel"></div>
        
        </td></tr>

         <tr><td colspan=2 valign=top class=lille>
         <div id="div_tildeljoblisten"></div>
         </td></tr>
        
         </table>
         </form>
        </div>

            
          <%select case lto
          case "dencker", "intranet - local"
          jdu_lmt = 20
          case "epi"
          jdu_lmt = 100
          case else
          jdu_lmt = 50
          end select %>        



        <div id="div_denneuge" style="position:absolute; left:779px; top:433px; width:212px; height:500px; visibility:<%=jinf_knap_vzb%>; display:<%=jinf_knap_dsp%>; z-index:100; border:1px #cccccc solid; padding:2px; background-color:#ffffff; overflow:auto; z-index:99999999994444;">
	     <table cellpadding=2 cellspacing=0 border=0 width=100%>
         <form method="post" action="#">
         <tr bgcolor="#5C75AA"><td class=alt><b>
         <%=tsa_txt_376 %>:</b>
         </td></tr>
         <tr><td class=lille>
         Viser aktive job (maks <%=jdu_lmt %>) hvor Du er: <br /><input type="checkbox" value="1" id="denneuge_jobans" name="denneuge_jobans" /> Jobansvarlig / jobejer
         <br />
         Vis seneste 2 uger + 
         <select id="denneuge_ignrper" style="font-size:9px; width:50px;">
            <option value="1">1 uge</option>
            <option value="2">2 uge</option>
            <option value="3">3 uge</option>
            <option value="4">4 uge</option>
         </select> frem
         <input type="hidden" value="<%=jdu_lmt %>" id="denneuge_limit"/>
         <!--
         <input type="checkbox" value="1" name="denneuge_ignrper" id="denneuge_ignrper" /> Slutdato uge <%=datepart("ww", tjekdag(1), 2,2) %> <br /> 
         ellers +1 md. og forrige 3 md.
         -->
         <br /><input type="checkbox" value="1" name="denneuge_alfa" id="denneuge_alfa" />  Sortér alfabetisk
         </td></tr>
         <input type="hidden" value="<%=lto%>" name="dennuge_lto" id="dennuge_lto" />
         

         <tr><td>

         <div id="jobidenneuge">
        </div>

         </td></tr>
         </form>
         </table>
        </div>


	    
        <div id="div_todolist" style="position:absolute; left:779px; top:433px; width:212px; height:200px; visibility:hidden; display:none; z-index:100; border:1px #cccccc solid; padding:2px; background-color:#ffffff; overflow:auto;">
	    
       
        <table cellpadding=2 cellspacing=0 border=0 width=100%>
         <tr bgcolor="#5C75AA"><td id="todotd" colspan="2" style="padding:3px; height:20px;" class=alt><b>ToDo liste</b> (seneste 20)</td></tr>
         <tr><td colspan="2"> <a href="webblik_todo.asp?nomenu=1" class=rmenu target="_blank">Se / rediger ToDo listen >></a></td></tr>
         <tr><td class=lille style="padding-left:5px;">ToDo's</td><td class=lille align=right style="padding-right:5px;">Afslut</td></tr>
        <% 
        
        idag_365 = dateadd("d", -365, now)
        idag_365 = year(idag_365) &"/"& month(idag_365) & "/" & day(idag_365)

        strSQLtodo = "SELECT t.navn, t.id, t.parent, t.dato, t.level, t.delt, "_
	    &" t.afsluttet, t.sortorder, trn.medarbid, trn.todoid, t.tododato, t.forvafsl, t.public, t.parent "_
	    &" FROM todo_rel_new trn "_
	    &" LEFT JOIN todo_new t ON (t.id = trn.todoid AND t.afsluttet <> 1)"_
	    &" WHERE ((medarbid = "& useMrn &" AND trn.todoid = t.id) OR (t.public = 1)) AND afsluttet = 0 AND t.dato > '"& idag_365 &"' GROUP BY t.id ORDER BY t.dato DESC, t.id DESC LIMIT 20"

        'AND parent = 0
        'Response.Write strSQLtodo

        oRec.open strSQLtodo, oConn, 3
        t = 0
        while not oRec.EOF

        select case right(t, 1)
        case 0,2,4,6,8
        tBg = "#FFFFFF"
        case else
        tBg = "#EFf3ff"
        end select


        if t = 0 AND cdate(oRec("dato")) = cdate(date()) then
        tbg = "#FFFFe1"
        end if

        if len(trim(oRec("navn"))) > 200 then
        todoTxt = left(oRec("navn"), 200) & ".."
        else
        todoTxt = oRec("navn")
        end if

        if len(todoTxt) > 15 then
        len_todoTxt = len(todoTxt)
        left_todoTxt = "<b>"& left(todoTxt, 15) & "</b>"
        right_todoTxt = right(todoTxt, len_todoTxt- 15)
        
        todoTxt = left_todoTxt & right_todoTxt 

        else

        todoTxt = "<b>"& todoTxt &"</b>"

        end if

        if oRec("public") = 0 then
        privatTxt = ""
        else
        privatTxt = "(off.lig)"
        end if

        if oRec("forvafsl") = 1 then
        todoDatoTxt = "Forv. udført: " & day(oRec("tododato")) & ". "& left(monthname(month(oRec("tododato"))), 3)&" "& year(oRec("tododato"))
        else
        todoDatoTxt = ""
        end if

       %><tr bgcolor="<%=tBg%>">
       <td class=lille valign=top style="border-top:1px #cccccc solid; padding:4px 5px 4px 5px;">
       <%if oRec("parent") <> 0 then 
       
       strSQLpar = "SELECT navn FROM todo_new WHERE id = "& oRec("parent")
       oRec2.open strSQLpar, oConn, 3
       if not oRec2.EOF then%>
       <%=left(oRec2("navn"), 20) %> >> 
       <%end if
       oRec2.close %>


       <%end if %>
       
       <%=todoTxt%> - <span style="color:#999999; font-size:9px;"><%=day(oRec("dato")) & ". "& left(monthname(month(oRec("dato"))), 3)&" "& year(oRec("dato"))%></span>
       <span style="color:darkred; font-size:9px;"><%=todoDatoTxt %></span></td>
       <td align=right valign=top style="border-top:1px #cccccc solid; padding:2px 5px 2px 5px;">
       <%if oRec("public") <> 1 then %>
       <input id="todoid_<%=oRec("id")%>" class="todoid" type="checkbox" value="1" />
       <%else %>
       <span style="color:#5582d2; font-size:9px;"><%=privatTxt %></span>
       <%end if %></td>
           
       </tr>
       
       <%
        t = t + 1
        oRec.movenext
        wend
        oRec.close
        %>
        <tr>
         <td valign=top colspan=2 style="border-top:1px #cccccc solid; padding:4px 5px 4px 5px;">&nbsp;</td>
        </tr>
        </table>

       
        </div>


	<%
    '****************** Slut højreDIV --->
    end if   'if browstype_client = "ip" then



	if cint(projektgrpOK) = 0 then
	
	
 
            
	
	'Response.end
	else%>
	 
	
	        

	
	     <%
		
				strSQL = "SELECT j.id, j.jobnavn, j.jobnr, k.kid, k.kkundenavn, k.kkundenr, "_
				&" j.jobans1, j.jobans2, jobstartdato, jobslutdato, j.beskrivelse, j.lukafmjob, "_
				&" m1.mnavn AS medjobans1, m1.mnr AS medjobans1mnr, m1.init AS medjobans1init, "_
				&" m2.mnavn AS medjobans2, m2.mnr AS medjobans2mnr, m2.init AS medjobans2init, "_
				&" m3.mnavn AS medkundeans1, m3.mnr AS mnr, m3.init AS init, j.jobstatus, "_
				&" j.budgettimer, j.ikkebudgettimer, j.jobtpris, j.fastpris, j.jobfaktype, s.navn AS saftnavn, "_
				&" s.status AS saftstatus, s.aftalenr AS saftnr, s.id AS aftid, jo_bruttooms, "_
				&" kundekpers, rekvnr, kk.navn AS kkpersnavn, k.sdskpriogrp, j.usejoborakt_tp, j.job_internbesk, j.restestimat, j.stade_tim_proc, k.useasfak "_
				&" FROM job j, kunder k, medarbejdere m "_
				&" LEFT JOIN medarbejdere m1 ON m1.mid = j.jobans1"_
				&" LEFT JOIN medarbejdere m2 ON m2.mid = j.jobans2"_
				&" LEFT JOIN medarbejdere m3 ON m3.mid = k.kundeans1"_
				&" LEFT JOIN serviceaft s ON (s.id = j.serviceaft)"_
				&" LEFT JOIN kontaktpers kk ON (kk.id = kundekpers)"_
				&" WHERE j.id = " & jobids(0) & " AND k.kid = j.jobknr"
				
				'Response.write strSQL
				'Response.flush
				oRec.open strSQL, oConn, 3
				if not oRec.EOF then
				
				'*** Jobbet er blevet lukket ved timregistrering ***
				if (cint(oRec("jobstatus")) <> 1 AND cint(oRec("jobstatus")) <> 3) OR cdate(oRec("jobstartdato")) > cDate(now) then 'tjekdag(7)
				
				Response.Redirect "timereg_akt_2006.asp?jobid=0&showakt=1"
                
				
				end if
				
				'** til timepriser **'
				j_intjobtype = oRec("fastpris")
				j_usejoborakt_tp = oRec("usejoborakt_tp") 
				j_timerforkalk = oRec("budgettimer") + oRec("ikkebudgettimer")
				j_budget = oRec("jobtpris")
				
				j_jobnavn_nr = oRec("jobnavn") & " ("& oRec("jobnr") &")"

                if oRec("jobstatus") = 3 then
                j_jobnavn_nr = j_jobnavn_nr & "<span style='font-size:12px; color:red;'> - Tilbud<span>"
	            end if

                restestimat = oRec("restestimat")
                stade_tim_proc = oRec("stade_tim_proc")
                aftid = oRec("aftid")

                if len(oRec("jo_bruttooms")) <> 0 then
		        jobbudget = oRec("jo_bruttooms")
		        else
		        jobbudget = 0
		        end if     

                useasfak = oRec("useasfak")
	    %>

		<div id="jinf_knap" style="position:absolute; left:779px; top:632px; width:210px; visibility:<%=jinf_knap_vzb%>; display:<%=jinf_knap_dsp%>; z-index:100000; border:1px #cccccc solid; padding:2px; background-color:#ffffff;">
	    <table cellpadding=2 cellspacing=0 border=0 width=100%>
	     <tr bgcolor="#5C75AA"><td style="padding:3px; height:20px;" class=alt><b>Links</b></td></tr>
	     <tr bgcolor="#FFFFFF">		        
				<td>
                <a href="materialer_indtast.asp?id=<%=jobid%>&fromsdsk=<%=fromsdsk%>&aftid=<%=aftid%>" target="_blank" class=rmenu><%=tsa_txt_021 %> >></a>
				</td>
	    </tr> 
        <tr>
            <td><%
			    if cint(dsksOnOff) = 1 then
			    %>	
				<a href="javascript:popUp('stopur_2008.asp','1000','650','20','20')" class=rmenu><%=tsa_txt_080 %> >></a>
				<%end if%></td>
        </tr>
         <tr>
            <td>
			 <%if level <= 3 OR level = 6 then
				%>
				<a href="week_real_norm_2010.asp" class=rmenu><%=tsa_txt_356 %> >></a>
			    <%end if
				%></td>
        </tr>
		</table>
        </div>

        <%
				
				
				
				lukafmjob = oRec("lukafmjob")
				kid = oRec("kid")
				
				
				
				end if
				oRec.close
				%>
			
                <%response.flush %>
                
                
                <%
            
                    '**** Valgt medarbejer *****'
                    strSQLmn = "SELECT mnavn, mnr, init, madr, mpostnr, mcity, mland FROM medarbejdere WHERE mid = "& usemrn 
                    oRec4.open strSQLmn, oConn, 3
                    if not oRec4.EOF then
                    
                    mMnavn = oRec4("mnavn")
                    mMnr = oRec4("mnr")
                    mInit = oRec4("init")
                    mAdr = oRec4("madr")
                    mPostnr = oRec4("mpostnr")
                    mCity = oRec4("mcity")
                    mLand = oRec4("mland")
                    
                    end if
                    oRec4.close
                 %>
                 
                 
                 
                        
                <!-- 
                '********************************************************
                Kontakt Kørsel info div 
                
                
                
                '********************************************************
                -->
	            <div id="korseldiv" name="korseldiv" style="position:absolute; left:100px; top:500px; background-color:#ffffff; width:600px; border:10px #CCCCCC solid; padding:10px; visibility:hidden; display:none; z-index:2000000; height:700px; overflow:auto;">
				<div style="width:100%; border:0px #8caae6 solid; padding:20px;">
				<a href="javascript:popUp('kontaktpers.asp?func=opr&id=0&kid=<%=kid%>&rdir=treg','550','500','20','20')" target="_self" class=remnu>+ Tilføj ny kontakt</a>
				<table border=0 cellspacing=0 cellpadding=0 width=100%>
				<form action="#" method="post">
				<tr><td align=right><a href="#" id="jq_hideKpersdiv" onClick="hideKpersdiv()" class=red>[x]</a></td></tr>
				
				<tr>
				    <td><h3><%=tsa_txt_360 %>:</h3>
                    <%=tsa_txt_361 %><br />
                    <%=tsa_txt_365 %><br /><br />
				   
				     <input id="ko0chk" value="1" type="checkbox" /> <select id="ko0sel" style="font-size:9px;">
                            <option value=1><%=tsa_txt_362 %></option>
                            <option value=0>-</option>
                            <option value=2><%=tsa_txt_363 %></option>
                            <option value=3><%=tsa_txt_364 %></option>
                            </select>
                        <b><%=tsa_txt_367%></b>
                        <div id="ko0">
                        <%if len(trim(mAdr)) <> 0 then %>
                        <%=mMnavn %><br />
                        <%=mAdr%><br />
                        <%=mPostnr &" "& mCity%><br />
                        <%=mLand%> 
                        <%end if %>
                        </div>
				   </td>
				</tr>
				
				<tr>
				    <td>
                        <br /><input id="ko1chk" type="checkbox" />
                           <select id="ko1sel" style="font-size:9px;">
                            <option value=2><%=tsa_txt_363 %></option>
                            <option value=1><%=tsa_txt_362 %></option>
                            <option value=3><%=tsa_txt_364 %></option>
                            <option value=0>-</option>
                            </select>
                        
                         <b><%=tsa_txt_331 %></b>
                        <%
                        strSQLklicens = "SELECT k.kid, k.kkundenavn, k.adresse, k.postnr, k.city, k.land FROM kunder k WHERE useasfak = 1"
                        oRec.open strSQLklicens, oConn, 3
                        if not oRec.EOF then
                        
                        koAdr = oRec("adresse")
                        koPostnr = oRec("postnr")
                        koCity = oRec("city")
                        koLand = oRec("land")
                        koKid = oRec("kid")
                        kokkundenavn = oRec("kkundenavn")
                        
                        end if
                        oRec.close
                        %>
                        <br /><img src="../ill/blank.gif" width=1 height=5 /><br>
                        <div id="ko1" style="padding:5px; border:1px #cccccc solid; top:5px; background-color:#ffffe1;">
                        <%
                        if len(trim(koAdr)) <> 0 then %>
                        <%=kokkundenavn %><br />
                        <%=koAdr %> <br />
                        <%=koPostnr&" "&koCity %> <br />
                        <%=koLand%> <br />
                        <%end if %>
                        </div>
                        <br />
                        
                        <input id="ko1kid" type="hidden" value="k_<%=koKid%>" />
                        
                     
                        
				   </td>
				</tr>
				
				<tr><td align=right><input id="indlaes_koadr_1" type="button" value=" Indlæs >> " onclick="koadr()" />
                    </tr>
				
				<tr><td valign=top>
                <br /><br />
				<h4>Du har kørt til:<br /><span style="font-size:10px;">Søg flere gange for at tilføje flere via punkter</span></h4>
                

				<%=tsa_txt_022 %> (<%=lcase(tsa_txt_243) %>), <%=tsa_txt_024 %>:<br />
				
				<b><%=tsa_txt_078 %>:</b> 
                    <input id="FM_sog_kpers_dist" name="FM_sog_kpers_dist" value="" type="text" style="width:200px;" /> 
                    &nbsp;<input id="FM_sog_kpers_but" type="button" value="<%=tsa_txt_078 %> >> " />
                    <input id="FM_sog_kpers_dist_kid" value="0" type="hidden" />
                    <br />
                    <input id="FM_sog_kpers_dist_all" value="1" type="checkbox" /> <%=tsa_txt_359 %>
                    <!--<br /><input id="FM_sog_kpers_dist_more" value="1" type="checkbox" /> <=tsa_txt_372 %>-->
				
				<br /><br />
				<div id="filialerogkontakter">
				<!--jquery TXT -->
				</div>
				<br />&nbsp
				</td></tr>
				<tr><td align=right>
				<%call erkmDialog() 
				%>
				    <input id="koKmDialog" value="<%=kmDialogOnOff%>" type="hidden" />
                    <input id="koFlt" value="0" type="hidden" />
                    <input id="koFltx" value="1" type="hidden" />
                    
                    <input id="Button1" type="button" value=" Indlæs >> " onclick="koadr();" /><br />&nbsp;
                    </td></tr>
				</form>
				</table>
				</div>
				</div>
				<!--KM dialog div -->
				<%response.flush %>
				
				
				
				
				
				<!-- Kontakt info div -->
	            <div id="kpersdiv" name="kpersdiv" style="position:absolute; left:0px; top:0px; width:400px; border:2px #6CAE1C solid; padding:10px; visibility:hidden; display:none; background-color:#ffffff; z-index:2000000; height:300px; overflow:auto;">
				<table border=0 cellspacing=0 cellpadding=0 width=100%>
		        <tr><td align=right><a href="#" onClick="hideKpersdiv()" class=red>[x]</a></td></tr>
				<tr><td valign=top>
		        <b><%=tsa_txt_022 %>:</b><br />
				
				<%
				
				if kid <> 0 then
				kid = kid 
				else
				kid = 0
				end if
				
				'strSQL2 = "SELECT kp.navn, kp.email, kp.dirtlf, kp.mobiltlf, "_
				'&" kp.adresse As kpadr, kp.postnr As kpzip, kp.town As kptown, kp.land As kpland, kp.afdeling AS kpafd, "_
				'&" k.kkundenavn, k.kkundenr, k.adresse, k.postnr, k.city, k.telefon, k.land FROM kontaktpers kp, kunder k WHERE kp.kundeid = "& kid &" AND k.kid =" & kid
			    
			    strSQL2 = "SELECT kp.id, kp.navn, kp.email, kp.dirtlf, kp.mobiltlf, "_
				&" kp.adresse As kpadr, kp.postnr As kpzip, kp.town As kptown, kp.land As kpland, kp.afdeling AS kpafd, "_
				&" k.kid, k.kkundenavn, k.kkundenr, k.adresse, k.postnr, k.city, k.telefon, k.land FROM kunder k "_ 
				&" LEFT JOIN kontaktpers kp ON (kp.kundeid = k.kid)"_
				&" WHERE k.kid =" & kid
			    
			    oRec2.open strSQL2, oConn, 3
                x = 2
                while not oRec2.EOF
                
                        if x = 2 then%>
                         <br /><b><%=oRec2("kkundenavn") %></b> (<%=oRec2("kkundenr") %>)
                         <div id="Div4" style="padding:5px; border:1px #cccccc solid; top:5px; background-color:#DCF5BD;">
                        <%if len(trim(oRec2("adresse"))) <> 0 then %>
                        <%=oRec2("adresse") %> <br />
                        <%=oRec2("postnr") %>, <%=oRec2("city") %> <br />
                        <%=oRec2("land") %> <br />
                        <%end if %>
                       
                        
                            <%if len(trim(oRec2("telefon"))) <> 0 then %>
                            <%=tsa_txt_023 %>: <%=oRec2("telefon") %> <br />
                           
                            <%end if %>
                         
                          </div>
                            
                         <%x = x + 1 %>    
                         
                         <br /><br /><b><%=tsa_txt_024 %>:</b><br />
                         
                        <%end if %>
                        
                <%if len(trim(oRec2("id"))) <> 0 then %>
                <br /><b><%=oRec2("navn")%></b>
                        <div id="Div5" style="padding:5px; border:1px #cccccc solid; top:5px; background-color:#EFF3FF;">
                        <%if len(trim(oRec2("kpadr"))) <> 0 then %>
                        <%=oRec2("kpafd") %> <br />
                        <%=oRec2("kpadr") %> <br />
                        <%=oRec2("kpzip") %>, <%=oRec2("kptown") %> <br />
                        <%=oRec2("kpland") %> <br />
                        <%end if %>
                        
                
                <%if len(trim(oRec2("email"))) <> 0 then %>
                <%=tsa_txt_025 %>: <a href="mailto:<%=oRec2("email")%>" class=vmenu><%=oRec2("email")%></a><br>
                <%end if %>
                
                <%if len(trim(oRec2("dirtlf"))) <> 0 then %>
                <%=tsa_txt_026 %>: <%=oRec2("dirtlf")%><br>
                <%end if %>
                
                <%if len(trim(oRec2("mobiltlf"))) <> 0 then %>
                <%=tsa_txt_027 %>: <%=oRec2("mobiltlf")%><br><br />
                <%end if %>
                
                </div>
                
                <%end if %>
                
                <%
                x = x + 1
                oRec2.movenext
                wend
                oRec2.close
				%>
				
				<br />&nbsp
				</td></tr>
				
				</table>
				
				</div>
				
    
  
	
	
	<%
	
	end if '******* Denne ??
	'Response.flush
	
	
	call browsertype()
	if browstype_client = "mz" then
	topAdd_1 = 16
	else
	topAdd_1 = 0
	end if
	
	

   
	
     '*** SMILEY faneblade ****'
    
   
    '*** Smiley *** 
    'sdTop = (333+topAdd)
    'if browstype_client = "ch" then
    'smdivTop = "122" 
    'else
    '    if browstype_client = "mz" then
    '    smdivTop = "138"
    '    else
    '    smdivTop = "141"
    '    end if
    'end if

     %>           
             
             <!--   
  <div id="s" style="position:absolute; left:25px; top:<%=smdivTop%>px; width:200px; visibility:<%=sVzb%>; display:<%=sDsp%>; z-index:2;">
   <table cellpadding=0 cellspacing=0 border=0>
        <tr>
            <td align=center id="s1_td" style="white-space:nowrap; width:100px; padding:1px; background-color:#EFF3FF; border-right:1px #d6dff5 solid; border-bottom:0px;">
			<a href="#" id="s1_k" class=rmenu>+ <%=tsa_txt_345 %></a></td>
            
           
			
        </tr>
    </table>
    </div>
    -->
	    
		<%
	
	
	    if smilaktiv <> 0 AND showakt <> 0 AND cint(projektgrpOK) <> 0 then
	
	   
	    call medrabSmilord(usemrn)
	    
	    '**** Smiley vises om fredagen, første gang me logger på. ***'
	    if datepart("w", now, 2,2) = 5 AND datepart("h", now, 2,2) >= 13 AND session("smvist") <> "j" AND showakt = "1" then
    	smVzb = "visible"
    	smDsp = ""
    	session("smvist") = "j"
    	showweekmsg = "j"
    	else
    	smVzb = "hidden"
    	smDsp = "none"
    	showweekmsg = "n"
    	end if
    	
    	'showweekmsg = "j"
    	
	        if cint(smilaktiv) = 1 then%>
	        <div id="s1" style="position:absolute; left:50px; top:395px; width:550px; visibility:<%=smVzb%>; display:<%=smDsp%>; z-index:20000000; background-color:#FFFFFF; padding:10px; border:10px #CCCCCC solid;">
	       <%
	        call showsmiley()
	        
	        call afslutuge()
	        
	        %><br />&nbsp;
	     
	        <%
	        useYear = year(tjekdag(4))
	        call smileystatus(usemrn, 1)
	        %>
	        <br />&nbsp;
	        </div>
        	
	        <%end if
    	
	    end if

    
    
    %>
    <form action="timereg_akt_2006.asp?FM_use_me=<%=usemrn%>&func=db&fromsdsk=<%=fromsdsk%>" method="post" name="timeregfm" id="timeregfm">
    
    <%	
	


	
	if showakt <> 0 AND cint(projektgrpOK) <> 0 then
	
         
    
    tTop = 0 '53 + topAdd_1 '(225+topAdd)
	tLeft = 20
	tWdth = 745
	tId = "timeregdiv"
    %>
    <br /><br /><br /><br /><br />

    <div style="position:relative; left:20px; padding:10px 0px 0px 0px;">
     <h4>Personlig aktiv-jobliste: <span style="font-size:9px;"><%=meTxt %><br />


   <%
   sortByUline1 = "Prioitet"
   sortByUline2 = "Jobnavn"
   sortByUline3 = "Slutdato"
   sortByUline4 = "Kunde"

   select case sortBy
   case 1
   sortByUline1 = "<u>Prioitet</u>"
   case 2
   sortByUline2 = "<u>Jobnavn</u>"
   case 3
   sortByUline3 = "<u>Slutdato</u>"
   case 4
   sortByUline4 = "<u>Kunde</u>"
   case else
   sortByUline1 = "<u>Prioitet</u>"
   end select %>

   Sorter efter: <a href="timereg_akt_2006.asp?showakt=1&FM_mnr=<%=usemrn%>&sortby=2&tildeliheledageSet=<%=tildeliheledageVal%>" class=rmenu><%=sortByUline2 %></a> / <a href="timereg_akt_2006.asp?showakt=1&FM_mnr=<%=usemrn%>&sortby=1&tildeliheledageSet=<%=tildeliheledageVal%>" class=rmenu><%=sortByUline1 %></a> / 
   <a href="timereg_akt_2006.asp?showakt=1&FM_mnr=<%=usemrn%>&sortby=3&tildeliheledageSet=<%=tildeliheledageVal%>" class=rmenu><%=sortByUline3 %></a> / <a href="timereg_akt_2006.asp?showakt=1&FM_mnr=<%=usemrn%>&sortby=4&tildeliheledageSet=<%=tildeliheledageVal%>" class=rmenu><%=sortByUline4 %></a></span></h4>
   </div>


    <%
	
	call tableDivWid(tTop,tLeft,tWdth,tId, dTimVzb, dTimDsp)


    
     call meStamdata(usemrn)

	
	%>

	
   	
	
	<table cellspacing="0" cellpadding="2" border="0" width="100%" bgcolor="#8caae6"><!-- id="inputTable" onkeydown="doKeyDown()" -->
     <%if smilaktiv <> 0 AND showakt <> 0 AND cint(projektgrpOK) <> 0 then %>
     <tr bgcolor="#ffffff">
	    <td colspan=9 valign=top>
        	<a href="#" id="s1_k" class=rmenu style="background-color:#FFDFDF;">+ <%=tsa_txt_345 %></a>
        </td>
    </tr>
    <%end if %>
	
	<tr bgcolor="#ffffff">
	    <td colspan=7 valign=top>
	    <%
	    oskrift = tsa_txt_031 &" "& datepart("ww", tjekdag(1), 2 ,2) & " "& datepart("yyyy", tjekdag(1), 2 ,3) 
	    %>
	<b><%=oskrift %></b> <a href="#" name="anc_s0" id="pagesetadv_but" class="rmenu">(+ <%=tsa_txt_302%> for <%=lcase(tsa_txt_305)%>, <%=tsa_txt_281%> etc.)</a>
	<br />
	<%if cint(intEasyreg) = 1 then%>
	<span><h3>Easyreg. listen 
	    <%if level = 1 then %>
	    <a href="javascript:popUp('easyreg_2010.asp','800','720','20','20')" class=rmenu>(+ <%=lcase(tsa_txt_251)%>)</a>
	    <%end if %>
	</h3>  </span>
	<%else %>

	<!--
    <h3><=j_jobnavn_nr %>  <a href="#" id="showjobbesk" class=rmenu>(+ <=tsa_txt_029 %>: <=len(trim(jobbeskrivelse)) &" "& tsa_txt_371 &"." %>)</a></h3>
    -->
	
	<%end if %>


   
	</td>
	<td colspan=2 valign=top align=right>
	                        <!-- skift uge pile -->
	                        <table cellpadding=0 cellspacing=0 border=0 width=80>
	                        <tr>
	                        <td valign=top align=right><a href="timereg_akt_2006.asp?showakt=1&strdag=<%=day(useDatePrevWeek)%>&strmrd=<%=month(useDatePrevWeek)%>&straar=<%=year(useDatePrevWeek)%>&tildeliheledageSet=<%=tildeliheledageVal%>"><img src="../ill/nav_left_blue.png" border="0" /></a></td> <!-- jobid=<=jobid%>&usemrn=<=usemrn%>& &fromsdsk=<=fromsdsk%> -->
                           <td style="padding-left:20px;" valign=top align=right><a href="timereg_akt_2006.asp?showakt=1&strdag=<%=day(useDateNextWeek)%>&strmrd=<%=month(useDateNextWeek)%>&straar=<%=year(useDateNextWeek)%>&tildeliheledageSet=<%=tildeliheledageVal%>"><img src="../ill/nav_right_blue.png" border="0" /></a></td>
	                        </tr>
	                        </table>
	</td>
	</tr>
	
	
	
	
    
    <!--<input id="showjobinfo" name="showjobinfo" type="hidden" value="<=showjobinfo %>" />-->
	
    
		
		<tr bgcolor="#FFFFFF">
        <td colspan=7 valign=top>
        <div id="indlaes_settings" style="border:0px #999999 solid; visibility:; display:;">
		 
		 <!-- luk job ved indtatning --->
        <%
	            lukafm = 0
	            strSQL = "SELECT lukafm FROM licens WHERE id = 1"
	            oRec.open strSQL, oConn, 3
	            while not oRec.EOF
	            lukafm = oRec("lukafm")
	            oRec.movenext
	            wend
	
	            oRec.close

                '*** Altid = 1 20110927 ***'
                lukafm = 1
	
	     %>
                    <!--<div style="position:relative; border:4px #cccccc solid; padding:5px; width:550px;">-->
                    <table cellpadding=1 cellspacing=0 border=0 width=100%><%
	
	                if cint(lukafm) = 100 then%>
	
		
		                 <%select case lukafmjob
		                 case 1
		                 %>
		                    <tr><td colspan=2>
		                    <input type="checkbox" name="xFM_lukjob" id="xFM_lukjob" value="1"> <span style="color:#000000;">
		                    <%if lto = "bowe" then%>
		                    <%=tsa_txt_040 %>
		                    <%else%>
		                    <%=tsa_txt_041 %>
		                    <%end if
		    
		                    %></span></td></tr><%
		                 case 2
		                 %>
		    
		
		                 <tr><td>
		   
		                    <input type="checkbox" name="xFM_lukjob" id="xFM_lukjob" value="1"> </td><td> <span style="color:#000000;">
		                    <%if lto = "bowe" then%>
		                    <%=tsa_txt_043 %>
		                    <%else%>
		                    <%=tsa_txt_044 %>
		                    <%end if
		    
		                    %> </span>
		                    <input type="hidden" name="xFM_kopierjob" id="xFM_kopierjob" value="1">
		    
		                  </td></tr>
		                <%
		                 end select
        
        
                        end if

                        'if cint(useasfak) = 1 then '*** Kun interne (licens ejer)
                        %>
		                
		               
                        

                         <input id="tildeliheledage" name="tildeliheledage" type="hidden" value="<%=tildeliheledageVal %>"/>
		
		                <%'*** Min indtast ETS **'
		
		                if lto = "ets-track" OR lto = "xintranet - local" then %>
		                <tr><td colspan=2><input id="minindtast_on" name="minindtast_on" type="checkbox" value="1" CHECKED /> <span style="color:#000000;">Gennemtving min. indtastning 7,4 timer / 8,0 timer</span>
		                 </td></tr>
		                <%end if %>

                        <%'end if 'licensejer%>
		
		                <%
		                '*** Multi tildel ***'
		                'if level = 1 then
		                %>
		                <tr><td colspan=2>
                        <input id="tildelalle" name="tildelalle" type="checkbox" value="1" onclick="showmultiDiv()" /> <span style="color:#000000;"><%=tsa_txt_268 %> (<%=tsa_txt_357 %>)</span> <br />
                        
                        <input type="CHECKBOX" id="abnlukallejob" name="abnlukallejob" value="1" /> Luk (kollaps) alle aktiviteter på joblisten (indlæser timer)<br />
		
		                                <div id="multivmed" style="position:relative; left:0px; top:0px; padding:10px; visibility:hidden; display:none;">
		                                <%=tsa_txt_267 %>:<br /> <select name="tildelselmedarb" id="tildelselmedarb" size=10 multiple style="width:350px;">
				                                <%
					                                strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe FROM medarbejdere WHERE mansat <> 2 "& medarbgrpIdSQLkri &" ORDER BY Mnavn"
					                                oRec.open strSQL, oConn, 3
					
					                                while not oRec.EOF 
					
					                                if cint(oRec("Mid")) = cint(usemrn) then
					                                rchk = "SELECTED"
					                                else
					                                rchk = ""
					                                end if%>
					                                <option value="<%=oRec("Mid")%>" <%=rchk%>><%=oRec("mnavn")%></option>
					                                <%
					
					
					                                oRec.movenext
					                                wend
					                                oRec.close
				                                %></select>
				                                </div>
		
		
		
		                                <%
		                                'end if
		                                %>
		
		                                </td></tr>
                                        </table>

                                        <!--</div>-->
		                                <br />&nbsp;

        
                            
		                    </div> <!-- indlaes_settings -->
		
		              
		                </td>
		                <td colspan=2 align=right valign=bottom style="padding-bottom:5px;"><input type="submit" id="sm1" value="<%=tsa_txt_085 %> >> "></td>
	                </tr>
	
	                 <%
                     if cint(intEasyreg) = 1 then
     
                     %>
                     <tr bgcolor="#FFFFFF"><td colspan=9 style="height:5px;"><img src="ill/blank.gif" border=0 /></td></tr>
                     <tr bgcolor="#DCF5BD">
                        <td colspan=2 style="border-right:1px #CCCCCC solid; border-left:1px #CCCCCC solid; border-top:1px #CCCCCC solid;">Fordel timer på Easyreg. aktiviteter</td>
        
                        <%for e = 1 to 7 %>
                        <td style="border-right:1px #CCCCCC solid; border-top:1px #CCCCCC solid;">
                            <input id="ea_<%=e %>" type="text" style="width:46px; font-size:9px;" /><a href="#" id="ea_k_<%=e%>" class="ea_kom"><font color="#999999"> + </font></a></td>
                        <%next %>
     
     
                     </tr>
                     <tr bgcolor="#DCF5BD">
                     <td colspan=2 style="border-right:1px #CCCCCC solid; border-left:1px #CCCCCC solid; border-top:1px #CCCCCC solid; width:325px;">(0 = nulstil)</td>
                     <%for e = 1 to 7 %>
                            <td style="border-right:1px #CCCCCC solid; border-top:1px #CCCCCC solid; width:55px;">
                                <input id="ea_kn_<%=e %>" type="button" value="Calc. =" style="font-family:arial; font-size:9px;" /></td>
                        <%next %>
     

                     </tr>
                     <%
                     end if
                     %>
     

      

      </table>
    </div><!--tableDiv-->
     

     <div id="timereg_job" style="position:relative; border:0px; top:0px; left:20px; width:740px; visibility:visible; display:; z-index:1000">
    
     <%
    'tTop = 153
	'tLeft = 20
	'tWdth = 0 '760
	'tId = "timereg_job"
	
	'call tableDivWid(tTop,tLeft,tWdth,tId, dTimVzb, dTimDsp)
    
   

     %>
   
   
    
    <table id="incidentlist" cellspacing=0 cellpadding=0 border=0 width=100%>
  
    <%
    'Response.Write "jobid:"& request("jobid") & "<br>"
    'Response.Write "seljobid" & request("seljobid")
    
    %>
    
	<input type="hidden" name="tildeliheledageSet" id="Hidden2" value="<%=tildeliheledageVal%>">
    <input type="hidden" name="hideallbut_first" id="hideallbut_first" value="<%=hideallbut_first%>">
    <input type="hidden" name="jobid" id="xjobid" value="<%=seljobidUse%>">
	<input type="Hidden" name="FM_vistimereltid" value="<%=visTimerElTid%>">
	<input type="Hidden" name="FM_start_dag" value="<%=strdag%>">
	<input type="Hidden" name="FM_start_mrd" value="<%=strmrd%>">
	<input type="Hidden" name="FM_start_aar" value="<%=straar%>">
	<!--<input type="Hidden" name="searchstring" value="<=searchstring%>">-->
	<input type="hidden" name="FM_mnr" id="treg_usemrn" value="<%=usemrn%>">

	<input type="hidden" name="year" value="<%=strRegAar%>">
	<input type="hidden" id="datoMan" name="datoMan" value="<%=tjekdag(1)%>">
	<input type="hidden" name="datoTir" value="<%=tjekdag(2)%>">
	<input type="hidden" name="datoOns" value="<%=tjekdag(3)%>">
	<input type="hidden" name="datoTor" value="<%=tjekdag(4)%>">
	<input type="hidden" name="datoFre" value="<%=tjekdag(5)%>">
	<input type="hidden" name="datoLor" value="<%=tjekdag(6)%>">
	<input type="hidden" id="datoSon" name="datoSon" value="<%=tjekdag(7)%>">
	<input type="hidden" id="lastopendiv" name="lastopendiv" value="jobinfo">

    <input type="hidden" id="FM_session_user" name="FM_session_user" value="<%=session("user")%>">
    <input type="hidden" id="FM_now" name="FM_now" value="<%=formatdatetime(now, 2)%>">

 

	
    <%

	strAktidsThisJob = "0"
    strAktnavneThisJob = "x"
    fsVal = "nonexx99w"
    tregDatoer = "1-1-2001"
    iRowLoopsThisJob = 0
    
     strEaAktidsThisJob = 0
    iRowLoopsEAThisJob = 0
   

	totTimMan = 0
	totTimTir = 0
	totTimOns = 0
	totTimTor = 0
	totTimFre = 0
	totTimLor = 0
	totTimSon = 0
	
	'*** Brugergrupper er fundet i linie 2697 ***
	'call hentbgrppamedarb(usemrn)
	
	iRowLoop = 1
	m = 58 '16
    if cint(intEasyreg) = 1 then
    x = 3500 ' Bør kun være 750, men der kan forekomme mere end 100 Easyreg. aktiviteter indtil loft effketureres
    else
	x = 550 '350 '6500 '2500 (maks 100 aktiviteter * 7 dages timereg.)
    end if
			
	dim aktdata
	redim aktdata(x, m) 
	
					
					
					
    '*** Aktiviteter og Timer SQL MAIN **** 
	strSQL = "SELECT a.navn, a.id AS aid, a.fakturerbar, "_
	&" j.jobnr AS jnr, j.id AS jid, j.jobnavn AS jnavn, j.risiko, job_internbesk, "_
    &" j.serviceaft, j.beskrivelse AS jobbesk, j.jobans1, jobans2, jobstartdato, jobslutdato, j.fastpris, jo_bruttooms, j.rekvnr, j.kommentar, "_
	&" a.fase, j.jobstatus, "_
    &" tu.forvalgt_sortorder, tu.id AS tuid, tu.forvalgt_af, tu.forvalgt_dt, "_
    &" SUM(tman.timer) As tmantimer, SUM(ttir.timer) As ttirtimer, SUM(tons.timer) As tonstimer, SUM(ttor.timer) As ttortimer, SUM(tfre.timer) As tfretimer, SUM(tlor.timer) As tlortimer, SUM(tson.timer) As tsontimer"

    ', t.timer, t.tdato,
    't.sttid, t.sltid, 
    '&" t.godkendtstatus, "_
	'&" t.bopal, t.destination, 
    't.timerkom, t.offentlig,
    '&" a.tidslaas, a.tidslaas_st, a.tidslaas_sl, a.tidslaas_man, a.tidslaas_tir, a.aktbudget, a.budgettimer AS aktbudgettimer, a.bgr, a.antalstk, "_
	'&" a.tidslaas_ons, a.tidslaas_tor, a.tidslaas_fre, a.tidslaas_lor, a.tidslaas_son, 
    'a.incidentid, a.aktstartdato, a.aktslutdato, a.beskrivelse, 
	
	'if cint(intEasyreg) = 1 then
	strSQL = strSQL &", kkundenavn, kkundenr, kid, kundeans1 "
	'end if
	
	strSQL = strSQL &" FROM job j, aktiviteter a "
	
	if cint(intEasyreg) = 1 then
	strSQL = strSQL &", timereg_usejob AS tu "
	else
    strSQL = strSQL &" LEFT JOIN timereg_usejob AS tu ON (tu.jobid = j.id AND tu.medarb = "& usemrn &") "
    end if
	
    strSQL = strSQL &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "

	strSQL = strSQL &" LEFT JOIN timer tman ON"_
	&" (tman.taktivitetid = a.id "_ 
    &" AND tman.tmnr = "& usemrn &" AND tman.tdato = '"& varTjDatoUS_man &"' AND ("& replace(aty_sql_realhours, "tfaktim", "tman.tfaktim") &"))"

    strSQL = strSQL &" LEFT JOIN timer ttir ON"_
	&" (ttir.taktivitetid = a.id "_ 
    &" AND ttir.tmnr = "& usemrn &" AND ttir.tdato = '"& varTjDatoUS_tir &"' AND ("& replace(aty_sql_realhours, "tfaktim", "ttir.tfaktim") &"))"

    strSQL = strSQL &" LEFT JOIN timer tons ON"_
	&" (tons.taktivitetid = a.id "_ 
    &" AND tons.tmnr = "& usemrn &" AND tons.tdato = '"& varTjDatoUS_ons &"' AND ("& replace(aty_sql_realhours, "tfaktim", "tons.tfaktim") &"))"

    strSQL = strSQL &" LEFT JOIN timer ttor ON"_
	&" (ttor.taktivitetid = a.id "_ 
    &" AND ttor.tmnr = "& usemrn &" AND ttor.tdato = '"& varTjDatoUS_tor &"' AND ("& replace(aty_sql_realhours, "tfaktim", "ttor.tfaktim") &"))"

    strSQL = strSQL &" LEFT JOIN timer tfre ON"_
	&" (tfre.taktivitetid = a.id "_ 
    &" AND tfre.tmnr = "& usemrn &" AND tfre.tdato = '"& varTjDatoUS_fre &"' AND ("& replace(aty_sql_realhours, "tfaktim", "tfre.tfaktim") &"))"

    strSQL = strSQL &" LEFT JOIN timer tlor ON"_
	&" (tlor.taktivitetid = a.id "_ 
    &" AND tlor.tmnr = "& usemrn &" AND tlor.tdato = '"& varTjDatoUS_lor &"' AND ("& replace(aty_sql_realhours, "tfaktim", "tlor.tfaktim") &"))"

    strSQL = strSQL &" LEFT JOIN timer tson ON"_
	&" (tson.taktivitetid = a.id "_ 
    &" AND tson.tmnr = "& usemrn &" AND tson.tdato = '"& varTjDatoUS_son &"' AND ("& replace(aty_sql_realhours, "tfaktim", "tson.tfaktim") &"))"


	'&" OR t.tdato = '"& varTjDatoUS_man &"'"_
	'&" OR t.tdato = '"& varTjDatoUS_tir &"'"_
	'&" OR t.tdato = '"& varTjDatoUS_ons &"'"_
	'&" OR t.tdato = '"& varTjDatoUS_tor &"'"_
	'&" OR t.tdato = '"& varTjDatoUS_fre &"'"_
	'&" OR t.tdato = '"& varTjDatoUS_lor &"'))"
	
    'aty_sql_realhours
   
	    '*** Vis som easyreg. 
	    if cint(intEasyreg) = 1 then
            
            if cint(sogKontakter) <> 0 then
            strSQL = strSQL &" WHERE (k.kid = " & sogKontakter
            else
            strSQL = strSQL &" WHERE (k.kid <> 0 "
            end if

            strSQL = strSQL &" AND j.id <> 0 AND a.easyreg = 1 AND j.jobstatus = 1 "_ 
            &" AND tu.medarb = "& usemrn &" AND tu.jobid = j.id AND tu.easyreg <> 0) AND "& strSQLkri3

        
        else
            
            '** ign projektgrupper på aktiviteter. F.eks for admin bruger til at få vist ferie u. løn aktiviteter for en anden medarb. der ikke selv har rettigheder ****'
            if level = 1 AND ignorerProgrp = 1 then
            strSQLkri3 = " a.job = j.id AND aktstatus = 1 "
            else
            strSQLkri3 = strSQLkri3
            end if

            strSQL = strSQL &" WHERE ("& seljobidSQL &") AND (j.jobstartdato <= '"& varTjDatoUS_son &"' AND (j.jobstatus = 1 OR j.jobstatus = 3)) AND "& strSQLkri3
        end if
	
	
    
	
	'Response.Write  "strSQLkri3" &  strSQLkri3  
	
	if cint(ignJobogAktper) = 1 then
	useDateSQL = year(useDate) &"/"& month(useDate) &"/"& day(useDate)
	strSQL = StrSQL & " AND (a.aktstartdato <= '"& varTjDatoUS_son &"' AND a.aktslutdato >= '"& useDateSQL &"')"
	end if
	
    '*** Vis ikke disse typer på treg. siden **'
    strSQL = StrSQL & " AND ("& aty_sql_hide_on_treg &")"
    
	'if lto = "Q2con" then
	'strSQL = StrSQL & " ORDER BY j.jobnavn, j.id, LTRIM(a.navn)"
	'else
	    if cint(intEasyreg) = 1 then
	    strSQL = StrSQL & " GROUP BY j.id, a.id ORDER BY kkundenavn, j.jobnavn, j.id, a.fase, a.sortorder, LTRIM(a.navn) LIMIT 150" 'kkundenavn, 
	    else
	    strSQL = StrSQL & " GROUP BY j.id, a.id ORDER BY "& sortBySQL &" j.id, a.fase, a.sortorder, LTRIM(a.navn) LIMIT 550"
        'j.risiko, j.jobnavn, j.id, a.fase, a.sortorder, LTRIM(a.navn) LIMIT 550"
	    end if
	'end if
	
	
    'Response.write strSQL '& "<br><br>" & strSQLkri3

	Response.flush
	
    lastFirstLetter = ""
	tr_vzb = ""
    tr_dsp = ""
	lastFase = ""
    lastJobId = 0
    'tjkLastJidFak = 0
	
    antalJobLinier = 0
    LastJobLoop = 0
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
            

    if cint(intEasyreg) = 1 then
    strEaAktidsThisJob = strEaAktidsThisJob & "," & oRec("aid")
    iRowLoopsEAThisJob = iRowLoopsEAThisJob & "," & iRowLoop
    antalaktlinier = antalaktlinier + 1
    loops = antalaktlinier
    else              
	
	'aktdata(iRowLoop, 0) = oRec("tdato")
	'aktdata(iRowLoop, 1) = oRec("timer")
	aktdata(iRowLoop, 2) = oRec("navn")
	aktdata(iRowLoop, 3) = oRec("aid")
    
  

	aktdata(iRowLoop, 4) = oRec("jid")
	aktdata(iRowLoop, 5) = oRec("jnavn")
	'aktdata(iRowLoop, 6) = oRec("timerkom")
	'aktdata(iRowLoop, 7) = oRec("offentlig")
	'aktdata(iRowLoop, 11) = oRec("fakturerbar")
	'aktdata(iRowLoop, 12) = oRec("incidentid")
	'aktdata(iRowLoop, 13) = oRec("aktstartdato")
	'aktdata(iRowLoop, 14) = oRec("aktslutdato")
	'aktdata(iRowLoop, 15) = oRec("beskrivelse")
	'aktdata(iRowLoop, 16) = oRec("godkendtstatus")

   
	if LastJobLoop <> oRec("jid") then
	aktdata(iRowLoop, 11) = oRec("tmantimer")
	aktdata(iRowLoop, 12) = oRec("ttirtimer")
	aktdata(iRowLoop, 13) = oRec("tonstimer")
	aktdata(iRowLoop, 14) = oRec("ttortimer")
	aktdata(iRowLoop, 15) = oRec("tfretimer")
    aktdata(iRowLoop, 16) = oRec("tlortimer")
    aktdata(iRowLoop, 17) = oRec("tsontimer")
    jobRowLoop = iRowLoop
    else
    aktdata(jobRowLoop, 11) = aktdata(jobRowLoop, 11) + oRec("tmantimer")
	aktdata(jobRowLoop, 12) = aktdata(jobRowLoop, 12) + oRec("ttirtimer")
	aktdata(jobRowLoop, 13) = aktdata(jobRowLoop, 13) + oRec("tonstimer")
	aktdata(jobRowLoop, 14) = aktdata(jobRowLoop, 14) + oRec("ttortimer")
	aktdata(jobRowLoop, 15) = aktdata(jobRowLoop, 15) + oRec("tfretimer")
    aktdata(jobRowLoop, 16) = aktdata(jobRowLoop, 16) + oRec("tlortimer")
    aktdata(jobRowLoop, 17) = aktdata(jobRowLoop, 17) + oRec("tsontimer")

    end if
    
  
    LastJobLoop = oRec("jid")

	
	'aktdata(iRowLoop, 17) = oRec("tidslaas")
	'aktdata(iRowLoop, 18) = oRec("tidslaas_st")
	'aktdata(iRowLoop, 19) = oRec("tidslaas_sl")
	'aktdata(iRowLoop, 20) = oRec("tidslaas_man")
	'aktdata(iRowLoop, 21) = oRec("tidslaas_tir")
	'aktdata(iRowLoop, 22) = oRec("tidslaas_ons")
	'aktdata(iRowLoop, 23) = oRec("tidslaas_tor")
	'aktdata(iRowLoop, 24) = oRec("tidslaas_fre")
	'aktdata(iRowLoop, 25) = oRec("tidslaas_lor")
	'aktdata(iRowLoop, 26) = oRec("tidslaas_son")
	'aktdata(iRowLoop, 27) = oRec("bopal")
	'aktdata(iRowLoop, 28) = oRec("destination")

    if isNull(oRec("fase")) <> true then
	aktdata(iRowLoop, 29) = oRec("fase")
    else
	aktdata(iRowLoop, 29) = ""
    end if

    'aktdata(iRowLoop, 30) = oRec("aktbudget")
	aktdata(iRowLoop, 31) = oRec("jnr")

    'aktdata(iRowLoop, 33) = oRec("aktbudgettimer")
	'aktdata(iRowLoop, 34) = oRec("bgr")
    'aktdata(iRowLoop, 35) = oRec("antalstk")

    aktdata(iRowLoop, 40) = oRec("job_internbesk")
    aktdata(iRowLoop, 41) = oRec("jobbesk") 


   
	aktdata(iRowLoop, 32) = oRec("kkundenavn")
    aktdata(iRowLoop, 56) = "("& oRec("kkundenr") &")"

	aktdata(iRowLoop, 43) = oRec("kundeans1")
    aktdata(iRowLoop, 44) = oRec("jobans1")
    aktdata(iRowLoop, 50) = oRec("jobans2")

    aktdata(iRowLoop, 42) = oRec("kid")


    aktdata(iRowLoop, 45) = oRec("jobstartdato")
    aktdata(iRowLoop, 46) = oRec("jobslutdato")

    aktdata(iRowLoop, 47) = oRec("fastpris")
    aktdata(iRowLoop, 48) = oRec("jo_bruttooms")

    aktdata(iRowLoop, 49) = oRec("rekvnr")

    aktdata(iRowLoop, 51) = oRec("serviceaft")
    
    aktdata(iRowLoop, 52) = oRec("forvalgt_sortorder")	       
    aktdata(iRowLoop, 53) = oRec("tuid")

    aktdata(iRowLoop, 54) = oRec("forvalgt_af")
    aktdata(iRowLoop, 55) = oRec("forvalgt_dt")
    

    aktdata(iRowLoop, 58) = oRec("kommentar")

    'select case aktdata(iRowLoop, 34) 
    'case 1 
    'aktdata(iRowLoop, 36) = "Fork.: "& formatnumber(aktdata(iRowLoop, 33), 2)  & " t."
    'case 2
    'aktdata(iRowLoop, 36) = "Ant.: "& formatnumber(aktdata(iRowLoop, 35), 2)  & " stk."
    'case else
    'aktdata(iRowLoop, 36) = ""
    'end select

    '** Timer real + kost (lsået fra) **'
    'aktdata(iRowLoop, 37) = 0
    'aktdata(iRowLoop, 39) = 0
    'strSQLttot = "SELECT SUM(timer) AS timertot FROM timer WHERE taktivitetid = "& aktdata(iRowLoop, 3) & " GROUP BY taktivitetid"
    'oRec2.open strSQLttot, oConn, 3
    'if not oRec2.EOF then  
    'aktdata(iRowLoop, 37) = oRec2("timertot")
        
        '''if lto = "intranet - local" OR lto = "epi" then
        '', SUM(timer*kostpris*(kurs/100)) AS internKost
        ''aktdata(iRowLoop, 39) = oRec2("internkost") 
        '''end if

    'end if
    'oRec2.close

    '** Timer real kun valgte medarb **'
    'aktdata(iRowLoop, 38) = 0
    'strSQLtmnr = "SELECT SUM(timer) AS timermnr FROM timer WHERE taktivitetid = "& aktdata(iRowLoop, 3) & " AND tmnr = "& usemrn &" GROUP BY taktivitetid"
    'oRec2.open strSQLtmnr, oConn, 3
    'if not oRec2.EOF then  
    'aktdata(iRowLoop, 38) = oRec2("timermnr")
    'end if
    'oRec2.close

	

	
	'if len(trim(oRec("sttid"))) <> 0 then
    '   if left(formatdatetime(oRec("sttid"), 3), 5) <> "00:00" then
	'	aktdata(iRowLoop, 8) = left(formatdatetime(oRec("sttid"), 3), 5)
	'	else
	'	aktdata(iRowLoop, 8) = ""
	'	end if
	'else
	'aktdata(iRowLoop, 8) = ""
	'end if
	
	'if len(trim(oRec("sltid"))) <> 0 then
	'	if left(formatdatetime(oRec("sltid"), 3), 5) <> "00:00" then
	'	aktdata(iRowLoop, 9) = left(formatdatetime(oRec("sltid"), 3), 5)
	'	else
	'	aktdata(iRowLoop, 9) = ""
	'	end if
	'else
	'aktdata(iRowLoop, 9) = ""
	'end if

     end if
	
	iRowLoop = iRowLoop + 1
	oRec.movenext
	wend
	oRec.close
	
	'** If iRowLoop > 100 show alert **'


    '**
	
	
    if cint(intEasyreg) <> 1 then

	if cint(iRowLoop) <> 1 then
	
	dim flt
	Redim flt(7)
	
    loops = 0
	antalaktlinier = 0
	lastAktid = 0
	for iRowLoop = 1 TO UBOUND(aktdata)
		
		    
            'if iRowLoop = 1 then
          
            'end if
            
            
            loops = loops + 1
		    

             '**** Rettigheder
			if level <= 2 OR level = 6 then
			editok = 1
			else
				if cint(session("mid")) = aktdata(iRowLoop, 44) OR cint(session("mid")) = aktdata(iRowLoop, 50) OR (cint(aktdata(iRowLoop, 44)) = 0 AND cint(aktdata(iRowLoop, 50)) = 0) then
				editok = 1
				end if
			end if 
                                                        

          


		   if lastAktid <> aktdata(iRowLoop, 3) AND len(trim(aktdata(iRowLoop, 3))) <> 0 then


           

		
			if iRowLoop <> 0 then
			
					
					'for x = 1 to 7
					
					'Response.write flt(x) 
					
					'next
                  
                    antalaktlinier = antalaktlinier + 1 %>
			<!--</tr>-->
			<%end if%>
			
			
			<%
                                '*** Job ***'
                                if lastJobid <> aktdata(iRowLoop, 4) AND cint(intEasyreg) <> 1 then

                     
                   
                                                    if antalJobLinier <> 0 then
                                                    %>

                                                    <!--
                                                    <tr>
                                                    <td bgcolor="#FFFFFF" colspan=9 align=right valign=bottom style="padding:5px 5px 5px 5px;"><input type="submit" id="Submit1" value="<%=tsa_txt_085 %> >> "></td>
                                                    </tr>
                                                    
                                                     <tr>
                                                          <td colspan=9 style="background-color:#D6dff5; background-image:url('../ill/stripe_10.png'); height:2px;">
                                                            <img src="../ill/blank.gif" width=1 height=1 /></td>  </tr>
                                                            
                                                    </table>-->
                                                    
                                                    <!-- akttable -->
                                                    </div><!-- aktdiv -->
                                                    <% 
                                                    call jq_format(faVal)
                                                    faVal = jq_formatTxt
                                                    %>

                                                     <input type="hidden"  id="job_aktids_<%=lastJobid %>" value="<%=strAktidsThisJob %>" />
                                                      <input type="hidden" name="fasenavn" id="job_fasenavne_<%=lastJobid %>" value="<%=fsVal %>" />
 
                                                    <input name="iRowLoop" id="job_iRowLoop_<%=lastJobid %>" type="hidden" value="<%=iRowLoopsThisJob%>" />

                                                    
                                                             <%=jLuft%>
                                                           
                                                            


                                                    </td></tr>

                                                    <%
                                                    tregDatoer = "1-1-2001" 
                                                    strAktidsThisJob = "0"
                                                    fsVal = "nonexx99w"
                                                    iRowLoopsThisJob = 0


                                                 
                                                    
                                                    else %>

                                                     <tr>
               
                                                       <td>
                                                           <input type="hidden" name="SortOrder" value="0" />
	                                                    <input type="hidden" name="rowId" value="0" />
        

                                                   
                                                    <div id="Div1" style="position:relative; width:745px; border:0px; visibility:visible; display:; z-index:1000">
                                                    <table width=100% cellspacing=0 cellpadding=0 border=0>
                                                    <tr><td><img src="../ill/blank.gif" width="1" height="3" border="0" /></td></tr>
                                                    </table>
                                                    </div>

                                                    </td></tr>

            
                                                    <%end if %>

            
              

                                 <tr>
                                 <td valign=top>
                                 
                                    <%if sortby = 2 then 
                                    
                                    if ucase(left(aktdata(iRowLoop, 5),1)) <> lastFirstLetter then
                                    Response.Write "&nbsp;&nbsp;<br><br><span style=""background-color:#FFFFFF; width:45px; padding:3px;"">"& ucase(left(aktdata(iRowLoop, 5),1)) & "</span>"
                                    end if

                                    lastFirstLetter = ucase(left(aktdata(iRowLoop, 5),1)) 
                                    end if %> 


                                     <%if sortby = 4 then 
                                    
                                    if ucase(left(aktdata(iRowLoop, 32),40)) <> lastFirstLetter then
                                    Response.Write "&nbsp;&nbsp;<br><br><span style=""background-color:#FFFFFF; width:225px; padding:3px;"">"& left(aktdata(iRowLoop, 32),40) & "</span>"
                                    end if

                                    lastFirstLetter = ucase(left(aktdata(iRowLoop, 32),40)) 
                                    end if %> 

                                       <input type="hidden" name="SortOrder" value="<%=aktdata(iRowLoop, 52)%>" />
	                                <input type="hidden" name="rowId" value="<%=aktdata(iRowLoop, 53)%>" />
                                     
                                        <input type="hidden" id="FM_jobid_<%=antalJobLinier %>" value="<%=aktdata(iRowLoop, 4) %>" /> 
           

          

                                            <div id="ad_timereg_<%=aktdata(iRowLoop, 4)%>" style="position:relative; width:745px; border:0px; visibility:visible; display:; z-index:1000">
                                            <table width=100% cellspacing=0 cellpadding=0 border=0 bgcolor="#ffffff">
             
          
            
                                                 <tr>
                                                    <td style="background-color:#D6Dff5; border-bottom:0px #cccccc solid; padding:0px 0px 0px 0px;" colspan=5 class=lille><img src="../ill/blank.gif" width=1 height=2 border=0 /></td>
                                                 </tr>
                
                                                <tr>
                                                    <td style="width:240px; background-color:#5C75AA; border-bottom:3px #8caae6 solid; border-top:1px #FFFFFF solid; padding:3px 3px 0px 4px;" valign=top> <a class="a_treg" id="a_timereg_<%=aktdata(iRowLoop, 4)%>" href="#">+ <%=left(aktdata(iRowLoop, 5), 30) %> (<%=aktdata(iRowLoop, 31) %>)</a>
                                                    <%if editok = 1 then %>
                                                    <a href="jobs.asp?menu=job&func=red&id=<%=aktdata(iRowLoop, 4) %>&int=1&rdir=treg" style="font-size:9px; color:#6CAE1C;"><%=left(tsa_txt_251,3) %>.</a>
                                                    <%end if %>
                                                    <br /><span style="color:#cccccc; font-size:9px;"><%=left(aktdata(iRowLoop, 32), 30) & " " & aktdata(iRowLoop, 56) %></span>
                                                    </td>

                                                    <td style="width:15px; background-color:#5C75AA; border:0px #cccccc solid; border-bottom:3px #8caae6 solid; border-top:1px #FFFFFF solid; border-right:0px #cccccc solid; padding:3px 3px 0px 4px;" valign=top><a class="a_det" id="a_det_<%=aktdata(iRowLoop, 4)%>" href="#"><img src="../ill/document.png" border="0" alt="Stamdata, job beskrivelse & Materiale registrering"/></a></td>
                    
                                                    <td valign="top" style="width:165px; background-color:#FFFFFF; border-bottom:3px #8caae6 solid; padding:2px 2px 0px 1px;">




                                                     <table cellpadding=1 cellspacing=1 border=0 height=100%>
                                                    <tr><td valign=top>

                                                    <%'if cint(usemrn) <> cint(aktdata(iRowLoop, 54)) then %>
                                                    &nbsp;<span style="font-size:10px; color:#999999;"> <%=antalJobLinier %> // <i>Af  
                                                    <%call meStamdata(aktdata(iRowLoop, 54)) %>
                                                    <a href="mailto:<%=meEmail %>?subject=Vedr. job <%=left(aktdata(iRowLoop, 5), 35) %> (<%=aktdata(iRowLoop, 31) %>)" class="rmenu"><%=meInit %></a> d. <%=aktdata(iRowLoop, 55)%> 
                                                    </i></span>
                                                    <%'end if %>
                                                    </td>
                                                    <td valign=top style="padding:3px 0px 0px 3px;"><img src="../ill/pile_drag.gif" alt="Træk og sorter job" border="0"  class="imgDrag" /></td>
                                                    </tr>
                                                    <tr><td valign=bottom colspan=2>

                                                    <table cellpadding=1 cellspacing=1 border=0>
                                                    <tr>
                                                    <%for d = 1 to 7 
                                                    if aktdata(iRowLoop, d+10) > 0 then
                                                    bgThisDay = "yellowgreen"
                                                    hgtThisDay = aktdata(iRowLoop, d+10)
                                                    else
                                                    bgThisDay = "#cccccc"
                                                    hgtThisDay = 0
                                                    end if

                                                    if aktdata(iRowLoop, d+10) >= 7 then
                                                    bgThisDay = "green"
                                                    hgtThisDay = aktdata(iRowLoop, d+10)
                                                    end if

                                                    if aktdata(iRowLoop, d+10) >= 10 then
                                                    bgThisDay = "red"
                                                        if aktdata(iRowLoop, d+10) >= 15 then
                                                        hgtThisDay = 15
                                                        end if
                                                    end if

                                                    hgtThisDay = (hgtThisDay * 2)
                                                    %>

                                                    <td valign=bottom> <div style="height:<%=hgtThisDay%>; background-color:<%=bgThisDay%>;"><img src="ill/blank.gif" width=10 height=<%=hgtThisDay%> border=0 alt="<%=aktdata(iRowLoop, d+10)%> timer"/></div></td>
                                                    
                                                    <%next %>

                                                    </tr></table>
                                                    
                                                    </td></tr></table>
                                                    

                                                    </td>
                    
                    

                                                    <td valign="bottom" class=lille style="background-color:#FFFFFF; width:300px; border-bottom:3px #8caae6 solid; height:4px; padding:2px 2px 0px 30px;">
                                                    <input type="text" style="width:258px; font-size:9px; font-family:Arial; color:#999999; font-style:italic;" maxlength="101" value="Job tweet..(åben for alle)" class="FM_job_komm" name="FM_job_komm_<%=aktdata(iRowLoop, 4)%>" id="FM_job_komm_<%=aktdata(iRowLoop, 4)%>"> <a href="#" class="aa_job_komm" id="aa_job_komm_<%=aktdata(iRowLoop, 4)%>">Gem</a>
                                                    <div id="dv_job_komm_<%=aktdata(iRowLoop, 4)%>" style="width:258px; font-size:9px; font-family:Arial; color:#000000; font-style:italic; overflow-y:auto; overflow-x:hidden; height:30px;"><%=aktdata(iRowLoop, 58) %></div>
                                                    </td>
                                                    <td valign="top" align=right class=lille style="background-color:#FFFFFF; width:40px; padding:2px 2px 0px 10px; border-bottom:3px #8caae6 solid;">
                                                    <span style="color:#999999; font-size:9px; font-style:oblique;">Fjern<br /><input class="fjern_job" type="checkbox" value="1" id="fjern_job_<%=aktdata(iRowLoop, 4)%>"/></span></td>
                                                </tr>
               
                                            </table>
                                            </div>
                                            <%response.flush %>

                                </td>
                                </tr>

                                <tr><td>

                                 <!-- job detaljer / stamdata -->
                                <%
                                call jobbeskrivelse_stdata

                                %>
                                </td>
                                </tr>






                                <tr><td>
                                <%


                                '**** Lastfakdato ******'

                                
					                    lastfakdato = "1/1/2001"
					
					                    strSQLFAK = "SELECT f.fakdato FROM fakturaer f WHERE f.jobid = "& aktdata(iRowLoop, 4) &" AND f.fakdato >= '"& varTjDatoUS_man &"' AND faktype = 0 ORDER BY f.fakdato DESC"
					                    
                                        'Response.Write strSQLFAK
                                        'Response.flush       
                    
                                        oRec2.open strSQLFAK, oConn, 3
					                    if not oRec2.EOF then
						                    if len(trim(oRec2("fakdato"))) <> 0 then
						                    lastfakdato = oRec2("fakdato")
						                    end if
					                    end if
					                    oRec2.close
	                

	                                'call dageDatoer()
	          
            
                                %>
                                <input type="hidden" id="FM_job_lastfakdato_<%=aktdata(iRowLoop, 4)%>" value="<%=lastfakdato %>" />
                                <%




                                '*** end lastFakdato ****'

          
			                    'if request.Cookies("job_"&aktdata(iRowLoop, 4)&"") <> "visible" then 'AND antalJobLinier <> 0
                                job_vzb = "hidden"
			                    job_dsp = "none"
                                'else
			                    'job_vzb = "visible"
			                    'job_dsp = ""
			                    'end if
			
			
			
			                    '**** Aktiviteterne på hvert job ****'%>
                                <div id="div_timereg_<%=aktdata(iRowLoop, 4)%>" style="position:relative; background-color:#ffffff; height:350px; overflow:auto; padding:0px 3px 3px 3px; width:745px; border:3px #8caae6 solid; visibility:<%=job_vzb %>; display:<%=job_dsp%>; z-index:2000">
                                <!-- Henter aktiviteter --->
                                <table width=100%><tr><td style="padding:20px 20px 20px 20px;">Henter aktiviteter 1-4 sek...</td></tr></table>
                                

                            
                         
                                <%


                                 lastJobid = aktdata(iRowLoop, 4)
                                antalJobLinier = antalJobLinier + 1 

                                else 'Easyreg


                               
                              
                                end if





                                '************** End job ****'

            


            '***   akt linje **'
            '*** Fra JQ ******'
            



        lastAktid = aktdata(iRowLoop, 3)
        lastAktnavn = aktdata(iRowLoop, 2)

        if len(trim(aktdata(iRowLoop, 29))) <> 0 AND isNull(aktdata(iRowLoop, 29)) <> true then
	    fsVal = fsVal &"xx,xx "& aktdata(iRowLoop, 29)
	    else
	    fsVal = fsVal &"xx,xx nonexx99w"
        end if

        strAktidsThisJob = strAktidsThisJob & ","& aktdata(iRowLoop, 3)
       
        iRowLoopsThisJob = iRowLoopsThisJob & ","& iRowLoop



          

      
        end if 'lastAktid

            
          
         

    			
          
    
	  
	next
	
	
	
    
	%>
   
    
	
	<%end if 'iRowLoop%>
    <%end if 'Easyreg %>
	<!--</tr>-->

  
            <%
            if cint(intEasyreg) <> 1 then
            
                                            if iRowLoop <> 0 then
                                            %>
            
           
               
                                            </div>
                                            <!-- table Div-->

             


                                            <%end if 



                        
                            call jq_format(faVal)
                            faVal = jq_formatTxt %>

                            <input type="hidden" id="job_aktids_<%=lastJobid%>" value="<%=strAktidsThisJob %>" />
                            <input name="fasenavn" id="job_fasenavne_<%=lastJobid %>" type="hidden" value="<%=fsVal %>" />
                            <input type="hidden" id="FM_jobid_<%=antalJobLinier-1 %>" value="<%=lastJobid %>" /> 
            
           
                            <input name="iRowLoop" id="job_iRowLoop_<%=lastJobid%>" type="hidden" value="<%=iRowLoopsThisJob%>" />

           
                            </td></tr>
           
                            </table>

            <%
            
            else '** Easyreg

                                'if antalJobLinier <> 0 then%>
                                <tr><td>

                                                    <input type="hidden" id="FM_job_lastfakdato_0" value="2001-1-1" />
                                                        <input type="hidden"  id="job_aktids_0" value="<%=strEaAktidsThisJob %>" />
                                                      <input type="hidden" name="fasenavn" id="job_fasenavne_0" value="<%=fsVal %>" />
 
                                                    <input type="hidden" name="iRowLoop" id="job_iRowLoop_0" value="<%=iRowLoopsEAThisJob %>" /> 

                                  <div id="div_timereg_0" style="position:relative; background-color:#ffffff; height:1050px; overflow:auto; padding:0px 3px 3px 3px; width:745px; border:3px #8caae6 solid; visibility:visible; display:; z-index:2000">
                                <!-- Henter aktiviteter --->
                           
                                <table width=100%><tr><td style="padding:20px 20px 20px 20px;">Henter Easyreg. aktiviteter (maks. 150) 5-10 sek...<br /><br />&nbsp;</td></tr></table>
                                
                                
                                </div>
                                 </td></tr>
           
                            </table>
                                <%
                                'end if
                                'lastJobid = aktdata(iRowLoop, 4)
                                antalJobLinier = antalJobLinier + 1 
        
            
            end if

          
           






            if antalJobLinier = 0 then %>

            <br /><br />
             <div id="Div6" class="jqcorner" style="position:relative; background-color:#ffffff; padding:3px 3px 3px 3px; width:745px; border:10px #CCCCCC solid; left:0px; visibility:visible; display:; z-index:1000">
            <table cellspacing="0" cellpadding="2" border="0" width="100%"><tr><td style="padding:20px 3px 20px 3px;">
            <h4>Der kan ikke vises nogen job!</h4>

            <b>Det sklydes mest sandsynligt en af nedenstående årsager:</b><br /><br />
             - Du mangler at vælge job i søgefilteret ovenfor og klikke på "Vis job >>"<br />
            - Jobbets startdato ligger i en senere uge end den valgte<br />
            - Du har ikke adgang til de(t) valgte job / aktiviteter<br />
            - Jobbet er ikke valgt i din personlige aktivliste</td></tr></table>
            </div>

               <%end if %>

             <br /><br /><br /><br /><br /><br />


             <h4>Ugetotaler:</h4>
     <div id="dagstotaler" class="jqcorner" style="position:relative; background-color:#ffffff; padding:3px 3px 3px 3px; width:745px; border:1px #8caae6 solid; left:0px; visibility:visible; display:; z-index:1000">
   


     <table cellspacing="0" cellpadding="2" border="0" width="100%" bgcolor="#cccccc">

     <% call dageDatoer(2) %>
     <%=dageDatoerTxt %>
     

     <!--
	<tr bgcolor="#FFFFFF">
		<td colspan=2 style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;">Timer på viste job:</td>
		<td style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><input id="timer_1" type="text" size=4 value="<%=formatnumber(totTimMan, 2)%>" style="border:0px #8caae6 solid; background-color:#FFFFFF; font-size:10px;" /></td>
	<td style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><input id="timer_2" type="text" size=4 value="<%=formatnumber(totTimTir, 2)%>" style="border:0px #8caae6 solid; background-color:#FFFFFF; font-size:10px;" /></td>
	<td style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><input id="timer_3" type="text" size=4 value="<%=formatnumber(totTimOns, 2)%>" style="border:0px #8caae6 solid; background-color:#FFFFFF; font-size:10px;" /></td>
	<td style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><input id="timer_4" type="text" size=4 value="<%=formatnumber(totTimTor, 2)%>" style="border:0px #8caae6 solid; background-color:#FFFFFF; font-size:10px;" /></td>
	<td style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><input id="timer_5" type="text" size=4 value="<%=formatnumber(totTimFre, 2)%>" style="border:0px #8caae6 solid; background-color:#FFFFFF; font-size:10px;" /></td>
	<td bgcolor="#cccccc" style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><input id="timer_6" type="text" size=4 value="<%=formatnumber(totTimLor, 2)%>" style="border:0px #8caae6 solid; background-color:#cccccc; font-size:10px;" /></td>
	<td bgcolor="#cccccc" style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><input id="timer_7" type="text" size=4 value="<%=formatnumber(totTimSon, 2)%>" style="border:0px #8caae6 solid; background-color:#cccccc; font-size:10px;" /></td>
	</tr>
    -->
	
	
    <% 

    manTimer = 0
    tirTimer = 0
    onsTimer = 0
    torTimer = 0
    freTimer = 0
    lorTimer = 0
    sonTimer = 0

     for x = 1 to 7
      

            select case x
                case 1
                tjkTimerTotDato = varTjDatoUS_man
                case 2
                tjkTimerTotDato = varTjDatoUS_tir
                case 3
                tjkTimerTotDato = varTjDatoUS_ons
                case 4
                tjkTimerTotDato = varTjDatoUS_tor
                case 5
                tjkTimerTotDato = varTjDatoUS_fre
                case 6
                tjkTimerTotDato = varTjDatoUS_lor
                case 7
                tjkTimerTotDato = varTjDatoUS_son
                end select



                sel