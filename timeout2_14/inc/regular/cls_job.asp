
<%

    function lukkedato(jobId, jobStatus)

strLukkeDato = year(now) &"/"& month(now) &"/"& day(now)

 if jobStatus = 0 then 'Kun ved luk job
    strSQLjs = "SELECT jobstatus FROM job WHERE id = "& jobId
    oRec5.open strSQLjs, oConn, 3
    if not oRec5.EOF then
                                        
        if oRec5("jobstatus") <> 0 then 'job bliver lukket nu

        strSQLjlukkedt = "UPDATE job SET lukkedato = '"& strLukkeDato &"' WHERE id = " & jobId
        oConn.execute(strSQLjlukkedt)

        end if


    end if
    oRec5.close 
end if


end function




    '***** Sync Job slutDato ***********'
    public useSyncDato
    function syncJobSlutDato(jobid, jobnr, syncslutdato)

    fod = 0

     '***** slutdato på job følger altid sidste timereg. / sidste fakturadato ellers dagsdato ****'
    if syncslutdato = 1 then

    '** Sidste timeregdato **
    lastTregDato = "1-1-2002"
    strSQLsidstetreg = "SELECT tdato FROM timer WHERE tjobnr = '"& jobnr &"' ORDER BY tdato DESC"
    Response.write strSQLsidstetreg & "<br>"
    oRec5.open strSQLsidstetreg, oConn, 3
    if not oRec5.EOF then
                                
    lastTregDato = oRec5("tdato")

    fod = 1
    end if
    oRec5.close 

    '** Sidste fakdato **
    lastFakDato = "1-1-2002"
    strSQLsidstefak = "SELECT if(brugfakdatolabel = 1, labeldato, fakdato) AS lastfaktdato FROM fakturaer WHERE jobid = "& jobid &" AND shadowcopy <> 1 AND faktype = 0 ORDER BY fid DESC"
    oRec5.open strSQLsidstefak, oConn, 3
    if not oRec5.EOF then
                                
    lastFakDato = oRec5("lastfaktdato")

    fod = 1
    end if
    oRec5.close 


    'Response.write "fod " & fod & " lastTregDato: "& lastTregDato & " lastFakDato: "& lastFakDato &"<br>" 

    '*** Er der fundet en ****
    if fod = 1 then

            if cDate(lastTregDato) > cDate(lastFakDato) then
            useSyncDato = lastTregDato
            else
            useSyncDato = lastFakDato
            end if 

    else
    '*** else dd **'
    useSyncDato = year(now) & "/" & month(now) & "/" & day(now)

    end if

    useSyncDato = year(useSyncDato) & "/" & month(useSyncDato) & "/" & day(useSyncDato) 

    strSQLjobSyncDato = "UPDATE job SET jobslutdato = '"& useSyncDato &"' WHERE id = "& jobid    

    'Response.write strSQLjobSyncDato

    oConn.execute(strSQLjobSyncDato)

    'Response.end

    end if
    '****

    end function



     public totTerminBelobJob, totTerminBelobGrand
    function terminbelob(jobid, showhtml)

    

        totTerminBelobJob = 0
		strSQLm = "SELECT milepale.id AS id, milepale.navn AS navn, milepal_dato, "_
		&" milepale_typer.navn AS type, ikon, beskrivelse, milepale.editor, belob FROM milepale "_
		&" LEFT JOIN milepale_typer ON (milepale_typer.id = milepale.type) "_
		&" WHERE jid = "& jobid &" ORDER BY milepal_dato"
		x = 0
		oRec4.open strSQLm, oConn, 3
		while not oRec4.EOF 
		
        if showhtml = 1 then

		select case right(x, 1)
		case 0,2,4,6,8
		bgthis = "#FFFFFF"
		case else
		bgthis = "#EFF3FF"
		end select
		%>
		<tr bgcolor="<%=bgthis %>">
			<td valign="top" style="border-bottom:1px #CCCCCC solid; font-size:8px; white-space:nowrap;">
			<%=formatdatetime(oRec4("milepal_dato"), 2)%><br />
		    <a href="javascript:popUp('milepale.asp?func=red&id=<%=oRec4("id")%>&jid=<%=jobid%>&rdir=wip&type=1','650','500','250','120');" target="_self" class=rmenu><%=left(oRec4("navn"),10)%></a></td>
            <td valign="top" style="border-bottom:1px #CCCCCC solid;" align=right class=lille><%=formatnumber(oRec4("belob"),0) %></td>
              <td valign="top" style="border-bottom:1px #CCCCCC solid;" align=right class=lille> <a href="javascript:popUp('milepale.asp?func=slet&id=<%=oRec4("id")%>&jid=<%=oRec("id")%>&rdir=wip&type=1','650','500','250','120');" target="_self" class=red>x</a></td>
			
			
		</tr>
		
		<%
        end if

        totTerminBelobJob = totTerminBelobJob + oRec4("belob")
        totTerminBelobGrand = totTerminBelobGrand + oRec4("belob")

        Response.flush

		x = x + 1
		oRec4.movenext
		wend
		oRec4.close
		

    end function



 function lukjobmail(jstatus, lukjob, medarbejderid)
 'for j = 0 to UBOUND(lukjob)
            				
                            'if lukjob(j) <> 0 then
                            
                            if jstatus = 0 then
                            lukjobSQL = ", lukkedato = '"& year(now)&"/"& month(now)&"/"& day(now) &"'"
                            else
                            lukjobSQL = ""
                            end if

                            '*** Skifter status ****'
                            strSQLst = "UPDATE job SET jobstatus = "& jstatus &" "& lukjobSQL &" WHERE id = "& lukjob  'jobnr = "& intJobnr
                            'Response.write strSQLst 
                            'Response.end
		                    oConn.execute(strSQLst)

                            '** Ny jobstatus 
                            select case jstatus
                            case 0
                            jstatusTxt = "Lukket"
                            case 1
                            jstatusTxt = "Aktiv"
                            case 2
                            jstatusTxt = "Passiv / Til Fakturering"
                            case 3
                            jstatusTxt = "Tilbud"
                            case 4
                            jstatusTxt = "Til gennemsyn"
                            end select


				            '**** Finder jobansvarlige *****
				            strSQL = "SELECT job.id AS jid, jobnavn, jobnr, jobans1, jobans2, m1.mnavn AS m1mnavn, m1.email AS m1email,"_
				            &" m2.mnavn AS m2mnavn, m2.email AS m2email, kkundenavn, kkundenr FROM job "_
				            &" LEFT JOIN medarbejdere AS m1 ON (m1.mid = jobans1)"_
				            &" LEFT JOIN medarbejdere AS m2 ON (m2.mid = jobans2)"_
                            &" LEFT JOIN kunder AS k ON (k.kid = job.jobknr)"_
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
                            strkkundenavn = oRec("kkundenavn")
                            'kkundenr = oRec("kkundenr")
            				
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
		                            'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                            ' Sætter Charsettet til ISO-8859-1
		                            'Mailer.CharSet = 2

                                    Set myMail=CreateObject("CDO.Message")
                                    myMail.From="timeout_no_reply@outzource.dk" 'TimeOut Email Service 

		                            'Mailer.FromName = "TimeOut | " & afsNavn 
            
                                    'select case lto
                                    'case "hestia"
                                    'Mailer.FromAddress = "timeout_no_reply@outzource.dk"
                                    'case else
		                            'Mailer.FromAddress = afsEmail
                                    'end select
		                            
                                    'Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk" '"pasmtp.tele.dk"
						            'Mailer.AddRecipient "" & jobans1 & "", "" & jobans1email & ""
		                            'Mailer.AddRecipient "" & jobans2 & "", "" & jobans2email & ""

                                    myMail.To= ""& jobans1 &"<"& jobans1email &">"
                                    myMail.Cc= ""& jobans2 &"<"& jobans2email &">"

                                    
            						
                                    if lto = "synergi1" then
                                    'Mailer.AddRecipient "Bettina Schou", "bs@synergi1.dk"
                                    myMail.Bcc= "Bettina Schou<bs@synergi1.dk>"
                                    'Mailer.AddRecipient "TimeOut Support", "support@outzource.dk"
                                    end if

                                    if lto = "outz" then
                                    'Mailer.AddRecipient "TimeOut Support", "support@outzource.dk"
                                    myMail.Bcc= "TimeOut Support<support@outzource.dk>"
                                    end if

                                    if lto = "dencker" then
                                    'Mailer.AddRecipient "Anders Dencker", "ad@dencker.net"
                                    myMail.Bcc= "Anders Dencker<ad@dencker.net>"
                                    end if

                                    'if lto = "fe" then
                                    'Mailer.AddRecipient "Lone P", "lrp@feteknik.dk"
                                    'end if

                                    if lto = "essens" then
                                    'Mailer.AddRecipient "Astrid", "astrid@essens.info"
                                    myMail.Bcc= "Bolette Westh<bolette@essens.info>"
                                    end if

            						    
						            'Mailer.Subject = strkkundenavn &", "& jobnavnThis &" ("& intJobnr &") - Afsluttet."
                                    myMail.Subject= strkkundenavn &", "& jobnavnThis &" ("& intJobnr &") - Afsluttet af "& session("user")
		                           
                                    strBody = "Hej jobansvarlige.<br><br>"
                                    strBody = strBody &"Mit arbejde på:<br><br>"
                                    strBody = strBody &"Kunde: "& strkkundenavn &"<br>"
						            strBody = strBody &"Job: "& jobnavnThis &" ("& intJobnr &") er nu afsluttet.<br><br>"
                                    strBody = strBody &"Jobstatus er: "& jstatusTxt &"<br><br>"
		                            strBody = strBody &"Med venlig hilsen<br>"
		                            strBody = strBody & session("user") & "<br><br>"
            		                
            		
                                    myMail.HTMLBody= "<html><head></head><body>" & strBody & "</body>"

                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                                    'Name or IP of remote SMTP server
                                    
                                                if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                                   smtpServer = "webout.smtp.nu"
                                                else
                                                   smtpServer = "formrelay.rackhosting.com" 
                                                end if
                    
                                                myMail.Configuration.Fields.Item _
                                                ("http://schemas.microsoft.com/cdo/configuration/smtpserver")= smtpServer

                                    'Server port
                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
                                    myMail.Configuration.Fields.Update

                                    if len(trim(jobans1email)) <> 0 then
                                    myMail.Send
                                    end if

                                    set myMail=nothing



            		
		                            'Mailer.BodyText = strBody
            		
		                            'Mailer.sendmail()
		                            'Set Mailer = Nothing

                                    Response.Write "mail afsendt - " & request.form("jq_lukjob")
				            
                            end if' x



    end function

 %>