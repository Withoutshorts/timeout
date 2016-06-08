<%


'********** Er uge afsluttet??? ******'
	public ugeNrAfsluttet, showAfsuge, cdAfs, ugegodkendt, ugegodkendtaf, ugegodkendtTxt, ugegodkendtdt 
    function erugeAfslutte(sm_aar, sm_sidsteugedag, sm_mid)
            
            showAfsuge = 1
            ugeNrAfsluttet = "1-1-2044"
            
            strSQLafslut = "SELECT status, afsluttet, uge, ugegodkendt, ugegodkendtaf, ugegodkendtTxt, ugegodkendtdt FROM ugestatus WHERE WEEK(uge, 1) = "& sm_sidsteugedag &" AND YEAR(uge) = "& sm_aar &" AND mid = "& sm_mid
		    
            'if session("mid") = 1 then
            'Response.write strSQLafslut & "//" & sm_sidsteugedag & "//"& sldatoSQL
		    'Response.flush
            'end if
            
            oRec3.open strSQLafslut, oConn, 3 
		    if not oRec3.EOF then
    			
			    showAfsuge = 0
			    cdAfs = oRec3("afsluttet")
			    ugeNrAfsluttet = oRec3("uge")


                ugegodkendt = oRec3("ugegodkendt")
                ugegodkendtaf = oRec3("ugegodkendtaf")
                ugegodkendtTxt = oRec3("ugegodkendtTxt")
                ugegodkendtdt = oRec3("ugegodkendtdt")
    		
		    end if
		    oRec3.close 
            
            
    end function


function godekendugeseddel(thisfile, godkenderID, medarbid, stDatoSQL)
       
        'gkweekfundet = 0
        ugegodkendtdtnow = year(now) & "/" & month(now) & "/" & day(now)
        denneuge = datepart("ww", stDatoSQL, 2, 2)
	    detteaar = datepart("yyyy", stDatoSQL, 2,3)

        ugstatusSQL = "SELECT id, mid FROM ugestatus AS u WHERE mid = "& medarbid & " AND YEAR(u.uge) = '"& detteaar &"' AND  WEEK(u.uge, 1) = '"& denneuge &"'"
        'Response.Write ugstatusSQL
        'Response.end


        oRec6.open ugstatusSQL, oConn, 3
        if not oRec6.EOF then
        
            strSQLup = "UPDATE ugestatus SET ugegodkendt = 1, ugegodkendtaf = "& godkenderID &", ugegodkendtdt = '"& ugegodkendtdtnow &"' WHERE id = "& oRec6("id") 
	        'Response.Write strSQLup
            'Response.end
            oConn.execute(strSQLup)

        'gkweekfundet = 1
        end if
        oRec6.close

        
      

end function


function afviseugeseddel(thisfile, afsenderMid, modtagerMid, varTjDatoUS_man, varTjDatoUS_son, txt)

                            '*** Henter afsender **
				            strSQL = "SELECT mnavn, email FROM medarbejdere"_
				            &" WHERE mid = "& afsenderMid
				            oRec.open strSQL, oConn, 3
            				
				            if not oRec.EOF then
            				
				            afsNavn = oRec("mnavn")
				            afsEmail = oRec("email")
            				
				            end if
				            oRec.close

                             '*** Henter modtager **
				            strSQL = "SELECT mnavn, email FROM medarbejdere"_
				            &" WHERE mid = "& modtagerMid
				            oRec.open strSQL, oConn, 3
            				
				            if not oRec.EOF then
            				
				            modtNavn = oRec("mnavn")
				            modtEmail = oRec("email")
            				
				            end if
				            oRec.close


                            '*** Afvis ugeseddel ænder ikke statusd på timerne ***'
                            'strSQLup = "UPDATE timer SET godkendtstatus = 2, godkendtstatusaf = '"& afsNavn &"' WHERE tmnr = "& modtagerMid & " AND tdato BETWEEN '"& varTjDatoUS_man &"' AND '" & varTjDatoUS_son & "'" 
	                        'oConn.execute(strSQLup)

                            ugegodkendtTxt = txt
                              
                              ugegodkendtdtnow = year(now) & "/" & month(now) & "/" & day(now)
                              denneuge = datepart("ww", varTjDatoUS_man, 2, 2)
	                          detteaar = datepart("yyyy", varTjDatoUS_man, 2,3)

                            strSQLupUgeseddel = "UPDATE ugestatus SET ugegodkendt = 2, ugegodkendtaf = "& afsenderMid &", ugegodkendtdt = '"& ugegodkendtdtnow &"', ugegodkendtTxt = '"& ugegodkendtTxt &"' WHERE mid = "& modtagerMid & "  AND YEAR(uge) = '"& detteaar &"' AND  WEEK(uge, 1) = '"& denneuge &"'"
	                        'Response.Write strSQLup
                            'Response.flush
                            oConn.execute(strSQLupUgeseddel)
	    
	    'Response.Write strSQLup
	    'Response.end
	    

          if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\"&thisfile then

                                    
                           
                                    
                                    thisWeek = datepart("ww", varTjDatoUS_man, 2, 2)    

					  	            'Sender notifikations mail
		                            Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                            ' Sætter Charsettet til ISO-8859-1
		                            Mailer.CharSet = 2
		                            Mailer.FromName = "TimeOut | " & afsNavn 
		                            Mailer.FromAddress = afsEmail
		                            Mailer.RemoteHost = "webmail.abusiness.dk" '"pasmtp.tele.dk"
						            'Mailer.AddRecipient "" & afsNavn & "", "" & afsEmail & ""
		                            Mailer.AddRecipient "" & modtNavn & "", "" & modtEmail & ""
            						
                                    
                                    Mailer.Subject = "Ugeseddel uge: "& thisWeek &" - afvist"
		                            strBody = "Hej " & modtNavn &  vbCrLf & vbCrLf
                                    strBody = strBody &"Din ugeseddel uge: "& thisWeek &" - er afvist" & vbCrLf & vbCrLf
                                    strBody = strBody &"Begrundelse: " & vbCrLf 
						            strBody = strBody & txt & vbCrLf & vbCrLf
		                            strBody = strBody &"Med venlig hilsen" & vbCrLf
		                            strBody = strBody & afsNavn & vbCrLf & vbCrLf
            		                
            		
            		
		                            Mailer.BodyText = strBody
            		
		                            Mailer.sendmail()
		                            Set Mailer = Nothing

                                   
				            
                            end if' x


end function


function afslutugereminder(thisfile, afsenderMid, modtagerMid, varTjDatoUS_man, varTjDatoUS_son, txt)

                            '*** Henter afsender **
				            strSQL = "SELECT mnavn, email FROM medarbejdere"_
				            &" WHERE mid = "& afsenderMid
				            oRec.open strSQL, oConn, 3
            				
				            if not oRec.EOF then
            				
				            afsNavn = oRec("mnavn")
				            afsEmail = oRec("email")
            				
				            end if
				            oRec.close

                             '*** Henter modtager **
				            strSQL = "SELECT mnavn, email FROM medarbejdere"_
				            &" WHERE mid = "& modtagerMid
				            oRec.open strSQL, oConn, 3
            				
				            if not oRec.EOF then
            				
				            modtNavn = oRec("mnavn")
				            modtEmail = oRec("email")
            				
				            end if
				            oRec.close


                          
	    

          if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\"&thisfile then

                                    
                           
                                    
                                    thisWeek = datepart("ww", varTjDatoUS_man, 2, 2)    

					  	            'Sender notifikations mail
		                            Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                            ' Sætter Charsettet til ISO-8859-1
		                            Mailer.CharSet = 2
		                            Mailer.FromName = "TimeOut | " & afsNavn 
		                            Mailer.FromAddress = afsEmail
		                            Mailer.RemoteHost = "webmail.abusiness.dk" '"pasmtp.tele.dk"
						            'Mailer.AddRecipient "" & afsNavn & "", "" & afsEmail & ""
		                            Mailer.AddRecipient "" & modtNavn & "", "" & modtEmail & ""
            						
                                    
                                    Mailer.Subject = "Du har endnu ikke afslutte din ugeseddel uge: "& thisWeek 
		                            strBody = "Hej " & modtNavn &  vbCrLf & vbCrLf
                                    strBody = strBody &"Din ugeseddel uge: "& thisWeek &" - er endnu ikke afsluttet, husk at få den afsluttet snarest." & vbCrLf & vbCrLf
                                   
		                            strBody = strBody &"Med venlig hilsen" & vbCrLf
		                            strBody = strBody & afsNavn & vbCrLf & vbCrLf
            		                
            		
            		
		                            Mailer.BodyText = strBody
            		
		                            Mailer.sendmail()
		                            Set Mailer = Nothing

                                   
				            
                            end if' x


end function

public ntimPer, ntimMan, ntimTir, ntimOns, ntimTor, ntimFre, ntimLor, ntimSon, antalDageMtimer
public ntimManIgnHellig, ntimTirIgnHellig, ntimOnsIgnHellig, ntimTorIgnHellig, ntimFreIgnHellig, ntimLorIgnHellig, ntimSonIgnHellig
	function normtimerPer(medid, stDato, interval)
	
	'Response.Write "Kalder normtidpper<br>"
	
	'** Maks 30 typer. **'
	dim mtyperIntvDato, mtyperIntvTyp, mtyperIntvEndDato 
	redim mtyperIntvDato(30), mtyperIntvTyp(30), mtyperIntvEndDato(30)
	
	'** Aktuel medarbejdertype **'
	strSQLtype = "SELECT medarbejdertype, ansatdato, opsagtdato FROM medarbejdere WHERE mid = "& medid
	oRec2.open strSQLtype, oConn, 3
	if not oRec2.EOF then
	
	mtyperIntvTyp(0) = oRec2("medarbejdertype")
    mtyperIntvDato(0) = oRec2("ansatdato")
    mtyperIntvEndDato(0) = oRec2("opsagtdato")
	
	end if
	oRec2.close
	
	
	'** Hvis medarbjeder ansat dato er før licens startdato benyttes licens startDato **'
	call licensStartDato()
	ltoStDato = startDatoDag &"/"& startDatoMd &"/"& startDatoAar
	if cdate(ltoStDato) > cDate(mtyperIntvDato(0)) then
	mtyperIntvDato(0) = ltoStDato
	end if
	
    
    'Response.Write mtyperIntvDato(0) &" "& ltoStDato 
	 
	'*** Finder medarbejdertyper_historik ***'
    strSQLmth = "SELECT mid, mtype, mtypedato FROM medarbejdertyper_historik WHERE "_
    &" mid = "& medid &" ORDER BY mtypedato, id"
    
    'Response.Write strSQLmth
    'Response.flush
    
    t = 1
    oRec2.open strSQLmth, oConn, 3
    while not oRec2.EOF 
    
    mtyperIntvTyp(t) = oRec2("mtype")
    mtyperIntvDato(t) = oRec2("mtypedato")
    
    t = t + 1
    oRec2.movenext
    wend
    oRec2.close
    
    'Response.Write "t: " & t & " interval = "& interval &"<br>"
    'Response.Write mtyperIntvDato(0) &" "& ltoStDato 
    '*****************************************'
	
			
			stDatoDag = weekday(stDato, 2)
			ntimPer = 0
            antalDageMtimer = 0
			
			for c = 0 to interval
			
			            if c = 0 then
			            n = stDatoDag
			            datoCount = stDato
			            else
			                if n < 7 then
			                n = n + 1
			                else
			                n = 1
			                end if
			            datoCount = dateAdd("d", 1, datoCount)
			            end if
			            
			            'Response.Write "her: " & cDate(datoCount) &" >= "& cDate(mtyperIntvDato(0))

			            '*** Tjekker ansatDato **'
			            '*** Skal først tjekke dage efter ansat dato, og før opsagt dato ***'

			            if cDate(datoCount) >= cDate(mtyperIntvDato(0)) AND cDate(datoCount) < cDate(mtyperIntvEndDato(0)) then
			            
			            
			            'Response.Write "<u>cdate(datoCount) </u>"& cdate(datoCount) & "<br>"
			            
			            '*** Finder den medarbejder typer der passer til den valgte dato ***'
			            if t = 1 then
			            mtypeUse = mtyperIntvTyp(0) 
			            'Response.Write  " A mtypeUse: "&  mtypeUse & "<br>"
			            else
			               
			                for d = 2 to t  
			                    
			                    
			                    if cdate(datoCount) >= mtyperIntvDato(d-1) then 
    			                mtypeUse = mtyperIntvTyp(d-1)
    			                'Response.Write  cdate(datoCount) & " >= "& mtyperIntvDato(d-1)  &" - B mtypeUse: "&  mtypeUse & "<br>"
    			                else
    			                    if d = 2 then
    			                    mtypeUse = mtyperIntvTyp(d-1)
    			                    'Response.Write  cdate(datoCount) & " - C1 mtypeUse: "&  mtypeUse & "<br>"
			                    
    			                    else
			                        mtypeUse = mtypeUse
			                        'Response.Write  cdate(datoCount) & " - C mtypeUse: "&  mtypeUse & "<br>"
			                    
			                        end if
			                    end if
			            
			                next
			            
			            end if
			            
			           
			
			            select case n 
						case 7
						normdag = "normtimer_son"
						nd1 = 0
						case 1
						normdag = "normtimer_man"
						nd2 = 0
						case 2
						normdag = "normtimer_tir"
						nd3 = 0
						case 3
						normdag = "normtimer_ons"
						nd4 = 0
						case 4
						normdag = "normtimer_tor"
						nd5 = 0
						case 5
						normdag = "normtimer_fre"
						nd6 = 0
						case 6
						normdag = "normtimer_lor"
						nd7 = 0
						end select
						
					
					'** Finder normtimer på den valgte dag for den vagte type **'	
                    if len(trim(mtypeUse)) <> 0 then
                    mtypeUse = mtypeUse
                    else
                    mtypeUse = 0
                    end if

					strSQLnt = "SELECT "& normdag &" AS timer FROM "_
					&" medarbejdertyper t WHERE t.id = "& mtypeUse 
				
					
					'if medid = 1 then
					'Response.write strSQLnt & " // n:"& n & "<hr>"
					'Response.flush
					'end if 
					
					oRec3.open strSQLnt, oConn, 3 
					if not oRec3.EOF then
						
						ntimPerThis = oRec3("timer")
						
						select case n 
						case 7
						ntimSon = ntimPerThis
                        ntimSonIgnHellig = ntimSon
						case 1
						ntimMan = ntimPerThis
                        ntimManIgnHellig = ntimMan
						case 2
						ntimTir = ntimPerThis
                        ntimTirIgnHellig = ntimTir
						case 3
						ntimOns = ntimPerThis
						ntimOnsIgnHellig = ntimOns
                        case 4
						ntimTor = ntimPerThis
						ntimTorIgnHellig = ntimTor
                        case 5
						ntimFre = ntimPerThis
						ntimFreIgnHellig = ntimFre
                        case 6
						ntimLor = ntimPerThis
						ntimLorIgnHellig = ntimLor
                        end select
						
                        'Response.Write "ntimPerThis:" & ntimPerThis & "<br>"
						
					end if
					oRec3.close 
					
					'Response.write "datoCount " & formatdatetime(datoCount, 2) & "<br>"
					'Response.end

                    
          

					
				    call helligdage(datoCount, 0)
				        
				        if cint(erHellig) = 1 then
				            select case n 
						    case 7
						    ntimSon = 0
						    case 1
						    ntimMan = 0
						    case 2
						    ntimTir = 0
						    case 3
						    ntimOns = 0
						    case 4
						    ntimTor = 0
						    case 5
						    ntimFre = 0
						    case 6
						    ntimLor = 0
						    end select

                            ntimPerThis = 0

						end if
					
					if erHellig <> 1 then
					ntimPer = ntimPer + ntimPerThis	
                            
                            if ntimPerThis <> 0 then
                            antalDageMtimer = antalDageMtimer + 1
                            'Response.Write "antalDageMtimer:" & antalDageMtimer & " ntimPerThis: "& ntimPerThis &"<br>"
                            end if 

					else
					ntimPer = ntimPer
					end if
					
					 
					
					
					
					end if '** Ansat Dato **'
					
							
		    next
			
			if len(ntimPer) <> 0 then
			ntimPer = ntimPer
			else
			ntimPer = 0
			end if

            if antalDageMtimer <> 0 then
            antalDageMtimer = antalDageMtimer
            else
            antalDageMtimer = 1
            end if
			
			
            'Response.write "her"
            'Response.end

	end function


public normtimerStDag
	function nortimerStandardDag(mtypeUse)
	
	
	ntimStdag = 0
	workdays_divider = 0
	
	    for n = 1 to 7 
	                    select case n 
						case 7
						normdag = "normtimer_son"
						case 1
						normdag = "normtimer_man"
						case 2
						normdag = "normtimer_tir"
						case 3
						normdag = "normtimer_ons"
						case 4
						normdag = "normtimer_tor"
						case 5
						normdag = "normtimer_fre"
						case 6
						normdag = "normtimer_lor"
						end select
						
					
					'** Finder normtimer på den valgte dag for den vagte type **'	
					strSQLnt = "SELECT "& normdag &" AS timer FROM "_
					&" medarbejdertyper t WHERE t.id = "& mtypeUse 
				    
		            oRec3.open strSQLnt, oConn, 3 
					if not oRec3.EOF then
						
						if oRec3("timer") <> 0 then
						ntimStdag = ntimStdag + oRec3("timer")
						workdays_divider = workdays_divider + 1
						else
						ntimStdag = ntimStdag
						end if
						
					end if
					oRec3.close 
				
				next
	            
	            '** workdays_divider bør altid være = 5
                '** Da danske overenskomster altdi er baseret på en 5 dages arb. uge
	            if workdays_divider <> 0 then
	            workdays_divider = workdays_divider
	            else
	            workdays_divider = 1
	            end if
	            
	            normtimerStDag = ntimStdag/workdays_divider
	
	end function




function ugeseddel(medarbid,stdatoSQL,sldatoSQL,visning)

st_dato = day(stdatoSQL) &"-"& month(stdatoSQL) & "-"& year(stdatoSQL)
%>


 <!-- **** Ugeseddel *** -->
   <table cellpadding=0 cellspacing=0 border=0 width=100%>
   <tr><td style="padding:0px 10px 0px 10px;">
   
    <%
               if thisfile = "ugeseddel_2011.asp" then
               fmLink = "ugeseddel_2011.asp?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&varTjDatoUS_son="&varTjDatoUS_son 
               else
               fmLink = "weekpage_2010.asp?medarbid="&medarbid&"&st_dato="&st_dato
               end if


    oimg = "ikon_stat_joblog_48.png"
	oleft = 0
	otop = 0
	owdt = 500
	oskrift = tsa_txt_338 &" "& datepart("ww", sldatoSQL, 2, 2) &" - "& datepart("yyyy", sldatoSQL, 2, 2) 
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	call erugeAfslutte(datepart("yyyy", sldatoSQL,2,2), datepart("ww", sldatoSQL, 2, 2), medarbid) 
	        
	            if cint(showAfsuge) = 0 then
	            showAfsugeTxt = "Ja"
	            else
	            showAfsugeTxt = "Nej"
	            end if
	            %>
	
	
	Har medarbejder afsluttet uge via smiley: <b><%=showAfsugeTxt%></b>

	
	</td>
	<td align=right colspan=2>
	
	<%if visning = 1 then 'ugeseddel_2010 ..tregsiden %>

    <%if media <> "print" then %>
    <form action="<%=fmLink %>&func=opdaterstatus" method="post">
	<table cellpadding=0 cellspacing=0 border=0 width=80>
	<tr>
	<td valign=top align=right style="padding:0px 10px 0px 0px;"><a href="ugeseddel_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=prev_varTjDatoUS_man %>&varTjDatoUS_son=<%=prev_varTjDatoUS_son %>"><img src="../ill/nav_left_blue.png" border="0" /></a></td>
   <td style="padding-right:10px;" valign=top align=right><a href="ugeseddel_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=next_varTjDatoUS_man %>&varTjDatoUS_son=<%=next_varTjDatoUS_son %>"><img src="../ill/nav_right_blue.png" border="0" /></a></td>
	</tr>
	</table>
    <%end if %>
	
	<%else%>
	
    <form action="<%=fmLink %>&func=opdaterstatus" method="post">
	<table cellpadding=0 cellspacing=0 border=0 width=80>
	<tr>
	<td valign=top align=right style="padding:0px 10px 0px 0px;"><a href="weekpage_2010.asp?medarbid=<%=medarbid%>&st_Dato=<%=prevWeek%>&func=us"><img src="../ill/nav_left_blue.png" border="0" /></a></td>
   <td style="padding-right:10px;" valign=top align=right><a href="weekpage_2010.asp?medarbid=<%=medarbid%>&st_Dato=<%=nextWeek%>&func=us"><img src="../ill/nav_right_blue.png" border="0" /></a></td>
	</tr>
	</table>
	
	<%end if%>
	</td>
	
    </tr>
  
   <tr><td valign=top colspan=2 style="padding:10px;">


   
   <table cellpadding=2 cellspacing=0 border=0 width=100%>
   <tr bgcolor="#5582d2">
            <td class=alt style="height:20px;"><b><%=tsa_txt_183 %></b></td>
            <td class=alt><b><%=tsa_txt_237 %></b></td>
            <td class=alt><b><%=tsa_txt_022 %> (<%=tsa_txt_066 %>)</b><br />
            <b><%=tsa_txt_244 %></b></td>
            <td class=alt><b><%=tsa_txt_068 %></b><br /><%=tsa_txt_329 %><br /><%=tsa_txt_296 %></td>
            <td class=alt><b><%=tsa_txt_148 %></b></td>
            <td class=alt>Status
            <%if cint(ugegodkendt) <> 1 AND media <> "print" then  %>
            <br /><input type="submit" value="Opdater >>" style="font-size:8px;" />
            <%end if %></td>
   </tr>
  
   <%'*** Joblog denne uge ***' 
   
   call afsluger(medarbid, stdatoSQL, sldatoSQL)
   
   strSQL = "SELECT tid, taktivitetnavn, timer, tfaktim, tjobnavn, tjobnr, tdato, "_
   &" tknr, tknavn, tmnavn, tmnr, kkundenr, godkendtstatus, m.mnr, timerkom, init, a.fase, a.easyreg FROM timer "_
   &" LEFT JOIN medarbejdere AS m ON (m.mid = tmnr)"_
   &" LEFT JOIN aktiviteter AS a ON (a.id = taktivitetid)"_
   &" LEFT JOIN kunder K on (kid = tknr) WHERE tmnr = "& medarbid &" AND "_
   &" tdato BETWEEN '"& stdatoSQL &"' AND '"& sldatoSQL &"' ORDER BY tdato DESC "
   
   'Response.Write strSQL
   'Response.flush
   at = 0
   timertot = 0
   lwedaynm = ""
   timerDag = 0
   antalEasyReg = 0
   eaDag = 0
   atDag = 0
   'ugegodkendtTxt = ""
   
   oRec.open strSQL, oConn, 3
   while not oRec.EOF 
    
    select case right(at, 1)
    case 0,2,4,6,8
    bgcol = "#ffffff"
    case else
    bgcol = "#EFF3ff"
    end select
    
    call akttyper2009prop(oRec("tfaktim"))
    
    
    if lwedaynm <> weekdayname(weekday(oRec("Tdato"))) then
     if at <> 0 then%>
    <tr>
        <td colspan=5 align=right><%=left(lwedaynm, 3) %>. d <%=formatdatetime(lwedate, 1) %> <b><%=formatnumber(timerDag, 2) %> t.</b><br />
        <span style="color:#999999;">Fordelt på <%=atDag %> registreringer, heraf <%=eaDag %> Easyreg.</span><br /><br />&nbsp;</td>
        <td>&nbsp;</td>

    </tr>
    <%
    atDag = 0
    eaDag = 0
    timerDag = 0
    
    end if %>
    <tr bgcolor="#D6Dff5">
        <td colspan=6 valign=bottom style="height:30px; border-bottom:1px #8CAAE6 solid; padding:3px;"><b><%=weekdayname(weekday(oRec("Tdato"))) %></b></td>
    </tr>
    <%
    end if
    
    %>
   <tr bgcolor="<%=bgcol%>">
   <td class=lille valign=top style="border-bottom:1px #cccccc solid; padding-right:20px; height:20px; white-space:nowrap;"><%=left(weekdayname(weekday(oRec("Tdato"))), 3) %>. <%=day(oRec("Tdato")) &" "&left(monthname(month(oRec("Tdato"))), 3) &". "& right(year(oRec("Tdato")), 2)%></td>
   <td class=lille valign=top style="border-bottom:1px #cccccc solid; padding-right:20px;"><b><%=oRec("tmnavn") %></b> (<%=oRec("mnr") %>) <%if len(trim(oRec("init"))) <> 0 then %>
   - <%=oRec("init") %>
   <%end if %></td>
   <td class=lille valign=top style="border-bottom:1px #cccccc solid; padding-right:20px;"><b><%=oRec("tknavn") %> (<%=oRec("kkundenr") %>)</b><br />
   <%=oRec("tjobnavn") %> (<%=oRec("tjobnr") %>)</td>
   <td class=lille valign=top style="border-bottom:1px #cccccc solid;"><%=oRec("taktivitetnavn") %>
   
   <%if oRec("easyreg") <> 0 then %>
    <span style="color:#999999;">(E!)</span>
   <%antalEasyReg = antalEasyReg + 1  
   eaDag = eaDag + 1%>
   <%end if %>
   
    
    <%
	call akttyper(oRec("tfaktim"), 1)
	%>

    <span style="color:yellowgreen;">(<%=akttypenavn%>)</span>
	
   
   <%if len(trim(oRec("fase"))) <> 0 then %>
   <br />
   <span style="color:#5582d2;"><%=replace(oRec("fase"), "_", " ")%></span>
   <%end if %>

    <%if len(trim(oRec("timerkom"))) <> 0 then %>
    <br /><span style="color:#999999;">
    <%if len(oRec("timerkom")) > 100 then %>
    <i><%=left(oRec("timerkom"), 100) %>...</i>
    <%else %>
    <i><%=oRec("timerkom")%></i>
    <%end if %>
    </span>
    <%end if %>
   </td>
   <td class=lille align=right valign=top style="border-bottom:1px #cccccc solid;">
          <%
        '** er periode godkendt ***'
		        tjkDag = oRec("Tdato")
		        erugeafsluttet = instr(afslUgerMedab(oRec("tmnr")), "#"&datepart("ww", tjkDag,2,2)&"_"& datepart("yyyy", tjkDag) &"#")
                
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
                
                if oRec("godkendtstatus") = 1 then
                ugeerAfsl_og_autogk_smil = 1
                else
                ugeerAfsl_og_autogk_smil = ugeerAfsl_og_autogk_smil
                end if
                
                
                if (ugeerAfsl_og_autogk_smil = 0 AND cint(ugegodkendt) <> 1 AND media <> "print") OR (level = 1 AND media <> "print") then%>
                 <a href="#" onclick="Javascript:window.open('rediger_tastede_dage_2006.asp?id=<%=oRec("tid") %>', '', 'width=650,height=500,resizable=yes,scrollbars=yes')" class=rmenu><%=formatnumber(oRec("timer"), 2) %></a>
	            
                <%else %>
                 <%=formatnumber(oRec("timer"), 2) %>
                <%end if %>
                </td>
                <td align="center" valign="top" style="border-bottom:1px #cccccc solid; padding:2px 2px 2px 2px;" class=lille>
                
                <%
                if cint(ugegodkendt) <> 1 AND media <> "print" then 
                
                              erGk = 0
                             gkCHK0 = ""
                             gkCHK1 = ""
                             gkCHK2 = ""
                             gk2bgcol = "#999999"
                             gk1bgcol = "#999999"
                             gk0bgcol = "#999999"

                         select case cint(oRec("godkendtstatus"))
                         case 2
                         gkCHK2 = "CHECKED"
                         erGk = 2
                         gk2bgcol = "red"
                         case 1
                         gkCHK1 = "CHECKED"
                         erGk = 1
                         gk1bgcol = "green"
                         case else
                         gkCHK0 = "CHECKED"
                         erGk = 0
                         gk0bgcol = "#000000"
                         end select
                 
                         erGkaf = ""


                
                %>

                   <input name="ids" id="ids" value="<%=oRec("tid")%>" type="hidden" />
					   <span style="color:<%=gk1bgcol%>;">Godk:</span><input type="radio" name="FM_godkendt_<%=oRec("tid")%>" id="FM_godkendt1_<%=v%>" class="FM_godkendt_1" value="1" <%=gkCHK1 %>><br />
                       <span style="color:<%=gk2bgcol%>;">Afvist:</span><input type="radio" name="FM_godkendt_<%=oRec("tid")%>" id="FM_godkendt2_<%=v%>" class="FM_godkendt_2" value="2" <%=gkCHK2 %>><br />
                       <span style="color:<%=gk0bgcol%>;">Ingen:</span><input type="radio" name="FM_godkendt_<%=oRec("tid")%>" id="FM_godkendt0_<%=v%>" value="0" class="FM_godkendt_0" <%=gkCHK0 %>><br />

                       <%
                    
                       
                       else 
                       
                 select case cint(oRec("godkendtstatus"))
                 case 2
                 gkTxt = "Afvist"
                 gkbgcol = "red"
                 case 1
                 gkTxt = "Godkendt"
                 gkbgcol = "green"
                 case else
                 gkTxt = ""
                 gkbgcol = "#000000"
                 end select
                 
                 Response.write "<span style=""color:"& gkbgcol &";"">"& gkTxt &"</span>" 
                       
                       %>


                    <%end if %>

                </td>
   </tr>
   
   <%
   lwedaynm = weekdayname(weekday(oRec("Tdato")))
   lwedate = oRec("Tdato")

   if cint(aty_real) = 1 then
   timertot = timertot + oRec("timer")
   timerDag = timerDag + oRec("timer")
   end if
   at = at + 1

   atDag = atDag + 1
   oRec.movenext
   wend
   oRec.close
   
   
   if at <> 0 then%>
   
  

   

   <tr>
        <td colspan=5 align=right><%=left(lwedaynm, 3) %>. d <%=formatdatetime(lwedate, 1) %> <b><%=formatnumber(timerDag, 2) %> t.</b>
        <br />
        <span style="color:#999999;">Fordelt på <%=atDag %> registreringer, heraf <%=eaDag %> Easyreg.</span></td>
        <td>&nbsp;</td>
    </tr>
   <tr>
        <td align=right colspan=5><br /><br />Total denne uge: <b><%=formatnumber(timertot, 2) %> t.</b><br />
        <span style="color:#999999;">Fordelt på <%=(at) %> registreringer, heraf <%=antalEasyReg %> Easyreg.</span></td>
        <td>&nbsp;</td>
   </tr>

   <% if cint(ugegodkendt) <> 1 AND media <> "print" then  %>
     <tr>
   <td colspan=6 align=right><br /><input type="submit" value="Opdater Status >>" style="font-size:9px;" />
   </td>
   <%end if %>
       
   </tr>

   <tr>
   <td colspan=6 align=right class=lille><br />Kun registreringer på aktivitets typer der tæller med <br />
    i det daglige forventede timeforbrug er med i totaler.</td>
       
   </tr>
   
   <%else %>
    <tr>
        <td colspan=6><br /><br />Ingen registreringer i valgte uge.</td>
    </tr>
   <%end if %>
   </table>
   </form>
   
   
   </td>

   <%if media <> "print" then %>
   <td valign=top style="width:220px;"><br />
   <b>Normtimer / Real timer denne uge:</b><br />
   
   <%call normtimerPer(medarbid, stdatoSQL, 6) %>
   
   <div style="position:relative; left:0px; top:5px; width:200px; border:0px #CCCCCC solid;">
   <table width=100% cellspacing=0 cellpadding=0 border=0>
   <tr>
        <td style="height:20px; border-bottom:1px #999999 solid;">Timer</td>
        <td style="border-bottom:1px #cccccc solid;"">
        <div style="position:absolute; left:30px; z-index:500000000; top:0px; width:20px; border-right:1px #cccccc solid;">
        <table cellpadding=0 cellspacing=0 border=0 width=100%>
        <tr><td align=right style="padding:2px;" class=lille>2</td></tr>
        </table>
        </div>
        <div style="position:absolute; left:50px; z-index:500000000; top:0px; width:20px; border-right:1px #cccccc solid;">
        <table cellpadding=0 cellspacing=0 border=0 width=100%>
        <tr><td align=right style="padding:2px;" class=lille>4</td></tr>
        </table>
        </div>
        <div style="position:absolute; left:70px; z-index:500000000; top:0px; width:34px; border-right:1px #cccccc solid;">
        <table cellpadding=0 cellspacing=0 border=0 width=100%>
        <tr><td align=right style="padding:2px;" class=lille>7,4</td></tr>
        </table>
        </div>
        <div style="position:absolute; left:104px; z-index:500000000; top:0px; width:26px; border-right:1px #cccccc solid;">
        <table cellpadding=0 cellspacing=0 border=0 width=100%>
        <tr><td align=right style="padding:2px;" class=lille>10</td></tr>
        </table>
        </div>
        <div style="position:absolute; left:130px; z-index:500000000; top:0px; width:20px; border-right:1px #cccccc solid;">
        <table cellpadding=0 cellspacing=0 border=0 width=100%>
        <tr><td align=right style="padding:2px;" class=lille>12</td></tr>
        </table>
        </div>
        </td>
   </tr>
  <%tdheight = 31 %>

  <%
  dtMan =  dateAdd("d", 0, stdatoSQL)
  dtTir =  dateAdd("d", 1, stdatoSQL)
  dtOns =  dateAdd("d", 2, stdatoSQL)
  dtTor =  dateAdd("d", 3, stdatoSQL)
  dtFre =  dateAdd("d", 4, stdatoSQL)
  dtLor =  dateAdd("d", 5, stdatoSQL)
  dtSon =  dateAdd("d", 6, stdatoSQL)
  %>

   <tr>
       
         <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid; " class=lille>Man <br /> <%=day(dtMan)&"."& month(dtMan) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, stdatoSQL, nTimMan, 30, 20) %>&nbsp;</td>
         </tr>
         <tr>
         <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid; " class=lille>Tir <br /> <%=day(dtTir)&"."& month(dtTir) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, year(dateadd("d", 1, stdatoSQL)) &"-"& month(dateadd("d", 1, stdatoSQL)) &"-"& day(dateadd("d", 1, stdatoSQL)), nTimTir, 30, 51) %>&nbsp;</td>
         </tr>
         <tr>
         <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid;" class=lille>Ons <br /> <%=day(dtOns)&"."& month(dtOns) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, year(dateadd("d", 2, stdatoSQL)) &"-"& month(dateadd("d", 2, stdatoSQL)) &"-"& day(dateadd("d", 2, stdatoSQL)), nTimOns, 30, 82) %>&nbsp;</td>
        </tr>
         <tr>
        <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid;" class=lille>Tor <br /> <%=day(dtTor)&"."& month(dtTor) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, year(dateadd("d", 3, stdatoSQL)) &"-"& month(dateadd("d", 3, stdatoSQL)) &"-"& day(dateadd("d", 3, stdatoSQL)), nTimTor, 30, 113) %>&nbsp;</td>
         </tr>
         <tr>
         <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid;" class=lille>Fre <br /> <%=day(dtFre)&"."& month(dtFre) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, year(dateadd("d", 4, stdatoSQL)) &"-"& month(dateadd("d", 4, stdatoSQL)) &"-"& day(dateadd("d", 4, stdatoSQL)), nTimFre, 30, 144) %>&nbsp;</td>
         </tr>
         <tr>
         <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid;" class=lille>Lør <br /> <%=day(dtLor)&"."& month(dtLor) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, year(dateadd("d", 5, stdatoSQL)) &"-"& month(dateadd("d", 5, stdatoSQL)) &"-"& day(dateadd("d", 5, stdatoSQL)), nTimLor, 30, 175) %>&nbsp;</td>
        </tr>
         <tr>
        <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid;" class=lille>Søn <br /> <%=day(dtSon)&"."& month(dtSon) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, sldatoSQL, nTimSon, 30, 206) %>&nbsp;</td>
        
        
   </tr>
</table>
   
   </div>
   
   <!-- hvis level = 1 OR teamleder -->
   <%if level <=2 OR level = 6 then %>

       <br /><br />
       <!--
       <div style="border:1px #CCCCCC solid; padding:2px; width:180px;">
       -->

              


      
              
               
               <%if cint(showAfsuge) = 0 then
               
               
                if cint(ugegodkendt) <> 1 then 
                
                              if cint(ugegodkendt) = 2 then

                             call meStamdata(ugegodkendtaf)%>
                        <div style="color:#FFFFF; font-size:11px; background-color:#FF6666; padding:5px;"><b>Ugeseddel er afvist!</b><br />
                        <span style="font-size:9px; line-height:12px; color:#FFFFFF;"><i><%=ugegodkendtdt %> af <%=meNavn %></i></span>
                        <%if len(trim(ugegodkendtTxt)) <> 0 then %>
                        <br />
                        <span style="font-size:9px; line-height:12px; color:#000000;"><i><%=left(ugegodkendtTxt, 200) %></i></span>
                        <%end if %>
                        </div>

                        <%end if %>

                           <form action="<%=fmLink%>&func=godkendugeseddel" method="post">
                            <table width=90% cellpadding=0 cellspacing=0 border=0>
                            <tr><td class=lille><br />

                           <span style="font-size:11px;"><b>Godkend ugeseddel</b></span><br />
                           Når en ugeseddel godkendes, godkendes alle ugens registreringer automatisk.
                
                           <input id="Submit2" type="submit" value="Godkend ugeseddel >>" style="font-size:9px; width:120px;" />
             
                           </td></tr></table>
                           </form>

                <%else 
                        
                        call meStamdata(ugegodkendtaf)%>
                        <div style="color:green; font-size:11px; background-color:#DCF5BD; padding:5px;"><b>Ugeseddel er godkendt!</b><br />
                        <span style="font-size:9px; line-height:12px; color:#999999;"><i><%=ugegodkendtdt %> af <%=meNavn %></i></span></div>
                
                <%end if %>

               <%else %>
                
                 <form action="<%=fmLink%>&func=adviserugeafslutning" method="post">
                <table width=90% cellpadding=0 cellspacing=0 border=0>
                <tr><td class=lille>

                  <span style="font-size:11px;"><b>Godkend ugeseddel</b></span><br />
                 Ugeseddel kan IKKE godkendes før den er afsluttet af medarbejder.<br />
                
            

                <%if len(trim(request("showadviseringmsg"))) <> 0 then %>
                <br />
                  <div style="color:#000000; font-size:11px; background-color:#DCF5BD; padding:5px;"><b>Besked afsendt!</b></div>
                <%else %>
                  <br />
                  <b>Send email</b> med besked om at uge mangler af blive aflsuttet<br />
                <input id="Submit1" type="submit" value="Send besked >>" style="font-size:9px; width:120px;" />
                <%end if %>
           
             
               </td></tr></table>
                </form>
               <%end if 
               
               
               if cint(ugegodkendt) <> 2 AND cint(showAfsuge) = 0 then%>

               <form action="<%=fmLink%>&func=afvisugeseddel" method="post">
               <table width=90% cellpadding=0 cellspacing=0 border=0>
               <tr><td class=lille>
               <br />
               <span style="font-size:11px;"><b>Afvis ugeseddel</b></span><br />
               Begrundelse:<br />
               <input type="text" value="" name="FM_afvis_grund" style="width:150px; height:40px; font-size:9px;" />
                <input id="Submit3" type="submit" value="Afvis ugeseddel >>" style="font-size:9px; width:120px;" /><br />
                 Afsender email til medarbejer om at ugeseddel er afvist, og åbner evt. allerede godkendt ugeseddel op igen.
                  </td></tr></table>
                 
               </form>

               <%end if %>


   
             

   
     <!--</div>-->
   
   <%end if 'level %> 
    
    
   </td>
   <%else %>
   <td>&nbsp;</td>
   <%end if %>
   
   </tr>
    </table>
    
   
    

    <%end function
    
    
    
    function normrealgraf(usemrn, realDato, nTimDag, lft, tp)

if nTimDag > 0 then

nTimDagPX = formatnumber(nTimDag, 0)
%>
   
   <span id=nt style="position:absolute; left:<%=lft%>px; top:<%=tp+1%>px; width:<%=10*nTimDagPX%>px; z-index:500; height:13px; background-color:#D6Dff5; border-right:1px #8caae6 solid;">
   <table width=100% cellpadding=0 cellspacing=0 border=0><tr>
   <td align=right style="padding-right:5px;" class=lille><%
    if nTimDag <> 0 then
    nTimDag = nTimDag
    else
	nTimDag = 0
	end if
   
    %>
  
   <%=nTimDag %>
   
   </td></tr></table></span>
<%end if %>   
   
  
   
   
   
        <%strSQL = "SELECT sum(timer) AS timer_indtastet FROM timer WHERE Tmnr = "& usemrn &" AND Tdato='"& realDato &"' AND ("& aty_sql_realhours &")"
		intDayHours = 0
		
		'Response.Write strSQL
		'Response.flush
		oRec.open strSQL, oConn, 3
		
		
		
		if not oRec.EOF then
		intDayHours = oRec("timer_indtastet")
	    end if
	    oRec.close 
	    
	    if intDayHours <> 0 then
	    intDayHours = intDayHours
	    else
	    intDayHours = 0
	    end if
	    
	    %>
   
   <%if intDayHours > 0 then 
   if intDayHours > 18 then
   intDayHoursDivval = 18
   else
   intDayHoursDivval = intDayHours
   end if
   %>
   <span id=rt style="position:absolute; padding:0px 0px 0px 2px; left:<%=lft%>px; top:<%=tp+15%>px; width:<%=10*formatnumber(intDayHoursDivval,0)%>px; z-index:1000; height:14px; background-color:yellowgreen; border-right:1px #999999 solid;">
   <table callpadding=0 cellspacing=0 border=0 width=100%></tr><td class=lille style="padding-right:5px;" align=right><%=formatnumber(intDayHours, 2) %></td></table>
    
   </span>
   <%end if %>
   
 <%  
 end function 
    
 
 
public  fTimer, ifTimer, km, sTimer, flexTimer,  sundTimer, pausTimer 
public	fpTimer, feriePLTimer, fefriTimer, fefriTimerBr, ferieAFTimer, ferieAFPerTimer, ferieAFPerTimertimer 
public	ferieOptjtimer, ferieUdbTimer, fefriTimerUdb, sygTimer, BarnsygTimer
public	afspTimer, afspTimerTim, afspTimerBr, afspTimerUdb, afspTimerBrPer
public	afspTimerOUdb, natTimer, weekendTimer, weekendnattimer 
public	aftenTimer, aftenweekendTimer, adhocTimer, stkAntal
public 	lageTimer, e1Timer, e2Timer
public dagTimer, fefriplTimer, ferieAFulonTimer, fefriTimerPerBr
public ferieaarSlut, ferieaarStart, afspTimerTimPer, afspTimerPer, sygaarSt
public sygDage, sygDagePer, sygTimerPer, barnSygTimerPer, barnSygDagePer, barnSygDage
public startdatoFerieLabel, slutdatoFerieLabel, barsel, ferieAFPerstDato, feriefriAFPerstDato, barselPerstDato, barnsygPerstDato, sygPerstDato, omsorgPerstDato, afspadPerstDato  
public  omsorg, senior, divfritimer, fefriTimerPerBrTimer, ferieFriaarStart, ferieFriaarSlut, startdatoFerieFriLabel, slutdatoFerieFriLabel
public fefriTimerUdbPer, fefriTimerUdbPerTimer
public ferieAFulonPer, ferieAFulonPerTimer, ferieUdbPer, ferieUdbPerTimer, ferieOptjUlontimer, ferieOptjOverforttimer 
	    

        redim fTimer(1200), ifTimer(1200), km(1200), sTimer(1200), flexTimer(1200),  sundTimer(1200), pausTimer(1200) 
        redim fpTimer(1200), feriePLTimer(1200), fefriTimer(1200), fefriTimerBr(1200), afspTimerBrPer(1200), ferieAFTimer(1200) 
        redim ferieOptjtimer(1200), ferieUdbTimer(1200), fefriTimerUdb(1200), sygTimer(1200), BarnsygTimer(1200), fefriTimerUdbPer(1200), fefriTimerUdbPerTimer(1200)
        redim afspTimer(1200), afspTimerTim(1200), afspTimerBr(1200), afspTimerUdb(1200), ferieAFPerTimer(1200), ferieAFPerTimertimer(1200) 
        redim afspTimerOUdb(1200), natTimer(1200), weekendTimer(1200), weekendnattimer(1200)
        redim aftenTimer(1200), aftenweekendTimer(1200), adhocTimer(1200), stkAntal(1200)
        redim lageTimer(1200), e1Timer(1200), e2Timer(1200)
        redim dagTimer(1200), fefriplTimer(1200), ferieAFulonTimer(1200), fefriTimerPerBr(1200)
        redim afspTimerTimPer(1200), afspTimerPer(1200), sygDage(1200), sygTimerPer(1200), sygDagePer(1200)
        redim barnSygTimerPer(1200), barnSygDagePer(1200), barnSygDage(1200), barsel(1200)
        redim ferieAFPerstDato(1200), feriefriAFPerstDato(1200), barselPerstDato(1200), barnsygPerstDato(1200), sygPerstDato(1200), omsorgPerstDato(1200), afspadPerstDato(1200) 
        redim omsorg(1200), senior(1200), divfritimer(1200), fefriTimerPerBrTimer(1200)
        redim ferieAFulonPer(1200), ferieAFulonPerTimer(1200), ferieUdbPer(1200), ferieUdbPerTimer(1200)
        redim   ferieOptjUlontimer(1200), ferieOptjOverforttimer(1200)

	     
function fordelpaaaktType(intMid, startDato, slutDato, visning, akttype_sel, x)
         
         'Response.Write  "startDato, slutDato, visning: " & startDato &","& slutDato &","& visning 
         
	     fTimer(x) = 0
	     ifTimer(x) = 0
	     
	     km(x) = 0
	     
	     sTimer(x) = 0
	     
	     flexTimer(x) = 0
	     sundTimer(x) = 0
	     pausTimer(x) = 0
	     fpTimer(x) = 0
	     
	               
	     fefriTimer(x) = 0
	     fefriplTimer(x) = 0
	     
	     fefriTimerBr(x) = 0
	     fefriTimerPerBr(x) = 0 'dage
	     fefriTimerPerBrTimer(x) = 0

	     ferieAFTimer(x) = 0
	     ferieAFPerTimer(x) = 0 'dage
	     ferieAFPerTimertimer(x) = 0

	     ferieAFulonTimer(x) = 0
	     ferieOptjtimer(x) = 0
	     ferieOptjUlontimer(x) = 0
         ferieOptjOverforttimer(x) = 0


	     ferieUdbTimer(x) = 0

         ferieAFulonPer(x) = 0
         ferieAFulonPerTimer(x) = 0
         ferieUdbPer(x) = 0 
         ferieUdbPerTimer(x) = 0
	     
	     fefriTimerUdb(x) = 0
         fefriTimerUdbPer(x) = 0
         fefriTimerUdbPerTimer(x) = 0
	     
	     sygTimer(x) = 0
	     sygDage(x) = 0
	     sygTimerPer(x) = 0
	     sygDagePer(x) = 0
	     
	     BarnsygTimer(x) = 0
	     barnSygTimerPer(x) = 0
         barnSygDage(x) = 0
         barnSygDagePer(x) = 0
	     
	     afspTimer(x) = 0
	     afspTimerTim(x) = 0
	     
	     afspTimerTimPer(x) = 0
	     afspTimerPer(x) = 0
	     
	     afspTimerBr(x) = 0
	     afspTimerBrPer(x) = 0
	     afspTimerUdb(x) = 0
	     afspTimerOUdb(x) = 0
	     
	     dagTimer(x) = 0
	     
	     natTimer(x) = 0
	     
	     weekendTimer(x) = 0
	     
	     weekendnattimer(x) = 0
	     
	     aftenTimer(x) = 0
	     
	     aftenweekendTimer(x) = 0
	     
	     adhocTimer(x) = 0
	     
	     stkAntal(x) = 0
	     
	     lageTimer(x) = 0
	    
	     e1Timer(x) = 0
	     e2Timer(x) = 0

         barsel(x) = 0


         omsorg(x) = 0
         senior(x) = 0
         divfritimer(x) = 0
         
        per13wrt = 0
        per14wrt = 0
        per20wrt = 0
        per21wrt = 0
        per30wrt = 0
        per31wrt = 0
        
        afstemnul(x) = 0 
        
        aktiveTyper = akttype_sel

        ferieAFPerstDato(x) = "2001/12/31"
        feriefriAFPerstDato(x) = "2001/12/31"
        barselPerstDato(x) = "2001/12/31"
        barnsygPerstDato(x) = "2001/12/31"
        sygPerstDato(x) = "2001/12/31"
        omsorgPerstDato(x) = "2001/12/31"
        afspadPerstDato(x) = "2001/12/31" 
            
            
         '*** Ferie pl. kun fra DagsDato ***'
	    
	    startdatoSQLNow = year(now) &"/"& month(now) & "/"& day(now)
	    startdatoSQL = startdato
	    
	  
	        '**Indstiller ferieperiode til ferie år (planlagt felter > dd) '***
	        if cdate(slutdato) >= cdate("1-5-"& datepart("yyyy", slutdato, 2, 2)) AND cdate(slutdato) <= cdate("31-12-"& datepart("yyyy", slutdato, 2, 2)) then
	        'slutdatoFerie = datepart("yyyy", (dateadd("yyyy", "1",slutdato)), 2, 2) & "-04-30"
	        startdatoFerie = datepart("yyyy", slutdato, 2, 2) & "-05-01"
            slutdatoFerie = slutdato
            slutdatoFeriePl = datepart("yyyy", (dateadd("yyyy", "1",slutdato)), 2, 2) & "-04-30"

	        else
	        'slutdatoFerie = datepart("yyyy", slutdato, 2, 2) & "-04-30"
	        slutdatoFerie = slutdato
            startdatoFerie = datepart("yyyy", (dateadd("yyyy", "-1",slutdato)), 2, 2) & "-05-01"

           
            slutdatoFeriePl = datepart("yyyy", slutdato, 2, 2) & "-04-30"
	        end if

	        
	        ferieaarStart = datepart("yyyy", startdatoFerie, 2, 2)
	        ferieaarSlut = datepart("yyyy", slutdatoFerie, 2, 2)
	        
	        slutdatoFerieLabel = datepart("d",slutdatoFerie,2,2) &"/"& datepart("m",slutdatoFerie,2,2) &"/"& datepart("yyyy",slutdatoFerie,2,2)
            startdatoFerieLabel = datepart("d",startdatoFerie,2,2) &"/"& datepart("m",startdatoFerie,2,2) &"/"& datepart("yyyy",startdatoFerie,2,2)

            
            '****************************************************
            '**** Følger feriefridag kalender år eller ferie år REGEL: Kalender år
            '****************************************************
            select case lcase(lto)
            case "mi", "intranet - local" '** kalender år

            if visning <> 4 then '' saldo på timereg.
            startdatoFerieFri = year(slutdato)&"/1/1"
            slutdatoFerieFri = year(slutdato)&"/"&month(slutdato)&"/"&day(slutdato)
            else
            startdatoFerieFri = year(startdato)&"/1/1"
            slutdatoFerieFri = year(startdato)&"/12/31" '&month(startdato)&"/"&day(startdato)
            end if

            ferieFriaarStart = datepart("yyyy", startdatoFerieFri, 2, 2)
	        ferieFriaarSlut = datepart("yyyy", slutdatoFerieFri, 2, 2)
	        
	        slutdatoFerieFriLabel = datepart("d",slutdatoFerieFri,2,2) &"/"& datepart("m",slutdatoFerieFri,2,2) &"/"& datepart("yyyy",slutdatoFerieFri,2,2)
            startdatoFerieFriLabel = datepart("d",startdatoFerieFri,2,2) &"/"& datepart("m",startdatoFerieFri,2,2) &"/"& datepart("yyyy",startdatoFerieFri,2,2)


            case else ''** ferieår
            startdatoFerieFri = startdatoFerie
            slutdatoFerieFri = slutdatoFerie


            ferieFriaarStart = datepart("yyyy", startdatoFerieFri, 2, 2)
	        ferieFriaarSlut = datepart("yyyy", slutdatoFerieFri, 2, 2)
	        
	        slutdatoFerieFriLabel = datepart("d",slutdatoFerieFri,2,2) &"/"& datepart("m",slutdatoFerieFri,2,2) &"/"& datepart("yyyy",slutdatoFerieFri,2,2)
            startdatoFerieFriLabel = datepart("d",startdatoFerieFri,2,2) &"/"& datepart("m",startdatoFerieFri,2,2) &"/"& datepart("yyyy",startdatoFerieFri,2,2)

            end select

            'Response.write "Her:" & ferieFriaarStart & "<br>"

            
	        
          

	        
	        '*** hvis ansat dato er senere end licens st. dato **'
	        if cDate(meAnsatDato) > cDate(lisStDato) then
	        lisStDatoSQL = year(meAnsatDato) &"/"& month(meAnsatDato) &"/"& day(meAnsatDato) 'meAnsatDato
	        else
	        lisStDatoSQL = year(lisStDato) &"/"& month(lisStDato) &"/"& day(lisStDato)
	        end if
	        
	        if visning = 1 OR visning = 5 then '** MD --> Dato
	        lisStDatoSQL = lisStDatoSQL 'year(lisStDato) &"/"& month(lisStDato) &"/"& day(lisStDato)
	        else
	        lisStDatoSQL = year(startdato) &"/"& month(startdato) &"/"& day(startdato) 'startdato
	        slutdatoFerie = startdato
	        slutdatoFerie = slutdato
	        end if

            

              '*** Syg ***'
            sygaarSt = year(startdato) &"/1/1" 'year(slutdato) &"/1/1"

            if cint(visning) = 6 OR cint(visning) = 7 OR visning = 77 then
            sygaarSt = startdato
            end if
	        
	        
	        'Response.Write "<hr>" & aktiveTyper & "<hr>"
	        
	        alleaktivetyperSQL = aktiveTyper
	       
	       '** fjerner sumtyper **'
	       alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #-1#", "")
	       alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #-2#", "")
	       alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #-3#", "")
	       alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #-4#", "")
	       alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #-5#", "")
	       
	       
	       '** Trimer og sætter dato interval på hver enekelt type. Ferie skal altid være i ferie år **' 
	        alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #11#", " OR (af.tfaktim = 11 AND af.tdato BETWEEN '"& startdatoSQLNow &"' AND '"& slutdatoFeriePl &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #12#", " OR (af.tfaktim = 12 AND af.tdato BETWEEN '"& startdatoFerieFri &"' AND '"& slutdatoFerieFri &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #13#", " OR (af.tfaktim = 13 AND af.tdato BETWEEN '"& startdatoFerieFri &"' AND '"& slutdatoFerieFri &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #14#", " OR (af.tfaktim = 14 AND af.tdato BETWEEN '"& startdatoFerie &"' AND '"& slutdatoFerie &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #15#", " OR (af.tfaktim = 15 AND af.tdato BETWEEN '"& startdatoFerie &"' AND '"& slutdatoFerie &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #16#", " OR (af.tfaktim = 16 AND af.tdato BETWEEN '"& startdatoFerie &"' AND '"& slutdatoFerie &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #17#", " OR (af.tfaktim = 17 AND af.tdato BETWEEN '"& startdatoFerieFri &"' AND '"& slutdatoFerieFri &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #18#", " OR (af.tfaktim = 18 AND af.tdato BETWEEN '"& startdatoFerieFri &"' AND '"& slutdatoFerieFri &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #19#", " OR (af.tfaktim = 19 AND af.tdato BETWEEN '"& startdatoFerie &"' AND '"& slutdatoFerie &"')")
    	   
           alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #111#", " OR (af.tfaktim = 111 AND af.tdato BETWEEN '"& startdatoFerie &"' AND '"& slutdatoFerie &"')")
           alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #112#", " OR (af.tfaktim = 112 AND af.tdato BETWEEN '"& startdatoFerie &"' AND '"& slutdatoFerie &"')")

    	    '** Sætter periode på andre felter til valgt periode **'
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #1#", " OR (af.tfaktim = 1 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #2#", " OR (af.tfaktim = 2 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #5#", " OR (af.tfaktim = 5 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #6#", " OR (af.tfaktim = 6 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #7#", " OR (af.tfaktim = 7 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #8#", " OR (af.tfaktim = 8 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #9#", " OR (af.tfaktim = 9 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #10#", " OR (af.tfaktim = 10 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    
    	    '*** Sygdom / Barn Syg / Barsel / Omsorg mm **'
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #20#", " OR (af.tfaktim = 20 AND af.tdato BETWEEN '"& sygaarSt &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #21#", " OR (af.tfaktim = 21 AND af.tdato BETWEEN '"& sygaarSt &"' AND '"& slutdato &"')")
            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #22#", " OR (af.tfaktim = 22 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #23#", " OR (af.tfaktim = 23 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #24#", " OR (af.tfaktim = 24 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #25#", " OR (af.tfaktim = 25 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    
    	    '** Overarb / Afspad **'
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #30#", " OR (af.tfaktim = 30 AND af.tdato BETWEEN '"& lisStDatoSQL &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #31#", " OR (af.tfaktim = 31 AND af.tdato BETWEEN '"& lisStDatoSQL &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #32#", " OR (af.tfaktim = 32 AND af.tdato BETWEEN '"& lisStDatoSQL &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #33#", " OR (af.tfaktim = 33 AND af.tdato BETWEEN '"& lisStDatoSQL &"' AND '"& slutdato &"')")
    	    
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #50#", " OR (af.tfaktim = 50 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #51#", " OR (af.tfaktim = 51 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #52#", " OR (af.tfaktim = 52 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #53#", " OR (af.tfaktim = 53 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #54#", " OR (af.tfaktim = 54 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #55#", " OR (af.tfaktim = 55 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #60#", " OR (af.tfaktim = 60 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #61#", " OR (af.tfaktim = 61 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #81#", " OR (af.tfaktim = 81 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #90#", " OR (af.tfaktim = 90 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #91#", " OR (af.tfaktim = 91 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, "#-99# OR ", "")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, "#-99#,", "af.tfaktim = -99")
    	    
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ",", "")
    	    
    	    'alleaktivetyperSQL = replace(alleaktivetyperSQL, "#-99#", "")
	        'alleaktivetyperSQL_len = len(alleaktivetyperSQL)
	        'alleaktivetyperSQL = left(alleaktivetyperSQL, alleaktivetyperSQL_len - 13) 
    	   
    	    
	        aktiveTyper = ""
	        'end if
	    
	    'end if
	    
	    
	    if trim(alleaktivetyperSQL) <> "#-99#" AND trim(alleaktivetyperSQL) <> "#-99#," then
	    alleaktivetyperSQL = alleaktivetyperSQL
	    else
	    alleaktivetyperSQL = "af.tfaktim = -99"
	    end if
	    
	    'Response.Write "<hr>" & alleaktivetyperSQL & "<hr>"
	    'Response.flush
	    
	    '**** Timefordeling på akt.typer *****'
	    strSQLtim = "SELECT tfaktim, sum(af.timer) AS sumtimer, "_
	    &" sum(af.timer * a.faktor) AS sumenheder, af.tdato,  "_
	    &" a.faktor FROM timer af "_
	    &" LEFT JOIN aktiviteter a ON (a.id = af.taktivitetid) "_
	    &" WHERE af.tmnr = "& intMid &" AND ("& alleaktivetyperSQL &") "_
	    &" GROUP BY af.tfaktim, af.tmnr, af.tdato ORDER BY tdato DESC"
	    
	    
	    'AND af.tdato BETWEEN '"& startdatoSQL &"' AND '"& slutdato &"'
	    'if visning <> 1 then
	    'alleaktivetyperSQL = ""
	    'akttype_sel = ""
	    'end if
	    
	    'af.tfaktim = "& akttype_sel_arr(t) &"
	    'if session("mid") = 1 then
        'Response.Write "<br>Visning: "& visning &"<br>"& strSQLtim &"<hr>"
	    'Response.flush
	    'end if

        'if session("mid") = 1 then
	    'Response.Write "x:"& meAnsatDato &" "& x &" visning:" & visning &" sql: <br>"& strSQLtim & "<br><br>"
	    'Response.flush
        'end if
	   
	    oRec6.open strSQLtim, oConn, 3
	    while not oRec6.EOF
	        
	        
	     select case oRec6("tfaktim") 'akttype_sel_arr(t)
	     case 1
	     fTimer(x) = fTimer(x) + oRec6("sumtimer")
	     case 2
	     ifTimer(x) = ifTimer(x) + oRec6("sumtimer")
	     case 5
	     km(x) = km(x) + oRec6("sumtimer")
	     case 6
	     sTimer(x) = sTimer(x) + oRec6("sumtimer")
	     case 7
	     flexTimer(x) = flexTimer(x) + oRec6("sumtimer")
	     case 8
	     sundTimer(x) = sundTimer(x) + oRec6("sumtimer")
	     case 9
	     pausTimer(x) = pausTimer(x) + oRec6("sumtimer")
	     case 10
	     fpTimer(x) = fpTimer(x) + oRec6("sumtimer")
	     
         case 11    '**** Ferie pl ***'
	     
	     call normtimerPer(intMid, oRec6("tdato"), 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

          feriePLTimer(x) = feriePLTimer(x) + oRec6("sumtimer") / ntimPerUse         
	               
	                
	     case 12
	     
         '*** nomrtid skal være staddanrd uge uanset helleidag mm. **'
         call normtimerPer(intMid, oRec6("tdato"), 6)
	     
         if ntimPer <> 0 then
         'ntimPerUse = ntimPer/antalDageMtimer
         normTimerGns5 = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)  / 5
         else
         normTimerGns5 = 1
         end if 

          fefriTimer(x) =  fefriTimer(x) + oRec6("sumtimer") / normTimerGns5

	     case 13
	    
         call normtimerPer(intMid, oRec6("tdato"), 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

          fefriTimerBr(x) =  fefriTimerBr(x) + oRec6("sumtimer") / ntimPerUse
	                
                    'Response.write "per13wrt: "& per13wrt

                    if cint(per13wrt) = 0 then 

	                      ''*** FerieFridage afholdt i per **'
	                    strSQLper = "SELECT SUM(timer) AS sumTimerPer, tdato FROM timer af WHERE af.tmnr = "& intMid &" AND af.tfaktim = 13 AND "_
	                    &" af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tfaktim, af.tdato ORDER BY tdato DESC"
	                    
                        'if session("mid") = 1 then
	                    'Response.Write "####"& strSQLper
	                    'Response.flush
                        'end if
	                
	                    oRec5.open strSQLper, oConn, 3
	                    while not oRec5.EOF 

                         call normtimerPer(intMid, oRec5("tdato"), 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                        fefriTimerPerBr(x) = fefriTimerPerBr(x) + oRec5("sumtimerPer") / ntimPerUse
                        fefriTimerPerBrTimer(x) = fefriTimerPerBrTimer(x) + (oRec5("sumtimerPer")/1)

                        '*** Første dage i peridode ***'
                        feriefriAFPerstDato(x) = oRec5("tdato")

	                    oRec5.movenext
	                    wend
	                    oRec5.close

                         per13wrt = 1

                    end if
	     
	     case 14

         call normtimerPer(intMid, oRec6("tdato"), 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         ferieAFTimer(x) = ferieAFTimer(x) + (oRec6("sumtimer") / ntimPerUse)
       
         
         'Response.Write oRec6("sumtimer") & " " & oRec6("tdato") & " " & ntimPerUse & "<br>"
        


                    if cint(per14wrt) = 0 then 

	                     ''*** Ferie afholdt i per **'
	                     strSQLper = "SELECT SUM(timer) AS sumTimerPer, tdato FROM timer af WHERE af.tmnr = "& intMid &" AND af.tfaktim = 14 AND "_
	                     &" af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tfaktim, af.tdato ORDER BY tdato DESC"
	                
	                    'Response.Write strSQLper
	                    'Response.flush
	                   
	                    oRec5.open strSQLper, oConn, 3
	                    while not oRec5.EOF 

                         call normtimerPer(intMid, oRec5("tdato"), 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                        ferieAFPerTimer(x) = ferieAFPerTimer(x) + (oRec5("sumtimerPer") / ntimPerUse)
                        ferieAFPerTimertimer(x) = ferieAFPerTimertimer(x) + (oRec5("sumtimerPer") / 1)

                        '*** Første dage i peridode ***'
                        ferieAFPerstDato(x) = oRec5("tdato") 
                        
                        oRec5.movenext
	                    wend
	                    oRec5.close

                         per14wrt = 1

                    end if
	                
	     
	     
	     case 15
	     
         '*** nomrtid skal være staddanrd uge, baseret på 5 arb. dage uanset helleidag mm. **'
         call normtimerPer(intMid, oRec6("tdato"), 6)
	     if ntimPer <> 0 then
         'ntimPerUse = ntimPer/antalDageMtimer
         normTimerGns5 = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)  / 5
         else
         normTimerGns5 = 1
         end if 

         ferieOptjtimer(x) = ferieOptjtimer(x) + oRec6("sumtimer") / normTimerGns5


         'Response.write "<br>"
         'Response.write " ferieOptjtimer(x)" & ferieOptjtimer(x)

         case 111

         '*** nomrtid skal være staddanrd uge, baseret på 5 arb. dage uanset helleidag mm. **'
         call normtimerPer(intMid, oRec6("tdato"), 6)
	     if ntimPer <> 0 then
         'ntimPerUse = ntimPer/antalDageMtimer
         normTimerGns5 = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)  / 5
         else
         normTimerGns5 = 1
         end if 

         ferieOptjOverforttimer(x) = ferieOptjOverforttimer(x) + oRec6("sumtimer") / normTimerGns5
         
         
         case 112

          '*** nomrtid skal være staddanrd uge, baseret på 5 arb. dage uanset helleidag mm. **'
         call normtimerPer(intMid, oRec6("tdato"), 6)
	     if ntimPer <> 0 then
         'ntimPerUse = ntimPer/antalDageMtimer
         normTimerGns5 = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)  / 5
         else
         normTimerGns5 = 1
         end if 

         ferieOptjUlontimer(x) = ferieOptjOverforttimer(x) + oRec6("sumtimer") / normTimerGns5

         
         
         case 16
	   
        
         call normtimerPer(intMid, oRec6("tdato"), 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         ferieUdbTimer(x) = ferieUdbTimer(x) + oRec6("sumtimer") / ntimPerUse

                    
                       if cint(per16wrt) = 0 then 

	                     ''*** Ferie udbetalt afholdt i per **'
	                     strSQLper = "SELECT SUM(timer) AS sumTimerPer, tdato FROM timer af WHERE af.tmnr = "& intMid &" AND af.tfaktim = 16 AND "_
	                     &" af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tfaktim, af.tdato ORDER BY tdato DESC"
	                
	                    'Response.Write strSQLper
	                    'Response.flush
	                   
	                    oRec5.open strSQLper, oConn, 3
	                    while not oRec5.EOF 

                         call normtimerPer(intMid, oRec5("tdato"), 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                        ferieUdbPer(x) = ferieUdbPer(x) + (oRec5("sumtimerPer") / ntimPerUse)
                        ferieUdbPerTimer(x) = ferieUdbPerTimer(x) + (oRec5("sumtimerPer") / 1)

                      
                        
                        oRec5.movenext
	                    wend
	                    oRec5.close

                         per16wrt = 1

                    end if
                     

	     case 17
	    
         
         call normtimerPer(intMid, oRec6("tdato"), 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         fefriTimerUdb(x) = fefriTimerUdb(x) + oRec6("sumtimer") / ntimPerUse


                    
                       if cint(per17wrt) = 0 then 

	                     ''*** FerieFridage udbetalt i per **'
	                     strSQLper = "SELECT SUM(timer) AS sumTimerPer, tdato FROM timer af WHERE af.tmnr = "& intMid &" AND af.tfaktim = 17 AND "_
	                     &" af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tfaktim, af.tdato ORDER BY tdato DESC"
	                
	                    'Response.Write strSQLper
	                    'Response.flush
	                   
	                    oRec5.open strSQLper, oConn, 3
	                    while not oRec5.EOF 

                         call normtimerPer(intMid, oRec5("tdato"), 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                        fefriTimerUdbPer(x) = fefriTimerUdbPer(x) + (oRec5("sumtimerPer") / ntimPerUse)
                        fefriTimerUdbPerTimer(x) = fefriTimerUdbPerTimer(x) + (oRec5("sumtimerPer") / 1)

                      
                        
                        oRec5.movenext
	                    wend
	                    oRec5.close

                         per17wrt = 1

                    end if



	     
	     case 18 '**** Ferie frdage pl ****'
	     

         call normtimerPer(intMid, oRec6("tdato"), 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         fefriplTimer(x) = fefriplTimer(x) + oRec6("sumtimer") / ntimPerUse

	     case 19           
	     
         
          call normtimerPer(intMid, oRec6("tdato"), 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         ferieAFulonTimer(x) = ferieAFulonTimer(x) + (oRec6("sumtimer") / ntimPerUse)

            
            
                        
                        if cint(per19wrt) = 0 then 

	                     ''*** Ferie U løn afholdt i per **'
	                     strSQLper = "SELECT SUM(timer) AS sumTimerPer, tdato FROM timer af WHERE af.tmnr = "& intMid &" AND af.tfaktim = 19 AND "_
	                     &" af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tfaktim, af.tdato ORDER BY tdato DESC"
	                
	                    'Response.Write strSQLper
	                    'Response.flush
	                   
	                    oRec5.open strSQLper, oConn, 3
	                    while not oRec5.EOF 

                         call normtimerPer(intMid, oRec5("tdato"), 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                        ferieAFulonPer(x) = ferieAFulonPer(x) + (oRec5("sumtimerPer") / ntimPerUse)
                        ferieAFulonPerTimer(x) = ferieAFulonPerTimer(x) + (oRec5("sumtimerPer") / 1)

                      
                        
                        oRec5.movenext
	                    wend
	                    oRec5.close

                         per19wrt = 1

                        end if
              
	     case 20
	    

         call normtimerPer(intMid, oRec6("tdato"), 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         sygTimer(x) = sygTimer(x) + oRec6("sumtimer")
	     sygDage(x) = sygDage(x) + (oRec6("sumtimer") / ntimPerUse)           
	                
	                
	                

                     if cint(per20wrt) = 0 then 

	                     ''*** Syg. i per **'
	                    strSQLper = "SELECT SUM(timer) AS sumTimerPer, tdato "_
	                    &" FROM timer af "_
	                    &" WHERE af.tmnr = "& intMid &" AND af.tfaktim = 20 AND "_
	                    &" af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tfaktim, af.tdato ORDER BY tdato DESC"
	                    

                        'if session("mid") = 1 then
                        'Response.write strSQLper & "<br><br>"
                        'end if

	                    oRec5.open strSQLper, oConn, 3
	                    while not oRec5.EOF 

                         call normtimerPer(intMid, oRec5("tdato"), 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                        sygTimerPer(x) = sygTimerPer(x) + oRec5("sumtimerPer")
                        sygDagePer(x) = sygDagePer(x) + (oRec5("sumtimerPer") / ntimPerUse)
                        sygPerstDato(x) = oRec5("tdato")
                        
                      

	                    oRec5.movenext
	                    wend
	                    oRec5.close

                         per20wrt = 1

                    end if
	     
	     case 21
	     
         
         call normtimerPer(intMid, oRec6("tdato"), 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         BarnsygTimer(x) = BarnsygTimer(x) + oRec6("sumtimer")
	     barnSygDage(x) = barnSygDage(x) + (oRec6("sumtimer") / ntimPerUse)
	                
	                
	               

                     if cint(per21wrt) = 0 then 

	                     ''*** Barnsyg. i per **'
	                    strSQLper = "SELECT SUM(timer) AS sumTimerPer, tdato "_
	                    &" FROM timer af "_
	                    &" WHERE af.tmnr = "& intMid &" AND af.tfaktim = 21 AND "_
	                    &" af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tfaktim, af.tdato ORDER BY tdato DESC"
	                
	                    oRec5.open strSQLper, oConn, 3
	                    while not oRec5.EOF 

                         call normtimerPer(intMid, oRec5("tdato"), 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                        barnSygTimerPer(x) = barnSygTimerPer(x) + oRec5("sumtimerPer") 
                        barnSygDagePer(x) = barnSygDagePer(x) + (oRec5("sumtimerPer") / ntimPerUse)

                         
                        barnsygPerstDato(x) = oRec5("tdato")

                        oRec5.movenext
	                    wend
	                    oRec5.close

                         per21wrt = 1

                    end if

	     
         case 22
         
	     
         
         call normtimerPer(intMid, oRec6("tdato"), 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         barsel(x) = barsel(x) + (oRec6("sumtimer") / ntimPerUse) 
         barselPerstDato(x) = oRec6("tdato")
     


         case 23 'omsorg(x)
         
          call normtimerPer(intMid, oRec6("tdato"), 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         omsorg(x) = omsorg(x) + (oRec6("sumtimer") / ntimPerUse) 
         omsorgPerstDato(x) = oRec6("tdato")
       
         
         case 24 'senior(x)

     
         senior(x) = senior(x) + oRec6("sumtimer")

         case 25 'divfritimer(x) 

         
         
	     call normtimerPer(intMid, oRec6("tdato"), 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         divfritimer(x) = divfritimer(x) + oRec6("sumtimer")

	     case 30
	     afspTimer(x) = afspTimer(x) + oRec6("sumenheder")
	     afspTimerTim(x) = afspTimerTim(x) + oRec6("sumtimer")
	                
	                
	               
	                

                    if cint(per30wrt) = 0 then 

	                     ''*** Overarb. i per **'
	                    strSQLper = "SELECT SUM(timer) AS sumTimerPer, sum(af.timer * a.faktor) AS sumenhederPer, tdato "_
	                    &" FROM timer af "_
	                    &" LEFT JOIN aktiviteter a ON (a.id = af.taktivitetid) "_
	                    &" WHERE af.tmnr = "& intMid &" AND af.tfaktim = 30 AND "_
	                    &" af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tfaktim, af.tdato"
	                
	                
	                    oRec5.open strSQLper, oConn, 3
	                    while not oRec5.EOF 

                         call normtimerPer(intMid, oRec5("tdato"), 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                       
                         afspTimerPer(x) = afspTimerPer(x) + oRec5("sumenhederPer")
	                    afspTimerTimPer(x) = afspTimerTimPer(x) + oRec5("sumTimerPer")

	                    oRec5.movenext
	                    wend
	                    oRec5.close

                         per30wrt = 1

                    end if

	     
	     case 31
	     afspTimerBr(x) = afspTimerBr(x) + oRec6("sumtimer")
	     
	                

                    if cint(per31wrt) = 0 then 

	                     ''*** Afspadsering i per **'
	                strSQLper = "SELECT SUM(timer) AS sumTimerPer, tdato "_
	                &" FROM timer af "_
	                &" WHERE af.tmnr = "& intMid &" AND af.tfaktim = 31 AND "_
	                &" af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tfaktim, af.tdato ORDER BY tdato DESC"
	                
	                
	                    oRec5.open strSQLper, oConn, 3
	                    while not oRec5.EOF 

                         call normtimerPer(intMid, oRec5("tdato"), 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                       
                        afspTimerBrPer(x) = afspTimerBrPer(x) + oRec5("sumTimerPer")
                        afspadPerstDato(x) = oRec5("tdato")

	                    
	                    oRec5.movenext
	                    wend
	                    oRec5.close

                         per31wrt = 1

                    end if

	     
	     case 32
	     afspTimerUdb(x) = afspTimerUdb(x) + oRec6("sumtimer")
	     case 33
	     afspTimerOUdb(x) =  afspTimerOUdb(x) + oRec6("sumtimer")
	     case 50
	     dagTimer(x) = dagTimer(x) + oRec6("sumtimer")
	     case 51
	     natTimer(x) = natTimer(x) + oRec6("sumtimer")
	     case 52
	     weekendTimer(x) = weekendTimer(x) + oRec6("sumtimer")
	     case 53
	     weekendnattimer(x) = weekendnattimer(x) + oRec6("sumtimer")
	     case 54
	     aftenTimer(x) = aftenTimer(x) + oRec6("sumtimer")
	     case 55
	     aftenweekendTimer(x) = aftenweekendTimer(x) + oRec6("sumtimer")
	     case 60
	     adhocTimer(x) = adhocTimer(x) + oRec6("sumtimer")
	     case 61
	     stkAntal(x) = stkAntal(x) + oRec6("sumtimer")
	     case 81
	     lageTimer(x) = lageTimer(x) + oRec6("sumtimer")
	     case 90
	     e1Timer(x) = e1Timer(x) + oRec6("sumtimer")
	     case 91
	     e2Timer(x) = e2Timer(x) + oRec6("sumtimer")
	     end select
	    
	    
	    
	    oRec6.movenext
	    wend
	    oRec6.close
	    
	    if media <> "export" then
	    Response.flush
        end if
	    

	
	         '* Afspadsering ***'
	         
	        if afspTimer(x) <> 0 then
	         afspTimer(x) = afspTimer(x)
	         afstemnul(x) = 235
	         else
	         afspTimer(x) = 0
	         end if
	         
	         if afspTimerTim(x) <> 0 then
	         afspTimerTim(x) = afspTimerTim(x)
	         afstemnul(x) = 234
	         else
	         afspTimerTim(x) = 0
	         end if
	        
	        
	        
	    
	    
	    '** Afspad. Brugt **'
	   
	    
	         if afspTimerBr(x) <> 0 then
	         afspTimerBr(x) = afspTimerBr(x)
	         afstemnul(x) = 233
	         else
	         afspTimerBr(x) = 0
	         end if
	    
	    
	    '*** Afspad Udbetalt **'
	        if afspTimerUdb(x) <> 0 then
	         afspTimerUdb(x) = afspTimerUdb(x)
	         afstemnul(x) = 232
	         else
	         afspTimerUdb(x) = 0
	         end if
	         
	         '*** Afspad Ønskes Udbetalt **'
	         if afspTimerOUdb(x) <> 0 then
	         afspTimerOUdb(x) = afspTimerOUdb(x)
	         afstemnul(x) = 231
	         else
	         afspTimerOUdb(x) = 0
	         end if
	         
	         
	       
	  
	    
	    '** Ferie fridage ***'
	    
	    '*** Optj **'
	     if fefriTimer(x) <> 0 then
	     fefriTimer(x) = fefriTimer(x)
	     afstemnul(x) = 230
	     else
	     fefriTimer(x) = 0
	     end if
	     



	     '*** PL ***'
	     if fefriplTimer(x) <> 0 then
	      fefriplTimer(x) =  fefriplTimer(x)
	      afstemnul(x) = 229
	     else
	      fefriplTimer(x) = 0
	     end if
	     
	    
	    
	    
	    '** Brugt **'
	    
	    if fefriTimerBr(x) <> 0 then
	     fefriTimerBr(x) = fefriTimerBr(x)
	     afstemnul(x) = 228
	     else
	    fefriTimerBr(x) = 0
	     end if
	    
	    
	    '** Udbetalt **'
	     if fefriTimerUdb(x) <> 0 then
	     fefriTimerUdb(x) = fefriTimerUdb(x)
	     afstemnul(x) = 227
	     else
	    fefriTimerUdb(x) = 0
	     end if
	    

         '** Udbetalt Periode **'
	     if fefriTimerUdbPer(x) <> 0 then
	     fefriTimerUdbPer(x) = fefriTimerUdbPer(x)
	     afstemnul(x) = 227
	     else
	    fefriTimerUdbPer(x) = 0
	     end if
	    
	   
	    
	    '** Ferie Planlagt **'
	    if feriePLTimer(x) <> 0 then
	    feriePLTimer(x) = feriePLTimer(x)
	    afstemnul(x) = 226
	    else
	    feriePLTimer(x) = 0
	    end if
	    
	    
	    
	    '** Ferie Afholdt **'
	    if ferieAFTimer(x) <> 0 then
	    ferieAFTimer(x) = ferieAFTimer(x)
	    afstemnul(x) = 225
	    else
	    ferieAFTimer(x) = 0
	    end if
	   
	     '** Ferie Afholdt ulon **'
	    if ferieAFulonTimer(x) <> 0 then
	    ferieAFulonTimer(x) = ferieAFulonTimer(x)
	    afstemnul(x) = 224
	    else
	    ferieAFulonTimer(x) = 0
	    end if
	  
	   
	    
	    '** Ferie Optjent **'
	    if  ferieOptjtimer(x) <> 0 then
	    ferieOptjtimer(x) =  ferieOptjtimer(x)
	    afstemnul(x) = 223
	    else
	    ferieOptjtimer(x) = 0
	    end if
	    
        '*** Ferie optjent overført **'
        if ferieOptjOverforttimer(x) <> 0 then
	    ferieOptjOverforttimer(x) = ferieOptjOverforttimer(x)
	    afstemnul(x) = 240
	    else
	    ferieOptjOverforttimer(x) = 0
	    end if

        '** SENESTE afstemnul(x) 241 2012.12.21 
        if ferieOptjUlontimer(x) <> 0 then
	    ferieOptjUlontimer(x) = ferieOptjUlontimer(x)
	    afstemnul(x) = 241
	    else
	    ferieOptjUlontimer(x) = 0
	    end if
         


	    
	    '** Ferie  Udb **'
	    if ferieUdbTimer(x) <> 0 then
	    ferieUdbTimer(x) = ferieUdbTimer(x)
	    afstemnul(x) = 222
	    else
	    ferieUdbTimer(x) = 0
	    end if
	    
	    
	      '*** Fakturerbaretimer #1 **'
	    if fTimer(x) <> 0 then
	    fTimer(x) = fTimer(x)
	    afstemnul(x) = 221
	    else
	    fTimer(x) = 0
	    end if
	      
	     
	     
	    '*** Ikke Fakturerbare timer #2 ***'
	    if ifTimer(x) <> 0 then
        ifTimer(x) = ifTimer(x)
        afstemnul(x) = 220
        else
        ifTimer(x) = 0
        end if
	    
	    '*** Salg & Newbizz timer***'
	    if sTimer(x) <> 0 then
        sTimer(x) = sTimer(x)
        afstemnul(x) = 219
        else
        sTimer(x) = 0
        end if
	    
	    
	    
	        
	    '*** Frokost ***'
	    if fpTimer(x) <> 0 then
	    fpTimer(x) = fpTimer(x)
	    afstemnul(x) = 218
	    else
	    fpTimer(x) = 0
	    end if
	       
	    
	    
	         
	    '*** Syg ****'
	    if sygTimer(x) <> 0 then
         sygTimer(x) = sygTimer(x)
         afstemnul(x) = 217
         else
         sygTimer(x) = 0
         end if
	   
         
          if sygDage(x) <> 0 then
         sygDage(x) = sygDage(x)
        else
         sygDage(x) = 0
         end if

          if sygDagePer(x) <> 0 then
         sygDagePer(x) = sygDagePer(x)
        else
         sygDagePer(x) = 0
         end if


         if barsel(x) <> 0 then
         barsel(x) = barsel(x)
         afstemnul(x) = 236
         else
         barsel(x) = 0
         end if

          if omsorg(x) <> 0 then
         omsorg(x) = omsorg(x)
         afstemnul(x) = 237
         else
         omsorg(x) = 0
         end if

          if senior(x) <> 0 then
         senior(x) = senior(x)
         afstemnul(x) = 238
         else
         senior(x) = 0
         end if

          if divfritimer(x) <> 0 then
         divfritimer(x) = divfritimer(x)
         afstemnul(x) = 239
         else
         divfritimer(x) = 0
         end if


	   
	   '*** Barn syg ***'
	    if BarnsygTimer(x) <> 0 then
         BarnsygTimer(x) = BarnsygTimer(x)
         afstemnul(x) = 216
         else
         BarnsygTimer(x) = 0
         end if
	    
	    if barnSygDage(x) <> 0 then
         barnSygDage(x) = barnSygDage(x)
        else
         barnSygDage(x) = 0
         end if

          if barnSygDagePer(x) <> 0 then
         barnSygDagePer(x) = barnSygDagePer(x)
        else
         barnSygDagePer(x) = 0
         end if
	    
	   
	    
	    
	    
	    '*** KM ***'
	    if km(x) <> 0 then
         km(x) = km(x)
         afstemnul(x) = 215
         else
         km(x) = 0
         end if
	         
	         
	      '*** sundheds timer ***'
	    if sundTimer(x) <> 0 then
	    sundTimer(x) = sundTimer(x)
	    afstemnul(x) = 214
	    else
	    sundTimer(x) = 0
	    end if
	    
	   
	    
	   
	    '*** Flex timer ***'
	    if flexTimer(x) <> 0 then
        flexTimer(x) = flexTimer(x)
        afstemnul(x) = 213
        else
        flexTimer(x) = 0
        end if

	  
	    
	   
	    
	    
	    '** Pause timer **'
	    if pausTimer(x) <> 0 then
	    pausTimer(x) = pausTimer(x)
	    afstemnul(x) = 212
	    else
	    pausTimer(x) = 0
	    end if
	    
	    
	     '*** Dag ***'
	   if dagTimer(x) <> 0 then
         dagTimer(x) = dagTimer(x)
         afstemnul(x) = 211
         else
         dagTimer(x) = 0
         end if
	    
	   '*** Nat ***'
	   if natTimer(x) <> 0 then
         natTimer(x) = natTimer(x)
         afstemnul(x) = 210
         else
         natTimer(x) = 0
         end if
	    
	    
	  
	    '*** Weekend ***'
	   if weekendTimer(x) <> 0 then
	         weekendTimer(x) = weekendTimer(x)
	         afstemnul(x) = 209
	         else
	         weekendTimer(x) = 0
	         end if
	    
	    
	    '*** Weekend Nat ***'
	   if weekendnattimer(x) <> 0 then
	         weekendnattimer(x) = weekendnattimer(x)
	         afstemnul(x) = 208
	         else
	         weekendnattimer(x) = 0
	         end if
	    
	    
	   
	 '*** Aften ***'
	 if aftenTimer(x) <> 0 then
     aftenTimer(x) = aftenTimer(x)
     afstemnul(x) = 207
     else
     aftenTimer(x) = 0
     end if
	    
	    
	    
	   '*** Aften Weekend ***'
	    if aftenweekendTimer(x) <> 0 then
         aftenweekendTimer(x) = aftenweekendTimer(x)
         afstemnul(x) = 206
         else
         aftenweekendTimer(x) = 0
         end if
	    
	    
	   
	    
	    
	    
	    '*** Adhoc ***'
	   
	         if adhocTimer(x) <> 0 then
	         adhocTimer(x) = adhocTimer(x)
	         afstemnul(x) = 205
	         else
	         adhocTimer(x) = 0
	         end if
	    
	  
	    
	    
	    '*** StkAntal ***'
	   if stkAntal(x) <> 0 then
	         stkAntal(x) = stkAntal(x)
	         afstemnul(x) = 204
	         else
	         stkAntal(x) = 0
	         end if
	    
	  
	     
	    
	    
	     '*** Læge Timer ***'
	    if lageTimer(x) <> 0 then
	     lageTimer(x) = lageTimer(x)
	     afstemnul(x) = 203
         else
         lageTimer(x) = 0
         end if
         
         
         '*** E1 Timer ***'
	     if e1Timer(x) <> 0 then
	     e1Timer(x) = e1Timer(x)
	     afstemnul(x) = 202
         else
         e1Timer(x) = 0
         end if
         
         '*** E2 Timer ***'
	     if e2Timer(x) <> 0 then
	     e2Timer(x) = e2Timer(x)
	     afstemnul(x) = 201
         else
         e2Timer(x) = 0
         end if
	     
	        
	        
	     
	    
            

end function   





    
%>