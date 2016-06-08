<%

public ferieBal
function feriebal_fn(ferieOptjtimer_X, ferieOptjOverforttimer_X, ferieOptjUlontimer_X, ferieAFTimer_X, ferieAFulonTimer_X, ferieUdbTimer_X)

     select case lto
     case "esn", "tec"
     ferieBal  = ((ferieOptjtimer_X + ferieOptjOverforttimer_X + ferieOptjUlontimer_X) - (ferieAFTimer_X + ferieAFulonTimer_X + ferieUdbTimer_X))
     case "cst", "wwf"
     ferieBal  = ((ferieOptjtimer_X + ferieOptjOverforttimer_X + ferieOptjUlontimer_X) - (ferieAFTimer_X + ferieAFulonTimer_X + ferieUdbTimer_X))
     case "xwwf"
	 ferieBal  = ((ferieOptjtimer_X + ferieOptjOverforttimer_X + ferieOptjUlontimer_X) - (ferieAFTimer_X + ferieUdbTimer_X - (ferieAFulonTimer_X)))
     case else
     ferieBal  = ((ferieOptjtimer_X + ferieOptjOverforttimer_X + ferieOptjUlontimer_X) - (ferieAFTimer_X + ferieUdbTimer_X))
     end select


end function



'********* Norm Real Graf ****************************
function normRealGrafWeekPage(medarbid, stdatoSQL)


    'Response.write "stdatoSQL "& stdatoSQL

    call normtimerPer(medarbid, stdatoSQL, 6, 0) %>
   
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
       
         <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid; " class=lille>Man <br /> <%=day(dtMan)&"."& month(dtMan) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, stdatoSQL, nTimMan, 30, 21) %>&nbsp;</td>
         </tr>
         <tr>
         <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid; " class=lille>Tir <br /> <%=day(dtTir)&"."& month(dtTir) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, year(dateadd("d", 1, stdatoSQL)) &"-"& month(dateadd("d", 1, stdatoSQL)) &"-"& day(dateadd("d", 1, stdatoSQL)), nTimTir, 30, 53) %>&nbsp;</td>
         </tr>
         <tr>
         <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid;" class=lille>Ons <br /> <%=day(dtOns)&"."& month(dtOns) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, year(dateadd("d", 2, stdatoSQL)) &"-"& month(dateadd("d", 2, stdatoSQL)) &"-"& day(dateadd("d", 2, stdatoSQL)), nTimOns, 30, 85) %>&nbsp;</td>
        </tr>
         <tr>
        <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid;" class=lille>Tor <br /> <%=day(dtTor)&"."& month(dtTor) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, year(dateadd("d", 3, stdatoSQL)) &"-"& month(dateadd("d", 3, stdatoSQL)) &"-"& day(dateadd("d", 3, stdatoSQL)), nTimTor, 30, 117) %>&nbsp;</td>
         </tr>
         <tr>
         <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid;" class=lille>Fre <br /> <%=day(dtFre)&"."& month(dtFre) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, year(dateadd("d", 4, stdatoSQL)) &"-"& month(dateadd("d", 4, stdatoSQL)) &"-"& day(dateadd("d", 4, stdatoSQL)), nTimFre, 30, 149) %>&nbsp;</td>
         </tr>
         <tr>
         <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid;" class=lille>Lør <br /> <%=day(dtLor)&"."& month(dtLor) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, year(dateadd("d", 5, stdatoSQL)) &"-"& month(dateadd("d", 5, stdatoSQL)) &"-"& day(dateadd("d", 5, stdatoSQL)), nTimLor, 30, 180) %>&nbsp;</td>
        </tr>
         <tr>
        <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid;" class=lille>Søn <br /> <%=day(dtSon)&"."& month(dtSon) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, year(dateadd("d", 6, stdatoSQL)) &"-"& month(dateadd("d", 6, stdatoSQL)) &"-"& day(dateadd("d", 6, stdatoSQL)), nTimSon, 30, 211) %>&nbsp;</td>
        
        
   </tr>
</table>
   
   </div>

<%end function


'******************************************************









public ntimPer, ntimMan, ntimTir, ntimOns, ntimTor, ntimFre, ntimLor, ntimSon, antalDageMtimer
public ntimManIgnHellig, ntimTirIgnHellig, ntimOnsIgnHellig, ntimTorIgnHellig, ntimFreIgnHellig, ntimLorIgnHellig, ntimSonIgnHellig, nTimerPerIgnHellig
	function normtimerPer(medid, stDato, interval, io)

    if medid <> 0 then
    medid = medid
    else 
    medid = 0
    end if

    'Response.write "VAR: m: "& medid &" , stdato: "& stDato &", intervla: "& interval & "<br>"
	
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
            nTimerPerIgnHellig = 0
			
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
			            
			            
			            'Response.Write "<u>cdate(datoCount) </u>"& cdate(datoCount) & " aftrædelse : "& cDate(mtyperIntvEndDato(0))  &"<br>"
			            
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
				
					
					'if session("mid") = 1 then
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

                    
          
                    '**Ignorer helligdage
                    nTimerPerIgnHellig = nTimerPerIgnHellig + ntimPerThis	

					
				    call helligdage(datoCount, 0, lto)
				        
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

            if len(nTimerPerIgnHellig) <> 0 then
			nTimerPerIgnHellig = nTimerPerIgnHellig
			else
			nTimerPerIgnHellig = 0
			end if

            
            if antalDageMtimer <> 0 then
            antalDageMtimer = antalDageMtimer
            else
            antalDageMtimer = 1
            end if
			
			
            'Response.write "her"
            'Response.end

            if io = 1 then 'reduktion for ferie afholdt, ulon og planlagt

            stDatoSQLkri = year(stDato) &"-"& month(stDato) &"-"& day(stDato) 
            slDatoSQLkri = dateAdd("d", interval, stDatoSQLkri)
            slDatoSQLkri = year(slDatoSQLkri) &"-"& month(slDatoSQLkri) &"-"& day(slDatoSQLkri) 

             strSQLfeau = "SELECT sum(timer) AS sumtimer, tdato, month(tdato) AS month, tfaktim FROM timer WHERE tmnr = " & medid  & ""_
	         &" AND (tfaktim = 11 OR tfaktim = 13 OR tfaktim = 14 OR tfaktim = 18 OR tfaktim = 19 OR tfaktim = 22 OR tfaktim = 23 OR tfaktim = 24 OR tfaktim = 25 OR tfaktim = 115) "_
             &" AND tdato BETWEEN '"& stDatoSQLkri &"' AND '"& slDatoSQLkri &"' GROUP BY tmnr ORDER BY tdato"

            'response.write strSQLfeau & "<br>"

            oRec6.open strSQLfeau, oConn, 3
	        if not oRec6.EOF then
        
                'response.write "Sumtimer: "& oRec6("sumtimer") &"<br>"

                ntimPer = ntimPer - oRec6("sumtimer")
          
            oRec6.movenext
            end if
            oRec6.close          

            end if        


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







    
    





'**********************************************************    

dim thisWdt
redim thisWdt(7)
function normrealgraf(usemrn, realDato, nTimDag, lft, tp)



'************** Norm timer ***************
if nTimDag > 0 then

nTimDagPX = formatnumber(nTimDag, 0)
%>
   
   <span id=nt style="position:absolute; left:<%=lft%>px; top:<%=tp+1%>px; width:<%=10*nTimDagPX%>px; z-index:500; height:13px; background-color:#CCCCCC;">
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
<%end if 
    
    
    
    
    '**** Real timer ****************
        strSQL = "SELECT sum(timer) AS timer_indtastet FROM timer WHERE Tmnr = "& usemrn &" AND Tdato ='"& realDato &"' AND ("& aty_sql_realhours &")"
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

    
    thisWeekDay = datePart("w", realDato, 2,2)

   thisWdt(thisWeekDay) = 10*formatnumber(intDayHoursDivval,0)
   %>
   <span id="rt" style="position:absolute; padding:0px 0px 0px 2px; left:<%=lft%>px; top:<%=tp+15%>px; width:<%=thisWdt(thisWeekDay)%>px; z-index:1000; height:14px; background-color:yellowgreen;">
   <table callpadding=0 cellspacing=0 border=0 width=100%></tr><td class=lille style="padding-right:5px;" align=right><%=formatnumber(intDayHours, 2) %></td></table>
    
   </span>
   <%end if





    '*********** Timer_temp_timer ********'


        strSQL = "SELECT sum(timer) AS timer_temp FROM timer_import_temp WHERE medarbejderid = "& usemrn &" AND Tdato ='"& realDato &"' AND overfort = 0"
		intDayHours = 0
		
		'Response.Write strSQL
		'Response.flush
		oRec.open strSQL, oConn, 3
		
		
		
		if not oRec.EOF then
		intDayHours = oRec("timer_temp")
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

    thisWeekDay = datePart("w", realDato, 2,2)
    thisLeft = thisWdt(thisWeekDay) + lft
   %>
   <span id="tt" style="position:absolute; padding:0px 0px 0px 2px; left:<%=thisLeft%>px; top:<%=tp+15%>px; width:<%=10*formatnumber(intDayHoursDivval,0)%>px; z-index:1000; height:14px; background-color:#FFC0CB;">
   <table callpadding=0 cellspacing=0 border=0 width=100%></tr><td class=lille style="padding-right:5px;" align=right><%=formatnumber(intDayHours, 2) %></td></table>
    
   </span>
   <%end if %>




   
 <%  
 end function 















public aktidKorrFundet
function indlasKorrektioner(mid, korrKom, korrReal, dtInd)

                '**** Ændrer status på planlagt ferie til afholdt ferie ***'
		        '*** (indlæser ny registrering så historik beholdes) ***'
		        '**** Tjekker om der findes reg. i forvejen *****'
		        '113 Korrek. Kom
		        '114 Korr. Real
		        
		       
		        for f = 1 to 2
		        
		        if f = 1 then
		        
                aktype = 113
                timerThis = split(korrKom, ":")
		        
                
                for t = 0 To UBOUND(timerThis)

                if t = 0 then
                timerThisSQL = timerThis(t)
                else
                timerThisSQL = timerThisSQL &","& formatnumber((timerThis(t) * 100) / 60, 0)
                end if

                next


                else
		        aktype = 114
		        timerThis = korrReal
                timerThisSQL = timerThis
                end if
		        
		        aktid = 0
		        editDato = year(now)&"/"& month(now)&"/"&day(now)
                
                call meStamdata(mid)

		        
		        '** Finder navn og id på Korrektions akt. ***'
		        strSQLKoafn = "SELECT a.id, a.navn, k.kkundenavn, k.kkundenr, j.jobnr, j.jobnavn FROM job j"_
		        &" LEFT JOIN aktiviteter a ON (a.fakturerbar = "& aktype &" AND a.job = j.id) "_
		        &" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_
                &" WHERE j.jobstatus = 1 AND a.id <> 'NULL' GROUP BY a.id ORDER BY a.id DESC"
		        
		       'Response.write strSQLKoafn & "<br>"
               'Response.flush
		        
		        oRec3.open strSQLKoafn, oCOnn, 3
		        if not oRec3.EOF then
		        
		        if oRec3("id") <> "" then
		        aktid = oRec3("id")
		        aktnavn = oRec3("navn")

                jobnavn = oRec3("jobnavn")
                jobnr = oRec3("jobnr")

                knavn = oRec3("kkundenavn")
                knr = oRec3("kkundenr")

		        end if
		        
		        end if
		        oRec3.close
		        
		       if cdbl(aktid) <> 0 then
		        
		        
                       
                        tpThis = 0
                        kpThis = 0
                        kursThis = 100
                        kommThis = "Korrektion indlæst"
                        origin = 5
                        ekstTimer = 0

                        '*** Hent først evt. manuel indlæsning på dagen '**
                        strSQLfindesTimer = "SELECT timer FROM timer WHERE tdato = '"& year(dtInd) &"/"& month(dtInd) & "/"& day(dtInd) &"' AND tmnr = "& mid & " AND tfaktim = "& aktype & " AND taktivitetid = " & aktid
                        
                        'Response.write strSQLfindesTimer & "<br>"
                        'Response.flush
                        oRec4.open strSQLfindesTimer, oConn, 3
                        
                        while not oRec4.EOF 
                        ekstTimer = ekstTimer + (oRec4("timer"))
                        oRec4.movenext
                        wend
                        oRec4.close 

                        '** Tømmer ***
                        strSQLDelTimer = "DELETE FROM timer WHERE tdato = '"& year(dtInd) &"/"& month(dtInd) & "/"& day(dtInd) &"' AND tmnr = "& mid & " AND tfaktim = "& aktype & " AND taktivitetid = " & aktid
                        oConn.execute(strSQLDelTimer)

                        timerThisSQL = (timerThisSQL/1 + (ekstTimer/1))
                        timerThisSQL = replace(timerThisSQL, ",", ".")

                        'Response.write "(timerThisSQL/1 + ekstTimer/1) MID: "& mid &": " & timerThisSQL &"+"& ekstTimer & "<br>"


                        '*** Indlæser afholdt ferie ***'
		                strSQLKoins = "INSERT INTO timer "_
		                &"("_
		                &" timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, "_
		                &" timerkom, TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, "_
		                &" editor, kostpris, seraft, sttid, sltid, valuta, kurs, origin "_
		                &") "_
		                &" VALUES "_
		                &" (" _
		                & timerThisSQL &", "& aktype &", "_
		                & "'"& year(dtInd) &"/"& month(dtInd) & "/"& day(dtInd) &"', "_
		                & "'"& meNavn &"', "_
		                & mid &", "_
		                & "'"& jobnavn &"', "_
		                & "'"& jobnr &"', "_
		                & "'"& knavn &"', "_
		                & knr &", "_
		                & "'"& kommThis &"', "_
		                & aktid &", "_
		                & "'"& aktnavn &"', "_
		                & year(now) &", "_
		                & tpThis &", "_
		                & "'"& year(now) &"/"& month(now) & "/"& day(now) &"', 0, "_
		                & "'00:00:01', "_
		                & "'Korrektions indlæsning', "_
		                & kpThis &", 0, '00:00:00', '00:00:00', 1, "_
		                & kursThis &", "& origin &")"
		                
                       
		                
		                
		                
                        'Response.write strSQLKoins & "<br>"
                        'Response.flush

                        oConn.execute(strSQLKoins)
		            
                
               

               
                
                end if '*** aktid <> 0
                
                next
                
                if f = 1 then

                    if session("stempelur") <> 0 then 'er stempelur slået til og sder der tjekkes herfor
                    aktidKorrFundet = aktid
                    else
                    aktidKorrFundet = -1 
                    end if

                else

                aktidKorrFundet = aktid
                end if
                
                
                'Response.end

end function



public lonKorsel_lukketPerDt, lonKorsel_lukketIO
function lonKorsel_lukketPer(tjkDay, hr)

    'hr: job risiko -2 eller Stempelur altid -2
    lonKorsel_lukketPerDt = "1-1-2002"     
    strSQL5 = "SELECT lk_dato FROM lon_korsel WHERE lk_id <> 0 ORDER BY lk_dato DESC"
	oRec5.open strSQL5, oConn, 3 
	if not oRec5.EOF then
	    lonKorsel_lukketPerDt = oRec5("lk_dato") 
	end if
	oRec5.close 

'Response.Write "lonKorsel_lukketPerDt "& lonKorsel_lukketPerDt &" tjkDay " & tjkDay & "<br>"

if cdate(lonKorsel_lukketPerDt) >= cDate(day(tjkDay) &"-"& month(tjkDay) &"-"& year(tjkDay)) AND hr = -2 then
lonKorsel_lukketIO = 1
else
lonKorsel_lukketIO = 0
end if
end function


function lonKorsel_lukketPerPrDato(tjkDato)

    
    lonKorsel_lukketPerDt = licensstdato

    tjkDatoThisSQL = year(tjkDato) &"-"& month(tjkDato) &"-"& day(tjkDato)


    strSQL5 = "SELECT lk_dato FROM lon_korsel WHERE lk_id <> 0 AND lk_dato <= '"& tjkDatoThisSQL &"' ORDER BY lk_dato DESC"
	'Response.write tjkDatoSQL &"<br>"& strSQL5 & "<br>"
    oRec5.open strSQL5, oConn, 3 
	if not oRec5.EOF then
	    lonKorsel_lukketPerDt = oRec5("lk_dato") 
	end if
	oRec5.close 

    'Response.write "lonKorsel_lukketPerDt: "& lonKorsel_lukketPerDt

    'Response.end

end function

    
%>