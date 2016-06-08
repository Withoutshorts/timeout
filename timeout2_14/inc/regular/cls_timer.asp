 
 <%


     '*** Timer denne uge ****'
public manTimer, tirTimer, onsTimer, torTimer, freTimer, lorTimer, sonTimer
public manTxt, tirTxt, onsTxt, torTxt, freTxt, lorTxt, sonTxt
public totTimerWeek
    
    
function timerDenneUge(usemrn, lto, varTjDatoUS_man, aty_sql_typer)

    'response.write "her"
    'response.flush


     manTimer = 0
    tirTimer = 0
    onsTimer = 0
    torTimer = 0
    freTimer = 0
    lorTimer = 0
    sonTimer = 0

    manTxt = ""
    tirTxt = ""
    onsTxt = ""
    torTxt = ""
    freTxt = ""
    lorTxt = ""
    sonTxt = ""

     for x = 1 to 7
      

           
                tjkTimerTotDato = dateAdd("d", x-1, varTjDatoUS_man)
              

                tjkTimerTotDato = year(tjkTimerTotDato) &"/"& month(tjkTimerTotDato) &"/"& day(tjkTimerTotDato)

                select case lto
                case "dencker", "intranet - local", "epi", "epi_no", "epi_ab", "epi_sta", "epi_catitest"
                timGrpBy = "tknr"
                case else
                timGrpBy = "tjobnr"
                end select

            strSQLtim = "SELECT ROUND(SUM(timer),2) AS sumtimer, tjobnr, tjobnavn, Tknavn FROM timer WHERE ("& aty_sql_typer &") AND tdato = '"& tjkTimerTotDato &"' AND tmnr = "& usemrn & " GROUP BY "& timGrpBy


            'if lto = "synergi1" AND session("mid") = 1 then
            'Response.Write strSQLtim & "<br><br>"
            'Response.flush
            'end if
            
            
            oRec3.open strSQLtim, oConn, 3
            While not oRec3.EOF 

                'if isnull(oRec3("sumtimer")) <> true then


                if timGrpBy = "tknr" then


                select case x
                case 1
                manTimer = manTimer/1 + oRec3("sumtimer")/1
                manTxt = manTxt & "<br>"& left(oRec3("Tknavn"), 10) &": "& oRec3("sumtimer")
                case 2
                tirTimer = tirTimer/1 + oRec3("sumtimer")/1
                tirTxt = tirTxt & "<br>"& left(oRec3("Tknavn"), 10) &": "& oRec3("sumtimer")
                case 3
                onsTimer = onsTimer/1 + oRec3("sumtimer")/1
                onsTxt = onsTxt & "<br>"& left(oRec3("Tknavn"), 10) &": "& oRec3("sumtimer")
                case 4
                torTimer = torTimer/1 + oRec3("sumtimer")/1
                torTxt = torTxt & "<br>"& left(oRec3("Tknavn"), 10) &": "& oRec3("sumtimer")
                case 5
                freTimer = freTimer/1 + oRec3("sumtimer")/1
                freTxt = freTxt & "<br>"& left(oRec3("Tknavn"), 10) &": "& oRec3("sumtimer")
                case 6
                lorTimer = lorTimer/1 + oRec3("sumtimer")/1
                lorTxt = lorTxt & "<br>"& left(oRec3("Tknavn"), 10) &": "& oRec3("sumtimer")
                case 7
                sonTimer = sonTimer/1 + oRec3("sumtimer")/1
                sonTxt = sonTxt & "<br>"& left(oRec3("Tknavn"), 10) &": "& oRec3("sumtimer")
                end select


                else

                select case x
                case 1
                manTimer = manTimer/1 + oRec3("sumtimer")/1
                manTxt = manTxt & "<br><b>"& left(oRec3("Tknavn"), 10) &"</b><br>"& left(oRec3("tjobnavn"), 8) & ": "& oRec3("sumtimer")
                case 2
                tirTimer = tirTimer/1 + oRec3("sumtimer")/1
                tirTxt = tirTxt & "<br><b>"& left(oRec3("Tknavn"), 10) &"</b><br>"& left(oRec3("tjobnavn"), 8) & ": "& oRec3("sumtimer")
                case 3
                onsTimer = onsTimer/1 + oRec3("sumtimer")/1
                onsTxt = onsTxt & "<br><b>"& left(oRec3("Tknavn"), 10) &"</b><br>"& left(oRec3("tjobnavn"), 8) & ": "& oRec3("sumtimer")
                case 4
                torTimer = torTimer/1 + oRec3("sumtimer")/1
                torTxt = torTxt & "<br><b>"& left(oRec3("Tknavn"), 10) &"</b><br>"& left(oRec3("tjobnavn"), 8) & ": "& oRec3("sumtimer")
                case 5
                freTimer = freTimer/1 + oRec3("sumtimer")/1
                freTxt = freTxt & "<br><b>"& left(oRec3("Tknavn"), 10) &"</b><br>"& left(oRec3("tjobnavn"), 8) & ": "& oRec3("sumtimer")
                case 6
                lorTimer = lorTimer/1 + oRec3("sumtimer")/1
                lorTxt = lorTxt & "<br><b>"& left(oRec3("Tknavn"), 10) &"</b><br>"& left(oRec3("tjobnavn"), 8) & ": "& oRec3("sumtimer")
                case 7
                sonTimer = sonTimer/1 + oRec3("sumtimer")/1
                sonTxt = sonTxt & "<br><b>"& left(oRec3("Tknavn"), 10) &"</b><br>"& left(oRec3("tjobnavn"), 8) & ": "& oRec3("sumtimer")
                end select


                end if

                'else

               


                'end if

            oRec3.movenext
            wend 
            oRec3.close

            next

    totTimerWeek = (manTimer + tirTimer + onsTimer + torTimer + freTimer + lorTimer + sonTimer)



end function








     sub mtimerIperHgt

                    if timerThis <> 0 then 'hgttimmtyp
                    bdThis = 1
                    timerThis = formatnumber(timerThis,2)
                    timerThisTxt = formatnumber(timerThis, 2) & " t."
                    else
                    timerThisTxt = ""
                    timerThis = ""
                    bdThis = 0
                    end if 
            
                    'bgThis = "#EFf3FF"
                    'select case right(mt,1) 
                    'case 0,2,4,6,8
                    'bgThis = "#CCCCCC" '"yellowgreen"
                    'case 9
                    'bgThis = "#999999" '"green"
                    'case 2
                    'bgThis = "#F7F7F7" '"#C2D69A"
                    'case 0,5
                    'bgThis = "#C4C4C4" '"#6CAE1C"
                    bgThis = "#D6Dff5"
                    'case else
                    'bgThis = "#8CAAE6" '"#DCF5BD"
                    'end select
                    
                    


                    if hgttimmtyp = 0 then
                    bgThis = "#ffffff"
                    end if
                
                    if hgttimmtyp > 100 then
                    hgttimmtyp = 100
                    end if
          
                    %>

                        <tr>
                    <td align="right" style="width:60px; border-bottom:0px #CCCCCC solid; padding:0px 3px 0px 2px; font-size:8px; white-space:nowrap;">
                    <%=left(mNavn, 10)%>:</td>
                        <td class=lille style="width:100px;">
                    <div style="width:<%=hgttimmtyp%>px; padding:0px 0px 0px 3px; font-size:8px; background-color:<%=bgThis%>; border-bottom:0px #CCCCCC solid; border-right:0px #CCCCCC solid; white-space:nowrap;"><%=timerThisTxt %></div>
                    </td>
                                                               
                    </tr>
                    <%
                    if len(trim(timerThis)) <> 0 then
                    timerIalt = timerIalt/1 + timerThis/1
                    else
                    timerIalt = timerIalt
                    end if

     end sub
 
 public hgtKvo, hgtKvoTxt, mt, hgttimmtyp, mNavn
 function timer_fordeling_medarb_typer(jobnr, per, timerTotJob, stDatoMtyp, slDatoMtyp, jobid)
                

                stDatoMtypTxt = stDatoMtyp
                slDatoMtypTxt = slDatoMtyp

                stDatoMtyp = year(stDatoMtyp) &"/"& month(stDatoMtyp) &"/"& day(stDatoMtyp)
                slDatoMtyp = year(slDatoMtyp) &"/"& month(slDatoMtyp) &"/"& day(slDatoMtyp)

                antalManeder = datediff("m", stDatoMtyp, slDatoMtyp, 2,2)
                        
                
                timerTotJob = timerTotJob

                timerIalt = 0

                if per <> 0 then
                dvHgt = 100
                lmt = 10
                dvBd = "1"
                dvPd = "10px 0px 0px 4px"
                else
                dvHgt = 181
                lmt = 10
                dvBd = "1"
                dvPd = "4px 4px 4px 4px"
                end if


                '***** Fordeling på medarbejder / medarbejertyper
                call bdgmtypon_fn()
                if cint(bdgmtypon_val) = 1 then
                'case "epi", "xintranet - local", "epi_no", "epi_sta", "epi_ab"
                medTxt1 = "medarbejdertyper"
                tfordGrpBy = "tjobnr, mt.id"
                mnavnFlt = "mt.type"
                else
                medTxt1 = "medarbejdere"
                tfordGrpBy = "tjobnr, m.mid"
                mnavnFlt = "m.mnavn"
                end if
                %>
                
                <span style="font-size:9px;">
                <b>Fordeling på <%=medTxt1 %>:</b><br />
                   
                    <%if cint(per) <> 0 then 
                     
                        if cint(bdgmtypon_val) = 1 AND cint(bdgmtypon_prgrp) > 1 then 
                        stDatoMtyp = year(stDatoMtyp) &"/"& month(stDatoMtyp) &"/1"%>
                        Per.: 1. <%=left(monthname(month(stDatoMtypTxt)), 3) &" "& year(stDatoMtypTxt) &" - "& dainmo &". "& left(monthname(month(slDatoMtypTxt)), 3) &" "& year(slDatoMtypTxt)%> <!--<span style="color:#999999;">(<%=antalManeder %> måneder)</span>-->
                        <%
                        else%>
                        Per.: <%=formatdatetime(stDatoMtypTxt, 1) & " - "& formatdatetime(slDatoMtypTxt, 1)%><!--<span style="color:#999999;"><br />(<%=antalManeder %> måneder, maks. 10 medarb.)</span>  -->
                        <%end if
                    else %>
                    Periode: total uanset periode<span style="color:#999999;"><br /> (maks. 10 medarb./typer)</span>
                    <%end if %>
                    </span>
                    <table cellpadding=0 cellspacing=1 border=0>
                                                           

                                                             
                    <%
                                                                 
                                                                 
                    

                        hgttimmtyp = 0
                        timerThis = 0

                        if cint(bdgmtypon_val) = 1 AND cint(bdgmtypon_prgrp) > 1 then 
                            



                                        mt = 0
                                        for t = 0 to UBOUND(mtypeids)

                                 

                                            'if len(trim(mtypgrpnavn(t))) <> 0 then
                                
                        
                                            if per <> 0 then 'Vis total (konsolideret) eller periode  
                                            strSQLjl = "SELECT SUM(timer) AS timer FROM timer_konsolideret_tot WHERE jobid = "& jobid & " AND (mtype = "& mtypeids(t) & " "& mtypesostergp(t) &") AND dato BETWEEN '"& stDatoMtyp &"' AND '"& slDatoMtyp &"' GROUP BY jobid"
                                            else
                                            strSQLjl = "SELECT SUM(timer) AS timer FROM timer_konsolideret_tot WHERE jobid = "& jobid & " AND (mtype = "& mtypeids(t) & " "& mtypesostergp(t) &") GROUP BY jobid"
                                            end if
                                            'Response.write strSQLmtypbgt
                                            
                                            'if session("mid") = 1 then
                                            'Response.Write "<tr><td>"& strSQLjl &"</td></tr>"
                                            'Response.flush
                                            'end if        

                                            oRec2.open strSQLjl, oConn, 3
                                            if not oRec2.EOF then
                                
                                
                                            hgttimmtyp = formatnumber((oRec2("timer")/timerTotJob) * 100, 0)
                                            timerThis = oRec2("timer")
                                            mNavn = mtypenavne(t)


                                            call mtimerIperHgt
                        
                        
                                            mt = mt + 1

                                
                                            end if
                                            oRec2.close


                                            'end if
                        
                                    next   


                        else
                        strSQLjl = "SELECT sum(timer) AS timer, "& mnavnFlt &" AS navn FROM timer AS t "_
                        &" LEFT JOIN medarbejdere AS m ON (m.mid = t.tmnr)"_
                        &" LEFT JOIN medarbejdertyper AS mt ON (mt.id = m.medarbejdertype) "_
                        &" WHERE tjobnr = '"& jobnr & "'"  
                        
                        if per <> 0 then 
                         strSQLjl = strSQLjl &" AND tdato BETWEEN '"& stDatoMtyp &"' AND '"& slDatoMtyp &"'"
                        end if

                        strSQLjl = strSQLjl &" AND ("& aty_sql_realhours &") GROUP BY "& tfordGrpBy &" ORDER BY timer DESC LIMIT "&lmt

                     

                        'if session("mid") = 1 then
                        '    Response.Write strSQLjl
                        '    Response.flush
                        'end if

                        mt = 0
                        oRec2.open strSQLjl, oConn, 3
                        while not oRec2.EOF 

                        hgttimmtyp = formatnumber((oRec2("timer")/timerTotJob) * 100, 0) 
                        timerThis = oRec2("timer")
                        mNavn = oRec2("navn")


                        call mtimerIperHgt
                        
                        
                        mt = mt + 1

                        oRec2.movenext
                        wend
                        oRec2.close


                           end if







                        '*************************************************
                        '*** Præsentation
                        '*************************************************

                        if hgttimmtyp > 100 then
                        hgttimmtyp = 20
                        else
                        hgttimmtyp = hgttimmtyp/5 
                        end if

                        hgtKvo = 1
                        hgtKvoTxt = "> 100"

                        if timerIalt > 1000 then
                        hgtKvo = 10
                        hgtKvoTxt = "> 1000"
                        end if

                        if timerIalt > 5000 then
                        hgtKvo = 50
                        hgtKvoTxt = "> 10000"
                        end if
                                                                
                    %>
                                                               
                                                             
	    
                </table>
            


<%
 end function


                                           


'*** Finder evt. alternativ timepris på medarbejder ****
'*** Hvis der ikke findes alternative timepriser bruges default ****
public intTimepris, foundone, intValuta, valutaISO
function alttimepris(useaktid, intjobid, strMnr, upd) 
                            
                            intTimepris = 0
							timeprisAlt = 0
							foundone = "n"
							strSQL = "SELECT id AS tpid, jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta FROM timepriser WHERE jobid = "& intjobid &" AND aktid = "& useaktid &" AND medarbid =  "& strMnr
							'Response.Write strSQL & "<br>"
                            'Response.end	
						    oRec2.open strSQL, oConn, 3 
							if not oRec2.EOF then
							        
							        foundone = "y"
				                    intTimepris = oRec2("6timepris")
				                    intValuta = oRec2("6valuta")
									tpid = oRec2("tpid")
									timeprisAlt = oRec2("timeprisalt")
									
									if cdbl(intTimepris) = 0 AND cint(timeprisAlt) <> 6 then
									
							
							            Select case timeprisAlt
							            case 1
							            timeprisalernativ = "timepris_a1"
							            valutaAlt = "tp1_valuta"
							            case 2
							            timeprisalernativ = "timepris_a2"
							            valutaAlt = "tp2_valuta"
							            case 3
							            timeprisalernativ = "timepris_a3"
							            valutaAlt = "tp3_valuta"
							            case 4
							            timeprisalernativ = "timepris_a4"
							            valutaAlt = "tp4_valuta"
							            case 5
							            timeprisalernativ = "timepris_a5"
							            valutaAlt = "tp5_valuta"
							           
							            case else
							            timeprisalernativ = "timepris"
							            valutaAlt = "tp0_valuta"
							            end select
							            
							            
							        strSQL3 = "SELECT mid, "&timeprisalernativ&" AS useTimepris, "& valutaAlt &" AS useValuta, medarbejdertype FROM medarbejdere, medarbejdertyper WHERE mid =" & strMnr & " AND medarbejdertyper.id = medarbejdertype"
									
									'Response.Write strSQL3
									'Response.flush
									oRec3.open strSQL3, oConn, 3 
									if not oRec3.EOF then
									intTimepris = oRec3("useTimepris")
									intValuta = oRec3("useValuta")
									end if
									oRec3.close
									
									intTimepris = replace(intTimepris, ",", ".")
									
									
									
									if cint(upd) = 1 then
									strSQLupdtp = "UPDATE timepriser SET 6timepris = "& intTimepris &", 6valuta = "& intValuta &" WHERE id = " & tpid
									oConn.execute(strSQLupdtp)
									end if
									
									
									end if
									
									strSQLval = "SELECT valutakode FROM valutaer WHERE id = " & intValuta
									oRec3.open strSQLval, oConn, 3 
									if not oRec3.EOF then
									valutaISO = oRec3("valutakode")
									end if
									oRec3.close
									
									'Response.Write strSQLval &  "valutaISO" & valutaISO
									
						    
						    end if 
							oRec2.close 
						
						
						
end function


public timerforbrugt, kostpris, timeOms, tp, OmsReal, salgsOmkFaktisk, matSalgsprisReal, matKobsprisReal
function timeRealOms(jobnr, sqlDatostart, sqlDatoslut, nettoomstimer, fastpris, budgettimerIalt, aty_sql_realhours)
         
         
        timerforbrugt = 0
        kostpris = 0
        timeOms = 0   


        '*** Timer forbrugt ***
		'** Hvis mtypebudget slået til og Totalforbrug (konsolideret valgt)
        'Response.write "bdgmtypon_val:" & bdgmtypon_val & " visSimpel:" & visSimpel
        'Response.end
        if cint(bdgmtypon_val) = 1 AND visSimpel = 2 then
        strSQL2 = "SELECT sum(t.timer) AS timerforbrugt, sum(kost) AS kostpris, sum(t.belob) AS timeOms FROM timer_konsolideret_tot AS t WHERE jobid = "& oRec("id")
                    
                    if realfakpertot <> 0 then
		            strSQL2 = strSQL2 &" AND dato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"'"
		            end if
        
                    strSQL2 = strSQL2 &" GROUP BY jobid"

        else

        strSQL2 = "SELECT sum(t.timer) AS timerforbrugt, sum(kostpris*t.timer*(t.kurs/100)) AS kostpris, sum(t.timer*(t.timepris*(t.kurs/100))) AS timeOms FROM timer t WHERE t.tjobnr = '"& jobnr &"' AND ("& aty_sql_realhours &") "
         
        

        if realfakpertot <> 0 then
		strSQL2 = strSQL2 &" AND tdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"'"
		end if

	    strSQL2 = strSQL2 &" GROUP BY t.tjobnr"
	    end if	

		'if session("mid") = 1 then
        'Response.Write strSQL2 & "<br><br>"
        'end if

        'Response.end
		oRec2.open strSQL2, oConn, 3
		

		if not oRec2.EOF then
		timerforbrugt = oRec2("timerforbrugt")
        kostpris = oRec2("kostpris")
        timeOms = oRec2("timeOms")
		end if
		oRec2.close
		
        


		if len(timerforbrugt) <> 0 then
		timerforbrugt = timerforbrugt
		else
		timerforbrugt = 0
		end if

        if len(kostpris) <> 0 then
        kostpris = kostpris
        else
        kostpris = 0
        end if

        if len(timeOms) <> 0 then
        timeOms = timeOms
        else
        timeOms = 0
        end if

       

        '*** Timepriser / Omsætning ***'
		OmsReal = 0
        tp = 0

        '** Fastpris **'
		if cint(fastpris) = 1000 then '1: fastpris / 0: lbn timer / 1000: Deaktiveret altid lbn timer beregning på Real timer uanset om det er fastpris eller lbn. timer
                                                
            if budgettimerIalt <> 0 then
            tp = (nettoomstimer/budgettimerIalt)
            else
            tp = (nettoomstimer/1)
            
            end if

            OmsReal =  timerforbrugt * (tp)
                							
        else
		'** Beløb **'
		OmsReal = timeOms

        if timerforbrugt <> 0 then
        tp = (timeOms/timerforbrugt)
        else
        tp = 0
        end if
							                
        end if


        'Response.Write "tp:" & tp

         '**** Materiale forbrug ****'
        strSQLmat = "SELECT SUM(matkobspris * matantal * (kurs/100)) AS udgifterfaktisk,  SUM(matsalgspris * matantal * (kurs/100)) AS matSalgspris FROM materiale_forbrug WHERE jobid = "& oRec("id") 
        
        if realfakpertot <> 0 then
		strSQLmat = strSQLmat &" AND forbrugsdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"'"
		end if

        strSQLmat = strSQLmat &" GROUP BY jobid "

        salgsOmkFaktisk = 0
        
        oRec2.open strSQLmat, oConn, 3
		if not oRec2.EOF then
		salgsOmkFaktisk = oRec2("udgifterfaktisk")
        matKobsPrisReal = salgsOmkFaktisk 
        matSalgsprisReal = oRec2("matSalgspris")
        end if
		oRec2.close

        if len(trim(salgsOmkFaktisk)) <> 0 then
        salgsOmkFaktisk = salgsOmkFaktisk
        else
        salgsOmkFaktisk = 0
        end if

end function
   




public timerThis, pleaseUpdate, lastUpdatedRecordID, lastRecordID, delrecords
function timer_konsolider(lto,dothis)




lastMid = 0
lastJobid = 0
lastAkt = 0
lastMnth = 0

timerThis = 0
pleaseUpdate = 0
lastUpdatedRecordID = 0


'*** Vigtigt at der er langt tilbage i tiden ****
'*** Mindst 3 uger for EPI, så extsysid ikke bliver slettet fra timer og indlæst igen **'
select case lto 
case "dencker"
dateNow = dateAdd("m",-1, now)
dateNowSQL = year(dateNow) &"/"& month(dateNow) &"/1" 
dateBegin = dateAdd("m",-5, now) 
dateBeginSQL = year(dateBegin) &"/"& month(dateBegin) &"/1" 

strSQL = "SELECT tid, tmnr, tjobnr, taktivitetid, tdato, timer FROM timer WHERE timer < 0.5 AND tdato BETWEEN  '"& dateBeginSQL &"' AND '"& dateNowSQL &"' ORDER BY tmnr, tjobnr, taktivitetid, tdato" 'AND origin = 91

case "epi", "epi_no", "epi_sta", "epi_ab"
dateNow = dateAdd("m", -2, now)
dateNowSQL = year(dateNow) &"/"& month(dateNow) &"/31" 

dateBegin = dateAdd("m",-3, now) 
dateBeginSQL = year(dateBegin) &"/"& month(dateBegin) &"/1" 

strSQL = "SELECT tid, tmnr, tjobnr, taktivitetid, tdato, timer FROM timer WHERE timer < 4.5 AND tdato BETWEEN  '"& dateBeginSQL &"' AND '"& dateNowSQL &"' AND tfaktim = 91 ORDER BY tmnr, tjobnr, taktivitetid, tdato"

case "intranet - local", "outz" 
dateNow = dateAdd("yyyy", -2, now)
dateNowSQL = year(dateNow) &"/"& month(dateNow) &"/1"

dateBegin = dateAdd("yyyy",-10, now) 
dateBeginSQL = year(dateBegin) &"/"& month(dateBegin) &"/1" 
 
strSQL = "SELECT tid, tmnr, tjobnr, taktivitetid, tdato, timer FROM timer WHERE timer < 7.5 AND tdato BETWEEN  '"& dateBeginSQL &"' AND '"& dateNowSQL &"' AND tfaktim = 1 ORDER BY tmnr, tjobnr, taktivitetid, tdato"

end select

'Response.Write strSQL & "<br><br>"
'Response.flush

x = 0
oRec.open strSQL, oConn, 3
while not oRec.EOF 
					
                    			
							

                   
                        
                        if cdbl(x) <> 0 then
                            
                            select case lto
                            case "dencker"
                                
                                '*** skal timer konsolideres på sin egen akt eller på en fælles aktivitet ***'
                                'if lastAkt <> oRec("taktivitetid") then 
                                'pleaseUpdate = 1
                                'end if

                                if lastJobid <> oRec("tjobnr") then
                                pleaseUpdate = 1
                                end if

                                if lastMid <> oRec("tmnr") then
                                pleaseUpdate = 1
                                end if

                                '**** Skal de samles på månedsbasis ***'
                                if lastMnth <> month(oRec("tdato")) then
                                pleaseUpdate = 1
                                end if

                                '**** Skal de samles på dagsbasis ***'
                                'if lastDay <> day(oRec("tdato")) then
                                'pleaseUpdate = 1
                                'end if


                            case "epi", "epi_no", "epi_sta", "epi_ab"

                                 '*** skal timer konsolideres på sin egen akt eller på en fælles aktivitet ***'
                                'if lastAkt <> oRec("taktivitetid") then 
                                'pleaseUpdate = 1
                                'end if
                                

                                if lastJobid <> oRec("tjobnr") then
                                pleaseUpdate = 1
                                end if

                                if lastMid <> oRec("tmnr") then
                                pleaseUpdate = 1
                                end if

                                '**** Skal de samles på månedsbasis ***'
                                'if lastMnth <> month(oRec("tdato")) then
                                'pleaseUpdate = 1
                                'end if

                                '**** Skal de samles på dagsbasis ***'
                                if lastDay <> day(oRec("tdato")) then
                                pleaseUpdate = 1
                                end if


                            case "intranet - local", "outz"

                                
                                '*** skal timer konsolideres på sin egen akt eller på en fælles aktivitet ***'
                                if lastAkt <> oRec("taktivitetid") then 
                                pleaseUpdate = 1
                                end if

                            
                                if lastJobid <> oRec("tjobnr") then
                                pleaseUpdate = 1
                                end if

                                if lastMid <> oRec("tmnr") then
                                pleaseUpdate = 1
                                end if

                                '**** Skal de samles på månedsbasis ***'
                                if lastMnth <> month(oRec("tdato")) then
                                pleaseUpdate = 1
                                end if

                                '**** Skal de samles på dagsbasis ***'
                                'if lastDay <> day(oRec("tdato")) then
                                'pleaseUpdate = 1
                                'end if

                            end select

                        end if

                    
                     if cint(pleaseUpdate) = 1 then '**indlæsrecord

                        call updateTimer

                    else

                         if cdbl(lastUpdatedRecordID) <> cdbl(lastRecordID) then
                         delrecords = delrecords & " OR tid = "& lastRecordID
                         end if

                    end if
                    
                    timerThis = timerThis + formatnumber(oRec("timer"))           

                   
                    

                    'Response.Write oRec("tid") &", "& oRec("tmnr") &", "& oRec("tjobnr") &", "& oRec("taktivitetid") &", "& oRec("tdato") &", "& oRec("timer") & "<br>"	

                    '**** Alle records indlæses til konsolider, hvis nu der skal rulles tilbage ****'
                    timerK = replace(oRec("timer"), ",", ".")
                    datok = year(oRec("tdato")) & "/" & month(oRec("tdato")) & "/" & day(oRec("tdato"))
                    kortDato = year(now) & "/" & month(now) & "/" & day(now) 

                    
                    strSQLkon = "INSERT IGNORE INTO timer_konsolideret (tk_dato, tk_timerid, tk_mnr, tk_jobnr, tk_aid, tk_timer, dato_kons) "_
                    &" VALUES ('"& datok &"', "& oRec("tid") &", "& oRec("tmnr") &", '"& oRec("tjobnr") &"', "& oRec("taktivitetid") &", "& timerK &", '"& kortDato &"')"
                    oConn.execute(strSQLkon)

                    'Response.write strSQLkon
                    'Response.end
                    'Response.Write timerK
                    'Response.end



lastMid = oRec("tmnr")
lastJobid = oRec("tjobnr")
lastAkt = oRec("taktivitetid")
lastMnth = month(oRec("tdato"))
lastRecordID = oRec("tid")
lastDay = day(oRec("tdato"))
 

x = x + 1
oRec.movenext
wend
oRec.close



call updateTimer



end function



'**** Func til konsolider timer *****'
sub updateTimer
                    

                    '**** Opdaterer ***'
                    '*** 91 = Easyreg
                     if len(trim(lastRecordID)) <> 0 then
                     lastRecordID = lastRecordID
                     else
                     lastRecordID = 0
                     end if
                     
                     timerThis = replace(timerThis, ".", "")
                     timerThis = replace(timerThis, ",", ".")
                     strSQLupd = "UPDATE timer SET timer = " & timerThis & " WHERE tid = " & lastRecordID
                     ', origin = 91
                     'Response.write strSQLupd & "<br>"
                     'Response.Flush
                     'Response.end

                     lastUpdatedRecordID = lastRecordID
                    
                     oConn.execute(strSQLupd)
                     timerThis = 0
                     pleaseUpdate = 0


                     '***** Renser ****'
                    strSQLdel = "DELETE FROM timer WHERE tid = 0 "& delrecords

                    'Response.write "<br>"& strSQLdel & "<hr>"
                    oConn.execute(strSQLdel)

                    delrecords = ""

                    'Response.flush


end sub



   %>