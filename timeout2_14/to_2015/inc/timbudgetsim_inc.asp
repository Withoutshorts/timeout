<%

sub fy_relprdato_fm
    %>

     <div class="col-lg-2 pad-t5"> Finansår:
                                         <select class="form-control input-small inputclsLeft" name="FM_fy" onchange="submit();">
                                        <%for fyC = -2 to 3
            
                                            

                                            if fyC = -2 then
                                            fyYear = dateAdd("yyyy", -2, now) 
                                            else
                                            fyYear = dateAdd("yyyy", 1, fyYear)
                                            end if

                                            if fyStMd = "7" then
                                            fyEndYear = dateAdd("yyyy", 1, fyYear)
                                            else
                                            fyEndYear = "" 'year(fyYear)
                                            end if

                                           

                                            'fyEndDt = dateAdd("d",-1, aktBudgettjkOnRegAarSt)
                                            'fyEndDt = day(fyEndDt)&". "&left(monthname(month(fyEndDt)), 3)&" "&fyEndYear

                                            if year(y0) = year(fyYear) then
                                            ySel = "SELECTED"
                                            else
                                            ySel = ""
                                            end if


                                            select case lto
                                            case "wwf" '1-7
                                             %>
                                                <option value="<%=year(fyYear) %>" <%=ySel %>>FY: <%=datepart("yyyy", fyYear, 2,2)+1 %>  
                                                     <%if fyEndYear <> "" then  %>
                                                     &nbsp;(<%=right(datepart("yyyy", fyYear, 2,2), 2)%>/<%=right(datepart("yyyy", fyEndYear, 2,2), 2)%>)
                                                     <%end if %>
                                                </option>
                                             <%case else %>
                                                <option value="<%=year(fyYear) %>" <%=ySel %>>FY: <%=datepart("yyyy", fyYear, 2,2)%>  
                                                     <%if fyEndYear <> "" then  %>
                                                     &nbsp;(<%=right(datepart("yyyy", fyYear, 2,2), 2)%>/<%=right(datepart("yyyy", fyEndYear, 2,2), 2)%>)
                                                     <%end if %>
                                                </option>

                                             <%end select %>

                                        <%next %>

                                    </select>
                                      </div>

                                    <div class="col-lg-2 pad-t5 pad-r30"> Vis realiseret pr. dato:
                                           <input class="form-control input-small inputclsLeft" type="text" name="FM_visrealprdato" value="<%=visrealprdato %>" />
                                     
                                      </div>

                                       <!--  <div class="col-lg-2 pad-t5"> Medarb. init: 
                                            <input type="text" class="form-control input-small inputclsLeft" value="<%=medarbinitsTXT%>" name="sogmedarb" placeholder="Medarb. init" style="width:100px;" DISABLED />
                                       </div>-->
                                        
                                        <%if func = "forecast" then %>
                                        <div class="col-lg-1 pad-t5">&nbsp;</div>
                                        <div class="col-lg-4 pad-t5">Medarb.linier vis:<br />
                                         <input type="radio" value="0" id="visKunFCFelter0" name="FM_visKunFCFelter" <%=visKunFCFelterCHK0 %> onclick="submit();"/> Overblik&nbsp;&nbsp;
                                         <input type="radio" value="1" id="visKunFCFelter1" name="FM_visKunFCFelter" <%=visKunFCFelterCHK1 %> onclick="submit();"/> Forecast&nbsp;&nbsp;
                                         <input type="radio" value="2" id="visKunFCFelter2" name="FM_visKunFCFelter" <%=visKunFCFelterCHK2 %> onclick="submit();"/> Realiseret&nbsp;&nbsp;
                                         <input type="radio" value="3" id="visKunFCFelter3" name="FM_visKunFCFelter" <%=visKunFCFelterCHK3 %> onclick="submit();"/> Saldo
                                        <!-- <input type="checkbox" id="visKunFCFelter" name="FM_visKunFCFelter" value="1" <%=visKunFCFelterCHK %> /> Vis kun forecast felter (for medarbejdere)-->
                                 
                                        </div>
                                        <%end if %>


                                       <div class="col-lg-2 pad-t20 pad-r30"><button type="submit" class="btn btn-secondary btn-sm">Search >></button>
                                       </div>

<%
end sub



public budgetFY0GT
function jobaktbudgetfelter(jobnr, jobid, aktid, h1aar, h2aar, h1md, h2md, budgettp)

    jh1 = 0
    jh2 = 0
    jh1h2 = 0
    timerTilH2 = 0

     if aktid <> 0 then
    'SUM(timer*timepris) AS realoms, AVG(timepris) AS gnstp
    'thisTimer = oRec("sumtimer")
    strSQLAktKri = " AND taktivitetid = "& aktid
    strSQLAktResKri = " AND aktid = "& aktid
    sqlRGrpBY = "aktid"
    else
    strSQLAktKri = ""
    strSQLAktResKri = " AND aktid <> 0 " 'AND aktid = 0
    sqlRGrpBY = "jobid" 
    end if

    'if lcase(sogVal) = "all" OR aktid = 0 then 'berenger sum af forecast total på job

        
                    strSQLmedrd = "SELECT SUM(timer) AS sumtimer FROM ressourcer_md WHERE jobid = " & jobid & " "& strSQLAktResKri &"  GROUP BY "& sqlRGrpBY 
                     'response.Write "strSQLmedrd: " & strSQLmedrd
                     'response.flush
                
                     oRec3.open strSQLmedrd, oConn, 3
                     if not oRec3.EOF then

                        jhGT = oRec3("sumtimer")

                     end if
                     oRec3.close
        
                     if cdbl(jhGT) <> 0 then
                     jhGTTxt = formatnumber(jhGT,0)
                     else
                     jhGTTxt = ""
                     end if
                
                     strSQLmedrd = "SELECT SUM(timer) AS sumtimer FROM ressourcer_md WHERE jobid = " & jobid & " "& strSQLAktResKri &" AND aar = "& h1aar &" AND md = "& h1md  & " GROUP BY " & sqlRGrpBY 
                    
                    'if session("mid") = 1 then 
                    'response.Write "strSQLmedrd: " & strSQLmedrd
                    'response.flush
                    'end if            

                     oRec3.open strSQLmedrd, oConn, 3
                     if not oRec3.EOF then

                        jh1 = oRec3("sumtimer")

                     end if
                     oRec3.close

                    
                     useH2 = 900
                     if cint(useH2) = 1 then

                    strSQLmedrd = "SELECT SUM(timer) AS sumtimer FROM ressourcer_md WHERE jobid = " & jobid & " "& strSQLAktResKri &" AND aar = "& h2aar &" AND md = "& h2md & " GROUP BY " & sqlRGrpBY   
                     oRec3.open strSQLmedrd, oConn, 3
                     if not oRec3.EOF then

                        jh2 = oRec3("sumtimer")

                     end if
                     oRec3.close

                     end if

               

                  

    'end if


                    '*** Realtimer ***'
                    select case lto
                    case "wwf"
                    visrealprdatoStartSQL = h1aar & "-7-1"
                    case else
                    visrealprdatoStartSQL = h1aar & "-1-1"
                    end select
                    
                    visrealprdatoSQL = year(visrealprdato) &"-"& month(visrealprdato) &"-"& day(visrealprdato)
                    realomsH1 = 0
                    gnstp = 0
                    realTimeOms = 0
    
                    if aktid <> 0 then
                    timergprBY = "taktivitetid"
                    else
                    timergprBY = "tjobnr"
                    end if
                

                                     strSQLrealTimerJob = "SELECT SUM(timer) AS sumtimer, SUM(timer*timepris) AS realoms, AVG(timepris) AS gnstp FROM timer WHERE tjobnr = '" & jobnr & "' "& strSQLAktKri &" AND (tfaktim = 1 OR tfaktim = 2) AND tdato BETWEEN '"& visrealprdatoStartSQL &"' AND '"& visrealprdatoSQL &"' GROUP BY "& timergprBY   
                                     'response.write strSQLrealTimerJob
                                     'response.flush
                                     oRec3.open strSQLrealTimerJob, oConn, 3
                                     if not oRec3.EOF then

                                        thisTimer = oRec3("sumtimer")
                                        realomsH1 = oRec3("realoms")
                                        realTimeOms = oRec3("realoms")

                                        if isNull(oRec3("sumtimer")) <> true AND oRec3("sumtimer") <> 0 then
                                        gnstp = oRec3("realoms") / oRec3("sumtimer") 'oRec3("gnstp") '
                                        else
                                        gnstp = 0
                                        end if

                                     end if
                                     oRec3.close

                    
                     if realTimeOms <> 0 then
                     realTimeOmsTxt = formatnumber(realTimeOms, 0)
                     else
                     realTimeOmsTxt = ""
                     end if


                     if gnstp <> 0 then
                     gnstpTxt = formatnumber(gnstp, 0)
                     else
                     gnstpTxt = ""
                     end if
                     


                     realomsH1 = realomsH1
                     jh1h2 = jh1+jh2
                     timerTilH2 = formatnumber((jh1 - (thisTimer)), 0)

                     
                     thisTimer = formatnumber(thisTimer, 0)





                    '** RAMME FY ***'
                     if aktid <> 0 then
                     aktSQLkri = " AND aktid = "& aktid 
                     grpBy = " aktid"

                     aktSQLkriAvgTp = " AND r.aktid = "& aktid
                     grpByTp = " r.aktid"

                     else
                     aktSQLkri = " AND aktid <> 0 "
                     grpBy = " jobid"

                     aktSQLkriAvgTp = " AND r.aktid <> 0 "
                     grpByTp = " r.jobid"
                        
                     end if


                    rammeFY0 = 0
                    fctimepris = 0
                    fctimeprish2 = 0 
                    
                    fcAvgtimepris = 0
                    fctimeOms = 0
                    fctimerAntal = 0
            
                    tpFundetAkt = 0            

                    '*** Findes tp på aktivitet eller nedarves - findes den for en medarb. gælder det ALLE, så ka nden ikke stå til nedarv ***'

                    if aktid <> 0 then
                    strSQLmedtp = "SELECT 6timepris FROM timepriser WHERE jobid = " & jobid & " AND aktid = "& aktid 
                    oRec3.open strSQLmedtp, oConn, 3
                    if not oRec3.EOF then

                    tpFundetAkt = 1
    
                    end if
                    oRec3.close
                    end if


                    if cint(tpFundetAkt) = 1 then
                    strSQLAvgTp = "SELECT COALESCE(SUM(r.timer*tp.6timepris),0) AS fctimeOms, COALESCE(SUM(r.timer),0) AS fctimerAntal FROM ressourcer_md AS r "_
                    &" LEFT JOIN timepriser AS tp ON (tp.jobid = "& jobid &" AND tp.aktid = r.aktid AND medarbid = medid) "_
                    &" WHERE r.jobid = "& jobid &" "& aktSQLkriAvgTp &" "& onlyThisMedarbidsFC & " GROUP BY "& grpByTp

                    else 'nedarves
                    strSQLAvgTp = "SELECT COALESCE(SUM(r.timer*tp.6timepris),0) AS fctimeOms, COALESCE(SUM(r.timer),0) AS fctimerAntal FROM ressourcer_md AS r "_
                    &" LEFT JOIN timepriser AS tp ON (tp.jobid = "& jobid &" AND medarbid = medid) "_
                    &" WHERE r.jobid = "& jobid &" "& aktSQLkriAvgTp &" "& onlyThisMedarbidsFC & " GROUP BY "& grpByTp
                    end if  
                

                    
                    'response.write strSQLAvgTp & "<br><br>"
                    'response.flush

                     oRec3.open strSQLAvgTp, oConn, 3
                     while not oRec3.EOF 

                        fctimerAntal = fctimerAntal + oRec3("fctimerAntal")
                        fctimeOms = fctimeOms + oRec3("fctimeOms")
            
                        rammeFY0 = 0' oRec3("timer")
                        fctimepris = 0 'oRec3("fctimepris")
                        fctimeprish2 = 0 'oRec3("fctimeprish2")
                       
                        
                     oRec3.movenext
                     wend 
                     oRec3.close

                    if (fctimerAntal <> 0) then
                    fcAvgtimepris = (fctimeOms/fctimerAntal)
                    else
                    fcAvgtimepris = 0 
                    end if

                    fctimepris = fcAvgtimepris


                     '** RAMME FY1 ***'
                     rammeFY1 = 0
                     strSQLrammeFY1 = "SELECT SUM(timer) AS timer FROM ressourcer_ramme WHERE jobid = " & jobid & " "& aktSQLkri &" AND medid = 0 AND aar = "& year(Y1) &" GROUP BY "& grpBy 

                    'response.write strSQLrammeFY1
                    'response.flush

                     oRec3.open strSQLrammeFY1, oConn, 3
                     if not oRec3.EOF then
                          rammeFY1 = oRec3("timer")
                     end if
                     oRec3.close


                    '** RAMME FY2 ***'
                     rammeFY2 = 0
                     strSQLrammeFY2 = "SELECT SUM(timer) AS timer FROM ressourcer_ramme WHERE jobid = " & jobid & " "& aktSQLkri &" AND medid = 0 AND aar = "& year(Y2) & " GROUP BY "& grpBy    

                    'response.write "<br>FY2: "& strSQLrammeFY2 & "<br><br>"
                    'response.flush

                     oRec3.open strSQLrammeFY2, oConn, 3
                     if not oRec3.EOF then

                        rammeFY2 = oRec3("timer")

                     end if
                     oRec3.close


                    if rammeFY0 <> 0 then
                    rammeFY0 = rammeFY0
                    else
                    rammeFY0 = ""
                    end if

            
                    if rammeFY1 <> 0 then
                    rammeFY1 = rammeFY1
                    else
                    rammeFY1 = ""
                    end if
                
                    if rammeFY2 <> 0 then
                    rammeFY2 = rammeFY2
                    else
                    rammeFY2 = ""
                    end if



        budgetFY0h1 = formatnumber((jh1/1 * fctimepris/1), 0)
        budgetFY0h2 = formatnumber(((realomsH1/1) + jh2/1 * fctimeprish2/1), 0)
        budgetGT = formatnumber((jhGT/1 * fctimepris/1), 0)

         '*** WWF I = 0 / N ***
        'if (thisTimer = 0) then
        'budgetFY0 = formatnumber(budgetFY0h1/1, 0)
        'else 
        'budgetFY0 = formatnumber(budgetFY0h2/1, 0)
        'end if



          '*** ØKO ***
        if (thisTimer = 0) then
        budgetFY0 = formatnumber(budgetGT/1, 0)
        else 
        budgetFY0 = formatnumber(budgetGT/1, 0)
        end if
        


        realomsH1 = formatnumber(realomsH1,0)

            if aktid = 0 then
            budgetFY0GT = budgetFY0GT + jhGT'budgetFY0 
            end if


            if aktid <> 0 then

            aktFMname = "akt"

            aktclassFY0 = "jobakt_fy0_"& jobid
            aktclassFY1 = "jobakt_fy1_"& jobid
            aktclassFY2 = "jobakt_fy2_"& jobid

            aktclassh1 = "jobakt_tph1_"& jobid
   
            aktclassh2 = "jobakt_tph2_"& jobid

            jobaktfltDisabled = ""

            else

            jobaktfltDisabled = "DISABLED"
            aktFMname = ""

            end if


    'if (rammeFY0 < jh1h2) then
    'bgthisFY0 = "lightpink"
    'else
    'bgthisFY0 = ""
    'end if

    
    if aktid <> 0 then
    h1cls = "jh1_"& jobid
    h2cls = "jh1_"& jobid
    jobDisAbled = ""
    else
    h1cls = "x"
    h2cls = "x"
    jobDisAbled = "DISABLED" 
    end if

    'visKunFCFelter
    
    %>
      <td><input type="text" name="FM_<%=aktFMname%>timebudget_FY0" id="FM_timerbudget_FY0_<%=jobid%>_<%=aktid %>" class="jobakt_budgettimer_FY <%=aktclassFY0 %> form-control input-small" value="<%=rammeFY0 %>" style="width:50px; background-color:<%=bgthisFY0%>;" <%=jobDisAbled %> /></td>
     <td><input type="text" name="FM_<%=aktFMname%>timebudget_FY1" id="FM_timerbudget_FY1_<%=jobid%>_<%=aktid %>" class="jobakt_budgettimer_FY1 <%=aktclassFY1 %> form-control input-small" value="<%=rammeFY1 %>" style="width:50px;"  <%=jobDisAbled %>/></td>
        <td><input type="text" name="FM_<%=aktFMname%>timebudget_FY2" id="FM_timerbudget_FY2_<%=jobid%>_<%=aktid %>" class="jobakt_budgettimer_FY2 <%=aktclassFY2 %> form-control input-small" value="<%=rammeFY2 %>" style="width:50px;" <%=jobDisAbled %> /></td>
       
        
                <input type="hidden" name="FM_<%=aktFMname%>timebudget_FY0" value="##" />
                <input type="hidden" name="FM_<%=aktFMname%>timebudget_FY1" value="##" />
                <input type="hidden" name="FM_<%=aktFMname%>timebudget_FY2" value="##" />


        <%if cint(timesimh1h2) = 1 then %>
        <td><input type="text" style="width:40px;" disabled value="<%=rammeFY0 %>" class="form-control input-small" /></td>
        <%end if %>
        
        <%
         
            if len(trim(fctimepris)) <> 0 then
            fctimepris = formatnumber(fctimepris, 0)
            else
            fctimepris = 0
            end if  
            
            if len(trim(budgettp)) <> 0 then
            budgettp = budgettp
            else
            budgettp = 0
            end if

             if cdbl(fctimepris) > cdbl(budgettp) then
                    bgJAfc = "lightpink"
                 else
                    bgJAfc = ""
             end if %>

        <!-- H1 Ressouce timer og TP -->
        <td><input type="text" id="h1t_jobakt_<%=jobid%>_<%=aktid%>" class="<%=h1cls%> jh1 form-control input-small" value="<%=jh1%>" style="width:60px;" disabled /></td>
    
            
            <%if cint(timesimtp) = 1 then%> 
            <td style="padding-right:2px;"><input type="text" name="" class="form-control input-small" style="width:60px;" value="<%=fctimepris %>" DISABLED /></td>
            <%end if %>     


        <!-- REAL timer -->
        <td style="white-space:nowrap;"><input type="text" id="" style="width:60px;" value="<%=thisTimer %>" class="form-control input-small" DISABLED />
        <input type="hidden" id="realtimer_jobakt_<%=jobid%>_<%=aktid %>" value="<%=thisTimer%>" /></td>

           <td><input type="text" style="width:60px;" name="" id="" value="<%=gnstpTxt %>" class="form-control input-small <%=aktclassh1 %>" DISABLED />
           
                <%if aktid <> 0 then %>
                <input type="hidden" style="width:60px; background-color:<%=bgJAfc%>;" name="FM_<%=aktFMname%>fctimepris_FY0" id="fctimepris_h1_jobakt_<%=jobid%>_<%=aktid %>" value="0" class="jobakt_tph1 form-control input-small <%=aktclassh1 %>" />
                <%end if %>
                <!-- <td></td> -->

               </td>

        <td style="padding-right:2px;"><input type="text" name="" class="form-control input-small" style="width:60px;" value="<%=realTimeOmsTxt%>" DISABLED />
        <input type="hidden" id="realtimep_jobakt_<%=jobid%>_<%=aktid %>" value="<%=realTimeOms%>" /></td>

        <%if cint(timesimtp) = 100 then '*** 1??%> 
        <td style="padding-right:2px;"><input type="text" name="" class="form-control input-small" style="width:60px;" value="<%=realomsH1 %>" DISABLED /></td>
        <%end if %>
        
        <%if cint(timesimh1h2) = 1 then %>
        <!-- H2 Ressouce timer og TP -->
        <td style="white-space:nowrap;"><input type="text" id="h2tilradighed_jobakt_<%=jobid%>_<%=aktid %>" class="form-control input-small" value="<%=timerTilH2%>" style="width:60px;" disabled /></td>
        <td><input type="text" id="h2t_jobakt_<%=jobid%>_<%=aktid%>" class="<%=h2cls%> jh2 form-control input-small" value="<%=jh2%>" style="width:60px;" disabled  /></td>
        <td><input type="text" style="width:60px;" name="FM_<%=aktFMname%>fctimeprish2_FY0" id="fctimepris_h2_jobakt_<%=jobid%>_<%=aktid %>" value="<%=fctimeprish2 %>" class="jobakt_tph2 form-control input-small <%=aktclassh2 %>" /></td>
        
        <td><input type="text" id="h12t_jobakt_<%=jobid%>_<%=aktid%>" style="width:60px;" value="<%=jh1h2%>" class="form-control input-small" disabled /></td>
        <%end if %>

        <%if cint(timesimtp) = 1 then%> 
        <td style="background-color:#FFFFFF;"><input type="text" id="budget_h1_jobakt_<%=jobid%>_<%=aktid %>" class="form-control input-small" style="width:60px; border:0px;" value="<%=budgetFY0h1 %>" DISABLED /></td>
        <%end if %>        

        <%if cint(timesimh1h2) = 1 then %>
        <td style="background-color:#FFFFFF;"><input type="text" id="budget_h2_jobakt_<%=jobid%>_<%=aktid %>" class="form-control input-small" style="width:60px; border:0px; background-color:<%=bgthis%>;" value="<%=budgetFY0h2 %>" /></td>
        <%end if %>

        <td><input type="text" id="budget_jobakt_<%=jobid%>_<%=aktid %>" class="form-control input-small" style="width:80px; border:0px;" value="<%=jhGTTxt%>" DISABLED /></td><!-- budgetFY0 -->

        
                <%if aktid <> 0 then %>
                <input type="hidden" name="FM_<%=aktFMname%>fctimepris_FY0" value="##" />
                <input type="hidden" name="FM_<%=aktFMname%>fctimeprish2_FY0" value="##" />
                <%end if %>
               


    <%
end function




public progrpTot, thisPrgRealTimerGT
function medarbfelter(jobnr, jobid, aktid, h1aar, h2aar, h1md, h2md, bgttpris)

        thisMedarbRealTimerGTthisAkt = 0
        'thisMedarbRealTimerGT = 0
        h1_jobaktTot = 0
        h2_jobaktTot = 0 
        pvzb = "hidden"
        pdsp = "none"
        for p = 0 to phigh - 1



                                '** REAL timer pr. dato GT ***'
                                 select case lto
                                case "wwf"
                                visrealprdatoStartSQL = h1aar & "-7-1"
                                case else
                                visrealprdatoStartSQL = h1aar & "-1-1"
                                end select
                                
                                visrealprdatoSQL = year(visrealprdato) &"-"& month(visrealprdato) &"-"& day(visrealprdato)

                                if aktid <> 0 then
                                timerPgrpAktSQLkri = "taktivitetid = "& aktid
                                timerGrpBy = "taktivitetid"
                                else
                                timerPgrpAktSQLkri = "taktivitetid <> 0"
                                timerGrpBy = "tjobnr"
                                end if
           


                      if p = 0 then 'Total på job / akt


                       

                         '**** FORECAST IAALT FY JOBAKT
                        
                        if aktid <> 0 then
                        aktKri = " AND r.aktid = "& aktid
                        grpBy = "r.aktid"
                        else
                        aktKri = " AND r.aktid <> 0 "
                        grpBy = "r.jobid"
                        end if

                    

                         strSQLjobFcGt = "SELECT SUM(timer) AS sumtimer  "
        
                         if cint(timesimtp) = 1 then 
                         strSQLjobFcGt = strSQLjobFcGt & ", SUM(timer*6timepris) AS sumtimerBelob "
                         end if

                         strSQLjobFcGt = strSQLjobFcGt & " FROM ressourcer_md AS r "


                         if cint(timesimtp) = 1 then 
                         '"& replace(aktKri, "r.aktid", "tp.aktid") &"

                                '*** Tjekker om der findes tp på aktivitet eller om der nedarves ***'
                                '**** strSQltpfindesAkt = "SELECT from timepriser WHERE aktid = "

                                '*** ØKO SKAL der er altid være angivet en timepris pr. medarb. pr akt.  ***'
                                '*** Men til at berenge omsæ. bruges timeprisen på jobbet, da der eller kommer en grp fejl og timesum vises 1 gang pr. aktivitet (altså 7 dobbelt)
                            

                         strSQLjobFcGt = strSQLjobFcGt & " LEFT JOIN timepriser AS tp ON (tp.jobid = " & jobid &" AND tp.aktid = 0 AND tp.medarbid = r.medid)" 
                         end if


                         strSQLjobFcGt = strSQLjobFcGt & " WHERE r.jobid = " & jobid &" "& aktKri &" AND (r.medid <> 0) AND (aar = "& h1aar &" AND md = "& h1md &") GROUP BY "& grpBy 
                        
                         'if session("mid") = 1 then
                         'response.Write "<br>strSQLjobFcGt: " & strSQLjobFcGt & " - p: "& p
                         'response.flush
                         'end if                

                         jobFcGt = 0
                         jobFcGtBelob = 0 
                         oRec3.open strSQLjobFcGt, oConn, 3
                         if not oRec3.EOF then

                             jobFcGt = oRec3("sumtimer")

                             if cint(timesimtp) = 1 then 
                             jobFcGtBelob = oRec3("sumtimerBelob")
                             end if

                         end if
                         oRec3.close 

                   

                         if aktid <> 0 then
                         jobFcGtGT = jobFcGtGT + jobFcGt 
                         jobFcGtBelobGT = jobFcGtBelobGT + jobFcGtBelob
                         end if
                         


                         if jobFcGt <> 0 then 
                         jobFcGtTxt = formatnumber(jobFcGt, 0)
                         else
                         jobFcGtTxt = ""
                         end if

                         if jobFcGtBelob <> 0 then 
                         jobFcGtBelobTxt = formatnumber(jobFcGtBelob, 0) & " DKK"
                         else
                         jobFcGtBelobTxt = ""
                         end if

                        if isNull(jobFcGtBelob) <> true AND isNull(jobbgtBelobArr(jobid)) <> true then
                        jobFcGtSaldo = jobbgtBelobArr(jobid) - jobFcGtBelob
                        jobFcGtSaldo = formatnumber(jobFcGtSaldo,0)
                        else
                        jobFcGtSaldo = 0
                        end if

                        if isNull(jobFcGt) <> true AND isNull(jobbgtTimerArr(jobid)) <> true then
                        jobFcGtTimerSaldo = formatnumber(jobbgtTimerArr(jobid) - jobFcGt, 0)
                        else
                        jobFcGtTimerSaldo = 0
                        end if
                    
                         if cint(filtervisallejobvlgtmedarb) = 1 OR minit <> "" then
                         fcjobaktGTtype = "hidden"
                         else
                         fcjobaktGTtype = "text"
                         end if

                
                        if aktid = 0 then 'job


                            if (cdbl(jobFcGt) > cdbl(jobbgtTimerArr(jobid))) then
                            fcBgCol = "lightpink"
                            else
                            fcBgCol = ""
                            end if

                        else

                            if (cdbl(jobFcGt) > cdbl(aktbgtTimerArr(aktid))) then
                            fcBgCol = "lightpink"
                            else
                            fcBgCol = ""
                            end if
                        

                        end if
                         
                       '********************************************************************
                       '*** JOB AKT linje **************************************************
                       '********************************************************************
                        %>
           
                         <td style="text-align:right;"> 
                                
                                <input type="<%=fcjobaktGTtype %>" id="fcjobaktgt_<%=jobid%>_<%=aktid%>" value="<%=jobFcGtTxt%>" class="form-control input-small fcjobgt fcjobgt_<%=jobid%>" style="width:60px; background-color:<%=fcBgCol%>; text-align:right;" DISABLED />
                             
                                    <%if fcjobaktGTtype = "hidden" then %>
                                    <span style="background-color:<%=fcBgCol%>;"><%=jobFcGtTxt%></span>
                                    <%end if %>
                              
                                    <%if cint(timesimtp) = 1 then 'aktid <> 0 %>
                                    <span style="font-size:10px; padding:2px 2px 2px 2px; white-space:nowrap;" id="fcjobaktBelgts_<%=jobid%>_<%=aktid%>" class="fcjobaktBelgts"><%=jobFcGtBelobTxt %></span>
                                    <input type="hidden" id="fcjobaktBelgt_<%=jobid%>_<%=aktid%>" class="fcjobaktBelgt" value="<%=jobFcGtBelob %>"/>
                             
                                    <%end if %>
                         </td>

                        <%

                               

                                thisjobAktRealTimer = 0
                                
                                
                                if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1 OR cint(visKunFCFelter) = 2 OR cint(visKunFCFelter) = 3 then
                                ', SUM(timer*timepris) AS realoms, AVG(timepris) AS gnstp
                                'AND (tmnr = 0 "& replace(medarbIPgrp(p), "medid", "tmnr") &")
                                strSQLrealTimerJobAkt = "SELECT SUM(timer) AS sumtimer FROM timer WHERE tjobnr = '" & jobnr & "' AND "& timerPgrpAktSQLkri &"  AND (tfaktim = 1 OR tfaktim = 2) AND tdato BETWEEN '"& visrealprdatoStartSQL &"' AND '"& visrealprdatoSQL &"' GROUP BY "& timerGrpBy   
                                'response.write strSQLrealTimerJobAkt & "<br>"
                                'response.flush
                                oRec3.open strSQLrealTimerJobAkt, oConn, 3
                                if not oRec3.EOF then

                                thisjobAktRealTimer = oRec3("sumtimer")
                                'thisMedarbrealomsH1 = oRec3("realoms")
                                'thisMedarbgnstp = oRec3("gnstp")

                                end if
                                oRec3.close
                                 
                                end if
                               


                                if len(trim(thisjobAktRealTimer)) <> 0 AND thisjobAktRealTimer <> 0 then
                                thisjobAktRealTimerTxt = formatnumber(thisjobAktRealTimer, 0)
                                else
                                thisjobAktRealTimerTxt = ""
                                end if

                                thisjobAktRealTimerGT = thisjobAktRealTimerGT + thisjobAktRealTimer

                                if (cdbl(jobFcGt) < cdbl(formatnumber(thisjobAktRealTimer, 0))) then
                                fcBgCol = "lightpink"
                                else
                                fcBgCol = ""
                                end if

                                thisJobAktSaldo = 0
                                thisJobAktSaldo = (jobFcGt - thisjobAktRealTimer)
            
                                if thisJobAktSaldo <> 0 AND len(trim(thisJobAktSaldo)) <> 0 then
                                thisJobAktSaldoTxt = " = " & formatnumber(thisJobAktSaldo, 0)
                                else
                                thisJobAktSaldoTxt = ""
                                end if

            
                       %>
           
                         <td style="text-align:right;">
                            <input type="<%=fcjobaktGTtype %>" id="realjobaktgt_<%=jobid%>_<%=aktid %>" class="form-control input-small realjobgt realjobgt_<%=jobid%>" style="width:60px; background-color:<%=fcBgCol%>; text-align:right;" value="<%=thisjobAktRealTimerTxt%>" DISABLED  /> 

                             <%if fcjobaktGTtype = "hidden" then %>
                                     <span style="background-color:<%=fcBgCol%>;"><%=thisjobAktRealTimerTxt%></span>
                                    <%end if %>

                         </td>
                         <td style="text-align:left; white-space:nowrap;">
                            <%=thisJobAktSaldoTxt %>
                         </td>


                            <%if cint(timesimtp) = 1 then %>
                             <td style="text-align:right; white-space:nowrap;"> 

                             <%if aktid = 0 then%>
                              <span id="saldoTgts_<%=jobid%>_<%=aktid %>"><%=jobFcGtTimerSaldo %> t.</span>
                               <div style="font-size:10px;" id="saldoBelgts_<%=jobid%>_<%=aktid %>"><%=jobFcGtSaldo %> DKK</div>
                              <%end if %>

                             </td>

                            <%end if


                      end if 'p



                     '** Total FORECAST på gruppen '*****
                        if aktid <> 0 then
                        aktKri = " AND aktid = "& aktid
                        grpBy = "aktid"
                        else
                        aktKri = " AND aktid <> 0 "
                        grpBy = "jobid"
                        end if


                     strSQLmedrd = "SELECT SUM(timer) AS sumtimer FROM ressourcer_md WHERE jobid = " & jobid &" "& aktKri &" AND (medid = 0 "& medarbIPgrp(p) &") AND (aar = "& h1aar &" AND md = "& h1md &") GROUP BY "& grpBy 
                     'response.Write "<br>strSQLmedrd: " & strSQLmedrd & " - p: "& p
                     'response.flush
                     progrpTot = 0
                     oRec3.open strSQLmedrd, oConn, 3
                     if not oRec3.EOF then

                        progrpTot = oRec3("sumtimer")

                     end if
                     oRec3.close 

                     if progrpTot <> 0 then 
                     progrpTotTxt = formatnumber(progrpTot, 0)
                     else
                     progrpTotTxt = ""
                     end if


                                
                                
                                '************************************************************************************************
                                '** REAL timer pr. dato på Gruppen ***'
                               '************************************************************************************************

                                thisPrgRealTimer = 0
                            
                                if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 2 OR cint(visKunFCFelter) = 3 then
                                ', SUM(timer*timepris) AS realoms, AVG(timepris) AS gnstp
                                'AND (tmnr = 0 "& replace(medarbIPgrp(p), "medid", "tmnr") &")
                                strSQLrealTimerProgrp = "SELECT SUM(timer) AS sumtimer FROM timer WHERE tjobnr = '" & jobnr & "' AND "& timerPgrpAktSQLkri &" AND (tmnr = 0 "& replace(medarbIPgrp(p), "medid", "tmnr") &") AND (tfaktim = 1 OR tfaktim = 2) AND tdato BETWEEN '"& visrealprdatoStartSQL &"' AND '"& visrealprdatoSQL &"' GROUP BY "& timerGrpBy   
                                'response.write strSQLrealTimerAktmedarb & "<br>"
                                'response.flush
                                oRec3.open strSQLrealTimerProgrp, oConn, 3
                                if not oRec3.EOF then

                                thisPrgRealTimer = oRec3("sumtimer")
                                'thisMedarbrealomsH1 = oRec3("realoms")
                                'thisMedarbgnstp = oRec3("gnstp")

                                end if
                                oRec3.close

                                end if

                                if len(trim(thisPrgRealTimer)) <> 0 AND thisPrgRealTimer <> 0 then
                                thisPrgRealTimerTxt = formatnumber(thisPrgRealTimer, 0)
                                else
                                thisPrgRealTimerTxt = ""
                                end if

                                if aktid <> 0 then
                                thisPrgRealTimerGT = thisPrgRealTimerGT + thisPrgRealTimer
                                end if    

                                thisPrgSaldo = 0
                                thisPrgSaldo = (progrpTot - thisPrgRealTimer)


                                if len(trim(thisPrgSaldo)) <> 0 AND thisPrgSaldo <> 0 then
                                thisPrgSaldoTxt = " = "& formatnumber(thisPrgSaldo, 0)
                                             
                                else
                                thisPrgSaldoTxt = ""
                                end if

                              if thisPrgSaldo < 0 then
                                 salBgCol = "lightpink"
                               else
                                 salBgCol = ""
                               end if
                                
            
                                 if aktid = 0 then 'job


                                    if (cdbl(progrpTot) > cdbl(jobbgtTimerArr(jobid))) then
                                    fcBgCol = "lightpink"
                                    else
                                    fcBgCol = ""
                                    end if

                                else

                                    if (cdbl(progrpTot) > cdbl(aktbgtTimerArr(aktid))) then
                                    fcBgCol = "lightpink"
                                    else
                                    fcBgCol = ""
                                    end if
                        

                                end if


            
            '*** PROJEKTGRUPPE ********************%>

             
               
            <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1 then %>
             <td><input type="text" id="afd_jobakt_<%=jobid%>_<%=aktid %>_<%=p %>" class="form-control input-small afd_jobakt afd_jobakt_<%=p %> afd_jobakt_<%=jobid%>_<%=aktid %>" style="width:60px; text-align:right; background-color:<%=fcBgCol%>;" value="<%=progrpTotTxt %>" DISABLED />
            
                 <%if cint(timesimtp) = 1 then %>
                 <span id="afd_jobaktbels_<%=jobid%>_<%=aktid %>_<%=p %>" class="afd_jobaktbels afd_jobaktbels_<%=p %> afd_jobaktbels_<%=jobid%>_<%=aktid %>" style="text-align:left; font-size:10px; display:none; visibility:hidden;">0 DKK</span>
                 <input type="hidden" id="afd_jobaktbel_<%=jobid%>_<%=aktid %>_<%=p %>" class="afd_jobaktbel afd_jobaktbel_<%=p %> afd_jobaktbel_<%=jobid%>_<%=aktid %>" value="0" />

                 <%end if %>

             </td>
            <%end if %>
             
            <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 2 then %>
            <td style="text-align:right;">
            <%=thisPrgRealTimerTxt %>
            </td>
            <%end if

            if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 3 then%>
                 <td style="text-align:left; background-color:<%=salBgCol%>;"><%=thisPrgSaldoTxt %></td>
             <%end if
            
           


            if (cint(visprogrpid) = 0) AND cint(filtervisallejobvlgtmedarb) = 1 then
                
            else     
                
            for m = 0 TO mhigh - 1 
                
                
                
                
                if antalp(p) = antalm(m,0) then

                     h1 = 0
                     h2 = 0
                     h1h2 = 0
                
                         strSQLmedrd = "SELECT SUM(timer) AS sumtimer FROM ressourcer_md WHERE jobid = " & jobid & " AND aktid = "& aktid &" AND medid = "& antalm(m,1) &" AND aar = "& h1aar &" AND md = "& h1md  &" GROUP BY medid" 
                         'response.Write "strSQLmedrd: " & strSQLmedrd
                         'response.flush
                
                         oRec3.open strSQLmedrd, oConn, 3
                         if not oRec3.EOF then

                            h1 = oRec3("sumtimer")

                         end if
                         oRec3.close


                        useH2 = 900
                        if cint(useH2) = 1 then
                            strSQLmedrd = "SELECT SUM(timer) AS sumtimer FROM ressourcer_md WHERE jobid = " & jobid & " AND aktid = "& aktid &" AND medid = "& antalm(m,1) &" AND aar = "& h2aar &" AND md = "& h2md &" GROUP BY medid " 
                             oRec3.open strSQLmedrd, oConn, 3
                             if not oRec3.EOF then

                                h2 = oRec3("sumtimer")

                             end if
                             oRec3.close
                        end if


                h1h2 = h1+h2
                
               


                if len(trim(h1)) <> 0 then 
                h1_jobaktTot = h1_jobaktTot + h1   
                else
                h1_jobaktTot = h1_jobaktTot
                end if  

                if len(trim(h2)) <> 0 then 
                h2_jobaktTot = h2_jobaktTot + h2   
                else
                h2_jobaktTot = h2_jobaktTot
                end if  

                if aktid <> 0 then
                h1_medTot(antalm(m,1)) = h1_medTot(antalm(m,1))/1 + h1
                h2_medTot(antalm(m,1)) = h2_medTot(antalm(m,1))/1 + h2
                end if

                if len(trim(h1h2)) <> 0 AND h1h2 <> 0 then
                h1h2 = h1h2 'formatnumber(h1h2, 2)
                else
                h1h2 = ""
                end if

                h1Val = h1

                if len(trim(h1)) <> 0 AND h1 <> 0 then
                h1 = h1 'formatnumber(h1, 2)
                else
                h1 = ""
                end if

                if len(trim(h2)) <> 0 AND h2 <> 0 then
                h2 = h2 'formatnumber(h2, 2)
                else
                h2 = ""
                end if


               


                if aktid <> 0 then
                mh1jacls = "mh1h_akt_"& jobid & "_"& antalm(m,1)
                mh2jacls = "mh2h_akt_"& jobid & "_"& antalm(m,1)
                mt1jacls = "mh1t_akt_"& jobid & "_"& antalm(m,1)
                else
                mh1jacls = "mh1h_job_"& jobid & "_"& antalm(m,1)
                mh2jacls = "mh2h_job_"& jobid & "_"& antalm(m,1)
                mt2jacls = "mh2t_job_"& jobid & "_"& antalm(m,1)
                end if

               
                 
                    '*********** Timepriser ******
                    medarbTp = 0
                    tpFundetAkt = 0
                    if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1 then

                        if cint(timesimtp) = 1 then
                
                        strSQLmedtp = "SELECT 6timepris FROM timepriser WHERE jobid = " & jobid & " AND aktid = "& aktid &" AND medarbid = "& antalm(m,1)  
                         oRec3.open strSQLmedtp, oConn, 3
                         if not oRec3.EOF then

                            tpFundetAkt = 1
                            medarbTp = oRec3("6timepris")

                         end if
                         oRec3.close

                        end if

                        if cint(tpFundetAkt) = 0 then 'Nedarver
                        strSQLmedtp = "SELECT 6timepris FROM timepriser WHERE jobid = " & jobid & " AND aktid = 0 AND medarbid = "& antalm(m,1)  
                         oRec3.open strSQLmedtp, oConn, 3
                         if not oRec3.EOF then

                       
                            medarbTp = oRec3("6timepris")

                         end if
                         oRec3.close

                        end if

                    end if

                    
                    '** REAL timer pr. dato ***'
                    if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 2 OR cint(visKunFCFelter) = 3 then


                     select case lto
                    case "wwf"
                    visrealprdatoStartSQL = h1aar & "-7-1"
                    case else
                    visrealprdatoStartSQL = h1aar & "-1-1"
                    end select
                   
                    visrealprdatoSQL = year(visrealprdato) &"-"& month(visrealprdato) &"-"& day(visrealprdato)

                    thisMedarbRealTimer = 0
                    strSQLrealTimerAktmedarb = "SELECT SUM(timer) AS sumtimer FROM timer WHERE tjobnr = '" & jobnr & "' AND taktivitetid = "& aktid &" AND tmnr = "& antalm(m,1) &" AND (tfaktim = 1 OR tfaktim = 2) AND tdato BETWEEN '"& visrealprdatoStartSQL &"' AND '"& visrealprdatoSQL &"' GROUP BY tmnr"   
                    'response.write strSQLrealTimerAktmedarb & "<br>"
                    'response.flush
                    oRec3.open strSQLrealTimerAktmedarb, oConn, 3
                    if not oRec3.EOF then

                    thisMedarbRealTimer = oRec3("sumtimer")
                   
                    end if
                    oRec3.close

                    end if

                    thisMedarbRealTimerGT(antalm(m,1)) = thisMedarbRealTimerGT(antalm(m,1)) + thisMedarbRealTimer
                    thisMedarbRealTimerGTthisAkt = thisMedarbRealTimerGTthisAkt + thisMedarbRealTimer
                    

                
                    if thisMedarbRealTimer <> 0 then
                    thisMedarbRealTimerTxt = formatnumber(thisMedarbRealTimer, 0)
                    else
                    thisMedarbRealTimerTxt = ""
                    end if

                    thisMedarbSaldo = (h1Val - thisMedarbRealTimer)
                    if cdbl(thisMedarbSaldo) <> 0 then
                    thisMedarbSaldoTxt = formatnumber(thisMedarbSaldo, 0)
                    else
                    thisMedarbSaldoTxt = ""
                    end if 

                    mBelobH1tp = formatnumber(h1Val*medarbTp,0)
                    if mBelobH1tp <> 0 then
                    mBelobH1tpTxt = mBelobH1tp & " DKK"
                    else
                    mBelobH1tpTxt = ""
                    end if

                
                        if aktid = 0 then
                        'h1h2Disabled = "" '"disabled"
                        h1h2type = "hidden"
                        h1h2Maxl = 0
                        h1h2BGcol = "#CCCCCC"
                        else
                        h1h2Maxl = 10
                        h1h2type = "text"
                        'h1h2Disabled = ""
                        h1h2BGcol = "#ffffff"
                        end if

                %>


            <input type="hidden" name="FM_medid" value="<%=antalm(m,1) %>" />

            <%if cdbl(medarbTp) > cdbl(bgttpris) then
                bgHfc = "" '"lightpink"
             else
                bgHfc = ""
             end if
                
                
            if cdbl(thisMedarbRealTimer) > cdbl(h1Val) then
                 bgTotfc = "Lightpink"
            else
                 bgTotfc = "#Eff3ff"
            end if %>

            

            <%if cint(timesimtp) = 1 AND (cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1) then 
                
                select case lto
                case "oko", "intranet - local"
                tpmedarbDisabled = "DISABLED"
                case else
                tpmedarbDisabled = "DISABLED"
                end select
            
             %>
            <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input name="FM_tp" type="<%=h1h2type %>" value="<%=formatnumber(medarbTp,0) %>" style="width:45px; background-color:<%=bgHfc%>;" id="mh1t_jobaktmid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1)%>" class="<%=mt1jacls%> form-control input-small mh1t mh1t_jobaktmid_<%=jobid%>_<%=aktid%>" maxlength="<%=h1h2Maxl %>" <%=tpmedarbDisabled %> /></td>
            <%end if %>


            <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1 then %>
            <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;">
                <input type="<%=h1h2type %>" name="FM_H1" id="mh1h_jobaktmid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1)%>" class="<%=mh1jacls%> form-control input-small mh1 mh1h_jobaktmid_<%=jobid%>_<%=aktid%> mh1h_jobaktmid_<%=jobid%>_<%=aktid%>_<%=p %>" style="width:45px; background-color:<%=h1h2BGcol%>;" value="<%=h1 %>" maxlength="<%=h1h2Maxl %>" />
            </td>
           <%else %>
            <input type="hidden" name="FM_H1" value="<%=h1 %>" /> 
           <%end if %>


                <input type="hidden" name="" id="mh1h_jobaktafd_<%=jobid%>_<%=aktid %>_<%=antalm(m,1) %>" value="<%=p %>" />
                <input type="hidden" name="FM_H1_jobid" id="h1_jobid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1) %>" value="<%=jobid %>" />
                <input type="hidden" name="FM_H1_aktid" id="h1_aktid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1) %>" value="<%=aktid %>" />
           
                

            <%if (cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 2) then %>
            <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>; text-align:right;">
                
               
                 <input type="<%=h1h2type %>" name="" id="" class="form-control input-small mreal" style="width:45px;" value="<%=thisMedarbRealTimerTxt %>" DISABLED />
                 <input type="hidden" name="" id="mreal_jobaktmid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1)%>" class="<%=mrealjacls%> mreal_jobaktmid_<%=jobid%>_<%=aktid%>" value="<%=thisMedarbRealTimer%>"/>
                
                <!--
                <input type="hidden" name="FM_H2" id="mh2h_jobaktmid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1)%>" class="<%=mh2jacls%> form-control input-small mh2 mh2h_jobaktmid_<%=jobid%>_<%=aktid%>" style="width:45px; background-color:<%=h1h2BGcol%>;" value="<%=h2 %>" maxlength="<%=h1h2Maxl %>" />
            
                <input type="hidden" name="FM_H2_jobid" id="h2_jobid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1) %>" value="<%=jobid %>" />
                <input type="hidden" name="FM_H2_aktid" id="h2_aktid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1) %>" value="<%=aktid %>" />
                -->

            </td>    
            <%end if %>

             <%if (cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 3) then %>
            <td class="afd_p_<%=p%> afd_p" style="white-space:nowrap; visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="<%=h1h2type %>" name="FM_H1H2" id="mh12h_jobaktmid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1) %>" class="form-control input-small fs_<%=jobid%>_<%=aktid %>_<%=p %>" style="width:45px; background-color:<%=bgTotfc%>;" value="<%=thisMedarbSaldoTxt%>" disabled />

                    <%if cint(timesimtp) = 1 then%>     
                    <input type="hidden" id="mh12tp_jobaktmid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1) %>" class="fs_<%=jobid%>_<%=aktid %>_<%=p %> mbel_<%=jobid%>_<%=aktid %>" value="<%=mBelobH1tp %>" />
                    <span id="mh12tps_jobaktmid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1) %>" class="fss_<%=jobid%>_<%=aktid %>_<%=p %> mbels_<%=jobid%>_<%=aktid %>" style="font-size:10px; white-space:nowrap;"><%=mBelobH1tpTXT%></span>
                    <%end if %>
            </td>
            <%end if %>
            

            <%
               
                  
            end if

            response.flush      
            next 'm
               
            end if'Visprogrp og vis alle job%>
            
            
        <%
         response.flush   
         next 'p%>
        
         <input type="hidden" id="real_jobaktmid_<%=jobid%>_<%=aktid %>" class="realjobaktgt" value="<%=thisMedarbRealTimerGTthisAkt%>" />
         <input type="hidden" id="h1h_jobaktmid_<%=jobid%>_<%=aktid %>" value="<%=h1_jobaktTot%>" />
         <!--BB:<input type="text" id="h1t_jobaktmid_<%=jobid%>_<%=aktid %>" value="0" />-->
         <input type="hidden" id="h2h_jobaktmid_<%=jobid%>_<%=aktid %>" value="<%=h2_jobaktTot%>" />

           

           

    <%
end function    
    


function opdaterRessouceRamme(f, FY0, FY1, FY2, jobBudgetFY0, fctimeprisFY0, fctimeprish2FY0, jobBudgetFY1, jobBudgetFY2, jobids, aktid)


        fctimepris = 0

        select case f
        case 0 
        FYuse = FY0
        Tiuse = trim(jobBudgetFY0)
        fctimepris = trim(fctimeprisFY0)
        fctimeprish2 = trim(fctimeprish2FY0)

        fctimepris = replace(fctimepris, ".", "")
        fctimepris = replace(fctimepris, ",", ".")

        
        fctimepris = replace(fctimepris, "#","")
        
        if len(trim(fctimepris)) <> 0 then
        fctimepris = fctimepris
        else
        fctimepris = 0
        end if

        fctimeprish2 = replace(fctimeprish2, ".", "")
        fctimeprish2 = replace(fctimeprish2, ",", ".")

        case 1
        FYuse = FY1
        Tiuse = trim(jobBudgetFY1)
        case 2
        FYuse = FY2
        Tiuse = trim(jobBudgetFY2)
        end select

        Tiuse = replace(Tiuse, ".", "")
        Tiuse = replace(Tiuse, ",", ".")

        if len(trim(Tiuse)) <> 0 then
        

            'response.write "Tiuse: " & Tiuse & "<br>"
            'response.flush

            jobBudgetRammeFindes = 0
            strSQLjobRamme = "SELECT id FROM ressourcer_ramme WHERE jobid = "& jobids &" AND medid = 0 AND aktid = "& aktid &" AND aar = "& FYuse
            
            'response.write strSQLjobRamme
            'response.flush
            oRec3.open strSQLjobRamme, oConn, 3
            if not oRec3.EOF then

                if cint(Tiuse) = 0 then '** Delete
                
                strSQlrammeDel = "DELETE FROM ressourcer_ramme  WHERE id = " & oRec3("id")
                
                'response.write strSQlrammeDel
                'response.flush
                oConn.execute(strSQlrammeDel)    

                else

                jobBudgetRammeFindes = oRec3("id")
                end if
        
            end if 
            oRec3.close

           


            if cdbl(jobBudgetRammeFindes) <> 0 then

                if f = 0 then 'opdater timepris aktuelt budget/forecast
                strSQLupdateJobRamme = "UPDATE ressourcer_ramme SET timer = " & Tiuse & ", fctimepris = "& fctimepris &", fctimeprish2 = "& fctimeprish2 &" WHERE id = " & jobBudgetRammeFindes
                else
                strSQLupdateJobRamme = "UPDATE ressourcer_ramme SET timer = " & Tiuse & " WHERE id = " & jobBudgetRammeFindes
                end if

           
            else
                if f = 0 then 'opdater timepris aktuelt budget/forecast
                strSQLupdateJobRamme = "INSERT INTO ressourcer_ramme (timer, jobid, aktid, medid, aar, fctimepris, fctimeprish2) VALUES ("& Tiuse &", "& jobids &", "& aktid &", 0, "& FYuse &", "& fctimepris &", "& fctimeprish2 &")"
                else
                strSQLupdateJobRamme = "INSERT INTO ressourcer_ramme (timer, jobid, aktid, medid, aar, fctimepris, fctimeprish2) VALUES ("& Tiuse &", "& jobids &", "& aktid &", 0, "& FYuse &", 0, 0)" 
                end if
            end if

            'response.write strSQLupdateJobRamme & "<br><br>"
            'response.flush
            oConn.execute(strSQLupdateJobRamme)


          
        end if'tiuse

       

end function 
    
%>
