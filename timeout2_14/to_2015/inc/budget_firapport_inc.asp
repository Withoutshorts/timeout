<%
public budgetFY0GT
function jobaktbudgetfelter(jobnr, jobid, aktid, h1aar, h2aar, h1md, h2md)

    jh1 = 0
    jh2 = 0
    jh1h2 = 0
    timerTilH2 = 0

     if aktid <> 0 then
    'SUM(timer*timepris) AS realoms, AVG(timepris) AS gnstp
    'thisTimer = oRec("sumtimer")
    strSQLAktKri = " AND taktivitetid = "& aktid
    strSQLAktResKri = " AND aktid = "& aktid
    else
    strSQLAktKri = ""
    strSQLAktResKri = " AND aktid = 0 "
    end if

    'if lcase(sogVal) = "all" OR aktid = 0 then 'berenger sum af forecast total på job

        
                    
                
                     strSQLmedrd = "SELECT SUM(timer) AS sumtimer FROM ressourcer_md WHERE jobid = " & jobid & " "& strSQLAktResKri &" AND aar = "& h1aar &" AND md = "& h1md  & " GROUP BY jobid" 
                     'response.Write "strSQLmedrd: " & strSQLmedrd
                     'response.flush
                
                        oRec3.open strSQLmedrd, oConn, 3
                     if not oRec3.EOF then

                        jh1 = oRec3("sumtimer")

                     end if
                     oRec3.close

                    strSQLmedrd = "SELECT SUM(timer) AS sumtimer FROM ressourcer_md WHERE jobid = " & jobid & " "& strSQLAktResKri &" AND aar = "& h2aar &" AND md = "& h2md & " GROUP BY jobid"   
                     oRec3.open strSQLmedrd, oConn, 3
                     if not oRec3.EOF then

                        jh2 = oRec3("sumtimer")

                     end if
                     oRec3.close

               

                  

    'end if


    '*** Realtimer ***'
    visrealprdatoStartSQL = h1aar & "-7-1"
    visrealprdatoSQL = year(visrealprdato) &"-"& month(visrealprdato) &"-"& day(visrealprdato)
    realomsH1 = 0
    gnstp = 0

    

 

                     strSQLrealTimerJob = "SELECT SUM(timer) AS sumtimer, SUM(timer*timepris) AS realoms, AVG(timepris) AS gnstp FROM timer WHERE tjobnr = '" & jobnr & "' "& strSQLAktKri &" AND tdato BETWEEN '"& visrealprdatoStartSQL &"' AND '"& visrealprdatoSQL &"' GROUP BY tjobnr"   
                     oRec3.open strSQLrealTimerJob, oConn, 3
                     if not oRec3.EOF then

                        thisTimer = oRec3("sumtimer")
                        realomsH1 = oRec3("realoms")
                        gnstp = oRec3("gnstp")

                     end if
                     oRec3.close

                    

     gnstp = formatnumber(gnstp, 2)
     realomsH1 = realomsH1
     jh1h2 = jh1+jh2
     timerTilH2 = formatnumber((jh1 - (thisTimer)), 0)




                    '** RAMME FY ***'
                     rammeFY0 = 0
                    fctimepris = 0
                    fctimeprish2 = 0 
                    strSQLrammeFY0 = "SELECT timer, fctimepris, fctimeprish2 FROM ressourcer_ramme WHERE jobid = " & jobid & " AND aktid = "& aktid &" AND medid = 0 AND aar = "& year(Y0)    

                    'response.write strSQLrammeFY0
                    'response.flush

                     oRec3.open strSQLrammeFY0, oConn, 3
                     if not oRec3.EOF then

                        rammeFY0 = oRec3("timer")
                        fctimepris = oRec3("fctimepris")
                        fctimeprish2 = oRec3("fctimeprish2")

                     end if
                     oRec3.close




                     '** RAMME FY1 ***'
                     rammeFY1 = 0
                     strSQLrammeFY1 = "SELECT timer FROM ressourcer_ramme WHERE jobid = " & jobid & " AND aktid = "& aktid &" AND medid = 0 AND aar = "& year(Y1)   

                    'response.write strSQLrammeFY1
                    'response.flush

                     oRec3.open strSQLrammeFY1, oConn, 3
                     if not oRec3.EOF then
                          rammeFY1 = oRec3("timer")
                     end if
                     oRec3.close


                    '** RAMME FY2 ***'
                     rammeFY2 = 0
                     strSQLrammeFY2 = "SELECT timer FROM ressourcer_ramme WHERE jobid = " & jobid & " AND aktid = "& aktid &" AND medid = 0 AND aar = "& year(Y2)    

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

        if (thisTimer = 0) then
        budgetFY0 = formatnumber(budgetFY0h1/1, 0)
        else 
        budgetFY0 = formatnumber(budgetFY0h2/1, 0)
        end if


        realomsH1 = formatnumber(realomsH1,0)

      if aktid = 0 then
      budgetFY0GT = budgetFY0GT + budgetFY0 
      end if


    if aktid <> 0 then

    aktFMname = "akt"

    aktclassFY0 = "jobakt_fy0_"& jobid
    aktclassFY1 = "jobakt_fy1_"& jobid
    aktclassFY2 = "jobakt_fy2_"& jobid

    aktclassh1 = "jobakt_tph1_"& jobid
   
    aktclassh2 = "jobakt_tph2_"& jobid

    else

    aktFMname = ""

    end if


    if (rammeFY0 < jh1h2) then
    bgthisFY0 = "red"
    else
    bgthisFY0 = ""
    end if

    
    if aktid <> 0 then
    h1cls = "jh1_"& jobid
    h2cls = "jh1_"& jobid
    else
    h1cls = "x"
    h2cls = "x"
    end if
    
    %>
      <td><input type="text" name="FM_<%=aktFMname%>timebudget_FY0" id="FM_timerbudget_FY0_<%=jobid%>_<%=aktid %>" class="jobakt_budgettimer_FY <%=aktclassFY0 %>" value="<%=rammeFY0 %>" style="width:40px; background-color:<%=bgthisFY0%>;" /></td>
     <td><input type="text" name="FM_<%=aktFMname%>timebudget_FY1" id="FM_timerbudget_FY1_<%=jobid%>_<%=aktid %>" class="jobakt_budgettimer_FY1 <%=aktclassFY1 %>" value="<%=rammeFY1 %>" style="width:40px;" /></td>
        <td><input type="text" name="FM_<%=aktFMname%>timebudget_FY2" id="FM_timerbudget_FY2_<%=jobid%>_<%=aktid %>" class="jobakt_budgettimer_FY2 <%=aktclassFY2 %>" value="<%=rammeFY2 %>" style="width:40px;" /></td>
        <td><input type="text" style="width:40px;" disabled value="<%=rammeFY0 %>" /></td>
        <td><input type="text" style="width:60px;" name="FM_<%=aktFMname%>fctimepris_FY0" id="fctimepris_h1_jobakt_<%=jobid%>_<%=aktid %>" value="<%=fctimepris %>" class="jobakt_tph1  <%=aktclassh1 %>" /></td>
        <td><input type="text" id="h1t_jobakt_<%=jobid%>_<%=aktid%>" class="<%=h1cls%> jh1" value="<%=jh1%>" style="width:40px;" disabled /></td>
        <td style="white-space:nowrap;"><input type="text" id="" style="width:40px;" value="<%=thisTimer %>" disabled />
        <input type="hidden" id="realtimer_jobakt_<%=jobid%>_<%=aktid %>" value="<%=thisTimer%>" /></td>
        <td align="right" style="padding-right:2px;"><%=gnstp %><input type="hidden" id="realtimep_jobakt_<%=jobid%>_<%=aktid %>" value="<%=gnstp%>" /></td>
        <td align="right" style="padding-right:2px;"><%=realomsH1 %></td>
        <td style="white-space:nowrap;"><input type="text" id="h2tilradighed_jobakt_<%=jobid%>_<%=aktid %>" value="<%=timerTilH2%>" style="width:40px;" disabled /></td>
        <td><input type="text" style="width:60px;" name="FM_<%=aktFMname%>fctimeprish2_FY0" id="fctimepris_h2_jobakt_<%=jobid%>_<%=aktid %>" value="<%=fctimeprish2 %>" class="jobakt_tph2 <%=aktclassh2 %>" /></td>
        <td><input type="text" id="h2t_jobakt_<%=jobid%>_<%=aktid%>" class="<%=h2cls%> jh2" value="<%=jh2%>" style="width:40px;" disabled  /></td>
        <td><input type="text" id="h12t_jobakt_<%=jobid%>_<%=aktid%>" style="width:40px;" value="<%=jh1h2%>" disabled /></td>
        <td style="background-color:#FFFFFF;"><input type="text" id="budget_h1_jobakt_<%=jobid%>_<%=aktid %>" style="width:60px; border:0px; background-color:<%=bgthis%>;"" value="<%=budgetFY0h1 %>" /></td>
        <td style="background-color:#FFFFFF;"><input type="text" id="budget_h2_jobakt_<%=jobid%>_<%=aktid %>" style="width:60px; border:0px; background-color:<%=bgthis%>;"" value="<%=budgetFY0h2 %>" /></td>
        <td style="background-color:lightpink;"><input type="text" id="budget_jobakt_<%=jobid%>_<%=aktid %>" style="width:60px; border:0px; background-color:lightpink;" value="<%=budgetFY0 %>" /></td>

    <%
end function





function medarbfelter(jobid, aktid, h1aar, h2aar, h1md, h2md)

    %>



                    
        <%
        h1_jobaktTot = 0
        h2_jobaktTot = 0 
        pvzb = "hidden"
        pdsp = "none"
        for p = 0 to phigh - 1
            
            
            
            
            %>
           
             <td><input type="text" id="afd_jobakt_<%=jobid%>_<%=aktid %>_<%=p %>" style="width:60px;" disabled /></td>
            
            <%for m = 0 TO mhigh - 1 
                
                if antalp(p) = antalm(m,0) then

                     h1 = 0
                     h2 = 0
                     h1h2 = 0
                
                     strSQLmedrd = "SELECT SUM(timer) AS sumtimer FROM ressourcer_md WHERE jobid = " & jobid & " AND aktid = "& aktid &" AND medid = "& antalm(m,1) &" AND aar = "& h1aar &" AND md = "& h1md  &" GROUP BY jobid, aktid, medid" 
                     'response.Write "strSQLmedrd: " & strSQLmedrd
                     'response.flush
                
                        oRec3.open strSQLmedrd, oConn, 3
                     if not oRec3.EOF then

                        h1 = oRec3("sumtimer")

                     end if
                     oRec3.close

                    strSQLmedrd = "SELECT SUM(timer) AS sumtimer FROM ressourcer_md WHERE jobid = " & jobid & " AND aktid = "& aktid &" AND medid = "& antalm(m,1) &" AND aar = "& h2aar &" AND md = "& h2md  &" GROUP BY jobid, aktid, medid " 
                     oRec3.open strSQLmedrd, oConn, 3
                     if not oRec3.EOF then

                        h2 = oRec3("sumtimer")

                     end if
                     oRec3.close

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
                else
                mh1jacls = "mh1h_job_"& jobid & "_"& antalm(m,1)
                mh2jacls = "mh2h_job_"& jobid & "_"& antalm(m,1)
                end if

                'if aktid = 0 then 
                'GT1clas = "h1h_jobaktmidGT_"& jobid
                'GT2clas = "h2h_jobaktmidGT_"& jobid
                'else
                'GT1clas = "h1h_jobaktmidGT_"& jobid
                'GT2clas = "h2h_jobaktmidGT_"& jobid
                'end if %>


            <input type="hidden" name="FM_medid" value="<%=antalm(m,1) %>" />
            <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" name="FM_H1" id="mh1h_jobaktmid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1)%>" class="<%=mh1jacls%> mh1 mh1h_jobaktmid_<%=jobid%>_<%=aktid%>" style="width:40px;" value="<%=h1 %>" />
                <input type="hidden" name="FM_H1_jobid" id="h1_jobid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1) %>" value="<%=jobid %>" />
                <input type="hidden" name="FM_H1_aktid" id="h1_aktid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1) %>" value="<%=aktid %>" />
            </td>
            <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" name="FM_H2" style="width:40px;" id="mh2h_jobaktmid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1)%>" class="<%=mh2jacls%> mh2 mh2h_jobaktmid_<%=jobid%>_<%=aktid%>" value="<%=h2 %>" />
                <input type="hidden" name="FM_H2_jobid" id="h2_jobid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1) %>" value="<%=jobid %>" />
                <input type="hidden" name="FM_H2_aktid" id="h2_aktid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1) %>" value="<%=aktid %>" /></td>    
            <td class="afd_p_<%=p%> afd_p" style="white-space:nowrap; visibility:<%=pvzb%>; display:<%=pdsp%>;">= <input type="text" name="FM_H1H2" id="mh12h_jobaktmid_<%=jobid%>_<%=aktid %>_<%=antalm(m,1) %>" class="fs_<%=jobid%>_<%=aktid %>_<%=p %>" style="width:40px;" value="<%=h1h2%>" disabled /></td>
            
            <%
               
                  
            end if

            response.flush      
            next 'm%>
            
            
        <%
         response.flush   
         next 'p%>


           

            <!--totaler på job / akt alle medarbejdere -->
            <input type="hidden" id="h1h_jobaktmid_<%=jobid%>_<%=aktid %>" value="<%=h1_jobaktTot%>" />
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

            'response.write strSQLupdateJobRamme
            'response.flush
            oConn.execute(strSQLupdateJobRamme)


          
        end if'tiuse

       

end function 
    
%>
