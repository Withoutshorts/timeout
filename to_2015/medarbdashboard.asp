

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--include file="inc/budget_firapport_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->


<script type="text/javascript" src="js/libs/excanvas.compiled.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.js"></script>

<!--
<script type="text/javascript" src="js/plugins/flot/jquery.flot.orderBars.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.resize.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.stack.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.tooltip.min.js"></script>
<script type="text/javascript" src="js/demos/flot/stacked-vertical.js"></script>
<script type="text/javascript" src="js/demos/flot/vertical.js"></script>
<script type="text/javascript" src="js/demos/flot/scatter.js"></script>
<script type="text/javascript" src="js/demos/flot/line.js"></script>
 -->
<!--
<script type="text/javascript" src="js/demos/flot/donut_saelger2.js"></script>
-->
<script type="text/javascript" src="js/demos/flot/donut.js"></script>

<script type="text/javascript" src="js/plugins/flot/jquery.flot.pie.js"></script>


<%call menu_2014 %>


<% if len(trim(request("FM_periode"))) <> 0 then
   usePeriode = request("FM_periode")
   else
   usePeriode = 1 'måned
   end if 


   
    

    usePeriodeSel0 = ""
    usePeriodeSel1 = ""
    usePeriodeSel2 = ""
    usePeriodeSel3 = ""

    select case usePeriode
    case 0
    usePeriodeSel0 = "CHECKED"
    case 1
    usePeriodeSel1 = "CHECKED"
    case 2
    usePeriodeSel2 = "CHECKED"
    case 3
    usePeriodeSel3 = "CHECKED"
    case else
    usePeriodeSel0 = "CHECKED"
    end select

    
    if len(trim(request("FM_aar"))) <> 0 then
    useYear = request("FM_aar")
    else
    useYear = year(now)
    end if
    
    if len(trim(request("FM_mth"))) <> 0 then
    useMonth = request("FM_mth")
    else
    useMonth = month(now)
    end if

    mthSel1 = ""
    mthSel2 = ""
    mthSel3 = ""
    mthSel4 = ""
    mthSel5 = ""
    mthSel6 = ""
    mthSel7 = ""
    mthSel8 = ""
    mthSel9 = ""
    mthSel10 = ""
    mthSel11 = ""
    mthSel12 = ""

    select case useMonth
    case 1
    mthSel1 = "SELECTED"
    case 2
    mthSel2 = "SELECTED"
    case 3
    mthSel3 = "SELECTED"
    case 4
    mthSel4 = "SELECTED"
    case 5
    mthSel5 = "SELECTED"
    case 6
    mthSel6 = "SELECTED"
    case 7
    mthSel7 = "SELECTED"
    case 8
    mthSel8 = "SELECTED"
    case 9
    mthSel9 = "SELECTED"
    case 10
    mthSel10 = "SELECTED"
    case 11
    mthSel11 = "SELECTED"
    case 12
    mthSel12 = "SELECTED"
    end select

    if len(trim(request("FM_kv"))) <> 0 then
    useQuater = request("FM_kv")
    else
    useQuater = 1
    end if

    useQuaterSel1 = ""
    useQuaterSel2 = ""
    useQuaterSel3 = ""
    useQuaterSel4 = ""

    select case useQuater
    case 1
    useQuaterSel1 = "SELECTED"
    case 2
    useQuaterSel2 = "SELECTED"
    case 3
    useQuaterSel3 = "SELECTED"
    case 4
    useQuaterSel4 = "SELECTED"
    case else
    useQuaterSel1 = "SELECTED"
    end select

    
    if len(trim(request("FM_medarb"))) <> 0 then
    usemrn = request("FM_medarb")
    else
    usemrn = session("mid") 
    end if

    if len(trim(request("FM_progrp"))) <> 0 then
    pgrp = request("FM_progrp")
    else
    pgrp = 10
    end if


    jobidsStr = " AND (j.id <> 0 "
    if len(trim(request("FM_job"))) <> 0 then
    jobid = request("FM_job")
        
            if instr(jobid, ",") <> 0 then
            jobidsStr = " AND (j.id = 0 "
            jobids = split(request("FM_job"), ", ") 

                for d = 0 to UBOUND(jobids)

                    jobidsStr = jobidsStr & " OR j.id = " & jobids(d)
                
                next

            '** Nulstil jobid
            jobid = 0
            end if
    
    else
    jobid = 0
    end if

    jobidsStr = jobidsStr & ")"

    select case usePeriode
    case 1
    maksHoursBar = 185
    case 2
    maksHoursBar = 185 * 3
    case 3
    maksHoursBar = 185 * 12
    end select
    %>

<div class="wrapper">
    <div class="content">
    
    <div class="container">
 
      

        <div class="portlet">
        <h3 class="portlet-title"><u>Medarbejder Dashboard</u></h3>
        

        <div class="well">
              <form action="medarbdashboard.asp?sogsubmitted=1" method="POST">
                <div class="row">
                    <div class="col-lg-2">
                    <h4 class="panel-title-well">Søgefilter</h4></div>
                    <div class="col-lg-10">&nbsp</div>
                 
                    </div>
                  
                    <div class="row">
                    <div class="col-lg-2">Projektgruppe:</div>
                    <div class="col-lg-2">Navn:</div>
                    <div class="col-lg-2"><input type="radio" name="FM_periode" value="1" onchange="submit()" <%=usePeriodeSel1 %> /> Måned:</div>
                    <div class="col-lg-2"><input type="radio" name="FM_periode" value="2" onchange="submit()" <%=usePeriodeSel2 %> /> Kvartal:</div>
                    <div class="col-lg-2"><input type="radio" name="FM_periode" value="3" onchange="submit()" <%=usePeriodeSel3 %> /> År:</div>
             
                   
                

            </div>
            <div class="row">

                        <div class="col-lg-2">
                                 <select name="FM_progrp" class="form-control input-small" size="7" onchange="submit();">
                                <%
                                   
                                    if level = 1 then
                                    strSQLpgrp = "SELECT id, navn FROM projektgrupper WHERE id <> 0 ORDER BY navn"

                                    else

                                    strSQLpgrp = "SELECT id, navn FROM projektgrupper WHERE id = 10 ORDER BY navn"

                                    end if

                                   

                                    oRec3.open strSQLpgrp, oConn, 3
                                    while not oRec3.EOF 

                                    if cint(pgrp) = cint(oRec3("id")) then
                                    pSel = "SELECTED"
                                    else
                                    pSel = ""
                                    end if
                                    %>
                                    <option value="<%=oRec3("id") %>" <%=pSel %>><%=oRec3("navn")%></option>
                                    <%
                                    oRec3.movenext
                                    wend 
                                    oRec3.close
                                 %>
                  
                                </select>
                

                </div>
                <div class="col-lg-2">

                                 <select name="FM_medarb" class="form-control input-small" size="7" onchange="submit();">
                                <%

                                    if level = 1 then
                                    
                                    call medarbiprojgrp(pgrp, session("mid"), 0, -1)
                                    
                                    strSQLm = "SELECT mid, mnavn, init FROM medarbejdere WHERE mansat = 1 AND (mid = 0 "& medarbgrpIdSQLkri &") ORDER BY mnavn"

                                    else

                                    strSQLm = "SELECT mid, mnavn, init FROM medarbejdere WHERE mansat = 1 AND (mid = "& session("mid") &") ORDER BY mnavn"

                                    end if

                                   

                                    oRec3.open strSQLm, oConn, 3
                                    while not oRec3.EOF 

                                    if cint(usemrn) = cint(oRec3("mid")) then
                                    mSel = "SELECTED"
                                    else
                                    mSel = ""
                                    end if
                                    %>
                                    <option value="<%=oRec3("mid") %>" <%=mSel %>><%=oRec3("mnavn")& " ["& oRec3("init") &"]" %></option>
                                    <%
                                    oRec3.movenext
                                    wend 
                                    oRec3.close
                                 %>
                  
                                </select>
                

                      </div>
                      <div class="col-lg-2">
                          <select name="FM_mth" class="form-control input-small" onchange="submit()">
                          <option value="1" <%=mthSel1 %>>Jan</option>
                          <option value="2" <%=mthSel2 %>>Feb</option>
                          <option value="3" <%=mthSel3 %>>Mar</option>
                          <option value="4" <%=mthSel4 %>>Apr</option>
                          <option value="5" <%=mthSel5 %>>Maj</option>
                          <option value="6" <%=mthSel6 %>>Jun</option>
                          <option value="7" <%=mthSel7 %>>Jul</option>
                          <option value="8" <%=mthSel8 %>>Aug</option>
                          <option value="9" <%=mthSel9 %>>Sep</option>
                          <option value="10" <%=mthSel10 %>>Okt</option>
                          <option value="11" <%=mthSel11 %>>Nov</option>
                          <option value="12" <%=mthSel12 %>>Dec</option>
                   
                        </select>
                      </div>

                     <div class="col-lg-2">
                <select name="FM_kv" class="form-control input-small" onchange="submit()">
                  <option value="1" <%=useQuaterSel1 %>>Kvt. 1</option>
                  <option value="2" <%=useQuaterSel2 %>>Kvt. 2</option>
                  <option value="3" <%=useQuaterSel3 %>>Kvt. 3</option>
                  <option value="4" <%=useQuaterSel4 %>>Kvt. 4</option>
                </select>
                 </div>

                      <div class="col-lg-2">
                        <select name="FM_aar" class="form-control input-small" onchange="submit()">
                            <%
                           yearSt = datePart("yyyy", dateAdd("yyyy", -3, now), 2,2) 
                           yearEnd = yearSt + 4 
     
                           for y = yearSt TO yearEnd  
                        
                        
                               if cint(y) = cint(useYear) then
                                ySel = "SELECTED"
                               else
                                ySel = ""
                                end if %>
                          <option value="<%=y %>" <%=ySel%>><%=y %></option>
                          <%next %>
                 
                        </select>
                    </div>
                     
                <!--
                 
                   <div class="col-lg-4">
                      
                       <%           
                           
                           strSQLKproj = "SELECT kkundenavn, kkundenr, jobnavn, jobnr, j.id AS jid, kid FROM job AS j LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                           &" WHERE risiko > -1"_
                           &" ORDER BY kkundenavn, jobnavn" 
                           '(j.salgsans1 = "& usemrn &" OR j.salgsans2 = "& usemrn &" OR j.salgsans3 = "& usemrn &" OR j.salgsans4 = "& usemrn &" OR j.salgsans5 = "& usemrn &") 
                           '&" AND jobstartdato > '2015-01-01'"_
                           'response.write strSQLKproj
                           'response.flush

                           %>
                              <select name="FM_job" class="form-control input-small" multiple size="7">
                                  <%if cint(jobid) =  0 then
                                    jSel = "SELECTED"
                                    else
                                    jSel = ""
                                    end if %>

                                    <option value="0" <%=jSel%>>Alle</option>
                                    <%
                                    lastKid = 0
                                   k = 0
                                    oRec3.open strSQLKproj, oConn, 3
                                    while not oRec3.EOF 

                                    if lastKid <> oRec3("kid") then
                                  
                                    if k > 0 then %>
                                    <option value="0" DISABLED></option>
                                    <%end if %>

                                     <option value="0" DISABLED><%=oRec3("kkundenavn")%></option>
                                    <%
                                    end if


                                    if cint(jobid) = cint(oRec3("jid")) OR instr(jobidsStr, "OR j.id = "& oRec3("jid")) <> 0 then
                                    jSel = "SELECTED"
                                    else
                                    jSel = ""
                                    end if
                                    %>
                                    <option value="<%=oRec3("jid") %>" <%=jSel %>><%=oRec3("jobnavn") & " ("& oRec3("jobnr") &")"%></option>
                                    <%

                                    lastKid = oRec3("kid")
                                    k = k + 1
                                    oRec3.movenext
                                    wend 
                                    oRec3.close
                           
                           %>
                            </select>
                       </div>
                    -->

                         <div class="col-lg-1">
                           <button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Søg >></b></button>
                      </div>

                    </div>
                  
                 
            </form>
            </div>


            

       



        <%
         
         call akttyper2009(2)

         call meStamdata(usemrn)
        
            
        select case usePeriode
        case 0 'uge
        strDatoStartKri = useYear & "-9-1"
         strDatoEndKri = useYear & "-11-1"
        case 1 'month
        strDatoStartKri = useYear & "-"& useMonth &"-1"
        strDatoEndKri = dateAdd("m", 1, strDatoStartKri)
        strDatoEndKri = dateAdd("d", -1, strDatoEndKri)
        strDatoEndKri = year(strDatoEndKri) &"-"& month(strDatoEndKri) &"-"& day(strDatoEndKri)
        case 2 'quater

            select case useQuater
            case 1 
            strDatoStartKri = useYear & "-1-1"
            strDatoEndKri = useYear & "-3-31"
            case 2
            strDatoStartKri = useYear & "-4-1"
            strDatoEndKri = useYear & "-6-30"
            case 3
            strDatoStartKri = useYear & "-7-1"
            strDatoEndKri = useYear & "-9-30"
            case 4
            strDatoStartKri = useYear & "-10-1"
            strDatoEndKri = useYear & "-12-31"
            end select

        case 3 'year
        strDatoStartKri = useYear & "-1-1"
         strDatoEndKri = useYear & "-12-31"
        case else 'mth
        strDatoStartKri = useYear & "-9-1"
        strDatoEndKri = useYear & "-11-1"
        end select
           
        '** Mtype mål ***
        fo = 0
        foBA = 0
        dim fomrnavn, fomrmal, fomrid, fomrBAnavn, fomrBAmal, fomrBAid
        redim fomrnavn(0), fomrmal(0), fomrid(0), fomrBAnavn(0), fomrBAmal(0), fomrBAid(0)
        strSQLsel_mmff = "SELECT mmff_id, mmff_mal, fomr.navn AS fomrnavn, mmff_fomr, business_area_label, business_unit, fomr.id AS fomrid FROM fomr "_
        &"LEFT JOIN mtype_mal_fordel_fomr ON (mmff_fomr = fomr.id AND mmff_mtype = "& meType & ") WHERE fomr.id <> 0 LIMIT 30"
        
        'response.write strSQLsel_mmff
        'response.flush

        oRec4.open strSQLsel_mmff, oConn, 3
        while not oRec4.EOF
            
            if oRec4("business_unit") = "Salg" then


            redim preserve fomrBAnavn(foBA), fomrBAmal(foBA), fomrBAid(foBA)
            fomrBAid(foBA) = oRec4("fomrid")
            fomrBAnavn(foBA) = oRec4("business_area_label")
            fomrBAmal(foBA) = oRec4("mmff_mal")

            'response.write "fomrBAnavn(foBA): " & fomrBAnavn(foBA) & "<br>"

            foBA = foBA + 1 
    
            else

            redim preserve fomrnavn(fo), fomrmal(fo), fomrid(fo)
            fomrid(fo) = oRec4("fomrid")
            fomrnavn(fo) = oRec4("fomrnavn")

            if isNull(oRec4("mmff_mal")) <> true then
            fomrmal(fo) = oRec4("mmff_mal")
            else
            fomrmal(fo) = 0
            end if

            fo = fo + 1
            end if

        
        oRec4.movenext
        wend 
        oRec4.close

        '*** Forretningsområde aktids
        dim fomraktids
        redim fomraktids(fo-1)
        for f = 0 TO fo - 1 

        'if f = 0 OR lastF <> f then
        fomraktids(f) = "AND (taktivitetid = 0"
        'end if
        
        strSQLsel_mmff = "SELECT for_aktid FROM fomr_rel WHERE for_fomr = "& fomrid(f)

        'response.write strSQLsel_mmff
        'response.flush

        oRec4.open strSQLsel_mmff, oConn, 3
        while not oRec4.EOF
            
            if oRec4("for_aktid") <> 0 then
            fomraktids(f) = fomraktids(f) & " OR taktivitetid = "& oRec4("for_aktid")
            end if
            
        oRec4.movenext
        wend 
        oRec4.close

        fomraktids(f) = fomraktids(f) &")"

        lastF = f
        next


       '*** Forretningsområde aktids BA
        dim fomraktidsBA
        redim fomraktidsBA(foBA-1)

            foBAc = 0
            for foBAc = 0 TO foBA - 1 

            'if f = 0 OR lastF <> f then
            fomraktidsBA(foBAc) = "AND (taktivitetid = 0"
            'end if
        
            strSQLsel_mmff = "SELECT for_aktid FROM fomr_rel WHERE for_fomr = "& fomrBAid(foBAc)

            'response.write strSQLsel_mmff
            'response.flush

            oRec4.open strSQLsel_mmff, oConn, 3
            while not oRec4.EOF
            
                if oRec4("for_aktid") <> 0 then
                fomraktidsBA(foBAc) = fomraktidsBA(foBAc) & " OR taktivitetid = "& oRec4("for_aktid")
                end if
            
            oRec4.movenext
            wend 
            oRec4.close

            fomraktidsBA(foBAc) = fomraktidsBA(foBAc) &")"

            lastF = foBAc
            next

        

        '** Normtimer
        natalDage = dateDiff("d", strDatoStartKri, strDatoEndKri, 2,2)
        call normtimerPer(usemrn, strDatoStartKri, natalDage, 0)
        ntimPer = ntimPer


        '** Target timepris
        tgtimpris = 0 '600


            '**** RESSOURCE FORECAST timer TOTAL til INDEX ******'
           
            jobisMedLevDato = " AND (jobid = 0" 
            strSQLjobmLevDato = "SELECT j.id FROM job AS j WHERE jobslutdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"'"
            
            'response.Write strSQLjobmLevDato
            'response.flush
            oRec5.open strSQLjobmLevDato, oConn, 3
            while not oRec5.EOF 
            
            jobisMedLevDato = jobisMedLevDato & " OR jobid =  " & oRec5("id") 

            oRec5.movenext
            wend 
            oRec5.close 

            jobisMedLevDato = jobisMedLevDato & ")"

            useEndMonth = month(strDatoEndKri)
            strSQLResPer = " AND aar = "& useYear  &"" 'I FY = AND (md >= "& useMonth &" AND md <= "& useEndMonth &")"
           
            indexFc = 0

            strSQLforecast = "SELECT COALESCE(SUM(timer),0) AS fctimer FROM ressourcer_md WHERE medid = "& usemrn &" "& jobisMedLevDato &" GROUP BY medid"
            'strSQLResPer
            'response.write strSQLforecast
            'response.flush
            
            oRec5.open strSQLforecast, oConn, 3

            if not oRec5.EOF then
            indexFc = oRec5("fctimer")
            end if
            oRec5.close

            if indexFc <> 0 then
            indexFc = indexFc
            else
            indexFc = 1
            end if



    

        '*******************************
        '*** Alle timer periode 
         strSQlmtimer = "SELECT SUM(timer) AS sumtimer FROM timer AS t WHERE tmnr = "& usemrn & " AND ("& aty_sql_realHours &") AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"'" 
            
        medarbSelHoursGT = 0
        'response.write strSQlmtimer & "<hr>"
        'response.flush

         oRec.open strSQlmtimer, oConn, 3
         if not oRec.EOF then

        medarbSelHoursGT = oRec("sumtimer")  
           

         end if
         oRec.close 


        if medarbSelHoursGT <> 0 then
        medarbSelHoursGT = medarbSelHoursGT
        else
        medarbSelHoursGT = 0
        end if


           



        '*******************************
        '*** Fakurerbare timer periode 
         strSQlmtimer = "SELECT SUM(timer) AS sumtimer FROM timer AS t WHERE tmnr = "& usemrn & " AND ("& aty_sql_realHoursFakbar &") AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"'" 
            
        faktimerGTselmedarb = 0
        'response.write strSQlmtimer & "<hr>"
        'response.flush

         oRec.open strSQlmtimer, oConn, 3
         if not oRec.EOF then

        faktimerGTselmedarb = oRec("sumtimer")  
           

         end if
         oRec.close 


        if faktimerGTselmedarb <> 0 then
        faktimerGTselmedarb = faktimerGTselmedarb
        else
        faktimerGTselmedarb = 0
        end if




        '** Real. timer SEL MEDARB Fordelt på Fomr
         dim hoursDB, hoursDBindicator, hoursAvgRate, hoursAvgindicator, utilz, hoursAvgindicato, hours, hoursOms, hoursBA
         dim utilzDB, utilzCol, utilzDBFortegn, utilzDBBA
         dim hoursDBindex
         redim utilzDB(fo-1), utilzCol(fo-1), utilzDBFortegn(fo-1), hours(fo-1), hoursOms(fo-1) 
         redim hoursDB(fo-1), hoursDBindicator(fo-1), hoursAvgRate(fo-1), hoursAvgindicato(fo-1), utilz(fo-1)
         redim hoursDBindex(fo-1), hoursBA(foBA-1), utilzDBBA(foBA-1)
             
         medarbSelDBGT = 0
         medarbSelOmsGT = 0
         medarbSelDBindexGT = 0
         medarbSelHoursFomrGT = 0


       for f = 0 TO fo - 1 

        '*****************************************
        '************** DE 7 MAIN FORM ***********
         strSQlmtimer = "SELECT SUM(timer) AS sumtimer, SUM(timer*timepris) AS sumtimeOms, SUM(timer*kostpris) AS sumtimeKost FROM timer WHERE tmnr = "& usemrn & " AND ("& aty_sql_realhours &") AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"' "& fomraktids(f)
            
         hoursDB(f) = 0
         hoursDBindicator(f) = 0
         hoursAvgRate(f) = 0   
         hoursAvgindicato(f) = 0
         hoursOms(f) = 0 
         hoursDBindex(f) = 0
         hours(f) = 0

        'response.write strSQlmtimer & "<hr>"
        'response.flush

         oRec.open strSQlmtimer, oConn, 3
         if not oRec.EOF then

            if isNULL(oRec("sumtimer")) <> true then
            hours(f) = oRec("sumtimer")
            hoursDB(f) = oRec("sumtimeOms") - oRec("sumtimeKost")
            hoursOms(f) = oRec("sumtimeOms")


            hoursDBindicator(f) = (oRec("sumtimeKost") / oRec("sumtimeOms")) * 100 'dbproc
            hoursAvgRate(f) = oRec("sumtimeOms") / oRec("sumtimer")

            if hoursAvgRate(f) > tgtimpris then
            hoursAvgindicato(f) = 100 - ((tgtimpris / hoursAvgRate(f) * 100))
            else
            hoursAvgindicato(f) = (hoursAvgRate(f) / tgtimpris * 100)
            end if

          

            else
            hours(f) = 0
            hoursDB(f) = 0
            hoursOms(f) = 0

            
            hoursDBindicator(f) = 0
            hoursAvgRate(f) = 0

            hoursAvgindicato(f) = 0
            
           
          
            end if

            
           

         end if
         oRec.close 


           

            
            medarbSelHoursFomrGT = medarbSelHoursFomrGT + hours(f)
            'medarbSelHoursGT = medarbSelHoursGT + hours(f)
            medarbSelDBGT = medarbSelDBGT + hoursDB(f)
            medarbSelOmsGT = medarbSelOmsGT + hoursOms(f)

            'response.write "medarbSelHoursGT: " & formatnumber(medarbSelHoursGT, 0) &"<br>"
            

        next

            

            '** IKKE i forhold til norm, men i forhold til reg timer
            for f = 0 TO fo - 1 

            utilz(f) = 0
            utilzDB(f) = 0
            utilzCol(f) = ""
            utilzDBFortegn(f) = ""

                    utilz(f) = (hours(f) - medarbSelHoursGT)

                    if hours(f) <> 0 then

                    utilz(f) = hours(f)

                        if medarbSelHoursGT > hours(f) then
                        utilzDB(f) = (hours(f) / medarbSelHoursGT) * 100
                        utilzCol(f) = "danger"
                        utilzDBFortegn(f) = "-"
                        else
                        utilzDB(f) = (medarbSelHoursGT / hours(f)) * 100
                        utilzCol(f) = "success"
                        utilzDBFortegn(f) = "+"
                        end if

                    else

                    utilz(f) = 0
                    utilzDB(f) = 0
                    utilzCol(f) = "success"
                    utilzDBFortegn(f) = "+"

                    end if

            next


            if medarbSelOmsGT <> 0 then
            medarbSelOmsGT = formatnumber(medarbSelOmsGT, 0)
            else
            medarbSelOmsGT = 0
            end if

            if medarbSelDBGT <> 0 then
            medarbSelDBGT = formatnumber(medarbSelDBGT, 0)
            else
            medarbSelDBGT = 0
            end if

            if medarbSelHoursGT <> 0 then
            medarbSelHoursGT = formatnumber(medarbSelHoursGT, 2)
            else
            medarbSelHoursGT = 0
            end if

            medarbSelIndex = 0
            if cdbl(indexFc) > cdbl(medarbSelHoursGT) then
            medarbSelIndex = (1 - (medarbSelHoursGT/indexFc)) + 1
            else
            medarbSelIndex = (indexFc/medarbSelHoursGT)
            end if

            'if (medarbSelIndex) < 1 then
            'medarbSelIndex = medarbSelIndex
            'else
            'medarbSelIndex = 0
            'end if

            medarbSelDBindexGT = (medarbSelIndex) * medarbSelDBGT

         





        '** Real. timer MEDARB PROJEKTGRUPPE ***'
        call hentbgrppamedarb(usemrn)
        
        '** FJERN ALLE GRUPPEN
        projektgruppeIds = replace(projektgruppeIds, ", 10", "") 

        call medarbiprojgrp(projektgruppeIds, usemrn, 0, -1)


        medarbgrpIdSQLkri = replace(medarbgrpIdSQLkri, "mid", "tmnr")



        
        visPeers = 0
        if cint(visPeers) = 1 then

        strSQlmtimerProjgrp = "SELECT SUM(timer) AS sumtimer, SUM(timer*timepris) AS sumtimeOms, SUM(timer*kostpris) AS sumtimeKost, "_
        &" SUM(timer*timepris-timer*kostpris) AS pghoursDB FROM timer WHERE "_
        &" (tmnr = 0 "& medarbgrpIdSQLkri &") AND ("& aty_sql_realhours &") AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"' GROUP BY tmnr"
            
         pghoursDB = 0
         pghoursDBindicator = 0
         pghoursAvgRate = 0   
         pghoursAvgindicator = 0
         pgutilz = 0
         m = 1
        'response.write strSQlmtimerProjgrp
        'response.flush

         oRec.open strSQlmtimerProjgrp, oConn, 3
         while not oRec.EOF


            pghoursDB = pghoursDB + oRec("pghoursDB")
            pghoursDBindicator = pghoursDBindicator + (oRec("sumtimeKost") / oRec("sumtimeOms")) * 100 'dbproc

         

            m = m + 1

         oRec.movenext
         wend 
         oRec.close 


            

            pghoursDB = pghoursDB / m-1
            pghoursDBTxt = replace(formatnumber(pghoursDB, 0), ".", "")
            
            pghoursDBindicator = pghoursDBindicator / m-1


        end if 'vis Peers

            
            f = 1
            if hoursDBindicator(f) < 20 then
            hoursDBindicatorColor = "danger"
            else
            hoursDBindicatorColor = "success"
            end if

            if hoursAvgindicato(f) < 0 then
            hoursAvgindicatoColor = "danger"
            hoursAvgindicatoFortegn = "-"
            else
            hoursAvgindicatoColor = "success"
            hoursAvgindicatoFortegn = "+"
            end if

            if cdbl(medarbSelIndex) < 1 then
            medarbSelIndexColor = "danger"
            medarbSelIndexFortegn = "-"
            else
            medarbSelIndexColor = "success"
            medarbSelIndexFortegn = "+"
            end if
            

            


            '***********************************************************************************************************************************'
            '*** Personlig DB 1 og DB 2, Fakturaindex og Timefordelign
            '***********************************************************************************************************************************'

                '*** PIE 1 ***
               %>
                 <form>
                <input type="hidden" id="timefordeling_proc_norm" value="<%=formatnumber(ntimPer, 0) %>" />
                <!--<input type="hidden" id="timefordeling_proc_fravar" value="<%=20 %>" />-->
                <%for f = 0 TO fo - 1  %>

              
                <input type="hidden" id="timefordeling_proc_<%=f %>" value="<%=replace(formatnumber(utilzDB(f), 0), ".", "") %>" />
                <input type="hidden" id="" value="<%=replace(formatnumber(utilz(f), 0), ".", "") %>" />
                <input type="hidden" id="timefordeling_navn_<%=f %>" value="<%=fomrnavn(f)%>" />
                <%next %>
                </form>

              

     


        <div class="portlet-body">
            <br />
            <div class="row">
                <div class="col-lg-11">
                    <h4 class="portlet-title"><u>Nøgletal <span style="font-size:12px; font-weight:lighter;"><%=meTxt %></span> </u>
                    </h4>
                </div>
            </div>


                 <div class="row">
            
                 <div class="col-lg-4">


                    
                      <div class="row-stat">
                        <p class="row-stat-label">DB1</p>
                        <h3 class="row-stat-value"><%=formatnumber(medarbSelDBGT, 2)%> <span style="font-size:14px;">DKK</span></h3>
                     
                          <br /><br />
                        <p class="row-stat-label">DB2</p>
                        <h3 class="row-stat-value"><%=formatnumber(medarbSelDBindexGT, 2)%> <span style="font-size:14px;">DKK</span></h3>
                        <!--<span class="label label-succes row-stat-badge"><%=formatnumber(jo_db2_proc, 0)%> %</span> -->
                          <br /> <span style="font-size:11px; color:#999999;"> [ Timer * (timepris - Kost.) * Faktureringsindex ]</span>

                     </div> <!-- /.row-stat -->
         

                     
                 </div>
                     <div class="col-lg-3">
                         Faktureringsindex: 
                         <span class="label label-<%=medarbSelIndexColor%> row-stat-badge"><%=formatnumber(medarbSelIndex, 2)%></span>
                         <br /> <span style="font-size:11px; color:#999999;">[ Forecast timer/ Real. fakturerbare timer ]</span>
                         <br /><br />
                         Norm: <%=formatnumber(ntimper, 2) %> t.<br />
                         Timer: <%=formatnumber(medarbSelHoursGT, 2) %> t.<br />
                         Heraf fakturerbare: <b><%=formatnumber(faktimerGTselmedarb, 2) %> t.</b><br />
                         Forecasttimer: <%=formatnumber(indexFc, 2) %> t.<br />
                         <span style="font-size:11px; color:#999999;">Projekt med lev. dato i valgt periode.</span>
                         <br /><br />
                        <div class="row-stat">
                        <p class="row-stat-label">NPS</p>
                        <h3 class="row-stat-value">8,2</h3>
                
                      </div> <!-- /.row-stat -->

                 </div>
                       <div class="col-lg-1">&nbsp;</div>
                      
                      <div class="col-lg-4">
                     Tidsfordeling - Fomr. % (<%=formatnumber(medarbSelHoursFomrGT, 2) %> t.)
                             
                    <div id="donut-chart" class="chart-holder" style="width: 95%;"></div>
                    

                 </div>

                  </div><!-- END ROW -->


            <!-- KUN SÆLGERE -->

            <%
              erSaelgerOk = 0  
              strSQLprogrp = "SELECT medarbejderid FROM progrupperelationer WHERE projektgruppeid = 2 OR projektgruppeid = 12"
                
              oRec5.open strSQLprogrp, oConn, 3
              if not oRec5.EOF then
              erSaelgerOk = 1  
              end if
                
              oRec5.close%>

            <%if erSaelgerOk = 1 then %>
            <br />
              <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseOne">Kunder top 10 - <span style="font-size:12px; font-weight:lighter;"><%=meTxt %></span></a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseOne" class="panel-collapse collapse">
                        <div class="panel-body">

                


            
                        <div class="row">
                        <div class="col-lg-8">
                     
                        <table class="table table-stribed">
                            <thead>
                                <tr>
                                    <th style="width: 50%">Kunder (kundenr)</th>
                                    <th style="width: 20%">Forecast</th>

                                    <th style="width: 20%">Realiseret</th>
                                    <th style="width: 10%"><!--Indikator--></th>

                                     <%for foBAc = 0 TO foBA - 1%>
                                    <th><%=fomrBAnavn(foBAc)%></th>

                                    <%next%>

                                </tr>
                            </thead>
                            <tbody>
                                <%
                                 '** Real. timer SEL MEDARB
                                 strSQlmtimer_k = "SELECT SUM(t.timer) AS sumtimer, SUM(t.timer*t.timepris) AS sumtimeOms, SUM(t.timer*t.kostpris) AS sumtimeKost, "_
                                 &" t.tknavn, t.tknr, t.tjobnavn, t.tjobnr, SUM(r.timer) AS fctimer "_
                                 &" FROM timer AS t "_
                                 &" LEFT JOIN job AS j ON (jobnr = tjobnr)"_ 
                                 &" LEFT JOIN ressourcer_md AS r ON (medid = "& usemrn &" AND jobid = j.id)"_
                                 &" WHERE t.tmnr = "& usemrn & " AND ("& aty_sql_realhours &") AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"' GROUP BY tknr ORDER BY sumtimer DESC LIMIT 10"
            
                                'response.write strSQlmtimer_k
                                'response.flush
                                
                                 oRec.open strSQlmtimer_k, oConn, 3
                                 while not oRec.EOF 

                                    if isNull(oRec("fctimer")) <> true then
                                    fctimer = oRec("fctimer")
                                    else
                                    fctimer = 0
                                    end if

                                    if cdbl(fctimer) > cdbl(oRec("sumtimer")) then
                                    indicatorCol = "yellowgreen"
                                    else
                                    indicatorCol = "crimson"
                                    end if
                                    %>

                                <tr>
                                    <td><%=oRec("tknavn") %> (<%=oRec("tknr") %>)</td>
                                    <td class="text-right;"><%=formatnumber(fctimer, 2) %></td>
                                    <td class="text-right;"><%=formatnumber(oRec("sumtimer"), 2) %></td>
                                    <td><div style="background-color:<%=indicatorCol%>; width:10px; height:10px;">&nbsp;</div></td>

                                    <%for foBAc = 0 TO foBA - 1
                                        
                                        '*****************************************************************************
                                        '************** SALGS FORM Business Area = 'SALG' fordelt på kunde ***********
                                         strSQlmtimer_fomr = "SELECT SUM(timer) AS sumtimer FROM timer "_
                                         &" WHERE tmnr = "& usemrn & " AND ("& aty_sql_realhours &") AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"' "& fomraktidsBA(foBAc) & " AND tknr = "& oRec("tknr") &" GROUP BY tknr"
            
                                         
                                        'response.write strSQlmtimer_fomr & "<hr>"
                                        'response.flush
                                        thisTimerFomrBA = 0
                                         oRec5.open strSQlmtimer_fomr, oConn, 3
                                         if not oRec5.EOF then

                                            if isNULL(oRec5("sumtimer")) <> true then
                                            hoursBA(foBAc) = oRec5("sumtimer")
                                            else
                                            hoursBA(foBAc) = 0
                                            end if
                                            
                                           if isNULL(oRec5("sumtimer")) <> true then
                                            thisTimerFomrBA = formatnumber(oRec5("sumtimer"), 0)
                                           else
                                            thisTimerFomrBA = 0
                                           end if

                                         end if
                                         oRec5.close 

                                            %>
                                            <td style="text-align:right;"><%=thisTimerFomrBA %></td>

                                            <%

                                         medarbSelHoursBAGT = (medarbSelHoursBAGT*1 + hoursBA(foBAc)*1)
           

                                     next

                                   
                                        'medarbSelHoursBAGT = replace(formatnumber(medarbSelHoursBAGT,0), ".", "")
                                                

                                        
                                        %>
                                   
                                </tr>

                                <%
                                 oRec.movenext
                                 wend
                                 oRec.close 
                                    
                                    
                                 for foBAc = 0 TO foBA - 1

                                            'response.write "medarbSelHoursGT: " & medarbSelHoursGT & " hoursBA(foBA): "& hoursBA(foBA) &"#<br>"

                                            if medarbSelHoursGT > hoursBA(foBAc) then
                                                 if medarbSelHoursGT <> 0 then
                                                 utilzDBBA(foBAc) = (hoursBA(foBAc) / medarbSelHoursGT)*100
                                                 else
                                                 utilzDBBA(foBAc) = 0
                                                 end if
                                            else
                                                if hoursBA(foBAc) <> 0 then
                                                utilzDBBA(foBAc) = (hoursBA(foBAc) / medarbSelHoursGT) * 100
                                                else
                                                utilzDBBA(foBAc) = 0
                                                end if
                                            end if
                                        
                                             

                                  next   
                                    
                                    
                                 %>


                               
                            </tbody>
                        </table>
                    </div>

                            <%
                 '*** PIE 2 ***
               %>
                 <form>
               
                <!--<input type="hidden" id="timefordeling_proc_fravar" value="<%=20 %>" />-->
                <%for foBAc = 0 TO foBA - 1  %>

              
                <input type="hidden" id="timefordeling_proc_saelger_<%=foBAc %>" value="<%=replace(formatnumber(utilzDBBA(foBAc), 0), ".", "") %>" />
                <input type="hidden" id="timefordeling_navn_saelger_<%=foBAc %>" value="<%=fomrBAnavn(foBAc)%>" />
                <%next %>
                </form>


                           
                    <div class="col-lg-4">
                    <!--Tidsfordeling Fomr. - Salg %-->
                    <div id="xdonut-chart_saelger" class="chart-holder" style="width: 95%;"></div>
                    </div>


                    </div>
                    </div><!-- END panel-body -->
              </div><!-- END panel-collapse -->
             </div><!-- END panel deafult -->

            <%end if %>

          



               <!-- JoB HVIS jobsansvarlig -->

                <%
                 jobansFundet = 0
                 jobansKrijobids = " AND (tjobnr = '0' "    
                 strSQKjobans = "SELECT jobnr FROM job WHERE (jobans1 = "& usemrn &" OR jobans2 = "& usemrn &") AND jobslutdato >= '"& strDatoStartKri &"'" 
                 
                    'response.write strSQKjobans

                    
                 oRec.open strSQKjobans, oConn, 3
                 while not oRec.EOF 
                 
                       
                    jobansKrijobids = jobansKrijobids & " OR tjobnr = '" & oRec("jobnr") & "'"

                    jobansFundet = 1
                 oRec.movenext
                 wend
                 oRec.close 
                 
                    jobansKrijobids = jobansKrijobids & ")"   
                    


                 if cint(jobansFundet) = 1 then
                 %>

             
                   <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseTwo">Projektansvarlig</a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseTwo" class="panel-collapse collapse">
                        <div class="panel-body">

            

                 <div class="row">
            
                 <div class="col-lg-12">
                   
                     <table class="table table-stribed">
                            <thead>
                                <tr>
                                    <th style="width: 25%">Kunder (kundenr)</th>
                                    <th style="width: 25%">Job</th>
                                    <th style="width: 10%">Lev. dato</th>
                                    <th style="width: 10%">Forecast</th>
                                    <th style="width: 10%">Realiseret</th>
                                    <th style="width: 10%">Omsætning</th>
                                    <th style="width: 10%"><!--Indikator--></th>


                                </tr>
                            </thead>
                            <tbody>
                                <%
                                 '** Real. timer SEL MEDARB JOBANSVARLIG
                                 strSQlmtimer_k = "SELECT SUM(t.timer) AS sumtimer, SUM(t.timer*t.timepris) AS sumtimeOms, SUM(t.timer*t.kostpris) AS sumtimeKost, "_
                                 &" t.tknavn, t.tknr, t.tjobnavn, t.tjobnr, j.id AS jid, jobslutdato "_
                                 &" FROM timer AS t "_
                                 &" LEFT JOIN job AS j ON (j.jobnr = tjobnr) "_
                                 &" WHERE tjobnr <> 0 "& jobansKrijobids &" AND ("& aty_sql_realhours &") GROUP BY tjobnr ORDER BY jobslutdato DESC, t.tknr, tjobnavn DESC LIMIT 30"
            
                                'response.write strSQlmtimer_k
                                'response.flush

                                 oRec.open strSQlmtimer_k, oConn, 3
                                 while not oRec.EOF 

                                    'if isNull(oRec("fctimer")) <> true then
                                    'fctimer = oRec("fctimer")
                                    'else
                                    fctimer = 0
                                    'end if
                                    strSQL_r = "SELECT COALESCE(SUM(timer), 0) AS fctimer FROM ressourcer_md AS r WHERE jobid = " & oRec("jid") 
                                    oRec2.open strSQL_r, oConn, 3
                                    if not oRec2.EOF then
                                    
                                    fctimer = oRec2("fctimer")

                                    end if
                                    oRec2.close 

                                    if fctimer <> 0 then
                                    fctimer = formatnumber(fctimer, 0)
                                    else
                                    fctimer = 0
                                    end if



                                    if cdbl(fctimer) > cdbl(oRec("sumtimer")) then
                                    indicatorCol = "yellowgreen"
                                    else
                                    indicatorCol = "crimson"
                                    end if
                                    %>

                                <tr>
                                    <td><%=oRec("tknavn") %> (<%=oRec("tknr") %>)</td>
                                    <td><%=oRec("tjobnavn") %> (<%=oRec("tjobnr") %>)</td>
                                    <td><%=oRec("jobslutdato") %></td>
                                    <td class="text-right;"><%=formatnumber(fctimer, 2) %></td>
                                    <td class="text-right;"><%=formatnumber(oRec("sumtimer"), 2) %></td>
                                    <td class="text-right;"><%=formatnumber(oRec("sumtimeOms"), 2) %></td>
                                    <td><div style="background-color:<%=indicatorCol%>; width:10px; height:10px;">&nbsp;</div></td>

                                  
                                   
                                </tr>

                                <%
                                 oRec.movenext
                                 wend
                                 oRec.close %>


                               
                            </tbody>
                        </table>
                 </div>
                  
                     
                  </div><!-- END ROW -->
                
                    </div><!-- END panel-body -->
              </div><!-- END panel-collapse -->
             </div><!-- END panel deafult -->

               <%end if 'jobansfundet %>





             <!-- KUNDER HVIS kundeansvarlig -->

                <%
                 kundeansFundet = 0
                 kansKrijobids = " AND (tjobnr = '0' "    
                 strSQKjobans = "SELECT jobnr, kid FROM kunder "_
                 &" LEFT JOIN job ON (jobknr = kid AND jobslutdato >= '"& strDatoStartKri &"') WHERE (kundeans1 = "& usemrn &" OR kundeans2 = "& usemrn &") " 
                 
                    'response.write strSQKjobans
                    'response.flush
                    
                 oRec.open strSQKjobans, oConn, 3
                 while not oRec.EOF 
                 
                       
                    kansKrijobids = kansKrijobids & " OR tjobnr = '" & oRec("jobnr") & "'"

                    kundeansFundet = 1
                 oRec.movenext
                 wend
                 oRec.close 
                 
                    kansKrijobids = kansKrijobids & ")"   
                    


                 if cint(kundeansFundet) = 1 then
                 %>

             
                   <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseTree">Kundeansvarlig - <span style="font-size:12px; font-weight:lighter;">Keyaccount</span></a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseTree" class="panel-collapse collapse">
                        <div class="panel-body">

            

                 <div class="row">
            
                 <div class="col-lg-12">
                   
                     <table class="table table-stribed">
                            <thead>
                                <tr>
                                    <th style="width: 25%">Kunder (kundenr)</th>
                                    <th style="width: 25%">Job</th>
                                    <th style="width: 10%">Lev. dato</th>
                                    <th style="width: 10%">Forecast</th>
                                    <th style="width: 10%">Realiseret</th>
                                    <th style="width: 10%">Omsætning</th>
                                    <th style="width: 10%"><!--Indikator--></th>


                                </tr>
                            </thead>
                            <tbody>
                                <%
                                 '** Real. timer SEL MEDARB KUNDEANSVARLIG
                                 strSQlmtimer_k = "SELECT SUM(t.timer) AS sumtimer, SUM(t.timer*t.timepris) AS sumtimeOms, SUM(t.timer*t.kostpris) AS sumtimeKost, "_
                                 &" t.tknavn, t.tknr, t.tjobnavn, t.tjobnr, j.id AS jid, jobslutdato "_
                                 &" FROM timer AS t "_
                                 &" LEFT JOIN job AS j ON (j.jobnr = tjobnr) "_
                                 &" WHERE tjobnr <> 0 "& kansKrijobids &" AND ("& aty_sql_realhours &") GROUP BY tjobnr ORDER BY t.tknr, tjobnavn DESC LIMIT 30"
            
                                'response.write strSQlmtimer_k
                                'response.flush

                                 oRec.open strSQlmtimer_k, oConn, 3
                                 while not oRec.EOF 

                                    'if isNull(oRec("fctimer")) <> true then
                                    'fctimer = oRec("fctimer")
                                    'else
                                    fctimer = 0
                                    'end if
                                    strSQL_r = "SELECT COALESCE(SUM(timer), 0) AS fctimer FROM ressourcer_md AS r WHERE jobid = " & oRec("jid") 
                                    oRec2.open strSQL_r, oConn, 3
                                    if not oRec2.EOF then
                                    
                                    fctimer = oRec2("fctimer")

                                    end if
                                    oRec2.close 

                                    if fctimer <> 0 then
                                    fctimer = formatnumber(fctimer, 0)
                                    else
                                    fctimer = 0
                                    end if



                                    if cdbl(fctimer) > cdbl(oRec("sumtimer")) then
                                    indicatorCol = "yellowgreen"
                                    else
                                    indicatorCol = "crimson"
                                    end if
                                    %>

                                <tr>
                                    <td><%=oRec("tknavn") %> (<%=oRec("tknr") %>)</td>
                                    <td><%=oRec("tjobnavn") %> (<%=oRec("tjobnr") %>)</td>
                                    <td><%=oRec("jobslutdato") %></td>
                                    <td class="text-right;"><%=formatnumber(fctimer, 2) %></td>
                                    <td class="text-right;"><%=formatnumber(oRec("sumtimer"), 2) %></td>
                                    <td class="text-right;"><%=formatnumber(oRec("sumtimeOms"), 2) %></td>
                                    <td><div style="background-color:<%=indicatorCol%>; width:10px; height:10px;">&nbsp;</div></td>

                                  
                                   
                                </tr>

                                <%
                                 oRec.movenext
                                 wend
                                 oRec.close %>


                               
                            </tbody>
                        </table>
                 </div>
                  
                     
                  </div><!-- END ROW -->
                
                    </div><!-- END panel-body -->
              </div><!-- END panel-collapse -->
             </div><!-- END panel deafult -->

               <%end if 'jobansfundet %>
        



            <!------------------------------- END PAGE ---------------------------------------------------------------------->


            

  </div>
</div>







<!--#include file="../inc/regular/footer_inc.asp"-->