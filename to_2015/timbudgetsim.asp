<%response.buffer = true

 %>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/webblik_func.asp"-->
<!--XXXinclude file="../timereg/inc/ressource_belaeg_jbpla_inc.asp"-->
<!--#include file="../timereg/inc/isint_func.asp"-->
<!--#include file="../to_2015/inc/timbudgetsim_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--XXXinclude file="../inc/regular/header_lysblaa_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->


<style type="text/css">
    	input { text-align:right; }
        .inputclsLeft { text-align:left; }

        /* #div1 {width:350px;height:270px; padding:20px;border:5px solid #999999; z-index:20000; position:absolute; top:100px; left:600px;} */
</style>



<script src="js/timbudgetsim_jav.js"></script>

<%

if session("user") = "" then


	errortype = 5
	call showError(errortype)
	else
	


            '*****************************************************************************
            '***************** VARIABLE **************************************************

            call timesimon_fn()


            if len(trim(request("FM_sog"))) <> 0 then
            sogVal = request("FM_sog")
            sogValTxt = sogVal
            else
            sogVal = "WF9)8/NXN76E#"
            sogValTxt = ""
            end if

             call aktBudgettjkOn_fn()
             fyStMd = month(aktBudgettjkOnRegAarSt)

            if len(trim(request("FM_fy"))) <> 0 then
            y0 = "01-"&fyStMd&"-"& request("FM_fy")
            addValyear = Y0
            else

                select case lto
                case "wwf" 'juli = Bør sættes udfra: aktBudgettjkOnRegAarSt
                
                if month(now) >= 7 then
                addValyear = dateAdd("yyyy", 0, now)
                else
                addValyear = dateAdd("yyyy", -1, now)
                end if

                case else
                addValyear = dateAdd("yyyy", 0, now)
                end select

            y0 = "01-"&fyStMd&"-"& year(addValyear)
            end if

             y1 = dateAdd("yyyy", 1, y0)
             y2 = dateAdd("yyyy", 2, y0) 


           

             h1aar = datePart("yyyy", addValyear, 2,2)
     
             select case lto
             case "wwf"
             h1md = 7 'juli = Bør sættes udfra: aktBudgettjkOnRegAarSt
             case else
             h1md = 1 'juli = Bør sættes udfra: aktBudgettjkOnRegAarSt
             end select    

             h2aar = datePart("yyyy", dateAdd("yyyy", 1, addValyear), 2,2) 
             h2md = 1 'jan


            if len(trim(request("FM_visrealprdato"))) <> 0 then
               visrealprdato = request("FM_visrealprdato")

                if isDate(visrealprdato) = false then
                visrealprdato = formatdatetime(now, 2)
                end if

            else
                select case lto
                case "wwf"
                visrealprdato = "1-1-"& h2aar 'juli = Bør sættes udfra: aktBudgettjkOnRegAarSt
                case else
                visrealprdato = "1-1-"& h1aar
                end select
    
            end if

            func = request("func")

        

                if request("issubmitted") = "1" then
                filterKunAktiveMedarb = request("FM_viskunmedarbMtimFc")
                
                if cint(filterKunAktiveMedarb) <> 0 then
                filterKunAktiveMedarbCHK = "CHECKED"
                else
                filterKunAktiveMedarbCHK = ""
                end if
        
                response.Cookies("EPR")("viskunmedarbMtimFc") = filterKunAktiveMedarb

                else


                    if request.Cookies("EPR")("viskunmedarbMtimFc") = "1" AND request("issubmitted") <> "1" then
                    filterKunAktiveMedarb = 1
                    filterKunAktiveMedarbCHK = "CHECKED"
                    else
                    filterKunAktiveMedarb = 0
                    filterKunAktiveMedarbCHK = ""
                    end if
        
                end if


                if request("issubmitted") = "1" then
                visKunFCFelter = request("FM_visKunFCFelter")
                response.Cookies("EPR")("FM_visKunFCFelter") = visKunFCFelter

                else


                    if request.Cookies("EPR")("FM_visKunFCFelter") <> "" AND request("issubmitted") <> "1" then
                    visKunFCFelter = request.Cookies("EPR")("FM_visKunFCFelter")
                    else
                    visKunFCFelter = 0
                    end if
        
                end if
                
                visKunFCFelterCHK0 = ""
                visKunFCFelterCHK1 = ""
                visKunFCFelterCHK2 = ""
                visKunFCFelterCHK3 = ""

                select case cint(visKunFCFelter) 
                case 0
                visKunFCFelterCHK0 = "CHECKED"
                case 1
                visKunFCFelterCHK1 = "CHECKED"
                case 2
                visKunFCFelterCHK2 = "CHECKED"
                case 3
                visKunFCFelterCHK3 = "CHECKED"
                case else
                visKunFCFelterCHK0 = "CHECKED"
                end select
              



                if request("issubmitted") = "1" then 
                filtervisallejobvlgtmedarb = request("FM_visallejobvlgtmedarb")

                if cint(filtervisallejobvlgtmedarb) <> 0 then
                filtervisAlleJobVlgtMedarbCHK = "CHECKED"
                else
                filtervisAlleJobVlgtMedarbCHK = ""
                end if

                response.Cookies("EPR")("visallejobvlgtmedarb") = filtervisallejobvlgtmedarb

                else


                    if request.Cookies("EPR")("visallejobvlgtmedarb") = "1" AND request("issubmitted") <> "1" then
                    filtervisallejobvlgtmedarb = 1
                    filtervisAlleJobVlgtMedarbCHK = "CHECKED"
                    else
                    filtervisallejobvlgtmedarb = 0
                    filtervisAlleJobVlgtMedarbCHK = ""
                    end if
        
                end if

   


             jbs = 5000
             akts = 50000

             dim strJobTxtTds, strAktTxtTds
             redim strJobTxtTds(jbs), strAktTxtTds(akts)

             public jobFcGtGT


            '*****************************************************************************


   select case func
   case "opd"

  
    '*** Job ****'
    jobids = split(request("FM_jobid"), ", ")
    jobbudgettimer = split(request("FM_jobbudgettimer"), ", ##, ") '", "

    'response.write "--------------------" & request("FM_jobbudgettimer")
    jobtpris = split(request("FM_jobtpris"), ", ##, ")

    'jobBudgetFY0 = split(request("FM_timebudget_FY0"), ", ##, ")
    'fctimeprisFY0 = split(request("FM_fctimepris_FY0"), ", ##, ")

    'response.write "<br>--------------------" & request("FM_timebudget_FY0")
    'response.flush

    'fctimeprish2FY0 = split(request("FM_fctimeprish2_FY0"), ", ##, ")

    'jobBudgetFY1 = split(request("FM_timebudget_FY1"), ", ##, ")
    'jobBudgetFY2 = split(request("FM_timebudget_FY2"), ", ##, ")

    '** FY ÅRSTAL 
    FY0 = request("FY0") 
    FY1 = request("FY1")
    FY2 = request("FY2")
    
   

    '*** Medarbejdere på job ***'

    'response.write request("FM_h1")
    'response.end

  

    for j = 0 TO UBOUND(jobids)




              jobbudgettimer(j) = replace(jobbudgettimer(j), ", ##", "")
              if len(trim(jobbudgettimer(j))) <> 0 then
              jobbudgettimer(j) = replace(jobbudgettimer(j), ".", "")
              jobbudgettimer(j) = replace(jobbudgettimer(j), ",", ".")
              else
              jobbudgettimer(j) = 0
              end if

              jobtpris(j) = replace(jobtpris(j), ", ##", "")
              if len(trim(jobtpris(j))) <> 0 then
              jobtpris(j) = replace(jobtpris(j), ".", "")
              jobtpris(j) = replace(jobtpris(j), ",", ".")
              else
              jobtpris(j) = 0
              end if

       

        '** !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        '** Hvis job er lbn timer beregn jobTpris / jo_bruttooms
        '** !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

        strSQLupdateJob = "UPDATE job SET budgettimer = " & jobbudgettimer(j) & ", jo_gnstpris = "& jobtpris(j) &" WHERE id = " & jobids(j)
        'response.write strSQLupdateJob
        'response.flush
        oConn.execute(strSQLupdateJob)



        '**** DER SKAL KUN OPDATES PÅ AKTIVITETER / JOB ER EN SUM AF AKTIVITETER
      
              'jobBudgetFY0(j) = replace(jobBudgetFY0(j), ", ##", "")
              'if len(trim(jobBudgetFY0(j))) <> 0 then
              'jobBudgetFY0(j) = replace(jobBudgetFY0(j), ".", "")
              'jobBudgetFY0(j) = replace(jobBudgetFY0(j), ",", ".")
              'else
              'jobBudgetFY0(j) = 0
              'end if


              'fctimeprisFY0(j) = replace(fctimeprisFY0(j), ", ##", "")
              'if len(trim(fctimeprisFY0(j))) <> 0 then
              'fctimeprisFY0(j) = replace(fctimeprisFY0(j), ".", "")
              'fctimeprisFY0(j) = replace(fctimeprisFY0(j), ",", ".")
              'else
              'fctimeprisFY0(j) = 0
              'end if


              'fctimeprish2FY0(j) = replace(fctimeprish2FY0(j), ", ##", "")
              'if len(trim(fctimeprish2FY0(j))) <> 0 then
              'fctimeprish2FY0(j) = replace(fctimeprish2FY0(j), ".", "")
              'fctimeprish2FY0(j) = replace(fctimeprish2FY0(j), ",", ".")
              'else
              'fctimeprish2FY0(j) = 0
              'end if


              'jobBudgetFY1(j) = replace(jobBudgetFY1(j), ", ##", "")
              'if len(trim(jobBudgetFY1(j))) <> 0 then
              'jobBudgetFY1(j) = replace(jobBudgetFY1(j), ".", "")
              'jobBudgetFY1(j) = replace(jobBudgetFY1(j), ",", ".")
              'else
              'jobBudgetFY1(j) = 0
              'end if


              'jobBudgetFY2(j) = replace(jobBudgetFY2(j), ", ##", "")
              'if len(trim(jobBudgetFY2(j))) <> 0 then
              'jobBudgetFY2(j) = replace(jobBudgetFY2(j), ".", "")
              'jobBudgetFY2(j) = replace(jobBudgetFY2(j), ",", ".")
              'else
              'jobBudgetFY2(j) = 0
              'end if

         
        
        'for f = 0 to 2
        '** RAMME FY0-FY2 **'

        'if cint(timesimh1h2) = 1 then
            
        '    call opdaterRessouceRamme(f, FY0, FY1, FY2, jobBudgetFY0(j), fctimeprisFY0(j), fctimeprish2FY0(j), jobBudgetFY1(j), jobBudgetFY2(j), jobids(j), 0) 
        'else
            
        '    call opdaterRessouceRamme(f, FY0, FY1, FY2, jobBudgetFY0(j), fctimeprisFY0(j), 0, jobBudgetFY1(j), 0, jobids(j), 0) 
       
        'end if
        

        'next


    next

    'response.write "Job opdateret...<br><br>"
    'response.flush

    'response.flush


    '*** Aktiviteter ***'
    aktids = split(request("FM_aktid"), ", ")

    'response.write request("FM_aktid")
    'response.flush

    jobaktids = split(request("FM_aktjobid"), ", ")
    aktbudgettimer = split(request("FM_aktbudgettimer"), ", ##, ")
    aktpris = split(request("FM_aktpris"), ", ##, ")

    aktBudgetFY0 = split(request("FM_akttimebudget_FY0"), ", ##, ")


    'response.write request("FM_aktbudgettimer") & "<br>"
    'response.Write request("FM_aktfctimepris_FY0") & "<br>"
    'response.write request("FM_aktfctimeprish2_FY0") & "<br>"
    'response.flush

    aktfctimeprisFY0 = split(request("FM_aktfctimepris_FY0"), ", ##, ")
    aktfctimeprish2FY0 = split(request("FM_aktfctimeprish2_FY0"), ", ##, ")

    aktBudgetFY1 = split(request("FM_akttimebudget_FY1"), ", ##, ")
    aktBudgetFY2 = split(request("FM_akttimebudget_FY2"), ", ##, ")

    for j = 0 TO UBOUND(aktids)



              aktbudgettimer(j) = replace(aktbudgettimer(j), ", ##", "")
              if len(trim(aktbudgettimer(j))) <> 0 then
              aktbudgettimer(j) = replace(aktbudgettimer(j), ".", "")
              aktbudgettimer(j) = replace(aktbudgettimer(j), ",", ".")
              else
              aktbudgettimer(j) = 0
              end if

              aktpris(j) = replace(aktpris(j), ", ##", "")
              if len(trim(aktpris(j))) <> 0 then
              aktpris(j) = replace(aktpris(j), ".", "")
              aktpris(j) = replace(aktpris(j), ",", ".")
              else
              aktpris(j) = 0
              end if


     
        '** !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! ******
        '** beregn aktpris pga timepris
        '** !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! ******

        strSQLupdateakt = "UPDATE aktiviteter SET budgettimer = " & aktbudgettimer(j) & ", aktbudget = "& aktpris(j) &" WHERE id = " & aktids(j)
        'response.write strSQLupdateakt
        'response.flush

        oConn.execute(strSQLupdateakt)


              aktBudgetFY0(j) = replace(aktBudgetFY0(j), ", ##", "")
              if len(trim(aktBudgetFY0(j))) <> 0 then
              aktBudgetFY0(j) = replace(aktBudgetFY0(j), ".", "")
              aktBudgetFY0(j) = replace(aktBudgetFY0(j), ",", ".")
              else
              aktBudgetFY0(j) = 0
              end if


              aktfctimeprisFY0(j) = replace(aktfctimeprisFY0(j), ", ##", "")
              if len(trim(aktfctimeprisFY0(j))) <> 0 then
              aktfctimeprisFY0(j) = replace(aktfctimeprisFY0(j), ".", "")
              aktfctimeprisFY0(j) = replace(aktfctimeprisFY0(j), ",", ".")
              else
              aktfctimeprisFY0(j) = 0
              end if


              'aktfctimeprish2FY0(j) = replace(aktfctimeprish2FY0(j), ", ##", "")
              'if len(trim(aktfctimeprish2FY0(j))) <> 0 then
              'aktfctimeprish2FY0(j) = replace(aktfctimeprish2FY0(j), ".", "")
              'aktfctimeprish2FY0(j) = replace(aktfctimeprish2FY0(j), ",", ".")
              'else
              'aktfctimeprish2FY0(j) = 0
              'end if


              aktBudgetFY1(j) = replace(aktBudgetFY1(j), ", ##", "")
              if len(trim(aktBudgetFY1(j))) <> 0 then
              aktBudgetFY1(j) = replace(aktBudgetFY1(j), ".", "")
              aktBudgetFY1(j) = replace(aktBudgetFY1(j), ",", ".")
              else
              aktBudgetFY1(j) = 0
              end if


              aktBudgetFY2(j) = replace(aktBudgetFY2(j), ", ##", "")
              if len(trim(aktBudgetFY2(j))) <> 0 then
              aktBudgetFY2(j) = replace(aktBudgetFY2(j), ".", "")
              aktBudgetFY2(j) = replace(aktBudgetFY2(j), ",", ".")
              else
              aktBudgetFY2(j) = 0
              end if



        for f = 0 to 2
        '** RAMME FY0-FY2 **'

          
            
                call opdaterRessouceRamme(f, FY0, FY1, FY2, aktBudgetFY0(j), aktfctimeprisFY0(j), 0, aktBudgetFY1(j), aktBudgetFY2(j), jobaktids(j), aktids(j))
       
       

        'call opdaterRessouceRamme(f, FY0, FY1, FY2, aktBudgetFY0(j), aktfctimeprisFY0(j), aktfctimeprish2FY0(j), aktBudgetFY1(j), aktBudgetFY2(j), jobaktids(j), aktids(j))

        next

    next



    'response.write "Aktiviteter opdateret...<br><br>"
    'response.end

    response.redirect "timbudgetsim.asp?FM_fy="&FY0&"&FM_visrealprdato="& visrealprdato &"&FM_sog="&sogVal
   
    case "opdfc"




    '*** Job ****'
    jobid = request("jobid")
   
    '** FY ÅRSTAL 
    FY0 = request("FY0") 
   
    visrealprdato = visrealprdato
     

    medids = split(request("FM_medid"), ", ")
    h1s = split(request("FM_h1"), ", ")

    select case lto
    case "oko", "intranet - local" 

    case else

    if cint(timesimtp) = 1 AND cint(visKunFCFelter) <> 1 then
    'if cint(timesimtp) = 1 then
    timepris = split(request("FM_tp"), ", ")
    end if

    end select
    'h2s = split(request("FM_h2"), ", ")

    h1_jobids = split(request("FM_H1_jobid"), ", ")
    h1_aktids = split(request("FM_H1_aktid"), ", ")

    'h2_jobids = split(request("FM_H2_jobid"), ", ")
    'h2_aktids = split(request("FM_H2_aktid"), ", ")
   

    aar1 = FY0

    select case lto
    case "wwf"
    md1 = "7"
    case else
    md1 = "1"
    end select

    aar2 = FY0+1
    md2 = "1"



    'response.write "aar1: " & aar1
    'response.end

      for m = 0 TO UBOUND(medids)

                

                if len(trim(h1s(m))) <> 0 then
                h1s(m) = h1s(m)
                else
                h1s(m) = 0
                end if

                'if len(trim(h2s(m))) <> 0 then
                'h2s(m) = h2s(m)
                'else
                'h2s(m) = 0
                'end if

                select case lto
                case "oko", "intranet - local" 

                case else

                    if cint(timesimtp) = 1 AND cint(visKunFCFelter) <> 1 then
                    timepris(m) = replace(timepris(m), ".", "")
                    timepris(m) = replace(timepris(m), ",", ".")
                    end if

                end select


                h1s(m) = replace(h1s(m), ".", "")
                h1s(m) = replace(h1s(m), ",", ".")

                'h2s(m) = replace(h2s(m), ".", "")
                'h2s(m) = replace(h2s(m), ",", ".")
        
                fc1Findes = 0


                '* KUN på aktiviteter
                if h1_aktids(m) <> 0 then

                '*** RENSER UD I FORECAST FORETAGET FRA RESSOURCEFORECAST SIDEN PÅ UGE NIVEAU
                '*** KUN TIMER I JULI for ikke at slette fra andre FY 
                 strSQLdelresUge = "DELETE FROM ressourcer_md WHERE jobid = " & h1_jobids(m) & " AND aktid = "& h1_aktids(m) &" AND medid = "& medids(m) &" AND (aar = "& aar1 &" AND md = "& md1 &") AND uge <> 0"
                    'response.write strSQLupdmedres & "<br>"
                    'response.flush
                oConn.execute(strSQLdelresUge)
                '*****************************************************************************



                strSQLdelmedres = "SELECT id FROM ressourcer_md WHERE jobid = " & h1_jobids(m) & " AND aktid = "& h1_aktids(m) &" AND medid = "& medids(m) &" AND aar = "& aar1 &" AND md = "& md1
                'response.Write strSQLdelmedres 
                'response.End
    
                oRec3.open strSQLdelmedres, oConn, 3
                if not oRec3.EOF then 

                    
                    if h1s(m) > 0 then
                    strSQLupdmedres = "UPDATE ressourcer_md SET timer = "& h1s(m) &" WHERE id = " & oRec3("id")  
                    'response.write strSQLupdmedres & "<br>"
                    'response.flush

                    oConn.execute(strSQLupdmedres)
                    end if

                    if h1s(m) = 0 then
                    strSQLdelmedres = "DELETE FROM ressourcer_md WHERE id = " & oRec3("id")  
                    'response.write strSQLupdmedres & "<br>"
                    'response.flush

                    oConn.execute(strSQLdelmedres)
                    end if


                fc1Findes = 1

                end if
                oRec3.close

            
                '**** Indsætter nye hvsi ikke findes
                if cint(fc1Findes) = 0 then


                    if h1s(m) > 0 then
                    strSQLinsmedres = "INSERT INTO ressourcer_md (jobid, aktid, medid, aar, md, timer) VALUES ("& h1_jobids(m) &", "& h1_aktids(m) &", "& medids(m) &", "& aar1 &", "& md1 &", "& h1s(m) &")"  
                    'response.write strSQLinsmedres & "<br>"
                    'response.flush

                    oConn.execute(strSQLinsmedres)
                    end if


                end if


                h2inUse = 999
                if cint(h2inUse) = 1 then

                    fc2Findes = 0

                    strSQLdelmedres = "SELECT id FROM ressourcer_md WHERE jobid = " & h2_jobids(m) & " AND aktid = "& h2_aktids(m) &" AND medid = "& medids(m) &" AND aar = "& aar2 &" AND md = "& md2
                    oRec3.open strSQLdelmedres, oConn, 3
                    if not oRec3.EOF then 

                    
                        if h2s(m) > 0 then
                        strSQLupdmedres = "UPDATE ressourcer_md SET timer = "& h2s(m) &" WHERE id = " & oRec3("id")  
                        'response.write strSQLupdmedres & "<br>"
                        'response.flush

                        oConn.execute(strSQLupdmedres)
                        end if


                        if h2s(m) = 0 then
                        strSQLdelmedres = "DELETE FROM ressourcer_md WHERE id = " & oRec3("id")  
                        'response.write strSQLupdmedres & "<br>"
                        'response.flush

                        oConn.execute(strSQLdelmedres)
                        end if


                    fc2Findes = 1

                    end if
                    oRec3.close

                end if 'h2inUse


            
                 '**** Indsætter nye hvsi ikke findes
               

                 if cint(fc2Findes) = 0 AND cint(h2inUse) = 1 then

                    if h2s(m) > 0 then
                    strSQLinsmedres = "INSERT INTO ressourcer_md (jobid, aktid, medid, aar, md, timer) VALUES ("& h2_jobids(m) &", "& h2_aktids(m) &", "& medids(m) &", "& aar2 &", "& md2 &", "& h2s(m) &")"  
                    'response.write strSQLinsmedres & "<br>"
                    'response.flush

                    oConn.execute(strSQLinsmedres)
                    end if

                end if


                 if cint(timesimtp) = 1 AND cint(visKunFCFelter) <> 1 then
                
                     '** LTO
                     select case lto
                     case "oko", "intranet - local" '** Opdater ikke timepriser


                     case "xxx"'** Opdater timeprsier
            
                      '** Problem med at timepris bliver sat = 0. Tjek for fejl. 20160119
                     strSQLmedtp = "UPDATE timepriser SET 6timepris = '"& timepris(m) &"' WHERE jobid = " & h1_jobids(m) & " AND aktid = "& h1_aktids(m) &" AND medarbid = "& medids(m)  
                     oConn.execute(strSQLmedtp)


                     case else

                    

                     end select

                 end if


                end if' h1_aktids(m) <> 0


            next


    'response.write "Medarbejdere opdateret...<br><br>"
    'response.end


    
        

     response.redirect "timbudgetsim.asp?FM_fy="&FY0&"&jobid="&jobid&"&func=forecast&FM_visrealprdato="& visrealprdato
    
   case "forecast"


       


        'response.write "<div style=""position:absolute; top:200px; left:200px; border:10px #999999 solid; z-index:10000000000000000000000;"">visrealprdato: "& visrealprdato & "</div>"
    
        visrealprdatoStartSQL = h1aar & "-7-1"
        visrealprdatoSQL = year(visrealprdato) &"-"& month(visrealprdato) &"-"& day(visrealprdato)

         jbs = 5000
         akts = 50000
         p = 14
         m = 120
         mhigh = 0
         phigh = 0
         dim antalm, antalp, h1_medTot, h2_medTot, h1_medTotGT, h2_medTotGT, h1_medTotGTGT
         public thisMedarbRealTimerGT, thisMedarbRealTimerGTAndre, aktbgtTimerArr, jobbgtTimerArr
         redim antalm(m,2), antalp(p), h1_medTot(m), h2_medTot(m), h1_medTotGT(m), h2_medTotGT(m), h1_medTotGTGT(m), thisMedarbRealTimerGT(m), thisMedarbRealTimerGTAndre(m)
         redim aktbgtTimerArr(akts), jobbgtTimerArr(jbs)
         public mhigh, phigh


        if len(trim(request("jobid"))) <> 0 then
        jobid = request("jobid")
        else
        jobid = 0
        end if

        if request("issubmitted") = "1" AND len(trim(request("FM_progrpid"))) = 0 then 'DiSabled = -1

                progrpid = 0
                visprogrpid = -1 'progrpid
                response.Cookies("EPR")("progrpid") = progrpid

        else
        
            if len(trim(request("FM_progrpid"))) <> 0 then
            
                progrpid = request("FM_progrpid")
                visprogrpid = progrpid
                response.Cookies("EPR")("progrpid") = progrpid
        
            else
            
                if request.Cookies("EPR")("progrpid") <> "" AND request.Cookies("EPR")("progrpid") <> "0" AND request("issubmitted") <> "1" then
                progrpid =  request.Cookies("EPR")("progrpid")
                visprogrpid = progrpid
                else
                progrpid = 0
                visprogrpid = 0
                end if

            end if

        end if

        
        'response.write "visprogrpid "& visprogrpid & " progrpid: "& progrpid & " AND "& request("FM_progrpid")


        if len(trim(request("FM_minit"))) <> 0 then
            minit = request("FM_minit")
            viskunminit = 1
            response.Cookies("EPR")("simminit") = minit
        else
            
            if request.Cookies("EPR")("simminit") <> "" AND request("issubmitted") <> "1" then
            minit =  request.Cookies("EPR")("simminit")
            viskunminit = 1
            else
            minit = ""
            viskunminit = 0
            end if

        end if

        if len(trim(minit)) <> 0 then
        minitSQlkri = " AND init LIKE '"& minit &"'"
        else
        minitSQlkri = ""
        end if


        onlyThisMedarbids = ""
        thisJobnr = 0
        
            if cint(filterKunAktiveMedarb) = 1 then

                strSQLjobnr = "SELECT jobnr FROM job WHERE id = "& jobid
                oRec3.open strSQLjobnr, oConn, 3
                if not oRec3.EOF then
        
                thisJobnr = oRec3("jobnr")

                end if
                oRec3.close

            end if


        if cint(filterKunAktiveMedarb) = 1 then

            if cint(filtervisallejobvlgtmedarb) = 1 then

                jobRealAlleSQLkri = "tjobnr <> '0'"
                jobFCAlleSQLkri = "jobid <> 0"

            else

                jobRealAlleSQLkri = "tjobnr = '" & thisJobnr & "'"
                jobFCAlleSQLkri = "jobid = "& jobid &""

            end if

        
        visJobdatoStartSQL = dateAdd("yyyy", -4, visrealprdatoStartSQL) '3 år tilbage
        visJobdatoStartSQL = year(visJobdatoStartSQL) & "-" & month(visJobdatoStartSQL) & "-" & day(visJobdatoStartSQL)


        '** REAL Timer --> Kun medarb. med timer på
        onlyThisMedarbids = " AND (mid = 0 "
        onlyThisMedarbidsFC = " AND (medarbid = 0 "
        strSQLmedids = "SELECT tmnr FROM timer "_
        & "LEFT JOIN job AS j ON (j.jobnr = tjobnr AND jobstatus = 1 AND jobstartdato >= '"& visJobdatoStartSQL &"')"_
        &" WHERE "& jobRealAlleSQLkri &" AND tdato BETWEEN '"& visrealprdatoStartSQL &"' AND '"& visrealprdatoSQL &"' AND (tfaktim = 1 OR tfaktim = 2) AND jobstatus = 1 AND jobstartdato >= '"& visJobdatoStartSQL &"' GROUP BY tmnr"
           
        'response.Write strSQLmedids
        'response.flush
    
        oRec3.open strSQLmedids, oConn, 3
        while not oRec3.EOF 
        
        onlyThisMedarbids = onlyThisMedarbids & " OR mid = "& oRec3("tmnr") 

        oRec3.movenext
        wend
        oRec3.close


        '*** FORECAST ***
        strSQLmedids = "SELECT medid FROM ressourcer_md WHERE "& jobFCAlleSQLkri &" AND (aar = "& h1aar &" AND md = "& h1md &") GROUP BY medid"
         oRec3.open strSQLmedids, oConn, 3
        while not oRec3.EOF 
        

        if instr(onlyThisMedarbids, "#"& oRec3("medid") &"#") = 0 then
        onlyThisMedarbids = onlyThisMedarbids & " OR mid = "& oRec3("medid") 
        end if

        if instr(onlyThisMedarbidsFC, "#"& oRec3("medid") &"#") = 0 then
        onlyThisMedarbidsFC = onlyThisMedarbidsFC & " OR medarbid = "& oRec3("medid") 
        end if

        oRec3.movenext
        wend
        oRec3.close

        onlyThisMedarbids = onlyThisMedarbids & ")"
        onlyThisMedarbidsFC = onlyThisMedarbidsFC & ")"

        end if

    %>

        
      

        <div id="wrapper">
        <div class="to-content-hybrid-fullzize" style="position:absolute; top:20px; left:20px; background-color:#FFFFFF;">


           

     <div id="load" style="position:absolute; display:; visibility:visible; top:260px; left:400px; width:300px; background-color:#ffffff; border:10px #cccccc solid; padding:20px; z-index:100000000;">

	    <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	    <img src="../ill/outzource_logo_200.gif" />
	    <br />
	    Forventet loadtid:
	    <%

	    'exp_loadtid = 30
	    'exp_loadtid = (((len(akttype_sel) / 3) * (len(antalvlgM) / 3)) / 50)  
        if cint(filtervisallejobvlgtmedarb) <> 1 then%> 
	    <b>ca. 5 sek.</b>
            <%else %>
            <b>ca. 10-20 sek.</b>
        <%end if %>
	    </td><td align=right style="padding-right:40px;">
	    <img src="../inc/jquery/images/ajax-loader.gif" />
	
	    </td></tr></table>

	    </div>

        <%response.flush 


          
                if cint(timesimh1h2) = 1 then %>
                <div class="container" style="width:1800px">
                <%else %>
                <div class="container" style="width:1280px">
                <%end if %>

                   <div class="portlet">
                       <form method="post" action="timbudgetsim.asp?func=forecast&issubmitted=1">
                           <input type="hidden" name="FY0" value="<%=year(y0)%>" />
                            <input type="hidden" name="FY1" value="<%=year(y1)%>" />
                            <input type="hidden" name="FY2" value="<%=year(y2)%>" />
                            
                       
                           
                        <h3 class="portlet-title">
                            <u>Timebudget simulering <span style="font-size:14px; font-weight:lighter; letter-spacing:0.5px;"> - Forecast medarbejdere</span></u>
                         </h3>
                              <div class="portlet-body">
                        <!-- NAVN / SORTERING / ID -->
                        <section>
                            <div class="well">
                             <div class="row">
                                     <div class="col-lg-4 pad-t5">Job:<br /> <select name="jobid" id="jq_jobid" class="form-control input-small" onchange="submit();">
                                         <%strSQLjob = "SELECT jobnavn, jobnr, id FROM job WHERE jobstatus = 1 AND (risiko > -1 OR risiko = -3) ORDER BY jobnavn" 
                                             
                                            oRec.open strSQLjob, oConn, 3
                                            while not oRec.EOF 

                                             if cint(jobid) = cint(oRec("id")) then
                                             jSel = "SELECTED"
                                             else
                                             jSel = ""
                                             end if

                                             %>
                                             <option value="<%=oRec("id") %>" <%=jSel %>><%=oRec("jobnavn") & " ("& oRec("jobnr") &")" %></option>
                                             <%
                                             oRec.movenext
                                             wend 
                                            oRec.close%>

                                                                       </select></div>
                                  
                                       </div>
                                 <div class="row">
                                      <div class="col-lg-2 pad-t5">Vis projektgruppe:<br />
                                          <select name="FM_progrpid" id="progrpid" class="form-control input-small" onchange="submit();">
                                              <option value="0">Alle</option>

                                          
                                          <%strSQLprogrp = "SELECT p.navn AS pgrpnavn, p.id AS pid FROM projektgrupper AS p WHERE p.orgvir = 1 ORDER BY p.navn" 

                                           oRec.open strSQLprogrp, oConn, 3
                                            while not oRec.EOF 

                                             if cint(progrpid) = cint(oRec("pid")) then
                                             pSel = "SELECTED"
                                             else
                                             pSel = ""
                                             end if

                                             %>
                                             <option value="<%=oRec("pid") %>" <%=pSel %>><%=oRec("pgrpnavn")%></option>
                                             <%
                                             oRec.movenext
                                             wend 
                                            oRec.close%>
                                              
                                              
                                              </select>

                                      </div>
                                     <div class="col-lg-5 pad-t5"><br /><input type="checkbox" id="viskunemedarbfcreal" name="FM_viskunmedarbMtimFc" value="1" <%=filterKunAktiveMedarbCHK %> /> Vis kun medarbejdere med forecast eller realiserede timer på i FY</div>
                                  
                                     </div>
                                     <div class="row">
                                      <div class="col-lg-2 pad-t5">Medarb. initaler: <input type="text" name="FM_minit" id="minit" class="form-control input-small" style="text-align:left;" value="<%=minit %>" placeholder="Init" /></div> 
                                          
                                          
                                          <div class="col-lg-4 pad-t5 pad-r30"><br /><input type="checkbox" name="FM_visallejobvlgtmedarb" id="allejob" value="1" <%=filtervisAlleJobVlgtMedarbCHK %> /> Vis alle job for valgte medarbejder / projektgrp. <span style="color:#999999; font-size:11px;"><br />(Projektgrp. kun overblik, kan ikke redigeres)</span></div>
                                
                                         </div>

                             <div class="row">

                                 <%call fy_relprdato_fm %>

                            </div>
                             <div class="clearfix"></div>

                            </div>
                              </section>
                              </div>
                            </form>






                           <form method="post" action="timbudgetsim.asp?func=opdfc">
                           <input type="hidden" id="viskunminit" value="<%=viskunminit%>" />
                           <input type="hidden" id="visallejob" value="<%=filtervisallejobvlgtmedarb%>" />
                           <input type="hidden" name="FM_visrealprdato" value="<%=visrealprdato %>" />
                           <input type="hidden" name="FY0" value="<%=year(y0)%>" />
                           <input type="hidden" name="jobid" value="<%=jobid%>" />

                            
                            <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button><br />&nbsp;
                             
                        
     <%'if lcase(sogVal) = "all2" then
     dvlft = 0
     alft = 150
     'else
     'dvlft = 1250
     'alft = 1250
     'end if 


    
 '********************************************************************************* 
'*** Henter job og aktiviteter TIL FORECASTSIDEN  ************************************** *****'
'*********************************************************************************
i = 0

if cint(filtervisallejobvlgtmedarb) = 1 then
'if cint(filtervisallejobvlgtmedarb) <> 10 then

visJobdatoStartSQL = dateAdd("yyyy", -4, visrealprdatoStartSQL) '3 år tilbage
visJobdatoStartSQL = year(visJobdatoStartSQL) & "-" & month(visJobdatoStartSQL) & "-" & day(visJobdatoStartSQL)

strSQLkrijob = " j.id <> 0" 
strSQLkrijobDatoKri = "AND jobstartdato > '"& visJobdatoStartSQL &"'"
else
strSQLkrijob = " j.id = "& jobid
strSQLkrijobDatoKri = ""
end if
    

lastknavn = ""
lastjobnavn = ""
lastFase = ""
strSQLjob = "SELECT jobnavn, jobnr, j.id AS jid, jobknr, j.budgettimer AS jobbudgettimer, jo_gnstpris, "_
&" kkundenavn, k.kid, a.id AS aid, a.navn AS aktnavn, a.budgettimer AS aktbudgettimer, aktbudget, aktbudgetsum, a.fase, jobtpris FROM job AS j "_
&" LEFT JOIN kunder AS k ON (kid = jobknr) "_
&" LEFT JOIN aktiviteter AS a ON (a.job = j.id AND (a.fakturerbar = 1 OR a.fakturerbar = 2)) "_
&" WHERE "& strSQLkrijob &" AND (j.risiko > -1 OR j.risiko = -3) AND jobstatus = 1 AND a.navn IS NOT NULL "& strSQLkrijobDatoKri &""_
&" GROUP BY a.id ORDER BY jobnavn, jobnr, a.fase, a.sortorder, a.navn LIMIT 5000"

'"& jobid &"
'** SKAL AKTIVITET VÆRE AKTVI?
'** strSQLkrijob
'response.write strSQLjob
'response.Flush
x = 0
oRec.open strSQLjob, oConn, 3
while not oRec.EOF
    


        '************* Aktvitetslinjer ************** 
          if cint(filtervisallejobvlgtmedarb) = 1 then
           aktLinjerDsp = "none"
           aktLinjerWsb = "hidden"
           jobPlusMinusiconLnk = "+"
           cssBgCol = "#FFFFFF"
          else
           aktLinjerDsp = ""
           aktLinjerWsb = "visible"
           jobPlusMinusiconLnk = "-"
           cssBgCol = "#FFFFE1"
          end if


         if lastjobnavn <> lcase(oRec("jobnavn")) then


              strJobTxtTds(oRec("jid")) = strJobTxtTds(oRec("jid")) & "<tr id='tr_job_"& oRec("jid") &"' style='background-color:"& cssBgCol &";'><td style='white-space:nowrap;'><span style=""color:#5582d2;"" id='an_"& oRec("jid") &"' class='fp_jid'><b>["& jobPlusMinusiconLnk &"]</b></span>&nbsp;<b>"& left(oRec("jobnavn"), 20) &"</b> ("& oRec("jobnr") &") "_
              &"</td><td style='white-space:nowrap; text-align:right;'>"& formatnumber(oRec("jobbudgettimer"), 0) &" t.</td>"

              jobbgtTimerArr(oRec("jid")) = formatnumber(oRec("jobbudgettimer"), 0)
       
              i = i + 1
              x = x + 1

        end if 
               
         
          
          
         




        if lastFase <> lcase(oRec("fase")) AND len(trim(oRec("fase"))) <> 0 then 

         if cint(visKunFCFelter) = 0 then
         cspFaser = 1
         else
         cspFaser = 1
         end if

        'strAktTxtTds(oRec("aid")) = strAktTxtTds(oRec("aid")) & "<tr style=""display:"& aktLinjerDsp &"; visibility:"& aktLinjerWsb &";"" class='tr_"& oRec("jid") &"'><td colspan=""100"">&nbsp;</td></tr>"
        strAktTxtTds(oRec("aid")) = strAktTxtTds(oRec("aid")) & "<tr style=""display:"& aktLinjerDsp &"; visibility:"& aktLinjerWsb &";"" class='tr_"& oRec("jid") &"'><td colspan="& cspFaser &" style='padding-left:30px;'><b>"& replace(oRec("fase"), "_", " ")  &"</b></td></tr>"
        
        end if
       
        if isNull(oRec("aktbudgettimer")) <> true then
        aktbgtTimer = formatnumber(oRec("aktbudgettimer"), 0)
        else
        aktbgtTimer = 0
        end if 

        aktbgtTimerArr(oRec("aid")) = aktbgtTimer
        
        strAktTxtTds(oRec("aid")) = strAktTxtTds(oRec("aid")) &  "<tr class='tr_"& oRec("jid") &"' style=""display:"& aktLinjerDsp &"; visibility:"& aktLinjerWsb &";""><td style='white-space:nowrap; padding-left:40px;'>"& left(oRec("aktnavn"), 20) &" "_
        &"</td><td style='white-space:nowrap; text-align:right;'>"& aktbgtTimer &" t.</td>"
        '&" <input type='hidden' id='bgt_akt_"& oRec("aid") &"' value='"& aktbgtTimer &"'>"
       

        if isNull(oRec("fase")) <> true then
        lastFase = lcase(oRec("fase"))
        else
        lastFase = ""
        end if

    i = i + 1
    lastknavn = lcase(oRec("kkundenavn"))
    lastjobnavn = lcase(oRec("jobnavn")) 
    'lastJid = oRec("jid")
    x = x + 1
oRec.movenext    
wend 
oRec.close 
      
        
public medarbIPgrp
redim medarbIPgrp(1000)  
            
        
    %>




<table class="table table-bordered" style="background-color:#FFFFFF;">
    
     <thead>
         <tr> 

            <th style="width:10%;">Job (nr) / Akt.</th>
            <th style="width:3%;">Budget<br /><span style="font-size:9px;">FY GT</span></th>
            <th style="width:3%;">Forecast<br /><span style="font-size:9px;">FY GT</span></th>
            <th style="width:3%; white-space:nowrap;">Real.<br /><span style="font-size:9px;">Pr. dato GT</span></th> 
            <th style="width:3%; white-space:nowrap;">Saldo<br /><span style="font-size:9px;">Pr. dato GT</span></th> 
           <%
             
             if cint(progrpid) <> 0 then  
             progrpidSQLkri = " AND (p.id = "& progrpid &")"
             else
             progrpidSQLkri = ""
             end if

            if cint(filterKunAktiveMedarb) = 1 then
             prgrRelMids =  replace(onlyThisMedarbids, "mid", "medarbejderId")
            else
             prgrRelMids = " AND (medarbejderId <> 0) "
            end if

             strSQLMedarb = "SELECT mnavn, init, mid, p.navn AS pgrpnavn, p.id AS pid FROM projektgrupper AS p "_
             &"LEFT JOIN progrupperelationer AS pr ON (projektgruppeId = p.id "& prgrRelMids &")"_
             &"LEFT JOIN medarbejdere ON (mid = medarbejderId) WHERE mansat <> 2 "& medarbinitSQL &" AND projektgruppeId <> 10 "_
             &" AND p.orgvir = 1 AND p.id <> 10 AND init <> '' "& minitSQlkri &" "& progrpidSQLkri &"  ORDER BY pgrpnavn, mnavn LIMIT 80"
            
               'response.write strSQLMedarb
               'response.Flush
              
              lastAfd = ""
              'antalm = 0
               p = 0
               m = 0
             progrpTot = 0
             oRec5.open strSQLMedarb, oConn, 3
             while not oRec5.EOF
                  

              
               
                if lastAfd <> oRec5("pgrpnavn") then
               
                if cint(visKunFCFelter) <> 1 then
                progrpCps = 2
                else 
                progrpCps = 1
                end if

                'colspan="<%=progrpCps
                %>
          
                 <th style="width:150px; white-space:nowrap;" class="td_p" id="td_p_<%=p %>"><span class="sp_p" id="sp_p_<%=p %>"><u><%=left(oRec5("pgrpnavn"), 20) %></u> 


                  

                     <%if (cint(visprogrpid) = 0) AND cint(filtervisallejobvlgtmedarb) = 1 then
                         
                     else %>
                     <span style="color:#5582d2;" id="an_p_<%=p %>">[+]</span></span>
                     <%end if %>


                     
                         <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1 then %>
                         <br />
                         <span style="font-size:9px;">Forecast FY </span>
                         </th>
                         <%end if %>

                   


                         <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 2 then 
                             
                         if cint(visKunFCFelter) = 0 then%>
                         <th><span style="font-size:9px; color:#999999;"><%=left(oRec5("pgrpnavn"), 10)%></span>
                         <%end if %>

                         <br /><span style="font-size:9px;">Real. dt.</span></th>
                         <%end if %>

                         <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 3 then
                          if cint(visKunFCFelter) = 0 then%>
                         <th><span style="font-size:9px; color:#999999;"><%=left(oRec5("pgrpnavn"), 10)%></span>
                         <%end if %>
                         <br /><span style="font-size:9px;">Saldo</span></th>
                         <%end if %>

                    
                <%


                ' if cint(filtervisallejobvlgtmedarb) = 1 OR len(trim(minit)) <> 0 then
                ' if p = 0 then
                ' pvzb = "visible"
                ' pdsp = ""
                ' else
                 pvzb = "hidden"
                 pdsp = "none"
                'end if

                antalp(p) = oRec5("pid")
                p = p + 1
                phigh = p
             

               
                end if 


               
                'redim preserve antalm(m,2)
                 'antalm(p,0) = oRec5("pid")
                 antalm(m,0) = oRec5("pid") 
                 antalm(m,1) = oRec5("mid")

                 medarbIPgrp(p-1) = medarbIPgrp(p-1) & " OR medid = "& oRec5("mid")

                 '** Norm 6 md ***'
                 ntimPer = 0
                 'call normtimerPer(oRec5("mid"), "1-1-"& datepart("yyyy", y0, 2,2), 30, 0)
                 'ntimPer = ntimPer * 6
                 select case lto
                 case "oko"
                 ntimPer = 1650
                 case else
                 ntimPer = 1604 '1582, 906
                 end select
                 antalm(m,2) = ntimPer


                if (cint(visprogrpid) = 0) AND cint(filtervisallejobvlgtmedarb) = 1 then

                else
                

                 if (cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1) then

                    if cint(timesimtp) = 1 then
                    %>
                    
                     <th class="alt afd_p_<%=p-1 %> afd_p" style="white-space:nowrap; visibility:<%=pvzb%>; display:<%=pdsp%>; vertical-align:bottom; background-color:#FFFFFF;"><b><%="[ "& oRec5("init") & " ]" %></b>
                         <br /><span style="font-size:9px;">Tpris.</span></th>
                         <th class="alt afd_p_<%=p-1 %> afd_p" valign="bottom" style="visibility:<%=pvzb%>; vertical-align:bottom; display:<%=pdsp%>; background-color:#FFFFFF;"><span style="font-size:9px;">Forecast <br>Timer</span></th>
               
                    <%else %>
                     <th class="alt afd_p_<%=p-1 %> afd_p" valign="bottom" style="visibility:<%=pvzb%>; vertical-align:bottom; display:<%=pdsp%>; background-color:#FFFFFF;"><b><%="[ "& oRec5("init") & " ]" %></b>
                        <br /><span style="font-size:9px;">Forecast</span></th>
               
                    <%end if 'vistp %>

                <%end if  %>


                <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 2  then %>
                <th class="alt afd_p_<%=p-1 %> afd_p" valign="bottom" style="visibility:<%=pvzb%>; vertical-align:bottom; display:<%=pdsp%>; background-color:#FFFFFF;">
                    <%if cint(visKunFCFelter) = 2 then %>
                    <b><%="[ "& oRec5("init") & " ]" %></b><br />
                    <%end if %>
                    <!--H2--><span style="font-size:9px;">Real. dt.</span></th>
                <%end if %>

                <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 3  then %>
                <th class="afd_p_<%=p-1 %> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>; vertical-align:bottom; background-color:#EFF3FF;">
                     <%if cint(visKunFCFelter) = 3 then %>
                    <b><%="[ "& oRec5("init") & " ]" %></b><br />
                    <%end if %>
                    <span style="font-size:9px;">Sal. Fc/Real.</span></th>
                <%end if %>

             <% end if'  if visprogrpid = 1 AND cint(filtervisallejobvlgtmedarb) = 1 then

                 'minits(m) = oRec5("init")
                m = m + 1
                mhigh = m

                

                lastAfd = oRec5("pgrpnavn")

             oRec5.movenext
             wend
             oRec5.close%>

         </tr>
    </thead>

<% 

'********************************************************************
'*** MEDARBEJDER I Projekgruppe LOOP ********************************
'********************************************************************
'
'public medarbIPgrp
'redim medarbIPgrp(1000)
'for p = 0 to phigh - 1
'medarbIPgrp(p) = ""
'medarbgrpIdSQLkri = ""
'call medarbiprojgrp(antalp(p), 0, 0, -1)
'medarbIPgrp(p) = replace(medarbgrpIdSQLkri, "mid", "medid")
'next

'm = mhigh 'stille tilbage til før loop medarbiprojgrp()

'response.write "M: "& m

'********************************************************************
'*** MEDARBEJDER LOOP ***********************************************
'********************************************************************

lastknavn = ""
lastjobnavn = ""
strSQLjob = "SELECT jobnavn, jobnr, j.id AS jid, jobknr, "_
&" kkundenavn, k.kid, a.id AS aid, a.fase, jo_gnstpris, aktbudget FROM job AS j "_
&" LEFT JOIN kunder AS k ON (kid = jobknr) "_
&" LEFT JOIN aktiviteter AS a ON (a.job = j.id AND (a.fakturerbar = 1 OR a.fakturerbar = 2)) "_
&" WHERE "& strSQLkrijob &" AND (j.risiko > -1 OR j.risiko = -3) AND jobstatus = 1 AND a.navn IS NOT NULL "& strSQLkrijobDatoKri &""_
&" GROUP BY a.id ORDER BY jobnavn, jobnr, a.fase, a.sortorder, a.navn LIMIT 5000"
'"& jobid &"
'response.write strSQLjob
'response.Flush

p = 0
x = 0
lastFase = ""
oRec.open strSQLjob, oConn, 3
while not oRec.EOF
    
   
                
            '************* Joblinjer ************** %>
            <%if lastjobnavn <> lcase(oRec("jobnavn")) then 
            %><tr><%

            response.write(strJobTxtTds(oRec("jid")))

           
            call medarbfelter(oRec("jobnr"), oRec("jid"), 0, h1aar, h2aar, h1md, h2md, oRec("jo_gnstpris")) 
               
            %>
             </tr>
            <% 'i = i + 1
            end if
               

            'select case right(x, 1)
            'case 0,2,4,6,8
            bgthis = "#FFFFFF"
            'case else
            'bgthis = "#F3F3F3"
            'end select%>


      <%'************* Aktvitetslinjer ************** 
          if cint(filtervisallejobvlgtmedarb) = 1 then
          aktLinjerDsp = "none"
          aktLinjerWsb = "hidden"
          else
          aktLinjerDsp = ""
          aktLinjerWsb = "visible"
          end if
          
          %>
   
     
    <tr class="tr_<%=oRec("jid") %>" style="background-color:<%=bgthis%>; visibility:<%=aktLinjerWsb%>; display:<%=aktLinjerDsp%>;">
            
     
            <%
                
            response.write(strAktTxtTds(oRec("aid")))    
            response.flush
             
            call medarbfelter(oRec("jobnr"), oRec("jid"), oRec("aid"), h1aar, h2aar, h1md, h2md, oRec("aktbudget")) 
             
            %>

      
     </tr>
            <% 'i = i + 1
   
            lastknavn = lcase(oRec("kkundenavn"))
            lastjobnavn = lcase(oRec("jobnavn")) 
        
                if isNull(oRec("fase")) <> true then
                lastFase = lcase(oRec("fase"))
                else
                lastFase = ""
                end if

           x = x + 1



        response.flush


        oRec.movenext    
        wend 
        oRec.close 


    '*** totaler ***
     pvzb = "hidden"
     pdsp = "none"
     %>

     <input type="hidden" value="<%=x-1 %>" id="xhigh" />

    <tr><td colspan="100"><br />Total</td></tr>

    <%if cint(filtervisallejobvlgtmedarb) <> 1 then %>

    <tr>
      <td>Total (dette job)</td>

        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      <%
  
      for p = 0 to phigh - 1
      %>
           
             <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1 then %>
             <td>&nbsp;</td>
             <%end if %>

             <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 2 then %>
             <td>&nbsp;</td>
             <%end if %>

           <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 3 then %>
             <td>&nbsp;</td>
             <%end if 

            for m = 0 TO mhigh - 1 

                if antalp(p) = antalm(m,0) then

                if antalm(m,2) < h1_medTot(antalm(m,1)) then
                bgThisGT = "Lightpink"
                else
                bgThisGT = ""
                end if

                h1h2medTot = (h1_medTot(antalm(m,1))/1 + h2_medTot(antalm(m,1))/1 )
                h1h2medTot = formatnumber(h1h2medTot, 0)

                h1_medTotGTsaldo = (h1_medTot(antalm(m,1)) - thisMedarbRealTimerGT(antalm(m,1)))
                h1_medTotGTsaldo = formatnumber(h1_medTotGTsaldo, 0)
                %>
               
                <%if cint(timesimtp) = 1 AND (cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1) then  %>
                <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;">&nbsp;</td>
                <%end if %>
                
                <%if(cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1) then %>
                <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" value="<%=h1_medTot(antalm(m,1)) %>" id="totmedarbh1_<%=antalm(m,1) %>" class="form-control input-small" style="width:45px;" /></td>
                <%end if %>

                <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 2 then %>
                <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" value="<%=formatnumber(thisMedarbRealTimerGT(antalm(m,1)), 0)%>"  class="form-control input-small" style="width:45px;" disabled /></td>
                <%end if %>        

                <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 3 then %>
                <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" value="<%=h1_medTotGTsaldo %>" class="form-control input-small" style="width:45px; background-color:#Eff3ff;" disabled /></td>
                <%end if

                end if
            next


    next
                    
    %></tr>


    <!-- TOTal på tværs af alle job -->

 

     <tr>
      <td>Total (andre job)</td>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
      <%
      
      for p = 0 to phigh - 1
      %>
           
             <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1 then %>
             <td>&nbsp;</td>
             <%end if %>

             <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 2 then %>
             <td>&nbsp;</td>
             <%end if %>

           <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 3 then %>
             <td>&nbsp;</td>
             <%end if 


            for m = 0 TO mhigh - 1 


                 '** GT FORECAST pr. medarb. H1 og H2 '*** 


                     if antalp(p) = antalm(m,0) then



                             h1GT = 0
                             h2GT = 0
                             h1h2GT = 0
                
                            if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1 then 

                             strSQLmedrd = "SELECT SUM(timer) AS sumtimer FROM ressourcer_md WHERE jobid <> "& jobid &" AND aktid <> 0 AND medid = "& antalm(m,1) &" AND ((aar = "& h1aar &" AND md = "& h1md  &") OR (aar = "& h2aar &" AND md = "& h2md  &"))  GROUP BY medid" 
                             'response.Write "strSQLmedrd: " & strSQLmedrd
                             'response.flush
                
                             oRec3.open strSQLmedrd, oConn, 3
                             if not oRec3.EOF then

                                h1GT = oRec3("sumtimer")

                             end if
                             oRec3.close

              
                            'if aktid <> 0 then
                            h1_medTotGT(antalm(m,1)) = h1_medTotGT(antalm(m,1))/1 + h1GT
                            h2_medTotGT(antalm(m,1)) = h2_medTotGT(antalm(m,1))/1 + h2GT
                            'end if
              
                            end if



              
                
                
                    '*********** Timepriser ******
                    medarbTp = 0
                    if cint(timesimtp) = 1 AND (cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1)  then

                     
                    strSQLmedtp = "SELECT 6timepris FROM timepriser WHERE jobid <> " & jobid & " AND aktid = 0 AND medarbid = "& antalm(m,1)  
                     oRec3.open strSQLmedtp, oConn, 3
                     if not oRec3.EOF then

                        medarbTp = oRec3("6timepris")

                     end if
                     oRec3.close

                    end if


                    if (cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 2) then

                    '*** Timer ANDRE JOB ***'
                  '** REAL timer pr. dato ***'
                    visrealprdatoStartSQL = h1aar & "-7-1"
                    visrealprdatoSQL = year(visrealprdato) &"-"& month(visrealprdato) &"-"& day(visrealprdato)

                 
                    strSQLrealTimerAktmedarb = "SELECT SUM(timer) AS sumtimer, SUM(timer*timepris) AS realoms, AVG(timepris) AS gnstp FROM timer WHERE tjobnr <> '" & jobnr & "' AND tmnr = "& antalm(m,1) &" AND tdato BETWEEN '"& visrealprdatoStartSQL &"' AND '"& visrealprdatoSQL &"' GROUP BY tmnr"   
                    'response.write strSQLrealTimerAktmedarb
                    'response.flush
                    oRec3.open strSQLrealTimerAktmedarb, oConn, 3
                    if not oRec3.EOF then

                    thisMedarbRealTimerGTAndre(antalm(m,1)) = oRec3("sumtimer")
                    'thisMedarbrealomsH1 = oRec3("realoms")
                    'thisMedarbgnstp = oRec3("gnstp")

                    end if
                    oRec3.close

                    end if
                   

                 if thisMedarbRealTimerGTAndre(antalm(m,1)) <> 0 then
                 thisMedarbRealTimerGTTXT = formatnumber(thisMedarbRealTimerGTAndre(antalm(m,1)), 0)
                 else
                 thisMedarbRealTimerGTTXT = ""
                 end if

                 h1_medTotGTsaldo = 0
                 h1_medTotGTsaldo = (h1_medTotGT(antalm(m,1)) - thisMedarbRealTimerGTAndre(antalm(m,1)))
                h1_medTotGTsaldo = formatnumber(h1_medTotGTsaldo, 0)
                 
                if h1_medTotGTsaldo < 0 then
                h1_medTotGTsaldoBGcol = "lightpink"
                else
                h1_medTotGTsaldoBGcol = "#Eff3ff"
                end if


                if cint(timesimtp) = 1 AND (cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1) then  %>
                <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;">&nbsp;</td>
                <%end if %>

                <% if(cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1) then %>
                <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" value="<%=h1_medTotGT(antalm(m,1)) %>" id="totmedarbh1GT_<%=antalm(m,1) %>" class="form-control input-small" style="width:45px;" /></td>
                <!--<td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" value="<%=h2_medTotGT(antalm(m,1)) %>" id="totmedarbh2GT_<%=antalm(m,1) %>"  class="form-control input-small" style="width:40px;" disabled /></td>-->
                <%end if %>


                <%if (cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 2) then %>
                <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" value="<%=thisMedarbRealTimerGTTXT %>"  class="form-control input-small" style="width:45px;" disabled /></td>
                <%end if %>
         
                <%if (cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 3) then %>
                <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" value="<%=h1_medTotGTsaldo %>" class="form-control input-small" style="width:45px; background-color:<%=h1_medTotGTsaldoBGcol%>; border:0;" disabled /></td>
                
                <%end if
                end if

            next


    next


                     
    %></tr>

    <%end if 'filtervisallejobvlgtmedarb  %>

     <tr style="background-color:lightblue;">
      <td><b>Grandtotal</b></td>

         <td>&nbsp;</td>

          <%if cdbl(jobFcGtGT) <> 0 then
           jobFcGtGTTxt = formatnumber(jobFcGtGT,0) 'Sammenregnet på aktiviteter
          else
            jobFcGtGTTxt = 0
          end if %>

        <td style="text-align:right;"><%=jobFcGtGTTxt %></td>

         <%if cdbl(thisPrgRealTimerGT) <> 0 then
             thisPrgRealTimerGTTxt = formatnumber(thisPrgRealTimerGT,0)
           else
            thisPrgRealTimerGTTxt = 0
            end if


             thisSaldoGT = formatnumber(jobFcGtGT - thisPrgRealTimerGT, 0) 
             
             if thisSaldoGT <> 0 then
             thisSaldoGTTxt = " = " & thisSaldoGT
             else
             thisSaldoGTTxt = ""
             end if
              %>

         
         <td style="text-align:right;"><%=thisPrgRealTimerGTTxt %></td>
         <td  style="text-align:left;"><%=thisSaldoGTTxt %></td>
      <%
      
      for p = 0 to phigh - 1
      %>
           
              <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1 then %>
             <td>&nbsp;</td>
             <%end if %>

             <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 2 then %>
             <td>&nbsp;</td>
             <%end if %>

           <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 3 then %>
             <td>&nbsp;</td>
             <%end if 

            for m = 0 TO mhigh - 1 


                 '** GT FORECAST pr. medarb. H1 og H2 '*** 


                     if antalp(p) = antalm(m,0) then


                        thisMedarbRealTimerGTGT = (thisMedarbRealTimerGT(antalm(m,1)) + thisMedarbRealTimerGTAndre(antalm(m,1)))
                        if len(trim(thisMedarbRealTimerGTGT)) <> 0 then
                        thisMedarbRealTimerGTGTtxt = formatnumber(thisMedarbRealTimerGTGT,0)
                        else
                        thisMedarbRealTimerGTGTTxt = ""
                        end if

                         h1_medTotGTGTsaldo = 0
                       

                    h1_medTotGTGT(antalm(m,1)) = (h1_medTotGT(antalm(m,1)) + h1_medTot(antalm(m,1)))
                    h1_medTotGTGT(antalm(m,1)) = formatnumber(h1_medTotGTGT(antalm(m,1)),0)

                     h1_medTotGTGTsaldo = (h1_medTotGTGT(antalm(m,1)) - thisMedarbRealTimerGTGT)
                     h1_medTotGTGTsaldo = formatnumber(h1_medTotGTGTsaldo ,0)

                        if h1_medTotGTGTsaldo < 0 then
                        h1_medTotGTGTsaldoBGcol = "lightpink"
                        else
                        h1_medTotGGTTsaldoBGcol = "#Eff3ff"
                        end if

                    'if cdbl(antalm(m,2)) < cdbl(h1_medTotGTGT(antalm(m,1))) then
                    'bgThisGTGT = "lightpink"
                    'else
                    'bgThisGTGT = ""
                    'end if

                %>

                 <%if cint(timesimtp) = 1 AND (cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1) then  %>
                <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;">&nbsp;</td>
                <%end if %>


                 <% if(cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1) then %>
                <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" value="<%=h1_medTotGTGT(antalm(m,1)) %>" id="totmedarbh1GTGT_<%=antalm(m,1)%>" class="form-control input-small" style="width:45px;"/></td>
                <%end if %>      

                <!--<td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" value="<%=h2_medTotGT(antalm(m,1)) %>" id="totmedarbh2GT_<%=antalm(m,1)%>"  class="form-control input-small" style="width:45px;" disabled /></td>
                <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" value="<%=h1h2medTotGT %>" class="form-control input-small" style="width:45px; background-color:#Eff3ff; border:0;" disabled /></td>
                -->

                <%if (cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 2) then %>
                 <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" value="<%=thisMedarbRealTimerGTGTtxt %>"  class="form-control input-small" style="width:45px;" disabled /></td>
                <%end if %>
         
                <%if (cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 3) then %>
                <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" value="<%=h1_medTotGTGTsaldo %>" class="form-control input-small" style="width:45px; background-color:<%=h1_medTotGTGTsaldoBGcol%>;" disabled /></td>
                <%end if
                end if

            next


    next

    %></tr>
    
    
    
    <tr>
       <td>Norm: (tilnærmet)</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      <%


     '*** Normtimer ***
     for p = 0 to phigh - 1
      %>
           
                         <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1 then %>
             <td>&nbsp;</td>
             <%end if %>

             <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 2 then %>
             <td>&nbsp;</td>
             <%end if %>

           <%if cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 3 then %>
             <td>&nbsp;</td>
             <%end if 

                
                for m = 0 TO mhigh - 1
                    
                    if antalp(p) = antalm(m,0) then %>

                     <%if cint(timesimtp) = 1 AND (cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1) then  %>
                    <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;" style="visibility:<%=pvzb%>; display:<%=pdsp%>;">&nbsp;</td>
                    <%end if %>

                     <% if(cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 1) then %>
                    <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" value="<%=antalm(m,2) %>" id="totmedarbn1_<%=antalm(m,1)%>" class="form-control input-small" style="width:45px;" disabled/></td>
                    <%end if %>

                    <%if (cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 2) then %>
                    <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" value="<%=antalm(m,2) %>" id="totmedarbn2_<%=antalm(m,1)%>" class="form-control input-small" style="width:45px;" disabled/></td>
                    <%end if %>            

                    <%if (cint(visKunFCFelter) = 0 OR cint(visKunFCFelter) = 3) then %>
                    <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;">&nbsp;</td>
                    <%end if
                    end if
                next

     next
    %></tr><%

    '*********************************************************************************************************************************






    %>
    <input type="hidden" value="<%=phigh-1 %>" id="phigh" />
    <input type="hidden" value="<%=mhigh-1 %>" id="mhigh" />

    </table>
      <%if fyStMd = 7 then
      saveMthTxt ="Juli"
    else
      saveMthTxt ="Jan."
    end if
    %>

<span style="color:#999999">Alle timer bliver gemt i <%=saveMthTxt %> måned i det valgte FY.</span>

    <br />
                             <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
    </form>


    <!-- </div> -->

     </div><!-- portlet body -->


     </div><!-- portlet -->
    </div><!-- container -->
    </div><!-- content -->
    </div><!-- wrapper -->

        
        
         
        <%
   case else    


    'if lcase(sogVal) = "all" then

    if len(trim(request("sogmedarb"))) <> 0 then
    medarbinits = request("sogmedarb")
    medarbinitsTXT = medarbinits
    medarbinitSQL = " AND init LIKE '"& medarbinits & "%'"
    else
    medarbinits = ""
    medarbinitSQL = ""
    medarbinitsTXT = ""
     'response.end
    end if

        if sogVal = "all" then
        sogValSQLKri = "AND (jobstatus = 1)"
        else
        sogValSQLKri = "AND (kkundenavn LIKE '"& sogVal &"%' OR jobnavn LIKE '%"& sogVal &"%' OR jobnr LIKE '"& sogVal &"%' OR jobnr = '"& sogVal &"') "
        end if

    lmt = 2500
    pdiv_vzb = "hidden" '"visible"
    pdiv_dsp = "none"

    jdiv_vzb = "visible"
    jdiv_dsp = ""

   



 %>

 

 <%call menu_2014() %>
<!-------------------------------Sideindhold------------------------------------->

    <!--<div id="div1" ondrop="drop(event)" ondragover="allowDrop(event)">Træk hertil</div>-->

<div id="wrapper">
  
<div class="to-content-hybrid-fullzize" style="position:absolute; top:102px; left:90px; background-color:#FFFFFF;">

        <div id="load" style="position:absolute; display:; visibility:visible; top:260px; left:400px; width:300px; background-color:#ffffff; border:10px #cccccc solid; padding:20px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	<br />
	Forventet loadtid:
	<%

	'exp_loadtid = 30
	'exp_loadtid = (((len(akttype_sel) / 3) * (len(antalvlgM) / 3)) / 50)  %> 
	<b>ca. 3-10 sek.</b>
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	
	</td></tr></table>

	</div>

    <%response.flush 


        
        if cint(timesimh1h2) = 1  then %>
            <div class="container" style="width:1800px">
        <%else 
            
             if cint(timesimtp) = 1 then%>
             <div class="container" style="width:1500px">
             <%else %>
             <div class="container" style="width:1280px">
            <%end if %>

        <%end if %>

           <div class="portlet">
               <form method="post" action="timbudgetsim.asp">
                <h3 class="portlet-title">
                    <u>Timebudget simulering</u>
                 </h3>

                <div class="portlet-body">
                        <!-- NAVN / SORTERING / ID -->
                        <section>
                            <div class="well">
                                <div class="row">

                                     <div class="col-lg-3 pad-t5">Kunde/Job søg:
                                         <input class="form-control input-small inputclsLeft" type="text" value="<%=sogValTxt %>" name="FM_sog" placeholder="Søg på kunde ell. jobnavn, ell. jobnr." />
                                         
                                      </div>

                                     <%call fy_relprdato_fm %>
                                      </div><!-- row -->



   
                               
       </div>
      </section>
      </div>
    </form>




<form method="post" action="timbudgetsim.asp?func=opd&FM_sog=<%=sogVal %>">
    <input type="hidden" name="FY0" value="<%=year(y0)%>" style="width:40px;" />
    <input type="hidden" name="FY1" value="<%=year(y1)%>" style="width:40px;" />
    <input type="hidden" name="FY2" value="<%=year(y2)%>" style="width:40px;" />
    <input type="hidden" name="FM_visrealprdato" value="<%=visrealprdato %>" />
    <input type="hidden" name="timesimh1h2" id="timesimh1h2" value="<%=timesimh1h2 %>" />
    
    

<!-- MAIN -->

   


<div id="d_jobakt" style="position:relative; visibility:<%=jdiv_vzb%>; display:<%=jdiv_dsp%>; left:0; top:0; overflow-x:scroll; width:100%; background-color:#ffffff;">
    
<h4>Budget</h4>

     <!-- Opdater/Annuller knapper -->
   
        <!--<div style="margin-top:15px; margin-bottom:15px;">
        
            <a class="btn btn-default" href="#" role="button" id="a_afdm">Forecast medarbejdere</a>
          
         
        </div>
    -->

   
                                <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button><br />&nbsp;
                              
    <!-- Job tabel -->
    <table class="table table-striped table-bordered" id="xtable-2">

     <thead>
           <tr>
        <th>&nbsp;</th>
        <th>&nbsp;</th>
        <th>&nbsp;</th>
        <th>A1</th>
        <th>A2</th>
        <th>A3</th>
        <th>C</th>
        <th>D</th>
        <th>E</th>
         <%if cint(timesimh1h2) = 1 then %>
        <th>F</th>
         <%end if %>
        <th>G</th>

        <%if cint(timesimtp) = 1 then%> 
        <th>H</th>
        <%end if %>

        <th>I</th>
        <th>I2</th>
        <th>I3</th>
        <%if cint(timesimh1h2) = 1 then %>
        <th>J</th>
        <th>K</th>
        <th>L</th>
        <th>M</th>
        <%end if %>

        <%if cint(timesimtp) = 1 then%> 
        <th>N</th>
        <%end if %>

        <%if cint(timesimh1h2) = 1 then %>
        <th>O</th>
        <%end if %>
         <th>P</th>

         </tr>


    </thead>


        <%select case lto 'Bør følge finasår func

            case "wwf"
            addOneYear = 1
            case else
            addOneYear = 0
            end select

             %>
    

    <thead>
          <tr>
        <th style="vertical-align:bottom;">Job</th>
        <th style="vertical-align:bottom;">Jobnr.</th>
        <th style="vertical-align:bottom;">Aktivitet</th>
        <th style="vertical-align:bottom;">Bgt.timer<br /><span style="font-size:9px;">fra job</span></th>
        <th style="vertical-align:bottom;">Bgt. tp.<br /><span style="font-size:9px;">fra job</span></th>
        <th style="vertical-align:bottom;">Budget<br /><span style="font-size:9px;">fra job</span></th>
        <th style="border-bottom:2px red solid; vertical-align:bottom; white-space:nowrap;">FY <%=datepart("yyyy", y0, 2,2)+addOneYear %><br /><span style="font-size:9px;">Timebgt.</span></th>
        <th style="vertical-align:bottom; vertical-align:bottom; white-space:nowrap;">FY <%=datepart("yyyy", y1, 2,2)+addOneYear %><br /><span style="font-size:9px;">Timebgt.</span></th>
        <th style="vertical-align:bottom; vertical-align:bottom; white-space:nowrap;">FY <%=datepart("yyyy", y2, 2,2)+addOneYear %><br /><span style="font-size:9px;">Timebgt.</span></th>
        <%if cint(timesimh1h2) = 1 then %>
        <th style="vertical-align:bottom;">Opr.</th>
        <%end if %>
        
        <th style="background-color:#DCF5BD; vertical-align:bottom; white-space:nowrap;">Forecast<br /><span style="font-size:9px;">Timer FY <%=datepart("yyyy", y0, 2,2)+addOneYear %></span></th>

        <%if cint(timesimtp) = 1 then%> 
        <th style="background-color:#DCF5BD; vertical-align:bottom; white-space:nowrap;">Forecast <br /><span style="font-size:9px;">Tp. FY <%=datepart("yyyy", y0, 2,2)+addOneYear %></span></th>
        <%end if %>

        <th style="background-color:#FFFFFF; vertical-align:bottom; white-space:nowrap;">Real.<br /><span style="font-size:9px;">Timer FY 1.7 - <br /><%=visrealprdato %></span></th>
        <th style="background-color:#FFFFFF; vertical-align:bottom; white-space:nowrap;">Real.<br /><span style="font-size:9px;">Gns. Tp.<br /> DKK/t.</span></th>
        <th style="background-color:#FFFFFF; vertical-align:bottom; white-space:nowrap;">Real. Oms.<br /><span style="font-size:9px;">FY 1.7 - <br /><%=visrealprdato %> DKK</span></th>
        <%if cint(timesimh1h2) = 1 then %>
        <th style="background-color:#FFFFFF; vertical-align:bottom; white-space:nowrap;">Til rådigh.<br /><span style="font-size:9px;">Timer H2</span> <span style="color:#999999; font-size:9px;">[C-I]</span></th>
        <th style="background-color:#b7e281; vertical-align:bottom; white-space:nowrap;">Forecast<br /><span style="font-size:9px;">Tim. H2 <br />FY <%=datepart("yyyy", y0, 2,2)+addOneYear %></span></th>
        <th style="background-color:#b7e281; vertical-align:bottom; white-space:nowrap;">Forecast<br /><span style="font-size:9px;">Tp. H2 <br />FY <%=datepart("yyyy", y0, 2,2)+addOneYear %></span></th>
        <th style="background-color:#99cc58; vertical-align:bottom; white-space:nowrap;">Forecast<br /><span style="font-size:9px;">H1+H2 <br />FY <%=datepart("yyyy", y0, 2,2)+addOneYear %></span></th>
        <%end if %>

        <%if cint(timesimtp) = 1 then%>  
        <th style="background-color:#DCF5BD; vertical-align:bottom; white-space:nowrap;">Forecast<br />budget <br /><span style="font-size:9px;">FY <%=datepart("yyyy", y0, 2,2)+addOneYear %> </span><span style="color:#999999; font-size:9px;">[G*H]</span></th>
        <%end if %>

        <%if cint(timesimh1h2) = 1 then %>
        <th style="background-color:#FFFFFF; vertical-align:bottom; white-space:nowrap;">Budget <br /><span style="font-size:9px;">H2 FY <%=datepart("yyyy", y0, 2,2)+addOneYear %> </span><br /><span style="color:#999999; font-size:9px;">[I3 + K*L]</span></th>
         <%end if %>
        <th style="background-color:lightpink; vertical-align:bottom; white-space:nowrap;">Forecast <br />
            <span style="font-size:9px;">Tim. Grandtotal<br /> alle FY</span>
            
            <!--<span style="font-size:9px;">FY <%=datepart("yyyy", y0, 2,2) %> </span>-->

          <!--<br /> <span style="color:darkred; font-size:9px;"> IF I = 0 : N<br />
            ELSE : O</span>-->

        </th>

         </tr>


    </thead>

<% 

'********************************************************************************* 
'*** Henter job og aktiviteter MAIN  ********************************************'
'*********************************************************************************
i = 0

    

lastknavn = ""
lastjobnavn = ""
lastFase = ""
strSQLjob = "SELECT jobnavn, jobnr, j.id AS jid, jobknr, j.budgettimer AS jobbudgettimer, jo_gnstpris, "_
&" kkundenavn, k.kid, a.id AS aid, a.navn AS aktnavn, a.budgettimer AS aktbudgettimer, aktbudget, aktbudgetsum, a.fase, jobtpris FROM job AS j "_
&" LEFT JOIN kunder AS k ON (kid = jobknr) "_
&" LEFT JOIN aktiviteter AS a ON (a.job = j.id AND (a.fakturerbar = 1 OR a.fakturerbar = 2) AND a.navn IS NOT NULL) "_
&" WHERE (j.risiko > -1 OR j.risiko = -3) AND j.jobstatus = 1 "& sogValSQLKri &""_
&" GROUP BY a.id ORDER BY kkundenavn, jobnavn, jobnr, a.fase, a.sortorder, a.navn LIMIT "& lmt

'response.write strSQLjob
'response.Flush
x = 0
oRec.open strSQLjob, oConn, 3
while not oRec.EOF
    

    


    if oRec("jobbudgettimer") <> 0 then
    jobbudgettimer = formatnumber(oRec("jobbudgettimer"), 0)
    else
    jobbudgettimer = ""
    end if

    if oRec("jo_gnstpris") <> 0 then
    jo_gnstpris = formatnumber(oRec("jo_gnstpris"), 0)
    else
    jo_gnstpris = ""
    end if

    if oRec("jobtpris") <> 0 then
    jobbudget = formatnumber(oRec("jobtpris"), 0)
    else
    jobbudget = ""
    end if

    if oRec("aktbudgettimer") <> 0 then
    aktbudgettimer = formatnumber(oRec("aktbudgettimer"), 0)
    else
    aktbudgettimer = ""
    end if

    if oRec("aktbudget") <> 0 then
    aktpris = formatnumber(oRec("aktbudget"), 0)
    else
    aktpris = ""
    end if

    if oRec("aktbudgetsum") <> 0 then
    aktbudgetsum = formatnumber(oRec("aktbudgetsum"), 0)
    else
    aktbudgetsum = ""
    end if


    if lastknavn <> lcase(oRec("kkundenavn")) then%>
    <tr style="background-color:#EFF3FF;"><td colspan="2" style="padding:10px 1px 2px 2px;"><b><%=oRec("kkundenavn") %></b></td>
        <td colspan="2"><!--<input type="submit" value="Opdater >>" />-->&nbsp;</td>
        <td colspan="70">&nbsp;</td></tr>
    <%end if %>
    
            <%if lastjobnavn <> lcase(oRec("jobnavn")) then %>
            <tr>
           
            <input type="hidden" name="FM_jobid" value="<%=oRec("jid") %>" />
            <input type="hidden" id="FM_aktid_<%=i %>" value="0" />
            <input type="hidden" id="FM_jobid_<%=i %>" value="<%=oRec("jid") %>" />
           
            <td style="white-space:nowrap;"><span style="color:#5582d2;" id="sp_<%=oRec("jid") %>" class="sp_jid"><b>[+]</b></span>&nbsp;<%=left(oRec("jobnavn"), 20) %></td><td><%=oRec("jobnr") %></td>
       
             <td><a class="btn btn-info btn-xs" href="timbudgetsim.asp?jobid=<%=oRec("jid")%>&func=forecast&FM_visrealprdato=<%=visrealprdato%>" role="button" target="_blank">Forecast</a></td>
             <td><input type="text" style="width:60px;" id="budgettimer_jobakt_<%=oRec("jid")%>_0" name="FM_jobbudgettimer" class="jobakt_budgettimer_job form-control input-small" value="<%=jobbudgettimer%>" /></td>
            <td><input type="text" style="width:60px;" id="budgettimep_jobakt_<%=oRec("jid")%>_0" name="FM_jobtpris" class="jobakt_budgettimep_job form-control input-small" value="<%=jo_gnstpris%>" /></td>
            <td><input type="text" style="width:80px;" id="budgetbelob_jobakt_<%=oRec("jid")%>_0" name="FM_jobbudget" class="jobakt_budgettimep_job form-control input-small" value="<%=jobbudget%>" disabled /></td>

                <input type="hidden" name="FM_jobbudgettimer" value="##" />
                <input type="hidden" name="FM_jobtpris" value="##" />
                <input type="hidden" name="FM_jobbudget" value="##" />

                <%call jobaktbudgetfelter(oRec("jobnr"), oRec("jid"), 0, h1aar, h2aar, h1md, h2md, jo_gnstpris) %>

                 <%'call medarbfelter(oRec("jid") , 0, h1aar, h2aar, h1md, h2md) %>
             </tr>
           <% 
             
                 
             i = i + 1
             x = x + 1
           end if 
               
               
     select case right(x, 1)
    case 0,2,4,6,8
    bgthis = "#FFFFFF"
    case else
    bgthis = "#F3F3F3"
    end select%>

    <%if lastFase <> lcase(oRec("fase")) then %>
    <tr class="tr_<%=oRec("jid") %>" style="background-color:<%=bgthis%>; visibility:hidden; display:none;">
        <td colspan="100"><%=replace(oRec("fase"), "_", " ") %></td>
    </tr>
    <%
        'strAktTxtTds(oRec("aid")) = strAktTxtTds(oRec("aid")) & "<tr><td colspan=""100"">"& replace(oRec("fase"), "_", " ")  &"</td></tr>"
        
        end if %>
        
    <tr class="tr_<%=oRec("jid") %>" style="background-color:<%=bgthis%>; visibility:hidden; display:none;">
        <input type="hidden"  name="FM_aktjobid" id="FM_jobid_<%=i %>" value="<%=oRec("jid") %>" />
        <input type="hidden" name="FM_aktid" value="<%=oRec("aid") %>" />
        <input type="hidden" id="FM_aktid_<%=i %>" value="<%=oRec("aid") %>" />
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td style="white-space:nowrap;"><%=left(oRec("aktnavn"), 20) %></td>

        <td><input type="text" name="FM_aktbudgettimer" id="budgettimer_jobakt_<%=oRec("jid")%>_<%=oRec("aid") %>" class="form-control input-small jobakt_budgettimer_job jobakt_budgettimer_<%=oRec("jid")%>" style="width:60px;" value="<%=aktbudgettimer %>" /></td>
         <td><input type="text" style="width:60px;" name="FM_aktpris" id="budgettimep_jobakt_<%=oRec("jid")%>_<%=oRec("aid") %>" class="form-control input-small jobakt_budgettimep_job jobakt_budgettimep_<%=oRec("jid")%>" value="<%=aktpris %>" /></td>
          <td><input type="text" style="width:80px;" name="FM_aktbudgetsum" id="budgettimep_jobakt_<%=oRec("jid")%>_<%=oRec("aid") %>" class="form-control input-small jobakt_budgettimep_job jobakt_budgettimep_<%=oRec("jid")%>" value="<%=aktbudgetsum %>" disabled /></td>


                <input type="hidden" name="FM_aktbudgettimer" value="##" />
                <input type="hidden" name="FM_aktpris" value="##" />
                <input type="hidden" name="FM_aktbudgetsum" value="##" />
                
              
        

           <%call jobaktbudgetfelter(oRec("jobnr"), oRec("jid"), oRec("aid"), h1aar, h2aar, h1md, h2md, aktpris) %>


            <%'call medarbfelter(oRec("jid"), oRec("aid"), h1aar, h2aar, h1md, h2md) %>

      
     </tr>
    <%

        'strAktTxtTds(oRec("aid")) = strAktTxtTds(oRec("aid")) &  "<td>&nbsp;</td>"
        'strAktTxtTds(oRec("aid")) = strAktTxtTds(oRec("aid")) &  "<td>&nbsp;</td>"  


        if len(trim(oRec("aktnavn"))) <> 0 then
        strAktTxtTds(oRec("aid")) = strAktTxtTds(oRec("aid")) & "<td style='white-space:nowrap;'>"& left(oRec("aktnavn"), 20) &"</td>"
        end if
       

        if isNull(oRec("fase")) <> true then
        lastFase = lcase(oRec("fase"))
        else
        lastFase = ""
        end if


    i = i + 1
    lastknavn = lcase(oRec("kkundenavn"))
    lastjobnavn = lcase(oRec("jobnavn")) 
    'lastJid = oRec("jid")
    x = x + 1
oRec.movenext    
wend 
oRec.close 


       

%>

    <input type="hidden" value="<%=x-1 %>" id="xhigh" />
    

    <%

    '******************** Medarbejdere LOOP ****************************

        budgetFY0GT = formatnumber(budgetFY0GT, 0)

%>

<tr><td colspan=4>Total:</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
     <%if cint(timesimh1h2) = 1 then %>
    <td>&nbsp;</td>
    <%end if %>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <%if cint(timesimtp) = 1 then%> 
    <td>&nbsp;</td>
    <%end if %>

     <%if cint(timesimh1h2) = 1 then %>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <%end if %>

    <%if cint(timesimtp) = 1 then%> 
    <td>&nbsp;</td>
    <%end if %>

    <%if cint(timesimh1h2) = 1 then %>
    <td>&nbsp;</td>
    <%end if %>

    <td><input type="text" class="form-control input-small" style="width:80px; text-align:right;" id="budgetgt" value="<%=budgetFY0GT %>" disabled /></td></tr>

</table>
   
                              
  <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
            <br />&nbsp;
                             
</div>



      
</form>     

    </div><!-- portlet -->
    </div><!-- container -->
    



     </div><!-- content -->
    </div><!-- wrapper -->

<%end select 
    
end if 'session%>

<!--#include file="../inc/regular/footer_inc.asp"-->


	
