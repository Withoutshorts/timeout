  

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
   
	<!--#include file="../inc/errors/error_inc.asp"-->
	<!--#include file="../inc/regular/global_func.asp"-->
	<!--#include file="inc/convertDate.asp"-->
	<!--#include file="inc/timereg_akt_2006_inc.asp"-->
    <!--#include file="inc/timereg_hojrediv_inc.asp"-->
	<!--#include file="inc/timereg_matrix_timespan_inc.asp"-->
    <!--#include file="inc/smiley_inc.asp"-->
	<!--#include file="inc/isint_func.asp"-->
    <!--#include file="../inc/regular/treg_func.asp"-->
    <!--#include file="../inc/regular/topmenu_inc.asp"-->
 
	
	


    
	
	<%

  
	
	
	   '**** Søgekriterier AJAX **'
        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")

        case "FN_tjktimer_traveldietexp"
                '** Traveldietexp

                 

                usemrn = request("treg_usemrn")
                timerTastet = request("timer_tastet")
                tdatoSQL = request("tdato")
                tdatoSQL = year(tdatoSQL) & "-"& month(tdatoSQL) & "-" & day(tdatoSQL)
                strAfgSQLtimer = " AND (tdato = '"& tdatoSQL &"')"
                
               
                call traveldietexp_fn()

                

                if cint(traveldietexp_on) = 1 then' tjk maks antal timer pr. dag ved rejse

                        '***TJK Om det er en rejse dag **************************************
                        errejsedag = 0
                        strSQLtravl = "SELECT diet_mid FROM Traveldietexp WHERE diet_mid = " & usemrn & " AND (diet_stdato <= '"& tdatoSQL &" 23:59' AND diet_sldato >= '"& tdatoSQL &" 00:01') ORDER BY diet_stdato"
                        
                        
                        oRec5.open strSQLtravl, oConn, 3
                        while not oRec5.EOF
               
                        errejsedag = 1
                       
                        oRec5.movenext
                        wend
                        oRec5.close

                     


                if cint(errejsedag) = 1 then

                        
        
                        '*******************
                        call akttyper2009(2)

                       
                        'Timer brugt på dagen
                        timerBrugt = 0
                        strSQLtimreal = "SELECT COALESCE(SUM(timer), 0) AS timerbrugt FROM timer WHERE tmnr = "& usemrn &" "& strAfgSQLtimer &" AND ("& aty_sql_realhours &") GROUP BY tdato, tmnr"
                        t = 0
                      
                        oRec5.open strSQLtimreal, oConn, 3
                        while not oRec5.EOF
               
                        timerBrugt = oRec5("timerbrugt")
                 
                        t = 1
                        oRec5.movenext
                        wend
                        oRec5.close
            
                    
                       

                       if len(trim(traveldietexp_maxhours)) <> 0 AND traveldietexp_maxhours > -1 then
                       traveldietexp_maxhours = traveldietexp_maxhours * 1
                       else
                       traveldietexp_maxhours = 0
                       end if

               
                       if len(trim(timerBrugt)) <> 0 AND timerBrugt <> 0 then
                       timerBrugt = replace(timerBrugt, ".", ",") * 1
                       else
                       timerBrugt = 0
                       end if
               
                       if len(trim(timerTastet)) <> 0 AND timerTastet <> 0 then
                       timerTastet = replace(timerTastet, ".", ",") * 1
                       else
                       timerTastet = 0
                       end if

              
                       if timerTastet*1 > 0 then
                       feltTxtVal = (traveldietexp_maxhours*1 - (timerTastet*1 + timerBrugt*1))
                       end if
        
               


                response.write feltTxtVal

                end if 'er rejsedag

                end if 'traveldietexp_on

                response.end

        case "FN_tjktimer_forecast"
                '** FORECAST ALERT 
         

                aktid = request("aktid")
                timerTastet = request("timer_tastet")
                usemrn = request("treg_usemrn")
                ibudgetaar = request("ibudgetaar")
                ibudgetmd = request("ibudgetaarMd")  
                aar = request("ibudgetUseAar")
                md = request("ibudgetUseMd")

                '** Afgræns indenfor budgetår
                if cint(ibudgetaar) = 1 then

                    if ibudgetmd <> 1 then '7: 1 juli - 30 juni

                        'Beregn evt. days inmonth
                       
                        if md >= 7 then 'jul-dec sammeår
                        strAfgSQL = "AND ((aar = "& aar &" AND md >= "& ibudgetmd &") OR (aar = "& aar+1 &" AND md < "& ibudgetmd &"))"
                        strAfgSQLtimer = "AND (tdato BETWEEN '"& aar &"-"& ibudgetmd &"-01' AND '"& aar+1 &"-"& ibudgetmd-1 &"-30')"
                        else
                        strAfgSQL = "AND ((aar = "& aar-1 &" AND md >= "& ibudgetmd &") OR (aar = "& aar &" AND md < "& ibudgetmd &"))"
                        strAfgSQLtimer = "AND (tdato BETWEEN '"& aar-1 &"-"& ibudgetmd &"-01' AND '"& aar &"-"& ibudgetmd-1 &"-30')"
                        end if
                    
                      else

                        strAfgSQL = "AND (aar = "& aar &")"
                        strAfgSQLtimer = "AND (tdato BETWEEN '"& aar &"-01-01' AND '"& aar &"-12-31')"
                
                    end if

                else

                    strAfgSQL = ""
                    strAfgSQLtimer = ""
        
                end if 

                feltTxtVal = 0
                
                '** Tjekker resouceforecast
                strSQLtimfc = "SELECT COALESCE(SUM(timer), 0) AS fctimer FROM ressourcer_md WHERE aktid = "& aktid & " AND medid = "& usemrn &" "& strAfgSQL
                

                fctimer = 0
                oRec5.open strSQLtimfc, oConn, 3
                while not oRec5.EOF
               
                fctimer = oRec5("fctimer")
                 
                oRec5.movenext
                wend
                oRec5.close

                
                'SKAL DER KUNNE TASTES HVIS DER IKKE ER ANGIVET FORECAST = NEJ HÅRD STYRRING?
                timerBrugt = 0
               
        
                
                        '** Finder timeforbrug på AKTIVITET total på tværs
                        '** HUSK BUDGET ÅR
                        strSQLtimbudget = "SELECT COALESCE(SUM(timer), 0) AS timerbrugt FROM timer WHERE taktivitetid = "& aktid &" AND tmnr = "& usemrn &" "& strAfgSQLtimer &" GROUP BY taktivitetid"
                        t = 0
                        
                        oRec5.open strSQLtimbudget, oConn, 3
                        while not oRec5.EOF
               
                        timerBrugt = oRec5("timerbrugt")
                 
                        t = 1
                        oRec5.movenext
                        wend
                        oRec5.close

        
               

               if len(trim(fctimer)) <> 0 AND fctimer <> 0 then
               fctimer = replace(fctimer, ".", ",") * 1
               else
               fctimer = 0
               end if

               
               if len(trim(timerBrugt)) <> 0 AND timerBrugt <> 0 then
               timerBrugt = replace(timerBrugt, ".", ",") * 1
               else
               timerBrugt = 0
               end if
               
               if len(trim(timerTastet)) <> 0 AND timerTastet <> 0 then
               timerTastet = replace(timerTastet, ".", ",") * 1
               else
               timerTastet = 0
               end if

              
               if timerTastet*1 > 0 then
               feltTxtVal = (fctimer*1 - (timerTastet*1 + timerBrugt*1))
               end if
        
               


                response.write feltTxtVal
                response.end




        case "FN_tjktimer_budget"
                '*** FORKALK ALERT


                aktid = request("aktid")
                timerTastet = request("timer_tastet")
                
                'feltTxt = timerTastet
                feltTxtVal = 0
                
            

                '*** SKAL DET KUN VÆRE TEC / ESN --> Kontrolpanel
             
                '** Finder budget på AKTIVITET Kun hvis grundlag er timer
                strSQLtimbudget = "SELECT budgettimer FROM aktiviteter WHERE id = "& aktid & " AND bgr = 1"
                a = 0
                budgettimer = 0
                oRec5.open strSQLtimbudget, oConn, 3
                while not oRec5.EOF
               
                budgettimer = oRec5("budgettimer")
                 
                a = 1
                oRec5.movenext
                wend
                oRec5.close

                'KUN HVIS DER ER ANGIVET ET budget på aktiviteten tjekkes det er om det er overskreddet
                timerBrugt = 0
                if a = 1 AND budgettimer > 0 then
        
                
                         '** Finder timerforbrug på AKTIVITET total på tværs
                        strSQLtimbudget = "SELECT COALESCE(SUM(timer), 0) AS timerbrugt FROM timer WHERE taktivitetid = "& aktid &" GROUP BY taktivitetid"
                        t = 0
                        
                        oRec5.open strSQLtimbudget, oConn, 3
                        while not oRec5.EOF
               
                        timerBrugt = oRec5("timerbrugt")
                 
                        t = 1
                        oRec5.movenext
                        wend
                        oRec5.close

        
               

                       if len(trim(budgettimer)) <> 0 AND budgettimer <> 0 then
                       budgettimer = replace(budgettimer, ".", ",") * 1
                       else
                       budgettimer = 0
                       end if

               
                       if len(trim(timerBrugt)) <> 0 AND timerBrugt <> 0 then
                       timerBrugt = replace(timerBrugt, ".", ",") * 1
                       else
                       timerBrugt = 0
                       end if
               
                       if len(trim(timerTastet)) <> 0 AND timerTastet <> 0 then
                       timerTastet = replace(timerTastet, ".", ",") * 1
                       else
                       timerTastet = 0
                       end if

              

                 'feltTxtVal = 100 'timerBrugt
                 'response.write "timerTastet: " & timerTastet*1 &" timerBrugt: " & timerBrugt*1 &" = "& timerTastet*1 + timerBrugt*1 &" :: budgettimer: "& budgettimer
               
                 'response.end    


                 '** Er ferie opbrugt ***'
                '** LAVES IKKE ENDNU. MANGE LOKALER REGLER FOR FERIE.
                '** MÅ DE TRÆKKE OVER I FERIEDAGE, SKAL DE INDLÆSES PÅ FERIE AFHOLDT UDEN LØN? 
                '** SKAL ADMINISTRATIONEN / LØNSYSTEMET SELV BESLUTTE om ferie skal indlæses som afholdt???
                'strSQLferie = "SELECT"
                
                if (timerTastet*1 + timerBrugt*1) > budgettimer*1 AND budgettimer*1 > 0 AND a = 1 then
                feltTxtVal = (budgettimer*1 - (timerTastet*1 + timerBrugt*1))
                end if
        
                       

                end if 


                response.write feltTxtVal
                response.end



        case "FN_sog_lager"

        strGrp = request("matreg_grp")
        strLev = request("matreg_lev")
        strSog = request("matreg_sog")

       
       %>
                <form>
                    <table cellpadding="0" cellspacing="0" border="0" width="100%">
                            <%call matregLagerheader(0) %>
       

                        <% 
                
                                useSog = 1

                                 if len(trim(strSog)) <> 0 then
                                    sogeKri = strSog
                                else
                                    sogeKri = ""
                                end if    
                    
                   
                                if len(trim(strLev)) <> 0 AND strLev <> "null" then
                                    sogLevSQLkri = " AND m.leva = "& strLev
                                else
                                    sogLevSQLkri = " AND m.leva <> -1 "
                                end if    

                    
                                if len(trim(strGrp)) <> 0 AND strGrp <> "null" then
                                    sogMatgrpSQLkri = " m.matgrp = "& strGrp
                                else
                                    sogMatgrpSQLkri = " m.matgrp <> -1 "
                                end if    
                
                    
                                vis = 0
                
                            call matfelter_lager(vis) %>

                        </table>
                    </form>

        

                    <%
        case "FN_indlaest_allerede_mat"

                aktid = request("matreg_aktid")
                medid = request("matreg_medid")
                regdato = request("matreg_regdato") 
                regdatoTXT = replace(formatdatetime(regdato, 2), "/", ".")
                regdato = year(regdato)&"/"&month(regdato)&"/"&day(regdato)
             


                strSQLmatreg = "SELECT m.matnavn, m.matantal, m.matsalgspris, m.valuta, matenhed, v.valutakode FROM materiale_forbrug AS m "_
                &"LEFT JOIN valutaer AS v ON (v.id = m.valuta) WHERE m.aktid = "& aktid &" AND usrid = "& medid & " AND forbrugsdato = '"& regdato &"'"
                
                'if lto = "dencker" then
                'response.write strSQLmatreg
                'response.end
                'end if
                m = 0
                oRec5.open strSQLmatreg, oConn, 3
                while not oRec5.EOF
               
                    if m = 0 then
                    strIndLaestAllerede = "<b>Dine indlæste materialer / udlæg d. "& regdatoTXT &"</b><br>"
                    end if        

                    strIndLaestAllerede = strIndLaestAllerede & oRec5("matantal") &" "& oRec5("matenhed") &". "& oRec5("matnavn") &" "& oRec5("matsalgspris") &" "& oRec5("valutakode") &"<br />"

                m = m + 1
                oRec5.movenext
                wend
                oRec5.close


                '*** ÆØÅ **'
                call jq_format(strIndLaestAllerede)
                strIndLaestAllerede = jq_formatTxt
                response.write strIndLaestAllerede & "<br>&nbsp;"
                response.end
                

        case "FN_indlaes_mat"

       

        '**** Værdier ****'

        strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
        jobid = request("matreg_jobid")
        medid = request("matreg_medid")


        '** Expences / Materialeforbrug
        '** OnTheFly: 1 / Fra lager : 0 
        if len(request("matreg_onthefly")) <> 0 AND request("matreg_onthefly") <> 0 then
		otf = 1 'request("matreg_onthefly")       
		else
		otf = 0
		end if      

        matId = request("matreg_matid")
        aktid = request("matreg_aktid")
        aftid = request("matreg_aftid")

        if len(trim(request("matreg_antal"))) <> 0 then
		intAntal = replace(request("matreg_antal"), ",", ".")
		else
		intAntal = 0
		end if

        regdato = request("matreg_regdato")
        valuta = request("matreg_valuta")
        intkode = request("matreg_intkode")
                
        if len(trim(request("matreg_personlig"))) <> 0 then
        personlig = 1
        else
        personlig = 0
        end if
                
                
        if len(request("matreg_bilagsnr")) <> 0 then
        bilagsnr = request("matreg_bilagsnr") 
        else
        bilagsnr = ""
        end if

        pris = request("matreg_pris")
        salgspris = request("matreg_salgspris")
        gruppe = request("matreg_gruppe")

        navn = replace(request("matreg_navn"), "'", "")
       
        if len(trim(request("matreg_varenr"))) <> 0 then
        varenr = request("matreg_varenr")
        else
        varenr = 0
        end if
        
        matreg_opdaterpris = request("matreg_opdaterpris")
        opretlager = request("matreg_opretlager")
        betegnelse = request("matreg_betegn")

        mat_func = request("matreg_func")

        'response.write "onthefly: "& otf &" aktid:"& aktid &" Dato: "& regdato &" Navn: "& navn & " varenr: "& varenr &" opretlager: "& opretlager &" valuta: "& valuta &"<br>"
        'response.end
        matregid = 0
        matava = 0

        call indlaes_mat(matregid, otf, medid, jobid, aktid, aftid, matId, strEditor, strDato, intAntal, regdato, valuta, intkode, personlig, bilagsnr, pris, salgspris, navn, gruppe, varenr, opretlager, betegnelse, mat_func, matreg_opdaterpris, matava)
	

        response.end

        case "FN_showfilter"
                
                if len(trim(request("jq_newfilterval"))) <> 0 then
                showfilter = request("jq_newfilterval")
                else
                showfilter = 0
                end if

                response.cookies("tsa_2012")("showfilter") = showfilter '4 'showfilter
                
               

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
            	

                       
            			
                        '***
                        medarbejderid = request("vlgtmedarb")
                        jstatus = request("jstatus")
                        lukjob = request.form("jq_lukjob") 'split(request("FM_lukjob"), ",")


                        'response.write "jstatus: "& jstatus
                        'response.end

                        call lukjobmail(jstatus, lukjob, medarbejderid)
                       
            	
	              
                       


	            end if '** Luk job **'

              
                Response.end

        case "FN_fjernjob"
                  
                 
                 jobid = request("jobid")   
                 medid = request("usemrn")

                 '** Nulstillerforvalgt job for denne medarbejder ****'
                 strSQLUpdFvOff = "UPDATE timereg_usejob SET forvalgt = 0 WHERE medarb = "& medid &" AND jobid = "& jobid 
                 oConn.Execute(strSQLUpdFvOff)
                 '******************************************************'
         
         
         case "FN_tilfojakt"
                  
                  '*** bruges på 14 / 2 mds listen ***'
                 
                 'jobid = request("jobid")   
                 medid = request("usemrn")
                 aktids = split(request("aktids"), ",")
                 dtNow = year(now) & "/" & month(now) & "/"& day(now)

                 'Response.write "aktids:"&  request("aktids")
                 'Response.end

                 for a = 0 to UBOUND(aktids)
                     
                     thisJob = 0
                     strSQLjob = "SELECT job FROM aktiviteter WHERE id = "& trim(aktids(a)) 
                     oRec5.open strSQLjob, oConn, 3
                     if not oRec5.EOF then
                     thisJob = oRec5("job")
                     end if
                     oRec5.close

                     strSQLsel = "SELECT id FROM timereg_usejob WHERE aktid = "& trim(aktids(a)) &" AND medarb = "& medid 

                

                             oRec5.open strSQLsel, oConn, 3
                             if not oRec5.EOF then 

                             strSQLInsAkt = "UPDATE timereg_usejob SET forvalgt = 1, forvalgt_dt = '"& dtNow &"' WHERE id = "& oRec5("id")
                             'Response.write "A"& strSQLInsdAkt & "<br>"
                             'Response.flush
                             oConn.Execute(strSQLInsAkt)

                             else

                              strSQLInsAkt = "INSERT INTO timereg_usejob SET jobid = "& thisJob &", aktid = "& trim(aktids(a)) &", forvalgt = 1, forvalgt_dt = '"& dtNow &"', medarb = "& medid 
                              'Response.write "B" &strSQLInsdAkt & "<br>"
                              'Response.flush
                              oConn.Execute(strSQLInsAkt)

                             end if
                             oRec5.close

                    

                 next
        
         
          case "FN_lasakt" 
         
                 aktid = request("aktid") 
                 medid = request("usemrn")

                     thisJob = 0
                     strSQLjob = "SELECT job FROM aktiviteter WHERE id = "& aktid 
                     oRec5.open strSQLjob, oConn, 3
                     if not oRec5.EOF then
                     thisJob = oRec5("job")
                     end if
                     oRec5.close

                 '*** Hvis den findes sættes den til -1 eller oprettes den **'

                 strSQLsel = "SELECT id FROM timereg_usejob WHERE aktid = "& aktid &" AND medarb = "& medid 

                 oRec5.open strSQLsel, oConn, 3
                 if not oRec5.EOF then 

                 strSQLLasAkt = "UPDATE timereg_usejob SET forvalgt = 2 WHERE id = "& oRec5("id")
                 'Response.write "A"& strSQLInsdAkt & "<br>"
                 'Response.flush
                 oConn.Execute(strSQLLasAkt)

                 else

                  strSQLLasAkt = "INSERT INTO timereg_usejob SET jobid = "& thisJob &", aktid = "& aktid &", forvalgt = 2, medarb = "& medid 
                  'Response.write "B" &strSQLInsdAkt & "<br>"
                  'Response.flush
                  oConn.Execute(strSQLLasAkt)

                 end if
                 oRec5.close
                 

          case "FN_sliakt" 
         
                 aktid = request("aktid") 
                 medid = request("usemrn")


                  thisJob = 0
                     strSQLjob = "SELECT job FROM aktiviteter WHERE id = "& aktid 
                     oRec5.open strSQLjob, oConn, 3
                     if not oRec5.EOF then
                     thisJob = oRec5("job")
                     end if
                     oRec5.close

                 '*** Hvis den findes sættes den til -1 eller oprettes den **'

                 strSQLsel = "SELECT id FROM timereg_usejob WHERE aktid = "& aktid &" AND medarb = "& medid 

                

                 oRec5.open strSQLsel, oConn, 3
                 if not oRec5.EOF then 

                 strSQLLasAkt = "UPDATE timereg_usejob SET forvalgt = 1 WHERE id = "& oRec5("id")
                 'Response.write "A"& strSQLInsdAkt & "<br>"
                 'Response.flush
                 oConn.Execute(strSQLLasAkt)

                 else

                  strSQLLasAkt = "INSERT INTO timereg_usejob SET jobid = "& thisJob &", aktid = "& aktid &", forvalgt = 1, medarb = "& medid 
                  'Response.write "B" &strSQLInsdAkt & "<br>"
                  'Response.flush
                  oConn.Execute(strSQLLasAkt)

                 end if
                 oRec5.close
                 

                 
         case "FN_fjernakt"
         
                 aktid = request("aktid") 
                 medid = request("usemrn")


                  thisJob = 0
                     strSQLjob = "SELECT job FROM aktiviteter WHERE id = "& aktid 
                     oRec5.open strSQLjob, oConn, 3
                     if not oRec5.EOF then
                     thisJob = oRec5("job")
                     end if
                     oRec5.close

                 '*** Hvis den findes sættes den til -1 eller oprettes den **'

                 strSQLsel = "SELECT id FROM timereg_usejob WHERE aktid = "& aktid &" AND medarb = "& medid 

                

                 oRec5.open strSQLsel, oConn, 3
                 if not oRec5.EOF then 

                 strSQLDelAkt = "UPDATE timereg_usejob SET forvalgt = -1 WHERE id = "& oRec5("id")
                 'Response.write "A"& strSQLInsdAkt & "<br>"
                 'Response.flush
                 oConn.Execute(strSQLDelAkt)

                 else

                  strSQLInsAkt = "INSERT INTO timereg_usejob SET jobid = "& thisJob &", aktid = "& aktid &", forvalgt = -1, medarb = "& medid 
                  'Response.write "B" &strSQLInsdAkt & "<br>"
                  'Response.flush
                  oConn.Execute(strSQLInsAkt)

                 end if
                 oRec5.close
                
                 
         
         case "FN_fjernAllejob"
                  
                 
                 medid = request("usemrn")

                 '** Nulstillerforvalgt job for denne medarbejder ****'
                 strSQLUpdFvOff = "UPDATE timereg_usejob SET forvalgt = 0 WHERE medarb = "& medid 
                 oConn.Execute(strSQLUpdFvOff)
                 '******************************************************'


        case "FM_sortOrder"
               'Call AjaxUpdateTreg("timereg_usejob","forvalgt_sortorder","")
                Call AjaxUpdate("job","risiko","")
        
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
                 
        




        '*********************************************************************************************************************************************         
        '*** Henter aktiviteter ***'
        '*********************************************************************************************************************************************         
        case "FN_showakt"        
            


            '*** Variable ****
            call smileyAfslutSettings()

            


            felt = ""
            
            

            level = session("rettigheder")

            jobid = request("jobid")
            intEasyreg = request("FM_easyreg")

            sortByVal = request("sortByVal")



            if len(trim(request("visallemedarb_bl"))) <> 0 then
            visallemedarb_bl = request("visallemedarb_bl")
            
            dim aktidsjobProdgrp, aktidpjob        ', aprg2, aprg3, aprg4, aprg5, aprg6, aprg7, aprg8, aprg9, aprg10, aktidpjob 
            redim aktidsjobProdgrp(50), aktidpjob(50)  'aprg1(50), aprg2(50), aprg3(50), aprg4(50), aprg5(50), aprg6(50), aprg7(50), aprg8(50), aprg9(50), aprg10(50), aktidpjob(50)

            else
            visallemedarb_bl = 0
            end if
            
            if len(trim(request("jobnr_bl"))) <> 0 then
            jobnr_bl = request("jobnr_bl")
            else
            jobnr_bl = 0
            end if

            'Response.write "jobnr_bl : "& jobnr_bl 
            'Response.write "<br>visallemedarb_bl : "& visallemedarb_bl 
            'Response.end
            'redim aktdata_jq(100, 30)
            
            
            '**** aktiviteter på valgte job **********'
            '**** bruges ikke på 14 / 2 md listen ****'
            if instr(request("job_aktids"),",") <> 0 then
            aktidsSQL = " WHERE (a.id = "&  replace(request("job_aktids"),",", "  OR a.id = ")
            else
            aktidsSQL = " WHERE (a.id = 0)"
            end if
           
          
            jobstatus = request("FM_jobstatus")
            ignorerakttype = request("ignakttype")


            call positiv_aktivering_akt_fn()

            'Response.write positiv_aktivering_akt_val & "<br>"
            'Response.flush
            
            if cint(positiv_aktivering_akt_val) <> 1 then

                if ignorerakttype <> "1" then ' ignorer typer på tilbud (vis alle akt.typer)

                if jobstatus = 3 then 'tlbud
                'strSQLAktStatusKri = "AND a.fakturerbar = 6 AND a.aktstatus = 1" '** Kun salgs aktivitets typer
                strSQLAktStatusKri = " AND ((j.jobstatus = 3 AND a.fakturerbar = 6) AND a.aktstatus = 1)"
                else
                strSQLAktStatusKri = " AND ((j.jobstatus = 1 AND a.aktstatus = 1) OR (j.jobstatus = 3 AND a.fakturerbar = 6 AND a.aktstatus = 1))" '** alle aktivitets typer
                end if

                else

                strSQLAktStatusKri = " AND ((j.jobstatus = 1 OR j.jobstatus = 3) AND a.aktstatus = 1)" '** alle aktivitets typer
           

                end if

            else

            strSQLAktStatusKri = " AND ((j.jobstatus = 1 OR j.jobstatus = 3) AND a.aktstatus = 1)"

            end if
            
            nyakt = 0
          

            job_iRowLoops = split(request("job_iRowLoops"),",")
            job_iRowLoops_cn = job_iRowLoops(1) 

            aktidsSQLOptions = ""
          
          
            foundone = "n"
            usemrn = request("usemrn")
            
           

            dim aktidsPrMid
            redim aktidsPrMid(200)

            stDato = request("stDato")
            
            '*** 5 ell 6 *** 14 dage / 2 md. 
            if cint(sortByVal) = 6 then
            stDato_30 = dateadd("d", -60, stDato)
            else
            stDato_30 = dateadd("d", -14, stDato)
            end if
            
            slDato = request("slDato")

            'dim tjekDato
            redim tjekdag(7)
            tjekdag(1) = stDato
            
            for x = 2 to 7
            tjekdag(x) = dateAdd("d", x-1, stDato)
            next

            '*** Skal akt lukkes for timereg. hvis forecast budget er overskrddet..?
            call aktBudgettjkOn_fn()
            

            '** Norm tid for ugen'**
            call normtimerPer(usemrn, tjekdag(1), 6, 0) 
            '** gns timer pr. iforhold til samlet arbejdsuge, baseret på en 5 dages arbejdsuge. ***'
            '** Således at det er underordnet om man indtaster på en tirsdag (7,5) eller en fredag (7,0) '****'
            normTimerGns5 = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)  / 5
            normTimerprUge = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)                         


            'Response.write "<hr>"& normTimerGns5 &".."& tjekdag(1)
          
            normtimer_man = ntimMan
            normtimer_tir = ntimTir
            normtimer_ons = ntimOns
            normtimer_tor = ntimTor
            normtimer_fre = ntimFre
            normtimer_lor = ntimLor
            normtimer_son = ntimSon

            call akttyper2009(2)
            
            '** Medarb. i projektgrupper
            call medariprogrpFn(usemrn)

            call ersmileyaktiv()
            
            
            '*** ÆØÅ **'

            acn = 0
            acnDtlinje = 0
            acnForFaseIsWrt = 0
         

            visHRliste = request("FM_hr")

            lastFakDato = request("lastFakDato")
            
            ignorertidslas = request("ingTlaas")
           
            visTimerElTid = request("FM_vistimereltid")

            ignJobogAktper = request("FM_ignJobogAktper")

            visSimpelAktLinje = request("FM_visSimpelAktLinje")

            call visAktSimpel_fn()

            'Response.Write "ignJobogAktper" & ignJobogAktper
            'Response.end

            useDateStSQL = year(stDato) &"/"& month(stDato) &"/"& day(stDato)
            useDateSlSQL = year(slDato) &"/"& month(slDato) &"/"& day(slDato)

            useDateSt_30SQL = year(stDato_30) &"/"& month(stDato_30) &"/"& day(stDato_30)

           


            if ignJobogAktper = "1" then
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
         
                                    

                                  
             
                                    '*** aktiviteter **'
                                    aktnavn = ".."
                                    job_fase = ""
                                    last_job_jid = 0

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
                                    lastKnavn = "A"
                                    
                                    '** Tjekker om det er et internt job -2 **'
                                    risiko = 0
                                    strSQLrisiko = "SELECT risiko FROM job WHERE id = "& jobid
                                    
                                    oRec6.open strSQLrisiko, oConn, 3 
                                    if not oRec6.EOF then
                                    risiko = oRec6("risiko")
                                    end if
                                    oRec6.close


                                    

                                    if len(trim(request("viskunForecast"))) <> 0 AND request("viskunForecast") <> 0 then
                                    viskunForecast = 1
                                    else
                                    viskunForecast = 0
                                    end if

                                    
                                    
                                    
                                    '*** Blandetliste
                                    '*************************************************************************************************************************'
                                    '**** Henter kun akt med timer på + netop tilføjede (aktid <>0) ved blandetliste 14 dage / 2 md / Via lle medarb på job **'
                                    '*************************************************************************************************************************'
                                    if (cint(sortByVal) = 5 OR cint(sortByVal) = 6) AND cint(intEasyreg) <> 1 AND cint(visallemedarb_bl) <> 1 then
                                    
                                   

                                   
                                            aktidsMtimWrt = ""
                                            aktidsLasWrt = ""
                                            aktidsSkjWrt = ""

                                            aktidsMtimSQL = "WHERE (a.id = 0 " 
                                            aktidsSQLOptions = "WHERE (a.id = 0 "

                                     
                                     
                                  
                                            '-1 = Skjult (altid med options listen i bunden af timereg.siden) 
                                            '0 = findes ikke, da de kunde findes i timereg-usejob hvis de er tiføjet
                                            '1 = Vis, vises kun med timer inde for de seneste 14 dg / 2 md, eller hvis de er tilføjt indefor de senste 14 dage
                                            '2 = Vises altid selvom der ikke er timer indenfor seneste 14 ddg / 2 md, og de er tilføjet for mere ned 14 / 2 md siden 

                                    

                                            '*** Henter og netop tilføjede og låste (faste aktiviteter) ***'
                                            '*** Kun specifkke aktiviteter kan være forvalgt til 1 eller skjult -1
                                            strSQLaktIdUseJob = "SELECT aktid, forvalgt FROM timereg_usejob WHERE (aktid <> 0 AND (forvalgt = 1 AND forvalgt_dt >= '"& useDateSt_30SQL &"') OR (forvalgt = 2)) AND medarb = "& usemrn  
                                    
                                            'if session("mid") = 1 then
                                            'Response.write "<hr>"& strSQLaktIdUseJob 
                                            'Response.flush
                                            'end if

                                   
                                            oRec6.open strSQLaktIdUseJob, oConn, 3 
                                            While not oRec6.EOF

                                            aktidsMtimSQL = aktidsMtimSQL & " OR a.id = " & oRec6("aktid")
                                            
                        
                                            '** Skal ikke tiføejs til options, da de job bliver vist på timereg. delen
                                            'aktidsSQLOptions = aktidsSQLOptions & " OR a.id = " & oRec6("aktid")
                                            'aktidsMtimWrt = aktidsMtimWrt & "#"& oRec6("aktid") & "#," 
                                    
                                            if oRec6("forvalgt") = 2 then
                                            aktidsLasWrt = aktidsLasWrt & "#"& oRec6("aktid") & "#,"
                                            end if 

                                            oRec6.movenext
                                            wend
                                            oRec6.close


                                            '*** Tilføjer alle dem med timebudget / forecast på
                                            '*** INDEN FOR FY?
                                            if cint(viskunForecast) = 1 then
                                                strSQLaktIdFc = "SELECT aktid FROM ressourcer_md WHERE aktid <> 0 AND medid = "& usemrn  
                                    
                                                oRec6.open strSQLaktIdFc, oConn, 3 
                                                While not oRec6.EOF

                                                aktidsMtimSQL = aktidsMtimSQL & " OR a.id = " & oRec6("aktid")
                                            
                                                '** Skal ikke tiføejs til options, da de job bliver vist på timereg. delen
                                                aktidsLasWrt = aktidsLasWrt & "#"& oRec6("aktid") & "#,"
                                            

                                                oRec6.movenext
                                                wend
                                                oRec6.close
                                            end if


                                             '*** Tilføjer ALTID alle dem fra intern
                                            select case lto
                                            case "oko", "adra", "xintranet - local"
                                               
                                                    strSQLaktIdInterne = "SELECT a.id AS aktid, "_
                                                    &" a.projektgruppe1, a.projektgruppe2, a.projektgruppe3, a.projektgruppe4, a.projektgruppe5, a.projektgruppe6, a.projektgruppe7, a.projektgruppe8, a.projektgruppe9, a.projektgruppe10 "_
                                                    &" FROM job AS j "_
                                                    &" LEFT JOIN aktiviteter AS a ON (a.job = j.id AND aktstatus = 1) WHERE j.risiko = -2 AND jobstatus = 1 AND aktstatus = 1 GROUP BY a.id" 
                                    
                                                    'if session("mid") = 1 then
                                                    'response.write strSQLaktIdInterne
                                                    'response.flush
                                                    'end if
                                                    
	                                                call medariprogrpFn(usemrn)
                                                    medariprogrpTxtArr = split(medariprogrpTxt, ",")
                                                    
                                                    oRec6.open strSQLaktIdInterne, oConn, 3 
                                                    While not oRec6.EOF
                                                    
                                                        aktidsjobProdgrpThisAkt = ""

                                                        '*** ADgang via projektgrupper **'
                                                        aktidsjobProdgrpThisAkt = "#"& oRec6("projektgruppe1") &"#,"_
                                                        & "#"& oRec6("projektgruppe2") &"#,"_
                                                        & "#"& oRec6("projektgruppe3") &"#,"_
                                                        & "#"& oRec6("projektgruppe4") &"#,"_                                                                
                                                        & "#"& oRec6("projektgruppe5") &"#,"_
                                                        & "#"& oRec6("projektgruppe6") &"#,"_                                                
                                                        & "#"& oRec6("projektgruppe7") &"#,"_
                                                        & "#"& oRec6("projektgruppe8") &"#,"_
                                                        & "#"& oRec6("projektgruppe9") &"#,"_
                                                        & "#"& oRec6("projektgruppe10") &"#,"

                                                            mprGrpC = 0
                                                            for mprGrpC = 0 TO UBOUND(medariprogrpTxtArr)
                                                    

                                                                  'if session("mid") = 1 then
                                                                 'response.write  "aktid: "& oRec6("aktid") &" aktidsjobProdgrpThisAkt: "& aktidsjobProdgrpThisAkt &", MedarbiGrp: #"& medariprogrpTxtArr(mprGrpC) &"#<br>"
                            
                                                                  'end if

                                                                if instr(aktidsjobProdgrpThisAkt, medariprogrpTxtArr(mprGrpC)) <> 0 then

                                                                aktidsMtimSQL = aktidsMtimSQL & " OR a.id = " & oRec6("aktid")
                                            
                                                                '** Skal ikke tiføejs til options, da de job bliver vist på timereg. delen
                                                                aktidsLasWrt = aktidsLasWrt & "#"& oRec6("aktid") & "#,"
                                                                exit for

                                                                end if

                                                                 'if session("mid") = 1 then
                                    
                                                                  'response.write  "aktidsMtimSQL: "& aktidsMtimSQL &"<br><br>"
                                                                 'end if

                                            
                                                             next

                                                  

                                                    oRec6.movenext
                                                    wend
                                                    oRec6.close

                            
                                            'if session("mid") = 1 then
                                            'response.write aktidsMtimSQL
                                            'response.flush
                                            'end if

                                              
                                            end select

                                 


                                             '*** Henter fjenet/skjulte (udelad) SAMT aktivteter fra personligaktivliste aktiviteter ***'
                                             '*** Følger jo også projektgrupper, timreg_usejob er afledt af projektgrupper på job ***'
                                  
                                            if cint(positiv_aktivering_akt_val) = 1 then
                                            useJobaktKri = "aktid <> 0 AND ((forvalgt = -1) OR (forvalgt = 1 AND forvalgt_dt <= '"& useDateSt_30SQL &"'))"
                                            else
                                            'Samme job kan findes flere gange, men det er ok
                                            'useJobaktKri = "(jobid <> 0 AND aktid = 0 AND forvalgt = 1) OR (aktid <> 0 AND ((forvalgt = -1) OR (forvalgt = 1 AND forvalgt_dt <= '"& useDateSt_30SQL &"')))"
                                            useJobaktKri = "((jobid <> 0 AND forvalgt = 1 AND aktid = 0) OR (aktid <> 0 AND forvalgt = -1) OR (aktid <> 0 AND forvalgt = 1 AND forvalgt_dt <= '"& useDateSt_30SQL &"'))" 
                                            end if

                                            strSQLaktIdUseJob = "SELECT aktid, jobid, forvalgt FROM timereg_usejob WHERE "& useJobaktKri &" AND medarb = "& usemrn &" GROUP BY jobid, aktid"
                                    
                                            'if session("mid") = 1 then
                                            'Response.write "<hr>"& strSQLaktIdUseJob 
                                            'Response.flush
                                            'end if

                                            oRec6.open strSQLaktIdUseJob, oConn, 3 
                                            While not oRec6.EOF

                                    

                                            if cint(positiv_aktivering_akt_val) = 1 then
                                               aktidsSQLOptions = aktidsSQLOptions & " OR a.id = " & oRec6("aktid")
                                               if oRec6("forvalgt") = "-1" then 'altid gemt (vises på tilføjlisten)
                                               aktidsSkjWrt = aktidsSkjWrt & "#"& oRec6("aktid") & "#," 
                                               end if
                                            else
                                        
                                                if oRec6("aktid") <> 0 then
                                                    if oRec6("forvalgt") = "-1" then 'altid gemt (vises på tilføjlisten)
                                                    aktidsSkjWrt = aktidsSkjWrt & "#"& oRec6("aktid") & "#,"
                                                    else 'Hvis der er timer indefor de senete 14 dage vises de på treglistne ellers på tilføj listen
                                                    'aktidsSQLOptions = aktidsSQLOptions & " OR aktid = " & oRec6("aktid")
                                                    end if
                                                else
                                                aktidsSQLOptions = aktidsSQLOptions & " OR a.job = " & oRec6("jobid")
                                        
                                                end if
                                     
                                        
                                         
                                            end if
                                    
                                            oRec6.movenext
                                            wend
                                            oRec6.close



        
                                            strSQLaktTimerIUge = "SELECT t.taktivitetid, SUM(timer) AS timeriuge, tmnavn, tmnr FROM timer AS t "  
                                            strSQLaktTimerIUge = strSQLaktTimerIUge &" WHERE (tmnr = "& usemrn &" AND tdato BETWEEN '"& useDateSt_30SQL &"' AND '"& useDateSlSQL &"')"
                                            strSQLaktTimerIUge = strSQLaktTimerIUge &" GROUP BY t.taktivitetid"
                                           
                                 
                                 
                                           'Response.write "<br>"& strSQLaktTimerIUge
                                           'Response.end
                                            z = 0
                                            oRec6.open strSQLaktTimerIUge, oConn, 3 
                                            While not oRec6.EOF

                                               if instr(aktidsSkjWrt, "#"& oRec6("taktivitetid") &"#") = 0 then
                                                aktidsMtimSQL = aktidsMtimSQL & " OR a.id = " & oRec6("taktivitetid")
                                                'aktidsMtimWrt = aktidsMtimWrt & ",#"& oRec6("taktivitetid") &"#"
                                        
                                                end if
                                        

                                      
                                        
                                    
                                            oRec6.movenext
                                            wend
                                            oRec6.close
                                    
                                    
                                    
                                            if cint(positiv_aktivering_akt_val) = 1 then
                                            aktidsSQLOptions = aktidsSQLOptions
                                            else
                                            tAktidsSQL = replace(aktidsSQL, "WHERE (", " OR ") 
                                            aktidsSQLOptions = aktidsSQLOptions & tAktidsSQL  
                                            end if


                                            aktidsSQL = aktidsMtimSQL  






                                   
                                    '*********************************************
                                    '*** Alm.liste sorteret efter job, kunde etc.
                                    '*********************************************
                                    else


                                            '*************************************************************************************************
                                            '*** Henter kun akt med ressource timer for valgte medarb                      *******************
                                            '*************************************************************************************************
                                            
                                            
                                            '*** Henter kun aktiviteter med forecast på ****'
                                            forecastAktids = " AND (a.id = 0"
                                            if cint(viskunForecast) = 1 then
                                              
                                                kforCastSQL = "SELECT aktid FROM ressourcer_md WHERE medid = " & usemrn & " AND jobid = "& jobid &" GROUP BY medid, aktid"
                                                oRec5.open kforCastSQL , oConn, 3
                                                while not oRec5.EOF 
                                                forecastAktids = forecastAktids & " OR a.id = "&oRec5("aktid") 
                                                oRec5.movenext
                                                wend
                                                oRec5.close       

                                                
                                                if cint(risiko) < 0 then '+ alle interne

                                                kforCastSQL = "SELECT a.id AS aktid FROM aktiviteter AS a WHERE a.job = "& jobid &" GROUP BY id"
                                                oRec5.open kforCastSQL , oConn, 3
                                                while not oRec5.EOF 
                                                forecastAktids = forecastAktids & " OR a.id = "&oRec5("aktid") 
                                                oRec5.movenext
                                                wend
                                                oRec5.close    


                                                end if



                                            end if

                                            forecastAktids = forecastAktids & ")"

                                            if cint(viskunForecast) = 1 then
                                            forecastAktids = forecastAktids
                                            else
                                            forecastAktids = ""
                                            end if

                                            '*************************************************************************************************





                                                if cint(intEasyreg) <> 1 AND visHRliste <> "1" AND risiko >= 0 then 'viser altid ALLE på HR listen. Bruger alm. projektgrupper selvom positiv_aktivering_akt_ er slået til

                                   
                                                '*********************************************************************************************************
                                                '*** Positiv tilmeling af aktiviteter slået til ***'
                                                '*********************************************************************************************************
                                                      midsMtimArrTxt = "0"                            
                                                    
                                                            



                                                                if cint(positiv_aktivering_akt_val) = 1 OR cint(visallemedarb_bl) = 1 then
                                                                aktidsMtimSQL = "WHERE (a.id = 0 "

                                                    
                                                                        if cint(visallemedarb_bl) = 1 then '** Vis alle medarbejdere på et job MMMI løsningen
                                                                
                                                                                            '** Er der søgt på et jobnr **

                                                                                            jobid = 0
                                                                                            if jobnr_bl <> "0" then
                                                                                            strJobnrJobnavnSog = " j.jobnr = '"& jobnr_bl &"'"
                                                                                            else
                                                                                            strJobnrJobnavnSog = " j.jobnr = 0"
                                                                                            end if


                                                                                            strSQLjob = "SELECT j.id, a.id AS aktid, "_
                                                                                            &" a.projektgruppe1, a.projektgruppe2, a.projektgruppe3, a.projektgruppe4, a.projektgruppe5, a.projektgruppe6, a.projektgruppe7, a.projektgruppe8, a.projektgruppe9, a.projektgruppe10 "_
                                                                                            &" FROM job AS j "_
                                                                                            &" LEFT JOIN aktiviteter AS a ON (a.job = j.id) "_ 
                                                                                            &" WHERE "& strJobnrJobnavnSog
                                                                        
                                                                                                'Response.write strSQLjob & "<hr>"
                                                                                                'Response.end
                                                                        

                                                                                            'aktidsjob = " OR a.id = 0 "
                                                                                            an = 0
                                                                                            oRec6.open strSQLjob, oConn, 3
                                                                                            while not oRec6.EOF 

                                                                                            jobid = oRec6("id")
                                                                        
                                                                                            aktidpjob(an) = aktidpjob(an) & " OR a.id = "& oRec6("aktid")

                                                                                            aktidsjob = aktidsjob & " OR a.id = "& oRec6("aktid")
                                                                        
                                                                                            aktidsjobProdgrp(an) = "#"& oRec6("projektgruppe1") &"#,"_
                                                                                            & "#"& oRec6("projektgruppe2") &"#,"_
                                                                                            & "#"& oRec6("projektgruppe3") &"#,"_
                                                                                            & "#"& oRec6("projektgruppe4") &"#,"_                                                                
                                                                                            & "#"& oRec6("projektgruppe5") &"#,"_
                                                                                            & "#"& oRec6("projektgruppe6") &"#,"_                                                
                                                                                            & "#"& oRec6("projektgruppe7") &"#,"_
                                                                                            & "#"& oRec6("projektgruppe8") &"#,"_
                                                                                            & "#"& oRec6("projektgruppe9") &"#,"_
                                                                                            & "#"& oRec6("projektgruppe10") &"#,"
                                                                       

                                                                       
                                                                                            an = an + 1
                                                                                            oRec6.movenext
                                                                                            wend 
                                                                                            oRec6.close
                                                                        
                                                                                            anEnd = an - 1
                                                                        
                                                                        
                                                                      


                                                                           '*** Medarb. med rettighder til job og aktiviteter via timereg_usejob ****'
                                                                          strSQLaktIdUseJob = "SELECT "_
                                                                          &" tu.medarb AS mid FROM timereg_usejob tu LEFT JOIN medarbejdere AS m ON (m.mid = tu.medarb) "_
                                                                          &" WHERE jobid = "& jobid & " AND m.mansat = 1 ORDER BY mnavn LIMIT 50" 

                                                                        else


                                                                          strSQLaktIdUseJob = "SELECT aktid FROM timereg_usejob WHERE medarb = "& usemrn &" AND jobid = "& jobid &" AND aktid <> 0"
        
                                                                        end if 'cint(visallemedarb_bl) = 1
                                                   
                                                             
    

                                                                       'Response.write strSQLaktIdUseJob & "<hr>"
                                                                       'Response.end
                                                  
                                                                        anMaTU = 0
                                                                        oRec6.open strSQLaktIdUseJob, oConn, 3 
                                                                        While not oRec6.EOF

                                                                    
                                                                                            if cint(visallemedarb_bl) = 1 then
                                                
                                                
                                                

                                                                                                    if instr(midsMtimArrTxt, ",#"& oRec6("mid") &"#") = 0 then

                                                                                                                'call projgrp(-1,level,oRec6("mid"),1)
                                                
                                                                                                            z = z + 1

                                                                                                            midsMtimArrTxt = midsMtimArrTxt & ",#"& oRec6("mid") &"#"
                                                

                                                                                                            aktidsPrMid(z) = " WHERE (a.id = 0"
                                                                                                     end if
                                                                                
                                                                                


                                                                                                    '** Medarbejder projektgrupper **'
                                                                                                    medarbPGrp = "#0#" 
                                                                                                    strMpg = "SELECT projektgruppeId, medarbejderId, teamleder FROM progrupperelationer WHERE medarbejderId = "& oRec6("mid") & " GROUP BY projektgruppeId"
                                                                        
                                                                                                        'Response.write strMpg & "<hr>"
                                                                                                        'Response.Flush

                                                                                                    oRec5.open strMpg, oConn, 3
                                                                                                    while not oRec5.EOF
        
                                                                                                    medarbPGrp = medarbPGrp & ",#"& oRec5("projektgruppeId") &"#"         
        
                                                                                                    oRec5.movenext
                                                                                                    wend
                                                                                                    oRec5.close 

                                                                        

                                                                                                            for i = 0 to anEnd
                                                                                                            '** Tjek at medarbejder har adtgang til aktivitetne via projektgruppen
                                                                                                            medarbPGrpArr = split(medarbPGrp, ",")
                                                                                                                for p = 0 TO UBOUND(medarbPGrpArr)
                                                                                                                if instr(aktidsjobProdgrp(i), medarbPGrpArr(p)) <> 0 then
                                                                                                                aktidsPrMid(z) = aktidsPrMid(z) & aktidpjob(i) 'aktidsjob  'aktidsPrMid(z) &  " OR a.id = "& aktidpjob(i) ' & oRec6("taktivitetid")
                                                                                                                end if
                                                                                                                next
                                                                                                            next
                            
                                        

                                                                                                   'Response.write "Her: "& aktidsPrMid(z) & "<br><br>"                                                        

                                                                                            else
                                                                                            '* Positiv aktivering adgang til specifik aktivtet ***'
                                                                                            aktidsMtimSQL = aktidsMtimSQL & " OR a.id = " & oRec6("aktid")

                                                                                            end if
                                                    

                                                                        anMaTU = anMaTU + 1
                                                                        oRec6.movenext
                                                                        wend
                                                                        oRec6.close

                                    
                                                                        aktidsSQL = aktidsMtimSQL
                                   
                                                              
        
        
                                                            end if '** if cint(positiv_aktivering_akt_val) = 1 OR cint(visallemedarb_bl) = 1 then


                                                      

                                                end if '**  if cint(intEasyreg) <> 1 AND visHRliste <> "1" AND risiko >= 0 then 

                                               
                                     
                                    
                                    end if '***** Blandetliste 14 dg/2md / Almvisning sorteret efter Job/Kunde




                                    
                                    strJobnrJobnavnSog = ""

                                    if jobnr_bl <> "0" then
                                    'strJobnrJobnavnSog = " AND (j.jobnr LIKE '"& jobnr_bl &"%' OR j.jobnavn LIKE '"& jobnr_bl &"%')"
                                    end if
                                    

                                
                                    'lastFase = ""
                                    antalmedarbids = 1
                                    '*** ER Der valgt alle medarb **'
                                    if cint(visallemedarb_bl) = 1 then
                                    
                                    'Response.write "midsMtimArrTxt: " & midsMtimArrTxt & "<br>"
                                    'Response.end                            

                                    midsMtimArr = replace(midsMtimArrTxt, "#", "")
                                    medarbids = split(midsMtimArr, ",")
                                    mnn = 0
                                        for mnn = 1 to UBOUND(medarbids) 'skal ikke ha nr 0 med
                                            antalmedarbids = antalmedarbids + 1  
                                            'Response.write "medarbids: "& medarbids(mnn) & "<br>"
                                                                      
                                        next

                                    else
                                    antalmedarbids = 1
                                    
                                    midsMtimArr = "0,1"
                                    medarbids = split(midsMtimArr, ",")

                                  
                                    end if
                                    




                                    'Response.write "antalmedarbids" & antalmedarbids 
                                    'Response.end
                                    lastMnr = 0
                                    for mn = 1 to UBOUND(medarbids)
                                    
                
                                    if cint(visallemedarb_bl) = 1 then
                                    aktidsSQL = aktidsPrMid(mn)
                                    usemrn = trim(medarbids(mn))
                                    call meStamdata(usemrn)
                                    'Response.write "##"&  medarbids(mn) & "<br>"
                                    end if
                                    
                                   
                                    InternJobOskriftIsWrt = 0
                                    wrtBreak = 0

                                    '******************************************************************************************************************************
                                    '********** MAIN aktivitetliste SQL ***'
                                    '**** Ved normalvisning, dvs. sorteret efter kunde, job etc. er projektgrupepr tjekket på forhånd på timereg. sidens alm load.
                                    strSQLakt = "SELECT a.id AS aid, j.id AS jid, navn AS aktnavn, a.fase, tidslaas, tidslaas_st, tidslaas_sl, tidslaas_man, tidslaas_tir, "_
                                    &"tidslaas_ons, tidslaas_tor, tidslaas_fre, tidslaas_lor, tidslaas_son, k.kkundenavn, k.kkundenr, k.kid, j.jobnavn, j.jobnr, j.risiko, j.jobstatus,  "_
                                    &" a.aktstartdato, a.aktslutdato, a.incidentid, a.fakturerbar, a.bgr, j.jobans1, j.jobans2, a.budgettimer AS aktbudgettimer, a.antalstk, a.beskrivelse AS abesk, brug_fasttp, fasttp, fasttp_val, v.valutakode AS iso "
                                    
                                    strSQLakt = strSQLakt &" FROM aktiviteter AS a "_
                                    &" LEFT JOIN job AS j ON (j.id = a.job) "_
                                    &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                                    &" LEFT JOIN valutaer AS v ON (v.id = a.fasttp_val) "
                                  
                                
                                    
                                    if intEasyreg = 1 then

                                         ordBySQL =  "a.fase, a.sortorder, a.navn"
                                         lmt = "0,150"
                                    
                                    else

                                        if cint(sortByVal) = 5 OR cint(sortByVal) = 6 then 
                                        
                                        select case lto
                                        case "oko", "adra", "xintranet - local"
                                        ordBySQL =  "a.sortorder, a.navn"
                                        case else
                                        ordBySQL =  "k.kkundenavn, j.jobnavn, a.fase, a.sortorder, a.navn"
                                        end select
                                        lmt = "0,100"
                                      

                                        
                                        else
                                        ordBySQL =  "a.fase, a.sortorder, a.navn"
                                        lmt = "0,100"
                                        end if 

                                    end if


                                   

                                    strSQLakt = strSQLakt & aktidsSQL &")" & strSQLAktDatoKri & " "& strSQLAktStatusKri &" "& forecastAktids &" "& strJobnrJobnavnSog &" GROUP BY a.id ORDER BY " & ordBySQL & " LIMIT " & lmt
                                    
                                    'Response.Write "<br> sql: "& strSQLakt & "<br>"
                                    'Response.write "lto" & lto
                                    'Response.flush
                                   
                                    
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

                                            call akttyper2009Prop(job_fakturerbar)

                                            if cint(aty_real) = 1 then
                                            aty_real_cls = "aty_real_cls_1"
                                            else
                                            aty_real_cls = "aty_real_cls_0"
                                            end if

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

                                            job_jobstatus = oRec6("jobstatus")

                                            select case job_jobstatus
                                            case 3
                                            job_jobstatusTxt = "<span style='font-size:9px; color:#000000;'> - Tilbud</span>"
                                            case else
                                            job_jobstatusTxt = ""
                                            end select 


                                            brug_fasttp = oRec6("brug_fasttp")
                                            fasttp = oRec6("fasttp")
                                            isoKode = oRec6("iso")

                                           

                                              '*********************************************************************************************************
                                              '*** AKTIVITETSLINJE ****'
                                              '*** AKT. NAVN AKTIVITETSNAVN ****
                                              '*** AKTIVITETER PÅ JOB ****'
                                              '*********************************************************************************************************
                                    


                                        'strAktiviteter = strAktiviteter & "<tr><td bgcolor='#FFFFFF' colspan='2' valign='top' style='padding:5px 5px 4px 2px; border:1px #cccccc solid; border-bottom:0px; border-left:0px;'>"& aktnavn &"</td>"
                                        'strAktiviteter = strAktiviteter & "</tr></table>"
                                          

                                                            '** Blandet liste medarb. overskift ****'
                                                             if cint(visallemedarb_bl) = 1 AND cdbl(lastMnr) <> cdbl(usemrn) then
                                                             strAktiviteter = strAktiviteter & "<tr><td colspan=9 bgcolor='#FFFFFF' style='padding:30px 0px 2px 2px; border-bottom:1px #cccccc solid;'><b>"& meNavn & " ("& meNr &")</b><br>"
                                                            strAktiviteter = strAktiviteter & ""& job_kundenavn &"<br>"
                                                            strAktiviteter = strAktiviteter & ""& job_jobnavn & " ("& job_jobnr &")"
                                                            strAktiviteter = strAktiviteter & "</td></tr>"

                                                             lastMnr = usemrn
                                                             end if



                                          
                                           '**** Faser ***'
                                           '*** Tilføj VZB iflg Cookie ***'
                                           if cint(intEasyreg) <> 1 AND cint(sortByVal) <> 5 AND cint(sortByVal) <> 6 then

                                                   if lastFase <> lcase(job_fase) AND len(trim(job_fase)) <> 0 AND job_fase <> "nonexx99w" AND job_fase <> "_" then
                                                    
				                                        strAktiviteter = strAktiviteter & "<tr><td bgcolor='#FFFFFF' colspan='9' valign='top' style='padding:0px 0px 0px 0px; height:10px; border-bottom:0px #cccccc solid;'>"_
                                                        &"<img src='../ill/blank.gif' width='1' height='1' border='0' alt='' /><br /></td></tr>"
                                                        
                                                        'strAktiviteter = strAktiviteter & dageDatoerTxt
            
			                                            strAktiviteter = strAktiviteter & "<tr><td bgcolor='#Eff3ff' colspan='2' valign='top' style='padding:5px 5px 4px 2px; border:1px #cccccc solid; border-bottom:0px; border-left:0px;'>"
			                                            
                                                            'select case lto
                                                            'case "intranet - local", "demo"
                                                             fsNavn = ""
			                                                'case else
                                                            'fsNavn = tsa_txt_329 & ": "
                                                            'end select         

                                                        strAktiviteter = strAktiviteter & "<a id='afaselnk_"&jobid&"_"&lcase(job_fase)&"' name='"&lcase(job_fase)&"' href='#"&lcase(job_fase)&"' class='afase'> + "& fsNavn &" "& replace(job_fase, "_", " ")&"</a></td>"
                
                                                                                           

                                                        strAktiviteter = strAktiviteter & "<td bgcolor='#FFFFFF' colspan='7' valign='top' style='padding:5px 5px 4px 2px; border-bottom:1px #cccccc solid;'>&nbsp;&nbsp;S&oslashg: <input type=""text"" id='fs_"&jobid&"_"&lcase(job_fase)&"' class=""fasesog"" style=""width:140px; border:1px #999999 solid;""></td></tr>"

                                                            
                                                            if len(trim(job_fase)) <> 0 then

                                                        
			                                                select case lto
                                                           
                                                            case "dencker", "oko" 'født med faser åbne
                                                            acnForFase = acnForFase + acn
                                                            tr_vzb = "visible"
			                                                tr_dsp = ""
                                                            case else

                                                            if acnForFaseIsWrt = 0 then
                                                            acnForFase = acnForFase + acn 
                                                            acnForFaseIsWrt = 1
                                                            else
                                                            acnForFase = acnForFase
                                                            end if
                                                            
                                                            tr_vzb = "hidden"
			                                                tr_dsp = "none"
                                                            end select

                                                                                                            

                                                            else
                                                                    
                                                            tr_vzb = "visible"
			                                                tr_dsp = ""
                                                            
                                                            end if
                                                      
                

                                                         '*** Faser ***'
			                                            strAktiviteter = strAktiviteter & "<input name=""faseshow"" id='faseshow_"&jobid&"_"&lcase(job_fase)&"' type=""hidden"" value='"& tr_vzb &"' />"
			                                            strAktiviteter = strAktiviteter & "<input name=""faseshowid"" type=""hidden"" value='"&jobid&"_"&lcase(job_fase) &"' />"
                                                        strAktiviteter = strAktiviteter & "<input name=""fasejobid"" id=""fasejobid_"& jobid &"_"& lcase(job_fase) &""" type=""hidden"" value='"&jobid&"' />"
                                                        '** end faser **'
			
			
			                                        else 
			                                               if len(trim(tr_vzb)) = 0 then
			                                                    'select case lto
                                                                'case "intranet - local", "demo"
                                                                'tr_vzb = "hidden"
			                                                    'tr_dsp = "none"
                                                                'case else
                                                                tr_vzb = "visible"
			                                                    tr_dsp = ""
                                                                'end select
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
                                          
                                          
                                          

                                          '**** Uger ****'
                                          '** Dato ooverskrifter ***'
                                          'call dageDatoer(1)
                                          if cint(intEasyreg) <> 1 AND (cint(sortByVal) = 5 OR cint(sortByVal) = 6) then
                                              if acnDtlinje >= 8 then
                                                  
                                                  if lastKnavn <> left(job_kundenavn,1) then
                                                  strAktiviteter = strAktiviteter & "<tr><td bgcolor='#FFFFFF' colspan='9' valign='top' style='padding:20px 0px 0px 10px;'>"_
                                                  & "<h4>"& ucase(left(job_kundenavn,1)) &" <span style=""color:#999999; font-size:10px;""> - "& job_kundenavn &"</span></h4></td></tr>"
                                                  strAktiviteter = strAktiviteter & dageDatoerTxt
                                                  acnDtlinje = 0
                                                  end if

                                              end if
                                          end if


                                          '*** Fravemarkering af nye ***'
                                          
                                          if cint(intEasyreg) = 1 OR (cint(sortByVal) = 5 OR cint(sortByVal) = 6) then
                                          
                                          if cdbl(nyakt) = cdbl(job_aid) then
                                          aktBGcol = "#FFFF99"
                                          else
                                            select case right(acn, 1)
                                            case 0,2,4,6,8
                                            aktBGcol = "#F7F7F7"
                                            case else
                                            aktBGcol = "#FFFFFF"
                                            end select
                                          end if      
                                         
                                          else

                                           aktBGcol = "#FFFFFF"

                                          end if
                                          
                                          if cint(intEasyreg) <> 1 AND (cint(sortByVal) = 5 OR cint(sortByVal) = 6) then
                                          strAktiviteter = strAktiviteter & "<tr class='td_"&job_aid&"' id='td_"&job_aid&"' style='visibility:"&tr_vzb&"; display:"&tr_dsp&"; background-color:"& aktBGcol &";'>"
                                          else
                                          strAktiviteter = strAktiviteter & "<tr class='td_"&jobid&"_"&lcase(job_fase)&"' id='"& job_aid &"' style='visibility:"&tr_vzb&"; display:"&tr_dsp&"; background-color:"& aktBGcol &";'>"
                                          end if
                                          
                                          'strAktiviteter = strAktiviteter & "<tr>"
				                          if visSimpelAktLinje = "1" OR visHRliste = "1" then
                                          aTdAlgn = "middle"
                                          else
                                          aTdAlgn = "top"
                                          end if
                                                        


                                        
                                          '** AKTIVITETS BASIS DATA **'
                                          strAktiviteter = strAktiviteter & "<td valign='"&aTdAlgn&"' height='30' style='padding-right:3px; border-bottom:1px #cccccc solid; background-color:"& aktBGcol &";'>"
                                
				                          strAktiviteter = strAktiviteter & "<input type='hidden' name='FM_jobid' id='FM_jobid_"& job_iRowLoops_cn &"' value='"& job_jid &"'>"

                                          strAktiviteter = strAktiviteter & "<input type='hidden' name='FM_mat_aktid' id='FM_mat_aktid_"& job_iRowLoops_cn &"' value='"& job_aid &"'>" 
                                          strAktiviteter = strAktiviteter & "<input type='hidden' name='FM_aktivitetid' id='FM_aktivitetid' value='"& job_aid &"'>" 
                                          strAktiviteter = strAktiviteter & "<input type='hidden' class='an_"&jobid&"_"&lcase(job_fase)&"' name='jq_aktivitetnavn' id='an_"& job_aid &"' value='"& left(aktnavn, 50) &"'>"
                                          if cint(visallemedarb_bl) = 1 then
                                          strAktiviteter = strAktiviteter & "<input type='hidden' name='FM_medid' id='FM_medid' value='"& medarbids(mn) &"'>"
                                          else
                                          strAktiviteter = strAktiviteter & "<input type='hidden' name='FM_medid' id='FM_medid' value='"& usemrn &"'>"
                                          end if  
                                             
                                                            


                                                        
                                                            '****************************** 
                                                            '** AKTIVITETSNAVN og DETALJER 
                                                            '******************************
                                             
                                                  

                                                            '** Kunde **' 
                                                            if (cint(intEasyreg) = 1 OR cint(sortByVal) = 5 OR cint(sortByVal) = 6) AND cint(visallemedarb_bl) <> 1 then
                                                                    
		                                                    
                                                                    '*** Kontakt ***
                                                                    '*** ÆØÅ **'
                                                                    call jq_format(job_kundenavn)
                                                                    job_kundenavn = jq_formatTxt
                                                                    
                                                                    if lto <> "oko" AND lto <> "xintrant - local" then
                                                                    strAktiviteter = strAktiviteter & "<span style='font-size:10px; color:#000000; line-height:10px;'>"& left(job_kundenavn, 50) &"</span><br>"
		                                                            end if
		        
		                                                            '*** jobnavn ****'
                                                                    '*** ÆØÅ **'
                                                                    call jq_format(job_jobnavn)
                                                                    job_jobnavn = jq_formatTxt

                                                                       'select case lcase(lto)
                                                                        'case "kejd_pb", "kejd_pb2", "intranet - local", "dencker", "outz"
                                                                        if visSimpelAktLinje = "1" then 'simpel
                                                                        jPx = 10
                                                                        aPx = 10
                                                                        jLft = 35 
                                                                        else
                                                                        
                                                                        
                                                                        if cint(intEasyreg) <> 1 AND (cint(sortByVal) = 5 OR cint(sortByVal) = 6) AND cint(visallemedarb_bl) <> 1 then
                                                                        jPx = 11
                                                                        aPx = 10
                                                                        else
                                                                        jPx = 10
                                                                        aPx = 10
                                                                        end if

                                                                        jLft = 35
                                                                        end if


                                
                                                                   

                                                                        if cint(InternJobOskriftIsWrt) = 1 then

                                                                        wrtBreak = 0
                                                                        else


                                                                           
                                                                                if level = 1 OR lcase(lto) = "kejd_pb" OR lcase(lto) = "kejd_pb2" then
                                                                                strAktiviteter = strAktiviteter & "<a href=""jobs.asp?menu=job&func=red&id="&job_jid&"&int=1&rdir=treg"" style=""font-size:"&jPx&"px; color:#000000; line-height:"&jPx+2&"px;"">"& left(job_jobnavn,jLft) &" ("& job_jobnr &")</a> "& job_jobstatusTxt &""
                                                                                else
                                                                                strAktiviteter = strAktiviteter & "<span style='font-size:"& jPx &"px; color:#000000; line-height:"&jPx+2&"px;'>"& left(job_jobnavn, jLft) & " ("& job_jobnr &")</span>"
                                                                                end if

                                                                           

                                                                        wrtBreak = 1
                                                                        
                                                                        if lto = "oko" AND cdbl(job_jid) = 3 then   
                                                                        InternJobOskriftIsWrt = 1
                                                                        end if

                                                                        end if
                                    
                                                                    


                                                                    if cint(intEasyreg) <> 1 AND (cint(sortByVal) = 5 OR cint(sortByVal) = 6) AND cint(visallemedarb_bl) <> 1 then
                                                                    
                                                            
                                                                            select case lto
                                                                            case "oko", "xintranet - local"

                                                                            case else

                                                                                strAktiviteter = strAktiviteter & "&nbsp;<span class='fjern_akt_bl' id='fjern_akt_"& job_aid &"' style=""color:darkred; font-size:9px; cursor:hand;"">[X]</span>"
                                                                    

                                                                                if instr(aktidsLasWrt, "#"& job_aid &"#") <> 0 then
                                                                                 strAktiviteter = strAktiviteter & "&nbsp;<span style=""cursor:pointer;""><img src=""../ill/las_on.gif"" class='sli_akt_bl' id='sli_akt_"& job_aid &"' border=0 alt=""Slip aktivitet""></span>&nbsp;"
                                                                    
                                                                                else
                                                                                 strAktiviteter = strAktiviteter & "&nbsp;<span style=""cursor:pointer;""><img src=""../ill/las.gif"" class='las_akt_bl' id='las_akt_"& job_aid &"' border=0 alt=""L&aring;s aktivitet""></span>&nbsp;"
                                                                    
                                                                                end if

                                                                           'strAktiviteter = strAktiviteter & "&nbsp;<span style=""color:#999999; font-size:9px; font-style:oblique;"">Fjern<br /><input class=""fjern_job"" type=""checkbox"" value=""1"" id='fjern_job_"&job_jid&"'/></span>"
                                                                            end select
                                                                    end if
                                                                    

                                                                    if cint(wrtBreak) = 1 then
                                                                    strAktiviteter = strAktiviteter & "<br>"
                                                                    end if

                                                                    
                                                                     '*** Fase ***'
                                                                    if len(trim(job_fase)) <> 0 AND trim(job_fase) <> "nonexx99w" AND job_fase <> " " then

                                                                    '*** ÆØÅ **'
                                                                    call jq_format(job_fase)
                                                                    job_fase = jq_formatTxt

                                                                    strAktiviteter = strAktiviteter & "<span style='font-size:9px; line-height:10px; color:#999999;'><b>fase:</b> "& left(job_fase, 30) &"</span><br>"
                                                                    end if 


                                                                    '** Aktnavn **'
                                                                    '*** ÆØÅ **'
                                                                    call jq_format(aktnavn)
                                                                    aktnavn = jq_formatTxt
				                                                    strAktiviteter = strAktiviteter & "<span style='font-size:"&aPx&"px; font-weight:lighter; color:#3B5998;'><b>"& left(aktnavn, 50) &"</b></span>"

                                                                   

                                                                    
                                                                 
                                                               
                                                                    if cint(intEasyreg) = 1 then
                                                                    strAktiviteter = strAktiviteter & "&nbsp;<span class=""fjeasyregakt"" id='"& job_aid &"' style=""color:darkred; font-size:9px; cursor:hand;"">[X]</span>"
                                                                    end if
                                                               
                                                                    'strAktiviteter = strAktiviteter & "<br>"
                                                                    if cint(visallemedarb_bl) = 1 then
                                                                    strAktiviteter = strAktiviteter & "<br><span style=""color:darkred; font-size:9px;"">"& meNavn & "</span>"
                                                                    end if
                                                               



                                                               else
                                                                       '*******************************************************                                 
                                                                       '** Aktnavn **' (Standard visning)
                                                                       '*******************************************************

                                                                       
                                                                       strAktiviteter = strAktiviteter & replace(aktnavn, "(ej fakturerbar)", "") 
                                                                                                                       

                                                                                        
                                                             end if

                                                             
                            
                                                            '***************************************************************************************
                                                            '*** AKTIVITETS LINJE DETALJER 
                                                            '***************************************************************************************
                                            

                                                             
                                                             if cint(intEasyreg) <> 1 then

                                                                   

                                                            
                                                                                '** Incident ID **'
                                                                                if job_incidentid <> 0 then
				                                                                strAktiviteter = strAktiviteter & "<br><span style='font-size:9px; line-height:10px; color:darkred;'><i>"& tsa_txt_036 &": "& job_incidentid &"</i></span>" 
				                                                                end if
                                                            
                                                                 

				                                                   
                                                                     
		                                                    
                                                                
                                                                            '******************************************************************************************'
                                                                            '***** Aktlinje detaljer. Budget, timepris, real. tid // Activities details, dates etc ****'
                                                                            '******************************************************************************************'


                                                                   



                                                                        resTimerThisM = 0
                                                                        restimeruspecThisM = 0
                                                                        SQLdatoKriTimer = ""
                                                                        if cint(aktBudgettjkOn) = 1 then
                                                                        '*******************************************
                                                                        '*** Ressource timer **'
                                                                        '*******************************************
                                                                       
                                                                                    
                                                                                    if month(useDateStSQL) < month(aktBudgettjkOnRegAarSt) then                                                                          
                                                                                       useRgnArr =  dateAdd("yyyy", -1, useDateStSQL)
                                                                                       useRgnArrNext = useDateStSQL
                                                                                    else
                                                                                       useRgnArr = useDateStSQL
                                                                                       useRgnArrNext =  dateAdd("yyyy", 1, useDateStSQL)

                                                                                    end if
                                                                     


                                                                        if cint(aktBudgettjkOn_afgr) <> 0 then
                                                                        sqlBudgafg = " AND ((md >= "& month(aktBudgettjkOnRegAarSt) & " AND aar = "& year(useRgnArr) & ") OR (md < "& month(aktBudgettjkOnRegAarSt) & " AND aar = "& year(useRgnArrNext) & "))" 
                                                                        else
                                                                        sqlBudgafg = ""
                                                                        end if

                                                                        strSQLres = "SELECT IF(aktid <> 0, SUM(timer), 0) AS restimer, IF(aktid = 0, SUM(timer), 0) AS restimeruspec FROM ressourcer_md WHERE ((jobid = "& job_jid &" AND aktid = "& job_aid &" AND medid = "& usemrn &") OR (jobid = "& job_jid &" AND aktid = 0 AND medid = "& usemrn &")) "& sqlBudgafg &"  GROUP BY aktid, medid" 
                                                                        'if session("mid") = 1 then
                                                                        'Response.write strSQLres & " aktBudgettjkOnRegAarSt = "& aktBudgettjkOnRegAarSt &" : useDateStSQL : "& useDateStSQL  &" // useDateSlSQL "& useDateSlSQL & "<br>"
                                                                        'Response.end 
                                                                        'end if


                                                                        oRec2.open strSQLres, oConn, 3
                                                                        while not oRec2.EOF 
                                                                        resTimerThisM = resTimerThisM + oRec2("restimer")
                                                                        restimeruspecThisM = restimeruspecThisM + oRec2("restimeruspec")
                                                                        oRec2.movenext
                                                                        wend
                                                                        oRec2.close
                                                                        

                                                                        end if

                                                           

                                                                        '** tot timeforbrug alle aktiviter til at holde op mod uspecificeret **
                                                                        '** kun akt. der tæller med i daglig timereg.                        **
                                            
                                                                       

                                                                        if cint(aktBudgettjkOn_afgr) <> 0 then

                                                                                    if month(useDateStSQL) < month(aktBudgettjkOnRegAarSt) then                                                                          
                                                                                       useRgnArrYear =  year(dateAdd("yyyy", -1, useDateStSQL))
                                                                                       useRgnArrSQLDate = useRgnArrYear &"/"& month(aktBudgettjkOnRegAarSt) &"/"& day(aktBudgettjkOnRegAarSt)

                                                                                       useRgnArrYear =  year(useDateStSQL)
                                                                                       useRgnArrDt = dateAdd("d", -1, aktBudgettjkOnRegAarSt)
                                                                                       useRgnArrNextSQLDate = useRgnArrYear &"/"& month(useRgnArrDt) &"/"& day(useRgnArrDt)
                                                                                       
                                                                                    else
                                                                                       useRgnArrYear =  year(useDateStSQL)
                                                                                       useRgnArrSQLDate = useRgnArrYear &"/"& month(aktBudgettjkOnRegAarSt) &"/"& day(aktBudgettjkOnRegAarSt)
                                                                                       
                                                                                       useRgnArrYear =  year(dateAdd("yyyy", 1, useDateStSQL))
                                                                                       useRgnArrDt = dateAdd("d", -1, aktBudgettjkOnRegAarSt)
                                                                                       useRgnArrNextSQLDate = useRgnArrYear &"/"& month(useRgnArrDt) &"/"& day(useRgnArrDt)

                                                                                    end if
                                                                    

                                                                                    SQLdatoKriTimer = " AND tdato BETWEEN '"& useRgnArrSQLDate &"' AND '"& useRgnArrNextSQLDate &"'"

                                                                        end if

                                                                      

                                                                        '** Timer real kun valgte medarb **'
                                                                        'aktdata(iRowLoop, 38) = 0
                                                                        akt_timerTot_medarb = 0
                                                                        job_timerTot_medarb = 0
                                                                        if cint(visAktlinjerSimpel_medarbrealtimer) = 1 OR cint(visAktlinjerSimpel_restimer) = 1 then
                                                                        '** Real. timer. Bruges også til balance på ressource ofrecast
                                                                       
                                                                            strSQLtmnr = "SELECT SUM(timer) AS timermnr FROM timer WHERE taktivitetid = "& job_aid & " AND tmnr = "& usemrn &" "& SQLdatoKriTimer &" GROUP BY taktivitetid"
                                                                        
                                                                       
                                                                            oRec2.open strSQLtmnr, oConn, 3
                                                                            if not oRec2.EOF then  
                                                                            akt_timerTot_medarb = oRec2("timermnr")
                                                                            end if
                                                                            oRec2.close



                                                                     
                                                                            strSQLtmnr = "SELECT SUM(timer) AS timermnr FROM timer WHERE tjobnr = '"& job_jobnr & "' AND tmnr = "& usemrn &" AND ("& aty_sql_realhours &") "& SQLdatoKriTimer &" GROUP BY tjobnr"
                                                                            'Response.write strSQLtmnr
                                                                            'Response.flush
                                                                            oRec2.open strSQLtmnr, oConn, 3
                                                                            if not oRec2.EOF then  
                                                                            job_timerTot_medarb = oRec2("timermnr")
                                                                            end if
                                                                            oRec2.close

                                                                    

                                                                        end if
                                                                



                                                                        if cint(aktBudgettjkOn) = 1 then

                                                                                    '** 1) Er der timebudget/forkalk. på akt. tjekkes for den enkelte aktivitet. 
                                                                                    '** 2) Er der uespecificeret budfget tjekeks der samlet på alle akt. på job
                                                                                    '** 3) Er der ikke budget vises svag lyserød markering og grå felter
                                                            
                                                                                    '*** HUSK BG farver ***'
                                                                                     resforecastMedOverskreddet = 0
                                                            
                                                                                    '** 1) Den enkelte akt
                                                                                    if resTimerThisM <> 0 then

                                                                                            if cdbl(akt_timerTot_medarb) >= cdbl(resTimerThisM) then
                                                                                            resforecastMedOverskreddet = 1
                                                                                            else
                                                                                            resforecastMedOverskreddet = 0
                                                                                            end if

                                                                                    else
                                                                    

                                                                                            '** 2) Uspec / all akt samlet
                                                                                            if restimeruspecThisM <> 0 then

                                                                                            if cdbl(job_timerTot_medarb) >= cdbl(resTimeruspecThisM) then
                                                                                            resforecastMedOverskreddet = 1
                                                                                            else
                                                                                            resforecastMedOverskreddet = 0
                                                                                            end if

                                                                                            else '*** IKKE noget forecast overhovedet

                                                                                            resforecastMedOverskreddet = 2

                                                                                            end if

                                                                                    end if


                                                                        end if 'aktBudgettjkOn


                                                           


                                                            
                                                         



                                                            'end if 'easyreg





                                                    


                                                    
                                                    if ((visSimpelAktLinje <> "1" AND visHRliste <> "1" AND job_internt > -1) AND _
                                                    (job_fakturerbar = 1 OR job_fakturerbar = 2 OR job_fakturerbar = 5 OR job_fakturerbar = 6 OR job_fakturerbar = 90)) _
                                                    OR (lto = "oko" AND left(aktnavn, 10) = "Intern tid") then 'fakturerbar, ikke fakbar, Km og Salg + E1
                                                                      

                                                   


                                                                      '***********************************
                                                                      '**** AktType ******
                                                                      '***********************************

                                                                      'lto = "wwf"
                                                                      'select case  lto
                                                                      'case "unik", "xintranet - local", "wwf", "mmmi"
                                                                      'case else
                                                                     
                                                                    if cint(visAktlinjerSimpel_akttype) = 1 then

                                                                        strAktiviteter = strAktiviteter & "&nbsp;<span style='font-family:Arial; line-height:11px; font-size:9px; color:#6CAE1C;'>("
				                                                         call akttyper(job_fakturerbar, 1)

                                                                         '*** ÆØÅ **'
                                                                        call jq_format(akttypenavn)
                                                                        akttypenavn = jq_formatTxt

                                                                        strAktiviteter = strAktiviteter & lcase(akttypenavn) & ")</span>"
                                                                        'end select
                                                              
                                                                        if job_fakturerbar = 15 OR job_fakturerbar = 12 then 
                                                                         strAktiviteter = strAktiviteter & "<span style='line-height:11px; font-size:10px; color:#999999;'><br>Gns. antal timer pr. dag baseret p&aring; norm. timer pr. uge / 5 dages arb. uge.</span>"
                                                                        end if

                                                                          
                                                                
                                                                     end if

                                                                      strAktiviteter = strAktiviteter & "<br>"

                                                                    '*****************************************


                                                                    'Resten af info evt. kun på fakturerbare job_fakturerbar = 1 



                                                                    '**************************************************'
                                                                    '** Datoer **'
                                                                    '**************************************************'
                                                                    
                                                                    'select case lto
                                                                    'case "mmmi", "wwf", "unik"

                                                                    'case else
                                                                    
                                                                     if cint(visAktlinjerSimpel_datoer) = 1 then

                                                                    strAktiviteter = strAktiviteter & "<span style='font-size:9px; line-height:12px; color:#5582d2;'>"
                                                                    strAktiviteter = strAktiviteter & datepart("d", aktstartdato, 2,2) &". "& left(monthname(datepart("m", aktstartdato, 2,2)), 3) &". "& datepart("yyyy", aktstartdato, 2,2) & " - " 
				                                                    strAktiviteter = strAktiviteter & datepart("d", aktslutdato, 2,2) &". "& left(monthname(datepart("m", aktslutdato, 2,2)), 3) &". "& datepart("yyyy", aktslutdato, 2,2)
				                                                    strAktiviteter = strAktiviteter & "</span><br>"
                                                    
                                                                    end if

                                                                    'end select

                                                                    
		                                                             

                                                                     '*******************************************************'
                                                                     '***** Timepris ****************************************'
                                                                     '*******************************************************

                                                                    'select case lto
                                                                    'case "mmmi", "wwf", "unik"

                                                                    'case else
                                
                                                                      if cint(visAktlinjerSimpel_medarbtimepriser) = 1 then
                                                                     strAktiviteter = strAktiviteter & "<span style='font-family:Arial; font-size:9px; color:#999999;'>"

                                                                    '** Timepris **'
                                                                    '** Admin eller jobansvarlig **'
                                                                    if level = 1 OR session("mid") = job_jobans1 OR session("mid") = job_jobans2 then
                                                                            
                                                                            
                                                                            if cint(brug_fasttp) <> 1 then
                                                                            
                                                                            call akttyper2009Prop(job_fakturerbar)


				                                                                    if cint(aty_fakbar) = 1 OR job_fakturerbar = 5 then '** Viser tpriser på fakturerbare eller kørsel
				
				                                                               
				                                                                
				                                                                                '*** tjekker for alternativ timepris på aktivitet
				                                                                                call alttimepris(job_aid, jobid, usemrn, 0)  
				      
				                                                                                '** Er der alternativ timepris på jobbet
                                                                                                if foundone = "n" then
                                                                                                call alttimepris(0, jobid, usemrn, 0) 
                                                                                                end if
				       
				                                                                                aktTpThis = formatnumber(intTimepris, 2) & " " & valutaISO
				                                                                       
				
				                                                                    strAktiviteter = strAktiviteter & " timepris: </span><span style='font-family:Arial; font-size:9px; color:#5582d2;'>" & aktTpThis & "</span><br>"  
				
				                                                                    end if

                                                                            else

                                                                            strAktiviteter = strAktiviteter & " timepris: </span><span style='font-family:Arial; font-size:9px; color:#5582d2;'>" & formatnumber(fasttp,2) & " " & isoKode &"</span><br>"  

                                                                            end if
				
				
				
				
				                                                       end if 'level

                                                                       end if 'visAktlinjerSimpel_medarbtimepriser
                                                                       'end select
                                                                       '*******************************************************************************
                                                           
                                                           
                                                           

                                                           '*****************'
                                                           '**** Budget *****'
                                                           'select case lto
                                                           'case "mmmi", "wwf", "unik"
                                                           
                                                           'case else

                                                            if cint(visAktlinjerSimpel_timebudget) = 1 then
                                                                    
                                                            select case job_bgr 
                                                            case 1 
                                                            job_bgr_Txt = "<span style='font-family:Arial; font-size:9px; color:#999999; line-height:12px;'> Forkalk.: </span><span style='font-family:Arial; font-size:9px; line-height:12px; color:#5582d2;'>" & formatnumber(job_aktbudgettimer, 2)  & " t."
                                                            case 2
                                                            job_bgr_Txt = "<span style='font-family:Arial; font-size:9px; color:#999999; line-height:12px;'> Stk.: </span><span style='font-family:Arial; font-size:9px; line-height:12px; color:#5582d2;'>"& formatnumber(job_antalstk, 2)  & " stk."
                                                            case else
                                                            job_bgr_Txt = ""
                                                            end select
                                                                
                                                            if len(trim(job_bgr_Txt)) <> 0 then
                                                            strAktiviteter = strAktiviteter & job_bgr_Txt & " "
                                                            end if

                                                            strAktiviteter = strAktiviteter & "</span>"

                                                            end if '** visAktlinjerSimpel_timebudget
                                                           ' end select


                                                            '**************************************************'
                                                            '** Timer real **'
                                                            '**************************************************'

                                                            

                                                            if cint(visAktlinjerSimpel_realtimer) = 1 _
                                                            OR (lto = "oko" AND left(aktnavn, 10) = "Intern tid") then
                                                    
                                                
                                                                if cint(visAktlinjerSimpel_realtimer) = 1 then

                                                                akt_timertot = 0
                                                                strSQLttot = "SELECT SUM(timer) AS timertot FROM timer WHERE taktivitetid = "& job_aid &" "& SQLdatoKriTimer &" GROUP BY taktivitetid"
                                                                oRec2.open strSQLttot, oConn, 3
                                                                if not oRec2.EOF then  
                                                                    akt_timertot = oRec2("timertot")
                                                                end if
                                                                oRec2.close

                                                           


                                                                if job_bgr <> 2 then 'bg = timer ell. IKKE angivet
                                                                    if akt_timertot > job_aktbudgettimer then
                                                                    fcolreal = "red"
                                                                    else
                                                                    fcolreal = "#2c962d"
                                                                    end if
                                                                else 'antal
                                                                fcolreal = "#5582d2"
                                                                end if

                                                            
                                                          
                                                                strAktiviteter = strAktiviteter & "<span style='font-family:Arial; font-size:9px; line-height:12px; color:#999999;'> Real.: <span style='font-family:Arial; font-size:9px; line-height:12px; color:"&fcolreal&";'>"& formatnumber(akt_timertot, 2) &"</span>"

                                                                end if 


                                                            if cint(visAktlinjerSimpel_medarbrealtimer) = 1 then
                                                            strAktiviteter = strAktiviteter & " (egne: <span style='font-family:Arial; font-size:9px; color:#5582d2;'>"& formatnumber(akt_timerTot_medarb, 2) &"</span>)"
                                                            end if

                                                             
                                                            if (lto = "oko" AND left(aktnavn, 10) = "Intern tid") then
                                                            strAktiviteter = strAktiviteter & "<span style='font-family:Arial; font-size:9px; color:#5582d2;'>Real. timer: <b>"& formatnumber(akt_timerTot_medarb, 2) &"</b></span>"
                                                            end if
                                                            
                                                            strAktiviteter = strAktiviteter & "<br>"
                        
                                                            

                                                           end if ' if cint(visAktlinjerSimpel_realtimer) = 1 then
                                                           '****************************************




                                                          
                                                            '****************************************
                                                            '**** Forcast medarb *****'
                                                            '****************************************
                                                            

                                                            if (cint(aktBudgettjkOn) = 1) AND (job_fakturerbar = 1 OR job_fakturerbar = 90) then
                                                                         
                                                              
                                                                     
                                                                    
                                                                    if cint(resforecastMedOverskreddet) = 2 then 'ikke noget forecast

                                                                    resTimerThisMSaldo = formatnumber((resTimerThisM - akt_timerTot_medarb), 2)
                                                                    tmforbrugTxt = akt_timerTot_medarb

                                                                    else

                                                                     if cdbl(resTimerThisM) <> 0 then
                                                                    resTimerThisMSaldo = formatnumber((resTimerThisM - akt_timerTot_medarb), 2)
                                                                    tmforbrugTxt = akt_timerTot_medarb
                                                                    else 'uspec
                                                                    resTimerThisMSaldo = formatnumber((restimeruspecThisM - job_timerTot_medarb), 2)
                                                                    tmforbrugTxt = job_timerTot_medarb
                                                                    end if

                                                                    end if

                                                                    if cint(resforecastMedOverskreddet) = 1 then
                                                                    resFcol = "red"
                                                                    resBGcol = "#ffdfdf"
                                                                    resTxt = " = <span style=""color:red; background-color:#ffdfdf; font-size:9px; line-height:12px;""><b>"& resTimerThisMSaldo &"</b></span>"
                                                                    'resFBgcol = "red"
                                                                    else

                                                                        if cint(resforecastMedOverskreddet) = 2 then '** Ikke noget forecast overhovedet Hverken på akt. eller uspec
                                                                        resFcol = "#999999"
                                                                        resBGcol = "#F7F7F7"
                                                                         resTxt = " = <span style=""color:#999999; background-color:#F7F7F7; font-size:9px; line-height:12px;""><b>"& resTimerThisMSaldo &"</b></span>"
                                                                        else
                                                                        resFcol = "#5582d2"
                                                                        resBGcol = "#FFFFFF"
                                                                          resTxt = " = <span style=""color:#5582d2; font-size:9px; line-height:12px;""><b>"& resTimerThisMSaldo &"</b></span>"
                                                                        end if
                                                                    'resFBgcol = "#FFFFFF"

                                                                    end if

                                                                
                                                                if cint(visAktlinjerSimpel_restimer) = 1 then

                                                                strAktiviteter = strAktiviteter & "<span style='font-family:Arial; font-size:9px; line-height:12px; color:#999999;'>Forecast - Real.: </span><span style='font-family:Arial; font-size:9px; line-height:12px; white-space:nowrap; color:"& resFcol &";'>"
                                                                
                                                                
                                                                'if cint(resforecastMedOverskreddet) <> 2 then 'ikke angivet forecast
                                

                                                                if cdbl(restimeruspecThisM) > 0 AND cdbl(restimerThisM) = 0 then '** der findes uspecificerede og der er IKKE angivet forecast / budget på akt.
                                                                strAktiviteter = strAktiviteter & formatnumber(restimeruspecThisM, 2) &" (uspec.) - "& formatnumber(tmforbrugTxt, 2) &"</span> " & resTxt 
                                                                else
                                                                strAktiviteter = strAktiviteter & formatnumber(resTimerThisM, 2) &" - "& formatnumber(tmforbrugTxt, 2) &"</span>" & resTxt 
                                                                end if

                                                        
                                                                end if '**  if cint(visSimpelAktLinje_realtimer) = 1 then
                                                                


				                                            end if '** aktBudgettjkOn

                                                           


				                                            end if '** visSimpelAktLinje









                                                            if visHRliste = "1" then
                                                            
                                                                    if job_fakturerbar = 15 OR job_fakturerbar = 12 OR job_fakturerbar = 111 OR job_fakturerbar = 112 then 
                                                                     strAktiviteter = strAktiviteter & "<span style='line-height:11px; font-size:10px; color:#999999;'><br>Gns. antal timer pr. dag baseret p&aring; norm. timer pr. uge / 5 dages arb. uge.</span>"
                                                                    end if

                                                            end if

				
				                                                          
				                                                         '*** Aktivitets beskrivelse ***'
				                                                         if len(trim(job_abesk)) <> 0 then
		                                                              
		                                                                strAktiviteter = strAktiviteter & "<div id='div_aktbesk_"& job_aid &"' style='position:relative; visibility:visible; background-color:"& aktBGcol &"; display:; left:0px; top:0px; width:200px; height:14px; overflow:hidden; z-index:200;'>"
		                                                                strAktiviteter = strAktiviteter & "<table bgcolor='"& aktBGcol &"' border='0' cellspacing='0' cellpadding='0' width='100%' ><tr><td style='padding:0px;' valign=top class=lille>"
		                                                                
		                                                                
		                                                                strAktiviteter = strAktiviteter & "<a class='a_showhideaktbesk' id='aktbesk_"& job_aid &"' href='#'>"& job_abesk &"</a></td></tr></table></div>"
		                                                                                    
		                                                                                        
		                                                                end if
				                                            



                                                              






                                                            end if 'easyreg



                                                            '** gl. stopurs link **'
				                                            strAktiviteter = strAktiviteter & "</td><td valign=top style='padding:8px 3px 3px 3px; border-bottom:1px #cccccc solid; border-right:1px #cccccc solid; background-color:"& aktBGcol &";'>&nbsp;</td>"
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
                                            timerNul_fundet_via_slettet_preval = 0
                                            destinationThis = 0
                                            origin = 0


                                            'if t = 100 then
                                            SQldatoX = year(tjekdag(x)) &"/"& month(tjekdag(x)) &"/"& day(tjekdag(x))    
				                            strSQLtim = "SELECT timer, timerKom, tfaktim, bopal, godkendtstatus, offentlig, sttid, sltid, destination, origin FROM timer WHERE taktivitetid = "& job_aid &" AND tdato = '"& SQldatoX &"' AND tmnr = "& usemrn

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
                                            origin = oRec2("origin")                              

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

                                            if cdbl(timerThis) = 0 then
                                            timerNul_fundet_via_slettet_preval = 1
                                            end if
                                            oRec2.movenext
                                            wend
                                            oRec2.close

                                            'end if


                                            '*** ÆØÅ **'
                                            call jq_format(timerKomThis)
                                            timerKomThis = jq_formatTxt
                                      
                                            
                                         

						                   		
								
								                    
                                                    '*** Tjekker om det er en helligdag **'
                                                    call helligdage(tjekdag(x), 0, lto)
	                                                'helligdageCol(x) = erHellig

								
								                    if len(timerKomThis) <> 0 then
								                    kommtrue = "+"
								                    else
								                    kommtrue = "<font color='#999999'>+</font>"
								                    end if 
						
								
								                    if x = 6 OR x = 7 OR erHellig = 1 then
								                    bgColor = "gainsboro"
								                    else
								                    bgColor = aktBGcol '"#FFFFFF"
								                    end if
								                    


                                                    call akttyper2009Prop(tfaktimThis) 

                                                    if cdbl(Timerthis) <> 0 OR timerNul_fundet_via_slettet_preval = 1 then

                                                                'if level <> 1 then
                                                                erIndlast = 1 '0
                                                                'else
                                                                'erIndlast = 0
                                                                'end if

                                                    else
                                                                

                                                                '** Preudfyldt Dage og projektgrupper ***'
                                                                aty_pre_prg_arr = split(aty_pre_prg, ",")
                                                                projgrpOk = 0
                                                                for p = 0 to UBOUND(aty_pre_prg_arr)
                                                                if instr(medariprogrpTxt, "#"& trim(aty_pre_prg_arr(p)) &"#") then                      
                                                                   projgrpOk = 1
                                                                end if  
                                                                next
                                                        

                                                                if aty_pre <> 0 AND instr(aty_pre_dg, x) <> 0 AND erHellig <> 1 AND cint(projgrpOk) = 1 then
								                                    
                                                               
                                                
                                                  
                                                                       '*** Der skal kun være indtastest pre_val hvis der ikke allerede er indtastet timer på dagen ***'        
                                                                        preVal = ""
                                                                        timerThisDayPreval = 0
                                                                        sqlDatoIdag = year(tjekdag(x)) &"/"& month(tjekdag(x)) &"/"& day(tjekdag(x))
                                                                        strSQLtimerAlleredeInd = "SELECT ROUND(SUM(timer),2) AS sumtimer FROM timer WHERE ("& aty_sql_realhours &") AND tdato = '"& sqlDatoIdag &"' AND tmnr = "& usemrn &" GROUP BY tmnr"
                                                                        oRec2.open strSQLtimerAlleredeInd, oConn, 3
                                                                        if not oRec2.EOF then                         
                                                                        timerThisDayPreval = oRec2("sumtimer")
								                                        end if                    
                                                                        oRec2.close

                                                                        if timerThisDayPreval = 0 then
                                                                        preVal = aty_pre
                                                                        end if

                                                                else
								                                preVal = ""
								                                end if

                                                                TimerThis = preVal
                                                                erIndlast = 0


                                                    end if





								                    
								                    '** tjek erIndlast = 0
                                                   
                                                      
                                                    if cint(intEasyreg) <> 1 then

								                    call fakfarver(lastfakdato, tjekdag(1), tjekdag(x), erIndlast, usemrn, tfaktimThis, job_tidslass_1, job_tidslass_2, job_tidslass_3, job_tidslass_4, job_tidslass_5, job_tidslass_6, job_tidslass_7, job_tidslass, job_tidslassSt, job_tidslassSl, ignorertidslas, godkendtstatus, job_internt, origin, resforecastMedOverskreddet)
								                    
                                                    'Response.Write "her  lastfakdato " & lastfakdato
                                                    'Response.end
                                                    felt = felt & "<td valign='top' bgcolor='"&bgColor&"' style='border-bottom:1px #cccccc solid; border-right:1px #cccccc solid; padding-top:6px;' id='td_"&job_iRowLoops_cn&""&x&"'>"
								                     
								
								                    if visTimerElTid <> 1 OR (cint(tfaktimThis) <> 1 AND cint(tfaktimThis) <> 2) OR aty_pre <> 0 OR cint(intEasyreg) = 1 then 'OR job_internt <= -1
								                    '** Alm. timefelt
                                                    timtype = "text"
								                    tidtype = "hidden"
								                    br = ""
                                                    showdageflt = 1
								                    
                                                    else
								                    '** Klokkeslet
                                                    timtype = "hidden"
								                    tidtype = "text"
								                    br = "<br>"
								                    showdageflt = 0
                                                    
                                                    end if
								                    

                                                    '*********************************************
                                                    '******* TIME FELTER *************************
                                                    '*********************************************
                                                    if (maxl <> 0) OR timtype = "hidden"  then 'Timefelt åben
								                        
                                                        if cint(origin) <> 0 then
                                                        felt = felt &"<input type='hidden' name='FM_timer' value='-9003'>" 'fra TT 
                                                        else
                                                        felt = felt &"<input class='dcls_"&x&" "&aty_real_cls&"_"&x&"' type='"& timtype &"' name='FM_timer' id='FM_timer_"&job_iRowLoops_cn&"_"&x&"' value='"& timerThis &"' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; border:"&fmborcl&"; height:16px; width:"& tfeltwth &"px; font-family:arial; font-size:10px;"" onkeyup=""doKeyDown()"" onfocus=""markerfelt('"&job_iRowLoops_cn&""&x&"','"&x&"'), showKMdailog('"&job_fakturerbar&"', '"&job_iRowLoops_cn&""&x&"', '"& bopalThis &"', '"& job_iRowLoops_cn &"', '"& job_kid &"')"";>"
						                                end if		                   

                                                    else 'Timefelt Lukket 
                                                    'uge lukket, / admin / ell. vis som klokkeslet, så skal timefeltet altid være hidden og ikke en span
                                                    'Eller fra TimeTag eller TimeOut mobile?? 20160104
                                                    
                                                        if cint(origin) <> 0 then 'Må ikke opdatere timer på registrering fra Excel, Vietnam, TiemTag, Timeout mobile. Da det dobler op pga. der kan være flere linjer der er grundlag
                                                        felt = felt &"<input type='hidden' name='FM_timer' value=''>" 'Brug evt -9003 ligesom ved klokkeslet
								                        else
                                                        felt = felt &"<input type='hidden' name='FM_timer' value='"& timerThis &"'>" 
								                        end if

                                                        felt = felt &"<div style='position:relative; background-color:"&fmbgcol&"; top:5px; border:0; height:12px; padding:3px; width:"& tfeltwth &"px; font-family:arial; font-size:10px; line-height:12px;'>"& timerThis
                                
                                                        if cint(origin) <> 0 then
                                                        felt = felt &"<span style=""font-size:9px; color:#999999;""> (se ugeseddel)</span>"
                                                        end if

                                                        felt = felt &"</div>"

                                                    end if
                                                    
                                                    felt = felt &"<input type='hidden' class='"& aty_real_cls &"_opr_"&x&"' name='FM_timer_opr' id='FM_timer_opr_"&job_iRowLoops_cn&"_"&x&"' value='"& timerThis &"'>"
								                    'felt = felt &"<input type='text' name='FM_aktty' id='FM_aktty_"&job_iRowLoops_cn&"_"&x&"' value='"& tfaktimThis &"'>"

                                                    felt = felt & "<input type='hidden' name='FM_mat_dato' id='FM_mat_dato_"& job_iRowLoops_cn &"_"&x&"' value='"& tjekdag(x) &"'>"  
		                                 
                                                     

                                                    '***** Vis som dage på fravæers typer ***'
                                                    if showdageflt = 1 then

                                                   

                                                  

                                                    select case job_fakturerbar
                                                    case 8,11,12,13,14,15,16,17,18,19,20,21,22,23,26,27,28,29,50,51,52,91,111,112,115,120,121,122,125 '** Ferie + Sygdom + E2
                                                        

                                                                'Tildel ferie, ferie u løn, overført og tildel ferie fridag skal ingnorerer helligdage, Omsorg 
                                                                if job_fakturerbar = 15 OR job_fakturerbar = 12 OR job_fakturerbar = 111 OR job_fakturerbar = 112 _
                                                                OR ((job_fakturerbar = 50 OR job_fakturerbar = 51 OR job_fakturerbar = 52) AND (lto = "tec" OR lto = "esn")) _
                                                                OR job_fakturerbar = 120 OR job_fakturerbar = 121 OR job_fakturerbar = 122 then 
                                                                
                                                                nTimDagFlt = normTimerGns5
                                                                else
                                                                    select case x
                                                                    case 1
                                                                    nTimDagFlt = normtimer_man
                                                                    case 2
                                                                    nTimDagFlt = normtimer_tir
                                                                    case 3
                                                                    nTimDagFlt = normtimer_ons
                                                                    case 4
                                                                    nTimDagFlt = normtimer_tor
                                                                    case 5
                                                                    nTimDagFlt = normtimer_fre
                                                                    case 6
                                                                    nTimDagFlt = normtimer_lor
                                                                    case 7
                                                                    nTimDagFlt = normtimer_son
                                                                    end select 
                                                                end if


                                                            
                                                            if cdbl(nTimDagFlt) <> 0 AND len(trim(timerThis)) <> 0 AND isNull(timerThis) <> true then ' AND 
                                                            timerThisBeregnDg = timerThis
                                                   

                                                            dageThis = formatnumber(timerThisBeregnDg/nTimDagFlt)

                                                            if right(dageThis, 2) <> "00" then
                                                            dageThis = formatnumber(timerThisBeregnDg/nTimDagFlt, 1)
                                                            else
                                                            dageThis = formatnumber(timerThisBeregnDg/nTimDagFlt, 0)
                                                            end if
                                                            'dageThis = 1
                                                            else
                                                            dageThis = ""
                                                            
                                                            end if
                                                        
                                                            
                                                   

                                                     
                                                             if (maxl <> 0) then 'uge lukket / admin
                                                             felt = felt &"<br><span style=""font-size:8px; color:#999999;"">"&tsa_txt_381&":<input class='dagecls' type='text' name='FM_dager' id='FM_dager_"&job_iRowLoops_cn&"_"&x&"' value='"& dageThis &"' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; border:"&fmborcl&"; height:14px; width:24px; font-family:arial; font-size:9px; line-height:12px;"" onkeyup=""doKeyDown()"" onfocus=""markerfelt('"&job_iRowLoops_cn&""&x&"','"&x&"')"";></span>"
                                                             else
                                                             felt = felt &"<br><span style=""font-size:8px; color:#999999;"">"&tsa_txt_381&":<input type='hidden' name='FM_dager' value='"& dageThis &"'></span><span style=""background-color:"&fmbgcol&"; border:"&fmborcl&"; height:14px; width:24px; padding:1px; font-family:arial; font-size:9px; line-height:12px;"">"& dageThis &"</span>"
                                                             end if


                                                                felt = felt &"<input type='hidden' name='FM_normt' id='FM_normt_"&job_iRowLoops_cn&"_"&x&"' value='"& nTimDagFlt &"'>"
                                                     
                                                             case else
                                                     
                                                             felt = felt &"<input type='hidden' name='FM_dager' id='FM_dager_"&job_iRowLoops_cn&"_"&x&"' value=''>"
                                                     
                                                            end select

                                                    else '** Hvis som start / slut tid
                                                        
                                                      felt = felt &"<input type='hidden' name='FM_dager' id='FM_dager_"&job_iRowLoops_cn&"_"&x&"' value=''>"
                                                    

                                                    end if

                                                    felt = felt &"<input type='hidden' name='FM_akttype' id='FM_akttp_"&job_iRowLoops_cn&"_"&x&"' value='"& tfaktimThis &"'>" 
                                                    
                                                    
                                                    if (maxl <> 0 OR tidtype = "hidden") then 'uge lukket / admin eller hvis hvis som timer skal klokkeslet felter altid være hidden. (hvis ikke span)
                                                  
                                                    felt = felt &"<input type='"& tidtype &"' name='FM_sttid' id='FM_sttid' value='"& timer_sttid &"' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; height:16px; width:"& tfeltwth-10 &"px; font-family:arial; font-size:10px; border:"&fmborcl&";"">"_
								                    &""& br &"<input type='"& tidtype &"' name='FM_sltid' id='FM_sltid' value='"& timer_sltid &"' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; height:16px; width:"& tfeltwth-10 &"px; font-family:arial; font-size:10px; border:"&fmborcl&";"">"
								                   
                        
                                                    else

                                                    
                                                        '"&fmborcl&";
                                                  
                                                   
                                    
                                                        if cint(origin) <> 0 then 'Må ikke opdatere timer på registrering fra Excel, Vietnam, TiemTag, Timeout mobile. Da det dobler op pga. der kan være flere linjer der er grundlag
                                                        felt = felt &"<input type='hidden' name='FM_sttid' value=''><span style=""background-color:"&fmbgcol&"; height:12px; width:"& tfeltwth-10 &"px; font-family:arial; font-size:10px; padding:3px; line-height:12px; border:0px;"">"& timerThis &"</span>"_
								                        &""& br &"<input type='hidden' name='FM_sltid' value=''><span style=""font-size:9px; color:#999999;""> (se ugeseddel)</span>"
								                
								                        else
                                                        
                                                        felt = felt &"<input type='hidden' name='FM_sttid' value='"& timer_sttid &"'><span style=""background-color:"&fmbgcol&"; height:12px; width:"& tfeltwth-10 &"px; font-family:arial; font-size:10px; padding:3px; line-height:12px; border:0px;"">"& timer_sttid &"</span>"_
								                        &""& br &"<input type='hidden' name='FM_sltid' value='"& timer_sltid &"'><span style=""background-color:"&fmbgcol&"; height:12px; width:"& tfeltwth-10 &"px; font-family:arial; font-size:10px; padding:3px; line-height:12px; border:0px;"">"& timer_sltid &"</span>"
								                
								                        end if

                                                        
                                                

                                                
                                                    end if


								                    if maxl <> 0 then 'ugeaflsuttet
								                    felt = felt &"&nbsp;<a class=""a_showkom2"" id='anc_"&job_iRowLoops_cn&"_"&x&"' name='"& job_jid &"'>"& kommtrue &"</a>" ' href=""#anc_s0""  name onClick=""expandkomm('"&job_iRowLoops_cn&"', '"&x&"', '"& job_jid &"')""
								                    end if
								
								                    felt = felt &"<input type='hidden' name='FM_timer' id='FM_timer' value='xx'><input type='hidden' name='FM_dager' id='FM_dager' value='xx'><input type='hidden' name='FM_datoer' id='FM_datoer' value='"& tjekdag(x) &"'>"_
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
                                                    
                                                    felt = felt &"<input class='dcls_"&x&"' type='"& timtype &"' name='FM_timer' id='FM_timer_"&job_iRowLoops_cn&"_"&x&"' value='"& timerThis &"' maxlength='5' style=""border:1px #999999 solid; height:16px; width:45px; font-family:arial; font-size:10px;"";><input type='hidden' name='FM_dager' id='FM_dager' value='xx'>"
								                    
                                                    felt = felt &"<input type='hidden' name='FM_sttid' id='FM_sttid' value='"& timer_sttid &"'><input type='hidden' name='FM_sltid' id='FM_sltid' value='"& timer_sltid &"'>"
								

                                                    felt = felt &"<input type='hidden' name='FM_timer' id='FM_timer' value='xx'><input type='hidden' name='FM_dager' id='FM_dager' value='xx'><input type='hidden' name='FM_datoer' id='FM_datoer' value='"& tjekdag(x) &"'>"_
                                                    &"<input class='d_kom_cls_"&x&"' type='hidden' name='FM_kom_"&job_iRowLoops_cn&""&x&"' id='FM_kom_"&job_iRowLoops_cn&""&x&"' value='"&timerKomThis&"'>"_
                                                    &"<input type='hidden' name='FM_feltnr' id='FM_feltnr' value='"&job_iRowLoops_cn&""&x&"'>"

                                                    felt = felt &"</td>"

                                                    end if
								
								
							    
							                      
						       
						                    %>
						                   
						
						                 
				                    <%
				                    ugedagUdskrevet = ugedagUdskrevet & "," & weekday(tjekdag(x), 2) & "#"
				                    'Response.write ugedagUdskrevet
				                    %>
				
				                    <%'end if%>
			                    <%next 'x 1 - 7

                                
								                    

                strAktiviteter = strAktiviteter & felt & "</tr>"


            '*** akt end
            aktidsMtimWrt = aktidsMtimWrt &",#"&job_aid&"#"
            lastKnavn = left(job_kundenavn,1)    
            last_job_jid = job_jid
            lastFase = lcase(job_fase)
            job_iRowLoops_cn = job_iRowLoops_cn + 1
            acn = acn + 1

            if acnForFaseIsWrt = 0 then
            acnForFase = acnForFase + 1
            else
            acnForFase = acnForFase
            end if

            acnDtlinje = acnDtlinje + 1
            oRec6.movenext
            wend
            oRec6.close
            'next
            

            next 'mn


             'if cint(visallemedarb_bl) <> 1 then

                    '***************** Notificer hvis aktibudget er opbrugt ****************'
                    'strAktiviteter = strAktiviteter & "<tr><td bgcolor='#FFFFFF' colspan='9' style='padding:0px 0px 0px 0px;' id='aktnotificer_"& job_jid &"'>&nbsp;</td></tr>"

                    '***************** Notificer hvis forecast er opbrugt ****************'
                    'strAktiviteter = strAktiviteter & "<tr><td bgcolor='#FFFFFF' colspan='9' style='padding:0px 0px 0px 0px;' id='aktnotificer_fc_"& job_jid &"'>&nbsp;</td></tr>"

                     '***************** Notificer hvis maks timer på rejsedage overskreddet ****************'
                     strAktiviteter = strAktiviteter & "<tr><td bgcolor='#FFFFFF' colspan='9' style='padding:0px 0px 0px 0px;' id='aktnotificer_trvl_"& job_jid &"'>&nbsp;</td></tr>"

                    '***Submit i bubnd / tom liste besked
                    if (fo = 0 OR acn = 0) AND (sortByVal <> 5 AND sortByVal <> 6) then
                    strAktiviteter = strAktiviteter & "<tr><td bgcolor='#FFFFFF' colspan='9' valign='top' style='padding:20px 0px 20px 20px;'>"_
                    &"Der blev ikke fundet nogen <b>aktiviteter</b> p&aring; dette job i den valgte periode.<br><br> - Tjek at du har <b>rettigheder til aktiviteterne</b>, og at de ligger i det korrekte <b>tidsinterval</b> (startdato nyere end mandag i valgte uge).<br> - Hvis jobbet har status som et tilbud vises kun <b>salgs-aktiviteter.</b>"
                                    
                     if cint(aktBudgettjkOn) = 1 then
                     strAktiviteter = strAktiviteter & "<br> - Der <b>findes ikke forecast / timebudget</b> p&aring; aktiviteterne endnu. (og forecast-filter er sl&aring;et til)"
                     end if               
                    
                                    
                     strAktiviteter = strAktiviteter & "</td></tr>"


                    else

                    ''** Lukket TXT ***'
                    if len(trim(ugeafsluttetTxt)) <> 0 then
                    strAktiviteter = strAktiviteter & "<tr><td bgcolor='#FFFFFF' colspan='9' style='padding:20px 0px 5px 20px;'>"& ugeafsluttetTxt &"</td></tr>"
                    end if

            
                    '*** ÆØÅ **'
                    call jq_format(tsa_txt_085)
                    tsa_txt_085jq = jq_formatTxt

                    strAktiviteter = strAktiviteter & "<tr><td bgcolor='#FFFFFF' colspan='9' align='right' valign='top' style='padding:10px 0px 20px 10px; height:10px; border-bottom:0px #cccccc solid;'>"
                    strAktiviteter = strAktiviteter & "<input type='hidden' id='jq_antalakt_"&job_jid&"' value='"&acnForFase&"'>"
                    strAktiviteter = strAktiviteter & "<input type='hidden' id='jq_antalakt_0' value="&acn&">"
                    strAktiviteter = strAktiviteter & "<input type=""submit"" value='"& tsa_txt_085jq &" >>' >"_
                    &"</td></tr>"
                    end if

            
            'end if
            
          



                '****************************************************************************************************
                '***** Blandet liste, Tilføj Options, tilføj timer på ***'
                '****************************************************************************************************

               select case lto 
               case "oko", "xintranet - local"
                                    
               case else 
                

               if cint(visallemedarb_bl) <> 1 then


                if (cint(sortByVal) = 5 OR cint(sortByVal) = 6) AND cint(intEasyreg) <> 1 then 
                '** tilføj timer på **
                strAktiviteter = strAktiviteter &"<tr bgcolor=#FFFFFF><td colspan=9><br><br><br><b>Tilf&oslash;j aktiviteter til liste:</b> (v&aelig;lg aktiviteter fra personlig aktiv-jobliste)</td></tr>"

                 
                strSQLaktopt = "SELECT a.id AS aid, j.id AS jid, navn AS aktnavn, a.fase, k.kkundenavn, k.kkundenr, k.kid, j.jobnavn, j.jobnr, a.fakturerbar, j.jobstatus FROM aktiviteter AS a "_
                &" LEFT JOIN job AS j ON (j.id = a.job) "_
                &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "
                                    
                ordBySQL =  "k.kkundenavn, j.jobnavn, a.fase, a.sortorder, a.navn"

                                       
               lmt = "0,500"
               
               select case lto
               case "wwf", "wwf2"
               wwfvisalleinterne = 1
               case else
               wwfvisalleinterne = 0
               end select
               
               if (wwfvisalleinterne = 1) then
               wwfvisalleinterneSQLkri = " OR (j.jobnr = 'WWF04')" 'Vis internet dog IKKE fravær //  eller - 2 undt. Ferie + ferie fri optjent AND (a.fakturerbar <> 12 AND a.fakturerbar <> 15 )
               else
               wwfvisalleinterneSQLkri = ""
               end if

  
               aktidsSQLOptions = replace(aktidsSQLOptions, ")", "")
               strSQLaktopt = strSQLaktopt & aktidsSQLOptions &")" & strSQLAktDatoKri & " "& strSQLAktStatusKri &" "& wwfvisalleinterneSQLkri &" AND (j.jobstatus = 1 OR j.jobstatus = 3) GROUP BY a.id ORDER BY " & ordBySQL & " LIMIT " & lmt
               
               'if session("mid") = 1 then
               'Response.write strSQLaktopt '"##"& strSQLaktopt
               'Response.flush
               'end if
               
               oplastJid = 0
               oplastKid = 0
               aktOptions = ""
               oRec6.open strSQLaktopt, oConn, 3                     
               while not oRec6.EOF 

               if isNull(oRec6("fase")) = true then
               fs = ""
               else
               fs = left(oRec6("fase"), 30) & ":"
               end if

               
               'if oplastKid = 0 OR oplastKid <> oRec6("kid") then
               'aktOptions = aktOptions & "<option value="& oRec6("aid") &" disabled></option>" 
               'aktOptions = aktOptions & "<option value="& oRec6("aid") &" disabled>"& left(oRec6("kkundenavn"), 30) &"</option>" 
               
               'else

               if oplastJid = 0 OR oplastJid <> oRec6("jid") then
               aktOptions = aktOptions & "<option value="& oRec6("aid") &" disabled></option>"
               end if

               'end if

               '**Vises ikke på tilføjlisten hvis de allerede er på timreg. listen
               if instr(aktidsMtimWrt, "#"&oRec6("aid")&"#,") <> 0 then
               dakAkt = "disabled"
               else
                    if oRec6("jobstatus") = 3 AND ignorerakttype <> "1" then 'tilbud
                        if oRec6("fakturerbar") = 6 then 'kun salgs akt
                         dakAkt = ""
                         aktOptions = aktOptions & "<option value="& oRec6("aid") &" "& dakAkt &">"& left(oRec6("kkundenavn"), 30) &" | "& left(oRec6("jobnavn"), 50) &" ("& oRec6("jobnr") &") | "& fs &" "& left(oRec6("aktnavn"), 50) &"</option>" 
                  
                        end if
                   else
                   dakAkt = ""
                   aktOptions = aktOptions & "<option value="& oRec6("aid") &" "& dakAkt &">"& left(oRec6("kkundenavn"), 30) &" | "& left(oRec6("jobnavn"), 50) &" ("& oRec6("jobnr") &") | "& fs &" "& left(oRec6("aktnavn"), 50) &"</option>" 
                   end if
               end if

              


               oplastJid = oRec6("jid")
               oplastKid = oRec6("kid")

               oRec6.movenext
               wend
               oRec6.close


                




                '*** ÆØÅ **'
                call jq_format(aktOptions)
                aktOptions = jq_formatTxt

               strAktiviteter = strAktiviteter &"<tr><td colspan=9 bgcolor=""#FFFFFF""><select id=""nyaktivitet"" size=""10"" multiple style=""width:650px; font-size:11px;""><option value=0>Kontakt | Job | Fase: Aktivitet</option>"& aktOptions &"</select><br>Valgte aktiviteter vil komme med p&aring; listen ovenfor. Der g&aring;r ca. 2-3 sek.</td></tr>"
               strAktiviteter = strAktiviteter &"<tr><td align=right colspan=9 bgcolor=""#FFFFFF"">&nbsp;<input id=""nyaktivitet_but"" value=""Tilf&oslash;j >>"" type=""button""><br><br><br>&nbsp;</td></tr>"
                end if
         
        
                end if   'if cint(visallemedarb_bl) <> 1 then      
           
               end select


    
           strAktiviteter = strAktiviteter & "</table><br>&nbsp;"
           
           Response.Write strAktiviteter 


           'call akttyper2009(2)
            

           Response.end
           
           
           
           


        

                
        case "FN_jobidenneuge"
          

         'Response.end

         stDato = request("stDato")
         slDato = request("slDato")

         'stDato = year(stDato) & "/"& month(stDato) & "/"& day(stDato)
         'slDato = year(slDato) & "/"& month(slDato) & "/"& day(slDato)

         jobans = request("jobans")
         'medid = request("medid")

         lto = request("lto")

         'select case lcase(lto)
         'case "dencker", "intranet - local"
         'Response.write "<br><br>Dencker V&aelig;rkt&oslash;j:<br>Oversigten er midlertidigt sl&aring;et fra pga. performance test."
         'Response.end
         'case else
         'end select


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
         &" WHERE jobstatus = 1 AND risiko > - 1 "& jobdtSQL &" "& jobansSQL &" GROUP BY j.id ORDER BY " & orderbySQL & " LIMIT 0,"& limit
         
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

         if (lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "xintranet - local") OR (level = 1 OR (oRec3("jobans1") = session("mid") OR oRec3("jobans2") = session("mid"))) then

         strJoiDUge = strJoiDUge &"<tr><td class=lille style=""border-bottom:1px #CCCCCC solid;""><a href='webblik_joblisten.asp?nomenu=1&rdir=treg&FM_kunde="&oRec3("kid")&"&jobnr_sog="&oRec3("jobnr")&"&FM_medarb_jobans="&session("mid")&"&st_sl_dato=2' class=rmenu target=""_blank"">"& left(oRec3("jobnavn"), 50) &" ("& oRec3("jobnr") &")</a>"_
         &"&nbsp;<a href=""jobs.asp?menu=job&func=red&id="& oRec3("jid") &"&int=1&rdir=treg"" style=""font-size:9px; color:#6CAE1C;"">"& left(tsa_txt_251, 3) &".</a><br>"_
         &"<span style=""font-size:9px; color:#999999;"">"& left(oRec3("kkundenavn"), 20) &"</span></td>"_
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

          strJoiDUge = strJoiDUge &"<tr><td class=lille style=""border-bottom:1px #CCCCCC solid;"">"& left(oRec3("jobnavn"), 50) &" ("& oRec3("jobnr") &")</a><br>"_
         &"<span style=""font-size:9px; color:#999999;"">"& left(oRec3("kkundenavn"), 20) &"</span></td>"_
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

        
              '<script src="inc/xtimereg_2006_load_func.js"></script> 
              

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
         call  medarbiprojgrp(selPrgrp, selmedarbid, 0, -1)
            
           
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
       
       
         Response.Write strTildelJob & "</select><input id=""tildel_sel_medarb_bt"" type=""button"" value="" >> "" style=""font-size:9px;""><br><br>"
         Response.end


        case "FN_getTildelJob"

                if len(trim(request.form("prgrp"))) <> 0 then
                selPrgrp = request.form("prgrp")
                else
                selPrgrp = 0
                end if


              selmedarbid = request("vlgtmedarb")
             
                

         
         '<script src="inc/xtimereg_2006_load_func.js"></script> 
             

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
                
        
        
        
        
        
        '*************************************************************************
        '**** Kørsel ***'
        '*************************************************************************
        case "FM_get_destinations"

                                    


            	if len(trim(request("kid"))) <> 0 then
				kid = request("kid") 
				else
				kid = 0
				end if
				
				
				
				
				    if len(trim(request("visalle"))) <> 0 then
				    visalle = request("visalle")
				    else
				    visalle = 0
				    end if
				    
				    'if visalle <> 0 then
				    '    strSQLKundKri = " AND kid <> 0"
				    'else
				    '    strSQLKundKri = " AND kid = " & kid
				    'end if
				
				
                                    if len(trim(request("jqmid"))) <> 0 then
                                    jqmrn = request("jqmid")
                                    else
                                    jqmrn = 0
                                    end if
				
				
				'if len(trim(request.form("cust"))) <> 0 then
				'sogeKri = " (kp.navn LIKE '"& request.form("cust") &"%' OR k.kkundenavn LIKE '"& request.form("cust") &"%')"
				'else
				'sogeKri = " kkundenavn <> 'sdfsdfd556irp'"
				'end if
				 
                                    strFil_Kon_Txt = ""

                            '**** Valgt medarbejer *****'
                    strSQLmnradr = "SELECT mnavn, mnr, init, madr, mpostnr, mcity, mland FROM medarbejdere WHERE mid = "& jqmrn 
                    'Response.write "<option>"& strSQLmnradr &"</option>"
                    
                                    
                     oRec2.open strSQLmnradr, oConn, 3
                    if not oRec2.EOF then
                    
                   
                    
                                strFil_Kon_Txt = strFil_Kon_Txt & "<option value='0'>"& oRec2("mnavn") &" "_
                                & left(oRec2("madr"), 20) &" "_
                                & oRec2("mpostnr") &" "& oRec2("mcity") & " "& oRec2("mland")
                                strFil_Kon_Txt = strFil_Kon_Txt &"</option>"
                           

                    end if
                    oRec2.close
                                   


                        strSQLklicens = "SELECT k.kid, k.kkundenavn, k.adresse, k.postnr, k.city, k.land FROM kunder AS k WHERE useasfak = 1"
                        oRec.open strSQLklicens, oConn, 3
                        if not oRec.EOF then
                        
                                if len(trim(oRec("kkundenavn"))) <> 0 then
                                
                                strFil_Kon_Txt = strFil_Kon_Txt & "<option value='k_"& oRec("kid") &"'>"& left(oRec("kkundenavn"), 50) &" "_
                                & left(oRec("adresse"), 20) &" "_
                                & oRec("postnr") &" "& oRec("city") &" "_
                                & oRec("land")
                                strFil_Kon_Txt = strFil_Kon_Txt & "</option>"
                                end if

                        end if
                        oRec.close
                      
                                 
                 
                          
                strSQLkperskm = "SELECT kp.id AS kpid, kp.navn AS kpersnavn, kp.email, kp.dirtlf, kp.mobiltlf, "_
				&" kp.adresse As kpadr, kp.postnr As kpzip, kp.town As kptown, kp.land As kpland, kp.afdeling AS kpafd, "_
				&" k.kid, k.kkundenavn, k.kkundenr, k.adresse, k.postnr, k.city, k.telefon, k.land FROM kunder k "_ 
				&" LEFT JOIN kontaktpers kp ON (kp.kundeid = k.kid AND kp.navn IS NOT NULL AND kp.adresse IS NOT NULL AND kp.postnr IS NOT NULL AND kp.town IS NOT NULL)"_
				&" WHERE useasfak <> 1 "_
                &" AND k.kkundenavn IS NOT NULL AND k.adresse IS NOT NULL ORDER BY k.kkundenavn, kp.navn"
			   
			    'Response.Write "<option>"& strSQLkperskm & "</option>"
			    'Response.end
			   
			
                lastKid = 0
                lastAlfabet = ""
			   
			    oRec2.open strSQLkperskm, oConn, 3
                x = xvalbegin
                k = 0
                while not oRec2.EOF
                
                   'strFil_Kon_Txt = strFil_Kon_Txt & "<option>"& oRec2("kkundenavn") &" </option>"
                       
                                 if lastAlfabet <> left(oRec2("kkundenavn"), 1) then
                                strFil_Kon_Txt = strFil_Kon_Txt & "<option DISABLED></option>"
                                strFil_Kon_Txt = strFil_Kon_Txt & "<option DISABLED></option>"
                            strFil_Kon_Txt = strFil_Kon_Txt & "<option DISABLED style=""font-size:20px; background-color:#F7F7F7;"">"& left(ucase(oRec2("kkundenavn")), 1) &"</option>"
                        end if
                              

                                    if len(trim(oRec2("kkundenavn"))) <> 0 AND lastKid <> oRec2("kid") then
                                strFil_Kon_Txt = strFil_Kon_Txt & "<option DISABLED></option>"
                                strFil_Kon_Txt = strFil_Kon_Txt & "<option value='k_"& oRec2("kid") &"'>"& oRec2("kkundenavn") &" "_
                                & left(oRec2("adresse"), 20) &" "_
                                & oRec2("postnr") &" "& oRec2("city")&" "_
                                & oRec2("land")
                                strFil_Kon_Txt = strFil_Kon_Txt &  "</option>"
                                end if
                                


                                       x = x + 1   
                        k = 1 
                       
                        
                
                               if len(trim(oRec2("kpid"))) <> 0 then
                                
                                               
                                        if len(trim(oRec2("kpersnavn"))) <> 0 then
                                        strFil_Kon_Txt = strFil_Kon_Txt & "<option value='kp_"& oRec2("kpid") &"'>.........."& oRec2("kpersnavn") &" "
                                       
                                        if len(trim(oRec2("kpafd"))) <> 0 then
                                        strFil_Kon_Txt = strFil_Kon_Txt &  "Afd:"&  oRec2("kpafd") &" "
                                        end if
                                        
                                        strFil_Kon_Txt = strFil_Kon_Txt & left(oRec2("kpadr"), 20) &" "_
                                        & oRec2("kpzip") &" "& oRec2("kptown") &" "& oRec2("kpland")
                                       
                                        
                                        strFil_Kon_Txt = strFil_Kon_Txt &  "</option>"
                                        end if
                                        

                                  x = x + 1
                                end if 'oRec2("id") <> 0    
                
                              
                lastAlfabet = left(oRec2("kkundenavn"), 1)
                lastknavn = oRec2("kkundenavn") 
                lastKid = oRec2("kid")
                oRec2.movenext
                wend
                oRec2.close

                                    'Response.Write strFil_Kon_Txt
                                    'Response.end

                           
                                    
                                   
                     
                
                
                '*** ÆØÅ **'
                call jq_format(strFil_Kon_Txt)
                strFil_Kon_Txt = jq_formatTxt
                Response.write strFil_Kon_Txt  
                                    Response.end
                'Response.Write strFil_Kon_Txt & "<input id=""kperfil_fundet"" type=""hidden"" value='"&x-1&"'>" 
                '&"</div><input type=""text"" value='"&x&"'>"
                
                 
                 
                 
        '*******************************************************************************            
        '**********         JOBBANKEN                                   ****************            
        '********* Henter valgte job i søge filter iforhold til søgning ****************
        '*******************************************************************************            
        '*******************************************************************************            
        case "FN_getCustDesc"
        
          
      
             'if request("FM_hentjava") = "1" then
            '<script language="javascript" src="inc/timereg_2011_load_func.js"></script>
            'end if
           
     
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


        '**** VIS HR ***'
        if len(trim(request("vishr"))) <> 0 then
        vishr = request("vishr")
        else
        vishr = 0
        end if

        if vishr <> 0 then
        Response.Write "Der er valgt HR liste. <br>Alle interne job, sygdom, ferie, interntid etc. (prioitet = -2)<br><br>Tilf&oslash;j job til HR listen ved at redigere job --> Avanceret --> S&aelig;t prioitet = -2"
        Response.end

        end if 
        

        '**jobansvarlig *****'
        if request("jobansv") = "1" then
        jobansvSQLkri = " AND (jobans1 = "& usemrn &" OR jobans2 = "& usemrn &" OR jobans3 = "& usemrn &" OR jobans4 = "& usemrn &" OR jobans5 = "& usemrn &")"
        else
        jobansvSQLkri = ""
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

            jobnrSmartRegKri = jobnrSmartRegKri & " OR j.jobnr = '"& oRec3("tjobnr") &"'"


                    
                    'Response.Write visGuide
                    'Response.end



                    if cint(visGuide) = 11 then 
                    dtNow = year(now) & "-"& month(now) & "-"& day(now)

                    strSQLuseguide = "SELECT id FROM job WHERE jobnr = '"& oRec3("tjobnr") &"'"
                    'Response.Write strSQLuseguide

                    oRec6.open strSQLuseguide, oConn, 3
                    if not oRec6.EOF then

                            forvalgt = 0
                            strSQLTUfindes = "SELECT forvalgt FROM timereg_usejob WHERE jobid = "& oRec6("id") &" AND medarb = "& usemrn
                            oRec5.open strSQLTUfindes, oConn, 3
                            if not oRec5.EOF then

                            forvalgt = oRec5("forvalgt") 

                            end if
                            oRec5.close

                            if cint(forvalgt) = 0 then
                            strSQLupdateTuseJob = "UPDATE timereg_usejob SET forvalgt = 1, forvalgt_sortorder = 0, forvalgt_af = "& session("mid") &", forvalgt_dt = '"& dtNow &"' WHERE jobid = "& oRec6("id") &" AND medarb = "& usemrn
                            'Response.write strSQLupdateTuseJob 
                            oConn.execute(strSQLupdateTuseJob)
                            end if

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

        
       


          sortby = request("sortby")
        'if sortby = 1 then
        'sqlOrderBy = sqlOrderBy & "j.jobnavn "
        'end if
        
        if sortby = 2 then
        sqlOrderBy = sqlOrderBy & "j.jobstartdato DESC, j.jobnr DESC "
        end if
        
        'if sortby = 3 then
        'sqlOrderBy = sqlOrderBy & "k.kkundenavn, j.jobnr DESC "
        'end if
        
        if sortby = 0 then
        sqlOrderBy = sqlOrderBy & "k.kkundenavn, j.jobnavn "
        end if
       
        
        if cint(visnyeste) = 1 then
        'sqlOrderBy = "j.jobstartdato DESC, j.jobnavn"
        lmt = " LIMIT 0, 25"
        dtWhcls = " AND j.jobstartdato <= '"& year(now) &"/"& month(now) &"/"& day(now) &"' "
        
        
        else
            
     
        
        if cint(viseasyreg) = 1 then
        
        lmt = ""
        dtWhcls = ""
        else
        
        select case lcase(lto)
        case "hvk_bbb"
        lmt = " LIMIT 0, 200"
        case "dencker"
        lmt = " LIMIT 0, 100"
        case "mi", "intranet - local"
        lmt = " LIMIT 0, 200"
        case "kejd_pb"
        lmt = " LIMIT 0, 100"
        case "jttek"
        lmt = " LIMIT 0, 200"
        case else
        lmt = " LIMIT 0, 100"
        end select
        'lmt = ""
        dtWhcls = ""
        
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
                                    
                                    if instr(jobsog, ",") <> 0 then
                                    
                                    jobsogArr = split(jobsog, ",")
                                    

                                   

                                    for j = 0 TO UBOUND(jobsogArr)

                                    jobSogTxt = trim(jobsogArr(j)) 

                                    if j = 0 then
                                      strJobSogKri = " AND ((kkundenavn LIKE '%"& jobSogTxt &"%' OR kkundenr = '"& jobSogTxt &"' OR Kinit = '"& jobkundesog &"') "
                                    else
                                        if j = 1 then
                                        strJobSogKri = strJobSogKri & " AND ((jobnr LIKE '"& jobSogTxt &"%' OR jobnavn LIKE '%"& jobSogTxt &"%') "
                                        else
                                        strJobSogKri = strJobSogKri & " OR (jobnr LIKE '"& jobSogTxt &"%' OR jobnavn LIKE '%"& jobSogTxt &"%') "
                                        end if
                                    end if

                                    next


                                    strJobSogKri = strJobSogKri & "))"
                                    
                                    else

                                        if instr(jobsog, ">") <> 0 then
                                        jobSogTxt = trim(replace(jobsog, ">", "")) 
                                         call erDetInt(SQLBless(trim(jobSogTxt)))

                                            if len(trim(jobSogTxt)) > 0 AND isInt = 0 then
                                            strJobSogKri = " AND (jobnr > "& jobSogTxt &")"
                                            else
                                              strJobSogKri = " AND (jobnr < 0)"
                                            end if 
                                        else

                                         if instr(jobsog, "<") <> 0 then
                                         jobSogTxt = trim(replace(jobsog, "<", ""))
                                            
                                            call erDetInt(SQLBless(trim(jobSogTxt)))
				                           
                                      
                                            if len(trim(jobSogTxt)) > 0 AND isInt = 0 then
                                            strJobSogKri = " AND (jobnr < "& jobSogTxt &")"
                                            else
                                            strJobSogKri = " AND (jobnr < 0)"
                                            end if 
                                         else	
                                    	
		                                strJobSogKri = " AND (jobnr LIKE '"& jobsog &"%' OR jobnavn LIKE '%"& jobsog &"%' OR kkundenavn LIKE '%"& jobsog &"%' OR kkundenr = '"& jobsog &"' OR Kinit = '"& jobkundesog &"') "
                                    
                                        end if
                                    end if
                                    end if
		else
		strJobSogKri = ""
		end if				
        
        'Response.write strJobSogKri
        'Response.flush

        
        if cint(ignProj) = 0 then
        strSQLwh = " (j.jobstatus = 1 OR j.jobstatus = 3) "& strPgrpSQLkri & "" 
	    else
	    '** Henter alle aktive job + tilbud **'
	    strSQLwh = " (j.jobstatus = 1 OR j.jobstatus = 3) " 
	    end if
        
        if Request.Form("cust") <> 0 then 'AND strJobSogKri = "" then
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
            & strSQLwh &" "& strJobSogKri &" "& dtWhcls &" AND kkundenavn <> '' "& jobansvSQLkri &" GROUP BY j.id ORDER BY "& sqlOrderBy & lmt
            
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
            
            strSQL = strSQL & "WHERE "& strSQLwh &" "& strJobSogKri &" "& dtWhcls &" AND kkundenavn <> '' "& jobansvSQLkri &" GROUP BY j.id ORDER BY " & sqlOrderBy & lmt
            
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
        
        'Response.Write strSQL & " <br>seljobid: " & seljobid
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
         
                                       
          strJobIkkeFundet = "<h4 style=""color:#999999;"">"& tsa_txt_391 &"</h4>"
          strJobIkkeFundet =  strJobIkkeFundet & "<b><< "& tsa_txt_392 &"</b><br>" 
          strJobIkkeFundet = strJobIkkeFundet &  tsa_txt_393 &"<br><br>" 

            '*** ÆØÅ **'
            call jq_format(strJobIkkeFundet)
            strJobIkkeFundet = jq_formatTxt
                                     
            response.write strJobIkkeFundet
          
        end if
        
       
        
        'Response.Write "</select>"
        '<input id='jobids_easyreg' type='text' value='"& jobids_easyreg &"' />"
        'Response.Write  "<option value='0'>"& strSQL & "</option>"
        end select
        Response.end
        end if 'JQ AJAX Slut 
	
	


    'loadA =  now
  

	if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else







    '*************** Faste vaiable ***********************
    if len(trim(request("rdir"))) <> 0 then
    rdir = request("rdir")
    else
    rdir = ""
    end if
    
    select case rdir
    case "timetag_web" 'timeout mobile
    origin = 12
    case "ugeseddel_2011" 'timetag
    origin = 11
    case else
    origin = 0
    end select



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
	
	
    select case lto
    case "dencker", "intranet - local", "mmmi", "unik"
    hideallbut_first = 3 'viser alle akt. åbne ved laod
    case else
	hideallbut_first = 0
    end select



     call aktBudgettjkOn_fn()
     if cint(aktBudgettjkOn) = 1 then 
            
     
            
        
    

                        if len(trim(request("isindstillingersubmitted"))) <> 0 then '** Er der søgt
                           
                            if request("viskunejobmbudget") = "1" then
                            viskunejobmbudgetCHK = "CHECKED"
                            viskunejobmbudget = 1
                            else
                            viskunejobmbudgetCHK = ""
                            viskunejobmbudget = 0
                            end if

                        else

                          

                               if cint(aktBudgettjkOnViskunmbgt) = 1 then 'AND cint(level) <> 1
                               viskunejobmbudgetCHK = "CHECKED"
                               viskunejobmbudget = 1
                               
                               else

                                  
                    
                                    'if request.cookies("tsa")("viskunejobmbudget") = "1" then 'AND request.cookies("tsa")("viskunejobmbudget") <> "" then
                                    'viskunejobmbudgetCHK = "CHECKED"
                                    'viskunejobmbudget = 1
                                    'else
                                    viskunejobmbudgetCHK = ""
                                    viskunejobmbudget = 0
                                    'end if

                                 end if

                        end if

        if cint(level) <> 1 then 
           bxbgtfilter_Type = "hidden"
        else 
           bxbgtfilter_Type = "checkbox"
        end if

                        response.cookies("tsa")("viskunejobmbudget") = viskunejobmbudget

    else


    end if

        

    if len(trim(request("sortby"))) <> 0 then
    sortBy = request("sortby")
    response.cookies("tsa")("sortby") = sortby
    else
        if request.cookies("tsa")("sortby") <> "" then
        sortBy = request.cookies("tsa")("sortby")
           
       
        else

        select case lto 
        case "oko"
        sortBy = 5
        case else
        sortBy = 4
        end select

        end if
    end if

   select case sortBy
   case 1
   sortBySQL = "j.risiko, j.jobnavn," 'tu.forvalgt_sortorder
   case 2,5,6
   sortBySQL = "j.jobnavn,"
   case 21
   sortBySQL = "j.id DESC,"
   case 3
   sortBySQL = "j.jobslutdato DESC,"
   case 4
   sortBySQL = "k.kkundenavn, j.jobnavn,"
   case else
   sortBySQL = "k.kkundenavn, j.jobnavn,"
   'sortBySQL = "tu.forvalgt_sortorder, j.jobnavn,"
   end select



        'Response.write "request(visallemedarb_blandetliste): " & request("visallemedarb_blandetliste") & "<br>"

        if len(trim(request("visallemedarb_blandetliste"))) <> 0 AND request("visallemedarb_blandetliste") <> "0" then
        visallemedarb_blandetliste = request("visallemedarb_blandetliste")

            if visallemedarb_blandetliste = 1 then
            visallemedarb_blandetlisteCHK1 = "CHECKED"
            else
            visallemedarb_blandetlisteCHK2 = "CHECKED"
            end if

        else
        visallemedarb_blandetliste = 0
        visallemedarb_blandetlisteCHK1 = ""
        visallemedarb_blandetlisteCHK2 = ""
        end if


     
	
	call ersmileyaktiv()
	'**************************************************************************************************************************************
	'**************************************************************************************************************************************
	'*************** Opdaterer timeregistreringer *********
	'**************************************************************************************************************************************
    '**************************************************************************************************************************************

    

    '*** Origin values *'''

        '0: Alm. timreg eller via Epi gl. tricTrac
        '1: Vietnam import
        '3: IBM SPSS
        '10 Excel upload
        '11: TimeTag  / Ugeseddel 
        '12: TimeOut Mobile
        'xx: Tilret fra ugeseddel / afvigelsese håndtering == Beholder origin



	if func = "db" then


    if len(trim(request("extsysid"))) <> 0 then
    extsysid = request("extsysid")
    else
    extsysid = 0
    end if

    'Response.write "extsysid " & extsysid & "<br>"
    


    if len(trim(request.cookies("easyreg"))) <> 0 then
    intEasyreg = request.cookies("easyreg") 
    else
    intEasyreg = 0
    end if
  
	
	if len(trim(request("jobid"))) <> 0 then
	jobid = request("jobid")
	else
	jobid = 0
	end if
	

    'response.write "HEj"
    'response.end

    '** Skal Indlæst MSG vises på timetagsiden
    session("timetag_web_indMsgShown") = "0"
	

	
	if request("tildelalle") = "1" then
	multitildel = 1
	else
	multitildel = 0
	end if
	
	if len(trim(request("tildeliheledageSet"))) <> 0 AND multitildel = 1 then

	'*** MULTITILDEL -- SKAL ALTID OMREGNE INDTASTNINGER I DAGE TIL TIMER IFHT. DEN NORM TIDE DEN NEKELTE MEDARBEJDER HAR
    tildeliheledage = 1
    'response.cookies("tsa")("tildelidage") = "1"
	else
	tildeliheledage = 0
    'response.cookies("tsa")("tildelidage") = "0"
	end if
	

    
	
	'*** Valgte medarbejdere ****'
    dim tildelselmedarb
    'redim tildelselmedarb(200)
	
    if len(trim(request("tildelselmedarb"))) <> 0 AND cint(multitildel) = 1 then
    tildelselmedarb = split(request("tildelselmedarb"), ", ")
    else
    tildelselmedarb = split(1, ", ")
    end if

    'Response.write "tildelselmedarb: " & request("tildelselmedarb") & "<br>"
	
	
	            '*** => Timer er indtastet på timereg. siden ***'
	            '*** Eller via stopur                        ***'
	            if stopur <> "1" then 
	            
	                
	                '**** Obligatoriske felter ****            
                    'jobid
                    'aktid
                    'medarbejderid
                    'dato


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
	                
                    if len(trim(request("FM_medid"))) <> 0 then
	                medarbejderids = split(request("FM_medid"), ",") 'request("FM_mnr")
                    else
                    dim medarbejderids
                    redim medarbejderids(0)
                    medarbejderids(0) = "0"
                    end if
                    

                    'response.write "HER:"

                    'response.write "FM_dager, FM_feltnr, dato;:"& request("FM_datoer") &" <br>jobid: "& request("FM_jobid") & "<br>aktid: " & request("FM_aktivitetid") & "<br>medid: "& request("FM_medid") & "<br>timer "& request("FM_timer")

                    'response.end

                    'Response.write "multitildel: "& multitildel & "<br>"
                    'Response.write "visallemedarb_blandetliste: " & visallemedarb_blandetliste & "<br>"
                    'Response.write "medarbejderid : "&  medarbejderid & "<br>"
                    'Response.write "medarbejderids : "&  request("FM_medid") & "<br>"
                    'Response.write "sortBy: "&  sortBy & "<br>"
                    'Response.end

                	
                    'response.write request("FM_timer") & "<br><br>"

	                tTimertildelt_temp = replace(request("FM_timer"), ", xx,", ";")
	                tTimertildelt_temp2 = replace(tTimertildelt_temp, ", xx", "")
	                tTimertildelt = split(tTimertildelt_temp2, ";") 
                	
                	'Response.Write "tTimertildelt:"& tTimertildelt &" FM: " & request("FM_timer") & "<br>"
                	'Response.end

                    'dage = split(request("FM_dager"), ", ")
                    tDage_temp = replace(request("FM_dager"), ", xx,", ";")
	                tDage_temp2 = replace(tDage_temp, ", xx", "")
	                tDage = split(tDage_temp2, ";") 
                    ''Response.Write "tDage:" & request("FM_dager")
                	'Response.end
                	
	                datoer = split(request("FM_datoer"), ",")
	                
	                'Response.Write "datoer:" & request("FM_datoer")
                	'Response.end
                	
                    'response.write request("FM_sttid") & "<br>"

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
				               
	                
	
			
			
			'*** Validerer at job og akt er udfyldt når der indlæses fra ugeseddel (timeTag) eller timeOut mobile, 
            '**  hvor der er mulighed for at akt og job ikke er udfyldt ***
            if rdir = "timetag_web" OR rdir = "ugeseddel_2011" then
            if len(trim(request("FM_jobid"))) = 0 OR (request("FM_jobid") = 0 AND (origin = 12 OR origin = 11)) then
                
                    
                    antalErr = 1
					errortype = 167
					useleftdiv = "t"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end

            
            end if


            'response.write "FM_timer:" &  request("FM_timer")
            'response.end

             timertjk = replace(request("FM_timer"), ", xx", "")



             if (len(trim(timertjk)) = 0) AND (origin = 12 OR origin = 11) then
                'OR formatnumber(timertjk,2) = 0
                    
                    antalErr = 1
					errortype = 176
					useleftdiv = "t"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end

            
            end if

                        'response.write "origin" & origin
                        'response.end


             if len(trim(request("FM_aktivitetid"))) = 0 OR (request("FM_aktivitetid") = 0 AND (origin = 12 OR origin = 11)) then


                    antalErr = 1
					errortype = 168
					useleftdiv = "t"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end

            
            end if


             if (len(trim(request("FM_datoer"))) = 0 OR ( IsDate(request("FM_datoer")) = false )) AND (origin = 12 OR origin = 11) then


                    antalErr = 1
					errortype = 175
					useleftdiv = "t"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end

            
            end if



                   

            '*** Er periode lukket *****'
            call timerIndlaesPeriodeLukket(medarbejderids(0), request("FM_datoer"), jobids(0))


            end if

			antalErr = 0
			for y = 0 to UBOUND(tTimertildelt)
			
				
                        'response.write tTimertildelt(y) & "<br>"
                        'response.flush

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

                if tSttid(y) = "24:00" then
                    tSttid(y) = "23:59"
                end if
				
                if tSttid(y) = "00:00" then
                    tSttid(y) = "00:01"
                end if

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

                if tSltid(y) = "24:00" then
                    tSltid(y) = "23:59"
                end if
				
                if tSltid(y) = "00:00" then
                    tSltid(y) = "00:01"
                end if
				
				
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
	
	'*****************************************
	'**** Hvis indlæsnng sker fra Stopur  ***'
	'*****************************************
	
	
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
	useDage = 0 'split(dagethis, ",")
	
	end if '***Fra  Stopur****'
	
	
	
	
	
	'**** tjekker om timer skal afrundes
    call timerround_fn()
	
	
	
	'*********************************************************************************
	'***** Indlæser timer array ****'
	'***** Indlæser i db ***
    '*********************************************************************************
	
	
	'Response.Write "<br><br>antalErr" & antalErr & "<br>"
	
	
	if antalErr = 0 then
	
	
	
	for m = 0 to UBOUND(tildelselmedarb)
	
    EmailNotificer = 0 'skal der udesendes notifikation tikl teamleder
    EmailNotificerTxt = ""

    EmailNotificer2 = 0 'skal der sendes email fra leder til medarb. 
    EmailNotificerTxt2 = ""
	'Response.Write tildelselmedarb(m) & "<br>"
	
    if cint(visallemedarb_blandetliste) <> 1 AND cint(visallemedarb_blandetliste) <> 2 then 'alle linier er samme medarbejder

	if cint(multitildel) = 1 then '** Multitildel / kopier til alle valgte medarbejdere valgt, så man behøver kunne tjekke pris 1 gang == bedre load
	medarbejderid = tildelselmedarb(m) 
	else
        if stopur <> "1" then  
	    medarbejderid = medarbejderids(m) 'kun ved sortBy 5 eller 6 og VisforalleMedarb = 1 skal pris findes ved hver enkelt reg. --> VFlyt ned i J loop
        else
        medarbejderid = medarbejderid
        end if
	end if
	
	    
       '**** Finder medarbejder timepris og kostpris ************
		call mNavnogKostpris(medarbejderid)

    end if 
	    
        'Response.Write "her" & jobids(0)



        '**** FORDEL MATRIX, FORDEL PÅ FLERE AKTIVETER ALT EFTER TIDSPUNKT
       
        'tTimertildeltThis = 0
        select case lto
        case "nonstop", "intranet - local"
            'useTimeMatrix = 1
            matrixNo = 7
            'origin = 12 'Mobile

        case else
        matrixNo = 0
        end select


       


                           

	    for j = 0 to UBOUND(jobids)
        mtrx = 0



                            for mtrx = 0 TO matrixNo 'Gennemløber flere aktiviteter hvis timer skal fordeles på underliggende MATRIX, pga. tidspunkter
                            '***** Aktiviteter skal være SKJULT på timereg. siden
                            matrixAktFundet = 0
                            aktFakturerbarId = 0

                            select case mtrx
                            case 0
                            aktids(j) = aktids(j)
                            orgAktId = aktids(j) 
                            case 1
                            aktFakturerbarId = 50 'dag
                            case 2
                            aktFakturerbarId = 54 'aften  
                            case 3
                            aktFakturerbarId = 52 'weekend
                            case 4
                            aktFakturerbarId = 55 'weekend aften
                            case 5
                            aktFakturerbarId = 25 '1 maj timer (helligdage) 
                            case 6
                            aktFakturerbarId = 51 'nat
                            case 7
                            aktFakturerbarId = 53 'weekend nat
                            end select

                              if mtrx > 0 then

                                '** Henter navn på ORG aktivitet
                                    strAktOrgNavn = ""
                                    strSQLaktorg = "SELECT navn FROM aktiviteter WHERE job = "& jobids(j) & " AND id = "& orgAktId &""
                                    oRec6.open strSQLaktorg, oConn, 3
                                    if not oRec6.EOF then

                                        strAktOrgNavn = left(oRec6("navn"), 6)

                                    end if
                                    oRec6.close

                                    '*** DAG / NAT / WEEKEND MATRIX aktiviteter
                                    strSQLaktmtrx = "SELECT id FROM aktiviteter WHERE job = "& jobids(j) & " AND navn LIKE '"& strAktOrgNavn &"%' AND fakturerbar = " & aktFakturerbarId

                                    oRec6.open strSQLaktmtrx, oConn, 3
                                    if not oRec6.EOF then

                                        aktids(j) = oRec6("id")
                                        matrixAktFundet = 1

                                    end if
                                    oRec6.close

                              end if


                           
                   
                    if cint(mtrx) = 0 OR cint(matrixAktFundet) = 1 then  

                         '**** SKAL LÆSES VED HVERT LOOP. ELLERS FEJLER BLA: TARIF Beregning
                        'Tjekker alle linier for medarbejderId / Det er blandet liste med blandede medarbejdere
                        'Multitildel og stopur ikke mulig her.
                        if (cint(visallemedarb_blandetliste) = 1 OR cint(visallemedarb_blandetliste) = 2) OR cint(matrixAktFundet) = 1 then 
                        medarbejderid = medarbejderids(j) 'kun ved sortBy 5 eller 6 og VisforalleMedarb = 1 skal pris findes ved hver enkelt reg. --> Flyt ned i J loop
                     
                        '**** Finder medarbejder timepris og kostpris ************
		                call mNavnogKostpris(medarbejderid)

                        end if 

                  

                

			    
			    '*****************************************************'
				'*** Henter Akt og Job oplysninger                  **'
				'*** Finder evt. alternativ medarbejder timepris    **'
				'*** på den valgte aktivitet                        **'
				'*****************************************************'
				
                 'foundone_0 = 0
                 'intTimepris_0 = 0
                 'intValuta_0 = 0
                
				tfaktimvalue = 0
                strSQL = "SELECT job.id AS jobid, aktiviteter.navn, fakturerbar, jobnavn, jobnr, "_
				&" kkundenavn, jobknr, fastpris, job.serviceaft, kostpristarif"_
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

                             '*** Overskrifer kostpris med kostpristarif fra medab.type
                             kostpristarif = oRec("kostpristarif")
                             select case kostpristarif
                             case "0"
                             dblkostpris = dblkostpris
                             case "A"
                             dblkostpris = mkostpristarif_A
                             case "B"
                             dblkostpris = mkostpristarif_B
                             case "C"
                             dblkostpris = mkostpristarif_C
                             case "D"
                             dblkostpris = mkostpristarif_D
                             case else
                             dblkostpris = dblkostpris 
                             end select

							 
							 '**** 2009 t.tfaktim er = a.fakturerbar (akt. type) også i timer tabellen. **'
							 tfaktimvalue = strFakturerbart
							 strFastpris = oRec("fastpris")
						
					        
					        '*** tjekker for alternativ timepris på aktivitet
							call alttimepris(aktids(j), jobids(j), medarbejderid, 1)
							
							'** Er der alternativ timepris på jobbet
							if foundone = "n" then
								
                                call alttimepris(0, jobids(j), medarbejderid, 1)
                                
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
	
	        'if lto = "nonstop" OR lto = "intranet - local" then 
	        'response.write "strAktNavn: "& strAktNavn &" kostpristarif: " & kostpristarif &" kostprsi: "& dblkostpris &"<br>"
            'end if
	        
	        
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
			
            if rdir = "timetag_web" OR rdir = "ugeseddel_2011" then 'ugeseddel er opdater aktivieter eller tilføje herreløse fra TimeTag
			yslut = 0
            else
            yslut = (ystart + 6)
            end if

            'Response.Write high_fetnr & "<br>"
			'Response.Write "ystart: "& ystart & "<br>"
            'Response.Write "yslut: "& yslut & "<br>"
            'Response.end

			isAktMedidWrt = ""
			multiDagArrUse = 0
			
			'for y = 0 to UBOUND(tTimertildelt)
			for y = ystart to yslut 'AND y <= z  
			    
			    '*** tjekker om felt id liggen i den rigtige aktlinie *'
			    '**  DVS mellem 11 - 17, eller 21-27 etc..          ***'
				if y >= ystart AND y <= yslut then
				
                'Response.Write "y: "& y &"<br>"
                'Response.Write "feltnr(y):"& feltnr(y)
                'Response.flush
				feltvalThis = cint(replace(feltnr(y), "_", ""))
				feltvalThis = cstr(feltvalThis)
				
				
			    if cint(origin) = 12 then '12 TimeTag Mobile, 11: TimeTag / Ugeseddel, 10: excel 
                    
                    if request("FM_kom_"&feltvalThis) <> "Kommentar" then
				    strKomm = "; "& replace(request("FM_kom_"&feltvalThis), "'", "&#39;")
                    else
                    strKomm = ""
                    end if
                else
                strKomm = replace(request("FM_kom_"&feltvalThis), "'", "&#39;")
                end if
				
			    '*** Tjekker at komm. er udfyldt ved pre-def aktiviteter ***'
				
                'response.write "tfaktimvalue: "& tfaktimvalue
                'response.end                        
                call akttyper2009prop(tfaktimvalue)
				
				
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

                    'if (session("mid") = 1) AND (lto = "intranet - local" OR lto = "epi" OR lto = "outz") then
				    'Response.Write "<hr>visTimerelTid: "& visTimerelTid &"<br>fltnt:"& feltnr(y)&"<br>"
                     '        Response.flush 
                    'Response.Write "aktid: "& aktids(j)&"<br>"
                    '        Response.flush
                    'Response.Write " timer:" & tTimertildelt(y) &"<br>"
                    '        Response.flush
                    'Response.Write " Dato: " & datoer(y)&"<br>"
                    '        Response.flush 
                    'Response.Write " tSttid(y): "& tSttid(y) &" - "& tSltid(y) &"<br>"
                     '       Response.flush
                    'Response.Write " medarbejderid: "& medarbejderid&"<br>"
                    '        Response.flush
                    'Response.end
                    'end if
				
				if len(trim(tTimertildelt(y))) <> 0 then 

				    tTimertildelt(y) = SQLBless(tTimertildelt(y))

                else
					
                    if visTimerelTid <> 0 then 'Timer eller klokke
                    tTimertildelt(y) = -9002 '-2 Opdater '** ÆNDREET 03032016 da der ellers kan komme fejl ind indlæsninger FRA TT ved flere registreringer pr. dag. nårder bruges klokkeslet
                    else
					tTimertildelt(y) = -9001 '-1 Slet
				    end if

				end if
				
				
                   
				
			    
				    '**** * Multitildel for flere dage ****'
                    '**** Stjerne funktion er brugt  ******'
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
				    call normtimerPer(medarbejderid, useDato, 0, 0)
    				
				        'Response.Write "timer:" & useTimer & " Dato: " & useDato & " ntimPer:"& ntimPer &" medarbejderid: "& medarbejderid&"<br>"
                        'Response.flush
    				
				        if cint(ntimPer) <> 0 then
    				    useTimerThis = useTimer
    				    else
    				    useTimerThis = 0
    				    end if
    				    
                        
    				        'response.write "tSttid(y)"&  tSttid(y)
                            'response.flush 
            				
				            '**** Indlæser timer i DB, alm timereg. i timer *****'
				            call opdaterimer(aktids(j), strAktNavn, tfaktimvalue, strFastpris, intJobnr, strJobnavn, strJobknr,_
				            strJobknavn, medarbejderid, strMnavn,_
				            useDato, useTimerThis, strKomm, intTimepris,_
				            dblkostpris, offentlig, intServiceAft, strYear,_
				            tSttid(y), tSltid(y), visTimerelTid, stopur, intValuta, bopal, destination, 0, 0, origin, extsysid, mtrx)
            				
    				    
    				    
				        isAktMedidWrt = isAktMedidWrt & ",#"&aktids(j)&"_"&medarbejderid&"#"
				        'end if
    				    
				        findTimerForStar = ""
				        
    				 
				    next 
				    
    				
    				
    				
				    else
				    '*** almindelig indtastning 
    				        
    				        
				            if left(trim(tTimertildelt(y)), 1) = "." then
			                useTimer = "0"&trim(tTimertildelt(y))
			                else
			                useTimer = tTimertildelt(y)
			                end if    
    			            
    				        muDag = 0
				            useDato = datoer(y)

                            if rdir = "xtimetag_web" OR rdir = "ugeseddel_2011" then
                            useDage = ""
                            usetSltid = ""
                            usetSttid = ""
                            else
				            useDage = tDage(y)
                            usetSltid = tSltid(y)
                            usetSttid = tSttid(y)
				            end if  
                                 
                           
				          

				            '*** Er akt/medarb. allerede opdateret via multiDatoer?****' 
				            '** Så skal de ikke opdateresd igen, da de så overskriver den allerede opdaterede værdi ***'
				            if instr(isAktMedidWrt, ",#"&aktids(j)&"_"&medarbejderid&"#") = 0 then
				            
				           
    	                    '**** Indlæser timer i DB, alm timereg. i timer *****'
				            
				            if aktids(j) <> 0 AND len(trim(aktids(j))) <> 0 then
				            

                            'response.write "tSttid(y)"&  usetSttid
                            'response.flush 

                             'Response.Write "<br>"& tSttid(y)& "usetSttid : "&  usetSttid &" timer:" & useTimer & " Dato: " & useDato & " ntimPer:"& ntimPer &" medarbejderid: "& medarbejderid&"<br>"
                             'Response.flush

                               
                            call opdaterimer(aktids(j), strAktNavn, tfaktimvalue, strFastpris, intJobnr, strJobnavn, strJobknr,_
				            strJobknavn, medarbejderid, strMnavn,_
				            useDato, useTimer, strKomm, intTimepris,_
				            dblkostpris, offentlig, intServiceAft, strYear,_
				            usetSttid, usetSltid, visTimerelTid, stopur, intValuta, bopal, destination, useDage, tildeliheledage, origin, extsysid, mtrx)
    				        
    				        
    				        isAktMedidDatoWrt = isAktMedidDatoWrt & ",#"&aktids(j)&"_"&medarbejderid&"_"&useDato&"#"
				            end if
				            
				            
				            end if
    				
    				
				    end if
			   
                                'lb = ""
				                'ub = ""

				end if 'y <= ystart


            'useTimeMatrix = 0
            'tTimertildeltThis = 0
			
			next 'y = ugedagene 1-7
			




			else ' = fra Stopur
                        '********************************************************************************************************
			            '********************************************************************************************************
			        
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
                
                
                            call timerIndlaesPeriodeLukket(medarbejderid, regdato)
               
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
				            dblkostpris, offentlig, intServiceAft, strYear, tSttid, tSltid, visTimerelTid, stopur, intValuta, bopal, destination, 0, 0, origin, extsysid, mtrx)
 
			            '*** Opdaterer stopur timer overført *****'
			            strSQLstopU = "UPDATE stopur SET timereg_overfort = 1 WHERE id = "& entrysids(j) 
	                    oConn.execute(strSQLstopU)
			       
			            end if '**** Fra Stopur ***'
                        '********************************************************************************************************
			            '********************************************************************************************************
			

                    end if 'mtrx = matrixAktFundet = 0

		
	                next 'mtrx KLOKKE/TIME MATRIX


		next 'j

        


                    'response.write "EmailNotificer: "& EmailNotificer &" EmailNotificerTxt: " & EmailNotificerTxt


                if cint(EmailNotificer) = 1 then
                    
                    call notificerEmail(medarbejderid, EmailNotificerTxt, 1, session("mid"))
                    
                end if 


                 if cint(EmailNotificer2) = 1 then
                    
                    if medarbejderid <> session("mid") then
                        call notificerEmail(session("mid"), EmailNotificerTxt2, 2, medarbejderid)
                    end if
                    
                end if 


          
        next 'm (multideldel)
	
	    
	    '** Tjekker indlæs værdier **' 
	    'Response.end
	
	
	  
        

           




	
	end if '** AntalErr **'
	
	
	 '** hvis der skal tjekkes values fra timeindlæsningen brus .end her ***'
	 '** 
     'Response.end

        '********************************************
        '***** Indlæs materialer fra TimeOut mobile
        '********************************************   
                               
        if cint(origin) = 12 then 'TimeOut mobile
            


                if len(trim(request("FM_jobid"))) <> 0 AND request("FM_jobid") <> 0 then
                jobid = request("FM_jobid")
                else
                jobid = 0
                end if
    

               thisMid = session("mid")

                if len(trim(request("FM_matids"))) <> 0 then


              


                aftid = 0 'skal sættes på jobid
                '**** Aftale på job ****'
                SQLaftid = "SELECT serviceaft FROM job WHERE id = " & jobid
                oRec3.open SQLaftid, oConn, 3
                if not oRec3.EOF then
                    aftid = oRec3("serviceaft")
                end if
         

                strEditor = session("user")
                strDato = year(now)&"/"&month(now)&"/"&day(now)
                regDatoSQL = strDato
            
                
                '*** Valuta kurs ***'
			    intValuta = intValuta 'fra job

                   

			    call valutaKurs(intValuta)
                dblKurs = dblKurs

                'Kan ikke sættes PT.
                intkode = 2 'Altid fakturerbar når der indtastes via mobil
                bilagsnr = 0 'Kommer ikke med på bilagsnummer
                personlig = 0 'ikke personlig

             

                'Altid 0 via mobil
                avaGrp = 0
                
                'aktid altid 0 via mobil
                aktid = 0

                matIds = split(request("FM_matids"), ", ") '0
                intAntals = split(request("FM_matantals"), ", ")
                
                
                for i = 0 to UBOUND(matIds)

                        prs = 1 '0: Brug indtastede priser, 1: Hent fast priser fra DB.
			            call matidStamdata(matIds(i), avaGrp, prs)
                        strNavn = strMatNavn 
                        strVarenr = strMatVarenr
                        strEnhed = strMatEnhed

                     if (dblSalgsPris) <> 0 then
                     dblKobsPrisBeregn = replace(dblKobsPris, ".", ",")
                     dblSalgsPrisBeregn = replace(dblSalgsPris, ".", ",")
                     ava = 1 - ((formatnumber(dblKobsPrisBeregn, 2)/1) / (formatnumber(dblSalgsPrisBeregn, 2)/1))
                     ava = replace(ava, ",", ".")
                     else
                     ava = 0
                     end if


                    'call insertMat_fn(matIds(i), intAntals(i), strNavn, strVarenr, dblKobsPris, dblSalgsPris, strEnhed, jobid, strEditor, strDato, thisMid, avaGrp, regDatoSQL, aftid, intValuta, intkode, bilagsnr, dblKurs, personlig)
                    call insertMat_fn(matIds(i), intAntals(i), strNavn, strVarenr, dblKobsPris, dblSalgsPris, strEnhed, jobid, aktid, strEditor, strDato, thisMid, avaGrp, regDatoSQL, aftid, intValuta, intkode, bilagsnr, dblKurs, personlig, ava)

                          

                      if strVarenr <> "0" then	
						
				         
			                  '** Nedskriver varelager ***
			                  strSQL2 = "UPDATE materialer SET antal = (antal - "& intAntals(i) &") WHERE id = "& matIds(i)
				              oCOnn.execute(strSQL2)
			             
			      
			          end if


                next


                end if 'request("FM_matids")








                '** Email notifikation ***************************'

               

                if len(trim(request("FM_lukjobstatus"))) <> 0 then

                    jstatus = request("FM_lukjobstatus")
                    call lukjobmail(jstatus, jobid, thisMid)

                end if


                '*** Nulstiller evt. forvalgt job udsendt via mail **'
                session("tomobjid") = 0


        end if 'origin



        
        
           '*** IF origin <> 12 (IKKE mobil WEB) ***
        '*** Tjekker typer og fjerner forudfyldt frokost. ***'
        '*** Der KAN ikke være frokost på dage med ferie, ferie fridag, barsel eller Omsorgsdage
        '*** Oprydnings funktion FLYT TIL kontrolpanel ***'
        if session("mid") = -100 then 'admin bruger session RETTIGHEDER / support
        if cint(multitildel) <> 1 then
        select case lto
        case "intranet - local", "kejd_pb"


                 
                    last24md = dateAdd("d", -760, now)
                    last24md = year(last24md)&"/"&month(last24md)&"/"&day(last24md)
                    strSQLsletFrokost = "SELECT timer, tdato, tmnr FROM timer WHERE ((tfaktim = 13 OR tfaktim = 14 OR tfaktim = 22 OR tfaktim = 23) AND tmnr <> 0 AND tdato >= '"& last24md &"') OR "_
                    &" ((tfaktim = 7 OR tfaktim = 8 OR tfaktim = 20 OR tfaktim = 21 OR tfaktim = 24 OR tfaktim = 25 OR tfaktim = 31 OR tfaktim = 81 OR tfaktim = 115) AND timer > 5 AND tmnr <> 0 AND tdato >= '"& last24md &"')"
                   
                        'response.write strSQLsletFrokost &"<br><br>"
                        'response.flush    
                    
                     oRec.open strSQLsletFrokost, oConn, 3
                    while not oRec.EOF 
                        
                            '** Skal være over 0 timer, da der ellers har været indlæst på dagne og der vil blive tilbudt nye 0,5 timer på timereg. siden **'
                            tdatoSQL = year(oRec("tdato"))&"/"&month(oRec("tdato"))&"/"&day(oRec("tdato"))
                            strSQlfrDel = "UPDATE timer SET timer = 0 WHERE tdato = '"& tdatoSQL & "' AND tmnr = "& oRec("tmnr") &" AND tfaktim = 10 AND timer > 0"
                            
                                response.write strSQlfrDel &"<br><br>"
                                response.flush 
                                oConn.execute(strSQlfrDel)
                                  


                    oRec.movenext
                    wend
                    oRec.close
                  

        case else

        end select
        end if

        'response.end
         end if


       
	
	    '**** Opdaterer stopurs siden / Redirect til Timregside ****'
	    if stopur = "1" then 
	    Response.Redirect "stopur_2008.asp?reload=1"
	    else

        select case rdir
        case "timetag_web"
        Response.Redirect "../timetag_web/timetag_web.asp?indlast=1"
        case "ugeseddel_2011"

        usemrn = request("usemrn")
        varTjDatoUS_man = request("varTjDatoUS_man")

        'response.write "ugeseddel_2011.asp?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man
        'response.end

        Response.Redirect "ugeseddel_2011.asp?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man

        

        case else
	    Response.Redirect "timereg_akt_2006.asp?showakt=1" 
        end select
      
	    end if
	
	
	
	end if '** func = db **'
	
	


    '**************************************************************'
    '************* Opdater WIP ingangværnede job og tilbud ********'
    '**************************************************************'
	
    if func = "opdigv" then

    call stadeopdater()

    


    Response.redirect "timereg_akt_2006.asp"

    end if

	
	
	'******************************************************'
	'************ Opdater Smiley **************************' (Bør får sin egen plads / side, da den kan afsluttes fra flere siders)
	'******************************************************'
	 if func = "opdatersmiley" then
	 
     call smiley_agg_fn()

     call opdaterSmiley
    
     usemrn = request("usemrn")
     rdir = request("rdir")
     varTjDatoUS_man = request("varTjDatoUS_man")
    
        
        '*** Vender tilbage til timereg ****'
        %>
        <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
        <div style="position:absolute; top:100px; left:200px; width:400px; padding:10px; border:10px #CCCCCC solid; background-color:#ffffff;">
        
        <table cellspacing="0" cellpadding="10" border="0" bgcolor="#ffffff" width=100%>
	    <tr>
	        <td bgcolor="#ffffff" style="padding-top:15px;">
                <h4>
                   
                  <%if cint(SmiWeekOrMonth) = 0 then 'uge aflsutning  %>
                    <%=tsa_txt_004 & " din "& tsa_txt_005 %> 
                   <%else %>
                     <%=tsa_txt_004 & " din "& tsa_txt_430 %>
                   <%end if %>
                    </h4>
             
               <%if cint(hidesmileyicon) = 0 then %>  
            <span><img src="../ill/<%=smileyImg%>" border=0 /></span>
                <%end if %>

	        </td>
	        <td valign="top" align=right bgcolor="#ffffff">
		    &nbsp;
		    </td>
	        </tr>
	        <tr>
	        <td colspan=2 valign=top>
            <b><%=tsa_txt_001%>:</b><br /><br />

            <%if instr(formatdatetime(cDateUge, 2), "1899") <> 0 then %>
            Der er opstået en fejl.<br />
            Du har glemt at vælge uge/måned!<br /><br />
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
	        <h4>
                <%if cint(SmiWeekOrMonth) = 0 then 'uge aflsutning  %>
                <%=tsa_txt_005 &" "& datepart("ww", cDateUgeTilAfslutning, 2, 2)%> 
                <%else 
                    
                afslutmd = dateAdd("m", -1, cDateUge)
                    
                  %>
                <%=monthname(month(afslutmd)) & ", "& year(afslutmd) %>
                <%end if %>
                <%=smileysttxt %></h4>
	        
           

	        </td>
	    </tr>
	<tr>
	<td>

    <%select case rdir
     case "ugeseddel_2011"
    %>
    <a href="ugeseddel_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=varTjDatoUS_man%>"><%=tsa_txt_006 %> >></a>
    <% 
    case "logindhist"
    %>
    <a href="logindhist_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=varTjDatoUS_man%>"><%=tsa_txt_006 %> >></a>
    <%
     case else %>
	 <a href="timereg_akt_2006.asp"><%=tsa_txt_006 %> >></a>
     <%end select %>
        
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

        call medarbiprojgrp(tuid_progrp, tuid_mid, 0, -1)

         
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
      '******************************************************************************************************************************
    '** Timereg siden start 
	'******************************************************************************************************************************
        
        	
	%>

     
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
  
	<SCRIPT src="inc/smiley_jav.js"></script>
	<SCRIPT src="inc/timereg_2012_4_func.js"></script>
   <SCRIPT src="inc/matind_2014_jav.js"></script>
   
	
    <%call browsertype()
    
    if browstype_client <> "ip" then%>
    

	
   
  
	
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

                    'Response.write "Her cookie uemrn = " & request.Cookies("tsa")("usemrn")
                    'response.end 
                    else
                    usemrn = session("mid")
                    end if
                end if
            	
            	response.cookies("timereg_2006")("usemrn") = usemrn

    
    '******************************************************************************
    '****************** Job *******************************************************
    '******************************************************************************


             

             
                '****** Sætter job aktive i timereg_usejob baseret på de valgte job i select boks i jobbanken +
                '****** de åbne job på personlig aktiv listen **'
                
               
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
                    'strSQLUpdFvOff = "UPDATE timereg_usejob SET forvalgt = 0 WHERE medarb = "& usemrn &"" 
                    'Response.Write "<br><br><br><br>Nulstilelr: " & strSQLUpdFvOff 
                    'oConn.Execute(strSQLUpdFvOff)
                    '******************************************************'
                   

                else	

                'Response.Write "<br><br>jobidjobidjobid<br><br>: jobidjobidjobid: " & jobid 
                    
                        
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
                dtNow = year(now) &"/"& month(now) &"/"& day(now)
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




                                idTU = 0
                                forvalgt = 0
                                strSQLTUfindes = "SELECT forvalgt FROM timereg_usejob WHERE jobid = "& thisJobId &" AND medarb = "& usemrn
                                'Response.Write "##"& strSQLTUfindes & "<br><br>"
                                oRec5.open strSQLTUfindes, oConn, 3
                                if not oRec5.EOF then

                                forvalgt = oRec5("forvalgt") 
                                'idTU = oRec5("id")

                                end if
                                oRec5.close

                                 
                                 'Response.Write "&nbsp;&nbsp;&nbsp;&nbsp; id:"& idTU &" der: "& forvalgt &"<br>"
                              
                                '** Sætter forvalgt job for denne medarbejder ****'
                                 if cint(forvalgt) = 0 then
                                strSQLUpdFvOn = "UPDATE timereg_usejob SET forvalgt = 1, forvalgt_sortorder = 0, forvalgt_af = "& session("mid") &", forvalgt_dt = '"& dtNow &"' WHERE medarb = "& usemrn &" AND jobid = "& thisJobId
                                'Response.write "<br><br>xxx<br>&nbsp;&nbsp;&nbsp;&nbsp;"& strSQLTUfindes &"::"& forvalgt &" ... "& strSQLUpdFvOn & "<br>"
                                'Response.flush
                                'Response.write strSQLUpdFvOn & "<hr>"
                                oConn.Execute(strSQLUpdFvOn)
                                end if
                                '*************************************************'

                        next

                          'Response.end

                else
                        
                        dtNow = year(now) &"/"& month(now) &"/"& day(now)
                        seljobidSQL = seljobidSQL & " j.id = "& jobid &" "
                        Redim jobids(0) 
                        jobids(0) = jobid

                                
                                forvalgt = 0
                                strSQLTUfindes = "SELECT id, forvalgt FROM timereg_usejob WHERE jobid = "& jobid &" AND medarb = "& usemrn
                                oRec5.open strSQLTUfindes, oConn, 3
                                if not oRec5.EOF then

                                forvalgt = oRec5("forvalgt") 

                                end if
                                oRec5.close


                                 'Response.Write "  &nbsp;&nbsp;&nbsp;&nbsp;  her<br>"

                                if cint(forvalgt) <> 1 then
                                '** Sætter forvalgt job for denne medarbejder ****'
                                strSQLUpdFvOn = "UPDATE timereg_usejob SET forvalgt = 1, forvalgt_sortorder = 0, forvalgt_af = "& session("mid") &", forvalgt_dt = '"& dtNow &"' WHERE medarb = "& usemrn &" AND jobid = "& jobid
                                oConn.Execute(strSQLUpdFvOn)
                                end if
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
	
        

		'if len(request.Cookies("tsa")("timereltid")) <> 0 then
		'visTimerElTid = request.Cookies("tsa")("timereltid")
		'else
            
            visTimerElTid = 0    
            strSQLmedarbSettingTimerEltid = "SELECT timer_ststop FROM medarbejdere WHERE mid = "& usemrn
            oRec6.open strSQLmedarbSettingTimerEltid, oConn, 3
            if not oRec6.EOF then

            visTimerElTid = oRec6("timer_ststop")

            end if
            oRec6.close

		    
		
        'end if
		
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


    '** Tilbuds / Akttype ***
    if len(request("FM_chkfor_igakttype")) <> 0 then
	    
	   
	    if len(request("FM_igakttype")) <> 0 then
	    ignorerakttype = 1
	    else
	    ignorerakttype = 0
	    end if
	
	else
	    
	    if request.Cookies("tsa")("igakttype") <> "" then
		ignorerakttype = request.Cookies("tsa")("igakttype")
		else
        select case lto
        case "xintranet - local" ', "essens", "synergi1", "qwert"
        ignorerakttype = 1
        case else
		ignorerakttype = 0
		end select
        end if
	   
	end if

    
	
	
	'**** Job og Akt. periode ***'
    call lukaktvdato_fn()

    if cint(lukaktvdato) = 1 then 'force fjern akt. ved dato overskreddet
        
        ignJobogAktper = 1

    else 

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

    end if
    
    '**** vis simpel akt linje ******'
    call visAktSimpel_fn()

        '** Sættes i kontrolpanel ***'
        if cint(visAktlinjerSimpel) = 1 then
        visSimpelAktLinje = 1
	    else
	    visSimpelAktLinje = 0
	    end if

    '** Opdatere medarb. indstillinger ***'
    if len(request("FM_chkfor_visSimpelAktLinje")) <> 0 then
       
        vistimereltid = request("FM_vistimereltid")

        strSQLtimer_ststop = "UPDATE medarbejdere SET timer_ststop = "& vistimereltid & " WHERE mid = "& usemrn
        oConn.execute(strSQLtimer_ststop)
        
    end if

    ' if len(request("FM_chkfor_visSimpelAktLinje")) <> 0 then
	'    
	'   
	'    if len(request("FM_visSimpelAktLinje")) <> 0 then
	'    visSimpelAktLinje = 1
	'    else
	'    visSimpelAktLinje = 0
	'    end if
	
	'else
	    
	'    if request.Cookies("tsa")("visSimpelAktLinje") <> "" then
    '	visSimpelAktLinje = request.Cookies("tsa")("visSimpelAktLinje")
	'	else

   '     select case lcase(lto)
   '     case "kejd_pb", "kejd_pb2", "intranet - local", "hestia", "tec", "esn"
'		visSimpelAktLinje = 1
	'	case else
    '    visSimpelAktLinje = 0
    '    end select

     '   end if
	   
	'end if

    'response.Cookies("tsa")("visSimpelAktLinje") = visSimpelAktLinje
	response.Cookies("tsa")("ignJobogAktper") = ignJobogAktper
	response.Cookies("tsa")("igtidslas") = ignorertidslas
    response.Cookies("tsa")("igakttype") = ignorerakttype


    if len(trim(jobid)) <> 0 then
    jobid = jobid
    else
    jobid = 0
    end if


    '**tildel dage iflg egen nomrtid sætte aktid via cookioe på indlæs timer ****'
                        
    'if request.cookies("tsa")("tildelidage") = "1" then
    tildeliheledageVal = 1
    tildeliheledageSEL = "CHECKED"
    'else
    'tildeliheledageVal = 0
    'tildeliheledageSEL = ""
    'end if
                
                
      



	
    'response.Cookies("tsa")("jobid") = jobid
	response.Cookies("showakt") = showakt
	response.Cookies("tsa")("usemrn") = usemrn
	response.Cookies("tsa")("timereltid") = visTimerElTid
	response.Cookies("tsa").expires = date + 32
	
	
	call erSDSKaktiv()


	
	level = session("rettigheder")
	timenow = formatdatetime(now, 3)
	

	
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
	



    vlgtmtypgrp = 0
    'call mtyperIGrp_fn(vlgtmtypgrp,1) 
    call fn_medarbtyper()	
	 
 '************************************************************************************************** 
  'Opbygger timereg SQL states
  '**************************************************************************************************
 	            

        visIgnorerprojgrp = 0
        select case lto
        case "wwf", "wwf2"
            if session("rettigheder") <= 2 then
            visIgnorerprojgrp = 1
            else
            visIgnorerprojgrp = 0
            end if
        case else
        visIgnorerprojgrp = 1
        end select
 	        
				
				
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
				
                'if len(trim(request("FM_kontakt"))) <> 0 then
				
                'if len(trim(request("FM_sog_kun_valgte"))) <> 0 then
                'sogKunValgteKontakt = "CHECKED"
                'else
                'sogKunValgteKontakt = ""
                'end if

				'response.cookies("tsa")("sogkunvalgte") = sogKunValgteKontakt
				'else
				'    if request.cookies("tsa")("sogkunvalgte") <> "" then
				'    sogKunValgteKontakt = "CHECKED"
				'    else
				'    sogKunValgteKontakt = ""
				'    end if
				'end if
                
				
				if len(request("FM_kontakt")) <> 0 then
				sogKontakter = request("FM_kontakt")
				
				
				response.cookies("tsa")("kontakt") = sogKontakter
	            'response.cookies("tsa").expires = date + 65
				
				else
				    
				        if request.Cookies("tsa")("kontakt") <> "" _
				        AND request.Cookies("tsa")("kontakt") <> 0 _
				        then
				        sogKontakter = request.Cookies("tsa")("kontakt")
				        else
				        sogKontakter = 0
                        end if
				end if
				

                '*** ignorer projektgrupper ***'
                '** Er der søgt **'
				if len(request("FM_kontakt")) <> 0 then

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

                    response.Cookies("tsa")("ignprogrp") = ignorerProgrp


                else
				    
                    if request.Cookies("tsa")("ignprogrp") <> "" then
                    ignorerProgrp = request.Cookies("tsa")("ignprogrp") 
                       
                        if cint(ignorerProgrp) = 1 then
					    selIgn = "CHECKED"
					    else
					    selIgn = ""
					    end if

                      else

                      ignorerProgrp = 0
				    selIgn = ""

				      end if
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

                if len(request("FM_hr")) <> 0 then
				    intHR = request("FM_hr") '1
					if cint(intHR) = 1 then
					selHR = "CHECKED"
					else
					selHR = ""
					end if
					
				else
				    if request.cookies("hr") = "1" then
				    intHR = 1
				    selHR = "CHECKED"
				    else
				    intHR = 0
				    selHR = ""
				    end if
				end if


           

                 
                 if len(request("FM_classicreg")) <> 0 then
				    intClassicreg = request("FM_classicreg") '1
					if cint(intClassicreg) = 1 then
					selClassic = "CHECKED"
					else
					selClassic = ""
					end if
					
				else
				    if request.cookies("classicreg") = "1" OR (intSmartreg = 0 AND intEasyreg = 0 AND intNyeste = 0 AND intHR = 0) then
				    intClassicreg  = 1
				    selClassic = "CHECKED"
				    else
				    intClassicreg  = 0
				    selClassic = ""
				    end if
				end if


                end if


                '*** Vis kun jobansvarlig ***'
                 if len(request("FM_kontakt")) <> 0 then 'er der søgt form (filter)
                    
                     if len(trim(request("FM_jobansv"))) <> 0 then
                     jobansvCHECKED = "CHECKED"
                     jobansv = 1
                     else
                     jobansvCHECKED = ""
                     jobansv = 0
                     end if
                     

                        

                    if len(trim(request("FM_visAlleMedarb"))) <> 0 then
                    visAlleMedarbCHK = "CHECKED"
                    visAlleMedarb = 1
                    else
                    visAlleMedarbCHK = ""
                    visAlleMedarb = 0
                    usemrn = session("mid")
                    end if


                   response.cookies("jobansv") = jobansv
				   response.cookies("tsa")("visAlleMedarb") = visAlleMedarb
                   
                   response.cookies("tsa").expires = date + 10

                     
                 else
                        
                        if request.cookies("jobansv") = "1" then
                        jobansvCHECKED = "CHECKED"
                        jobansv = 1
                        else
                        jobansvCHECKED = ""
                        jobansv = 0
                        end if


                    if request.cookies("tsa")("visAlleMedarb") = "1" then
                    visAlleMedarbCHK = "CHECKED"
                    visAlleMedarb = 1
                    else
                    visAlleMedarbCHK = ""
                    visAlleMedarb = 0
                    usemrn = session("mid")
                    end if
                     
                 end if

             
               


                 

			     sortVal = 0


         if session("mid") = 100000 then
                %>
                <div style="position:absolute; top:480px; z-index:100;">
                    response.cookies("tsa_2012")("showfilter"): <%=request.cookies("tsa_2012")("showfilter") %><br />
                    request.cookies(tsa)(visAlleMedarb)<%=request.cookies("tsa")("visAlleMedarb") %> // visAlleMedarb: <%=visAlleMedarb %><br />
                    session.SessionID = <%=session.SessionID %>
                </div>
                <%                
                end if	





	
    
    %>
     <div id="kalender" style="position:absolute; left:870px; top:102px; width:210px; visibility:<%=dKalVzb%>; display:<%=dKalDsp%>;">
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

    
    'call treg_3menu(thisfile)

    call menu_2014()

    '*** Alert ved forkal. timebudget over skreddet
    call akt_maksbudget_treg_fn
    %>

    
    
    
   
    
   
	
    
    <div id="loadbar" style="position:absolute; display:; visibility:visible; top:300px; left:200px; width:300px; background-color:#ffffff; border:10px #CCCCCC solid; padding:10px; z-index:100;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" /><br />&nbsp;
  
	</td></tr>
    <tr><td colspan=2>  <div id="load_cdown">Forventet loadtid: 4-7 sekunder...</div></td></tr>
    </table>

	</div>






<div id="sidediv" style="top:102px; left:90px;">
 	
	
 <%
    

				
				
	if cint(intHR) = 1 then 'HR LISTEN (tjekker egne rettigheder, da man selv skal adgang til at indtaste sygdom for at man må gøre det for andre)
        call hentbgrppamedarb(session("mid")) 
    else 'valgte medarb		
     	   '*** Hvis man indtaster for an anden medarb. ER TEAMLEDER allers kan session("mid") aldrig være <> usemrn og level <=2 OR level = 6. Admin + niveau 1 
            
            if (usemrn <> session("mid")) AND (level = 2 OR level = 6) AND cint(ignorerProgrp) = 1 then
                call hentbgrppamedarb(session("mid")) 
            else
	            call hentbgrppamedarb(usemrn)
            end if
     end if
				
	
	'*******************************************************************************'
	'***************************** Jobbanken Filter              *******************'
	'*******************************************************************************'
	
    call meStamdata(session("mid"))
    if ((lto = "mmmi" AND meType = 10) OR lto = "xintranet - local") OR request.cookies("tsa_2012")("showfilter") = "1" OR (session("forste") = "j" AND lto = "wwf") OR (session("forste") = "j" AND lto = "ngf") then
    autohidefilter = 1 '1 irriterende
    else
    autohidefilter = 0
    end if

    response.cookies("tsa_2012")("showfilter") = autohidefilter
    response.cookies("tsa_2012").expires = date + 65


     if cint(stempelurOn) = 1 then 
            lft = 630
        else
            lft = 720
        end if


     %>
    <div style="position:absolute; background-color:#ffffff; border:0px #5582d2 solid; padding:4px; top:82px; width:75px; left:<%=lft%>px; z-index:0;"><a href="<%=lnkUgeseddel%>" class="vmenu"><%=tsa_txt_337 %></a></div>
    
    <%if cint(stempelurOn) = 1 then %>
    <div style="position:absolute; background-color:#ffffff; border:0px #5582d2 solid; padding:4px; top:82px; width:85px; left:720px; z-index:0;"><a href="<%=lnkLogind%>" class="vmenu"><%=tsa_txt_340 %></a></div>
    <%end if
   
	call filterheaderid(102,20,745,pTxt,fiVzb,fiDsp,"filterTreg", "relative")%>

     <form action="timereg_akt_2006.asp" id="filterkri">



	<table cellspacing=0 width=100% cellpadding=5 border=0>

    
	
   
	<input type="hidden" name="FM_autohidefilter" id="FM_autohidefilter" value="<%=autohidefilter %>"><!--request.cookies("tsa")("showfilter") :request.cookies("tsa")("showfilter")  -->
    <input type="hidden" name="FM_sesMid" id="FM_sesMid" value="<%=session("mid") %>">
    <input type="hidden" name="FM_ch_medarb" id="FM_ch_medarb" value="0">
    <input type="hidden" name="" id="sprog_filterheader" value="<%=tsa_txt_384 %>">
    <!--<input type="text" name="tildeliheledageSet" id="Hidden3" value="<%=tildeliheledageVal%>" />-->
    <input type="hidden" name="dtson" id="datoSonSQL" value="<%=now%>">
	<%if level = 1 then
	SQLkriEkstern = ""
    end if


	
	''*** Finder medarbejdere i de progrp hvor man er teamleder ***'
    'Response.Write "level " & level
	 
     
    call medarb_teamlederfor


	%>
	<tr bgcolor="#ffffff">
	<td valign=top rowspan=2 style="padding-top:5px; padding-bottom:5px;">
        
 

    <%if len(trim(strSQLmids)) > 0 then 'Hvis man er level 1 eller teamleder vil len(trim(strSQLmids)) ALTID VÆRE > 16 %>
    <b><%=tsa_txt_077 %>:</b> <br />
    <input type="CHECKBOX" name="FM_visallemedarb" id="FM_visallemedarb" value="1" <%=visAlleMedarbCHK %> /> <%=tsa_txt_388 %> (<%=tsa_txt_357 %>)

      
   
	<br />
				<% 
				call medarb_vaelgandre
                %>

					<br /><br /><img src="ill/blank.gif" border=0 height=13 width=1 /><br />
				<%
                 else
                   %>
                <input type="hidden" name="usemrn" id="usemrn" value="<%=usemrn %>" />
                <%   
                 end if 'if len(trim(strSQLmids)) > 16

              
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

        <input type="hidden" name="FM_kontakt" id="FM_kontakt" value="0" />
        <!--
	<select name="FM_kontakt" id="FM_kontakt" style="width:250px;" size="12"> 
		
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
        -->
       	
       
       	<b><%=tsa_txt_078 %>:</b><br /><%=replace(tsa_txt_073, "XTX9", "<>") %><br />
        <!-- søg -->
		<input type="text" name="FM_sog_job_navn_nr" id="FM_sog_job_navn_nr" value="<%=show_sogjobnavn_nr%>" style="font-family:arial; font-size:11px; line-height:14px; width:250px; border:2px #6CAE1C solid;">
        <!--<br /><input type="checkbox" name="FM_sog_kun_valgte" id="FM_sog_kun_valgte" value="1" <%=sogKunValgteKontakt%>> Søg kun job på valgte kontakt-->
		<!-- ignorer projektgrupper -->
        <%
       
        
        if cint(visIgnorerprojgrp) = 1 then%>
        <br /><input type="checkbox" name="FM_ignorer_projektgrupper" id="FM_ignorer_projektgrupper" value="1" <%=selIgn%>> <%=tsa_txt_074%> <span id="a_ignproj_info" style="color:#999999;"><b>?</b></span>
        <div id="sp_ignproj_info" style="color:#999999; display:none; visibility:hidden; padding:5px; border:0px #cccccc solid;"><br /><b>Administrator:</b> Viser alle aktiviteter.<br /><b>Teamleder:</b> viser egne rettigheder på aktiviteter, så Du kan taste for andre medarbejdere på de aktiviteter Du har rettighed til.
        <p id="al_ignproj_info" style="color:#999999;">X</p></div>
        <%else %>
        <input type="checkbox" name="FM_ignorer_projektgrupper" id="FM_ignorer_projektgrupper" value="1" style="visibility:hidden;" >
        <%end if %> 
		<br />

        <%
        '**** Vis Guiden aktivejob LINK ****'
        showGuiden = 0
        call positiv_aktivering_akt_fn()
        if (cint(positiv_aktivering_akt_val) <> 1 OR level = 1) then 
            
            showGuiden = 1 

            select case lto 
            case "assurator"

             if level = 1 then
             showGuiden = 1 
            else
            showGuiden = 0
            end if

            case else 
            
             showGuiden = 1 
                
            end select
            
                      if cint(showGuiden) = 1 then%>
                    <a href="javascript:popUp('guiden_2006.asp?mid=<%=usemrn%>','850','620','150','120');" target="_self" class=rmenu>
		            <%=tsa_txt_082%> >>
		            </a>
                    <%else %>
		            &nbsp;
                    <%end if %>

        <%end if %>
           
         

				</td>
				
				
	    <td valign=top style="height:250px; padding-top:5px;"><b>Filter:</b><br />
            <input id="FM_nyeste" name="FM_nyeste" type="radio" value="1" <%=selNye%> /> <%=tsa_txt_334 %>
            
            <%
            
            call showEasyreg_fn()
            if cint(showEasyreg_val) = 1 then %>
            &nbsp;&nbsp;<input id="FM_easyreg" name="FM_easyreg" type="radio" value="1" <%=selEasy%> /> <%=left(tsa_txt_358, 7) %> 
            <%end if %>
		
			<%select case lto
            case "x"
            %>
            &nbsp;&nbsp;<input id="FM_smartreg" name="FM_smartreg" type="radio" value="1" <%=selSmart%> /> <%=tsa_txt_373 %>
            <%
            case else
            end select %>
            
            &nbsp;&nbsp;<input id="FM_classicreg" name="FM_classicreg" type="radio" value="1" <%=selClassic%> /> <%=tsa_txt_374 %> (vis alle)

            <%select case lto
            case "xintranet - local", "mmmi", "unik"
            %>
           
            <%
            case else

            %>
            &nbsp;&nbsp;<input id="FM_hr" name="FM_hr" type="radio" value="1" <%=selHR%> /> <%=tsa_txt_380 %>
            <%
            end select %>
           
            
            <br /><input id="FM_jobansv" name="FM_jobansv" type="checkbox" value="1" <%=jobansvCHECKED%> /> <%=tsa_txt_379 %>
            
         <br /><br /><br />
        
        <b><%=tsa_txt_336 %>:</b>
	    
        
	    <%if level <= 3 OR level = 6 then %>
	    &nbsp;<a href="jobs.asp?menu=job&func=opret&id=0&int=1&rdir=treg" class=rmenu><%=tsa_txt_335 %> >></a>
        <%end if %>
	    
    
	   
      
     <div id="sorterdiv" style="position:absolute; left:575px; top:132px; width:220px; border:0px #cccccc solid; font-size:9px; color:#999999;">
         <b><%=tsa_txt_341 %>:</b> <input id="sort0" name="sort" type="radio" value=0 CHECKED /> <%=tsa_txt_209 %>
             <input id="sort2" name="sort" type="radio" value=2  /> <%=tsa_txt_397 %> 
            
            <!--
            <input id="sort1" name="sort" type="radio" value=1 <%=sort1CHK%> /> <%=tsa_txt_342 %>&nbsp;
             <input id="sort2" name="sort" type="radio" value=2  /> <%=tsa_txt_343 %>&nbsp;
              <input id="sort0" name="sort" type="radio" value=0 CHECKED /> <%=tsa_txt_344&", "&lcase(tsa_txt_342) %>&nbsp;
              <input id="sort3" name="sort" type="radio" value=3  /> <%=tsa_txt_344&", "&lcase(tsa_txt_343) %>
              
              -->
              </div>
             


       
	    

           
         
	   
       <!--
	   <textarea id="fajl" cols="40" rows="40"></textarea>
       -->
       
       
       <div name="div_jobid" id="div_jobid" style="width:420px; height:220px; overflow:auto; padding:10px; border:1px #CCCCCC solid; font-size:11px;">
	    <!-- henter job fra jquery -->
	    </div>

        <%select case lto
        case "mi", "intranet - local", "hvk_bbb", "jttek"
        lmtTxt = "200"
        case else
        lmtTxt = "100"
        end select %>
       
        <span style="font-size:9px; color:#999999; padding:5px 0px 4px 0px;"><%=tsa_txt_394 &" "& lmtTxt &" "& tsa_txt_395 %></span>
        <br /><input type="checkbox" name="FM_viskunetjob" id="FM_viskunetjob" value="1" /><%=tsa_txt_396 %>
        
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
	    <input type="submit" id="sogsubmit" value="<%=tsa_txt_378%> >>">
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

			if pgeditok = 1000 then 'pgeditok = 1 DENNE ER FARLIG. ER LUKKET NED. ÅBEN IGEN PÅ REQUEST til REKLAME bureauer
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
	

        '*** Hvis padding på divs skal bredden være forskellig i browserne
        
    
            if browstype_client = "ie" then
            gbldivWdt = 745
            else
            gbldivWdt = 725
            end if





    '*************************************************************'
    '************ Højre mnubar på timereg. siden *****************'
    '**************** Aktive job, job til afv. +TODO LIST ********'
    '*************************************************************'

   
    call hojreDiv()









    '*********************************************************************************
	'*** Data til timreg.. siden, projektgrupper, medarb info mm. ********************
    '*****************************************************************************
        
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
				&" FROM job j "_
                &" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_  
				&" LEFT JOIN medarbejdere m1 ON m1.mid = j.jobans1"_
				&" LEFT JOIN medarbejdere m2 ON m2.mid = j.jobans2"_
				&" LEFT JOIN medarbejdere m3 ON m3.mid = k.kundeans1"_
				&" LEFT JOIN serviceaft s ON (s.id = j.serviceaft)"_
				&" LEFT JOIN kontaktpers kk ON (kk.id = kundekpers)"_
				&" WHERE j.id = " & jobids(0) 
                'AND k.kid = j.jobknr
				
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

		<div id="jinf_knap" style="position:absolute; left:870px; top:632px; width:210px; visibility:<%=jinf_knap_vzb%>; display:<%=jinf_knap_dsp%>; z-index:100; border:1px #cccccc solid; padding:2px; background-color:#ffffff;">
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
                 
                 
                 
                        
            
				
				
				
				
				
				<!-- Kontakt info div -->
	            <div id="kpersdiv" name="kpersdiv" style="position:absolute; left:0px; top:0px; width:400px; border:2px #6CAE1C solid; padding:10px; visibility:hidden; display:none; background-color:#ffffff; z-index:200; height:300px; overflow:auto;">
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
	
	end if '** denne?
    
	'Response.flush
	
	
	call browsertype()
	if browstype_client = "mz" then
	topAdd_1 = 16
	else
	topAdd_1 = 0
	end if
	

      '** Normtimer for ugen for viste medar ****'
    call normtimerPer(usemrn, tjekdag(1), 6, 0)



    normtimer_man = ntimMan
    normtimer_tir = ntimTir
    normtimer_ons = ntimOns
    normtimer_tor = ntimTor
    normtimer_fre = ntimFre
    normtimer_lor = ntimLor
    normtimer_son = ntimSon

 
    normTimerprUge = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig) 
    totNormTimer = (normtimer_man + normtimer_tir + normtimer_ons + normtimer_tor + normtimer_fre + normtimer_lor + normtimer_son)


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

    
	
      


	
     '*** SMILEY faneblade / afslut uge ****'
     if smilaktiv <> 0 AND showakt <> 0 AND cint(projektgrpOK) <> 0 then


        call medrabSmilord(usemrn)
        call smiley_agg_fn()
        call smileyAfslutSettings()

          call meStamdata(usemrn)

          '** Henter timer i den uge der skal afsluttes ***'
          call afsluttedeUger(year(now), usemrn)

         '** Er kriterie for afslutuge mødt? Ifht. medatype mindstimer og der må ikke være herreløse timer fra. f.eks TT
         call timeKriOpfyldt(lto, sidsteUgenrAfsl, meType, afslutugekri, usemrn)

    

         call timerDenneUge(usemrn, lto, tjkTimeriUgeSQL, akttypeKrism)
     
           'response.Write "SmiWeekOrMonth: "& SmiWeekOrMonth

         if cint(SmiWeekOrMonth) = 0 then 'uge aflsutning
         weekMonthDate = datepart("ww", tjekdag(7),2,2)
         else
         weekMonthDate = datepart("m", tjekdag(7),2,2)
         end if

         call erugeAfslutte(datepart("yyyy", tjekdag(7),2,2), weekMonthDate, usemrn) 


        %>

         <div style="position:relative; left:20px; width:745px; top:140px; padding:10px 0px 0px 0px;">


             <table cellspacing="0" cellpadding="2" border="0" width="100%" >
         <tr>
	    <td valign=top style="padding:4px; width:620px;">

            <%call smileyAfslutSettings()
             

             call smileyAfslutBtn(SmiWeekOrMonth)

             call ugeAfsluttetStatus(tjekdag(7), showAfsuge, ugegodkendt, ugegodkendtaf) %>
           
            
         

               </td>

             <td align="right" valign="top">

                    <!-- skift uge pile -->
                 <%  prevWeek = datepart("ww", useDatePrevWeek, 2,2) 
                     nextWeek = datepart("ww", useDateNextWeek, 2,2)  %>

	                        <table cellpadding="0" cellspacing="0" border="0" width="100%">
	                        <tr>
                             
	                        <td valign=top align=right style="width:140px; white-space:nowrap;"><a href="timereg_akt_2006.asp?showakt=1&strdag=<%=day(useDatePrevWeek)%>&strmrd=<%=month(useDatePrevWeek)%>&straar=<%=year(useDatePrevWeek)%>">< <%=tsa_txt_005 &" "& prevWeek %></a></td> <!-- jobid=<=jobid%>&usemrn=<=usemrn%>& &fromsdsk=<=fromsdsk%> -->
                           <td style="padding-left:20px; width:140px; white-space:nowrap;" valign=top align=right><a href="timereg_akt_2006.asp?showakt=1&strdag=<%=day(useDateNextWeek)%>&strmrd=<%=month(useDateNextWeek)%>&straar=<%=year(useDateNextWeek)%>"><%=tsa_txt_005 &" "&nextWeek %> ></a></td>
	                        </tr>
	                        </table>

             </td>
       
    </tr>
     </table>
      </div>


        
        <%
	    '**** Afslut uge ***'
	    '**** Smiley vises hvis sidste uge ikke er afsluttet, og dag for afslutninger ovwerskreddet ***' 12 
        if cint(SmiWeekOrMonth) = 0 then
        denneUgeDag = datePart("w", now, 2,2)
        s0Show_sidstedagsidsteuge = dateAdd("d", -denneUgeDag, now) 'now
        else
        s0Show_sidstedagsidsteuge = now
        end if

        '** finder kriterie for rettidig afslutning
        call rettidigafsl(s0Show_sidstedagsidsteuge)

        if cint(SmiWeekOrMonth) = 0 then
            s0Show_weekMd = datePart("ww", s0Show_sidstedagsidsteuge, 2,2)
        else
            s0Show_weekMd = datePart("m", s0Show_sidstedagsidsteuge, 2,2)
        end if

        
        '** tjekker om uge er afsluttet og viser afsluttet eller form til afslutning
		call erugeAfslutte(year(s0Show_sidstedagsidsteuge), s0Show_weekMd, usemrn)
      
       
        'if session("mid") = 1 then
        '    response.write "<br><br><br><br><br><br>"& datepart("w", now, 2,2) & "= 1 AND " & datepart("h", now, 2,2) &" session vist: " & session("smvist")  & " smiley_agg:  "& smiley_agg & "<br>"_
        '    &" now: "&  cDate(formatdatetime(now, 2)) & ">="& cDate(formatdatetime(cDateUge, 2)) & " ugeNrStatus: " & ugeNrStatus
            
        'end if

        if (cDate(formatdatetime(now, 2)) >= cDate(formatdatetime(cDateUge, 2)) AND cint(ugeNrStatus) = 0) OR cint(smiley_agg) = 1 then


	    'if (datepart("w", now, 2,2) = 1 AND datepart("h", now, 2,2) <= 23 AND session("smvist") <> "j") OR cint(smiley_agg) = 1 then

       
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

          
              

    	    
            '*** Auto popup ThhisWEEK now SMILEY
	        if cint(smilaktiv) = 1 then%>
	        <div id="s0" style="position:relative; left:20px; top:142px; width:725px; visibility:<%=smVzb%>; display:<%=smDsp%>; z-index:2; background-color:#FFFFFF; padding:10px; border:0px #CCCCCC solid;">
	      
                 <%
               
           '*** Viser sidste uge
            'weekSelected = tjekdag(7)

            '*** Viser denne uge
            weekSelectedThis = dateAdd("d", 7, now) 'tjekdag(7)

	        call showsmiley(weekSelectedThis, 1)

            
            call afslutkri(tjekdag(7), tjkTimeriUgeDt, usemrn, lto)


            if cint(afslutugekri) = 0 OR ((cint(afslutugekri) = 1 OR cint(afslutugekri) = 2) AND cint(afslProcKri) = 1) OR cint(level) = 1 then 
            
                

	        call afslutuge(weekSelectedTjk, 1, tjekdag(7), "")

             
         

            else
            

             '** Status på antal registrederede projekttimer i den pågældende uge    
             select case lto
             case "tec", "esn"
             case else %>
                
            <div style="color:#000000; background-color:#FFFFFF; padding:5px; border:1px red solid;">
                
            <%=tsa_txt_398 & ": "& totTimerWeek %> 
                
                 <%if afslutugekri = 2 then %>
                    <%=tsa_txt_399 %>
                    <%end if %>
                
                <%=" "&tsa_txt_140 %> / <%=afslutugeBasisKri %> = <b><%=afslProc %> %</b> <%=tsa_txt_095 %> <b><%=datePart("ww", tjkTimeriUgeDtTxt, 2,2) %></b>
                <br />(<%=left(weekdayname(weekday(formatdatetime(tjkTimeriUgeDt, 2))), 3) &". "& formatdatetime(tjkTimeriUgeDt, 2)%> - <%= left(weekdayname(weekday(formatdatetime(dateAdd("d", 6, tjkTimeriUgeDt), 2))), 3) &". "&formatdatetime(dateAdd("d", 6, tjkTimeriUgeDt), 2) %>)
               
                 <%=tsa_txt_400 %>: <b><%=afslutugekri_proc %> %</b>  <%=tsa_txt_401%>.
               </div>
                <br />
            <%

            end select

            end if
	        
	        %>
               
	      <br /><br />
	        <span id="se_uegeafls_a" style="color:#5582d2;">[+] <%=tsa_txt_402 &" "& year(tjekdag(4))%></span><br /><%
	        useYear = year(tjekdag(4))
            '** Hvilke uger er afsluttet '***
	        call smileystatus(usemrn, 1, useYear)
	        
                
                
                
             %>
	        <br />&nbsp;

               
	        </div>
        	
	        <%end if 'visSmliley


                'Response.write "<br><br><br><br><br><br><br><br><br><br>antalAfsDato + 3) < useDateSmileyTjkWeek:<br>"
                'Response.write "antalAfsDato: "& antalAfsDato &"<"& useDateSmileyTjkWeek
                    
                select case lto 
                case "intranet - local", "dencker"
                slip = 10 'antal uger ind der lukkes
                case else
                slip = 3
                end select   


                'response.write "cDateUge: "& cDateUge
                'response.flush
                indevarendeUgeDt = dateAdd("d", -7, cDateUge)
                'indevæørende ugne inde afslutnings tidspunkt
                if cdate(now) < cdate(indevarendeUgeDt) then
                slip = slip + 1
                end if

                %>

                                            <!--
                                <br><br><br><br><br><br><br><br><br><br><br /><br />             <br />
                        &nbsp;&nbsp;&nbsp;&nbsp;<%=cdate(now) &"< " & cdate(indevarendeUgeDt)%>

                        &nbsp;&nbsp;&nbsp;&nbsp;<%=antalAfsDato &" "& slip &" < "& useDateSmileyTjkWeek%>
                        -->

<%

                 if ((antalAfsDato + slip) < useDateSmileyTjkWeek) AND cint(smiley_agg) = 1 then
                    %>


                           <br><br><br><br><br><br><br><br><br><br><br /><br />                      
                    <div style="position:relative; left:20px; background-color:lightpink; width:720px; padding:10px;">
                    <h4>Du kan ikke længere registrere timer</h4> 
                        Det skyldes højst sandsynligt en af følgende årsager:<br /><br />
                        A) Du har for mange uafsluttede uger. <br /><br />
                        B) Din bruger er blevet de-aktiviteret. <br /><br />
                        C) Din fratrædelsesdato er overskreddet. <br /><br />&nbsp;

                        
                     
                        

                        </div>
                    

                <%
                    '** Stop indlæsning, så der ikke kan tidsregistrers ***'
                    if cint(level) <> 1 then
                    response.end
                    end if

                end if
                    
              


                'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>antalAfsDato: "& antalAfsDato & " useDateSmileyTjkWeek: "& useDateSmileyTjkWeek

               
     


    	
        else

                    %>


                       <!-- skift uge pile -->
                 <%  prevWeek = datepart("ww", useDatePrevWeek, 2,2) 
                     nextWeek = datepart("ww", useDateNextWeek, 2,2)  %>

                    <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
     

	                        <table cellpadding="0" cellspacing="0" border="0" width="745">
	                        <tr>
                             <td style="width:450px;">&nbsp;</td>
	                        <td valign=top align=right style="width:140px; white-space:nowrap;"><a href="timereg_akt_2006.asp?showakt=1&strdag=<%=day(useDatePrevWeek)%>&strmrd=<%=month(useDatePrevWeek)%>&straar=<%=year(useDatePrevWeek)%>">< <%=tsa_txt_005 &" "& prevWeek %></a></td> <!-- jobid=<=jobid%>&usemrn=<=usemrn%>& &fromsdsk=<=fromsdsk%> -->
                           <td style="padding-left:20px; width:50px; white-space:nowrap;" valign=top align=right><a href="timereg_akt_2006.asp?showakt=1&strdag=<%=day(useDateNextWeek)%>&strmrd=<%=month(useDateNextWeek)%>&straar=<%=year(useDateNextWeek)%>"><%=tsa_txt_005 &" "&nextWeek %> ></a></td>
	                        </tr>
	                        </table>

                      

        <%
	    end if 'visSmliley

    

   



    
   
	


    %>
    <form action="timereg_akt_2006.asp?FM_use_me=<%=usemrn%>&func=db&fromsdsk=<%=fromsdsk%>&isfiltersubmitted=1" method="post" name="timeregfm" id="timeregfm">
    <%
	
	if showakt <> 0 AND cint(projektgrpOK) <> 0 then
	
         
    
    tTop = 0 '53 + topAdd_1 '(225+topAdd)
	tLeft = 20
	tWdth = 745
	tId = "timeregdiv"
    
        
    if cint(smilaktiv) <> 0 then%>
    <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
    <%end if %>
    
  

    <div style="position:relative; left:20px; width:745px; padding:10px 0px 0px 0px;">
     <table width=100% cellspacing=1 cellpadding=2 border=0>
     <tr>
     <td valign="bottom">
    
    <%if cint(intEasyreg) = 1 then%>
	<h4>Easyreg. listen 
	    <%if level = 1 then %>
        <span>
	    <a href="javascript:popUp('easyreg_2010.asp','800','720','20','20')" class=rmenu><%=lcase(tsa_txt_251)%> >></a>
        </span>
	    <%end if %>
        <hr />
	
	<%
    
    nextDvminus = 18
    else %>


     <%if cint(intHR) <> 1 then  %>
     <%call meStamdata(usemrn) %>
        
	 <h4><a href="#" name="anc_s0" id="pagesetadv_but" style="font-size:9px; color:#999999; font-weight:lighter;">[ + <%=tsa_txt_302%> ]</a> 
      &nbsp;<a href="#" name="anc_s1" id="A4" onclick="showmultiDiv()" style="font-size:9px; color:#999999; font-weight:lighter;">[ + Multitildel ]</a>
         
            
      <br /><%=tsa_txt_385 %>: <span style="font-size:11px; font-weight:lighter;"><%=meTxt %>
	<hr />
	<%


   '***************************** SORTER EFTER, sortBy *******************
  


    select case lto
    case "xintranet - local", "xoko"

    'if level = 1 then
    'sorterEfterShow = 1
    'else
    sorterEfterShow = 0
    'end if

    case else
    sorterEfterShow = 1

    end select
    

   if cint(sorterEfterShow) = 1 then


   sortByUline1 = tsa_txt_413
   sortByUline2 = tsa_txt_414
   sortByUline21 = tsa_txt_415
   sortByUline3 = tsa_txt_416
   sortByUline4 = tsa_txt_417
   sortByUline5 = tsa_txt_418
   sortByUline6 = tsa_txt_419
  

   select case sortBy
   case 1
   sortByUline1 = "<u>"& tsa_txt_413 &"</u>" '"<u>Prioitet</u>"
   case 2
   sortByUline2 = "<u>"& tsa_txt_414 &"</u>" '"<u>Jobnavn</u>"
   case 21
   sortByUline21 = "<u>"& tsa_txt_415 &"</u>" '"<u>Jobnr</u>"
   case 3
   sortByUline3 = "<u>"& tsa_txt_416 &"</u>" '"<u>Slutdato</u>"
   case 4
   sortByUline4 = "<u>"& tsa_txt_417 &"</u>" '"<u>Kunde</u>"
   case 5
   sortByUline5 = "<u>"& tsa_txt_418 &"</u>" '"<u>"& sortByUline5 &"</u>"
   case 6
   sortByUline6 = "<u>"& tsa_txt_419 &"</u>" '"<u>"& sortByUline6 &"</u>"
   case else
   sortByUline1 = "<u>"& tsa_txt_413&"</u>" '"<u>Prioitet</u>"
   end select %>

   <%=tsa_txt_412 %>: <a href="timereg_akt_2006.asp?showakt=1&FM_mnr=<%=usemrn%>&sortby=2" class=rmenu><%=sortByUline2 %></a> 
   | <a href="timereg_akt_2006.asp?showakt=1&FM_mnr=<%=usemrn%>&sortby=21" class=rmenu><%=sortByUline21 %></a> 
   | <a href="timereg_akt_2006.asp?showakt=1&FM_mnr=<%=usemrn%>&sortby=1" class=rmenu><%=sortByUline1 %></a> 
   | <a href="timereg_akt_2006.asp?showakt=1&FM_mnr=<%=usemrn%>&sortby=3" class=rmenu><%=sortByUline3 %></a>
   | <a href="timereg_akt_2006.asp?showakt=1&FM_mnr=<%=usemrn%>&sortby=4" class=rmenu><%=sortByUline4 %></a>

   <%select case lto
   case "xunik", "xmmmi"
   case else %>
   | <%=tsa_txt_420 %>: <a href="timereg_akt_2006.asp?showakt=1&FM_mnr=<%=usemrn%>&sortby=5" class=rmenu><%=sortByUline5 %></a>
   <%end select %>
   
    <%select case lto
   case "wwf", "wwf2", "intranet - local", "unik", "mmmi"
   %>
   | <a href="timereg_akt_2006.asp?showakt=1&FM_mnr=<%=usemrn%>&sortby=6" class=rmenu><%=sortByUline6 %></a>
   <%
   case else %>
   
   <%end select


 
   
   
   call showuploadimport_fn() 
   if cint(showuploadimport) = 1 then%>
   
   <br /> <a href="../../ver2_10/timereg_net/importer_timer.aspx?lto=<%=lto%>&editor=<%=session("user") %>&mid=<%=usemrn %>" class=rmenu target="_blank"> Upload & importér .csv fil >></a>
   
   <%end if

   nextDvminus = 20


   end if 'sorterEfterShow lto, vis sorterings muligheder
   
   else ' HR mode%>
   <%
   call meStamdata(usemrn) %>

          
        
	 <h4><a href="#" name="anc_s0" id="pagesetadv_but" style="font-size:9px; color:#999999; font-weight:lighter;">[ + <%=tsa_txt_302%> ]</a> 
      &nbsp;<a href="#" name="anc_s1" id="A4" onclick="showmultiDiv()" style="font-size:9px; color:#999999; font-weight:lighter;">[ + Multitildel ]</a>
         
    <br />HR listen: <span style="font-size:11px; font-weight:lighter;"><%=meTxt %>	<hr /> 

   <%
   nextDvminus = 18 
   sortBy = 4 ' altid efter kunde
   end if %>
   

    <%
    end if 
    %>

     <%
	    'oskrift = tsa_txt_031 &" "& datepart("ww", tjekdag(1), 2 ,2) & " "& datepart("yyyy", tjekdag(1), 2 ,3) 
	    %>
	
 
     </span></h4>
    </td>
    

    </tr>
    </table>
	
   
	  <table cellspacing="0" cellpadding="2" border="0" width="100%" >
	<tr>
	    <td valign=top colspan=7 style="padding-top:4px; width:650px;">
	   
	
    <%
    
    'if lto = "wwf" OR lto = "wwf2" then
    'fcTxt = "timebudget"
    'else
    fcTxt = "forecast"
    'end if
    
     if browstype_client <> "ip" AND level = 1 then
     
                       
     
     if cint(sortby) <> 5 AND cint(sortby) <> 6 AND cint(intEasyreg) <> 1 then 
     %>
     <input id="tildelalle" name="tildelalle" type="hidden" value="0" /> <!--<span style="color:#000000;"><%=tsa_txt_268 %> (<%=tsa_txt_357 %>)</span>-->
     <%
     end if


  %>


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

                      
		
		                '*** Min indtast ETS **'
		
		                if lto = "ets-track" OR lto = "xintranet - local" then %>
		                <tr><td colspan=2><input id="minindtast_on" name="minindtast_on" type="checkbox" value="1" CHECKED /> <span style="color:#000000;">Gennemtving min. indtastning 7,4 timer / 8,0 timer</span>
		                 </td></tr>
		                <%end if


                        select case lto
                        case "mmmi", "sdutek"
                            showVisAlleMedTilknyt = 1
                        case else
                            showVisAlleMedTilknyt = 0
                        end select 

                        if showVisAlleMedTilknyt = 1 then

                                 if cint(level) = 1 then 
                                   if cint(sortBy) = 5 OR cint(sortBy) = 6 then%>
		                          <tr><td colspan=2 style="background-color:#FFFFFF; padding:10px;">
                              
                                        <input name="visallemedarb_blandetliste" id="visallemedarb_blandetliste" type="checkbox" value="1" /> Vis alle medarbejdere tilknyttet (via aktiv jobliste)
                                    
                                          job: <input type="text" name="jobnr_blandetliste" id="jobnr_blandetliste" style="width:150px;" /> <input type="button"  id="bn_visallemedarb_blandetliste" value=">>" /><br />
                                      <span style="font-size:9px; color:#CCCCCC;">Maks. 50 medarbejdere</span>
                                       </td></tr>
		              
		                        <%end if
                         
                                    end if 

                         end if    
                         %>

                        <%'end if 'licensejer
                        
                       
                        %>
		
		                

                       

		                <tr><td colspan=2 valign=top><br /><br />
                        
                         <%

                          
		                '*** Multi tildel ***'
		                'if browstype_client = "ip" then
		                %>
                        
                        
                        <!--
                        <input type="CHECKBOX" id="abnlukallejob" name="abnlukallejob" value="1" /> Luk (kollaps) alle aktiviteter på joblisten (indlæser timer)<br />
                        -->
		            
		                                <div id="multivmed" style="position:relative; padding:20px; background-color:#FFFFFF; width:520px; visibility:hidden; display:none;">
                                        

                                        <%strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe, medarbejdertype, mt.type AS mtypenavn, mt.id AS mtid "_
                                        &", normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, normtimer_fre, normtimer_lor, normtimer_son "_
                                         &" FROM medarbejdere AS m LEFT JOIN medarbejdertyper AS mt ON (mt.id = m.medarbejdertype) WHERE mansat <> 2 "& medarbgrpIdSQLkri &" ORDER BY mt.type, Mnavn" %>
                                        
                                        <h4>Multitildel <span><a href="#" class="red" onclick="showmultiDiv()">[x]</a></span></h4>
                                        <%=tsa_txt_268 %> (<%=tsa_txt_357 %>)<br /><br />
                                      
		                                <b><%=tsa_txt_267 %>:</b><br /> <select name="tildelselmedarb" id="tildelselmedarb" size=10 multiple style="width:500px;">
				                                <%
					                                
					                                oRec.open strSQL, oConn, 3
					
					                                m = 0
                                                    while not oRec.EOF 
					
					                                
                                                    if lastMtyp <> oRec("mtid") then
                                                     if m <> 0 then%>
                                                    <option DISABLED></option>
                                                    <%end if %>

                                                       <%normtidpruge = formatnumber(oRec("normtimer_man") + oRec("normtimer_tir") + oRec("normtimer_ons") + oRec("normtimer_tor") + oRec("normtimer_fre") + oRec("normtimer_lor") + oRec("normtimer_son"),2)  %>


                                                    <option DISABLED><%=oRec("mtypenavn") %> // norm: <%=normtidpruge %> t. pr. uge</option>
                                                    <%end if
                                                    
                                                    
                                                    if cint(oRec("Mid")) = cint(usemrn) then
					                                rchk = "SELECTED"
					                                else
					                                rchk = ""
					                                end if%>

                                                 
					                                <option value="<%=oRec("Mid")%>" <%=rchk%>><%=oRec("mnavn") &" "&"("& oRec("mnr") %>)</option>
					                                <%
					
					                                m = m + 1
                                                    lastMtyp = oRec("mtid")
					                                oRec.movenext
					                                wend
					                                oRec.close
				                                %></select>

                                                <br />
                                               
		                                            <input id="Checkbox1" name="tildeliheledageSet" type="checkbox" value="1" <%=tildeliheledageSEL %>/> <span style="color:#000000;"><%=replace(tsa_txt_277, "XXXX", "<br>")%> </span><br />&nbsp;
		                                            <!--<input id="tildeliheledage" name="tildeliheledage" type="text" value="<%=tildeliheledageVal %>"/>-->
                                                    </div>
		
		
		
		                                <%
		                                'end if
		                                %>
		
		                                </td></tr>
                                        </table>

                                        <!--</div>-->
		                                <br />&nbsp;

        
                            
		                    </div> <!-- indlaes_settings -->




    <%end if %>

	</td>
     <td colspan=2 align=right valign=bottom style="padding-bottom:5px;"><br /><input type="submit" id="sm1" value="<%=tsa_txt_085 %> >> "></td>
	
	</tr>
	
	
	
	
    
    <!--<input id="showjobinfo" name="showjobinfo" type="hidden" value="<=showjobinfo %>" />-->
	
    
		
	
                   

	
	                 <%
                     if cint(intEasyreg) = 1 then
     
                     %>
                
                     <tr bgcolor="#DCF5BD">
                        <td colspan=2 style="border-right:1px #CCCCCC solid; border-left:2px #CCCCCC solid; border-top:2px #CCCCCC solid; white-space:nowrap;">Fordel timer på Easyreg. aktiviteter</td>
        
                        <%for e = 1 to 7 %>
                        <td style="border-right:1px #CCCCCC solid; border-top:2px #CCCCCC solid;">
                            <input id="ea_<%=e %>" type="text" style="width:46px; font-size:9px;" /><a href="#" id="ea_k_<%=e%>" class="ea_kom"><font color="#999999"> + </font></a></td>
                        <%next %>
     
     
                     </tr>
                     <tr bgcolor="#DCF5BD">
                     <td colspan=2 style="border-right:1px #CCCCCC solid; border-left:2px #CCCCCC solid; border-top:1px #CCCCCC solid; border-bottom:2px #CCCCCC solid;">(0 = nulstil)</td>
                     <%for e = 1 to 7 %>
                            <td style="border-right:1px #CCCCCC solid; border-top:1px #CCCCCC solid; border-bottom:2px #CCCCCC solid; width:55px;">
                                <input id="ea_kn_<%=e %>" type="button" value="Calc. =" style="font-family:arial; font-size:9px;" /></td>
                        <%next %>
     

                     </tr>
                     <%
                     end if
                     %>
     

      

      </table>
    </div><!--tableDiv-->



     

     <div id="timereg_job" style="position:relative; border:0px; top:0px; left:20px; width:740px; visibility:visible; display:; z-index:100">
    
     <%
    'tTop = 153
	'tLeft = 20
	'tWdth = 0 '760
	'tId = "timereg_job"
	
	'call tableDivWid(tTop,tLeft,tWdth,tId, dTimVzb, dTimDsp)
    
   
    'if session("mid") = 21 AND lto = "dencker" OR lto = "intranet - local" then
    ' timeC = now
    'loadtime = datediff("s",timeA, timeC)
    'Response.write "tid: "& loadtime
    'end if
     
    if cint(sortBy) = 1 then 'sorterliste Drag'n drop mode'
         tblid = "incidentlist"
    else
         tblid = ""
    end if         
    %>
   
   
    
    <table id="<%=tblid%>" cellspacing=0 cellpadding=0 border=0 width=100%>
  
    <%
    'Response.Write "jobid:"& request("jobid") & "<br>"
    'Response.Write "seljobid" & request("seljobid")
    
    %>
    
	<!--<input type="hidden" name="tildeliheledageSet" id="Hidden2" value="<%=tildeliheledageVal%>">-->
    <input type="hidden" name="hideallbut_first" id="hideallbut_first" value="<%=hideallbut_first%>">
    <input type="hidden" name="jobid" id="sortBy5jobids" value="<%=seljobidUse%>">
    <input type="hidden" name="sortByval" id="sortByval" value="<%=sortBy%>">
    
	<input type="Hidden" name="FM_vistimereltid" value="<%=visTimerElTid%>">
	<input type="Hidden" name="FM_start_dag" id="treg_dag" value="<%=strdag%>">
	<input type="Hidden" name="FM_start_mrd" id="treg_md" value="<%=strmrd%>">
	<input type="Hidden" name="FM_start_aar" id="treg_aar" value="<%=straar%>">
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


    <input type="hidden" id="FM_hentjava" name="FM_hentjava" value="1">
   
   <input type="hidden" id="jq_level" name="jq_level" value="<%=level %>">
   <input type="hidden" id="jq_lto" name="jq_lto" value="<%=lto %>">
   <input type="hidden" id="jq_frokostalert" value="0">
   <input type="hidden" id="browstype" value="<%=browstype_client %>">     
   <input type="hidden" id="aktnotificer" value="0">
   <input type="hidden" id="aktnotificer_fc" value="0">  
   <input type="hidden" id="akt_maksbudget_treg" value="<%=akt_maksbudget_treg%>">  
   <input type="hidden" id="akt_maksforecast_treg" value="<%=akt_maksforecast_treg%>">
   <input type="hidden" id="aktBudgettjkOn_afgr" value="<%=aktBudgettjkOn_afgr%>">
   <input type="hidden" id="regskabsaarStMd" value="<%=datePart("m", aktBudgettjkOnRegAarSt, 2,2)%>">
   <input type="hidden" id="regskabsaarUseAar" value="<%=straar%>">
   <input type="hidden" id="regskabsaarUseMd" value="<%=strmrd%>">     

    <%call traveldietexp_fn()

    %>
        <input type="hidden" id="traveldietexp_on" value="<%=traveldietexp_on%>">
        <input type="hidden" id="traveldietexp_maxhours" value="<%=traveldietexp_maxhours%>">
        <input type="hidden" id="aktnotificer_trvl" value="0"> 
        <%

    


  
	
	'*** Brugergrupper er fundet i linie 2697 ***
	'call hentbgrppamedarb(usemrn)
	
	iRowLoop = 1
	m = 58 '16
    if cint(intEasyreg) = 1 then
    x = 3500 ' Bør kun være 750, men der kan forekomme mere end 100 Easyreg. aktiviteter indtil loft effketureres
    else
	x = 750 '550 '350 '6500 '2500 (maks 100 aktiviteter * 7 dages timereg.)
    end if
			
	dim aktdata
	redim aktdata(x, m) 
	
					
					
					
    '*** Aktiviteter og Timer SQL MAIN **** 
	strSQL = "SELECT a.navn, a.id AS aid, a.fakturerbar, "_
	&" j.jobnr AS jnr, j.id AS jid, j.jobnavn AS jnavn, j.risiko, job_internbesk, "_
    &" j.serviceaft, j.beskrivelse AS jobbesk, j.jobans1, jobans2, jobstartdato, jobslutdato, j.fastpris, jo_bruttooms, j.rekvnr, j.kommentar, j.kundekpers, j.budgettimer, "_
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
	
	
	strSQL = strSQL &", kkundenavn, kkundenr, kid, kundeans1 "
    strSQL = strSQL &" FROM job j"
    strSQL = strSQL &" LEFT JOIN aktiviteter AS a ON (a.job = j.id) "
	
   
	if cint(intEasyreg) = 1 then
	strSQL = strSQL &" LEFT JOIN timereg_usejob AS tu ON (tu.jobid = j.id) "
	else
        if cint(positiv_aktivering_akt_val) = 1 then
        strSQL = strSQL &" LEFT JOIN timereg_usejob AS tu ON (tu.jobid = j.id AND tu.aktid = a.id AND tu.medarb = "& usemrn &") "
        else
            
           'tjekker rettigheder for den der er logget på og ikke valgte medarb. Da man skal ha rettighede til f.ejs sygdom for at taste  den for andre. 
            '* Det er ligegyldt om den valgte medarb. har adgang. selv. F.eks hvsi SYGDOM eller lign. blvier indtastet fra centralt hold.
            '** SÆTTES i 6145 hentbgrppamedarb()
            strSQL = strSQL &" LEFT JOIN timereg_usejob AS tu ON (tu.jobid = j.id  AND tu.medarb = "& usemrn &") "
           
        end if
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
            
            '** ign projektgrupper på aktiviteter. F.eks for admin bruger til at få 
            '** vist ferie u. løn aktiviteter for en anden medarb. der ikke selv har rettigheder ****'
            '** Ell. HR mode & ADmin  **'
            if (level = 1 AND ignorerProgrp = 1) OR (cint(intHR) = 1 AND level = 1) then
            strSQLkri3 = " a.job = j.id AND aktstatus = 1 "
            else
            '*** func hentbgrppamedarb() L: 6528
            '*** Egne projektgrupper. Hvis teamleder, så den medarbejder 
            '**** man indtaster for hvis der er valgt en anden medarbejder end en selv ***' 
            strSQLkri3 = strSQLkri3
            end if

            '* hvis blandet liste bruges benyttes dette kald ikke, alle Aktiviteer hentes i jquery : OPTIEMRING ***'
            if cint(intEasyreg) = 1 OR cint(sortBy) = 5 OR cint(sortBy) = 6 then
            strSQLkri3 = strSQLkri3 & " AND a.id = 0 "
            end if

            if cint(intHR) <> 1 then '**HR mode viser alle interne RISIKO = -1
            strSQL = strSQL &" WHERE ("& seljobidSQL &") AND (j.jobstartdato <= '"& varTjDatoUS_son &"' AND (j.jobstatus = 1 OR j.jobstatus = 3))  AND "& strSQLkri3
            else
            strSQL = strSQL &" WHERE (j.risiko = -2 AND j.jobstatus = 1) AND "& strSQLkri3
            end if

        end if
	
	
    
	
	'Response.Write  "strSQLkri3" &  strSQLkri3  
	
	if cint(ignJobogAktper) = 1 AND intHR <> 1 then
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
	    strSQL = StrSQL & " GROUP BY j.id, a.id ORDER BY "& sortBySQL &" j.id, a.fase, a.sortorder, LTRIM(a.navn) LIMIT 750"
        'j.risiko, j.jobnavn, j.id, a.fase, a.sortorder, LTRIM(a.navn) LIMIT 550"
	    end if
	'end if
	
	
    'Response.Write strSQL  'j.id =
    'Response.write strSQL & " positiv_aktivering_akt_val:" & positiv_aktivering_akt_val '& "<br><br>" & strSQLkri3

    strJobFundet = "#0#,"
	Response.flush
	
    lastFirstLetter = ""
	tr_vzb = ""
    tr_dsp = ""
	lastFase = ""
    lastJobId = 0
    'tjkLastJidFak = 0
	
    antalJobLinier = 0
    LastJobLoop = 0
    
    
    'if lto = "dencker" AND session("mid") = 21 then
    'response.write strSQL
    'response.flush
    'end if

 
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
            

    if cint(intEasyreg) = 1 OR cint(sortBy) = 5 OR cint(sortBy) = 6 then
    strEaAktidsThisJob = strEaAktidsThisJob & "," & oRec("aid")
    iRowLoopsEAThisJob = iRowLoopsEAThisJob & "," & iRowLoop
    antalaktlinier = antalaktlinier + 1
    loops = antalaktlinier

        if cint(sortBy) = 5 OR cint(sortBy) = 6 then
        strJobFundet = strJobFundet & "#"& oRec("jid") &"#,"
        end if

    else              
	
	'aktdata(iRowLoop, 0) = oRec("tdato")
	'aktdata(iRowLoop, 1) = oRec("timer")
	aktdata(iRowLoop, 2) = oRec("navn")
	aktdata(iRowLoop, 3) = oRec("aid")
    
  

	aktdata(iRowLoop, 4) = oRec("jid")
    strJobFundet = strJobFundet & "#"& oRec("jid") &"#,"

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

    aktdata(iRowLoop, 27) = oRec("kundekpers")

    if isNull(oRec("fase")) <> true then
	aktdata(iRowLoop, 29) = oRec("fase")
    else
	aktdata(iRowLoop, 29) = ""
    end if

    'aktdata(iRowLoop, 30) = oRec("aktbudget")
	aktdata(iRowLoop, 31) = oRec("jnr")

    aktdata(iRowLoop, 33) = oRec("jobstatus")
   

	aktdata(iRowLoop, 34) = oRec("budgettimer")
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
    aktdata(iRowLoop, 20) = oRec("risiko")
    

   


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
	
	
    if cint(intEasyreg) <> 1 AND sortBy <> 5 AND sortBy <> 6 then

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
			if level <= 2 OR level = 6 OR lto = "kejd_pb" OR lto = "kejd_pb2" then
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
                  
                    antalaktlinier = antalaktlinier + 1
                    
                  %>
			<!--</tr>-->
			<%end if%>
			
			
			<%
                                '*** Job ***'
                                if lastJobid <> aktdata(iRowLoop, 4) AND cint(intEasyreg) <> 1 AND sortBy <> 5 AND sortBy <> 6 then

                     
                   
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

                                                    antalAktidsThisJob = 0
                                                    antalAktidsThisJobArr = split(strAktidsThisJob, ",")
                                                        for a = 0 to UBOUND(antalAktidsThisJobArr)
                                                        antalAktidsThisJob = antalAktidsThisJob + 1
                                                        next
                                                        %>
                                            
                                                     
                                                    <input type="hidden"  id="job_antalaktids_<%=lastJobid %>" value="<%=antalAktidsThisJob %>" />

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
                                                           <input type="hidden" name="SortOrder" class="SortOrder" value="0" />
	                                                    <input type="hidden" name="rowId" value="0" />
        

                                                   
                                                    <div id="Div1" style="position:relative; width:745px; border:0px; visibility:visible; display:; z-index:100">
                                                    <table width=100% cellspacing=0 cellpadding=0 border=0>
                                                    <tr><td><img src="../ill/blank.gif" width="1" height="3" border="0" /></td></tr>
                                                    </table>
                                                    </div>

                                                    </td></tr>

            
                                                    <%end if %>

            
                        



                                  
                               
                                 <tr>
                                 <td valign=top>
                                 
                                    <%
                                    '***** Jobnavn ***'
                                    if sortby = 2 then 
                                    
                                    if ucase(left(aktdata(iRowLoop, 5),1)) <> lastFirstLetter then
                                    Response.Write "&nbsp;&nbsp;<br><br><span style=""background-color:#FFFFFF; width:45px; padding:3px;"">"& ucase(left(aktdata(iRowLoop, 5),1)) & "</span>"
                                    end if

                                    lastFirstLetter = ucase(left(aktdata(iRowLoop, 5),1)) 
                                    end if %> 

                                        <%
                                     '*** Slutdato ***
                                     if sortby = 3 then 
                                    
                                    if month(aktdata(iRowLoop, 46)) <> lastFirstLetter then
                                    Response.Write "&nbsp;&nbsp;<br><br><span style=""background-color:#FFFFFF; width:225px; padding:3px;"">"&  monthname(month(aktdata(iRowLoop, 46))) &" "&  year(aktdata(iRowLoop, 46)) & "</span>"
                                    end if

                                    lastFirstLetter = month(aktdata(iRowLoop, 46)) 
                                    end if %> 


                                     <%
                                     '***** Kontakt ****'
                                     if sortby = 4 then 
                                    
                                    if ucase(left(aktdata(iRowLoop, 32),40)) <> lastFirstLetter then
                                    Response.Write "&nbsp;&nbsp;<br><br><span style=""background-color:#FFFFFF; width:225px; padding:3px;"">"& left(aktdata(iRowLoop, 32),40) & "</span>"
                                    end if

                                    lastFirstLetter = ucase(left(aktdata(iRowLoop, 32),40)) 
                                    end if %> 

                                       <input type="hidden" name="SortOrder" class="SortOrder" value="<%=antalJobLinier%>" />
	                                <input type="hidden" name="rowId" value="<%=aktdata(iRowLoop, 4)%>" />
                                     
                                        <input type="hidden" id="FM_jobid_<%=antalJobLinier %>" value="<%=aktdata(iRowLoop, 4) %>" />
                                       
           

          

                                            <div id="ad_timereg_<%=aktdata(iRowLoop, 4)%>" style="position:relative; width:745px; border:0px; visibility:visible; display:; z-index:100">
                                            <table width=100% cellspacing=0 cellpadding=0 border=0 bgcolor="#ffffff">
             
          
            
                                                 <tr>
                                                    <td style="background-color:#D6Dff5; border-bottom:0px #cccccc solid; padding:0px 0px 0px 0px;" colspan=5 class=lille><img src="../ill/blank.gif" width=1 height=2 border=0 /></td>
                                                 </tr>
                
                                                <tr>
                                                    <td style="width:240px; background-color:#5C75AA; border-bottom:3px #8caae6 solid; border-top:1px #FFFFFF solid; padding:3px 3px 0px 4px;" valign=top> <a class="a_treg" id="a_timereg_<%=aktdata(iRowLoop, 4)%>" href="#"><span id="sp_a_timereg_<%=aktdata(iRowLoop, 4)%>" style="font-size:14px; color:#CCCCCC;">[+]</span>&nbsp;<%=left(aktdata(iRowLoop, 5), 50) %> (<%=aktdata(iRowLoop, 31) %>)</a>
                                                    <%select case aktdata(iRowLoop, 33)
                                                    case 3
                                                    Response.write "<span style=""font-size:9px; color:#FFFFFF;""> - Tilbud</span> "
                                                    end select %>

                                                    <%if editok = 1 then %>
                                                    <a href="jobs.asp?menu=job&func=red&id=<%=aktdata(iRowLoop, 4) %>&int=1&rdir=treg" style="color:yellowgreen; font-weight:lighter;"><sub>[ <%=left(tsa_txt_251, 3) %>. ]</sub></a>
                                                    <%end if %>
                                                    <br /><span style="color:#cccccc; font-size:9px;"><%=left(aktdata(iRowLoop, 32), 30) & " " & aktdata(iRowLoop, 56) %></span>
                                                    </td>

                                                    <td style="width:15px; background-color:#5C75AA; border:0px #cccccc solid; border-bottom:3px #8caae6 solid; border-top:1px #FFFFFF solid; border-right:0px #cccccc solid; padding:3px 3px 0px 4px;" valign=top>
                                                    <%if aktdata(iRowLoop, 20) > -1 then %>
                                                    <span class="a_det" id="a_det_<%=aktdata(iRowLoop, 4)%>"><img src="../ill/document.png" border="0" alt="Stamdata, job beskrivelse & Materiale registrering"/></span>
                                                    <%else%>
                                                    &nbsp;
                                                    <%end if%>
                                                    </td>
                    
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
                                                    <td valign=top style="padding:3px 0px 0px 3px;">
                                                        <%  if cint(sortBy) = 1 then 'sorterliste Drag'n drop mode' %>
                                                        <img src="../ill/pile_drag.gif" alt='Klik, træk og sorter aktivitet' class="drag" border="0" />
                                                        <%else %>
                                                        &nbsp;
                                                        <%end if %>
                                                    </td>
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
                    
                    

                                                    <td valign="bottom" class=lille style="background-color:#FFFFFF; width:300px; border-bottom:3px #8caae6 solid; height:55px; padding:2px 2px 0px 30px;">

                                                    <%select case lto
                                                    case "mmmi", "unik", "intranet - local" 
                                                    shwTw = 0
                                                    case else
                                                    shwTw = 1
                                                    end select 
                                                    
                                                    if cint(shwTw) = 1 then%>
                                                    <input type="text" style="width:258px; font-size:9px; font-family:Arial; color:#999999; font-style:italic;" maxlength="101" value="Job tweet..(åben for alle)" class="FM_job_komm" name="FM_job_komm_<%=aktdata(iRowLoop, 4)%>" id="FM_job_komm_<%=aktdata(iRowLoop, 4)%>"> <a href="#" class="aa_job_komm" id="aa_job_komm_<%=aktdata(iRowLoop, 4)%>">Gem</a>
                                                    <div id="dv_job_komm_<%=aktdata(iRowLoop, 4)%>" style="width:258px; font-size:9px; font-family:Arial; color:#000000; font-style:italic; overflow-y:auto; overflow-x:hidden; height:30px;"><%=aktdata(iRowLoop, 58) %></div>
                                                    <%else %>
                                                    &nbsp;
                                                    <%end if %>
                                                    </td>
                                                    <td valign="top" align=right class=lille style="background-color:#FFFFFF; width:40px; padding:2px 2px 0px 10px; border-bottom:3px #8caae6 solid;">
                                                    <%if cint(intHR) <> 1 then %>
                                                    <span style="color:#999999; font-size:9px; font-style:oblique;">Fjern<br /><input class="fjern_job" type="checkbox" value="1" id="fjern_job_<%=aktdata(iRowLoop, 4)%>"/></span>
                                                    <%else %>
                                                    &nbsp;
                                                    <%end if %></td>
                                                </tr>
               
                                            </table>
                                            </div>
                                            <%response.flush %>

                                </td>
                                </tr>
                               


                                <tr><td>

                                 <!-- job detaljer / stamdata -->
                                <%
                                
                                if aktdata(iRowLoop, 20) > -1 then '** ikke interne
                                call jobbeskrivelse_stdata
                                end if
                                

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
                                <input type="hidden" id="FM_job_laststatus_<%=aktdata(iRowLoop, 4)%>" value="<%=aktdata(iRowLoop, 33) %>" />

                                
                                <%if antalJobLinier = 0 then ''første record ****
                                 %>
                                <input type="hidden" id="first_jobid" value="<%=aktdata(iRowLoop, 4)%>" />
                                <%
                                end if




                                '*** end lastFakdato ****'

          
			                    job_vzb = "hidden"
			                    job_dsp = "none"
                                
                                
                                dv_akt_hgt = 130
                             
			
			
			
			                    '**** Aktiviteterne på hvert job ****'%>
                                <div id="div_timereg_<%=aktdata(iRowLoop, 4)%>" style="position:relative; background-color:#ffffff; height:<%=dv_akt_hgt%>px; overflow:auto; padding:0px 3px 3px 3px; width:745px; border:0px #8caae6 solid; visibility:<%=job_vzb %>; display:<%=job_dsp%>; z-index:100">
                                <!-- Henter aktiviteter --->
                                

                            
                         
                                <%

                                
                                lastJobid = aktdata(iRowLoop, 4)
                                antalJobLinier = antalJobLinier + 1 

                                else 'Easyreg // Blandet liste


                               
                              
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
    <%end if 'Easyreg / Blandetliste %>
	<!--</tr>-->

  
            <%
            if cint(intEasyreg) <> 1 AND sortBy <> 5 AND sortBy <> 6 then
            
                            
                
                
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
            
            else '** Easyreg / blandet liste

                                'if antalJobLinier <> 0 then%>
                                <tr><td>

                                                    <input type="hidden" id="FM_job_lastfakdato_0" value="2001-1-1" />
                                                    <input type="hidden" id="FM_job_laststatus_0" value="0" />
                                                        <input type="hidden"  id="job_aktids_0" value="<%=strEaAktidsThisJob %>" />
                                                      <input type="hidden" name="fasenavn" id="job_fasenavne_0" value="<%=fsVal %>" />
 
                                                    <input type="hidden" name="iRowLoop" id="job_iRowLoop_0" value="<%=iRowLoopsEAThisJob %>" /> 

                                  <div id="div_timereg_0" style="position:relative; background-color:#ffffff; height:200px; overflow:auto; padding:0px 3px 3px 3px; width:745px; border:0px #8caae6 solid; visibility:visible; display:; z-index:2000">
                               
                                      
                                       <!-- Henter aktiviteter --->
                                 <!--
                                 <%if cint(intEasyreg) = 1 then%>
                                 <table width=100%><tr><td style="padding:20px 20px 20px 20px;">Henter Easyreg. aktiviteter 5-10 sek... 
                                <br /><span style="color:darkred; font-size:10px;">Maks. 150 Easyreg-aktiviteter</span><br /><br />&nbsp;</td></tr></table>
                                 <%else%>
                                <table width=100%><tr><td style="padding:20px 20px 20px 20px;">Henter blandet liste 5-10 sek... 
                                <br /><span style="color:darkred; font-size:10px;">Aktiviteter med timer på seneste 
                                <%if cint(sortBy) = 5 then %>
                                14 dage
                                <%else %>
                                2 måneder
                                <%end if %> 
                                vises åbne. (maks 100)<br />
                                Tilføj selv flere linier i bunden</span><br /><br />&nbsp;</td></tr></table>
                                    -->
                                <%end if %>
                                
                                
                                </div>
                                 </td></tr>
           
                            </table>
                                <%
                                'end if
                                'lastJobid = aktdata(iRowLoop, 4)
                                antalJobLinier = antalJobLinier + 1 
        
            
            end if







           
           '***** Antal job på listen ****
           %>
           <input id="antaljob" type="hidden" value="<%=antalJobLinier %>" />
           <%
           
           if cint(intHR) <> 1 AND (cint(sortBy) <> 5 AND cint(sortBy) <> 6) AND cint(intEasyreg) <> 1 then '** IKKE ved HR MODE /Easyreg / og Blandet liste

           strJobVist = replace(seljobidSQL, "OR", "")
           strJobVist = replace(strJobVist, " ", "")

           
           'Response.write strJobFundet & "<br>"
           strJobVistArr = split(strJobVist, "j.id=")

           'Response.Write "Viste job:<br>"
           afslutaktDiv = 0
           for f = 0 TO UBOUND(strJobVistArr)

         

           if instr(strJobFundet, "#"& strJobVistArr(f) & "#") = 0 AND len(trim(strJobVistArr(f))) <> 0 then

            if strJobVistArr(f) < 200000 then


               strSQLf = "SELECT jobnavn, jobnr, id, jobstartdato FROM job WHERE id = "& strJobVistArr(f) &" AND (jobstatus = 1 OR jobstatus = 3) ORDER BY jobnavn "
               oRec5.open strSQLf, oConn, 3
               if not oRec5.EOF then

                if cint(afslutaktDiv) = 0 then
                   %>
                   <br /><br />
                    <div id="Div2" class="jqcorner" style="position:relative; background-color:#ffffff; padding:3px 3px 3px 3px; width:<%=gbldivWdt-25%>px; border:10px #CCCCCC solid; left:0px; visibility:visible; display:; z-index:100">
                    <table cellspacing="0" cellpadding="2" border="0" width="100%"><tr><td style="padding:20px 3px 20px 3px;">
         
                   <h4>Kan ikke vise følgende job</h4>
                   TimeOut kan ikke åbne følgende job på din aktive jobliste af en eller flere grunde:<br /><br />

                   - Der er ikke oprettet aktiviteter på jobbet <br />
                   - Du ikke har adgang til aktiviteterne<br>
                   - Startdato på jobbet er efter ugens startdato<br />
                    - Du har overskredet maksimun antal aktivitslinjer på siden<br /><br /><br />

                    <b>Prøv at:</b><br />
                     - Klik på jobbet herunder for at <b>tilføje projektgrupper</b><br />
                     - Find jobbet <b>på joblisten</b> for at ændre <b>startdato</b><br />
                     - Tilføj <b>aktiviteter</b> på jobbet<br />
                     - Fjern job på den aktive jobliste (maksimum antal aktivitetslinjer er overskreddet)<br />

                   
                     <%afslutaktDiv = 1
                   end if

                   'Response.write "visIgnorerprojgrp:" &visIgnorerprojgrp & "<br><br>"

                    if cint(visIgnorerprojgrp) = 1 AND (lto = "synergi1" OR lto = "outz") then 'visIgnorerprojgrp = 1 DENNE ER FARLIG. ER LUKKET NED. ÅBEN IGEN PÅ REQUEST til REKLAME bureauer%>
                    <br /><a href="#" onclick="Javascript:window.open('../timereg/tilknytprojektgrupper.asp?id=<%=oRec5("id")%>&medid=<%=usemrn %>', '', 'width=650,height=500,resizable=yes,scrollbars=yes')" class=vmenu><%=oRec5("jobnavn") & " ("& oRec5("jobnr") &")" %> >></a>
				   <%
                   else
                   Response.Write "<br><b>"& oRec5("jobnavn") &"</b> ("& oRec5("jobnr") &")" 
                   end if

                   if CDate(varTjDatoUS_man) < cDate(oRec5("jobstartdato")) then
                   Response.write " <span style=""color:darkred;"">"& formatdatetime(oRec5("jobstartdato"), 1) & " !</span> <span style=""color:#999999;"">(startdato på job er efter den valgte uge)</span>"
                   end if
               
               end if
               oRec5.close
           
            end if


           end if

           next

           if cint(afslutaktDiv) = 1 then
           Response.Write "</td></tr></table></div>"
          end if 


          end if 'HR mode



            if antalJobLinier = 0 then %>

            <br /><br />
             <div id="Div6" class="jqcorner" style="position:relative; background-color:#ffffff; padding:3px 3px 3px 3px; width:720px; border:10px #CCCCCC solid; left:0px; visibility:visible; display:; z-index:100">
            <table cellspacing="0" cellpadding="2" border="0" width="100%"><tr><td style="padding:20px 3px 20px 3px;">
            <h4><%=tsa_txt_386 %>!</h4>

            <b><%=tsa_txt_421 %>:</b><br /><br />
            -  <%=tsa_txt_422 %><br />
            -  <%=tsa_txt_423 %><br />
            -  <%=tsa_txt_424 %><br />
            -  <%=tsa_txt_425 %></td></tr></table>
            </div>

               <%end if %>

             <br /><br /><br />&nbsp;<br /><br /><br />














             <%
             
             '******************************************************
             '** Vis ugetotaler i bunden ***'
             '******************************************************
             if (lto = "xintranet - local") OR lto = "xepi" OR (lto = "mmmi" AND meType = 10) then 
             
             else%>


             <h4><%=tsa_txt_387 %>:</h4>
     <div id="dagstotaler" style="position:relative; background-color:#ffffff; padding:3px 3px 3px 3px; width:745px; left:0px; visibility:visible; display:; z-index:100;">
     <table cellspacing="0" cellpadding="2" border="0" width="100%" bgcolor="#cccccc">

     <% 
     call dageDatoer(2) %>
     <%=dageDatoerTxt %>
     
       
    <% '** Henter timer i valgte uge ***'
    'response.write "<hr>"
    call timerDenneUge(usemrn, lto, varTjDatoUS_man, aty_sql_realhours)
        

        %>
        <tr bgcolor="#FFFFFF">
             <td colspan=2 valign=top style="border-bottom:1px #CCCCCC solid;"><b><%=tsa_txt_038 %>:</b></td>
        <%

        for tDay = 0 to 6
    
        select case tDay
        case 0
        timerThis = manTimer
        txtThis = manTxt
            normtimer_This = normtimer_man
        case 1
        timerThis = tirTimer
        txtThis = tirTxt
            normtimer_This = normtimer_tir
        case 2
        timerThis = onsTimer
        txtThis = onsTxt
            normtimer_This = normtimer_ons
        case 3
        timerThis = torTimer
        txtThis = torTxt
            normtimer_This = normtimer_tor
        case 4
        timerThis = freTimer
        txtThis = freTxt
            normtimer_This = normtimer_fre
        case 5
        timerThis = lorTimer
        txtThis = lorTxt
        normtimer_This = normtimer_lor
        case 6
        timerThis = sonTimer
        txtThis = sonTxt 
        normtimer_This = normtimer_son
        end select

        if cdbl(timerThis) > 12 then
        timerThisHgt = 12
        else
        timerThisHgt = timerThis
        end if

             timerThisHgt = (timerThisHgt * 100) / 4
             timerThisHgt = replace(timerThisHgt, ",", ".")

        if timerThis <> 0 then
            timerThisTxt = formatnumber(timerThis, 2)
        else
            timerThisTxt = ""
        end if

        if timerThis > 0 then

        if cdbl(timerThis) > 0 then
        bgColTma = "yellowgreen"
        'bgColTmaAlt = "#000000"
        end if

        if cdbl(timerThis) > 7 then
        bgColTma = "green"
        'bgColTmaAlt = "#000000"
        end if

         if cdbl(timerThis) > 10 then
        bgColTma = "red"
        'bgColTmaAlt = "#000000"
        end if

        else
            bgColTma = "#FFFFFF"
        end if

            %>  <td  align=right valign=bottom style="border-bottom:1px #CCCCCC solid;"><div style="background-color:<%=bgColTma%>; font-size:8px; height:<%=timerThisHgt%>px; width:65px; padding:2px; overflow-x:auto;"><u><%=timerThisTxt%></u>
            <%=txtThis %></div></td>
            <%
        next
        
     %>
   
	
   
	
 
   
  
	<td  align=right valign=bottom valign=top style="border-bottom:1px #CCCCCC solid; white-space:nowrap;">= <a href="<%=lnkUge%>" class=vmenu><%=formatnumber(totTimerWeek, 2)%></a></td>
	
    </tr>



      <tr bgcolor="#FFFFFF">
    <td colspan=2 ><%=tsa_txt_259 %>:</td>
    <td  align=right><%=formatnumber(normtimer_man, 2)%></td>
	<td  align=right><%=formatnumber(normtimer_tir, 2)%></td>
	<td  align=right><%=formatnumber(normtimer_ons, 2)%></td>
	<td  align=right><%=formatnumber(normtimer_tor, 2)%></td>
	<td  align=right><%=formatnumber(normtimer_fre, 2)%></td>
	<td  align=right><%=formatnumber(normtimer_lor, 2)%></td>
	<td  align=right><%=formatnumber(normtimer_son, 2)%></td>
	<td  align=right>= <%=formatnumber(totNormTimer, 2)%></td>
	
    </tr>


    <%
    manTimerBalance = (manTimer - normtimer_man) 
    tirTimerBalance = (tirTimer - normtimer_tir) 
    onsTimerBalance = (onsTimer - normtimer_ons) 
    torTimerBalance = (torTimer - normtimer_tor) 
    freTimerBalance = (freTimer - normtimer_fre) 
    lorTimerBalance = (lorTimer - normtimer_lor) 
    sonTimerBalance = (sonTimer - normtimer_son) 
    totBalance = (manTimerBalance + tirTimerBalance + onsTimerBalance + torTimerBalance + freTimerBalance + lorTimerBalance + sonTimerBalance)

    %>

    <tr bgcolor="#FFFFFF">
		<td colspan=2 style="border-bottom:1px #CCCCCC solid;" >Balance:</td>
		<td  align=right style="border-bottom:1px #CCCCCC solid;"><%=formatnumber(manTimerBalance, 2)%></td>
	<td  align=right style="border-bottom:1px #CCCCCC solid;"><%=formatnumber(tirTimerBalance, 2)%></td>
	<td  align=right style="border-bottom:1px #CCCCCC solid;"><%=formatnumber(onsTimerBalance, 2)%></td>
	<td  align=right style="border-bottom:1px #CCCCCC solid;"><%=formatnumber(torTimerBalance, 2)%></td>
	<td  align=right style="border-bottom:1px #CCCCCC solid;"><%=formatnumber(freTimerBalance, 2)%></td>
	<td  align=right style="border-bottom:1px #CCCCCC solid;"><%=formatnumber(lorTimerBalance, 2)%></td>
	<td  align=right style="border-bottom:1px #CCCCCC solid;"><%=formatnumber(sonTimerBalance, 2)%></td>
    <td  align=right style="border-bottom:1px #CCCCCC solid;">= <%=formatnumber(totBalance, 2)%></td>

	</tr>




    <%'*** If Stempelur = 1 then 
    call erStempelurOn
    if stempelurOn = 1 then
    
    call fLonTimerPer(varTjDatoUS_man, 6, 20, usemrn)

    
    
    end if %>

      

	<tr bgcolor="#ffffff">
		<td colspan=10 align=right valign=top style="padding-top:10px;">
		<input id="sm2" type="submit" value="<%=tsa_txt_085 %> >> "></td>
	</tr>
	<tr bgcolor="#ffffff"><td colspan=10>
    <br />Antal job: <%=antalJobLinier %>
    <br />Antal aktivitetslinier: <%=antalaktlinier %> (<%=loops %>)</td></tr>
    <input id="antalaktlinier" value="<%=antalaktlinier %>" type="hidden" />
   <input id="loops" value="<%=loops %>" type="hidden" />
    
  
    </table>

   
	 

    </div><!-- dagstotalerdiv -->
   <%end if %>
 
    
  



           
  
     



	
	
	<%
     '*************************************************************   
     '***** KOMMENTAR layer ***
     '************************************************************
        
        
     '*** EASY REG ***************%>
    
    <div id="kom_easy" name="kom_easy" style="position:absolute; visibility:hidden; display:none; left:80px; top:110px; z-index:950000000; width:400px; background-color:#ffffff; padding:10px 10px 10px 10px; border:10px #CCCCCC solid;">
	<table width=100% cellpadding=0 cellspacing=2 border=0><tr><td></td>
	<textarea cols="60" rows="10" id="FM_kommentar_easy" name="FM_kommentar_easy"></textarea><br>
        <input id="FM_kom_dagtype" value="1" type="hidden" />
      </td></tr>
      <tr><td align=right>
	<a href="#" class="vmenu" id="close_kom_easy"><%=tsa_txt_056 %>
    &nbsp;<img src="../ill/ikon_luk_komm.png" alt="<%=tsa_txt_056 %>" border="0"> </a><br />
    <br />Kommentar bliver først gemt når timer indlæses. Indlæses på alle aktiviteter den valgte dag.
	</td></tr></table>
	</div>







     <%'*** STANDARD KOMMENTAR ***' %>
    
	<div id="kom" style="position:absolute; visibility:hidden; display:none; z-index:30000; top:0px; left:0px; width:650px; background-color:#ffffFF; padding:20px; border:10px #CCCCCC solid;">
	<table width=100% cellspacing=0 cellpadding=0 border=0>
	<tr>
	<td >
    <h4>Tilføj til timeregistrering<br /><span style="font-size:11px; font-weight:lighter; color:#999999;">Tilføj kommentar og udlæg / materialeforbrug til din timeregistrering på denne aktivitet</span></h4>

	<span id="sp_komm" style="font-size:14px;">[-]</span> <span style="font-size:14px; font-weight:bolder;"><%=tsa_txt_051 %></span>
    </td>

     <td align="right" valign="top"><span style="color:#999999; font-size:14px;" id="sp_closekom">X</span></td>

	</tr> 
    
    <tr class="tr_kom" style="visibility:visible; display:;">
	<td colspan="2">
   
	   
	<input type="hidden" name="rowcounter" id="rowcounter" value="0">
    <input type="hidden" name="daytype" id="daytype" value="1">
    
        
          
	<label style="float:right; padding-right:50px; color:#999999;"><input type="text" name="antch" id="antch" size="3" maxlength="4" style="border:1px #cccccc solid;">  (<%=tsa_txt_052 %>) </label><br><br />
	<textarea cols="75" rows="10" id="FM_kommentar" name="FM_kommentar" onKeyup="antalchar();" style="border:1px #cccccc solid;"></textarea><br>
	
            <%if lto = "xx" then %>
            <br /><%=tsa_txt_053 %>?
	        <select name="FM_off" id="FM_off" style="font-size:9px;">
	        <option value="0" SELECTED><%=tsa_txt_054 %></option>
	        <option value="1"><%=tsa_txt_055 %></option>
	        </select>
            <%else %>
            <input type="hidden" name="FM_off" id="FM_off" value="0">
            <%end if %>

	        <%if kmDialogOnOff = 1 then %>
	        <br />
	        Ved kørselsregistrering,
	        er dette en bopælstur?
	        <select name="FM_bopalstur" id="FM_bopalstur" style="font-size:9px;">
	        <option value="0" SELECTED><%=tsa_txt_054 %></option>
	        <option value="1"><%=tsa_txt_055 %></option>
	        </select>
	        <%else %>
                <input name="FM_bopalstur" id="FM_bopalstur" value="0" type="hidden" />
	        <%end if %>

	</td>
   
	</tr>
    <tr class="tr_kom" style="visibility:visible; display:;"><td colspan="2" align=right style="border-bottom:10px #CCCCCC solid;"><br /><br />
        <input type="button" value="<%=tsa_txt_056 %> >>" id="clskom"  onclick="closekomm();"/>
         <br /><span style="color:#999999;">Kommentar gemmes først i forbindelse med timeregistrering.</span><br /><br />&nbsp;
        </td></tr>

    


      <input type="hidden" name="sdskKomm" id="sdskKomm" value="<%=fromdeskKomm%>">
<input type="hidden" name="lasttd" id="lasttd" value="0">
<input type="hidden" name="lastkmdiv" id="lastkmdiv" value="0">
<input type="hidden" name="lasttddaytype" id="lasttddaytype" value="0">

</form><!-- MAIN Timereg form -->



<%
'** Vis faneblad

'*** ALTID LUKKET ***'
'select case lto 
'case "xxx", "intranet - local", "dencker", "jttek"
'otf = 0
'case else
'otf = 1
'end select


'if cint(otf) = 1 then 
        
        'dvotf_vzb = "visible"
        'dvotf_dsp = ""

        dvlager_vzb = "hidden"
        dvlager_dsp = "none"

    'else

        dvotf_vzb = "hidden"
        dvotf_dsp = "none"

        'dvlager_vzb = "visible"
        'dvlager_dsp = ""

    'end if

%>


    <form>

        <input type="hidden" id="matreg_test" value="xx">
        <input type="hidden" id="matreg_lto" value="<%=lto %>" />   
        <input type="hidden" id="matreg_medid" value="<%=usemrn %>" />
        <!--<input type="hidden" id="matreg_onthefly" value="<%=otf %>" />-->
        <input type="hidden" id="matreg_matid" value="0" />
        <input type="hidden" id="matreg_jobid" name="jobid" value="0" />
        <input type="hidden" id="matreg_aktid" name="aktid" value="0" />
        <input type="hidden" id="matreg_aftid" name="aftid" value="0" />
        <input type="hidden" id="matreg_regdato_0" name="regdato" value="01-01-2002" />
        <input type="hidden" id="matregid" name="matregid" value="0" />
        <input type="hidden" id="matreg_func" name="matreg_func" value="dbopr" />
        
        
    
   
   
    <tr><td colspan="2">
        <br />

        <span id="sp_mat_forbrug" style="font-size:14px;">[+]</span> <span style="font-size:14px; font-weight:bolder;">Materiale forbrug</span><br /> [<span id="a_otf" style="color:#5582d2;">Opret "on th fly"</span>] &nbsp; [<span id="a_lager" style="color:#5582d2;">Tilføj fra lager</span>]<br />
        <span style="color:#999999;">Tilknyt materialeforbrug til denne timeregistrering.</span>&nbsp;
        </td></tr>

         

        <tr><td colspan="2" id="matreg_indlast_err" style="background-color:#ffffff; padding:4px;"></td></tr>
        <tr><td colspan="2" id="matreg_indlast" style="background-color:#ffffff; padding:4px;"></td></tr>

      

      
        <tr><td colspan="2" id="matreg_indlast_allerede" style="background-color:#ffffff; padding:4px; font-size:9px; color:#999999;"></td></tr>



    <!-- tilføj on the fly -->
       <tr><td colspan="2" valign="top">
           
           <div id="dv_mat_otf" style="position:relative; visibility:<%=dvotf_vzb%>; display:<%=dvotf_dsp%>;">
           <table>
   

     <%
         matantal = 1
         call matStFelter() %>

               </table>
           </div>
           </td>
           </tr>

          <tr id="dv_mat_otf_sb" style="visibility:<%=dvotf_vzb%>; display:<%=dvotf_dsp%>;"><td colspan="2" align=right><br /><br />
	<input type="button" id="matreg_sb" value="<%=tsa_txt_085 %> >>" />
  <br />
    <br /><!--(Kommentar bliver først gemt når timer indlæses)-->
    &nbsp;</td></tr>

        

       <tr><td colspan="2" valign="top">
            
        <div id="dv_mat_lager" style="position:relative; visibility:<%=dvlager_vzb%>; display:<%=dvlager_dsp%>;">
      <!-- tilføj fra lager -->

            <table cellpadding="2" cellspacing="0" border="0">

                <%call mat_lager_sogeKri %>

            </table>
            </div>
            <br /><br />

               <div id="dv_mat_lager_list" style="position:relative; visibility:<%=dvlager_vzb%>; display:<%=dvlager_dsp%>; height:400px; overflow-y:scroll;">
              <table cellpadding="0" cellspacing="0" border="0" width="100%">

                  <%call matregLagerheader(0) %>
                
                  </table>
            </div>
           </td>
           </tr>

   
     
	</table>


    

        </form><!-- oft / lager -->



  </div><!-- kom -->
  
  



  </div> <!-- sidediv-->
   </div>   <!-- sidediv-->








        <!-- 
                '********************************************************
                Kontakt Kørsel info div 
                
                
                
                '********************************************************
                -->
	            <div id="korseldiv" style="position:absolute; left:125px; top:300px; background-color:#ffffff; width:600px; border:10px #CCCCCC solid; padding:10px; visibility:hidden; display:none; z-index:200; height:400px; overflow:auto;">
				 
                    
                    <a href="javascript:popUp('kontaktpers.asp?func=opr&id=0&kid=<%=kid%>&rdir=treg','550','500','20','20')" target="_self" class=remnu>+ Tilføj ny kontakt</a>
				<table border=0 cellspacing=0 cellpadding=0 width=100%>
				<form action="#" method="post">
				<tr><td align=right><span id="jq_hideKpersdiv" style="color:#999999; font-size:14px;">X</span></td></tr>
				
				<tr>
				    <td><h3><%=tsa_txt_360 %>:</h3>
                    <%=tsa_txt_365 %><br /><%=tsa_txt_361 %><br /><br />

                        <select id="ko0" style="width:500px;"><option value="0">....</option></select><br /><br />

                        <h4>Du har kørt til:</h4>
                        <select id="ko1" style="width:500px;"><option value="0">....</option></select>




                        
				   </td>
				</tr>
				
				
			
				<input id="FM_sog_kpers_dist_kid" value="0" type="hidden" />
				
                    <!--
                    <input id="FM_sog_kpers_dist" name="FM_sog_kpers_dist" value="" type="text" style="width:200px;" /> 
                    &nbsp;<input id="FM_sog_kpers_but" type="button" value="<%=tsa_txt_078 %> >> " />
                   
                    <br />
                    <input id="FM_sog_kpers_dist_all" value="1" type="checkbox" /> <%=tsa_txt_359 %>
                     <input id="FM_sog_kpers_dist_more" value="1" type="checkbox" /> <%=tsa_txt_372 %>
                    -->
				
			
				<tr><td align=right>
				<%call erkmDialog()  'skal den vises
				%>
				    <input id="koKmDialog" value="<%=kmDialogOnOff%>" type="hidden" />
                    <input id="koFlt" value="0" type="hidden" /><!-- row -->
                    <input id="koFltx" value="1" type="hidden" /><!-- x dag -->
                    <!--<input id="kperfil_fundet" type="text" value="0">-->

                    <input id="indlaes_koadr_2" type="button" value=" Indlæs >> " /><br />&nbsp;
                    </td></tr>
				</form>
				</table>
				</div>
				<!--KM dialog div -->
				<%response.flush %>










 
  

	
	
	
	
	   <!---- Side indstillinger, Indstilinger for timereg.,Indlæs timer som, tal (timer),  start og slut tid  --->
	   <div id="pagesetadv" style="position:absolute; display:none; visibility:hidden; border:10px #CCCCCC solid; top:400px; left:125px; width:480px; height:295px; z-index:300; background-color:#ffffff; padding:10px;">
         
        <table cellspacing=0 cellpadding=2 border="0" bgcolor="#FFFFFF" width="100%">
         <form action="timereg_akt_2006.asp?FM_use_me=<%=usemrn%>&showakt=1&jobid=<%=jobid%>&fromsdsk=<%=fromsdsk%>&isindstillingersubmitted=1" method="post" name="timereg_indst" id="timereg_indst">
	   
       <tr bgcolor="#FFFFFF">
	   <td colspan=2 style="width:150px;"><h4><%=tsa_txt_307%></h4></td>
	   <td align=right><a href="#" id="pagesetadv_close" style="color:#999999; font-size:14px;">X</a></td></tr>
        <tr><td colspan=3>
        
		<b><%=tsa_txt_033 %>:</b><br /> 

         <%visTimerElTid = timer_ststop 'meStamdata %>

		<%if cint(visTimerElTid) <> 1 then
		chk0 = "CHECKED"
		chk1 = ""
		else
		chk1 = "CHECKED"
		chk0 = ""
		end if%>
		<input type="radio" name="FM_vistimereltid" id="FM_vistimereltid_0" value="0" <%=chk0%>> <%=tsa_txt_034 %> <input type="radio" name="FM_vistimereltid" id="FM_vistimereltid_1" value="1" <%=chk1%>> <%=tsa_txt_035 %> 
		
        <br><br />
		
		<%
		if cint(ignorertidslas) = 1 then
		ignorertidslasCHK = "checked"
		else
		ignorertidslasCHK = ""
		end if
		%>
		
		<input type="checkbox" name="FM_igtlaas" id="FM_igtlaas" value="1" <%=ignorertidslasCHK%>> <%=tsa_txt_046 %>
            <input id="FM_chkfor_ignorertidslas" name="FM_chkfor_ignorertidslas" type="hidden" value="1"/>
		
		
		<br />
        
        <%

		if cint(ignorerakttype) = 1 then
		ignorerakttypeCHK = "checked"
		else
		ignorerakttypeCHK = ""
		end if
		%>
		
		<input type="checkbox" name="FM_igakttype" id="FM_igakttype" value="1" <%=ignorerakttypeCHK%>> <%=tsa_txt_377 %> (ikke kun salgs-aktiviteter)
            <input id="FM_chkfor_igakttype" name="FM_chkfor_igakttype" type="hidden" value="1"/>
		
		
		<br />
		
		<%
        if cint(lukaktvdato) = 1 then 'sat i kontrolpanel

            %>

            	<input type="checkbox" name="" id="FM_ignJobogAktper" value="1" CHECKED DISABLED> <%=tsa_txt_286 %>
            <input type="hidden" name="FM_ignJobogAktper" id="Checkbox2" value="1"> 
            <%
        else

		if cint(ignJobogAktper) = 1 then
		ignJobogAktperCHK = "checked"
		else
		ignJobogAktperCHK = ""
		end if
		%>
		
		<input type="checkbox" name="FM_ignJobogAktper" id="FM_ignJobogAktper" value="1" <%=ignJobogAktperCHK%>> <%=tsa_txt_286 %>

        <%end if %>
         <input id="Hidden1" name="FM_chkfor_ignJobogAktper" type="hidden" value="1"/>


         <br />
		
		<%
		if cint(visSimpelAktLinje) = 1 then
		visSimpelAktLinjeCHK = "checked"
		else
		visSimpelAktLinjeCHK = ""
		end if
		%>
		
		<input type="checkbox" name="FM_visSimpelAktLinje" id="FM_visSimpelAktLinje" value="1" <%=visSimpelAktLinjeCHK%> DISABLED> <%=tsa_txt_382 %> (sættes i kontrolpanel)


         <input id="Hidden4" name="FM_chkfor_visSimpelAktLinje" type="hidden" value="1"/>
		
	   </td>
	   </tr>
             <tr><td colspan="3">
       
        <%

     'call aktBudgettjkOn_fn()
     if cint(aktBudgettjkOn) = 1 then 
            
        if bxbgtfilter_Type = "hidden" then

            %>
             <input type="checkbox" id="viskunejobmbudget" value="1" <%=viskunejobmbudgetCHK %> disabled/> <span style="color:#000000;">Vis kun job/aktiviteter med <%=fcTxt %> på, for valgte medarbejder (sættes i kontrolpanel)</span>
             <input type="hidden" value="<%=viskunejobmbudget %>" name="viskunejobmbudget" />
    
            <%
        else%>
             <input name="viskunejobmbudget" id="viskunejobmbudget" type="checkbox" value="1" <%=viskunejobmbudgetCHK %>/> <span style="color:#000000;">Vis kun job/aktiviteter med <%=fcTxt %> på, for valgte medarbejder (sættes i kontrolpanel)</span>
    
        <%end if%>
     
     <%end if 
         
         
       %>

            </td>
             </tr>


                         
	   <tr>
	   <td colspan=3 align=right><br /><br /><input type="submit" value="<%=tsa_txt_086 %> <> " style="height:18px; font-family:arial; font-size:10px;" /></td>
	   </tr>
	   	</form>
	   </table>
	   </div>
	
	
	
	<% 



    if lto = "xx" then ' slået fra i en peridoe

    if (antalJobLinier > 20 OR antalaktlinier > 200) AND session("forste") = "j" AND (sortBy <> 5 AND sortBy <> 6) then
    
                itop = 480
            ileft = 240
            iwdt = 300
            idsp = ""
            ivzb = "visible"
            iId = "tregaktmsg1"
            call sidemsgId(itop,ileft,iwdt,iId,idsp,ivzb)
            %>
			
            <b>Der er valgt mere end 20 job eller mere end 200 aktiviteter på denne liste</b><br><br />
	        Hvis der er vises mere end 20 job (<b><%=antalJobLinier %></b>) eller over 200 aktiviteter (<b><%=antalaktlinier%></b>), bliver visningen langsommere (> 10 sek.) og det er lettere at miste overblikket. 
            <br><br>
            Du bør derfor overveje at vælge færre job eller rydde op i antallet af aktiviteter.
            <br><br>
	        
	        
	</td></tr></table>
	</div>
	

<%

    end if


    end if 'xx


t = visGuide


%>



<div id="velkommen" style="width:1px; height:1px; display:none; visibility:hidden;"></div>


<%
'**** Stade indmelding ****'
if session("forste") = "j" then
call stadeindm(usemrn, 1, now)
end if
%>



<%response.flush %>


	
	
	
	
	
<!--pagehelp-->

<%
    q = 900
if showakt = 1 AND browstype_client <> "ip" AND q = 1 then

itop = 84
ileft = 655
iwdt = 120
ihgt = 0
ibtop = 140 
ibleft = 150
ibwdt = 600
ibhgt = 550
iId = "pagehelp"
ibId = "pagehelp_bread"
call sideinfoId(itop,ileft,iwdt,ihgt,iId,phDsp,phVzb,ibtop,ibleft,ibwdt,ibhgt,ibId)
%>
        <b><%=tsa_txt_278 %></b><br /><%=tsa_txt_279 %><br /><br />
		<!--<b><=tsa_txt_047 %>:</b>&nbsp;<=tsa_txt_048 %><br><br /> -->
		<b><%=tsa_txt_049 %>:</b>&nbsp;<%=tsa_txt_050 %> <br /><br />
		<b><%=tsa_txt_296 %>:</b><br />
		<%=tsa_txt_297 %>
		<br /><br />
		<b><%=tsa_txt_298 %>:</b><br />
		<%=tsa_txt_299 %>
		
		<br /><br /><b><%=tsa_txt_300 %></b><br />
		<%=tsa_txt_301 %>
		<br /><br /><b><%=tsa_txt_295 %>:</b><br />
		
		<%for f = 1 to 6
		
		select case f
		case 1
		xxfmborcl = "1px #999999 solid"
		xxfmbgcol = "#ffffff"
		xxtxt = tsa_txt_285 '"Åben. Timer ikke indlæst. (Pre-udfyldte aktiviteter, f.eks. frokost)"
		case 2
		xxfmborcl = "1px yellowgreen solid"
		xxfmbgcol = "#ffffff"
		xxtxt = tsa_txt_290 '"Åben. Timer indlæst i database."
		case 3
		xxfmborcl = "1px yellowgreen dashed"
		xxfmbgcol = "#cccccc"
		xxtxt = tsa_txt_291 '"Lukket for indtastning. Registrering er godkendt. (kan ydermere også være faktureret eller lukket via smiley-ordning)"
		case 4
		xxfmborcl = "1px #999999 dashed"
		xxfmbgcol = "#cccccc"
		xxtxt = tsa_txt_292 '"Lukket for indtastning. Er enten faktureret eller uge er lukket via smiley-ordning."
		case 5
		xxfmborcl = "1px #999999 dashed"
		xxfmbgcol = "#ffffff"
		xxtxt = tsa_txt_293 '"Lukket for indtastning. Uge er lukket via smiley ell. fast periode, men admin. brugere kan ændre."
		case 6
		xxfmborcl = "1px #cccccc dashed"
		xxfmbgcol = "#EfF3FF"
		xxtxt = tsa_txt_294 '"Tidslåst."
		end select
		
		Response.write "<input type='text' name='xx' id='xx' value='' maxlength='0' style=""background-color:"&xxfmbgcol&"; border:"&xxfmborcl&"; height:16px; width:45px; font-family:arial; font-size:10px; line-height:12px;""> "& xxtxt &"<img src='ill/blank.gif' height=3 width=1><br><br>"
								
		
		next %>
		
		
		</td>
	</tr>
    </table></div>
	
	<%end if %>
	


    <%
    response.flush
    
  
   


'** ikke længere første logon-. == Cookies må benyttes. ***
'** bruges bl.a til i kalender til valg af dato eller cookie **'
session("forste") = "n"


	
else





            itop = 550
            ileft = 20
            iwdt = 300
            idsp = ""
            ivzb = "visible"
            iId = "tregaktmsg2"
            call sidemsgId(itop,ileft,iwdt,iId,idsp,ivzb)
            %>
            			
                        
            	         <b><%=tsa_txt_347 %></b><br />
            	         <%=tsa_txt_348 %>
            	         <br /><br><b><%=tsa_txt_349 %></b><br />
            	         <%=tsa_txt_350 %>
	                    
	                    <br /><br />
            	        
	                    <%=strprojektgrpOK %>

                        <br /><br />
                        <b>Nulstil</b><br />
                        Hvis du søgningen er stoppet kan du nulstille søgningen ved at klikke her.<br />
                        <a href="timereg_akt_2006.asp?jobid=0&showakt=1" class=vmenu>Nulstil søgning >></a>
	                    <br /><br />
                       
	                   
	                    
	                   &nbsp;
	            </td></tr></table>
	            </div>
            	

            <%
       
        
end if 'projgrpOK, adgang via projgrp eller vi link til tilføj sig til job 



end if 'Session User





%>

	

<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />


<!--#include file="../inc/regular/footer_inc.asp"-->
