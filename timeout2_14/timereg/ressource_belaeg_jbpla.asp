<%
response.buffer = true
'timeA = now

    %>
         <!--#include file="../inc/connection/conn_db_inc.asp"-->
        <!--#include file="../inc/errors/error_inc.asp"-->
        <!--#include file="../inc/regular/global_func.asp"-->
        <!--#include file="../inc/regular/webblik_func.asp"-->
        <!--#include file="inc/ressource_belaeg_jbpla_inc.asp"-->
        <!--#include file="inc/isint_func.asp"-->
        <!--#include file="../inc/regular/stat_func.asp"-->

       
        <%


    call positiv_aktivering_akt_fn()

    public visAktiv, visRamme
    public strSQLa
    function hentaktiviteter(positiv_aktivering_akt_val, jid, mid, aty_sql_realhoursAkt)
                    
                
                    '** Fjerner akt der allerede er i bruge for denne medarbejder **'
                    'aktivIbrug = "AND (a.id <> 0 "

                    'strSQLaib = "SELECT aktid FROM ressourcer_md AS rmd WHERE rmd.medid = "& mid &" AND rmd.jobid = "& jid  'AND ((rmd.md >= "& startdatoMD &" AND rmd.aar = "& startdatoYY &") "& orandval &" (rmd.md <= "& slutdatoMD &" AND rmd.aar = "& slutdatoYY &")) GROUP BY aktid"
                    'response.write "strSQLaib " & strSQLaib
                    'response.flush            
        
                    'oRec4.open strSQLaib, oConn, 3
                    'while not oRec4.EOF 
                            
                    'aktivIbrug = aktivIbrug & " AND a.id <> "& oRec4("aid")

                    'oRec4.movnext
                    'wend
                    'oRec4.close                
            
                    'aktivIbrug = aktivIbrug & ")"


                  

                    aktivIbrug = ""
    
                    if cint(positiv_aktivering_akt_val) = 1 then
                    strSQLa = "SELECT navn AS aktnavn, a.id AS aid, fase FROM timereg_usejob AS tu "_
                    &" LEFT JOIN aktiviteter AS a ON (a.id = tu.aktid) "_
                    &" WHERE tu.jobid = "& jid &" AND tu.medarb = "& mid &" AND aktstatus = 1 "& aktivIbrug &" GROUP BY a.id ORDER BY fase, sortorder, navn"
                    else
                    strSQLa = "SELECT navn AS aktnavn, id AS aid, fase FROM aktiviteter AS a WHERE job = "& jid &" AND aktstatus = 1 AND ("& aty_sql_realhoursAkt &") "& aktivIbrug &" GROUP BY id ORDER BY fase, sortorder, navn"
                    end if


                    'if session("mid") = 1 then
                    'response.write strSQLa
                    'response.flush
                    'end if
                  

    end function



    function jobmedtimeruforecast(mid,vis)

    


        if len(trim(jobisWrt(mid))) <> 0 then
	    jobisWrt(mid) = jobisWrt(mid) & ")"
	    jobIdisWrt(mid) = jobIdisWrt(mid) & ")"
        else
	    jobisWrt(mid) = ""
	    jobIdisWrt(mid) = ""
        end if

    '********************************************************************************
    '*** finder job, medarb og kundeoplysninger, Real timer bliver fundet senere ***'
    '*** +Aktive job på personlig aktiv listen *************************************'
    '********************************************************************************

  

    strSQLt = "SELECT jobnavn, m.mnavn, m.mnr, tmnr, j.id AS jobid, j.jobnr, m.medarbejdertype, m.mid, "
    strSQLt = strSQLt & "SUM(t.timer) AS realtimer, "
    strSQLt = strSQLt & " k.kkundenavn, k.kkundenr, m.init, m.forecaststamp, jobans1, jobans2, "_
	&" m2.mnavn AS m2mnavn, m2.init AS m2init, m2.mnr AS m2mnr, m3.mnavn AS m3mnavn, m3.init AS m3init, m3.mnr AS m3mnr "_
    &" FROM timer t  "
	
	
	strSQLt = strSQLt &" LEFT JOIN medarbejdere m ON (m.mid = " & mid &")"_
	&" LEFT JOIN job j ON (j.jobnr = t.tjobnr AND "& jobSQLkri &" AND (j.risiko > -1 OR j.risiko = -3)) AND j.jobnavn IS NOT NULL"
	
	
    strSQLt = strSQLt &" LEFT JOIN kunder k ON (kid = j.jobknr)"_
	&" LEFT JOIN medarbejdere m2 ON (m2.mid = j.jobans1)"_
	&" LEFT JOIN medarbejdere m3 ON (m3.mid = j.jobans2)"
	
  

    strSQLt = strSQLt &" WHERE t.tmnr = "& mid &" "& jobisWrt(mid)
   
    strSQLt = strSQLt &" AND j.jobstatus = 1 AND t.tdato BETWEEN '"& tStDatoJobUforecast &"' AND '"& tSlDatoJobUforecast &"'"
	strSQLt = strSQLt &" GROUP BY t.tjobnr, t.tmnr ORDER BY t.tjobnavn"
    
    'if lto = "essens" then
    'Response.write "<br>Job med timre uden forecast: <b>"& strSQLt &"</b><br><br>"
	'Response.flush
    'end if
    
    
    jobMrealtimerPa = jobIdisWrt(mid)
    jobMrealtimerPa = replace(jobMrealtimerPa, ")", "")
    oRec2.open strSQLt, oConn, 3
    while not oRec2.EOF 

            if oRec2("realtimer") <> 0 then

                        medarbKundeoplysX(x, 0) = oRec2("jobid")
	                    medarbKundeoplysX(x, 1) = oRec2("jobnr")
                        medarbKundeoplysX(x, 2) = oRec2("jobnavn")
                        medarbKundeoplysX(x, 3) = oRec2("mid")
                        medarbKundeoplysX(x, 4) = oRec2("mnavn")
    
	                    medarbKundeoplysX(x, 5) = oRec2("kkundenavn")
	                    medarbKundeoplysX(x, 6) = oRec2("kkundenr")
	                    medarbKundeoplysX(x, 7) = oRec2("mnr")
	                    medarbKundeoplysX(x, 8) = oRec2("init")
	
	       
	                    medarbKundeoplysX(x, 9) = "2002-01-01" 'oRec("forecaststamp")
	                    medarbKundeoplysX(x, 10) = 0
	                    resid = -1
	       
	
	                    medarbKundeoplysX(x, 11) = oRec2("jobans1")
	                    medarbKundeoplysX(x, 12) = oRec2("jobans2")
	
	                    if len(trim(oRec2("m2mnavn"))) <> 0 then
                        medarbKundeoplysX(x, 13) = oRec2("m2mnavn") &" ("&oRec2("m2mnr")&")"
                        medarbKundeoplysX(x, 14) = oRec2("m2init") 
                        else
                        medarbKundeoplysX(x, 13) = ""
                        medarbKundeoplysX(x, 14) = ""
                        end if
    
                        if len(trim(oRec2("m3mnavn"))) <> 0 then
                        medarbKundeoplysX(x, 15) = oRec2("m3mnavn") &" ("&oRec2("m3mnr")&")"
                        medarbKundeoplysX(x, 16) = oRec2("m3init")
                        else
                        medarbKundeoplysX(x, 15) = ""
                        medarbKundeoplysX(x, 16) = ""
                        end if 
            

            
                        medarbKundeoplysX(x, 17) = 0
                        medarbKundeoplysX(x, 18) = ""

                        jobMrealtimerPa = jobMrealtimerPa & " AND jobid <> '"&  oRec2("jobid") & "'"
    
                    x = x + 1

            end if
    oRec2.movenext
    wend
    oRec2.close



    '***** Henter job der er med på aktivjoblisten for den enkelte medarbejder ****'
     if len(trim(jobMrealtimerPa)) <> 0 then
     jobMrealtimerPa = jobMrealtimerPa & ")"
     end if
    'jobIswrtKri = jobisWrt(mid)
    strSQLtua = "SELECT aktid, medarb AS mid, m.mnavn, jobid, j.jobnr, j.jobnavn, "_
    &" kkundenavn, kkundenr, m.mnr, m.init, jobans1, jobans2, "_
    &" m2.mnavn AS m2mnavn, m2.init AS m2init, m2.mnr AS m2mnr, m3.mnavn AS m3mnavn, m3.init AS m3init, m3.mnr AS m3mnr FROM timereg_usejob AS tu "
    
    strSQLtua = strSQLtua &" LEFT JOIN medarbejdere m ON (m.mid = " & mid &")"_
	&" LEFT JOIN job j ON (j.id = tu.jobid AND (j.risiko > -1 OR j.risiko = -3))"
	
    strSQLtua = strSQLtua &" LEFT JOIN kunder k ON (kid = j.jobknr)"_
	&" LEFT JOIN medarbejdere m2 ON (m2.mid = j.jobans1)"_
	&" LEFT JOIN medarbejdere m3 ON (m3.mid = j.jobans2)"_
    &" WHERE medarb = "& mid &" "& jobMrealtimerPa & " AND forvalgt = 1 AND aktid = 0 AND "& jobSQLkri &" AND (j.risiko > -1 OR j.risiko = -3) AND jobstatus = 1"
    
    'if lto = "essens" then
    'Response.write strSQLtua
    'Response.flush
    'end if
    oRec2.open strSQLtua, oConn, 3
    while not oRec2.EOF 

            

                        medarbKundeoplysX(x, 0) = oRec2("jobid")
	                    medarbKundeoplysX(x, 1) = oRec2("jobnr")
                        medarbKundeoplysX(x, 2) = oRec2("jobnavn")
                        medarbKundeoplysX(x, 3) = oRec2("mid")
                        medarbKundeoplysX(x, 4) = oRec2("mnavn")
    
	                    medarbKundeoplysX(x, 5) = oRec2("kkundenavn")
	                    medarbKundeoplysX(x, 6) = oRec2("kkundenr")
	                    medarbKundeoplysX(x, 7) = oRec2("mnr")
	                    medarbKundeoplysX(x, 8) = oRec2("init")
	
	       
	                    medarbKundeoplysX(x, 9) = "2002-01-01" 'oRec("forecaststamp")
	                    medarbKundeoplysX(x, 10) = 0
	                    resid = -1
	       
	
	                    medarbKundeoplysX(x, 11) = oRec2("jobans1")
	                    medarbKundeoplysX(x, 12) = oRec2("jobans2")
	
	                    if len(trim(oRec2("m2mnavn"))) <> 0 then
                        medarbKundeoplysX(x, 13) = oRec2("m2mnavn") &" ("&oRec2("m2mnr")&")"
                        medarbKundeoplysX(x, 14) = oRec2("m2init") 
                        else
                        medarbKundeoplysX(x, 13) = ""
                        medarbKundeoplysX(x, 14) = ""
                        end if
    
                        if len(trim(oRec2("m3mnavn"))) <> 0 then
                        medarbKundeoplysX(x, 15) = oRec2("m3mnavn") &" ("&oRec2("m3mnr")&")"
                        medarbKundeoplysX(x, 16) = oRec2("m3init")
                        else
                        medarbKundeoplysX(x, 15) = ""
                        medarbKundeoplysX(x, 16) = ""
                        end if 
            

            
                        medarbKundeoplysX(x, 17) = 0
                        medarbKundeoplysX(x, 18) = ""

           
    
                    x = x + 1

         
    oRec2.movenext
    wend
    oRec2.close


     
    'jobisWrt(mid) = " AND (t.tjobnr <> 0 "
    
    end function




if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	%>
	
	<%
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	'*** Akt. typer der tælles med i ugereg. ***'
	call akttyper2009(2)
	
	
	func = request("func")
	
	thisfile = "res_belaeg"
	'Response.write "func: " & func
	
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if


    if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
    progrp = 0
	end if



        if len(request("FM_medarb")) <> 0 then 'form submitted

        if len(trim(request("FM_expvisreal"))) <> 0 then
        expvisreal = request("FM_expvisreal")
        else
        expvisreal = 0
        end if

        response.cookies("tsa")("respreal") = expvisreal

        else
        
        if request.cookies("tsa")("respreal") <> "" then
        expvisreal = request.cookies("tsa")("respreal")
        else
        expvisreal = 0
        end if


        end if


        if cint(expvisreal) = 0 then
         expvisrealCHK = ""
        else
        expvisrealCHK = "CHECKED"
        end if

	'Response.Write "medid first: "& left(request("FM_medarb"), 1)
	'Response.end
	
	'*** Rettigheder på den der er logget ind **'
	medarbid = session("mid")
	
	vis = 0
	if len(trim(request("vis"))) <> 0 then
    vis = request("vis")
    else
    vis = 0
    end if

    '**** vis som timer ell. procent ***'
    if len(trim(request("FM_showasproc"))) <> 0 then
        showasproc = request("FM_showasproc")
    else
        
        if request.cookies("tsa")("showasproc") <> "" then
        showasproc = request.cookies("tsa")("showasproc")
        else
        showasproc = 0
        end if

    end if 

    if cint(showasproc) = 1 then
        showasproc0Sel = ""
        showasproc1Sel = "SELECTED"
    else
         showasproc0Sel = "SELECTED"
        showasproc1Sel = ""
    end if
    response.cookies("tsa")("showasproc") = showasproc 

    
    '** Job uden forecast
     if len(request("FM_medarb")) <> 0 then 'form submitted
    
    if len(trim(request("FM_vis_job_u_fc"))) <> 0 AND request("FM_vis_job_u_fc") <> 0 then
    vis_job_u_fc = 1 'request("FM_vis_job_u_fc")
    else
    vis_job_u_fc = 0
    end if

    else 

        if request.cookies("resbel")("ufc") <> "" then
        vis_job_u_fc = request.cookies("resbel")("ufc")
        else
        vis_job_u_fc = 0
        end if
    
    end if

 

    if cint(vis_job_u_fc) = 0 then
    vis_job_u_fcCHK = ""
    else
    vis_job_u_fcCHK = "CHECKED"
    end if


      '** Vis simpel (uden susbtotaler)
     if len(request("FM_medarb")) <> 0 then 'form submitted
    
    if len(trim(request("FM_vis_simpel"))) <> 0 then
    vis_simpel = request("FM_vis_simpel")
    else
    vis_simpel = 0
    end if

    else 

        if request.cookies("resbel")("vsi") <> "" then
        vis_simpel = request.cookies("resbel")("vsi")
        else
        vis_simpel = 0
        end if
    
    end if


    if len(trim(vis_simpel)) <> 0 then
    vis_simpel = vis_simpel
    else
    vis_simpel = 0
    end if


        'vis_simpel = 0
        vis_simpelCHK0 = ""
        vis_simpelCHK1 = ""
        vis_simpelCHK2 = ""


        'response.write "vis_simpel: " & vis_simpel
        'response.flush

    select case cint(vis_simpel) 
    case 1
        vis_simpelCHK1 = "SELECTED"
    case 2
        vis_simpelCHK2 = "SELECTED"
    case else
        vis_simpelCHK0 = "SELECTED"
    end select


    '** Job hvor forecast er overskreddet
    if len(request("FM_medarb")) <> 0 then 'form submitted
    
    if len(trim(request("FM_vis_job_fc_neg"))) <> 0 AND request("FM_vis_job_fc_neg") <> 0 then
    vis_job_fc_neg = 1 'request("FM_vis_job_fc_neg")
    else
    vis_job_fc_neg = 0
    end if

    else 

        if request.cookies("resbel")("fcneg") <> "" then
        vis_job_fc_neg = request.cookies("resbel")("fcneg")
        else
        vis_job_fc_neg = 0
        end if
    
    end if


    if cint(vis_job_fc_neg) = 0 then
    vis_job_fc_negCHK = ""
    else
    vis_job_fc_negCHK = "CHECKED"
    end if

    response.cookies("resbel")("fcneg") = vis_job_fc_neg
    response.cookies("resbel")("ufc") = vis_job_u_fc
    response.cookies("resbel")("vsi") = vis_simpel

    if len(request("FM_medarb")) <> 0 then 'OR func = "export"
	
	    if left(request("FM_medarb"), 1) = "0" then
	        
            if media <> "print" then
	        thisMiduse = replace(request("FM_medarb_hidden"), "-", "0")
    	    else
    	    thisMiduse = replace(request("FM_medarb"), "-", "0")
    	    end if
    	
    	intMids = split(thisMiduse, ", ")
	    else
	    thisMiduse = replace(request("FM_medarb"), "-", "0")
	    'Response.write "<br>thisMiduse "& thisMiduse
        intMids = split(thisMiduse, ", ")
	    end if
	
	else
	    
	    thisMiduse = session("mid") 
	    intMids = split(thisMiduse, ", ")
	   
	end if


    'Response.Write "FM_medarb:" &  request("FM_medarb")
	

	media = request("media")
	'print = request("print")
	
	'*** Periode ***
	dim timerTotalTotal, maanedY, medarbTotalTimer, medarbTimerReal, realTotalTotal
    dim timerMJobTotal, realMJobTotal, fc_lastjobnrWrt, fc_lastaktidWrt

    
	
	if len(request("periodeselect")) <> 0 then
	periodeSel = request("periodeselect")
	'periodeShow = ""& periodeSel &" Måneder frem"
	perBeregn = periodeSel - 1
	else
	
	        if len(request.cookies("resbel")("periodesel")) <> 0 then
	        periodeSel = request.cookies("resbel")("periodesel")
	        'periodeShow = ""& periodeSel &" Måneder frem"
	        perBeregn = periodeSel - 1
	        else
	        periodeSel = 6
	        'periodeShow = "6 Måneder frem"
	        perBeregn = 5
	        end if
	end if


	
	select case periodeSel
	'case 1
	'leftwdt = 35
	'rdimnb = 31
	'periodeShow = periodeShow & " dage"
	case 3
	stDatoDag = 1 
	rdimnb = 15
	leftwdt = 350
	periodeShow = periodeShow & " 3 måneder frem (uge/uge)"
	case 6
	stDatoDag = 14 '** for at være sikker på at der ikke bliver skrvet en uge ind, der hører til forrige måned.
	rdimnb = 7
	leftwdt = 500
	periodeShow = "6 Måneder frem"
	case 12
	stDatoDag = 14
	rdimnb = 13
	leftwdt = 850
	periodeShow = "12 Måneder frem"
	end select
	
    redimNo = 9500

	Redim timerTotalTotal(redimNo), maanedY(redimNo), medarbTotalTimer(redimNo), medarbTimerReal(redimNo), realTotalTotal(redimNo) 
    Redim timerMJobTotal(redimNo), realMJobTotal(redimNo), fcGrandtotal(redimNo), realGrandtotal(redimNo), jobTimerGrandGrand(redimNo), jobBgtGrandGrand(redimNo)
    Redim procMJobTotal(redimNo), procTotalTotal(redimNo),  medarbTotalProc(redimNo), fpGrandtotal(redimNo) 

    realignPerGrandGrand = 0
    realGrandGrand = 0
    fcignPerGrandGrand = 0 
    fcGrandGrand = 0
    jobTimerGrandGrand = 0
    jobBgtGrandGrand = 0
	
	response.cookies("resbel")("periodesel") = periodeSel
	
	
	'************** Indlæser variable til sidevisning *************

	'*** Hvis plus eller minus knap er brugt ***
	if len(request("selectplus")) <> 0 then
	selectplus = request("selectplus")
	else
	selectplus = 0
	end if
	
	if len(request("selectminus")) <> 0 then
	selectminus = request("selectminus")
	else
	selectminus = 0
	end if
	
	
        'response.write "request "& request("mdselect") & "minus: "& selectminus 

	
		if len(request("mdselect")) <> 0 then
				monththis = request("mdselect")
				
				if selectminus <> 0 then
				monththis = monththis - 1
				end if
				
				if selectplus <> 0 then
				monththis = monththis + 1
				end if
				
				
				select case monththis 
				case 0
				monththis = 12
				yearthis = request("aarselect")
				yearthis = yearthis - 1
				case 13
				monththis = 1
				yearthis = request("aarselect")
				yearthis = yearthis + 1
				case else
				yearthis = request("aarselect")
				end select

                response.cookies("resbel")("mm") = monththis
                response.cookies("resbel")("yyyy") = yearthis
				
			else

                
                if request.cookies("resbel")("mm") <> "" then


                 monththis = request.cookies("resbel")("mm")
                 yearthis = request.cookies("resbel")("yyyy")

                else
               

			    monththis = month(now)
			    yearthis = year(now)

                end if
		
        	end if
	
	
	
	jobSlutTEMP = dateadd("m", perBeregn, "1/"& monththis &"/"& yearthis) '** år ?
	
	monthUse = datepart("m", jobSlutTEMP)
	yearUse = datepart("yyyy", jobSlutTEMP)
	
	call dageimd(monthUse,yearUse)	
	mthDaysUse = mthDays
	
   select case periodeSel
   case 3
    jobStartKri = yearthis&"/"&monththis&"/"&stDatoDag

        dayOfWeek = datePart("w",jobStartKri,2,2)
        
        '1 = (dayOfWeek - d) 
        d = (dayOfWeek - 1)
        firstdayOfFirstWeek = dateAdd("d", -d, jobStartKri)
        if month(firstdayOfFirstWeek) <> monththis then
        firstdayOfFirstWeek = dateAdd("d", 7, firstdayOfFirstWeek)
        end if

        firstThursdayOfFirstWeek = firstdayOfFirstWeek 'dateAdd("d", +3, firstdayOfFirstWeek)

        jobStartKri = firstThursdayOfFirstWeek 'firstdayOfFirstWeek 'year(firstdayOfFirstWeek) &"-"& month(firstdayOfFirstWeek) &"-"& year(firstdayOfFirstWeek)

    case else
	jobStartKri = yearthis&"/"&monththis&"/"&stDatoDag
    end select

	jobSlutKri = yearUse&"/"&monthUse&"/"&mthDaysUse
	
     '*** Bruges til at tjekke det nøjagtige normtoimer pr. medarb. ved tilføje ny linie.
	datoInterval = datediff("d",jobStartKri,jobSlutKri,2,2)
    
    '**** Tilføjer en md hvis der er en uge går ind i ny måned.
    if periodeSel = 3 then
    jobSlutKri = dateAdd("d", 1, jobSlutKri)
    'jobSlutKri = dateAdd("d")	
    jobSlutKri = year(jobSlutKri)&"/"&month(jobSlutKri)&"/"&day(jobSlutKri)
    end if

     'response.write "jobSlutKri: "& jobSlutKri

	daynow = day(now)&"/"&month(now)&"/"&year(now)
	
	
	if len(trim(request("FM_jobsel"))) <> 0 then 		
	jobidsel = request("FM_jobsel")
	response.cookies("resbel")("jobidsel") = jobidsel
	else
	
	    if len(request.cookies("resbel")("jobidsel")) <> 0 then
		jobidsel = request.cookies("resbel")("jobidsel")
		else
		jobidsel = 0
		end if
	
	end if
	
	if cint(jobidsel) <> 0 then
	jobSQLkri = " j.id = " & jobidsel
	else
	jobSQLkri = " j.id <> 0 "
    'jobSQLkri = " m.mid <> 0 " 'hvis ingen job er ibrug må der ikke tjekkes på job id i wh clause da listen så bliver tom.
	end if

    'Response.Write "jobidsel: " & jobidsel


    if len(trim(request("FM_sog"))) <> 0 then
    sogTxt = request("FM_sog")
    else
        if trim(request.cookies("tsa")("sogval")) <> "" AND len(trim(request("mdselect"))) = 0 then
        sogTxt = request.cookies("tsa")("sogval")
        else
        sogTxt = ""
        end if
    end if

    
    
    if len(trim(sogTxt)) <> 0 then
    jobSogVal = sogTxt
    jobSQLkri = ""

                if instr(jobSogVal, ",") <> 0 then '** Komma **'
	            
                    jobSQLkri = " jobnr = 0 "
	                jobSogValuse = split(jobSogVal, ",")
	                for i = 0 TO UBOUND(jobSogValuse)
	                    jobSQLkri = jobSQLkri & " OR jobnr = '"& jobSogValuse(i) &"'"   
	                next
	    
	                jobSQLkri = "("&jobSQLkri&")"
	    
	            else


                    if instr(jobSogVal, "-") <> 0 then '** Interval
	                jobSogValuse = split(jobSogVal, "-")
	                jobSQLkri = "(jobnr >= '"& trim(jobSogValuse(0)) &"' AND jobnr <= '" & trim(jobSogValuse(1)) & "')"   
	                else 
                    jobSQLkri = "(jobnavn LIKE '" & jobSogVal &"%' OR jobnr = '"& jobSogVal &"')"
                    end if

	            end if
	
	
	Response.cookies("tsa")("sogval") = sogTxt	
    end if

		
	if len(trim(request("viskunjobiper"))) <> 0 then
	viskunjobiper = request("viskunjobiper")

	else
	    
	    if len(request.cookies("resbel")("viskunjobiper")) <> 0 then
		viskunjobiper = request.cookies("resbel")("viskunjobiper")
		else
		viskunjobiper = 0
		end if
	
	end if 
	

   response.cookies("resbel")("viskunjobiper") = viskunjobiper

      

    
     if len(trim(request("FM_hideEmptyEmplyLines"))) <> 0 then
             hideEmptyEmplyLines = request("FM_hideEmptyEmplyLines")
    else

        if len(request.cookies("resbel")("hideEmptyEmplyLines")) <> 0 AND len(trim(request("mdselect"))) = 0 then
		hideEmptyEmplyLines = request.cookies("resbel")("hideEmptyEmplyLines")
		else
		hideEmptyEmplyLines = 0
        
		end if
          
    end if

   
    if cint(hideEmptyEmplyLines) = 1 then
    hideEmptyEmplyLinesCHK = "CHECKED"
    else
    hideEmptyEmplyLinesCHK = ""
    end if

    response.cookies("resbel")("hideEmptyEmplyLines") = hideEmptyEmplyLines

    

      
         

    if len(trim(request("mdselect"))) <> 0 OR len(trim(request("FM_vis_kolonne_simpel"))) <> 0 then
           if len(trim(request("FM_vis_kolonne_simpel"))) <> 0 then
            vis_kolonne_simpel = request("FM_vis_kolonne_simpel")
           else
            vis_kolonne_simpel = 0
            end if
    else

        if len(request.cookies("resbel")("vis_kolonne_simpel")) <> 0 then
        vis_kolonne_simpel = request.cookies("resbel")("vis_kolonne_simpel")
		else
            vis_kolonne_simpel = 0
		end if
          
    end if

 

    response.cookies("resbel")("vis_kolonne_simpel") = vis_kolonne_simpel



    if len(trim(request("mdselect"))) <> 0 OR len(trim(request("FM_vis_kolonne_simpel_akt"))) <> 0 then
           if len(trim(request("FM_vis_kolonne_simpel_akt"))) <> 0 then
            vis_kolonne_simpel_akt = request("FM_vis_kolonne_simpel_akt")
           else
            vis_kolonne_simpel_akt = 0
            end if
    else

        if len(request.cookies("resbel")("vis_kolonne_simpel_akt")) <> 0 then
        vis_kolonne_simpel_akt = request.cookies("resbel")("vis_kolonne_simpel_akt")
		else
         vis_kolonne_simpel_akt = 0
		end if
          
    end if

   
    'response.write "vis_kolonne_simpel_akt: "& vis_kolonne_simpel_akt

    response.cookies("resbel")("vis_kolonne_simpel_akt") = vis_kolonne_simpel_akt



     '*** Licens indstillinger (overruler) ****'
          select case lto
            case "kejd_pb", "xintranet - local"
        
            vis_kolonne_simpelDisabled = "DISABLED"
            vis_kolonne_simpel_aktDisabled = "DISABLED"

            case else
		  
            vis_kolonne_simpelDisabled = ""
            vis_kolonne_simpel_aktDisabled = ""
            end select



     select case lto
            case "kejd_pb", "xintranet - local"
            vis_kolonne_simpel = 1
            case else
		    vis_kolonne_simpel = vis_kolonne_simpel
            end select


        select case lto
            case "kejd_pb", "xintranet - local"
            vis_kolonne_simpel_akt = 1
            case else
		    vis_kolonne_simpel_akt = vis_kolonne_simpel_akt
            end select

    
        
     if cint(vis_kolonne_simpel) = 1 then
    vis_kolonne_simpelCHK = "CHECKED"
    else
    vis_kolonne_simpelCHK = ""
    end if




     if cint(vis_kolonne_simpel_akt) = 1 then
    vis_kolonne_simpel_aktCHK = "CHECKED"
    else
    vis_kolonne_simpel_aktCHK = ""
    end if


    select case vis_kolonne_simpel_akt
    case 1
    visAktiv = 0
    case else
    visAktiv = 1
    end select 

    select case vis_kolonne_simpel
    case 1
    visRamme = 0
    case else
    visRamme = 1
    end select 
    
        
             



    '** Tjekekr variable ****'
    'response.write "vis_simpel: "& vis_simpel & " -- vis: "& vis & " -- media = "& media & " -- func: "& func
    'response.write  " -- viskunjobiper: "& viskunjobiper 
    'response.write  " -- hideEmptyEmplyLines: "&  hideEmptyEmplyLines
    'response.write  " -- vis_kolonne_simpel: "&  vis_kolonne_simpel 
    'response.write " -- vis_kolonne_simpel_akt: "& vis_kolonne_simpel_akt & " -- FM_vis_job_u_fc:" & vis_job_u_fc & " -- FM_vis_job_fc_neg:" & vis_job_fc_neg

   






	
	SELECT case viskunjobiper 
    case 1
	viskunjobiper1 = "CHECKED"
	viskunjobiper0 = ""
    'viskunjobiper2 = ""
	jobDatoKri = " AND jobslutdato >= '"& jobStartKri &"' "
   
	case else
	jobDatoKri = ""
	viskunjobiper1 = ""
	viskunjobiper0 = "CHECKED"
    'viskunjobiper2 = ""
	end select
	
	'if len(request("showonejob")) <> 0 then
	'showonejob = request("showonejob")
	'else
	showonejob = 0
	'end if
	
	
	

	
	if (level <= 2 OR level = 6) then 
	whSQLkri = " m.mid <> 0 "
	multiUse = "multiple"
	
	call licensStartDato()
	
	if licensklienter < 20 then
	    if licensklienter < 5 then
	    szi = 5
	    else
	    szi = 10
	    end if
	else
	    if licensklienter > 36 then
	    szi = 20
	    else
	    szi = 15
	    end if
	end if
	
	else
	szi = 1
	multiUse = ""
	whSQLkri = " m.mid = " & medarbSel
	end if
	
	
	'Response.Write "her" & medarbSel & "<br>"
	
	'if medarbSel = "0" Then
	'strMedidSQL = " m.mid <> 0 "
	'strMedidSel = 0
	'else
	 '   strMedidSQL = " (m.mid = 0"
	 '   medarbSelArr = split(trim(medarbSel), ",")
	 '   for m = 0 to UBOUND(medarbSelArr)
	 '   strMedidSQL = strMedidSQL & " OR m.mid = " & trim(medarbSelArr(m))
	 '   strMedidSel = strMedidSel & "#"& trim(medarbSelArr(m)) &"#"
	 '   next
	 '   strMedidSQL = strMedidSQL & ")"
     'end if
	


    '*** finder udfra valgte projektgrupper og medarbejdere
	'medarbSQlKri 
	'kundeAnsSQLKri
	
			    for m = 0 to UBOUND(intMids)
			    
			     if m = 0 then
			     medarbSQlKri = "(m.mid = " & intMids(m)
			    
			     else
			     medarbSQlKri = medarbSQlKri & " OR m.mid = " & intMids(m)
			    
			     end if
			     
			    next
			    
			    medarbSQlKri = medarbSQlKri & ")"
			    
	'medarbSQlKri = strMedidSQL    
	
	
	response.cookies("resbel").expires = date + 30
	
	'Response.write strMedidSQL & "<br>"& strMedidSel
	'Response.flush
	
	
	
	medarbIswrt = ""
	if func = "redalle" then
	
						
						 'firstLoop = 1
						'*** Datoer ****
						datoer = split(request("FM_dato"), ",")
						
                        'Response.Write request("FM_jobid") & "<hr>"
                        'Response.Write request("FM_aktid") & "<hr>"
						'Response.flush

						tmedid = split(request("FM_medarbid"), ",")
						tjobid = split(request("FM_jobid"), ",")
                        taktid = split(request("FM_aktid"), ",")
                        taktid_old = split(request("FM_aktid_old"), ",")
                      
                        
						
						tTimertildelt_temp = replace(request("FM_timer"), ", #,", ";")
						tTimertildelt_temp2 = replace(tTimertildelt_temp, ", #", "")
						
						'Response.write request("FM_jobid") & "<br>"
						'Response.write tTimertildelt_temp2
						'Response.flush
						
						tTimertildelt = split(tTimertildelt_temp2, ";") 'split(request("FM_timer"), ";")

                        select case periodeSel
                        case 3 
		                mdorweek = 1
                        case else 
                        mdorweek = 0
                        end select				

						'*** Validerer ***
						antalErr = 0
						for y = 0 to UBOUND(tTimertildelt)
						
							
							call erDetInt(SQLBless(trim(tTimertildelt(y))))
							if isInt > 0 then
								antalErr = 1
								errortype = 67
								'useleftdiv = "t"
								%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
								call showError(errortype)
								response.end
							end if	
							isInt = 0
						
						next


                       
                      			
						for y = 0 to Ubound(datoer)
						
						'Response.write "her<br>"
                        if y > 0 then
                                                
                        'if tmedid(y-1) <> tmedid(y) OR taktid_old(y-1) <> taktid_old(y) OR tjobid(y-1) <> tjobid(y) then 
                        '            firstLoop = 1
						'end if

                        end if


						if cint(instr(medarbIswrt, "#"&tmedid(y)&"#")) = 0 then
								
                                '*** Opdaterer medarbejder TimeStamp ***
								medarbStamp = year(now)&"/"&month(now)&"/"&day(now)
								strSQlmed = "UPDATE medarbejdere SET forecaststamp = '"& medarbStamp &"' WHERE mid = "& tmedid(y)		
								
								oConn.execute(strSQLmed)
								
								medarbIswrt = medarbIswrt & "#"& tmedid(y)&"#,"
						
						end if
							
							
												
												monththis = datepart("m", datoer(y), 2,2)
												'monththis = datepart("m", datoer(y))
												yearthis = datepart("yyyy", datoer(y))
												weekThis = datepart("ww", datoer(y), 2,2)
												
                                                'Response.write "datoer(y): "& datoer(y) & " tTimertildelt(y) : "& tTimertildelt(y) &"<br>"
												
                                            

													'*** Opretter nye records ***
												
													if len(trim(tTimertildelt(y))) <> 0 then
													
													
															if tTimertildelt(y) <> 0 then
																
																'*** Renser ud i tabellen 
                                                                 
																select case periodeSel
																case 3
																strSQLperKri = " AND uge = "& weekThis
                                                                'weekFaktor = 1 'gns antal uger pr. md
                                                                'normPer = 1
																case 6,12
                                                                'normPer = 30
														        'weekFaktor = 4 'gns antal uger pr. md
                                                                strSQLperKri = " AND md = "& monththis
																end select
																
																strSQLdel = "DELETE FROM ressourcer_md WHERE jobid = "& tjobid(y) &" AND medid = "& tmedid(y) &" AND aktid = "& taktid_old(y) &""_
																&"  AND aar =  "& yearthis & strSQLperKri
															    
															    'AND md = "& monththis &"
															    'Response.Write "<br>periodeSel"& periodeSel &" strSQLdel:" & strSQLdel &"<br>"
																'Response.flush
                                                                oConn.execute strSQLdel
															    
                                                                jan1 = "1-1-"& yearthis
                                                                nu = dateAdd("ww", weekThis-1, jan1)
                                                                'firstdayoffWeek = dateadd("d",(-datepart("w",jan1)),nu)
                                                                ugeDag = datePart("w",nu,2,2)

																if cint(showasproc) <> 1 then
																
                                                                    if len(trim(tTimertildelt(y))) <> 0 then
                                                                    timerThisMD = tTimertildelt(y) 'sqlbless(tTimertildelt(y))
                                                                    else
                                                                    timerThisMD = 0
                                                                    end if
                                                                   

                                                                    call normtimerPer(tmedid(y), nu, 6, 1)

                                                                    if len(trim(ntimPer)) <> 0 AND ntimPer > 0 then
                                                                    ntimPer = cint(ntimPer)
                                                                    else
                                                                    ntimPer = 1
                                                                    end if

                                                                    procThisMD = (timerThisMD/ntimPer) * 100	
                                                                    procThisMD = replace(formatnumber(procThisMD, 0), ",",".")
                                                                
                                    
                                                                else
									                            
                                                                    if len(trim(tTimertildelt(y))) <> 0 then
                                                                    procThisMD = tTimertildelt(y) 'sqlbless(tTimertildelt(y))
                                                                    else
                                                                    procThisMD = 1
                                                                    end if
                                                                    
                                                                    

                                                                
                                                                    call normtimerPer(tmedid(y), nu, 6, 1)
                                                                    timerThisMD = replace(formatnumber((ntimPer * (procThisMD/100)), 0), ",",".")						    
                                                                        'nTimerPerIgnHellig

                                                                     'response.write "timerThisMD: "& timerThisMD & " nTimerPerIgnHellig: "& nTimerPerIgnHellig &"<br>"

                                                                end if

                                                                'response.write "<br>dt: "& nu &"  ugeDag: "& ugeDag & " nTimerPerIgnHellig: "& nTimerPerIgnHellig & " formatnumber(tTimertildelt(y), 0): "& formatnumber(tTimertildelt(y), 0)

                                                                if tjobid(y) <> 0 then

                                                                '*** Renser ud i tabellen 
																select case periodeSel
																case 3
                                                                timerThisMD = replace(timerThisMD, ".", "")
                                                                timerThisMD = replace(timerThisMD, ",", ".")
                                                                procThisMD = replace(procThisMD, ".", "")
                                                                procThisMD = replace(procThisMD, ",", ".")
															   
                                                                strSQL = "INSERT INTO ressourcer_md"_
																& " (jobid, medid, md, aar, timer, uge, aktid, proc) VALUES ("_
																&" "& tjobid(y) &", "& tmedid(y) &", "& monththis &", "& yearthis &", "& timerThisMD &", "& weekThis &", "& taktid(y) &", "& procThisMD &")" 
																
                                                                
                                                                'if lto = "kejd_pb" then
                                                                'Response.Write "<br>strSQL A"& strSQL & "<br>"
                                                                'Response.flush
                                                                'end if
                                                                oConn.execute strSQL

																case 6,12
                                                                
                                                               ' if cint(firstLoop) = 1 OR dateDiff("m", thisDateGlobal, datoer(y), 2,2) >= 1 then 'first loop
                                                                firstOfmonth = "1-"& month(datoer(y)) &"-"& year(datoer(y))
                                                                'thisDateGlobal = firstOfmonth
                                                                'firstLoop = 0
                                                                'else
                                                                'firstOfmonth = thisDateGlobal
                                                                
                                                                'end if
                                                                'thisWeek53 = datePart("ww", firstOfmonth, 2,2)
                                            
                                                                thisMth = datepart("m", firstOfmonth, 2,2)
                                                                
                                    
                                                                nextMonth = dateAdd("m", 1, firstOfmonth)
                                                                nextMonthFirst = "1-"& month(nextMonth) & "-"& year(nextMonth)   
                                                                lastDInMonth = dateAdd("d", -1, nextMonthFirst) 'for at sikere vi kommer tilbage midt i ugen før. Der er den sidste uge i måneden før.
                                                                
                                                                dayOfLastWeekThis = datePart("w",lastDInMonth,2,2) - 1
                                                                firstdayOfLastWeekThis = dateAdd("d", -dayOfLastWeekThis, lastDInMonth)
                                    
                                                                dividerThisMth = datediff("w", firstOfmonth, firstdayOfLastWeekThis,2,2)
                                                        
                                                               

                                                                dayOfWeek = datePart("w",firstOfmonth,2,2)
                                                                d = (dayOfWeek - 1)
                                                                firstdayOfFirstWeek = dateAdd("d", -d, firstOfmonth)
                                                                firstThursdayOfFirstWeek = firstdayOfFirstWeek 'dateAdd("d", +3, firstdayOfFirstWeek)
                                                                
                                                                'if year(nextMonthFirst) <> year(firstOfmonth) then 'sidste uge i år ræekker ind i det nye år
                                                                'dividerThisMth = dividerThisMth - 1
                                                                'firstThursdayOfFirstWeek = dateAdd("d", 7, firstThursdayOfFirstWeek)
                                                                'end if

                                                             
                                                             
                                                                if cint(showasproc) <> 1 then
                                                                timerThisMD = formatnumber(timerThisMD/((dividerThisMth/1)+1),2)
                                                                
                                                                procThisMD = formatnumber(procThisMD/((dividerThisMth/1)+1),2)
                                                                
                                                
                                                                end if            


                                                                timerThisMD = replace(timerThisMD, ".", "") 
                                                                timerThisMD = replace(timerThisMD, ",", ".") 


                                                                procThisMD = replace(procThisMD, ".", "") 
                                                                procThisMD = replace(procThisMD, ",", ".")   
                                            
                                                                
                                                                for weekno = 0 to dividerThisMth 
                                                                
                                                            
                                                                thisDate = dateAdd("d", 7*weekno, firstThursdayOfFirstWeek)
                                                                
                                                                if datePart("m", thisDate, 2,2) < datePart("m", firstOfmonth, 2,2) OR datePart("yyyy", thisDate, 2,2) < datePart("yyyy", firstOfmonth, 2,2) then
                                                                thisDate = dateAdd("d", 7, thisDate)
                                                                firstThursdayOfFirstWeek = dateAdd("d", 7, firstThursdayOfFirstWeek)
                                                                end if

                                                                weekThis = datepart("ww", thisDate, 2,2)
                                                                monththis = datepart("m", thisDate, 2,2)
												                yearthis = datepart("yyyy", thisDate)
												               
                                                
                                                                            

                                                                'response.write "dateDiff(m, thisDateGlobal, datoer(y), 2,2): "& dateDiff("m", thisDateGlobal, datoer(y), 2,2)&" thisMth: "& thisMth &" monththis: "& monththis &"<br>   Y: "& y &" datoer(y): "& datoer(y) &" thisDateGlobal: " & thisDateGlobal& " DateDiff M: "& dateDiff("m", thisDateGlobal, datoer(y),2,2)  &" dividerThisMth: "& dividerThisMth & " thisDate: "& thisDate & " firstOfmonth: "& firstOfmonth &" lastDInMonth = "& lastDInMonth &"<br>"
                                                                '** Ved månedsindlæsning må der kune læses timer ind i den valgte måned. Derfor tjekkes om uge om torsdagen går over i næste måned
                                                                '** F.eks slut apr 2014 går over i maj og må ikke indlæses ***
                                                        
                                                                'if cint(thisMth) = cint(monththis) then

                                                                strSQL = "INSERT INTO ressourcer_md"_
																& " (jobid, medid, md, aar, timer, uge, aktid, proc) VALUES ("_
																&" "& tjobid(y) &", "& tmedid(y) &", "& thisMth &", "& yearthis &", "& timerThisMD &", "& weekThis &", "& taktid(y) &", "& procThisMD &")" 
																
                                                                'if lto = "wwf" then
                                                                'response.write "dividerThisMth" & dividerThisMth & " showasproc: "& showasproc & "<br>"
                                                                'Response.Write "strSQL A  "& strSQL & "<br><br>"
                                                                'end if                            

                                                                oConn.execute strSQL

                                                                'end if

                                                                'lastweekwrt = weekThis

                                                                
                                    
                                                                next
                                                                'thisDateGlobal = dateAdd("d", 7, thisDate)

																end select

															
																
															   'if lto = "wwf" then
															   'Response.end
                                                               'end if
																
																

                                                                end if
																
															
															else
																		
																'**** Sletter evt. gamle registreringer (Hvis 0 er angivet) ***	
																'*** Renser ud i tabellen 
																select case periodeSel
																case 3
																strSQLperKri = " AND uge = "& weekThis
																case 6,12
														        strSQLperKri = " AND md = "& monththis
																end select
																
																
																strSQLdel = "DELETE FROM ressourcer_md WHERE jobid = "& tjobid(y) &" AND medid = "& tmedid(y) &" AND aktid = "& taktid_old(y) &""_
																&" AND aar =  "& yearthis & strSQLperKri
															    'AND md = "& monththis &"
															    'Response.Write "<br>strSQldel B"& strSQldel & "<br>"
															    'Response.flush
															    
																oConn.execute strSQLdel
														
															end if
													end if
												
							
												
						next
	                    
                    
	                    'Response.end

                        '***** Års rammer ****'

                        'Response.Write request("FM_aarsramme_medarb")
                        'Response.end

                        arr_rids = split(request("FM_aarsramme_rid"), ",")
                        arr_medids = split(request("FM_aarsramme_medarb"), ",")
                        arr_jobids = split(request("FM_aarsramme_jobid"), ",") 
                        arr_aarids = split(request("FM_aarsramme_aar"), ",") 

                        arr_timerids_temp = replace(request("FM_aarsramme_timer"), ", #,", ";")
						arr_timerids_temp2 = replace(arr_timerids_temp, ", #", "")
						
						arr_timerids = split(arr_timerids_temp2, ";") 'split(request("FM_timer"), ";")
                        

                        'arr_aarids = split(request("FM_aarsramme_timer"), "# ,") 

                        
                        for a = 0 to UBOUND(arr_medids)
                            
                            arr_timerids(a) = replace(arr_timerids(a), ",", ".")

                            

                            if arr_rids(a) <> 0 then
                            strSQLaar = "UPDATE ressourcer_ramme SET timer = "& arr_timerids(a) &" WHERE id = " & arr_rids(a)
                            oConn.execute(strSQLaar)
                            else

                            if arr_timerids(a) <> 0 then
                            strSQLaar = "INSERT INTO ressourcer_ramme (jobid, medid, aar, timer) VALUES ("& arr_jobids(a) &", "& arr_medids(a) &", "& arr_aarids(a) &", "& arr_timerids(a) &")" 
                            'Response.Write strSQLaar & "<br>"
                            oConn.execute(strSQLaar)
                            end if

                            end if


                        next
       
            'Response.end

	
	
	


	
	'Response.end
	response.Redirect "ressource_belaeg_jbpla.asp?FM_progrp="&progrp&"&FM_medarb="&thisMiduse&"&FM_medarb_hidden="&thisMiduse&"&showonejob="&showonejob&"&id="&id&"&periodeselect="&periodeSel&"&jobselect="&jobnavn&"&mdselect="&request("mdselect")&"&aarselect="&request("aarselect")&"&selectplus="&selectplus&"&selectminus="&selectminus&"&FM_vis_job_u_fc="& vis_job_u_fc &"&FM_vis_job_fc_neg="&vis_job_fc_neg&"&FM_hideEmptyEmplyLines="&hideEmptyEmplyLines&"&FM_vis_simpel="&vis_simpel&"&FM_vis_kolonne_simpel="&vis_kolonne_simpel&"&FM_vis_kolonne_simpel_akt="&vis_kolonne_simpel_akt&"&FM_sog="& sogTxt 
	
	
	end if

    
    if func = "overfor" then


                                   


    if len(trim(request("FM_overfor"))) <> 0 then
    overfor = 1
    else
    overfor = 0
    end if 


    if cint(overfor) = 1 then

    overforfra = request("FM_overforfra") 
    overforfraAar = request("FM_overforfraAar")


   lastjob = 0

    select case periodeSel
    case 3
    perSQL = " uge = "& overforfra & ""
    perSQL2 = "WEEK(tdato) = " & overforfra

    if cint(overforfra) < 52 then
    overfortil = overforfra + 1
    overfortilAar = overforfraAar
    else
    overfortil = 1
    overfortilAar = overforfraAar + 1
    end if
    
                                                                                          
    perSQLtil = " uge = "& overfortil & ""
    perSQLtil2 = "WEEK(tdato) = " & overfortil

    case else
    perSQL = " md = "& overforfra & ""
    perSQL2 = "MONTH(tdato) = " & overforfra


    if cint(overforfra) < 12 then
    overfortil = overforfra + 1
    overfortilAar = overforfraAar
    else
    overfortil = 1
    overfortilAar = overforfraAar + 1
    end if

    perSQLtil = " md = "& overfortil & ""
    perSQLtil2 = "MONTH(tdato) = " & overfortil
    end select

    
    call akttyper2009(2)
  


    '**** Overfør alle ressourcetimer der ligger i denne uge / md til næste uge / måned
    '**** Overfør kun diff
    '**** Hvis der allerede er timer i den næste periode lægges disse til
    '**** Uspecificeret hentes altid først, ORDER BY aktid, da de først overføres, da specifikke aktiviteter ellers vil blive overskrevet


                                    strSQL = "SELECT SUM(timer) AS restimer, medid, aktid, jobid, j.jobnr AS jobnr, rmd.id AS rid, uge, md, aar FROM ressourcer_md AS rmd "_
                                    &" LEFT JOIN job AS j ON (j.id = jobid) WHERE " & perSQL & " AND aar = "& overforfraAar &" AND jobstatus = 1 GROUP BY medid, jobid, aktid ORDER BY medid, jobid, aktid DESC"
                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF 

                                    realtimerthisjobAkt = 0

                                    if lastjob <> oRec("jobid") then  
                                    aktidsThisJob = " AND taktivitetid <> 0"
                                    aktidsThisJobSQL = ""
                                    else
                                    aktidsThisJob = aktidsThisJob
                                    end if

                                    'response.Write strSQL & "<br>"
                                    'response.flush

                                        '*** Henter timer og beregner diff

                                        if oRec("aktid") <> 0 then
                                        aktKri = " AND taktivitetid = " & oRec("aktid")
                                        aktidsThisJob = aktidsThisJob & " AND taktivitetid <> " & oRec("aktid") 
                                        aktidsThisJobSQL = ""
                                        else
                                        aktKri = ""
                                        aktidsThisJobSQL = aktidsThisJob
                                        end if
                                            
                                                
                                                diffTimer = oRec("restimer")
                                                realtimerthisjobAkt = 0
                                                
                                                strSQLtimer = "SELECT SUM(timer) AS realtimertot FROM timer WHERE tjobnr = "& oRec("jobnr") &" AND " & perSQL2 &" AND YEAR(tdato) = "& overforfraAar &" AND tmnr = "& oRec("medid") & aktKri & " "& aktidsThisJobSQL &" AND ("& aty_sql_realhours &") GROUP BY tmnr"
                                                
                                                'response.write strSQLtimer & "<br>"
                                                'response.flush

                                                oRec2.open strSQLtimer, oConn, 3
                                                if not oRec2.EOF then

                                                realtimerthisjobAkt = oRec2("realtimertot")
                                                diffTimer = (diffTimer - realtimerthisjobAkt)
                                               
                                                end if
                                                oRec2.close


                                                '** Overfører kun hvis diffTimer > 0, dvs. hvis der mangler at blive udførst arbejde
                                                '** Finder tim i dne måned der skal overføres til
                                                nyeResTimer = 0    
                                                if cdbl(diffTimer) > 0 then
                                            

                                                 'response.write "MEdid: "& oRec("medid") & " jobnr:"& oRec("jobnr") &" aktid: "& oRec("aktid") & " timer: "& oRec("restimer") & " Real timer: "& realtimerthisjobAkt &" diffTimer: "& diffTimer &"<hr>"

                                                strSQL = "SELECT timer AS restimer, medid, aktid, jobid, rmd.id AS rid FROM ressourcer_md AS rmd "_
                                                &" WHERE " & perSQLtil & " AND aar = "& overfortilAar &" AND medid = "& oRec("medid") &" AND jobid = "& oRec("jobid") &" AND aktid = "& oRec("aktid") 
                                                oRec3.open strSQL, oConn, 3
                                                if not oRec3.EOF then 

                                                    nyeResTimer = (diffTimer + oRec3("restimer"))
                                                    nyeResTimer = replace(nyeResTimer, ",", ".")

                                                    strSQLUpdateResTimer = "UPDATE ressourcer_md SET timer = "& nyeResTimer &" WHERE id = "& oRec3("rid") 
                                                    
                                                    'response.write "strSQLUpdateResTimer: "& strSQLUpdateResTimer & "<br>"
                                                    'response.flush
                                                    oConn.execute(strSQLUpdateResTimer)

                                                
                                                else

                                                    nyeResTimer = diffTimer
                                                    nyeResTimer = replace(nyeResTimer, ",", ".")

                                                    select case periodeSel
                                                    case 3 'uge
                                                    ugetil = overfortil
                                                    
                                                    
                                                    weekdayy = datepart("w","1/1/"&overfortilAar, 2,2)                
                                            
                                                    if weekdayy > 4 then 'fredag, lør, søn == uge 53 dvs ekstra uge i diff
                                                    mdtil = dateAdd("w", ugetil + 1, "1/1/"&overfortilAar)
                                                    else
                                                    mdtil = dateAdd("w", ugetil, "1/1/"&overfortilAar)
                                                    end if                     
                                        

                                                    mdtil = month(mdtil)
                                                    proctil = 0
                                                    case else 'md
                                                    ugetil = datepart("ww", "15/"&overfortil&"/"&overfortilAar, 2,2) 'midt i måned
                                                    mdtil = overfortil            
                                                    proctil = 0
                                                    end select

                                                    strSQLInsertResTimer = "INSERT INTO ressourcer_md (timer, medid, aktid, jobid, md, aar, uge, proc) VALUES "_
                                                    &" ("& nyeResTimer &", "& oRec("medid") &", "& oRec("aktid") &", "& oRec("jobid") &", "& mdtil &", "& overfortilAar &","& ugetil &", "& proctil &")"
                                                    
                                                    'response.write "strSQLInsertResTimer: "& strSQLInsertResTimer & "<br>"
                                    
                                                    oConn.execute(strSQLInsertResTimer)

                                                end if 
                                                oRec3.close  



                                                    '*** Nedskriver den måned der overføres fra
                                                    nedskrivResTimer = (realtimerthisjobAkt)
                                                    nedskrivResTimer = replace(nedskrivResTimer, ",", ".")               
                                        
                                    
                                                   
                                    
                                                    if cint(realtimerthisjobAkt) > 0 then          

                                                   

                                                    strSQLUpdateFraResTimer = "UPDATE ressourcer_md SET timer = ("& nedskrivResTimer &") WHERE id = "& oRec("rid") 

                                                   
                                                    else

                                                    strSQLUpdateFraResTimer = "UPDATE ressourcer_md SET timer = 0 WHERE id = "& oRec("rid") 

                                                    end if


                                                     'renser ud i de andre uger ved månedsoverførsel
                                                    select case periodeSel
                                                    case 6,12 'måned
                                                    strSQLUpdateFraResTimerDel = "DELETE FROM ressourcer_md "_
                                                    &" WHERE " & perSQL & " AND aar = "& overforfraAar &" AND medid = "& oRec("medid") &" AND jobid = "& oRec("jobid") &" AND aktid = "& oRec("aktid") &" AND id <> "& oRec("rid")
                                                
                                                    'response.write "<br>strSQLUpdateFraResTimerDel" & strSQLUpdateFraResTimerDel
                                                    

                                                    oConn.execute(strSQLUpdateFraResTimerDel)

                                                  
                                                    end select

                                                    

                                                    oConn.execute(strSQLUpdateFraResTimer)


                                                end if 'Difftimer > 0

                                    lastjob = oRec("jobid")

                                    oRec.movenext
                                    wend
                                    oRec.close




    end if


'response.write "<br><br><br>HER"
'response.write "periodeSel: " & periodeSel
'response.write request("FM_overforfra")
'response.write request("FM_overfor")



    strLink = "ressource_belaeg_jbpla.asp?FM_progrp="&progrp&"&FM_medarb="&thisMiduse&"&FM_medarb_hidden="&thisMiduse&"&showonejob="&showonejob&"&id="&id&"&periodeselect="&periodeSel&"&jobselect="&jobnavn&"&mdselect="&request("mdselect")&"&aarselect="&request("aarselect")&"&selectplus="&selectplus&"&selectminus="&selectminus&"&FM_vis_job_u_fc="& vis_job_u_fc &"&FM_vis_job_fc_neg="&vis_job_fc_neg&"&FM_hideEmptyEmplyLines="&hideEmptyEmplyLines&"&FM_vis_simpel="&vis_simpel&"&FM_vis_kolonne_simpel="&vis_kolonne_simpel&"&FM_vis_kolonne_simpel_akt="&vis_kolonne_simpel_akt&"&FM_sog="& sogTxt 
	
    response.write "<br><br><div style=""padding:20px;"">Timer er overflyttet på alle aktive job.<br><a href="& strLink &">Videre >></a></div>"                                
    response.end                           
    end if


	
	if func = "opduge" then
	
	    strSQL = "SELECT md, aar, id FROM ressourcer_md WHERE id <> 0"
	    oRec.open strSQL, oConn, 3
	    While not oRec.EOF
	            
	            thisdate = "1/"& oRec("md")&"/"&oRec("aar")
	            thisWeek = datepart("ww", thisdate, 2,2)
	            
	            
	            strSQLu = "UPDATE ressourcer_md SET uge = "& thisWeek & " WHERE id = "& oRec("id")
	            'Response.write "<b>"& thisdate &"</b>: "& strSQLu & "<br>"
	            'Response.flush
	            'oConn.execute strSQLu
	            
	    
	    oRec.movenext
	    wend 
	    oRec.close
	
	
	'Response.Write "ok"
	Response.end
	
	end if
	
	
	
	
	
	
	
	
	
	
    
	
	intMid = 0
	intJid = 0
	
	startdato = jobStartKri
	slutdato =  jobSlutKri
	
	
	startdatoMD = month(jobStartKri)
	startdatoYY = year(jobStartKri)
	slutdatoMD =  month(jobSlutKri)
	slutdatoYY = year(jobSlutKri)
	
	'** Forgående 3 md. (totaler) ***
	startdatoMDtot = month(dateadd("m", -3, jobStartKri))
	startdatoYYtot = year(dateadd("m", -3, jobStartKri))
	slutdatoMDtot =  month(dateadd("m", -1, jobStartKri))
	slutdatoYYtot = year(dateadd("m", -1, jobStartKri))
	
	timregStdato = year(dateadd("m", -3, jobStartKri))&"/"& month(dateadd("m", -3, jobStartKri)) &"/"& day(dateadd("m", -3, jobStartKri))
	timregSldato = year(dateadd("m", -0, jobStartKri))&"/"& month(dateadd("m", -0, jobStartKri)) &"/"& day(dateadd("m", -0, jobStartKri))
	
	
	if media <> "print" AND media <> "eksport" AND media <> "chart" then%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    
	        <%if showonejob <> 1 then%>
	
	        <!--#include file="../inc/regular/topmenu_inc.asp"-->

            <!--
	        <div id="topmenu" style="position:absolute; left:0px; top:42px; visibility:visible;">
	        <!--<h4>Timeregistrering - Jobliste</h4>
	        <%call tsamainmenu(2)%>
	        </div>
	         <div id="Div1" style="position:absolute; left:15px; top:82px; visibility:visible;">
	                <%
	                call webbliktopmenu()
	                %>
	                </div>
                -->

            <%call menu_2014() %>

	        <%
	        topdiv = 102
	        leftdiv = 90
	        else
	        topdiv = 102
	        leftdiv = 20
	        end if
	
	else
	
	
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%
	
	
	
	topdiv = 0
	leftdiv = 40
    
	
	end if

	
	'*************************************************************************************	
	
	%>
	
	
		
	
     
    <script src="inc/res_bl_jav.js"></script>
	
    <!-------------------------------Sideindhold------------------------------------->
	
	<%
	
	
	


	select case media
    case "eksport"
    ldTop = 150
	ldLft = 200
    case "chart", "print"
    ldTop = 200
	ldLft = 350
	case else  
    ldTop = 400
	ldLft = 350
	
	end select

    if media <> "print" then
	%>
	<div id="load" style="position:absolute; display:; visibility:visible; top:<%=ldTop%>px; left:<%=ldLft%>px; width:300px; background-color:#ffffff; border:10px #6CAE1C solid; padding:20px; z-index:100000000;">
    <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" /><br />
	&nbsp;&nbsp;&nbsp;&nbsp;Vent veligst. Forventet loadtid: 3-25 sek.
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	</td></tr></table>
	
	</div>
	<% end if
	
	Response.Flush 
	
	'end if %>
	
	<%'Response.Flush %>
	<%'Response.End %>
	
	
	
	
	<%
	
	if media <> "eksport" then
	
	         if media <> "chart" then%>
	        <div id="overblik" style="position:absolute; left:<%=leftdiv%>px; top:<%=topdiv%>px; visibility:visible;">
	
	        <%end if
	
	'oimg = "ikon_ressourcer_48.png"
	oleft = -20
	otop = 0
	owdt = 800
   
    oskrift = "Ressource Forecast (timebudget)"
  
    if media = "print" then
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
    end if

   end if
	
	
	if media <> "print" then
	toppxThisDiv = 260
	leftpxThisDiv = 1110
	else
	toppxThisDiv = 20
	leftpxThisDiv = 910
	end if
	
	
	
	
	if media <> "print" ANd media <> "eksport" AND media <> "chart" then%>
	
	
	<%call filterheader_2013(0,0,1004,oskrift) %>
	
	<table cellspacing="0" cellpadding="0" border="0" width=100%>
	<form action="ressource_belaeg_jbpla.asp?showonejob=<%=showonejob%>&id=<%=id%>" method="post" name="mdselect" id="mdselect">
	
    <tr><td colspan="5"><span id="sp_med" style="color:#5582d2;">[+] Projektgrupper & Medarbejdere</span></td></tr>
	<tr id="tr_prog_med" style="display:none; visibility:hidden;">
     <%call progrpmedarb %>

    
	    
       
    </tr>
	<tr><td colspan="5"><span id="sp_job" style="color:#5582d2;">[+] Vælg job</span></td></tr>
	<tr id="tr_job" style="display:none; visibility:hidden;">
	
	<%if showonejob <> 1 then%>
	<td valign=top style="padding-top:20px;" colspan="2"><!--<h4>Vis forecast for job:<span style="font-size:11px; font-weight:lighter;"> (prioitet > -1)</span></h4>-->
        Job:<br />
	
	<%
		strSQL = "SELECT jobnavn, jobnr, j.id, k.kkundenavn, k.kkundenr, k.kid"_
		&" FROM job j "_
		&" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_
	    &" WHERE jobstatus = 1 "& jobDatoKri &" AND j.fakturerbart = 1 AND j.risiko > -1 ORDER BY k.kkundenavn, j.jobnavn"
		
		'3 = tilbud
        'if lto = "hvk_bbb" AND session("mid") = 1 then
		''Response.write strSQL
		'Response.flush
		'end if
        %>
        
		
		<select name="FM_jobsel" id="FM_jobsel" style="width:960px; font-size:11px;" size=10>
		<%if cint(jobidsel) = 0 then
        jallSel = "SELECTED"
        else
        jallSel = ""
        end if%>
        
        <option value="0" <%=jallSel %>>Alle</option>
		<%
				
				oRec.open strSQL, oConn, 3
                k = 0
				while not oRec.EOF
				
				if cint(jobidsel) = cint(oRec("id")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if

                if cdbl(lastKid) <> oRec("kid") then
                %>
                <option value="<%=oRec("id")%>" disabled></option>
                <option value="<%=oRec("id")%>" disabled><%=oRec("kkundenavn")%>&nbsp;(<%=oRec("kkundenr")%>)</option>
                <%
				end if
				
				%>
				<option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("jobnavn")%>&nbsp;(<%=oRec("jobnr")%>) <%=aftnavnognr %></option>
				<%

                lastKid = oRec("kid")
                k = k + 1
				oRec.movenext
				wend
				oRec.close
				
				%>
				
		</select>
		<br />
		
		       <input id="viskunjobiper0" name="viskunjobiper" value="0" type="radio" <%=viskunjobiper0 %> onclick="submit();" /> Vis <b>alle</b> aktive job.<br />
 
		       <input id="viskunjobiper1" name="viskunjobiper" value="1" type="radio" <%=viskunjobiper1 %> onclick="submit();" /> Vis <b>aktuelle</b> job (kun job, hvis slutdato ligger efter den valgte periode start)
        <br /><br />&nbsp;
        </td></tr>
        <tr><td style="padding-top:20px;">      
        <h4>Søg på jobnavn ell. nr.:</h4>

        <input id="sogtxt" name="FM_sog" value="<%=sogTxt %>" type="text" style="width:350px; font-size:14px; border:2px #6CAE1C solid;;" />
    <br />(% wildcard, <b>100-200</b> interval, eller <b>58, 89</b> for specifikke job)
            <br /><br /><br />

            <b>Filter:</b><br />
               <input type="checkbox" value="1" id="FM_vis_job_fc_neg" name="FM_vis_job_fc_neg" <%=vis_job_fc_negCHK%> /> <b>Forecast overskreddet</b> <span style="color:#999999;">(kun job og aktiviteter, hvor forecast (timebudget) er overskreddet)</span>
                <br />
               <input type="checkbox" value="1" id="FM_vis_job_u_fc" name="FM_vis_job_u_fc" <%=vis_job_u_fcCHK %> /> Vis aktive job <b>uden forecast</b> <span style="color:#999999;">(alle job på personlig aktivliste +/-5 md., maks 10 medarb.)</span> 
                <br />
           
            <input type="checkbox" value="1" id="Checkbox1" name="FM_hideEmptyEmplyLines" <%=hideEmptyEmplyLinesCHK %> /> Skjul medarbejder(e) uden forecast<br />        


            
           
                 <br /><img src="../blank.gif" border="0" width="0" height="7" />
            


             


   
	
	<%else%>
	<td>&nbsp;
    <input type="hidden" name="FM_jobid" id="FM_jobid" value="<%=jobid%>"><b><%=jobid%></b>
	<%end if%>

     
      
    
    
	</td>

	
    <td valign=top style="padding:20px 30px 10px 10px;"><b>Periode fra:</b><br /> 
	<select name="mdselect">
	<option value="<%=monththis%>" SELECTED><%=left(monthname(monththis),3)%></option>
	<option value="1">Jan</option>
	<option value="2">Feb</option>
	<option value="3">Mar</option>
	<option value="4">Apr</option>
	<option value="5">Maj</option>
	<option value="6">Jun</option>
	<option value="7">Jul</option>
	<option value="8">Aug</option>
	<option value="9">Sep</option>
	<option value="10">Okt</option>
	<option value="11">Nov</option>
	<option value="12">Dec</option>
	</select> <select name="aarselect">
	<option value="<%=yearthis%>" SELECTED><%=yearthis%></option>

    <%for yThis = 0 to 30  

    yUse = 2007 + yThis
    
    %>
    <option value="<%=yUse %>"><%=yUse %></option>
    <%
    next
    %> 

	
	
	</select> og 
	<select name="periodeselect" style="width:190px;" onchange="submit()">
	<option value="<%=periodeSel%>" SELECTED><%=periodeShow%></option>
	<!--<option value="1">1 Måned (dag/dag)</option>-->
	<option value="3">3 Måneder frem (uge/uge)</option>
	<!--<option value="4">4 Måneder frem (måned/måned)</option>-->
    <%select case lto 
        case "kejd_pb"
        case else %>
	<option value="12">12 Måneder frem</option>
	<option value="6">6 Måneder frem</option>
        <%end select %>
    
    </select>

        <%
	leftKnap = 170
	%>
    

     <br /><br /><br />
            <b>Præsentation:</b><br />
             <input type="checkbox" value="1" id="Checkbox2" <%=vis_kolonne_simpelDisabled %> name="FM_vis_kolonne_simpel" <%=vis_kolonne_simpelCHK %> /> Vis ikke års-ramme og saldo kolonner<br />
            <input type="checkbox" value="1" id="Checkbox3" <%=vis_kolonne_simpel_aktDisabled %> name="FM_vis_kolonne_simpel_akt" <%=vis_kolonne_simpel_aktCHK %> /> Vis ikke aktivitets kolonner           
            <br />

             <select style="width:350px;" id="FM_vis_simpel" name="FM_vis_simpel" onchange="submit();">
                <option value="0" <%=vis_simpelCHK0 %>>Vis alle linier (standard)</option>
                <option value="1" <%=vis_simpelCHK1 %>>Skjul alle sub-totaler (optimeret til indtastning)</option>
                <option value="2" <%=vis_simpelCHK2 %>>Vis Grafisk (optimeret til overblik)</option>
            </select>
            <!--<input type="checkbox" value="1" id="FM_vis_simpel" name="FM_vis_simpel" <%=vis_simpelCHK %> /> Skjul <b>sub-totaler</b> på job og medarbejdere (bedre overblik, hvis der er valgt mange medarbejdere) </span> 
                -->


    
   <br /><br /><br />
       Angiv forecast som  <select name="FM_showasproc" onchange="submit();"><option value="0" <%=showasproc0Sel %>>Timer</option><option value="1" <%=showasproc1Sel %>>Procent</option></select>
  
    <br />&nbsp;
    </td>
	


	</tr>

    <tr>
    <td valign=top style="padding-top:5px;">
     &nbsp;</td>
    <td align=right style="padding-right:40px; padding-top:4px;">
        <input id="Submit1" type="submit" value=" Søg >> " /></td>
    </tr>

	</form>
	</table>
	</div>
	<!-- filter header --->
	</td>
	</tr>
	</table>
	</div>
	<%
    else
        %>
        <input type="hidden" id="FM_vis_simpel" value="<%=vis_simpel %>" />
        <%    
    end if 'Print


    

    

    aarsRamme = datepart("yyyy",startdato, 2,2)
	

    if cint(visRamme) = 1 then


        select case periodeSel
        case 3
        tWdth = 1794
        case 12
        tWdth = 1794
        case else
        tWdth = 1694
	    end select


    else

        select case periodeSel
        case 3
        tWdth = 1294
        case 12
        tWdth = 1394
        case else
        tWdth = 1194
	    end select

    end if



	if media <> "eksport" AND media <> "chart" then
	
	tTop = 0
	tLeft = -40
	
    tId = "sidediv"
	tVzb = "visible"
	tDsp = ""
	
	if media <> "print" then
	'dtop = topdiv + 110
    tLeft = tLeft + 40
	tTop = 20   
      
        
        call tableDivWid(tTop,tLeft,tWdth,tId,tVzb,tDsp)
	
    else
   
	'dtop = topdiv
	call tableDiv(tTop,tLeft,tWdth)
	end if
	
	
	end if
	
  
	
	
    'Response.write "periodeSel: " & periodeSel


    '** HENTER ALLE FORVALGTE JOVB fra timereg_usejob ****'
    'strSQLtu = "SELECT jobid, medarb, aktid forvalgt FROM timereg_usejob AS tu WHERE tu.medarb = "& mids
    
    Response.flush
	
	
	'**** Akt typer der er med i ddagligt timeregnskab ***'
    aty_sql_realhoursAKt = replace(aty_sql_realhours, "tfaktim", "fakturerbar")
	
	isArsNormWrtEksp = 0
	skrivCsvFil = 1 
	public csvTxtTop, csvTxt
	public csvDatoerMD(35), csvDatoerAR(35), csvDatoerUge(35)
   
	public lastxmid
	lastxmid = 0
	
	mids = 1150
	dt = 300 '1000
	ji = 6

	dim midjiddato 
	redim midjiddato(mids, dt, ji) 

   
    dim n_arr
    redim n_arr(mids)
	
	antalX = 6000
	xId = 25
	
	'Response.flush
	
	public medarbKundeoplysX
	redim medarbKundeoplysX(antalX, xId) 
	
	'dim testarr
	'***  Md, Aar, Jobid, Mid **
	'select case lto
	'case "bminds", "execon"
	'redim testarr(53,12,44,900,110)
	'case else
	'redim testarr(53,12,500,40)
	'end select
	
	'*** Unique job id brugs i array, så hvis de når job id nr 901 fejler siden. **'
	
	lastYear = 0
	lastWeek = 0
	lastMid = 0
	lastJid = 0
    lastJid_Aid = 0
	'public n
    n = 0
    m = 0
	'wrtNewRow = 0
	
	lastmedarbId = 0
	totFC_job_med_IsWrt = "##-1##"
	
   
	if cint(startdatoYY) = cint(slutDatoYY)  then
	orandval = " AND "
	else
	orandval = " OR "
	end if
	
    public jobisWrt, jobIdisWrt
    redim  jobisWrt(4000), jobIdisWrt(4000) 
	
    medarbwrt = "#0#"
    fc_lastjobnrWrt = 0
    fc_lastaktidWrt = 0
	
	'************************************************************************
	'*** Main SQL ***
	'************************************************************************
	't = 0 alle med ress. timer i forecast
	't = 1 alle med realiserede timer og som ikke er med i forecast
	public x
	x = 0
	't1first = 1
	tStDato = year(jobStartKri) &"-"& month(jobStartKri) & "-1" 
	tSlDato = jobSlutKri
	
    tStDatoJobUforecast = dateAdd("m", - 6, jobStartKri)
    tStDatoJobUforecast = year(tStDatoJobUforecast) & "-" & month(tStDatoJobUforecast) & "-1"
	
    tSlDatoJobUforecast = dateAdd("m", 1, now)
    tSlDatoJobUforecast = year(tSlDatoJobUforecast) &"-"& month(tSlDatoJobUforecast) & "-" & day(tSlDatoJobUforecast)

	strSQL = "SELECT jobnavn, m.mnavn, m.mnr, j.id AS jobid, j.jobnr, j.jobstatus, m.medarbejdertype, j.jo_bruttooms, j.jo_udgifter_ulev, j.stade_tim_proc, j.budgettimer, j.restestimat, "_
    &" m.mid, a.navn AS aktnavn, "
    strSQL = strSQL & "sum(rmd.timer) AS restimer, SUM(rmd.proc) AS proc, rmd.md AS resmd, rmd.aar AS resaar, rmd.id AS resid, rmd.uge AS resuge, rmd.aktid AS aktid, "
	strSQL = strSQL & " k.kkundenavn, k.kkundenr, m.init, m.forecaststamp, jobans1, jobans2, "_
	&" m2.mnavn AS m2mnavn, m2.init AS m2init, m2.mnr AS m2mnr, m3.mnavn AS m3mnavn, m3.init AS m3init, m3.mnr AS m3mnr "_
    &" FROM medarbejdere m "
	
    strSQL = strSQL &" LEFT JOIN ressourcer_md rmd ON (rmd.medid = m.mid AND "_
	&" ((rmd.md >= "& startdatoMD &" AND rmd.aar = "& startdatoYY &") "& orandval &" (rmd.md <= "& slutdatoMD &" AND rmd.aar = "& slutdatoYY &")))"_
	&" LEFT JOIN job j ON (j.id = rmd.jobid)"
	
    'tu.jobid
    '&" LEFT JOIN timereg_usejob AS tu ON (tu.medarb = m.mid)"_

	strSQL = strSQL &" LEFT JOIN kunder k ON (kid = j.jobknr)"_
	&" LEFT JOIN medarbejdere m2 ON (m2.mid = j.jobans1)"_
	&" LEFT JOIN medarbejdere m3 ON (m3.mid = j.jobans2)"_
    &" LEFT JOIN aktiviteter a ON (a.id = rmd.aktid)"
	
	strSQL = strSQL &" WHERE "& medarbSQlKri &" "& deakSQL & "" 'strMedidSQL 
	
	strSQL = strSQL &" AND m.mid <> 0 AND "& jobSQLkri &"" ' AND j.jobnavn IS NOT NULL"

	strSQL = strSQL &" AND (j.risiko > -1 OR j.risiko = -3)"
    
    strSQL = strSQL &" GROUP BY m.mid, rmd.jobid, rmd.aktid, rmd.aar, rmd.md"
            
            if periodeSel = 3 then
            strSQL = strSQL &", rmd.uge"
            end if

    strSQL = strSQL &" ORDER BY m.mid, k.kkundenavn, rmd.jobid, rmd.aktid DESC, a.fase, a.sortorder, a.navn" 'rmd.aktid
   
	strSQL = strSQL &", rmd.aar, rmd.md"

            if periodeSel = 3 then
            strSQL = strSQL &", rmd.uge"
            end if

	
	'AND (m.mansat <> 2 AND m.mansat <> 3)
        'if lto = "essens" then
	    'Response.write strSQL &"<br><br>"
	    'Response.flush
        'end if
	
	oRec.open strSQL, oConn, 3 
	
	
	while not oRec.EOF 

    if x = 0 then
        jobisWrt(oRec("mid")) = " AND (t.tjobnr <> '0' "
        jobIdisWrt(oRec("mid")) = " AND (jobid <> '0' "
    end if
    

    '** Afslutter forrige og finder job med timer på udenne forecast pr. medarb **'
    if lastMid <> oRec("mid") AND x <> 0 then
        
       
        lastJid = 0
        n_arr(lastMid) = n-1 

        '*** Henter job uden forecast ***'
        if cint(vis_job_u_fc) = 1 then
        call jobmedtimeruforecast(lastMid, 1)
        
        jobisWrt(oRec("mid")) = " AND (t.tjobnr <> '0' "
        jobIdisWrt(oRec("mid")) = " AND (jobid <> '0' "
        medarbwrt = medarbwrt & ",#" & lastMid & "#"

        end if

        n = 0
     

    end if

        'if oRec("jobnr") <> "" AND instr(jobisWrt(oRec("mid")), " AND t.tjobnr <> "& oRec("jobnr")) = 0 then
        if lastJid <> oRec("jobid") then 
	    jobisWrt(oRec("mid")) = jobisWrt(oRec("mid")) & " AND t.tjobnr <> '"& oRec("jobnr") &"'"
        jobIdisWrt(oRec("mid")) = jobIdisWrt(oRec("mid")) & " AND jobid <> '"& oRec("jobid") &"'" 
	    end if
    
   
        if len(trim(oRec("resmd"))) <> "" then
	    showresdato = "1/"&oRec("resmd")&"/"&oRec("resaar")
	    showresuge = oRec("resuge")
        else
        showresdato = "1/1/2002"
	    showresuge = 1
        end if
	
	
	medarbKundeoplysX(x, 0) = oRec("jobid")
	medarbKundeoplysX(x, 1) = oRec("jobnr")
    medarbKundeoplysX(x, 2) = oRec("jobnavn")
    medarbKundeoplysX(x, 3) = oRec("mid")
    medarbKundeoplysX(x, 4) = oRec("mnavn")
    
	medarbKundeoplysX(x, 5) = oRec("kkundenavn")
	medarbKundeoplysX(x, 6) = oRec("kkundenr")
	medarbKundeoplysX(x, 7) = oRec("mnr")
	medarbKundeoplysX(x, 8) = oRec("init")
	
	'if t = 0 then
	medarbKundeoplysX(x, 9) = oRec("forecaststamp")

    'if cint(showasproc) = 1 then
    'medarbKundeoplysX(x, 10) = oRec("proc")
    'else
	medarbKundeoplysX(x, 10) = oRec("restimer")
	'end if

    '** procent altid indtastet i forhold til normtid i den valgete periode. **'
    medarbKundeoplysX(x, 23) = oRec("proc") 
    'medarbKundeoplysX(x, 24) = oRec("mdorweek")
   
            'response.write "JOB: "& oRec("jobnr") &" PROC: "& oRec("proc") & "<br>"


    resid = oRec("resid")
	'else
	'medarbKundeoplysX(x, 9) = "2002-01-01" 'oRec("forecaststamp")
	'medarbKundeoplysX(x, 10) = 0
	'resid = -1
	'end if
	
	
	medarbKundeoplysX(x, 11) = oRec("jobans1")
	medarbKundeoplysX(x, 12) = oRec("jobans2")
	
	if len(trim(oRec("m2mnavn"))) <> 0 then
    medarbKundeoplysX(x, 13) = oRec("m2mnavn") &" ("&oRec("m2mnr")&")"
    medarbKundeoplysX(x, 14) = oRec("m2init") 
    else
    medarbKundeoplysX(x, 13) = ""
    medarbKundeoplysX(x, 14) = ""
    end if
    
    if len(trim(oRec("m3mnavn"))) <> 0 then
    medarbKundeoplysX(x, 15) = oRec("m3mnavn") &" ("&oRec("m3mnr")&")"
    medarbKundeoplysX(x, 16) = oRec("m3init")
    else
    medarbKundeoplysX(x, 15) = ""
    medarbKundeoplysX(x, 16) = ""
    end if 
	
     medarbKundeoplysX(x, 17) = oRec("aktid")
    
     medarbKundeoplysX(x, 18) = oRec("aktnavn")
     'Response.Write medarbKundeoplysX(x, 0) &"_"& medarbKundeoplysX(x, 17) & "<br>"


    medarbKundeoplysX(x, 19) = oRec("jo_bruttooms")
    medarbKundeoplysX(x, 20) = oRec("budgettimer")
    medarbKundeoplysX(x, 21) = oRec("stade_tim_proc")
    medarbKundeoplysX(x, 22) = oRec("restestimat")   


    select case oRec("jobstatus")
            case 0
            medarbKundeoplysX(x, 25) = tsa_txt_020
            case 1
            medarbKundeoplysX(x, 25) = tsa_txt_019
            case 2
            medarbKundeoplysX(x, 25) = "Passiv/til fak."
            case 3
            medarbKundeoplysX(x, 25) = "Tilbud"
            case 4
            medarbKundeoplysX(x, 25) = "Gennemsyn"
    end select
    
	       
			if len(resid) <> 0 AND len(oRec("jobid")) <> 0 then
		         
		         
		         'if periodeSel = 3 OR periodeSel = 6 OR periodeSel = 12 then   
		            
				     
					        'Response.Write "N: " & n & "<br>"
					        'Response.flush
					        midjiddato(oRec("mid"),n, 0) = medarbKundeoplysX(x, 10) 'oRec("restimer") 'showrestimer(x) 
                            midjiddato(oRec("mid"),n, 6) = medarbKundeoplysX(x, 23) 'procent
					        midjiddato(oRec("mid"),n, 1) = oRec("jobid") 'useJobid(x) 
					        midjiddato(oRec("mid"),n, 2) = oRec("mid") 'xmid(x) 
					        midjiddato(oRec("mid"),n, 3) = showresdato 
					        midjiddato(oRec("mid"),n, 4) = showresuge
                            midjiddato(oRec("mid"),n, 5) = oRec("aktid")
                           
				      
				        
				  ' else
				   '*** MD / MD tjkker ikke om uge passer **'
				            
				            
				     

                            'if t = 0 then
					        'Response.Write "<br>dt: "& formatdatetime(showresdato, 2) &"n:"& n &" x:"& x &" arr: md: "& oRec("resmd") &",AAr: "& right(oRec("resaar"),2) &", Job: "& oRec("jobid") &", MId: "& oRec("mid") & ":Navn: "& medarbKundeoplysX(x, 4) &" Timer: "& medarbKundeoplysX(x, 10) &"<br>"
	                        'Response.Flush
        					'end if
					        'testarr(oRec("resuge"),right(oRec("resaar"),2),oRec("jobid"),oRec("mid")) = oRec("restimer")
        	                
        	                
        	                
					        
					        'Response.Write "N: " & n & "<br>"
					        'Response.flush
					'        midjiddato(oRec("mid"),n, 0) = medarbKundeoplysX(x, 10) 'timer 
                    '       midjiddato(oRec("mid"),n, 6) = medarbKundeoplysX(x, 23) 'procent

					'        midjiddato(oRec("mid"),n, 1) = oRec("jobid") 'useJobid(x) 
					'        midjiddato(oRec("mid"),n, 2) = oRec("mid") 'xmid(x) 
                            'if t = 0 then
					        'midjiddato(oRec("mid"),n, 3) = "1/"&oRec("resmd")&"/"&oRec("resaar") 'showresdato 
                            'else
                    '        midjiddato(oRec("mid"),n, 3) = showresdato 
                            'end if
					'        midjiddato(oRec("mid"),n, 4) = showresuge
                    '        midjiddato(oRec("mid"),n, 5) = oRec("aktid")

                        
				   
				   'end if

               
			
            end if

			lastWeek = showresuge	
			lastYear = datepart("yyyy", showresdato) 
			lastMonth = datepart("m", showresdato, 2,2) 
			lastMid = oRec("mid")
			lastJid = oRec("jobid")
            lastJid_Aid = oRec("jobid") &"_"& oRec("aktid")
            lastx = x
			
             
			
	
	
    n = n + 1

    

    'Response.Write "N:" & n & "<br>"
    'n_arr(oRec("mid")) = n
	
	x = x + 1
	oRec.movenext
	wend
	oRec.close 
	




    n_arr(lastMid) = n-1 


	    '**** Henter timer på job udne forecast på sidste medarbejer
        '**** Kun hvis der kun er 1 medarb. med forecast på, så bliver den ikke hentet i loopet 

         '*** Henter job uden forecast ***'
        if cint(vis_job_u_fc) = 1 then

         call jobmedtimeruforecast(lastMid, 2)
         medarbwrt = medarbwrt & ",#" & lastMid & "#"


        '*** Henter resten af e valgte medarbejdere uden forecast **'
	
   

        for m = 0 to UBOUND(intMids)
            
                 

            if instr(medarbwrt, "#"& intMids(m) &"#") = 0 then
            jobisWrt(intMids(m)) = " AND (t.tjobnr <> '0'"
            jobIdisWrt(intMids(m)) = " AND (jobid <> '0'"

            'Response.Write medarbwrt & " " & intMids(m) & " // " & instr(medarbwrt, "#"& intMids(m) &"#") & "<br>" 
            call jobmedtimeruforecast(intMids(m), 2)

            '** Hvis der er valgt alle job skal denne være slået til, da loopet ellers finder medarbejdere dobbelt. ***'
            'if jobidsel = 0 then
            'medarbwrt = medarbwrt & "1538,#" & intMids(m) & "#"
            'end if
            
            end if
        next

        end if  'vis_job_u_fc ***'




        '*** vis kolonner ****'
        if cint(vis_kolonne_simpel) = 1 then
        visRamme = 0
        visStatus = 0
        else
        visRamme = 1
        visStatus = 1
        end if

         '*** vis akt. kolonner ****'
        if cint(vis_kolonne_simpel_akt) = 1 then
        visAkt = 0
        else
        visAkt = 1
        end if
        
	%>
	
    

 <%if media <> "print" AND media <> "eksport" AND media <> "chart" then%>
<form action="ressource_belaeg_jbpla.asp?func=redalle&FM_progrp=<%=progrp%>&FM_medarb=<%=thisMiduse%>&FM_medarb_hidden=<%=thisMiduse%>&showonejob=<%=showonejob%>&id=<%=id%>&FM_vis_job_u_fc=<%=vis_job_u_fc %>&FM_vis_job_fc_neg=<%=vis_job_fc_neg %>&FM_vis_simpel=<%=vis_simpel%>&FM_hideEmptyEmplyLines=<%=hideEmptyEmplyLines %>&FM_vis_kolonne_simpel=<%=vis_kolonne_simpel %>&FM_vis_kolonne_simpel_akt=<%=vis_kolonne_simpel_akt%>" method="post" name="indlas" id="indlas">
<!--name="alledagemedarb" id="alledagemedarb"-->

<!--


<input type="hidden" name="inklweekend_2" id="inklweekend_2" value="0">
-->
<input id="Text1" name="FM_sog" value="<%=sogTxt %>" type="hidden"/>
<input type="hidden" name="mdselect" id="Hidden2" value="<%=monththis%>">
<input type="hidden" name="aarselect" id="aarselect" value="<%=yearthis%>">
<input type="hidden" name="periodeselect" id="periodeselect" value="<%=periodeSel%>">
<input type="hidden" name="jobselect" id="Hidden1" value="<%=jobnavn%>">


<table cellspacing=0 cellpadding=0 border=0 width=100%>
		<tr>
            <td>
                 <input type="hidden" name="selectminus" id="selectminus" value="0">
	<input type="hidden" name="selectplus" id="selectplus" value="0">
	
	 <table cellspacing=4 cellpadding=0 border=0 ><tr><td>
	<input type="button" name="submitplus" id="submitminus" value="<< <%=monthname(month(dateadd("m", -1, "1/"&monththis&"/"&yearthis)))%>" onClick="selminus()">
	</td><td>
	<input type="button" name="submitplus" id="submitplus" value="<%=monthname(month(dateadd("m", +1, "1/"&monththis&"/"&yearthis)))%> >>" onClick="seladd()">
    </td></tr></table>

            </td>
            
            <td align=right style="padding:10px 30px 5px 0px;">
		<input type="submit" value="Indlæs forecast >> "></td></tr>
</table>
<%end if %>




<%if media <> "eksport" AND media <> "chart" then

%>

<table border=0 cellspacing=1 cellpadding=0 bgcolor="#CCCCCC" width=100%>

<%end if

MedIdSQLKri = ""
MedTotSQLKri = ""	
'MedJobFcastSQLKri = ""

select case periodeSel
'case 1
'numoffdaysorweeksinperiode = datediff("d", startdato, slutdato)
case 3
numoffdaysorweeksinperiode = datediff("ww", startdato, slutdato)
case 6, 12
numoffdaysorweeksinperiode = datediff("m", startdato, slutdato)
end select

fob = 0
y = 0 
nday = startdato
weekrest = 0		
antalxx = x
xx = 0
'foundone = 0
daynowrpl = replace(daynow, "/", "-")
lastMonth = monththis

thisMTotRamme = 0

antalJobAktlinier = 0
antalJobAktMedlinier = 0
antalJobAktlinierGrand = 0
'antanlJoblinierprM = 0
thisMedarbJoblist = ""

lastAktJidExpVis_1 = "0_0_0"
dagsdato = now

if media <> "eksport" AND media <> "chart" then
call datoeroverskrift(1)
else
call datoeroverskriftCSV(1)
end if

'public MedTotSQLKri
'strjIdWrt = "0#," 


    'Response.write "HER: x: "& x &" antalx: "& antalxx
    'Response.flush

'**** Medarbejdere ****
x = 0
for x = 0 to antalxx - 1 'UBOUND(medarbKundeoplysX, 1) 'medarbKundeoplysX(x, 0)


    'Response.write "<br>HER: x: "& x &" antalx: "& antalxx
    'Response.flush



    '**** Ny overksift for de linier er ligger timer på men der ikke ligger forecast på ***'
  
	
	if lastxmid <> medarbKundeoplysX(x, 3) then	'xmid(x) 

    
	
	tdcolspan = 5
	strDageChkboxOne = "<input type=text size=10 name=FM_sel_dage id=FM_sel_dage value="&daynowrpl&">&nbsp;dd-mm-aaaa"
	
	
    
    
    
    
    
    if x > 0 then				
		
	
                    '*** Total pr. medarb linije ********************

		            if media <> "eksport" AND media <> "chart" AND cint(vis_job_fc_neg) <> 1 then
		                    
                            
                               
                                    call jobtotalprmedarb
                               
                           
                        if cint(jobidsel) = 0 then
                            'if cint(vis_simpel) <> 1 then
                            mTotthisMid = lastxmid
	                        call medarbtotal
                            'end if
	                    end if
	    
	                end if
	    
                    
                    

                    if cint(jobidsel) = 0 AND cint(vis_job_fc_neg) <> 1 then
                    
                        if media = "chart" then
                                
                                call meStamdata(medarbKundeoplysX(x-1, 3))
                                csvTxtA = csvTxtA & meTxt 

                                for y = 0 TO numoffdaysorweeksinperiode
                                csvTxtA = csvTxtA & ";"& medarbTotalTimer(y)
                                medarbTotalTimer(y) = 0 
                                next

                                csvTxtA = csvTxtA & "xx99123sy#z"
                        end if

                    end if
            
	end if 'x
	
	
    
                
	            
    
                if media <> "eksport" AND media <> "chart" then 
        
                    'Response.Write "#"&medarbwrt&"#, "& medarbKundeoplysX(x, 3) &"<br>"
	                'call medarbejderLinje(medarbKundeoplysX(x, 3),jobStartKri,datoInterval,rdimnb)
                    medarbwrt = medarbwrt & ",#" & medarbKundeoplysX(x, 3) & "#"
         
	            end if
	
    



	
		       if media <> "print" AND media <> "eksport" AND media <> "chart" then 'AND (instr(strjIdWrt, "#"&jobidSel&"#") = 0) then 
				    
				    call tilfojNylinie(medarbKundeoplysX(x, 3), visRamme, visAktiv)
				   
				    
				end if 'Print
		
		
	
	
	
	MedIdSQLKri = medarbKundeoplysX(x, 3)
	csvTmnavn = medarbKundeoplysX(x, 4)
	csvTmnr = medarbKundeoplysX(x, 7)
	csvTinit = medarbKundeoplysX(x, 8)
	csvTnormuge = ntimPer


	end if '*** lastXMid / MedarbejderID
	
		
		
		

        'Response.Write "lastJid: " & lastJid  &" <> medarbKundeoplysX(x, 0) " & medarbKundeoplysX(x, 0) & "  OR " & lastxmid &" <> "& medarbKundeoplysX(x, 3) &" AND "& len(medarbKundeoplysX(x, 0)) &" <> 0 <br>"
		
        '***** Total pr. job indenfor medarbejder
        if lastJid <> medarbKundeoplysX(x, 0) AND len(medarbKundeoplysX(x, 0)) <> 0 AND fob = 1 AND cint(vis_job_fc_neg) <> 1 then
            call jobtotalprmedarb
            'antanlJoblinierprM = 0
        end if


		
        
        
        '*** Tjekker om linje på medarb er skrevet **'
        if (lastJid_Aid <> medarbKundeoplysX(x, 0)&"_"& medarbKundeoplysX(x, 17) OR lastxmid <> medarbKundeoplysX(x, 3)) AND len(medarbKundeoplysX(x, 0)) <> 0 then
		
        isArsNormWrtEksp = 0


        



                   '**** Total forecast pr. medarb. pr. job / akt. ***'
                   '**** Toal Real pr. medarb. pr. job / akt. ********'

                    if inStr(totFC_job_med_IsWrt, "##"&x&"##") = 0 then
                        
                         forcastTimerTotAktMedignPer = 0
                         realTimerTotAktMedignPer = 0

                        '*** Forecast, total ign. periode ***
                        strSQLforcTotjobMid = "SELECT SUM(timer) AS sumTimerAktMedIgnPer FROM ressourcer_md WHERE medid = "& medarbKundeoplysX(x, 3) &" AND aktid = "& medarbKundeoplysX(x, 17) & " AND jobid = "& medarbKundeoplysX(x, 0) &" GROUP BY medid, jobid"
                        'Response.write strSQLforcTotjobMid & "<br><br>"
                        'Response.flush

                        oRec3.open strSQLforcTotjobMid, oConn, 3
                        if not oRec3.EOF then
                    
                        forcastTimerTotAktMedignPer = oRec3("sumTimerAktMedIgnPer") 

                        end if
                        oRec3.close 

                 

                        '*** Real. tot ign periode **'
                        if medarbKundeoplysX(x, 17) <> 0 then
                        aktSQLkri = " AND taktivitetid = "& medarbKundeoplysX(x, 17) 
                        else
                        aktSQLkri = ""
                        end if

                        strSQLRealTotjobMid = "SELECT SUM(timer) AS sumTimerAktMedIgnPer FROM timer WHERE tmnr = "& medarbKundeoplysX(x, 3) &" "& aktSQLkri &" AND tjobnr = '"& medarbKundeoplysX(x, 1) &"' AND ("& aty_sql_realhours &") GROUP BY tmnr, tjobnr"
                        'Response.write "<hr>medarbKundeoplysX(x, 17):"& medarbKundeoplysX(x, 17)  &" "& medarbKundeoplysX(x, 4) &":" & strSQLRealTotjobMid & "<br>"
                        
                        'Response.write strSQLRealTotjobMid & "<br>"
                        'Response.flush

                        oRec3.open strSQLRealTotjobMid, oConn, 3
                        if not oRec3.EOF then
                    
                        realTimerTotAktMedignPer = oRec3("sumTimerAktMedIgnPer") 

                        end if
                        oRec3.close 


                      
                      
                        realTimerTotJobMedignPerGrand = realTimerTotJobMedignPerGrand + realTimerTotAktMedignPer
                        realTimerTotJobMedignPer = realTimerTotAktMedignPer

                        realTimIngPerTxt = "<span style=""font-size:9px; color:#999999;"">"& formatnumber(realTimerTotAktMedignPer, 2) &"</span>"
                        
                        totFC_job_med_IsWrt = totFC_job_med_IsWrt & ",##"&x&"##"
                    
                    end if















        '*******************************************************************
        '***************** Ny job linje     ********************************
        '*******************************************************************


  
        '*** Vis kun akt. / job hvor forecast er overskrddet 
        'z = 900
        'if z = 900 then

          

        if cint(vis_job_fc_neg) = 0 OR (cint(vis_job_fc_neg) = 1 AND (cdbl(realTimerTotAktMedignPer) > cdbl(forcastTimerTotAktMedignPer))) then     
         '**efter medarbejder sum er skrevet           
          antalJobAktlinier = antalJobAktlinier + 1 
          'antalJobAktlinierPrJob = antalJobAktlinierPrJob + 1        
             
		
        if media <> "eksport" AND media <> "chart" then
            
                        if media <> "print" then
                            
                            call tilfojNyliniePaaJob(medarbKundeoplysX(x, 3), medarbKundeoplysX(x, 0))
                           

                        end if 'print
	    
                        %>	
        
                        <tr class="tr_medarb tr_medarb_<%=medarbKundeoplysX(x,3)%> tr_medarb_<%=medarbKundeoplysX(x,3)%>_<%=medarbKundeoplysX(x, 0)%>" style="visibility:hidden; display:none;">


                     <%
                         if lastxmid <> medarbKundeoplysX(x, 3) then
                         call medarbejderLinje(medarbKundeoplysX(x, 3),jobStartKri,datoInterval,rdimnb,1) 
                         
                         'lastxmid = medarbKundeoplysX(x, 3)
        
                         else
                         %>
                         <td style="background-color:#FFFFFF; padding:5px 5px 5px 5px; white-space:nowrap; color:#999999;" valign="top">
                             
                             <%
                                  if vis_simpel = 2 then
                                    
                                     call meStamdata(medarbKundeoplysX(x, 3))
        
      
                                 
                                    %>
                                  <i><%=meNavn%></i>
                                  <%end if %>  
                           

                             
                             &nbsp;</td>
                         <%
                         end if


                              
          %>





        <td bgcolor="#ffffff" style="padding:5px 5px 5px 5px; white-space:nowrap;" valign=top>
                        <%

            
                        '** SKAL Jobnavn og kundee navn skives igen (er det første linje på job pr. emdarb.)
                        if lastJid <> medarbKundeoplysX(x, 0) OR lastxmid <> medarbKundeoplysX(x, 3) OR (fob = 0 AND len(medarbKundeoplysX(x, 0)) <> 0) then%>
		
		
        
                        <%fob = 1%>
            

                        <%=medarbKundeoplysX(x, 5)%> (<%=medarbKundeoplysX(x, 6)%>) <!--&nbsp;<span class="btpush" style="padding:1px; float:right; background-color:#CCCCCC; font-size:9px;" id="btpush_<%=medarbKundeoplysX(x, 3) %>_<%=medarbKundeoplysX(x, 0) %>_<%=x%>_<%=y%>">>>|</span>--><br />

                        <b><%=medarbKundeoplysX(x, 2)%> (<%=medarbKundeoplysX(x, 1)%>)</b>  
            
                            <%
                             showJobst = 0   
                             select case lcase(medarbKundeoplysX(x, 25))
                             case "aktiv"
                                showJobst = 0
                             case "lukket"
                                jstaCol = "red"
                                showJobst = 1
                             case else
                                jstaCol = "#999999"
                                showJobst = 1
                             end select
                                
                                if cint(showJobst) = 1 then
                                %>
                                <span style="font-size:9px; color:<%=jstaCol%>;"><%=medarbKundeoplysX(x, 25) %></span>
                                <%end if%>
                            <br />
		
                        <%
                            if thisMedarbJoblist = "" then
                            thisMedarbJoblist = "<span id='sp_medarbjoblist_"&medarbKundeoplysX(x, 3)&"_"& medarbKundeoplysX(x, 0) &"' class='sp_medarbjoblist' style='color:#5582d2;'>" & medarbKundeoplysX(x, 2) & "</span>"
                            else
                            thisMedarbJoblist = thisMedarbJoblist &", <span id='sp_medarbjoblist_"&medarbKundeoplysX(x, 3)&"_"& medarbKundeoplysX(x, 0) &"' class='sp_medarbjoblist' style='color:#5582d2;'>"& medarbKundeoplysX(x, 2) & "</span>"
                            end if%>                    

		                <span style="font-size:9px; color:#999999;">
		                <%if len(trim(medarbKundeoplysX(x, 13))) <> 0 then %>
	                    <%=medarbKundeoplysX(x, 13) %>
		                <%end if %>
		
		                <%if len(trim(medarbKundeoplysX(x, 15))) <> 0 then %>
		                , <%=medarbKundeoplysX(x, 15) %>
		                <%end if %>
		                </span>
                    
                     
                        

                        
		
                        <%
                        else
                    
              
                        end if %>

                        <%
                            '**** Aktivitet ***'
                            if cint(visAkt) = 0 then 'viser ikke kolonne, men viser alligevel aktnavn hvis ikke linjen er uspec 

                            aktNavn = ""
                            if medarbKundeoplysX(x, 17) <> 0 then
                            strSQlAktnavn = "SELECT navn FROM aktiviteter WHERE id = "& medarbKundeoplysX(x, 17)
                            
                       
                            oRec6.open strSQlAktnavn, oConn, 3
                            if not oRec6.EOF then
                            aktNavn = oRec6("navn")
                            end if
                            oRec6.close


                            
                            %><br /><i><%=aktNavn %></i>
                            <%else %>
                            <i>Uspecificeret</i>
                            <%end if %>                

                        <%end if %>

           
        
            
        </td>

        
		

         <!-- Ramme -->
        <%if cint(visRamme) = 1 then %>
         
           <td bgcolor="#FFFFFF" valign=top style="padding:2px 2px 2px 2px; font-size:9px;">
               
              <%call arsRammeSub %>
               
               <%if media <> "print" then %>
               <input id="Text2" type="text" style="width:40px;" name="FM_aarsramme_timer"  value="<%=formatnumber(arrsNormThis, 0) %>" />
               <%else %>
               <%=formatnumber(arrsNormThis, 0) %>
               <%end if %>
               (år:<%=aarsRamme %>)

               <input id="Text6" type="hidden" name="FM_aarsramme_timer"  value="#" />
               <input id="nl_medid_<%=medarbKundeoplysX(x, 0)%>" type="hidden" name="FM_aarsramme_medarb"  value="<%=medarbKundeoplysX(x, 3) %>" />
               <input id="Hidden5" type="hidden" name="FM_aarsramme_rid"  value="<%=arrsIdThis %>" />
               <input id="nl_jobid_<%=medarbKundeoplysX(x, 0)%>" type="hidden" name="FM_aarsramme_jobid"  value="<%=medarbKundeoplysX(x, 0) %>" />
               <input id="Text5" type="hidden" name="FM_aarsramme_aar"  value="<%=aarsRamme %>" />

           <%thisMTotRamme = thisMTotRamme + arrsNormThis  %>
           </td>

         <%end if %>


        <%'else %>
       
        <!--<td bgcolor="#FFFFFF" valign=top style="padding:2px 2px 2px 2px;">&nbsp;</!--td> -->
      






        
		
        <%if len(trim(medarbKundeoplysX(x, 17))) <> 0 then
        medarbKundeoplysX(x, 17) = medarbKundeoplysX(x, 17)
        else
        medarbKundeoplysX(x, 17) = 0
        end if
         
            
            
        %>
		  
         

          <!-- Aktiviteter --->

            <%if cint(visAkt) = 1 then %>
          <td bgcolor="#FFFFFF" valign=top style="padding:2px 2px 2px 2px;">
                    <%
                    'WWF: aktive via timereg_usejob og faktuerbare
                    '*** PoSTIV aktivering af aktiviteter slået til
                  

               
                   


                    call hentaktiviteter(positiv_aktivering_akt_val, medarbKundeoplysX(x, 0), medarbKundeoplysX(x, 3), aty_sql_realhoursAkt)

                    'response.Write strSQLa & "<br>18:" & medarbKundeoplysX(x, 18)
                    'response.flush 
                    aSel = ""
                    aktText = ""
                    
                        if media <> "print" then%>
                    <select class="aaFM_jobid" id="aaFM_jobid_<%=x%>" name="sFM_aktid" style="width:100px;">
                    <option value="0">(uspecificeret)</option>
                    <option value="0">Vælg aktivitet..?</option>
                    <%
                    end if

                    oRec4.open strSQLa, oConn, 3
                    While not oRec4.EOF 

                     if ISNULL(oRec4("fase")) <> true AND len(trim(oRec4("fase"))) <> 0 then
                     fsNavn = " | fase: "& oRec4("fase")
                     else
                     fsNavn = ""
                     end if
                     
                     if cdbl(medarbKundeoplysX(x, 17)) = cdbl(oRec4("aid")) then
                     aSel = "SELECTED"
                     aktText = oRec4("aktnavn") & fsNavn
                     else
                     aSel = ""
                     end if
                     
                     
                      if media <> "print" then%>
                    <option value="<%=oRec4("aid")%>" <%=aSel %>><%=oRec4("aktnavn") &" "& fsNavn%></option>
                    <%end if

                    oRec4.movenext
                    wend
                    oRec4.close 
                    
                      if media <> "print" then%>
        
                    </select><br />
                    <%else %>
                    <%=aktText %>
                    <%end if %>
                    
                    
                    
                      <%if media <> "print" then 'AND cint(visAktiv) = 1 then %>
	                    	<a href="#" id="anl_a_<%=medarbKundeoplysX(x, 0)%>_<%=medarbKundeoplysX(x, 3) %>" class="rodlille">Tilføj ny linie +</a> &nbsp; 
		              <%end if%>
       
                    </td>
		            <%end if 'visAkt %>
		
	    <%end if 'media
		
         'if lto = "mmmi" then
         'Response.Write "numoffdaysorweeksinperiode" & numoffdaysorweeksinperiode & "<hr>"
		 'end if   

         


		'**************************************************************
		'***** Udkskriver forecast timer pr. md. på joblinje **********
		'**************************************************************
		
		'lastY = -1
        
            
    
		if y = 0 then
	    forcastTimerJobtot = 0
        end if			
					lastMonth = monththis
					y = 0
					for y = 0 TO numoffdaysorweeksinperiode
					    

					    'if lto = "mmmi" then
					    'Response.Write periodeSel &","& y &","& startdato &","& monthUse & ""& numoffdaysorweeksinperiode  &"<br>"
                        'end if
						
						call antaliperiode(periodeSel, y, startdato, monthUse)
					    'Response.Write "<br>Nday: "& nday 
								    
									timerThis = 0
                                    procThis = 0
                                    n = 0
                                    'divider = 0
                                    'foundone = 0	
									For n = 0 to n_arr(medarbKundeoplysX(x, 3)) 'medarbKundeoplysX(x, 17) + 1 ' AND foundone = 0 'UBOUND(midjiddato)
									


								    'Response.Write "n: "& n &" "& medarbKundeoplysX(x, 3) & "<br>"
									'***** Uge visning, tjkker ikke MD ****'
									if periodeSel = 3 then
									
								        if (midjiddato(medarbKundeoplysX(x, 3),n, 4) = newWeek) AND _
								          midjiddato(medarbKundeoplysX(x, 3),n, 1) = medarbKundeoplysX(x, 0) AND _
                                         midjiddato(medarbKundeoplysX(x, 3),n, 5) = medarbKundeoplysX(x, 17) AND _
                                          midjiddato(medarbKundeoplysX(x, 3),n, 2) = medarbKundeoplysX(x, 3) AND _
								           datepart("yyyy", midjiddato(medarbKundeoplysX(x, 3),n, 3)) = datepart("yyyy", nDay) then
    								        
								            timerThis = midjiddato(medarbKundeoplysX(x, 3),n, 0)
                                            procThis = midjiddato(medarbKundeoplysX(x, 3),n, 6)
                                           
    										
										    'foundone = 1
    										
    										
									    end if
								    
								    
								    else
								    '*** MD visning tjekker ikke om uge passer **'

                                  
								    
								            if (datepart("m", midjiddato(medarbKundeoplysX(x, 3),n, 3), 2,2) = datepart("m", nDay, 2,2)) AND _
								                midjiddato(medarbKundeoplysX(x, 3),n, 1) = medarbKundeoplysX(x, 0) AND midjiddato(medarbKundeoplysX(x, 3),n, 5) = medarbKundeoplysX(x, 17) AND  _
                                                midjiddato(medarbKundeoplysX(x, 3),n, 2) = medarbKundeoplysX(x, 3) AND _
								                datepart("yyyy", midjiddato(medarbKundeoplysX(x, 3),n, 3)) = datepart("yyyy", nDay) then
								        
								                timerThis = midjiddato(medarbKundeoplysX(x, 3),n, 0)
                                                procThis = midjiddato(medarbKundeoplysX(x, 3),n, 6)
                                                'Response.Write "<br>timerThis" & timerThis
										      
										        'foundone = 1
										        'dividerThisMth( = divider + 1 
										      
									 
								            end if
								    
								
								    end if
									
                                   

									
									loops = loops + 1
									next
									


                                    if periodeSel <> 3 then 'AND cint(medarbKundeoplysX(x, 24)) = 1  'fra uge til gns pr. md medarbKundeoplysX(x, 24) = 1 = ANGIVET SOM UGE
                                    divider = dividerThisMth(datepart("m", nDay, 2,2)) '+ 1
                                    procThis = formatnumber((procThis/divider),0)
                                    else
                                    procThis = formatnumber((procThis),0)
                                    end if
                                   

                                    'response.write "procThis: "& procThis & " divider: "& divider & "<br>"

                                    procThis = replace(procThis, ".", "")
									
									if media <> "eksport" AND media <> "chart" then 
									Response.flush
									end if
									
									'loops = 0
								    'timerThis = testarr(datepart("ww", nday, 2,2),right(datepart("yyyy", nday), 2),medarbKundeoplysX(x, 0),medarbKundeoplysX(x, 3)) 
									'timerthis = midjiddato(n, 0)
								
						'foundone = 0
						xx = 0
					    
					

                   
					
					if media <> "eksport" AND media <> "chart" then
					
					if timerThis = 0 then
					bgTimer = "#ffffff"
					else
                    bgTimer = "#DCF5BD"
                    end if
				
					end if
					
					lastMonth = newMonth
					dato = nday
					stThisDato = dato
					txtwidth = 60
					
                    if cint(showasproc) = 1 then 'vis som procent

                        if procThis = 0 then
					    showTimerThis = ""
					    else
					    showTimerThis = procThis
					    end if

                    
                    else
					
				        if timerThis = 0 then
					    showTimerThis = ""
					    else
					    showTimerThis = timerThis
					    end if

                    end if
					
					if media <> "eksport" AND media <> "chart" then
					%>
					<td valign=top bgcolor="<%=bgTimer%>" class=lille style="padding:2px 2px 2px 2px; width:<%=txtwidth%>px;">
					<%
					end if
					
					'*** Realiserede timer /medarbejder på job pr. md
					
					sumTimer = 0
					
					if periodeSel <> 3 then
					mmwwSQLkri = " MONTH(tdato) = "& datepart("m", nday, 2, 2) 
					else
					mmwwSQLkri = " WEEK(tdato, 1) = "& datepart("ww", nday, 2, 2) 
					end if

                    'nulstiller aktSQLkriExcludeAkOnUspec 
                    'if medarbKundeoplysX(x, 17) = 0 then 'Uspecificeret akt. (derfor vigtigt soretering akt DESC på main SQL)
                    'aktSQLkriExcludeAkOnUspec = ""
                    'end if

                     if medarbKundeoplysX(x, 1) <> fc_lastjobnrWrt then
                        aktSQLkriExcludeAkOnUspec = ""
                     end if

                    if medarbKundeoplysX(x, 17) <> 0 then 'specifik akt.
                    aktSQLkri = " AND taktivitetid = "& medarbKundeoplysX(x, 17)
                     
                        if medarbKundeoplysX(x, 1) <> fc_lastjobnrWrt then
                                aktSQLkriExcludeAkOnUspec = aktSQLkriExcludeAkOnUspec & " AND (taktivitetid <> "& medarbKundeoplysX(x, 17)
                        else
                                if medarbKundeoplysX(x, 17) <> fc_lastaktidWrt then
                                aktSQLkriExcludeAkOnUspec = aktSQLkriExcludeAkOnUspec & " AND taktivitetid <> "& medarbKundeoplysX(x, 17)
                                end if
                        end if
                        
                    else
                    aktSQLkri = ""
                        if medarbKundeoplysX(x, 17) = 0 AND aktSQLkriExcludeAkOnUspec <> "" AND instr(aktSQLkriExcludeAkOnUspec, ")") = 0 then
                        aktSQLkriExcludeAkOnUspec = aktSQLkriExcludeAkOnUspec & ")" 'trim til klar
                       
                        'aktSQLkriExcludeAkOnUspec = "" 'trim nulstillet
                        end if
                    end if

                   
										
                        					
				    strSQLtimer = "SELECT sum(timer) AS sumtimerFaktisk FROM timer WHERE tmnr = "& medarbKundeoplysX(x, 3) &" AND tjobnr = '"& medarbKundeoplysX(x, 1) &"' "& aktSQLkri 
                       
                         if medarbKundeoplysX(x, 17) = 0 then 'specifik akt.
                        strSQLtimer = strSQLtimer & aktSQLkriExcludeAkOnUspec 
                        end if  
                    
                        
                    strSQLtimer = strSQLtimer & " AND ("& aty_sql_realhours &") AND (YEAR(tdato) = "& datepart("yyyy", nday, 2, 2) &" AND " & mmwwSQLkri &")"
				    'Response.Write strSQLtimer
				    'Response.Flush
				    'Response.end
				    
				    oRec3.open strSQLtimer, oConn, 3
				    while not oRec3.EOF
    				
				    sumTimer = oRec3("sumtimerFaktisk")
    				
				    oRec3.movenext
				    wend
    				
				    oRec3.close
    				
    				

    				'Response.Write sumTimer & "<br>"
    				
				    if len(trim(sumTimer)) <> 0 then
				    sumTimer = sumTimer
				    else
				    sumTimer = 0
				    end if
				    
				    if len(trim(timerThis)) <> 0 then
				    timerForecast = timerThis 'showTimerThis
				    else
				    timerForecast = 0
				    end if
				    

                    if len(trim(procThis)) <> 0 then
				    procForecast = procThis/1 
				    else
				    procForecast = 0
				    end if
				    
				    '*** Afvigelse %
				    
				        if cdbl(timerForecast) <= cdbl(sumTimer) then
						    if cdbl(sumTimer) <> 0 then
                            afvThis = 100 - (timerForecast/sumTimer) * 100
                            else
                            afvThis = 0
                            end if
                        else
                            if cdbl(timerForecast) <> 0 then
                            afvThis = 100 - (sumTimer/timerForecast) * 100
                            'afvThis = replace(afvThis, "-", "")
                            else
                            afvThis = 0
                            end if
                        end if
				    
				   
				   'Response.Write "timerForecast:" & timerForecast & "<br>"
				   'Response.Write "sumTimer: " & sumTimer & "<br>"
				   
				   	if media <> "print" AND media <> "eksport" AND media <> "chart" then
				   	
				   	            '*** Er det admin eller jobansvarlig ****'
                                '*** Bør allerede være fundet i søgefilteret **
            				   	
				   	            'if (cint(session("mid")) = cint(medarbKundeoplysX(x, 3)) OR _
				   	            'session("mid") = medarbKundeoplysX(x, 11) OR _
				   	            'session("mid") = medarbKundeoplysX(x, 12)) AND _

				   	            if ((month(cdate(dagsdato)) <= month(cdate(nday)) AND year(cdate(dagsdato)) = year(cdate(nday))) OR (year(cdate(dagsdato)) < year(cdate(nday)))) OR level = 1 then
				   	            iptype = "text"
				   	            else 
				   	            iptype = "hidden"
				   	            end if
				   	            %> 
                				
                				
    				            <input type="<%=iptype%>" name="FM_timer" id="FM_timer_<%=medarbKundeoplysX(x, 3) %>_<%=medarbKundeoplysX(x, 0) %>_<%=x%>_<%=y%>" style="width:50px;" value="<%=showTimerThis %>">

                               
                                <%if iptype = "text" then%>
                                
                                <!--<input class="bt" id="bt_<%=medarbKundeoplysX(x, 3) %>_<%=medarbKundeoplysX(x, 0) %>_<%=x%>_<%=y%>" type="button" value=">>" style="font-size:8px;"  />-->

                                <br /><span class="bt" style="padding:1px; background-color:#CCCCCC; font-size:8px;" id="bt_<%=medarbKundeoplysX(x, 3) %>_<%=medarbKundeoplysX(x, 0) %>_<%=x%>_<%=y%>">>></span>
                                
                              

    				            <%else %>
    				            <%=showTimerThis %>
                                
    				            <%end if %>
                                
    				            <!--'*** hidden array felt. *** Bruges så der kan skelneshvis man bruger komma.-->
					            <input type="hidden" name="FM_timer" id="FM_timer" value="#">
					<%else 
					
					        if media <> "eksport" AND media <> "chart" then%>
					         <%=showTimerThis %> 
					        <%end if %>
					        
					<%end if 'Print / media%>


					
					
					<%if media <> "eksport" AND media <> "chart" then %>
					
					
					
					<%if sumTimer <> 0 then %>
                        <span style="font-size:9px; color:#5582d2;">
                            
                         <%if cint(showasproc) = 1 then %>
                                  f:  <%=formatnumber(timerThis, 0) %> t. <!-- (<=formatnumber(procThis, 0) %>%) --> / 
                                <%end if%>

					r: <%=formatnumber(sumTimer, 0)%> t.  <!--afv.: <%=formatnumber(afvThis, 0)%>%-->

                        </span>
					<%end if %>
				
                    	</td>
					<%
					
					end if 'media


                        '** Normtimer i periode til eksport
                        if media = "eksport" then


                            if lastxmid <> medarbKundeoplysX(x, 3) then
                              '** NormTimer pr medarb pr uge / md
                                'if periodesel = "6" OR periodesel = "12" then 
                                
                               

                                'call stKriInterval612(y, jobStartKri)

                                'response.write "lastxmid: "& lastxmid &" medarbKundeoplysX(x, 3)" & medarbKundeoplysX(x, 3) & " y: "& y &" jobStartKriThisM: "& jobStartKriThisM &" datoIntervalThisM: "& datoIntervalThisM &"<br>"
                                'response.flush
                            
	                            'call normtimerPer(medarbKundeoplysX(x, 3), jobStartKriThisM, 6, 0)
                                'normTimerMedarb = formatnumber(ntimPer*4,5, 0)        
                               
                                'response.write "normTimerMedarb: "& normTimerMedarb 
                                'response.flush
                                'else 
                            
                                'if y = 0 then
                                'jobStartKriNormT = day(jobStartKri) &"-"& month(jobStartKri) &"-"& year(jobStartKri)
                                'end if
                                'jobStartKri = dateAdd("d", y*7, jobStartKriNormT)
                                '** Altid pr. uge i udtræk, da det er det mest faste hodlepunkt mm. man henter en kolonne ud for hver md.
                                call normtimerPer(medarbKundeoplysX(x, 3), jobStartKri, 6, 0) 
                                normTimerMedarb = formatnumber(ntimPer, 1)  
                              
                          
                                'end if 

                            end if
                        end if


                    if (media = "eksport" AND vis <> 1)  then
                         
                         if lastxmid <> medarbKundeoplysX(x, 3) then
                           

                          
                            '*** Bregner pr. uge til at omregne mellem timer og %
                            '*** Er nødt til at øre det igen, da vi ovenfor berenger på både pr. md eller ugfe alt efter peridoevalg. herunder SKAL det altid være pr. uge
                             call normtimerPer(medarbKundeoplysX(x, 3), jobStartKri, 6, 0) 

                         end if

                        if lastxmid <> medarbKundeoplysX(x, 3) OR lastJid <> medarbKundeoplysX(x, 0) then
                            call arsRammeSub
                        end if
                    end if



					

                    if media = "eksport" then

                           

                    
                            if cint(vis) = 1 then 'Alm 


                                        if (lastAktJidExpVis_1 <> medarbKundeoplysX(x, 0) &"_"& medarbKundeoplysX(x, 17) &"_"& medarbKundeoplysX(x, 3)) then
                                        
                                        csvTxt = csvTxt & vbcrlf & medarbKundeoplysX(x, 5)&";"& medarbKundeoplysX(x, 6) &"" _
                                        &";" & medarbKundeoplysX(x, 2) & ";"& medarbKundeoplysX(x, 1) & ";"
                                        
                                        'if cint(visAktiv) = 1 then ALTID med da der kan væres flere linier på samme job ved uspec. Dette skal kunne ses i udtrækkket.
                                        csvTxt = csvTxt & medarbKundeoplysX(x, 18) & ";" & medarbKundeoplysX(x, 17) & ";" 
                                        'end if                

                                        csvTxt = csvTxt & medarbKundeoplysX(x, 4) &" ("& medarbKundeoplysX(x, 7) &");" _
					                    & medarbKundeoplysX(x, 8) &";" _ 
					                    & medarbKundeoplysX(x, 9) &";" _ 
					                    & normTimerMedarb


                      

                                        lastAktJidExpVis_1 = medarbKundeoplysX(x, 0) &"_"& medarbKundeoplysX(x, 17) &"_"& medarbKundeoplysX(x, 3)
                                       

                                        end if
                                        
                                      
                                         csvTxt = csvTxt &";"& showTimerThis 

                                        if cint(expvisreal) = 1 then
                                        
                                        if len(trim(sumTimer)) <> 0 then
                                        sumTimerExp = formatnumber(sumTimer,2)
                                        else
                                        sumTimerExp = ""
                                        end if

                                        csvTxt = csvTxt &";"& sumTimerExp
					                    end if

                            else 'PIVOT 
                                    

                                        csvTxt = csvTxt & vbcrlf & medarbKundeoplysX(x, 4) &" ("& medarbKundeoplysX(x, 7) &")"
					                    csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 8) 
					                    csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 9) &";" _ 
					                    & normTimerMedarb 
					

                                        'if (lastxmid <> medarbKundeoplysX(x, 3) OR lastJid <> medarbKundeoplysX(x, 0)) AND cint(isArsNormWrtEksp) = 0 then

                                        if cint(visRamme) = 1 then
                                        'csvTxt = csvTxt & ";"& formatnumber(ntimPer, 2) 
                                        csvTxt = csvTxt & ";"& aarsRamme & ";" & arrsNormThis
                                        isArsNormWrtEksp = 1
                                        end if

                                        'else
                                        'csvTxt = csvTxt & ";;;" 
                                        'end if

					                    csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 2) & ";" & medarbKundeoplysX(x, 1)

                                        'if cint(visAkt) = 1 then ALTID med i udtræk der kan være flere linier på ammme job. Skal fremgå
                                        csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 18) & ";" & medarbKundeoplysX(x, 17) 
					                    'end if                

					                    csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 13) & ";" & medarbKundeoplysX(x, 14)
					                    csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 15) & ";" & medarbKundeoplysX(x, 16)
					
					                    csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 5) & ";" & medarbKundeoplysX(x, 6)
					
					                    'if lto <> "execon" then
					                    csvTxt = csvTxt & ";"& csvDatoerUge(y) 
					                    'end if
					
					                    csvTxt = csvTxt & ";"& csvDatoerMD(y) &";"& csvDatoerAR(y) &";"& timerForecast 
                        
                                        if cint(visStatus) = 1 then
                                        csvTxt = csvTxt &";"& sumTimer &";"& formatnumber(afvThis, 0)
                                        end if
                                       

					                    'csvTxt = csvTxt & vbcrlf & ";;;;;;;"& csvDatoer(y) &";" & timerThis


                                      
                                      

                                      

                    
                                        
                            end if 'vis

                    end if 'media
                        
                      

                          
                         'csvTxt = csvTxt &";"&  medarbKundeoplysX(x, 0) &"_"& medarbKundeoplysX(x, 17) &"_"& medarbKundeoplysX(x, 3) &"   x: "& x &" y: "& y &":"& showTimerThis & "<br>"
                         'csvTxt = csvTxt &";"& showTimerThis 

                                          
                        %>


                   

					
					<%if media <> "eksport" AND media <> "chart" then  %>
					
                    <input type="hidden" class="ahFM_jobid_<%=x%>" name="FM_aktid" id="Hidden4" value="<%=medarbKundeoplysX(x, 17)%>">
                    <input type="hidden" name="FM_aktid_old" id="Hidden6" value="<%=medarbKundeoplysX(x, 17)%>">
					<input type="hidden" name="FM_jobid" id="Hidden3" value="<%=medarbKundeoplysX(x, 0)%>">
					<input type="hidden" name="FM_medarbid" id="FM_medarbid" value="<%=medarbKundeoplysX(x, 3)%>">
					<input type="hidden" name="FM_dato" id="FM_dato" value="<%=stThisDato%>">
                     
					
					
				
					
					<%
					end if

                      
                    fc_lastaktidWrt = medarbKundeoplysX(x, 17)
					fc_lastjobnrWrt = medarbKundeoplysX(x, 1)
					
					medarbTotalTimer(y) = medarbTotalTimer(y) + timerForecast
					timerTotalTotal(y) = timerTotalTotal(y) + timerForecast

                    medarbTotalProc(y) = medarbTotalProc(y) + procForecast
					procTotalTotal(y) = procTotalTotal(y) + procForecast

                    medarbTimerReal(y) = medarbTimerReal(y) + sumTimer
					realTotalTotal(y) = realTotalTotal(y) + sumTimer
					
					timerMJobTotal(y) = timerMJobTotal(y) + timerForecast
                    procMJobTotal(y) = procMJobTotal(y) + procForecast
                    
					realMJobTotal(y) = realMJobTotal(y) + sumTimer

                    '** Jobtotal
                    realTimerJobtot = realTimerJobtot + sumTimer 
					forcastTimerJobtot = forcastTimerJobtot + timerForecast
                    forcastProcJobtot = forcastProcJobtot + procForecast 

                    '** Medarbejder total
                    forecastProcMJobtotGrand = forecastProcMJobtotGrand + procForecast
                    forecastTimerMJobtotGrand = forecastTimerMJobtotGrand + timerForecast
				    realTimerMJobtotGrand = realTimerMJobtotGrand + sumTimer  
                    
                    lastY = y
                  

					next 'y periode

                       
                    
                    





                    saldoJobTot = (forcastTimerJobtot - realTimerJobtot)
                    saldoJobTotIgnper = (forcastTimerTotAktMedignPer - realTimerTotAktMedignPer)


					
					if media <> "eksport" AND media <> "chart" then 
                        
                    if cint(visStatus) = 1 then%>
                    
                    <td bgcolor="#ffffff" valign=top class=lille align=right style="padding:2px 2px 2px 2px; white-space:nowrap;"><%=formatnumber(medarbKundeoplysX(x, 20), 0) %><br />
                       <span style="font-size:10px; color:#999999;"><%=formatnumber(medarbKundeoplysX(x, 19), 0) %> DKK<br /></span>

                        <span style="font-size9px; color:#999999;">
                        <%if medarbKundeoplysX(x, 20) <> 1 then %>
                        rest.: <%=formatnumber(medarbKundeoplysX(x, 21), 0) %> t.
                        <%else %>
                        afsl.:<%=formatnumber(medarbKundeoplysX(x, 21), 0) %> %
                        <%end if %>

                        </span>
                    </td>

                    <td bgcolor="#ffffff" valign=top class=lille align=right style="padding:2px 2px 2px 2px;"><b><%=formatnumber(forcastTimerJobtot, 0) %></b> <!-- (<%=formatnumber(forcastProcJobtot/dividerAll, 0) %>%) --><br />
                    <span style="font-size:9px; color:#999999;"><%=formatnumber(forcastTimerTotAktMedignPer, 0) %></span></td>
			        <td bgcolor="#ffffff" valign=top class=lille align=right style="padding:2px 2px 2px 2px;"><%=formatnumber(realTimerJobtot, 2)%><br />
                    <%=realTimIngPerTxt %>
                   </td>	
		            <td bgcolor="#ffffff" valign=top class=lille align=right style="padding:2px 2px 2px 2px;">

                    <%if saldoJobTot > 0 then
                    bgColSaldoTot = "#6CAE1C" 
                    else
                    bgColSaldoTot = "darkred"
                    end if%>


                    <%if saldoJobTotIgnper > 0 then
                    bgColSaldoTotig = "#6CAE1C" 
                    else
                    bgColSaldoTotig = "darkred"
                    end if%>

                    <span style="color:<%=bgColSaldoTot%>;">
                    <b><%=formatnumber(saldoJobTot, 0) %></b></span><br />
                    <span style="font-size:9px; color:<%=bgColSaldoTotig%>;"><%=formatnumber(saldoJobTotIgnper, 0) %></span></td>
                    <%
                    end if 'visStatus
            		
		            end if




                    if media = "eksport" AND vis = "1" then
                    
                        if forcastTimerJobtot <> 0 then
                        forcastTimerJobtotExpTxt = forcastTimerJobtot
                        else
                        forcastTimerJobtotExpTxt = ""
                        end if

                        if forcastProcJobtot <> 0 then
                        forcastProcJobtotExpTxt = forcastProcJobtot
                        else
                        forcastProcJobtotExpTxt = ""
                        end if

                        if forcastTimerTotAktMedignPer <> 0 then
                        forcastTimerTotAktMedignPerExpTxt = forcastTimerTotAktMedignPer
                        else
                        forcastTimerTotAktMedignPerExpTxt = ""
                        end if


                        if realTimerJobtot <> 0 then
                        realTimerJobtotExpTxt = realTimerJobtot
                        else
                        realTimerJobtotExpTxt = ""
                        end if


                        if realTimerTotAktMedignPer <> 0 then
                        realTimerTotAktMedignPerExpTxt = realTimerTotAktMedignPer
                        else
                        realTimerTotAktMedignPerExpTxt = ""
                        end if

                                        if cint(visStatus) = 1 then 

                                         csvTxt = csvTxt  &";"& formatnumber(medarbKundeoplysX(x, 20), 0)
                                         csvTxt = csvTxt  &";"& formatnumber(medarbKundeoplysX(x, 19), 0)

                                         csvTxt = csvTxt & ";"& forcastTimerJobtotExpTxt
                                         csvTxt = csvTxt & ";"& forcastTimerTotAktMedignPerExpTxt
                                         csvTxt = csvTxt & ";"& realTimerJobtotExpTxt
                                         csvTxt = csvTxt & ";"& realTimerTotAktMedignPerExpTxt
                                         csvTxt = csvTxt & ";"& saldoJobTot
                                         csvTxt = csvTxt & ";"& saldoJobTotIgnper
                                    
                                         end if

                    end if

          
          
          
		                            if media <> "eksport" AND media <> "chart" then%>
		
		                            </tr>
                                    <%
                                          'Response.write "<tr><td colspan=1000>herT<br></td></tr>"
        
       
                                    end if 'Media          
        
        
         end if 'Forecast overskreddet
        
         'Response.write "<tr><td colspan=1000>herS<br></td></tr>"

        
        forcastTimerTotJobMedignPer = forcastTimerTotJobMedignPer + forcastTimerTotAktMedignPer
                    
        realTimerJobtotGrand = realTimerJobtotGrand + realTimerJobtot 

		forecastTimerJobtotGrand = forecastTimerJobtotGrand + forcastTimerJobtot
        forecastProcJobtotGrand = forecastProcJobtotGrand + forcastProcJobtot

        forcastTimerTotJobMedignPerGrand = forcastTimerTotJobMedignPerGrand + forcastTimerTotAktMedignPer 'forcastTimerTotJobMedignPer

        pageForecastGrandtot = pageForecastGrandtot + forcastTimerJobtot 
		pageRealGrandtot = pageRealGrandtot + realTimerJobtot


         '*** Grand Grand total i bunde ALLE medarbejdere / Alle job **'
        fcGrandGrand = fcGrandGrand + forcastTimerJobtot
        fpGrandGrand = fpGrandGrand + forcastProcJobtot
        fcignPerGrandGrand = fcignPerGrandGrand + forcastTimerTotAktMedignPer
        realGrandGrand = realGrandGrand + realTimerJobtot 
        realignPerGrandGrand = realignPerGrandGrand + realTimerTotAktMedignPer

       jobBgtGrandGrand = jobBgtGrandGrand +  medarbKundeoplysX(x, 19)
       jobTimerGrandGrand = jobTimerGrandGrand +  medarbKundeoplysX(x, 20)
        
        
        realTimerJobtot = 0 
		forcastTimerJobtot = 0
        forcastProcJobtot = 0
	    
        
         
        	

        'strjIdWrt = strjIdWrt & "#"&medarbKundeoplysX(x, 0)&"#," 
		lastJid = medarbKundeoplysX(x, 0)
	    lastJid_Aid = medarbKundeoplysX(x, 0) &"_"& medarbKundeoplysX(x, 17)

	

        lastJobans1 = medarbKundeoplysX(x, 11)
		lastJobans2 = medarbKundeoplysX(x, 12)
		

        
      

        'if cint(vis_job_fc_neg) = 0 OR (cint(vis_job_fc_neg) = 1 AND cdbl(realTimerTotAktMedignPer) > cdbl(forcastTimerTotAktMedignPer)) then



		

        'antanlJoblinierprM = antanlJoblinierprM + 1 
       
		
        end if '** Lastjobid
        
        'y = y + 1	
		'Response.Flush

       
        'linejCsvTxt = linejCsvTxt & csvTxt
        'csvTxt = ""
		
        if lastxmid <> medarbKundeoplysX(x, 3) then
        lastxmid = medarbKundeoplysX(x, 3)
        end if

        'response.write "lastxmid" & lastxmid

        next 'medarb loop
        
        'response.write "VIS: "& vis &"linejCsvTxt: "& csvTxt
                                            

        
		
		'response.end
        Response.flush
		
		







        '*************************************************************************************************
        '********** Grandtotaler i bunden ****************************************************************
        '*************************************************************************************************

       

		
		if media <> "eksport" then
		

        if cint(vis_job_fc_neg) <> 1 then

		'*****************************************
		'*** Tildel timer  / Guiden aktive job****
		'*****************************************
		
	
		
				
				
				        if media <> "chart" then 
				           
                                call jobtotalprmedarb
                           
                           
				        
                            if cint(jobidsel) = 0 then
                                '**** Medarb. total ***'
                                'if cint(vis_simpel) <> 1 then
                                mTotthisMid = lastxmid
                                call medarbTotal
                                'end if
                            end if
                            
                            'if cint(vis_simpel) <> 1 then
                            call grandtotal
                            'end if

                        end if

                        

                        if cint(jobidsel) = 0 then

                            if media = "chart" then
                            
                                    'Response.write "her"

                                    call meStamdata(lastxmid)
                                    csvTxtA = csvTxtA & meTxt 

                                    for y = 0 TO numoffdaysorweeksinperiode
                                    csvTxtA = csvTxtA & ";"& medarbTotalTimer(y)
                                    medarbTotalTimer(y) = 0 
                                    next

                                    csvTxtA = csvTxtA & "xx99123sy#z"
                            end if

				        end if


                  



                                      

				if media <> "print" AND media <> "chart" AND cint(hideEmptyEmplyLines) <> 1 then
				   

                '** Henter medarbejder linier, (til at oprette forecast) på valgte medarbejder der ikke er vist på siden ***'
                for m = 0 to UBOUND(intMids)
                     

                     'Response.Write "medarbwrt: " & medarbwrt & " intMids(m): "& intMids(m) & "<br>"

                     if instr("#"&medarbwrt&"#", intMids(m)) = 0 then
                     'call tilfojNylinie(intMids(m))
                      call tilfojNylinie(intMids(m), visRamme, visAktiv)
                      call medarbejderLinje(intMids(m),jobStartKri,datoInterval,rdimnb,0)
                     end if
                
                next

                end if 'Print


				
				   
				    

                  if media <> "chart" then
				
				        if medarbSel = "0" then
				        %>
				        <!--- Grandtotal --->
				
				        <tr>
					        <td height=20 bgcolor="#ffffff" align=right style="padding:2px 5px 2px 2px;"><b>Grandtotal forecast:</b><br />
					        Real.</font></td>
                            <td bgcolor="#ffffff">&nbsp;</td>
					
					        <%
					
						        for y = 0 TO numoffdaysorweeksinperiode
						    
						        '*** Afvigelse ***
			                    tottotAfv = 0
    			
    				    
				                if timerTotalTotal(y) >= realTotalTotal(y) then
                                    if timerTotalTotal(y) <> 0 then
                                    tottotAfv = 100 - (realTotalTotal(y)/timerTotalTotal(y)) * 100
                                    else
                                    tottotAfv = 0
                                    end if
                                else
                                    if realTotalTotal(y) <> 0 then
                                    tottotAfv = 100 - (timerTotalTotal(y)/realTotalTotal(y)) * 100
                                    else
                                    tottotAfv = 0
                                    end if
                                end if
						
						        %>
							
						        <td bgcolor="#ffffff" class=lille style="padding:2px 5px 2px 2px;"><b><%=formatnumber(timerTotalTotal(y),0)%></b> <br /> <%=formatnumber(realTotalTotal(y), 0)%> ~ <%=formatnumber(tottotAfv,0) %>%</td>
						        <%
						        next
						        %>
						        <td bgcolor="#ffffff" style="padding:2px 5px 2px 2px;" align=right><b><%=formatnumber(pageForecastGrandtot,0) %></b></td>
						        <td bgcolor="#ffffff" style="padding:2px 5px 2px 2px;" align=right><%=formatnumber(pageRealGrandtot,0) %></td>
				                <td bgcolor="#ffffff" style="padding:2px 5px 2px 2px;" align=right>&nbsp;</td>
                        </tr>
				        <%
				        end if 'medarbSel
				
				
		        %>
		        </table>
                <%end if 'chart%>


		
   <%
   else
   %>
      </table>
   <%end if 'cint(vis_job_fc_neg) <> 1 %>


		                <%if media <> "print" AND media <> "chart" then%>
		                        <table cellspacing=0 cellpadding=0 border=0 width=100%>
		                        <tr><td align=right style="padding:10px 30px 5px 0px;">
		                        <input type="submit" value="Indlæs forecast >> "></td></tr>
                                </table>
                        <input id="ymax" name="ymax" value="<%=y-1 %>" type="hidden" />
                        </form> 
                        
*) Normtimer er korrigeret for ferie, barsel og anden planlagt fravær.<br /><br />
f: forecast<br />
r: realiseret (indtastet på job/projekter)<br />
s: saldo<br />
n: norm.<br /><br />

Antal linier: <%=antalJobAktlinierGrand %><br />

                                    
                        
                        
                        
                        <%end if %>
            
            
<%end if '** eksport %>



       
</div>



		
	
	 

	 
	 
	 <%if media <> "eksport" AND media <> "chart" then %>
	 <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />&nbsp;
	 
	 
	 <%end if
				
		
		
		'***************** Eksport *************************'
        if media = "eksport" OR media = "chart" then
               


              
		
        if vis <> 1 then 'PIVOT
		            
                     csvTxtTop = "Medarbejder;Initialer;Sidst opd.;" 

                      select case periodesel 
                        case 6,12
                        csvTxtTop = csvTxtTop & "Norm. timer pr. md;" 
                        case else
                        csvTxtTop = csvTxtTop & "Norm. timer pr. uge;"
                        end select

		
                     if cint(visRamme) = 1 then 
                     csvTxtTop = csvTxtTop & "Årsramme år;Årsramme Timer;"
                     end if

                      csvTxtTop = csvTxtTop & "Job;Job nr.;"

                     'if cint(visAkt) = 1 then ALTID MED I UDTRÆK
                     csvTxtTop = csvTxtTop & "Aktnavn;Akt. id;"
                     'end if

                    csvTxtTop = csvTxtTop & "Jobansvarlig;Init;Jobejer;Init;Kunde;Kunde nr.;"
                    csvTxtTop = csvTxtTop & "Uge;"
                    csvTxtTop = csvTxtTop &"Måned;År;Forecast;"
                    
                    if cint(visStatus) = 1 then
                    csvTxtTop = csvTxtTop &"Realiseret;Afvigelse Forecast/Real. %"
                    end if

        else '** Alm eksport

       

            if media = "chart" then
                csvTxtTop = "Medarbejder;" 
                csvTxtTop = csvTxtTop & csvTxtTopMd & "xx99123sy#z"
            
            else
                csvTxtTop = "Kunde;Kunde nr.;Job;Job nr.;"
            
                'if cint(visAkt) = 1 then ALTID MED I UDTRÆK
                csvTxtTop = csvTxtTop & "Aktnavn;Akt. id;"
                'end if

                csvTxtTop = csvTxtTop & "Medarbejder;Initialer;Sidst opd.;" 

                'select case periodesel 
                'case 6,12
                'csvTxtTop = csvTxtTop & "Norm. timer pr. md (gns.);" 
                'case else
                csvTxtTop = csvTxtTop & "Norm. timer pr. uge (gns.);"
                'end select

                csvTxtTop = csvTxtTop & csvTxtTopMd 

                if cint(visStatus) = 1 then
                csvTxtTop = csvTxtTop & "Jobbudget timer;Jobbudget beløb;Forecast;Forecast (total, uanset valgt periode);Realiseret;Realiseret (total, uanset valgt periode); Saldo; Saldo (total, uanset valgt periode) "
                end if   
         
            end if
        
        
        end if

	

               if media = "chart" then
               
               ekspTxt = csvTxtTop & csvTxtA 
               ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
               else
               ekspTxt = csvTxtTop & csvTxt & csvTxtTotal
               ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
               end if

                'csvTxt
               'kspTxt = replace(ekspTxt, ";", chr(34))
               'datointerval = request("datointerval")
                       
                
                call TimeOutVersion()
                       
                filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	            filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
            	
	            Set objFSO = server.createobject("Scripting.FileSystemObject")
            	
	            if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\ressource_belaeg_jbpla.asp" then
            							
		            Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\"& toVer &"\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
		            Set objNewFile = nothing
		            Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\"& toVer &"\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)

                 
            	
	            else
            		
		            Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, false)
		            Set objNewFile = nothing
		            Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8, -1)
            		
	            end if
            	
	            file = "ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
            	
            	
	            objF.WriteLine(ekspTxt)


                                          

                                             %>
                                            <form>
                                            <input type="hidden" value="<%=file %>" name="FM_filename" id="FM_filename" />
                                            </form>
                                            <%
            	
            	
	                   objF.close

                if media = "chart" then%>

                                            
                <script src="inc/res_hc_jav.js"></script>

                   
                   <%
                   maddVal = 1150
                   for m = 0 to UBOUND(intMids) 
                   maddVal = maddVal + 50 
                   next%>
                    <div id="container" style="width: <%=maddVal%>px; height: 400px; border:0px;"></div>

                     
                                             

                <%else

	            %>
	            <table border=0 cellspacing=1 cellpadding=0 width="200">
	            <tr><td valign=top bgcolor="#ffffff" style="padding:5px;">
	            <img src="../ill/outzource_logo_200.gif" />
	            </td>
	            </tr>
	            <tr>
	            <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
	            <a href="../inc/log/data/<%=file%>" class=vmenu target="_blank" onClick="Javascript:window.close()">Din CSV. fil er klar >></a>
	            </td></tr>
	            </table>

                <%end if %>
	            
	          
	            
	            <%
                
                
                Response.end
	            Response.redirect "../inc/log/data/"& file &""	

         end if
                
                
                
                
                
                
                
                
           if media <> "print" AND media <> "eksport" AND media <> "chart" then
            
           
            ptop = 0 '72
            pleft = 1040
            pwdt = 180
            
           
            call eksportogprint(ptop, pleft, pwdt)

            lnk = "periodeSel="&periodeSel&"&FM_medarb="&thisMiduse&"&FM_medarb_hidden="&thisMiduse&"&jobselect="&jobnavn&"&mdselect="&monththis&"&aarselect="&yearThis&"&FM_vis_job_u_fc="&vis_job_u_fc&"&FM_vis_job_fc_neg="&vis_job_fc_neg&"&FM_hideEmptyEmplyLines="&hideEmptyEmplyLines&"&FM_vis_simpel="&vis_simpel&"&FM_vis_kolonne_simpel="&vis_kolonne_simpel&"&FM_vis_kolonne_simpel_akt="&vis_kolonne_simpel_akt&"&FM_sog="&sogTxt
            %>


               
            <form action="ressource_belaeg_jbpla.asp?media=eksport&<%=lnk%>" method="post" target="_blank">                     
            <tr>
                <td align=center>
                  <input type="image" src="../ill/export1.png" /></td>
                
                <td>
                    <input type="submit" value=".csv fil eksport (Pivot)" style="font-size:10px;" /></td>
                </tr>
            </form>
        
             <form action="ressource_belaeg_jbpla.asp?media=eksport&<%=lnk%>&vis=1" method="post" target="_blank">   
                                
                 <tr>
                <td align=center>
                <input type="image" src="../ill/export1.png" />
                </td>
                
                <td>
                <input type="submit" value=".csv fil eksport" style="font-size:10px;" />
                </td>
                </tr>

                 <tr><td colspan="2"><input type="checkbox" name="FM_expvisreal" value="1" <%=expvisrealCHK %> /> Vis realiserede timer<br />&nbsp;</td></tr>

            </form>
                

    <form action="ressource_belaeg_jbpla.asp?media=print&<%=lnk%>" method="post" target="_blank">   
                <tr>
                <td align=center>  <input type="image" src="../ill/printer3.png" />
                                 
                             </td><td> <input type="submit" value="Print version" style="font-size:10px;" /></td>
               </tr>
        </form>

        
         <%if cint(level) = 1 then %>

          <form action="ressource_belaeg_jbpla.asp?func=overfor&<%=lnk%>" method="post">   
                <tr>
                <td colspan="2"><br /><br /><b>Funktioner:</b><br />
                <input type="checkbox" value="1" name="FM_overfor" />Overfør timer ikke brugte ressource timer.
                    <span style="color:#999999;">Gælder for alle medarbejdere uanset viste. </span><br /><br />


                    <%if periodeSel = 3 then %>
                    Overfør saldo fra uge: <select name="FM_overforfra" style="font-size:10px;">
                                    <%for u = 1 to 53 
                                        
                                        if  datepart("ww", "1/"& monththis & "/" & yearthis, 2,2) = u - 1 then 'datepart("ww", now, 2,2)
                                        uSEL = "SELECTED"
                                        else
                                        uSel = ""
                                        end if%>
                                    <option value="<%=u %>" <%=uSel %>>Uge: <%=u & " - "& datepart("yyyy", now, 2,2)%></option>
                                    <%next %>

                                     </select> <br />til følgende uge.

                    <%else %>

                           Overfør saldo fra: <select name="FM_overforfra" style="font-size:10px;">
                                    <%for u = 1 to 12 
                                        
                                        if cint(monththis) = u then 'datepart("m", now, 2,2)
                                        uSEL = "SELECTED"
                                        else
                                        uSel = ""
                                        end if%>
                                    <option value="<%=u %>" <%=uSel %>><%=left(monthname(u), 3) &" "& yearthis %></option>
                                    <%next %>

                                     </select> <br />til følgende måned.

                    <%end if %>

                    <br /><br />
                    <input type="hidden" name="FM_overforfraAar" value="<%=yearthis %>" />
                            <span style="float:right;"><input type="submit" value="Overfør timer >>" style="font-size:10px;" /></span></td>
               </tr>
        </form>

        <%end if %>


<!--
                  <tr>
                            <td align=center><a href="ressource_belaeg_jbpla.asp?media=chart&<%=lnk%>&vis=1" target="_blank"  class='rmenu'>
                           &nbsp;<img src="../ill/chart_column.png" border="0" alt="" /></a>
                            </td><td><a href="ressource_belaeg_jbpla.asp?media=chart&<%=lnk%>&vis=1" target="_blank" class="rmenu">Graf</a></td>
                           

                           </tr>
    -->

<!--
                    <tr><td colspan="2"><br />
                        Nulstil saldo på alle forecast pr. <input type="text" value="<%=formatdatetime(now,2) %>" style="width:60px; font-size:9px;" /> <input type="button" value=">>" style="font-size:9px;" /><br />
                        <span style="color:#999999; font-size:9px;">Indlæser og nulstiller forecast +/- for alle medarbejdere på alle linier pr. valgte dato, så der kan indtastes et nyt forecast, i ny periode, fra 0</span> 
                        </td></tr>

    -->
               </table>
            </div>
            
            <%
            else
            Response.Write("<script language=""JavaScript"">window.print();</script>")
            end if%>
           
           
           </div>
			
			
			
		
	
	
	<br><br>&nbsp;
	
</div>




<br><br>&nbsp;<br><br>&nbsp;<br><br>&nbsp;<br><br>&nbsp;<br><br>&nbsp;<br><br>&nbsp;




	
<%end if%>

          

<!--#include file="../inc/regular/footer_inc.asp"-->


	
