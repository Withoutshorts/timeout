<%Response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	'session.LCID = 1033
	
        call bdgmtypon_fn()

        if cint(bdgmtypon_val) = 1 then 'budget på mtyper slået til
        vlgtmtypgrp = 0
        call mtyperIGrp_fn(vlgtmtypgrp,1)
        call fn_medarbtyper()
        end if

	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
	
	if len(request("optiprint")) <> 0 then
	optiPrint = request("optiprint")	
	else
	optiPrint = 0
	end if


    if optiprint = 6 then 'EPI BU fil

          ktfomrAarMaks = 100
          dim fomrNameTxt, fomrLabelTxt, fak_AfdTxt, kundetypeNavnTxt, kundetypeLabelTxt, strTxtExportJobUB
          redim fomrNameTxt(ktfomrAarMaks), fomrLabelTxt(ktfomrAarMaks), fak_AfdTxt(ktfomrAarMaks), kundetypeNavnTxt(ktfomrAarMaks), kundetypeLabelTxt(ktfomrAarMaks), strTxtExportJobUB(5000)


                 '*** Kundetyper ****'
                'strSQLfomr = "SELECT navn FROM kundetyper WHERE id > 0 ORDER BY id"
                'kt = 0
                'oRec4.open strSQLfomr, oConn, 3
				'while not oRec4.EOF
							 
                'kundetypeNavnTxt(kt) = oRec4("navn")
				'kundetypeLabelTxt(kt) = "" 'oRec4("business_area_label")
                
				
                'kt = kt + 1			    
		        'oRec4.movenext
        		'wend
				'oRec4.close


                '*** FOMR = UnderAFD, UNIT = AFD ****'
                strSQLfomr = "SELECT navn, business_area_label, business_unit FROM fomr WHERE id > 0 ORDER BY id"
                fo = 0
                oRec4.open strSQLfomr, oConn, 3
				while not oRec4.EOF
							 
                fomrNameTxt(fo) = oRec4("navn")
				fomrLabelTxt(fo) = oRec4("business_area_label")
                fak_AfdTxt(fo) = oRec4("business_unit")
				

                fo = fo + 1			    
		        oRec4.movenext
        		wend
				oRec4.close

    end if

    if len(trim(request("historisk_wip"))) <> 0 then
        historisk_wip = request("historisk_wip")
    else
        historisk_wip = 0
    end if 


    if len(trim(request("timertilfak"))) <> 0 then
    timertilfak = request("timertilfak")
    else
    timertilfak = 0
    end if


    

    if cint(historisk_wip) = 1 then
        historisk_wipTxt = "historisk"
    else
        historisk_wipTxt = "aktuel"
    end if

    if len(trim(request("sorttp"))) <> 0 AND request("sorttp") <> 0 then
    sorttp = 1
    else
    sorttp = 0
    end if

    if len(trim(request("visSimpel"))) <> 0 then
        visSimpel = request("visSimpel")
    else
        visSimpel = -1
    end if

    
    if len(trim(request("eksDataStd"))) <> 0 then 'fra joblisten

  
    eksDataStd = 1
    
    if len(trim(request("eksDataNrl"))) <> 0 then
    eksDataNrl = 1
    else
    eksDataNrl = 0
    end if

    if len(trim(request("eksDataJsv"))) <> 0 then
    eksDataJsv = 1
    else
    eksDataJsv = 0
    end if

    if len(trim(request("eksDataAkt"))) <> 0 then
    eksDataAkt = 1
    else
    eksDataAkt = 0
    end if

    if len(trim(request("eksDataMile"))) <> 0 then
    eksDataMile = 1
    else
    eksDataMile = 0
    end if
    

    

    else ' fra f.eks webblik joblisten

    eksDataStd = 1

    if len(trim(request("eksDataNrl"))) <> 0 then
    eksDataNrl = 1
    else
    eksDataNrl = 0
    end if

    if len(trim(request("eksDataJsv"))) <> 0 then
    eksDataJsv = 1
    else
    eksDataJsv = 0
    end if


    eksDataAkt = 0
    eksDataMile = 0

    end if

	
	level = session("rettigheder")
    thisMid = session("mid")



    '************** FILTYPE ******************
	select case optiPrint 
    case 0, 3, 4, 5, 6
    'ext = "txt"
	ext = "csv"
    'case 2
    'ext = "csv"
    'Response.Charset = "utf-8"
	case else
	ext = "txt"
	end select

    if len(trim(request("realfakpertot"))) <> 0 then
    realfakpertot = request("realfakpertot")
    else
    realfakpertot = 0
    end if

    sqlDatostart = request("sqlDatostart")
    sqlDatoslut = request("sqlDatoslut")
    
    sqlDatostartMtypTkons = left(sqlDatostart, 2) &"/"& datepart("m",sqlDatostart, 2,2) & "/1"

        'Response.write sqlDatostartMtypTkons 
        'Response.end

    'Response.write "realfakpertot" & realfakpertot
    'Response.end


    call TimeOutVersion()

    call akttyper2009(2)

    '************** STI OG MAPPE ******************
	
	Set objFSO = server.createobject("Scripting.FileSystemObject")
	
	if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\job_eksport.asp" then
							
		Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\jobexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."&ext&"", True, False)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\jobexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."&ext&"", 8)
	
	else
		
		Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\jobexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."&ext&"", True, false)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\jobexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."&ext&"", 8, -1)
		
	end if
	
	file = "jobexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."&ext&""
	
	

    call salgsans_fn()    


    jids = request("jids")
	jobIder = split(jids, ",") 



    '************** HEADER ******************
	
	'if optiPrint = 0 OR optiPrint = 2 then
	select case optiPrint
    case 0,2 
    strTxtExport = strTxtExport & "Kontakt;Kontakt id;Jobnavn;Jobnr.;Status;Startdato;Slutdato;Prioitet;"
        
         cellerForAkt_og_Faser = ";;;;;;;;;"

        if eksDataNrl <> 1 then 
        strTxtExport = strTxtExport & "Realiseret Timer;"
          cellerForAkt_og_Faser = cellerForAkt_og_Faser & ";"
        end if
     
           

            if eksDataNrl = 1 then
            cellerForAkt_og_Faser = cellerForAkt_og_Faser & ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"

                
                            if cint(bdgmtypon_val) = 1 then 
            
                                for t = 1 to UBOUND(mtypgrpids)

                                if len(trim(mtypgrpnavn(t))) <> 0 then
                                cellerForAkt_og_Faser = cellerForAkt_og_Faser & ";"
                                end if
            
                            next
            
                            end if
            
           end if


             if eksDataJsv = 1 then
             cellerForAkt_og_Faser = cellerForAkt_og_Faser & ";;;;;;;;;;;;;;;;"
                
                 if cint(showSalgsAnv) = 1 then
                 cellerForAkt_og_Faser = cellerForAkt_og_Faser & ";;;;;;;;;;;;;;;"
                 end if
            
             end if




                    if eksDataNrl = 1 then
	                strTxtExport = strTxtExport &"Fastpris:1/Lbn Timer:0;Forkalk. Timer;Budget. Bruttooms.;"
  
    
                          if cint(bdgmtypon_val) = 1 then 'budget på mtyper slået til
         
                         for t = 1 to UBOUND(mtypgrpids)

                            if len(trim(mtypgrpnavn(t))) <> 0 then
            
                            strTxtExport = strTxtExport &"Budget "&  mtypgrpnavn(t) & ";"

                            end if
                        next


                         end if
        
                    strTxtExport = strTxtExport &"Budget Salgsomk. (indkøb);Budget Intern kost.;Budget DB beløb;Budget DB %;"_
                    &"Realiseret Timer;~Real. Timepris;Real. Internkost. pr. time;Real. Bruttoomsætning;Real. Nettoomsætning;"
    
                          if cint(bdgmtypon_val) = 1 then 'budget på mtyper slået til
         
                                for t = 0 to UBOUND(mtypeids)
                                       strTxtExport = strTxtExport &"Real. timer: "&  mtypenavne(t) & ";Real. kostpris: "&  mtypenavne(t) & ";"
                                next


                         end if    
        
                    strTxtExport = strTxtExport &"Real. DB beløb;Real. DB %;"_
                    &"WIP DB beløb;WIP DB %;Faktureret;Faktisk Salgsomk. (indkøb);Faktisk Internkost.;Faktisk DB beløb;Faktisk DB %;"_
                    &"Timepris (faktisk);Stade (WIP) %;WIP pr. dato ("& historisk_wipTxt &");WIP angivet dato;WIP Omsætning; Bal. Faktureret-Oms. WIP;"

       

                    end if

    
         


                    if eksDataJsv = 1 then
	                strTxtExport = strTxtExport &"Kontaktpers;Jobansvarlig 1;Jobans1. %;Init 1;Jobejer 2;Jobans2. %;Init 2;Medansv. 3;Jobans3. %;Init 3;Medansv. 4;Jobans4. %;Init 4;Medansv. 5;Jobans5. %;Init 5;"
    
                    if cint(showSalgsAnv) = 1 then
                    strTxtExport = strTxtExport &"Salgsansv. 1;Salgsansv. 1 %;Salgsansv. 1 Init;Salgsansv. 2;Salgsansv. 2 %;Salgsansv. 2 Init;Salgsansv. 3;Salgsansv. 3 %;Salgsansv. 3 Init;Salgsansv. 4;Salgsansv. 4 %;Salgsansv. 4 Init;Salgsansv. 5;Salgsansv. 5 %;Salgsansv. 5 Init;"
                    end if

                    end if
    
                    if eksDataNrl = 1 then
                    strTxtExport = strTxtExport &"Aftale;Rekvnr.;Pre-konditioner opfyldt;Forretningsomr.;Projektgrupper (job);"
                    end if

                    if eksDataJsv = 1 then

                    if level = 1 then
                        if lto <> "epi" OR (lto = "epi" AND (thisMid = 6 OR thisMid = 11 OR thisMid = 59 OR thisMid = 1686 OR thisMid = 1720)) OR (lto = "epi_no" AND thisMid = 2) OR (lto = "epi_ab" AND thisMid = 2) OR (lto = "epi_sta" AND thisMid = 2) then 
                        strTxtExport = strTxtExport & "Virksomhedsandel;"
                        cellerForAkt_og_Faser = cellerForAkt_og_Faser & ";"
                        end if
                    end if
    
                    end if

                    'if optiPrint = 2 then
                    if eksDataNrl = 1 then
                    strTxtExport = strTxtExport & "Beskrivelse;"  ';Kommentar" & vbcrlf 
                    end if
    
                     strTxtExport = strTxtExport & "Kommentar;" 


                     if cint(eksDataAkt) = 1 then
                     strTxtExport = strTxtExport &"Fase;Aktiviteter;Aktivitet startdato;Aktivitet slutdato;"
                     end if

                      if cint(eksDataMile) = 1 then
                     strTxtExport = strTxtExport &"Milepæl;Dato;Type;Beløb;"
                     end if


   
    strTxtExport = strTxtExport & vbcrlf 
	
    case 6
	    
            strTxtExport = strTxtExport & "NAME;NAME_SHT;NAME afd.;NAV_NAME afd. Kode;Jobnavn;Jobnr.;Status;NAME_SHT;NAV_NAME;"

            strTxtExport = strTxtExport & vbcrlf 

    end select 'optiPrint
    'end if
	



    '************** BODY / LINJER ******************

    select case optiPrint
    case 0,1,2,6

    'if optiPrint = 0 OR optiPrint = 1 OR optiPrint = 2 then
	

    antaljobs = 0
	for i = 0 to Ubound(jobIder)
	
	
	strSQL = "SELECT j.dato, j.id, jobnavn, jobnr, kkundenavn, kid, jobknr, jobTpris, jobstatus, jobstartdato, risiko, "_
	&" jobslutdato, budgettimer, fakturerbart, Kkundenr, ikkebudgettimer, "_
    &" j.jobans1, jobans2, jobans3, jobans4, jobans5, "_
    &" jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, "

    if cint(showSalgsAnv) = 1 then
    strSQL = strSQL & " salgsans1, salgsans2, salgsans3, salgsans4, salgsans5, "_
    &" salgsans1_proc, salgsans2_proc, salgsans3_proc, salgsans4_proc, salgsans5_proc, "
    end if

    strSQL = strSQL &" virksomheds_proc, j.beskrivelse, fastpris, "_
    &" projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, "

	strSQL = strSQL & " m1.mnavn AS m1navn, m1.init AS m1init, m1.mnr AS m1nr, "_
	&" m2.mnavn AS m2navn, m2.init AS m2init, m2.mnr AS m2nr, "_
    &" m3.mnavn AS m3navn, m3.init AS m3init, m3.mnr AS m3nr, "_
    &" m4.mnavn AS m4navn, m4.init AS m4init, m4.mnr AS m4nr, "_
    &" m5.mnavn AS m5navn, m5.init AS m5init, m5.mnr AS m5nr, "

     if cint(showSalgsAnv) = 1 then
   strSQL = strSQL & " ms1.mnavn AS ms1navn, ms1.init AS ms1init, ms1.mnr AS ms1nr, "_
	&" ms2.mnavn AS ms2navn, ms2.init AS ms2init, ms2.mnr AS ms2nr, "_
    &" ms3.mnavn AS ms3navn, ms3.init AS ms3init, ms3.mnr AS ms3nr, "_
    &" ms4.mnavn AS ms4navn, ms4.init AS ms4init, ms4.mnr AS ms4nr, "_
    &" ms5.mnavn AS ms5navn, ms5.init AS ms5init, ms5.mnr AS ms5nr, "
    end if
    
    strSQL = strSQL &" rekvnr, serviceaft, kommentar, jo_dbproc, udgifter, jo_bruttooms, jo_bruttofortj, jo_udgifter_intern, jo_udgifter_ulev, "_
    &" restestimat, stade_tim_proc, kp.navn AS kontaktpers, preconditions_met "_
    &" FROM job AS j"

    strSQL = strSQL & " LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "

    strSQL = strSQL & " LEFT JOIN medarbejdere m1 ON (m1.mid = j.jobans1) "_
	&" LEFT JOIN medarbejdere m2 ON (m2.mid = j.jobans2) "_
    &" LEFT JOIN medarbejdere m3 ON (m3.mid = j.jobans3) "_
    &" LEFT JOIN medarbejdere m4 ON (m4.mid = j.jobans4) "_
    &" LEFT JOIN medarbejdere m5 ON (m5.mid = j.jobans5) "
    
    if cint(showSalgsAnv) = 1 then
    strSQL = strSQL & "LEFT JOIN medarbejdere ms1 ON (ms1.mid = j.salgsans1) "_
	&" LEFT JOIN medarbejdere ms2 ON (ms2.mid = j.salgsans2) "_
    &" LEFT JOIN medarbejdere ms3 ON (ms3.mid = j.salgsans3) "_
    &" LEFT JOIN medarbejdere ms4 ON (ms4.mid = j.salgsans4) "_
    &" LEFT JOIN medarbejdere ms5 ON (ms5.mid = j.salgsans5) "
    end if

    strSQL = strSQL &" LEFT JOIN kontaktpers AS kp ON (kp.id = kundekpers)"_
	&" WHERE j.id = "& jobIder(i) &" AND k.kid = j.jobknr ORDER BY j.jobnavn"
    
    
	'&" INTO OUTFILE 'c:\\"& file &"' FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '\' LINES TERMINATED BY '\r\n'" _

	'Response.write strSQL & "<hr>"
	'Response.flush
	
    'oConn.execute(strSQL)
    'Response.end

	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	
    jobid = oRec("id")


    prioitet = oRec("risiko")
  
	
	strJobnavn = oRec("jobnavn")
	strJobnr = oRec("jobnr")
	intJobstatus = oRec("jobstatus")
	strKunde = oRec("kkundenavn")
	strKnr = oRec("kkundenr")
	dtStdato = oRec("jobstartdato")
	dtSldato = oRec("jobslutdato")
	intFakbart = oRec("fakturerbart")
	intFakbaretimer = oRec("budgettimer")
	intIkkeFakbaretimer = oRec("ikkebudgettimer")
	timerialt = (intFakbaretimer + intIkkeFakbaretimer)
    budgettimerIalt = timerialt
	


    '** Jobsnavrlige *****'
    strJobans1 = oRec("m1navn")
    jobans_proc_1 = oRec("jobans_proc_1")
	strJobans1nr = oRec("m1nr")
	strJobans1init = oRec("m1init")

    if len(trim(strJobans1)) <> 0 then
    strJobans1nr = "("&strJobans1nr&")"
    else
    strJobans1nr = ""
    end if
	
    strJobans2 = oRec("m2navn")
    jobans_proc_2 = oRec("jobans_proc_2")
	strJobans2nr = oRec("m2nr")
	strJobans2init = oRec("m2init")

     if len(trim(strJobans2)) <> 0 then
    strJobans2nr = "("&strJobans2nr&")"
    else
    strJobans2nr = ""
    end if

    strJobans3 = oRec("m3navn")
    jobans_proc_3 = oRec("jobans_proc_3")
	strJobans3nr = oRec("m3nr")
	strJobans3init = oRec("m3init")

     if len(trim(strJobans3)) <> 0 then
    strJobans3nr = "("&strJobans3nr&")"
    else
    strJobans3nr = ""
    end if

    strJobans4 = oRec("m4navn")
    jobans_proc_4 = oRec("jobans_proc_4")
	strJobans4nr = oRec("m4nr")
	strJobans4init = oRec("m4init")

    if len(trim(strJobans4)) <> 0 then
    strJobans4nr = "("&strJobans4nr&")"
    else
    strJobans4nr = ""
    end if


    strJobans5 = oRec("m5navn")
    jobans_proc_5 = oRec("jobans_proc_5")
	strJobans5nr = oRec("m5nr")
	strJobans5init = oRec("m5init")

    if len(trim(strJobans5)) <> 0 then
    strJobans5nr = "("&strJobans5nr&")"
    else
    strJobans5nr = ""
    end if

    

     if cint(showSalgsAnv) = 1 then
    '** Salgsansvarlige ***'
    strsalgsans1 = oRec("ms1navn")
    salgsans_proc_1 = oRec("salgsans1_proc")
	strsalgsans1nr = oRec("ms1nr")
	strsalgsans1init = oRec("ms1init")

    if len(trim(strsalgsans1)) <> 0 then
    strsalgsans1nr = "("&strsalgsans1nr&")"
    else
    strsalgsans1nr = ""
    end if
	
    strsalgsans2 = oRec("ms2navn")
    salgsans_proc_2 = oRec("salgsans2_proc")
	strsalgsans2nr = oRec("ms2nr")
	strsalgsans2init = oRec("ms2init")

     if len(trim(strsalgsans2)) <> 0 then
    strsalgsans2nr = "("&strsalgsans2nr&")"
    else
    strsalgsans2nr = ""
    end if

    strsalgsans3 = oRec("ms3navn")
    salgsans_proc_3 = oRec("salgsans3_proc")
	strsalgsans3nr = oRec("ms3nr")
	strsalgsans3init = oRec("ms3init")

     if len(trim(strsalgsans3)) <> 0 then
    strsalgsans3nr = "("&strsalgsans3nr&")"
    else
    strsalgsans3nr = ""
    end if

    strsalgsans4 = oRec("ms4navn")
    salgsans_proc_4 = oRec("salgsans4_proc")
	strsalgsans4nr = oRec("ms4nr")
	strsalgsans4init = oRec("ms4init")

    if len(trim(strsalgsans4)) <> 0 then
    strsalgsans4nr = "("&strsalgsans4nr&")"
    else
    strsalgsans4nr = ""
    end if


    strsalgsans5 = oRec("ms5navn")
    salgsans_proc_5 = oRec("salgsans5_proc")
	strsalgsans5nr = oRec("ms5nr")
	strsalgsans5init = oRec("ms5init")

    if len(trim(strsalgsans5)) <> 0 then
    strsalgsans5nr = "("&strsalgsans5nr&")"
    else
    strsalgsans5nr = ""
    end if


    end if



    if level = 1 then
        if lto <> "epi" OR (lto = "epi" AND (thisMid = 6 OR thisMid = 11 OR thisMid = 59 OR thisMid = 1686 OR thisMid = 1 OR thisMid = 1720)) OR (lto = "epi_no" AND thisMid = 2) OR (lto = "epi_ab" AND thisMid = 2) OR (lto = "epi_sta" AND thisMid = 2) then
        virksomheds_proc = oRec("virksomheds_proc")
        end if
    end if

    if isNull(oRec("beskrivelse")) <> true then
	strJobBesk = oRec("beskrivelse")
    else
    strJobBesk = ""
    end if

    if isNull(oRec("kontaktpers")) <> true then
    kontaktpers = oRec("kontaktpers")
    else
    kontaktpers = ""
    end if

	dblBudget = oRec("jo_bruttooms") 'oRec("jobTpris")
    nettoomstimer = oRec("jobTpris")
	intFaspris = oRec("fastpris")
	rekvnr = oRec("rekvnr")
	serviceaft = oRec("serviceaft")
	komm = oRec("kommentar")
	jo_dbproc = oRec("jo_dbproc")
	joDbbelob = oRec("jo_bruttofortj")

    'udgifter = oRec("udgifter") 'salgsomk + internkost
	salgsomk = oRec("jo_udgifter_ulev")
    jo_udgifter_ulev = oRec("jo_udgifter_ulev")
    internKost = oRec("jo_udgifter_intern")

    preconditions_met = oRec("preconditions_met")

		'*** Export ****
		'objF.WriteLine("Id "&  chr(009) &" Kundenavn "&  chr(009) &"Kundenr "& chr(009) &"Jobnavn "& chr(009)&"Aktivitet "& chr(009) &"Kunde "& chr(009) &"Timer / Enheder"& chr(009) &"Heraf fakturerbare timer"& chr(009) &"Reg. Timepris (uanset fastpris/lbn.tim.)"& chr(009) &"Gns.fak. Timepris" & chr(009) & "Medarbejder "& chr(009) &"Kommentar ")
		
		select case intJobstatus
		case 0
		strStatus = "Lukket"
		case 1
		strStatus = "Aktiv"
		case 2
		strStatus = "Passiv"
		case 3
		strStatus = "Tilbud"
		end select 
		
		
        '******* Eksport version *****'
		if optiPrint = 0 OR optiPrint = 2 then





            if dblBudget <> 0 then
            jobbudget = dblBudget
            else
            jobbudget = 0
            end if


           '**** Faktureret
            'faktureretLastFakDato = day(sqlDatostart) &"/"& month(sqlDatostart) &"/"& year(sqlDatostart) 
            call stat_faktureret_job(oRec("id"), sqlDatostart, sqlDatoslut) '** cls_fak
            


             '**** Realiserede timer **************
            '** Tot
            '** I periode
            '** Siden sidste faktura

           
             if cint(timertilfak) = 1 then 'kun timer til fakturering i valgte periode (siden sidste faktura)
	
                if cDate(faktureretLastFakDato) > cdate(day(sqlDatostart) &"/"& month(sqlDatostart) &"/"& year(sqlDatostart)) then
                sqlDatostRealTimer = year(faktureretLastFakDato) &"/"& month(faktureretLastFakDato) &"/"& day(faktureretLastFakDato)
                startDatovlgt = faktureretLastFakDato
                lastFakbrugt = 1
                else
                sqlDatostRealTimer = sqlDatostart
                lastFakbrugt = 0
                end if

            else
            sqlDatostRealTimer = sqlDatostart
            end if
            
           


            'call timeRealOms '** cls_timer
            call timeRealOms(oRec("jobnr"), sqlDatostRealTimer, sqlDatoslut, nettoomstimer, oRec("fastpris"), budgettimerIalt, aty_sql_realhours) '** cls_timer

            'response.write "timertilfak: "& timertilfak & " faktureretLastFakDato: "& faktureretLastFakDato &" sqlDatostart: "& sqlDatostart & " sqlDatostRealTimer: "& sqlDatostRealTimer
            'response.end
            


       
		
		
		if timerforbrugt <> 0 then
		gnstimepris = faktureretTimerEnhStk/timerforbrugt
		else
		gnstimepris = 0
		end if


       
       
        udgifterFaktisk = salgsOmkFaktisk + kostpris
        db2bel = (faktureret-(udgifterFaktisk))

        call fn_dbproc(faktureret,db2bel)
        db2 = dbProc


        
        timerTildelt = timerialt



                    '*** WIP / Satde Historik eller aktuel ******'' 
                    if cint(historisk_wip) = 1 then
            
                        call wip_historik_fn(oRec("id"),sqlDatoslut)
                        wipdato = formatdatetime(sqlDatoslut, 2)
                        wipHistdato = wipHistdato


                    else

                    if len(oRec("restestimat")) <> 0 then
		            restestimat = oRec("restestimat")
		            else
		            restestimat = 0
		            end if

                    stade_tim_proc = oRec("stade_tim_proc")

                    wipdato = formatdatetime(now,2)
                    wipHistdato = oRec("dato")

                    end if
		
		
		select case stade_tim_proc '0 = rest i timer, 1 = i proc
		case 0
		totalforbrugt = (timerforbrugt + restestimat)
		case 1
		    
		    if timerforbrugt <> 0 AND restestimat <> 0 then
		    totalforbrugt = timerforbrugt * (100/restestimat) 
		    else
		    totalforbrugt = 0
		    end if '- (restestimat/timerTildelt) * 100)
		
		end select
		
        jobbudget = dblBudget

        
        if timerTildelt <> 0 then
        udg_internKostprTim = oRec("jo_udgifter_intern") / timerTildelt
        else
        udg_internKostprTim = oRec("jo_udgifter_intern") / 1
        end if 


         if timerforbrugt <> 0 then 
         gnskostprtime = kostpris/timerforbrugt
         else
         gnskostprtime = 0
         end if
        
        


        OmsWIP = (afsl_proc/100) * jobbudget 
        salgsOmkWIP = salgsOmkFaktisk 
        nettoWIP = ((afsl_proc/100) * (jobbudget)) - salgsOmkWIP 
        



         
        '*** bal WIP ****'

        if cint(stade_tim_proc) = 0 then 'rest estimat angivet i timer
		
		      if totalforbrugt <> 0 then
		
		        if restestimat = 0 then
		        stade = 100
        		else
        		stade = (timerforbrugt/totalforbrugt) * 100
		        end if
		    
		    else
		    
		    stade = 0
		    
		    end if
		    
		     afsl_proc = stade
		else
	        
	       afsl_proc = restestimat
		end if
        

        OmsWIP =  (afsl_proc/100) * jobbudget 
        balWIP = (faktureret - OmsWIP)



        wipGnskostprtime = gnskostprtime
        interKostWip = kostpris 'jo_udgifter_intern '(afsl_proc/100) *

        WipOmkIalt = interKostWip + salgsOmkWIP
        forvDbbel = (OmsWIP - (WipOmkIalt))
    
        call fn_dbproc(OmsWIP,forvDbbel)
        forvDb = dbProc

         


      

        
        'bruttoOmsReal = (OmsReal + salgsOmkFaktisk)
        bruttoOmsReal = (OmsReal + matSalgsprisReal) 
        RealOmk = salgsOmkFaktisk + kostpris
        realDbbelob = (bruttoOmsReal-RealOmk)
        OmsRealBrutto = bruttoOmsReal '(OmsReal/1+salgsOmkFaktisk/1)

        
         call fn_dbproc(bruttoOmsReal,realDbbelob)
        realDB = dbProc

		
		strTxtExport = strTxtExport & Chr(34) & strKunde & Chr(34) &";" & Chr(34) & strKnr & Chr(34) &";"& Chr(34) & strJobnavn & Chr(34) &";"& Chr(34) &strJobnr& Chr(34) &";"& Chr(34) &strStatus& Chr(34) &";" & Chr(34) & dtStdato & Chr(34) & ";"& Chr(34) & dtSldato& Chr(34) &";" & Chr(34) & prioitet & Chr(34) &";"
        
        if eksDataNrl <> 1 then 
        strTxtExport = strTxtExport & Chr(34) & formatnumber(timerforbrugt,2) & Chr(34) &";"
        end if
          
        
        if eksDataNrl = 1 then
        strTxtExport = strTxtExport & Chr(34) & intFaspris & Chr(34) &";"& Chr(34) & formatnumber(timerialt, 2) & Chr(34) &";"& Chr(34) & formatnumber(dblBudget,2) & Chr(34) &";"
        
                
            if cint(bdgmtypon_val) = 1 then 
            
                    for t = 1 to UBOUND(mtypgrpids)

                    if len(trim(mtypgrpnavn(t))) <> 0 then
                    thisMtypGrpBel = 0
                    strSQLmtypbgt = "SELECT SUM(belob) AS belob FROM medarbejdertyper_timebudget WHERE jobid = "& oRec("id") &" AND ("& mtypgrpids(t) &") GROUP BY jobid"
                    'Response.write strSQLmtypbgt
                    oRec6.open strSQLmtypbgt, oConn, 3
                    if not oRec6.EOF then
                    thisMtypGrpBel = oRec6("belob")
                    end if
                    oRec6.close
            
                    if len(trim(thisMtypGrpBel)) <> 0 then
                    thisMtypGrpBel = thisMtypGrpBel
                    else
                    thisMtypGrpBel = 0 
                    end if 

                    strTxtExport = strTxtExport & Chr(34) & formatnumber(thisMtypGrpBel,2) & Chr(34) &";"
            

                    end if
                    next

            end if
        
        strTxtExport = strTxtExport & Chr(34) & formatnumber(salgsomk,2) & Chr(34) &";"& Chr(34) & formatnumber(internKost,2) & Chr(34) &";"_
        & Chr(34) & formatnumber(joDbbelob,2) & Chr(34) &";"_
        & Chr(34) & formatnumber(jo_dbproc,0) & Chr(34) &";"_
        & Chr(34) & formatnumber(timerforbrugt,2) & Chr(34) &";"& Chr(34) & formatnumber(tp,2) & Chr(34) &";"& Chr(34) & formatnumber(gnskostprtime, 2) & Chr(34) &";"& Chr(34) & formatnumber(OmsRealBrutto,2) & Chr(34) &";"_
        & Chr(34) & formatnumber(OmsReal,2) & Chr(34) &";"
        
                 
        
         if cint(bdgmtypon_val) = 1 then 
            
                    for t = 0 to UBOUND(mtypeids)

                  
                    thisMtypTimerBelob = 0
                    thisMtypTimer = 0
                    thisMtyptimerkost = 0
        
                    
                    'if lto = "epi" then 'konsolideret

        
                        if cint(realfakpertot) <> 0 then 'timer o peridoe elle total == konsolideret
                        strSQLmtypbgt = "SELECT SUM(timer) AS mtyptimer, SUM(belob) AS mtyptimerbelob, SUM(kost) AS timerkost FROM timer_konsolideret_tot WHERE jobid = "& oRec("id") & " AND (mtype = "& mtypeids(t) & " "& mtypesostergp(t) &") AND dato BETWEEN '"& sqlDatostartMtypTkons &"' AND '"& sqlDatoslut &"' GROUP BY jobid"
                        else
                        strSQLmtypbgt = "SELECT SUM(timer) AS mtyptimer, SUM(belob) AS mtyptimerbelob, SUM(kost) AS timerkost FROM timer_konsolideret_tot WHERE jobid = "& oRec("id") & " AND (mtype = "& mtypeids(t) & " "& mtypesostergp(t) &") GROUP BY jobid"
                        end if

                    'else 'EVT ANDRE EPI VERSIONER 20151120 

                        'call mtyperIGrp_fn(mtypeids(t),1) 

                        'if cint(realfakpertot) <> 0 then
                        'strSQLmtypbgt = "SELECT sum(t.timer) AS mtyptimer, sum(kostpris*t.timer*(t.kurs/100)) AS timerkost, sum(t.timer*(t.timepris*(t.kurs/100))) AS mtyptimerbelob FROM timer t WHERE t.tjobnr = '"& jobnr &"' AND ("& aty_sql_realhours &") AND ("& medarbimType(t) &") AND (dato BETWEEN '"& sqlDatostartMtypTkons &"' AND '"& sqlDatoslut &"') GROUP BY tjobnr"
                        'else
                        'strSQLmtypbgt = "SELECT sum(t.timer) AS mtyptimer, sum(kostpris*t.timer*(t.kurs/100)) AS timerkost, sum(t.timer*(t.timepris*(t.kurs/100))) AS mtyptimerbelob FROM timer t WHERE t.tjobnr = '"& jobnr &"' AND ("& aty_sql_realhours &") AND ("& medarbimType(t) &") GROUP BY tjobnr"
                        'end if


                    'end if
           
                   

                    'Response.write strSQLmtypbgt
                    oRec6.open strSQLmtypbgt, oConn, 3
                    if not oRec6.EOF then
                        thisMtypTimerBelob = oRec6("mtyptimerbelob")
                        thisMtypTimer = oRec6("mtyptimer")
                        thisMtyptimerkost = oRec6("timerkost")
                    end if
                    oRec6.close


                    strTxtExport = strTxtExport & Chr(34) & formatnumber(thisMtypTimer,2) & Chr(34) &";" & Chr(34) & formatnumber(thisMtyptimerkost,2) & Chr(34) &";"
                   

                    next
        end if

        strTxtExport = strTxtExport & Chr(34) & formatnumber(realDbbelob,2) & Chr(34) &";"_
        & Chr(34) & formatnumber(realDB,0) & Chr(34) &";"_
        & Chr(34) & formatnumber(forvDbbel,2) & Chr(34) &";"_
        & Chr(34) & formatnumber(forvDb,0) & Chr(34) &";"_
        & Chr(34) & formatnumber(faktureret,2) & Chr(34) &";"_
        & Chr(34) & formatnumber(salgsOmkFaktisk,2) & Chr(34) &";"& Chr(34) & formatnumber(kostpris,2) & Chr(34) &";"_
        & Chr(34) & formatnumber(db2bel,2) & Chr(34) &";"_
        & Chr(34) & formatnumber(db2,0) & Chr(34) &";"_
        & Chr(34) & formatnumber(gnstimepris,2) & Chr(34) &";"& Chr(34) & formatnumber(afsl_proc, 0) & Chr(34) &";" & Chr(34) & wipdato & Chr(34) &";" & Chr(34) & wipHistdato & Chr(34) &";"& Chr(34) & formatnumber(OmsWIP, 2) & Chr(34) &";"& Chr(34) & formatnumber(balWIP, 2) & Chr(34) &";"
		end if


        if eksDataJsv = 1 then
        strTxtExport = strTxtExport & Chr(34) & kontaktpers & Chr(34) &";"& Chr(34) & strJobans1 & strJobans1nr & Chr(34) &";"& Chr(34) & jobans_proc_1 & Chr(34) &";"& Chr(34) & strJobans1init & Chr(34) &";"_
		& Chr(34) & strJobans2 & strJobans2nr & Chr(34) &";"& Chr(34) & jobans_proc_2 & Chr(34) &";"& Chr(34) & strJobans2init & Chr(34) &";"_
        & Chr(34) & strJobans3 & strJobans3nr & Chr(34) &";"& Chr(34) & jobans_proc_3 & Chr(34) &";"& Chr(34) & strJobans3init & Chr(34) &";"_
        & Chr(34) & strJobans4 & strJobans4nr & Chr(34) &";"& Chr(34) & jobans_proc_4 & Chr(34) &";"& Chr(34) & strJobans4init & Chr(34) &";"_
        & Chr(34) & strJobans5 & strJobans5nr & Chr(34) &";"& Chr(34) & jobans_proc_5 & Chr(34) &";"& Chr(34) & strJobans5init & Chr(34) &";"
        
            if cint(showSalgsAnv) = 1 then
            strTxtExport = strTxtExport & Chr(34) & strsalgsans1 & strsalgsans1nr & Chr(34) &";"& Chr(34) & salgsans_proc_1 & Chr(34) &";"& Chr(34) & strsalgsans1init & Chr(34) &";"_
		    & Chr(34) & strsalgsans2 & strsalgsans2nr & Chr(34) &";"& Chr(34) & salgsans_proc_2 & Chr(34) &";"& Chr(34) & strsalgsans2init & Chr(34) &";"_
            & Chr(34) & strsalgsans3 & strsalgsans3nr & Chr(34) &";"& Chr(34) & salgsans_proc_3 & Chr(34) &";"& Chr(34) & strsalgsans3init & Chr(34) &";"_
            & Chr(34) & strsalgsans4 & strsalgsans4nr & Chr(34) &";"& Chr(34) & salgsans_proc_4 & Chr(34) &";"& Chr(34) & strsalgsans4init & Chr(34) &";"_
            & Chr(34) & strsalgsans5 & strsalgsans5nr & Chr(34) &";"& Chr(34) & salgsans_proc_5 & Chr(34) &";"& Chr(34) & strsalgsans5init & Chr(34) &";"
           end if
        
        end if

        if eksDataNrl = 1 then

		        strTxtExport = strTxtExport & Chr(34) & serviceaft& Chr(34) &";"& Chr(34) & rekvnr & Chr(34) &";"& Chr(34) & preconditions_met & Chr(34) &";"
       

                '**** Fomr ****'
                strTxtFomr = ""
                strSQLf = "SELECT for_fomr, fomr.navn FROM fomr_rel "_
                &" LEFT JOIN fomr ON (fomr.id = for_fomr) WHERE for_jobid = "& jobIder(i) & " GROUP BY for_fomr"
                fo = 0
                oRec3.open strSQLf, oConn, 3
                while not oRec3.EOF 

                strTxtFomr = strTxtFomr & oRec3("navn") & ", "

                fo = fo + 1
                oRec3.movenext
                wend
                oRec3.close

                if fo <> 0 then
                len_strTxtFomr = len(strTxtFomr)
                left_strTxtFomr = left(strTxtFomr, len_strTxtFomr-2)
                strTxtFomr = left_strTxtFomr
                end if



                strTxtExport = strTxtExport & Chr(34) & strTxtFomr & Chr(34) & ";"

                
                strTxtExport = strTxtExport & Chr(34) 

                 for p = 1 to 10
        
                pgid = oRec("projektgruppe"&p)

                if pgid <> 1 then
                    call prgNavn(pgid, 200)
                    strTxtExport = strTxtExport & left(prgNavnTxt, 30) & vbCrLf
                end if
        
                next

                strTxtExport = strTxtExport & Chr(34) & ";"

                end if


        if eksDataJsv = 1 then

            if level = 1 then
                if lto <> "epi" OR (lto = "epi" AND (thisMid = 6 OR thisMid = 11 OR thisMid = 59 OR thisMid = 1686 OR thisMid = 1 OR thisMid = 1720)) OR (lto = "epi_no" AND thisMid = 2) OR (lto = "epi_ab" AND thisMid = 2) OR (lto = "epi_sta" AND thisMid = 2) then
                strTxtExport = strTxtExport  & Chr(34) & virksomheds_proc & Chr(34) & ";"
                end if
            end if
       
        
        end if


            'if optiPrint = 2 then
            htmlparseCSV(strJobBesk)
            strJobBesk = htmlparseCSVtxt 

            assci_format(strJobBesk)
            strJobBesk = assci_formatTxt

            strJobBesk = replace(strJobBesk, "&nbsp;", "")
            strJobBesk = replace(strJobBesk, ";", "")
            
            if eksDataNrl = 1 then
		    strTxtExport = strTxtExport & Chr(34) & trim(strJobBesk) & Chr(34) & ";" 
            end if
            
            if isNull(komm) <> true AND len(trim(komm)) > 0 then    

            komm = replace(komm, "</br>", " # ") 
            komm = replace(komm, "<br>", " # ")
            komm = replace(komm, "</ br>", " # ")
            htmlparseCSV(komm)
            strKomm = htmlparseCSVtxt 

            assci_format(strKomm)
            strKomm = assci_formatTxt

            strKomm = replace(strKomm, "&nbsp;", " ")
            strKomm = replace(strKomm, ";", "")
            
            else
            
            strKomm = ""    

            end if
            
		    strTxtExport = strTxtExport & Chr(34) & trim(strKomm) & Chr(34) & ";" 
            

            


             if cint(eksDataAkt) = 1 then

            '*** aktiviteter på job ***
	       strTxtExport = strTxtExport & vbcrlf

         

            lastFase = ""
            strSQL2 = "SELECT id, navn, fase, aktstartdato, aktslutdato FROM aktiviteter WHERE job = "& jobIder(i) & " ORDER BY fase, sortorder, navn"
	        oRec5.open strSQL2, oConn, 3
	        while not oRec5.EOF 
	        
                
                if lastFase <> oRec5("fase") AND isNull(oRec5("fase")) <> true AND len(trim(oRec5("fase"))) <> 0 then
                strTxtExport = strTxtExport & cellerForAkt_og_Faser & Chr(34) & replace(oRec5("fase"), "_", " ") & Chr(34) & vbcrlf 
                lastFase = oRec5("fase")
                end if

           

            strTxtExport = strTxtExport & cellerForAkt_og_Faser &";"& Chr(34) & trim(oRec5("navn")) & Chr(34) &";"& Chr(34) & oRec5("aktstartdato") & chr(34) &";"& Chr(34) & oRec5("aktslutdato") & Chr(34)  & vbcrlf 
	        oRec5.movenext
	        wend
	        oRec5.close
	        

         
            end if


            if cint(eksDataMile) = 1 then

            strTxtExport = strTxtExport & vbcrlf

            if cint(eksDataAkt) = 1 then
            cellerAkt = ";;;;"
            else
            cellerAkt = ""
            end if

            strSQL2 = "SELECT m.id, m.navn, milepal_dato, type, belob, mt.navn AS typenavn FROM milepale AS m LEFT JOIN milepale_typer AS mt ON (mt.id = m.type) WHERE m.jid = "& jobIder(i) & " ORDER BY milepal_dato"
	        oRec5.open strSQL2, oConn, 3
	        while not oRec5.EOF 
	        
                
            strTxtExport = strTxtExport & cellerForAkt_og_Faser & cellerAkt & Chr(34) & trim(oRec5("navn")) & Chr(34) &";"& Chr(34) & oRec5("milepal_dato") & chr(34) &";"& Chr(34) & oRec5("typenavn") & Chr(34) &";"& Chr(34) & oRec5("belob") & Chr(34)  & vbcrlf 
	        oRec5.movenext
	        wend
	        oRec5.close


            end if

         
           
            strTxtExport = strTxtExport & vbcrlf
           
		
		'else 'optiprint
        end if
		
        if optiPrint = 1 then 'Arb. kort
        
        htmlparseCSV(strJobBesk)
        strJobBesk = htmlparseCSVtxt 

        'encodeUTF8(strJobBesk)
        'strJobBesk = 

        strTxtExport = strTxtExport &"Kontakt: "&strKunde&" ("&strKnr&")" & vbcrlf
		strTxtExport = strTxtExport &"Job: " &strJobnavn&" ("&strJobnr&")"  & vbcrlf
		strTxtExport = strTxtExport &"Status: "& strStatus  & vbcrlf
		strTxtExport = strTxtExport &"Brutto Oms.: "& formatnumber(dblBudget, 2) & vbcrlf 
		strTxtExport = strTxtExport &"Udgifter: "& formatnumber(udgifter, 2) & vbcrlf 
		strTxtExport = strTxtExport &"DB %: "& formatnumber(jo_dbproc,0) & vbcrlf 
		strTxtExport = strTxtExport &"Timer forkalk.: "& formatnumber(timerialt, 2) & vbcrlf
		strTxtExport = strTxtExport &"Periode: "& dtStdato &" til " & dtSldato & vbcrlf
		strTxtExport = strTxtExport &"jobansvarlig1: "& strJobans1 & "("& strJobans1nr &") - "& strJobans1init & vbcrlf
		strTxtExport = strTxtExport &"jobansvarlig2: "& strJobans2 & "("& strJobans2nr &") - "& strJobans2init & vbcrlf
		strTxtExport = strTxtExport &"Beskrivelse:" & vbcrlf
		strTxtExport = strTxtExport & strJobBesk & vbcrlf & vbcrlf & vbcrlf 
		strTxtExport = strTxtExport &"Kommentar:" & vbcrlf & ""& komm & vbcrlf 
		strTxtExport = strTxtExport & vbcrlf & vbcrlf & vbcrlf
		strTxtExport = strTxtExport &"----------------------------------------------------------------------------------------------------" & vbcrlf & vbcrlf & vbcrlf
		
		
		end if 'optiprint


        if optiPrint = 6 then 'BU FIL


                '*** FOMR = UnderAFD, UNIT = AFD PÅ JOBBET ****'
                fomrNameTxtjob = ""
				fomrLabelTxtjob = ""
                fak_AfdTxtjob = ""

                
                '**** Fomr ****'
                '** JOB KAN KUN ANTAGE ET FORRETNINGSOMRÅDE **'
                strTxtFomr = ""
                strSQLf = "SELECT for_fomr FROM fomr_rel WHERE for_jobid = "& jobIder(i) & " GROUP BY for_fomr"
                oRec3.open strSQLf, oConn, 3
                if not oRec3.EOF then

             


                            strSQLfomr = "SELECT navn, business_area_label, business_unit FROM fomr WHERE id = "& oRec3("for_fomr")
              
                            oRec4.open strSQLfomr, oConn, 3
				            if not oRec4.EOF then
							 
                            fomrNameTxtjob = oRec4("navn")
				            fomrLabelTxtjob = oRec4("business_area_label")
                            fak_AfdTxtjob = oRec4("business_unit")
				
        		            end if
				            oRec4.close

                  
             
                end if
                oRec3.close

                

         
                'if i <= fo-1 then
                'strTxtExport = strTxtExport & fomrNameTxt(i-1) &";"& fomrLabelTxt(i) &";"
                'else
                'strTxtExport = strTxtExport &";;"
                'end if

                'if i <= kt-1 then
                'strTxtExport = strTxtExport & kundetypeNavnTxt(i-1) &";"& kundetypeLabelTxt(i-1) &";"
                'else
                'strTxtExport = strTxtExport &";;"
                'end if
        
               
                strTxtExportJobUB(antaljobs) = oRec("jobnavn") &";" & oRec("jobnr") &";"& oRec("jobstatus") &";"& fomrLabelTxtjob &";"& fak_AfdTxtjob &";" & vbcrlf
           
             

        end if
		
        antaljobs = antaljobs + 1
		
	end if
	oRec.close


      
                 

	
	next
	

         if optiPrint = 6 then 'UB FIL

                if ktfomrAarMaks < antaljobs then
                highEndArr = antaljobs
                else
                highEndArr = ktfomrAarMaks
                end if
   
                for jbs = 0 TO highEndArr

       

                        '** HVIS DERER VLAGT FÆRRE JOB EN der er Kundetyper / Fomr
                        '** Label er med 2 gange, da De ikke har angivet navn på Adf / Unit. i TO  
                        if jbs <= fo-1 then
                        strTxtExport = strTxtExport & fomrNameTxt(jbs) &";"& fomrLabelTxt(jbs) &";" & fak_AfdTxt(jbs) &";" & fak_AfdTxt(jbs) &";"
                        else
                        strTxtExport = strTxtExport &";;;;"
                        end if

                        'if jbs <= kt-1 then
                        'strTxtExport = strTxtExport & kundetypeNavnTxt(jbs) &";"& kundetypeLabelTxt(jbs) &";"
                        'else
                        'strTxtExport = strTxtExport &";"
                        'end if
        
                        if jbs <= antaljobs-1 then
                        strTxtExport = strTxtExport & strTxtExportJobUB(jbs)
                        else
                        strTxtExport = strTxtExport &";;;;;" & vbcrlf
                        end if

        


                next

         end if


    'else 'optiprint = 3 timepriser
    case 3

            if sorttp = 0 then
            sortTpBy = "a.fase, a.navn, m.mnavn"
            else
            sortTpBy = "m.mnavn, a.fase, a.navn"
            end if

    strTxtExport = strTxtExport & "Kontakt;Kontakt id;Job;Jobnr;Fase;Aktivitet;Medarbejder;Mnr.;Init.;Timepris;Valuta;" & vbcrlf

    for i = 0 to Ubound(jobIder)

                        strSQL = "SELECT tp.jobid, tp.aktid, tp.medarbid, tp.timeprisalt, tp.6timepris, tp.6valuta, j.jobnavn, j.jobnr, a.navn AS aktnavn, a.fase, m.mnavn, m.init, m.mnr, v.valutakode, k.kkundenavn, k.kkundenr FROM timepriser AS tp "_
                        &" LEFT JOIN medarbejdere AS m ON (m.mid = medarbid) "_
                        &" LEFT JOIN aktiviteter AS a ON (a.id = aktid) "_
                        &" LEFT JOIN job AS j ON (j.id = tp.jobid) "_
                        &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                        &" LEFT JOIN valutaer AS v ON (v.id = tp.6valuta) "_
                        &" WHERE jobid = "& jobIder(i) &" AND j.jobnavn IS NOT NULL AND m.mnavn IS NOT NULL GROUP BY tp.jobid, tp.aktid, tp.medarbid ORDER BY j.jobnavn, "& sortTpBy
						'Response.write strSQL & "<br>"
						'Response.flush
                        oRec2.open strSQL, oConn, 3 
						    
                            t = 0
                            while not oRec2.EOF 

                             if IsNull(oRec2("aktnavn")) = False OR oRec2("aktid") = 0 then

						    'thissel = oRec2("timeprisalt")
						    this6timepris = formatnumber(oRec2("6timepris"), 2)
						    this6valuta = oRec2("valutakode")

                            if lastJob <> oRec2("jobid") then
                            strTxtExport = strTxtExport & ";;;;;;;;;;;" & vbcrlf
                            end if

                            strTxtExport = strTxtExport & oRec2("kkundenavn") & ";"& oRec2("kkundenr") &";"& oRec2("jobnavn") & ";"& oRec2("jobnr") &";"& oRec2("fase") &";"& oRec2("aktnavn") &";"& oRec2("mnavn") &";"& oRec2("mnr") &";"& oRec2("init") &";" & this6timepris &";"& this6valuta & ";" & vbcrlf
                            

                            lastJob = oRec2("jobid")
                            t = t + 1
                            
                            end if 'aktnavn

                            oRec2.movenext
						    wend 
						oRec2.close 

    next


  
    end select
    'end if optiprint
	
	objF.WriteLine(strTxtExport)

	'Response.write strTxtExport
	
	objF.close

	Response.redirect "../inc/log/data/"& file &""	
	
end if
%>
<!--#include file="../inc/regular/footer_inc.asp"-->
		
