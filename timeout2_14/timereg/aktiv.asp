<%response.buffer = true%>

 
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/job_func.asp"-->
<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" -->
<!--#include file="inc/isint_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->

<%
'section for ajax calls
if Request.Form("AjaxUpdateField") = "true" then
Select Case Request.Form("control")
case "FM_sortOrder"
Call AjaxUpdate("aktiviteter","sortorder","")
End select
Response.End
end if



if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	func = request("func")
	if func = "dbopr" OR func = "opret" then
	id = 0
	else
	id = request("id")
	end if
	
	'Response.Write "func" & func & "<br>"
	'Response.Write "id:" & id
	'Response.end
	
	if len(trim(id)) <> 0 then
	id = id 
	else
	id = 0
	end if
	
	if len(request("nomenu")) <> 0 AND request("nomenu") <> 0 then
	nomenu = 1
	else
	nomenu = 0
	end if

    if len(trim(request("showtp"))) <> 0 then
    showtp = request("showtp")
    else

        call antalmedarb_fn(2)

        'response.write antalmedarb
        'response.end

        if cdbl(antalmedarb) > 40 then
        'select case lto
        'case "epi", "epi_as", "epi_bar", "epi_ab", ""
        showtp = 0
        'case else
        else
        showtp = 1
        end if
        'end select
    end if
	
	rdir = request("rdir")
	strjobnr = trim(Request("jobid"))
	jobid = strjobnr
	
	aktfavgp = request("aktfavgp")
	
	'** Viser lukkede / passive aktiviteter **'
	if func = "opdaktliste" then
	    if len(trim(request("vispasluk"))) <> 0 then
	    vispasluk = 1
	    vispaslukCHK = "CHECKED"
	    else
	    vispasluk = 0
	    vispaslukCHK = ""
	    end if
	
	else    
	    
	    if request.cookies("tsa")("vispasluk") <> "" then
	    vispasluk = request.cookies("tsa")("vispasluk")
	        if vispasluk = 1 then
	        vispasluk = 1
	        vispaslukCHK = "CHECKED"
	        else
	        vispasluk = 0
	        vispaslukCHK = ""
	        end if
	    
	    else
	        '** vis passive / aktive akt. ved hent side LTO indstilling (sættes ved opdaterliste)
	        select case lto
	        case "q2con"
	        vispasluk = 0
	        vispaslukCHK = ""
	        case else
	        vispasluk = 1
	        vispaslukCHK = "CHECKED"
	        end select
	    end if
	end if
	
	response.cookies("tsa")("vispasluk") = vispasluk
	

        visbeskCHK = "" 'sættes i jquery

   
	
	
	
	
	'******************************* funktioner **********************************************
	
	

	
	
	
	
	public strTdato
	public strUdato
	public intjobnrthis
	public totaltimertildelt
	public jobtotaltimer
	public ikkeBudgettimer
	public totalikkefaktimertildelt
	public jobBudget
	public jobAntal
	
	
	function budgettimerTildelt(jobidthis)
	strSQL = "SELECT jobstartdato, jobslutdato, budgettimer, ikkebudgettimer, jobnr, jobtpris FROM job WHERE id = " & jobidthis
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	intjobnrthis = oRec("jobnr")
		
		'*** Skal dato nedarves fra job?? (kun ved oprettelse) ***
		if func = "opret" then
		strTdato = oRec("jobstartdato")
		strUdato = oRec("jobslutdato")
		else
		strTdato = strTdato
		strUdato = strUdato
		end if
	
	jobBudget = oRec("jobtpris")
	jobtotaltimer = oRec("budgettimer")
	ikkeBudgettimer = oRec("ikkebudgettimer")
	
	jobAntal = jobtotaltimer + ikkeBudgettimer
	
	end if
	oRec.close
	
	if len(jobBudget) <> 0 then
	jobBudget = jobBudget
	else
	jobBudget = 0
	end if
	
	if len(jobtotaltimer) <> 0 then
	jobtotaltimer = jobtotaltimer
	else
	jobtotaltimer = 0
	end if
	
	
	strSQL = "SELECT sum(budgettimer) AS akttimer FROM aktiviteter WHERE job = " & jobidthis &" ORDER BY budgettimer"
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
		akttotaltimer = oRec("akttimer")
	end if
	oRec.close
	
	
	if len(akttotaltimer) <> 0 then
	akttotaltimer = akttotaltimer
	else
	akttotaltimer = 0
	end if
	
	
	'** Rest timer på job **'
	totaltimertildelt = (jobtotaltimer - akttotaltimer)
	
	if len(totaltimertildelt) <> 0 then
	totaltimertildelt = totaltimertildelt
	else
	totaltimertildelt = 0
	end if
	
	
	
	
	
	end function
	'******************************* funktioner slut  **********************************************
	
	
	
	
	select case func
	case "slet"
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	
	
	slttxtalt = ""
	slturlalt = ""
	
	slttxt = "Du er ved at <b>slette</b> en aktivitet.<br>"_
    &"Du vil samtidig slette alle timeregistreringer på denne aktivitet.<br>"_
	&"Timeregistreringerne vil <b>ikke kunne genskabes</b>. <br>"
    slturl = "aktiv.asp?menu=job&func=sletok&jobid="&jobid&"&jobnavn="&request("jobnavn")&"&id="&id&"&rdir="&rdir
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,210,100)
	
	
	
	case "sletok"
	'*** Her slettes en aktivitet ***
	
	call delakt(id)
	
	
	
	Response.redirect "aktiv.asp?menu=job&jobid="&jobid&"&jobnavn="&request("jobnavn")&"&rdir="&rdir
	
	case "opdprogp"
	
			jobid = request("jobid")
			strSQL = "SELECT projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM job WHERE id = "& jobid 
	
								oRec2.open strSQL, oConn, 3
									if not oRec2.EOF then
									
									oConn.execute("UPDATE aktiviteter SET"_
									&" projektgruppe1 = "& oRec2("projektgruppe1") &" , projektgruppe2 = "& oRec2("projektgruppe2") &", "_
									&" projektgruppe3 = "& oRec2("projektgruppe3") &", "_
									&" projektgruppe4 = "& oRec2("projektgruppe4") &", "_
									&" projektgruppe5 = "& oRec2("projektgruppe5") &", "_
									&" projektgruppe6 = "& oRec2("projektgruppe6") &", "_
									&" projektgruppe7 = "& oRec2("projektgruppe7") &", "_
									&" projektgruppe8 = "& oRec2("projektgruppe8") &", "_
									&" projektgruppe9 = "& oRec2("projektgruppe9") &", "_
									&" projektgruppe10 = "& oRec2("projektgruppe10") &" WHERE job = "& jobid &"")
									
									end if
								oRec2.close
			
            
        Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
        %>
        <br /><br />
        Projektgrupper fra job overført til aktiviteter.<br />
        <a href="#" onClick="javascript:window.close();">Luk vindue</a> 
        <%
            'Response.Write("<script language=""JavaScript"">window.opener.close();</script>")
        	'Response.redirect "aktiv.asp?menu=job&jobid="&jobid&"&jobnavn="&request("jobnavn")&"&rdir="&rdir
			'Response.redirect "jobs.asp?menu=job&shokselector=1&id="&jobid
			
	case "opdaktliste"
	
	
	
	if len(request("FM_aktnavn")) > 3 then
	len_aktnavn = len(request("FM_aktnavn"))
	left_aktnavn = left(request("FM_aktnavn"), len_aktnavn - 3)
	strAktnavn = left_aktnavn
	else
	strAktnavn = ""
	end if
	
	aktNavn = split(strAktnavn, ", #,")
	akttimer = split(request("FM_timer"), ", ")
	aktantalstk = split(request("FM_antalstk"), ", ")
	aktpris = split(request("FM_pris"), ", ")
	aktids = split(request("FM_listeid"), ",")
	aktstatus = split(request("FM_listestatus"), ",")
    'aktsortering = split(request("FM_sortorder"), ", ")
    aktEasyreg = split(request("FM_easyreg_aid"), ", ")
   
	
	if len(request("FM_fase")) > 3 then
	len_Fasenavn = len(request("FM_fase"))
	left_Fasenavn = left(request("FM_fase"), len_Fasenavn - 3)
	strFasenavn = left_Fasenavn
	else
	strFasenavn = ", #,"
	end if

    aktType = split(request("FM_fakturerbart"), ", ")

    'Response.Write strFasenavn
    'Response.end
	
    'Response.Write "request(FM_st_dato)" & request("FM_st_dato") & "<br>"
    'Response.flush

    aktStDato = split(request("FM_st_dato"), ", ")
    aktSlDato = split(request("FM_sl_dato"), ", ")

	aktFaser = split(strFasenavn, ", #,")
	aktBgr = split(request("FM_bgr"), ",")
	
	'Response.Write request("FM_akt_totpris")
	'Response.end
	
	akttotpris = split(request("FM_akt_totpris"), ", #")
	
	'Response.Write request("FM_fase")
	'Response.end
	

    'Response.Write request("FM_faktor")
    'Response.end
    aktFaktor = split(request("FM_faktor"), ", #,")
	aktSlet = split(request("FM_slet_aid"), ", ")
	
	'Response.Write aktSlet
	'Response.end
	
    treg_uojb_ea_on = 0

    'Response.write request("FM_beskrivelse_4099")
    'Response.end
	
	
	for t = 0 to UBOUND(aktids)
	err = 0
	
	    akttimer(t) = replace(akttimer(t), ".", "")
	    akttimer(t) = replace(akttimer(t), ",", ".")
	    
	     if len(trim(akttimer(t))) <> 0 then
	    akttimer(t) = akttimer(t)
	    else
	    akttimer(t) = 0
	    end if
	    
	    call erDetInt(akttimer(t))
	    
	      if cint(isInt) > 0 then
	      err = 1
	      end if
	      
	      isInt = 0
	    
	    aktantalstk(t) = replace(aktantalstk(t), ".", "")
	    aktantalstk(t) = replace(aktantalstk(t), ",", ".")
	    
	    if len(trim(aktantalstk(t))) <> 0 then
	    aktantalstk(t) = aktantalstk(t)
	    else
	    aktantalstk(t) = 0
	    end if
	    
	    call erDetInt(aktantalstk(t))
	        
	      if cint(isInt) > 0 then
	      err = 1
	      end if
	      
	      isInt = 0
	    
	    aktpris(t) = replace(aktpris(t), ".", "")
	    aktpris(t) = replace(aktpris(t), ",", ".")
	    
	    if len(trim(aktpris(t))) <> 0 then
	    aktpris(t) = aktpris(t)
	    else
	    aktpris(t) = 0
	    end if
	    
	    call erDetInt(aktpris(t))
	      
	      if cint(isInt) > 0 then
	      err = 1
	      end if
	      
	      isInt = 0
	      
	      akttotpris(t) = replace(akttotpris(t), "#", "")
	      akttotpris(t) = replace(akttotpris(t), ".", "")
	      akttotpris(t) = replace(akttotpris(t), ",", ".")
	      
	     
	      'Response.Write "err:" & err & "<br>"
	      'Response.Write "int:" & isInt &"<br>"
	      
	      'Response.flush
	    aktFaser(t) = trim(aktFaser(t))
		call illChar(aktFaser(t))
		aktFaser(t) = vTxt
		
	    
	    
	    aktNavn(t) = trim(aktNavn(t))
	    aktNavn(t) = replace(aktNavn(t), "'", "")
	    

        'aktStDato = 

        aktStDato(t) = replace(aktStDato(t), ".", "-")
        if isDate(aktStDato(t)) = true then
        aktStDato(t) = year(aktStDato(t)) & "/" & month(aktStDato(t)) & "/" & day(aktStDato(t))
        else
        aktStDato(t) = "2002-01-01"
        end if

        aktSlDato(t) = replace(aktSlDato(t), ".", "-")
        if isDate(aktSlDato(t)) = true then
        aktSlDato(t) =  year(aktSlDato(t)) & "/" & month(aktSlDato(t)) & "/" & day(aktSlDato(t))
        else
        aktSlDato(t) = "2002-01-01"
        end if

        if instr(aktFaktor(t), "#") <> 0 then
        len_aktFaktor = len(aktFaktor(t))
        left_aktFaktor = left(aktFaktor(t), len_aktFaktor - 3)
        aktFaktor(t) = left_aktFaktor

        'Response.Write "<br>her:" & aktFaktor(t) & "<br>"

        else
        aktFaktor(t) = aktFaktor(t)
        end if

        aktFaktor(t) = replace(aktFaktor(t), ".", "")
        aktFaktor(t) = replace(aktFaktor(t), ",", ".")
        
        


	    if cint(err) = 0 then
	    
	    aktSletval = request("FM_slet_aid_"& aktSlet(t) &"")
		
		if aktSletval <> "1" then

          aktEasyregval = request("FM_easyreg_aid_"& aktEasyreg(t) &"")
          
          if cint(aktEasyregval) <> 0 then
          aktEasyregval = aktEasyregval
          treg_uojb_ea_on = 1
          else
          aktEasyregval = 0
          end if

         aktBesk = request("FM_beskrivelse_"& trim(aktids(t)))
           
         call htmlreplace(aktBesk)
         aktBesk = htmlparseTxt
         aktBesk = replace(aktBesk, "'", "''")
		
		strSQL = "UPDATE aktiviteter SET navn = '"& aktNavn(t) &"', aktstatus = " & aktstatus(t) & ", fakturerbar = "& aktType(t) &", "_
		&" budgettimer = "& akttimer(t) &", aktbudget = "& aktpris(t) &", antalstk = "& aktantalstk(t) &", easyreg = "& aktEasyregval &", " _ 
        &" aktstartdato = '"& aktStDato(t) &"', aktslutdato = '"& aktSlDato(t) &"', "
		
        if len(trim(aktFaser(t))) <> 0 then
        strSQL = strSQL &" fase = '"& aktFaser(t) &"', "
        else
        strSQL = strSQL &" fase =  NULL, "
        end if
        
        strSQL = strSQL &" bgr = "& aktBgr(t) &", aktbudgetsum = "& akttotpris(t) &", faktor = ABS("& aktFaktor(t) &"), beskrivelse = '"& aktBesk &"'"_
		&" WHERE id = "& aktids(t)
		
        ', sortorder = "& aktsortering(t) &"
		'Response.write strSQL & "("& aktFaktor(t) &")<br>"
		'Response.flush
		oConn.execute(strSQL)

            '*** Opdaterer timereg. tabellen ***'
            oConn.execute("UPDATE timer SET "_
            & " TAktivitetNavn = '"& aktNavn(t) &"', "_
            & " TFaktim = "& aktType(t) &""_
            & " WHERE TAktivitetId = "& aktids(t) & "")


       else
		    
		    call delakt(aktids(t))
		
		end if
		
		'** Sync job ***'
		if len(trim(request("opdjobv"))) <> 0 then 
		opdjobv = 1		
		else
		opdjobv = 0
		end if		
		
		
		if cint(opdjobv) = 1 then
            call syncJob(jobid)
        end if
		
		end if
	
	next

    '*** Timereg_usejob Easyreg *****'
	if cint(treg_uojb_ea_on) = 1 then
    strSQLtu = "UPDATE timereg_usejob SET easyreg = "& jobid &" WHERE jobid = " & jobid
    oConn.execute(strSQLtu)
    end if

   
        

	'Response.end
	Response.redirect "aktiv.asp?menu=job&jobid="&jobid&"&jobnavn="&request("jobnavn")&"&rdir="&rdir&"&nomenu="&nomenu
							
			
	
	
	
	case "dbopr", "dbred"
	
	
	
	'*** Her indsættes en ny aktivitet i db ****
	            
	            '*tjekker om dato eksisterer og tildeler ny dato hvis den ikke gør ** 
				call dato_30(Request("FM_start_dag"), Request("FM_start_mrd"), Request("FM_start_aar"))
				strStartDay = strDay_30
				
				call dato_30(Request("FM_slut_dag"), Request("FM_slut_mrd"), Request("FM_slut_aar"))
				strSlutDay = strDay_30
				
				
				
	
	            slutaar = Request("FM_slut_aar")
	            if slutaar = "44" then
	            slutaar = "2044"
	            else
	            slutaar = slutaar
	            end if
	
	
		
		function SQLBlessK(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ".", ",")
		SQLBlessK = tmp
		end function
		
		function SQLBlessP(s)
		dim tmp2
		tmp2 = s
		tmp2 = replace(tmp2, ",", ".")
		SQLBlessP = tmp2
		end function
		
		
		            if len(request("FM_budgettimer")) <> 0 then
		            
		            useBudgettimer = request("FM_budgettimer")
			            
			            if useBudgettimer < 0 then
			            useBudgettimer = 0
			            else
			            useBudgettimer = replace(useBudgettimer, ".", "")
		                useBudgettimer = replace(useBudgettimer, ",", ".")
			            end if
			            
		            else
		            useBudgettimer = 0
		            end if
            		
		            if len(request("FM_faktor")) <> 0 then
		            dblFaktor = replace(request("FM_faktor"), ".", "")
		            dblFaktor = replace(dblFaktor, ",", ".")
		            else
		            dblFaktor = 1
		            end if
            		
		            if len(request("FM_budget")) <> 0 then
		            intBudget = replace(request("FM_budget"), ".", "")
		            intBudget = replace(intBudget, ",", ".")
		            else
		            intBudget = 0
		            end if
		            
		            if len(request("FM_akt_totpris")) <> 0 then
		            intBudgetTot = replace(request("FM_akt_totpris"), ".", "")
		            intBudgetTot = replace(intBudgetTot, ",", ".")
		            else
		            intBudgetTot = 0
		            end if
		            
		            
		            
		            if len(request("FM_antalstk")) <> 0 then
		            antalstk = replace(request("FM_antalstk"), ".", "")
		            antalstk = replace(antalstk, ",", ".")
		            else
		            antalstk = 0
		            end if
		            

                    if len(trim(request("FM_brug_fasttp"))) <> 0 then
                    brug_fasttp = 1
                    else
                    brug_fasttp = 0
                    end if

                     
                    if len(trim(request("FM_fasttp"))) <> 0 then
                    fasttp = request("FM_fasttp")

                        fasttp = replace(fasttp, ".", "")
		                fasttp = replace(fasttp, ",", ".")
			        else
                        fasttp = 0 
                    end if  

                    if len(trim(request("FM_fasttp_val"))) <> 0 then
                    fasttp_val = request("FM_fasttp_val")
                    else
                    fasttp_val = 0
                    end if
                    
                   
                    if len(trim(request("FM_brug_fastkp"))) <> 0 then
                    brug_fastkp = 1
                    else
                    brug_fastkp = 0
                    end if

                     if len(trim(request("FM_fastkp"))) <> 0 then
                    fastkp = request("FM_fastkp")

                        fastkp = replace(fastkp, ".", "")
		                fastkp = replace(fastkp, ",", ".")



			        else
                        fastkp = 0 
                    end if
                    
                    
                    '** Cost price always in Basic Valuta **'
                    'if len(trim(request("FM_fastkp_val"))) <> 0 then  
                    'fastkp_val = request("FM_fastkp_val")
                    'else
		            fastkp_val = 0
                    'end if
		
		
		if len(request("FM_navn")) = 0 OR instr(request("FM_navn"), "'") <> 0 then 'OR len(request("FM_jnr")) = 0 OR startDatoNum > slutDatoNum
		%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		
		<%
		errortype = 23
		call showError(errortype)
		else
			
			
			if instr(request("FM_fase"), "'") <> 0 then
			%>
		    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		    
		    <%
		    errortype = 141
		    call showError(errortype)
		    else
			
			
			
			
            call erDetInt(fasttp)
            if isInt > 0 then
					%>
					<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
				
					<%
					errortype = 160
					call showError(errortype)
					isInt = 0
		
				    else


                    call erDetInt(fastkp)
                    if isInt > 0 then
					%>
					<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					
					<%
					errortype = 160
					call showError(errortype)
					isInt = 0
		
				    else
			
			call erDetInt(useBudgettimer)
			
					if isInt > 0 then
					%>
					<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					
					<%
					errortype = 20
					call showError(errortype)
					isInt = 0
		
				    else
				
					        call erDetInt(dblFaktor)
					        if isInt > 0 then
					        %>
					        <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					        <%
					        errortype = 51
					        call showError(errortype)
					        isInt = 0
        					
        					
					        else
				
					                    
					                    call erDetInt(antalstk)
					                    if isInt > 0 then
					                    %>
					                    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					                   <%
					                    errortype = 122
					                    call showError(errortype)
					                    isInt = 0
                    					
                    					
                    					
					                    else
					
					
				                        '*** Tidslås ***
				                        if len(request("FM_tidslaas")) <> 0 then
				                        tidslaas = 1
				                        else
				                        tidslaas = 0
				                        end if
                        				
				                        if tidslaas = 1 then
				                        tidslaas_st = request("FM_tidslaas_start") '& ":00"
				                        tidslaas_sl = request("FM_tidslaas_slut")  '& ":00"
				                        else
				                        tidslaas_st = "07:30:00"
				                        tidslaas_sl = "23:30:00"
				                        end if
                        				
				                        if len(trim(request("FM_tidslaas_man"))) <> 0 then
				                        tidslaas_man = 1
				                        else 
				                        tidslaas_man = 0
				                        end if
                        				
				                        if len(trim(request("FM_tidslaas_tir"))) <> 0 then
				                        tidslaas_tir = 1
				                        else 
				                        tidslaas_tir = 0
				                        end if
                        				
				                        if len(trim(request("FM_tidslaas_ons"))) <> 0 then
				                        tidslaas_ons = 1
				                        else 
				                        tidslaas_ons = 0
				                        end if
                        				
				                        if len(trim(request("FM_tidslaas_tor"))) <> 0 then
				                        tidslaas_tor = 1
				                        else 
				                        tidslaas_tor = 0
				                        end if
                        				
				                        if len(trim(request("FM_tidslaas_fre"))) <> 0 then
				                        tidslaas_fre = 1
				                        else 
				                        tidslaas_fre = 0
				                        end if
                        				
				                        if len(trim(request("FM_tidslaas_lor"))) <> 0 then
				                        tidslaas_lor = 1
				                        else 
				                        tidslaas_lor = 0
				                        end if
                        				
				                        if len(trim(request("FM_tidslaas_son"))) <> 0 then
				                        tidslaas_son = 1
				                        else 
				                        tidslaas_son = 0
				                        end if
                        		        
	            
	            
	            
	           
				
				
				function SQLBless3(s3)
				dim tmp3
				tmp3 = s3
				tmp3 = replace(tmp3, ":", "")
				SQLBless3 = tmp3
				end function
				
				
				'if cint(tidslaas) = 1 then 
				
				
				idagErrTjek = day(now)&"/"&month(now)&"/"&year(now)
				
				
				for t = 1 to 2
				
				select case t
				case 1
				tdsl = tidslaas_st
				case 2
				tdsl = tidslaas_sl
				end select
				
				'Response.write SQLBless3(trim(tSttid(y))) & "<br>"
				'Response.flush
				
				call erDetInt(SQLBless3(trim(tdsl)))
				if isInt > 0 then
					antalErr = 1
					errortype = 63
					useleftdiv = "t"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end
				end if	
				isInt = 0


                '** Sortorder ***'
                if len(request("SortOrder")) <> 0 then
                sortorder = trim(request("SortOrder"))
                sortorder = replace(sortorder, ",", ".")

                        call erDetInt(sortorder)
				        if isInt > 0 then
					        antalErr = 1
					        errortype = 153
					        'useleftdiv = "t"
					        %><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					        call showError(errortype)
					        response.end
				        end if	
				        isInt = 0

                else
                sortorder =  1000
                end if
				
				'*** Punktum  i angivelse ved registrering af klokkeslet
				if instr(trim(tdsl), ".") <> 0 OR instr(trim(tdsl), ",") <> 0 then
					antalErr = 1
					errortype = 66
					useleftdiv = "t"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end
				end if	
				
				if len(trim(tdsl)) <> 0 then
				
				'Response.write idagErrTjek &" "& tSttid(y)&":00" &" - "& isdate(idagErrTjek &" "& tSttid(y)&":00") &"<br>"
					if isdate(idagErrTjek &" "& tdsl) = false then
						antalErr = 1
						errTid = tdsl
						errortype = 148
						useleftdiv = "t"
						%>
						<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
						<%
						call showError(errortype)
						response.end
					end if
				end if
				
			
			next
					
			'''end if 'tidslås
			
				
				function SQLBless(s)
				dim tmp
				tmp = s
				tmp = replace(tmp, "'", "''")
				SQLBless = tmp
				end function	
				
				if len(request("jobid")) <> 0 then
				jobid = request("jobid")
				else
				jobid = 0
				end if
				
				strNavn = replace(request("FM_navn"), "'", "''")
				
				strFase = trim(request("FM_fase"))
				call illChar(strFase)
				strFase = trim(vTxt)

                'Response.Write "vTxt" &  vTxt
				'Response.end
				
				if len(trim(strFase)) <> 0 then
				strFase = strFase
				else
				strFase = ""
				end if
				
				strbeskrivelse = replace(request("FM_beskrivelse"), "'", "''")
				
				'*** HTML Replace **'
				call htmlreplace(strbeskrivelse)
				strbeskrivelse = htmlparseTxt
				
				strOLDaktId = request("FM_OLDaktId")
				
				if len(request("FM_fakturerbart")) <> 0 then
				strFakturerbart = request("FM_fakturerbart")
				else
				strFakturerbart = 0
				end if
				
				int_aktFav = request("FM_favorit")
				
				intaktstatus = request("FM_aktstatus")
				strEditor = session("user")
				strDato = session("dato")
				
				intFomr = 0 'request("FM_fomr")
				
				if len(trim(request("FM_easyreg"))) <> 0 then
				easyreg = 1
				else
				easyreg = 0
				end if

                if len(trim(request("FM_avarenr"))) <> 0 then
               
                avarenr = request("FM_avarenr")
                 call alfanumerisk(avarenr)
                 avarenr = alfanumeriskTxt
               
                else
                avarenr = ""
                end if


             

                if len(trim(request("FM_kostpristarif"))) <> 0 then
                kostpristarif = request("FM_kostpristarif")
                else
                kostpristarif = 0
                end if

                  
                if len(trim(request("FM_fravalgt"))) <> 0 then
                fravalgt = 1
                else
                fravalgt = 0
                end if
				
				'*** Hvis den ikke er Easyreg i forvejen, Aktiver PUSH, gør den aktiv foralle medarbejdere **'
				'*** Gælder alle easy-reg aktiviter på jobbet ***'
				if cint(easyreg) = 1 then 'AND func <> "dbopr"
				
				olEasyreg = 0
				strSQLea = "SELECT easyreg FROM aktiviteter WHERE id = "& id
				oRec.open strSQLea, oConn, 3
				if not oRec.EOF then
				 olEasyreg = oRec("easyreg")
				end if
				oRec.close
				    
				    '** PUSH aktiv **'
				    if cint(olEasyreg) = 0 then
				    strSQLtreguse = "UPDATE timereg_usejob SET easyreg = " & jobid & " WHERE jobid = " & jobid 'id
                    'Response.Write strSQLtreguse & "<br>"
	   	            'Response.flush
	   	            oConn.execute(strSQLtreguse)
				    end if
				
				end if
				
				
				
				intBgr = request("FM_bgr")
					
				startDato = Request("FM_start_aar") & "/" & Request("FM_start_mrd") & "/" & strStartDay 
				slutDato = Request("FM_slut_aar") & "/" & Request("FM_slut_mrd") & "/" & strSlutDay 
				
				
			    jobids = split(request("FM_jnr"), ", ")
				jobidswrt = ""
				
				
                 '************************************'
                 '***** Forrretningsområder **********'
                 '************************************'

                    fomrArr = split(request("FM_fomr"), ",")

                    for_faktor = 0
                    for afor = 0 to UBOUND(fomrArr)
                    for_faktor = for_faktor + 1 
                    next

                    if for_faktor <> 0 then
                    for_faktor = for_faktor
                    else
                    for_faktor = 1
                    end if

                    for_faktor = formatnumber(100/for_faktor, 2)
                    for_faktor = replace(for_faktor, ",", ".")
                  

                  '**************************************'
				
				'Response.Write "her" & request("FM_jnr")
	            'Response.end



                '***** KONTO (kun hvis slået il i kontrolpanel) ********************
			    if len(trim(request("FM_aktkonto"))) <> 0 then
                aktkonto = request("FM_aktkonto")
                else
                aktkonto = 0
                end if


				
				for j = 0 TO UBOUND(jobids)
				
				if instr(jobidswrt, ",#"&jobids(j)&"#") = 0 then
				jobidswrt = jobidswrt & ",#"&jobids(j)&"#"
				
				
                
				
                '**** Projektgrupper ****'
                aj = 2 '** Hvad bliver opdateret? Aktivitet eller job 1 = job, 2 = akt.
                aid = 0 '** Aktid
                
                if len(trim(request("FM_pgrp_arvefode"))) <> 0 then
                pgrp_arvefode = request("FM_pgrp_arvefode")
                else
                pgrp_arvefode = -1
                end if 
                
                nedfod = pgrp_arvefode '-1 '*** ignorer fød/nedarv -1 = ignorer, 0 Nedarv,  1 = fød job
                'fm_alle = -1 'request("FM_alle") '*** bruges kun ved oprettelse af enkelt aktivitet. Skal den følge job's projektgrupper eller beholde sinde egne. 1 = Behold, 0/"" = Nedarv fra job, - 1 ignorer (ved jobopr / rediger job)
                jid = jobids(j) 'jobid
                
                firstLoop = 0

                'Response.write nedfod
                'Response.end

                call tilfojProGrp(func,aj,jid,aid,nedfod,request("FM_projektgruppe_1"),request("FM_projektgruppe_2"),request("FM_projektgruppe_3"),request("FM_projektgruppe_4"),request("FM_projektgruppe_5"),request("FM_projektgruppe_6"),request("FM_projektgruppe_7"),request("FM_projektgruppe_8"),request("FM_projektgruppe_9"),request("FM_projektgruppe_10"), firstLoop)

                        
				
				


				
				'**** Er det stamakt. eller alm. akt. ***'
				if jobids(j) = 0 then
				int_aktFav = int_aktFav
				rdir_int_aktFav = int_aktFav
				else
				int_aktFav = 0
				end if
				
				
				
				
				
				
				
						if func = "dbopr" then
						
						
						        '** ST DATO og SL DATO ved opret stamakt og tildel på job ***'
						        if jobids(j) <> 0 AND jobid = 0 then 
        						
        						
						        strSQL = "SELECT jobstartdato, jobslutdato FROM job WHERE id =" & jobids(j)	
	                            oRec.open strSQL, oConn, 3	
	                            if not oRec.EOF Then
	                            jobstdate = oRec("jobstartdato") 
	                            jobsldate = oRec("jobslutdato")
	                            else
	                            jobstdate = "1/1/2000"
	                            jobsldate = "1/1/2044"
	                            end if
	                            oRec.close
        	                    
	                            startDato = year(jobstdate) &"/"& month(jobstdate) &"/"& day(jobstdate)
	                            slutDato = year(jobsldate) &"/"& month(jobsldate) &"/"& day(jobsldate)
        	                    
	                            else
        	                    
	                            startDato = startDato 
	                            slutDato = slutDato 
        	                    
	                            end if 
						
						
						
						strSQL = "INSERT INTO aktiviteter (navn, beskrivelse, job, fakturerbar, aktfavorit, aktstartdato, aktslutdato, "_
						&" editor, dato, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, "_
						&" projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, "_
						&" budgettimer, aktstatus, fomr, faktor, aktbudget, tidslaas, tidslaas_st, tidslaas_sl, antalstk, "_
						&" tidslaas_man, tidslaas_tir, tidslaas_ons, "_
						&" tidslaas_tor, tidslaas_fre, tidslaas_lor, tidslaas_son, fase, bgr, aktbudgetsum, easyreg, sortorder, fravalgt, brug_fasttp, "_
                        &" brug_fastkp, fasttp, fasttp_val, fastkp, fastkp_val, avarenr, kostpristarif, aktkonto "_
						&") VALUES "_
						&" ('"& strNavn &"', "_
						&"'"& strBeskrivelse &"', "_ 
						&""& jobids(j) &", "_  
						&""& strFakturerbart &", "_ 
						&""& int_aktFav &", "_ 
						&"'"&startDato &"', "_ 
						&"'"&slutDato &"', "_
						&"'"&strEditor &"', "_
						&"'"&strDato &"', "_ 
						&""&strProjektgr1 &", "_ 
						&""&strProjektgr2 &", "_
						&""&strProjektgr3 &", "_
						&""&strProjektgr4 &", "_
						&""&strProjektgr5 &", "_
						&""&strProjektgr6 &", "_ 
						&""&strProjektgr7 &", "_
						&""&strProjektgr8 &", "_
						&""&strProjektgr9 &", "_
						&""&strProjektgr10 &", "_       
						&""& SQLBlessP(useBudgettimer) &", "_
						&""& intaktstatus &", "& intFomr &", "& dblFaktor &", "& intBudget &", "_
						&" "& tidslaas &", '"& tidslaas_st &"', '"& tidslaas_sl &"', "_
						&" "& antalstk &", "_
						&" "& tidslaas_man &", "& tidslaas_tir &", "& tidslaas_ons &", "_
						&" "& tidslaas_tor &", "& tidslaas_fre &", "& tidslaas_lor &", "& tidslaas_son 
						
						if len(trim(strFase)) <> 0 then
						strSQL = strSQL & ", '"& strFase &"', "
						else
						strSQL = strSQL & ", NULL , "
						end if
						
						strSQL = strSQL & intBgr &", "& intBudgetTot &", "& easyreg &", "& sortorder &", "& fravalgt &", "& brug_fasttp &","_
                        &""& brug_fastkp &", "& fasttp &", "& fasttp_val &", "& fastkp &", "& fastkp_val &", '"& avarenr &"', '"& kostpristarif &"', "& aktkonto &")"
						
						
						
						'Response.write strSQL & "<br>"
						'Response.flush
						'Response.end
						
						oConn.execute(strSQL)
						
						strSQL = ""
						
						else
						
						
						
						strSQLupd = "UPDATE aktiviteter SET "_
						&" navn = '"& strNavn &"', "_
						&" beskrivelse ='"& strBeskrivelse &"', "_ 
						&" job = "& jobids(j) &", "_  
						&" fakturerbar = "& strFakturerbart &", "_ 
						&" aktfavorit = "& int_aktFav &", "_ 
						&" projektgruppe1 = "& strProjektgr1 &", "_ 
						&" projektgruppe2 = "& strProjektgr2 &", "_ 
						&" projektgruppe3 = "& strProjektgr3 &", " _
						&" projektgruppe4 = "& strProjektgr4 &", " _
						&" projektgruppe5 = "& strProjektgr5 &", " _
						&" projektgruppe6 = "& strProjektgr6 &", "_ 
						&" projektgruppe7 = "& strProjektgr7 &", "_ 
						&" projektgruppe8 = "& strProjektgr8 &", " _
						&" projektgruppe9 = "& strProjektgr9 &", " _
						&" projektgruppe10 = "& strProjektgr10 &", " _
						&" aktstartdato = '"& startDato &"', "_ 
						&" aktslutdato = '"& slutDato &"', "_
						&" editor = '"& strEditor &"', "_
						&" budgettimer = "& useBudgettimer &", "_
						&" dato = '"& strDato &"', "_
						&" aktstatus = "& intaktstatus &", fomr = "& intFomr &", "_
						&" faktor = "& dblFaktor &", aktbudget = "& intBudget &", "_
						&" tidslaas = "& tidslaas &", tidslaas_st = '"& tidslaas_st &"', "_
						&" tidslaas_sl = '"& tidslaas_sl &"', antalstk = "& antalstk &", "_
						&" tidslaas_man = "& tidslaas_man &", tidslaas_tir = "& tidslaas_tir &", "_
						&" tidslaas_ons = "& tidslaas_ons &", tidslaas_tor = "& tidslaas_tor &", "_
						&" tidslaas_fre = "& tidslaas_fre &", tidslaas_lor = "& tidslaas_lor &", "_
						&" tidslaas_son = "& tidslaas_son &", bgr = "& intBgr &", aktbudgetsum = "& intBudgetTot &", easyreg = "& easyreg &", sortorder = "& sortorder &", fravalgt = " & fravalgt &", "_
                        &" brug_fasttp = "& brug_fasttp &", brug_fastkp = "& brug_fastkp &", fasttp = "& fasttp &", fasttp_val = "& fasttp_val &","_
                        &" fastkp = "& fastkp&", fastkp_val = "& fastkp_val &", avarenr = '"& avarenr &"', kostpristarif = '"& kostpristarif &"', aktkonto = "& aktkonto &"" 
						
						if len(trim(strFase)) <> 0 then
						strSQLupd = strSQLupd & ", fase = '"& strFase &"'"
						else
						strSQLupd = strSQLupd & ", fase = NULL"
						end if
						
						 
						strSQLupd = strSQLupd &" WHERE id = "& id 
						
						
						oConn.execute(strSQLupd)
						strSQLupd = ""
						
						
                        
                        tfaktimvalue = strFakturerbart
						


                        
                        '*** Finder nyeste kunde og job oplysninger ***'
                        
                        strKundenavn = "..if"
                        intKid = 0
                        strJobnavn = "..if"
                        strJobnr = 0
                        
                        strSQLjobogKunde = "SELECT k.kkundenavn, k.kid, j.jobnavn, j.jobnr FROM job AS j "_
                        &" LEFT JOIN kunder AS k ON (k.kid = jobknr) WHERE j.id = "& jobids(j) 
						oRec6.open strSQLjobogKunde, oConn, 3
                        if not oRec6.EOF then

                        strKundenavn = oRec6("kkundenavn")
                        intKid = oRec6("kid")
                        strJobnavn = oRec6("jobnavn")
                        strJobnr = oRec6("jobnr")

                        end if
                        oRec6.close

						'*** Opdaterer time tabellen ***
                        '*** timepriser kan afvige på det nye projekt ***'
						oConn.execute("UPDATE timer SET "_
						& " TAktivitetNavn = '"& strNavn &"', "_
						& " TFaktim = "& tfaktimvalue &", "_
                        & " Tknavn = '"& replace(strKundenavn, "'", "''") &"', Tknr = "& intKid &", "_
						& " Tjobnavn = '"& strJobnavn &"', "_
						& " Tjobnr = '"& strJobnr &"'"_
						& " WHERE TAktivitetId = "& strOLDaktId &""_
						& "")
						



                                        

                                        

						end if
						
			    '*************************************'
				'*** Opdater værdi og timer på job ***'
				'*************************************'		
				if len(trim(request("opdjobv"))) <> 0 then 
				opdjobv = 1		
				else
				opdjobv = 0
				end if		
				
				
				if cint(opdjobv) = 1 then
				
				call syncJob(jobids(j))
				
				           
				
				end if
				'**************************************'
				'*************************************'
						
						
				if func = "dbopr" then
						
						'*** Henter det netop oprettede akt-id ***
						strSQLid = "SELECT id FROM aktiviteter ORDER BY id DESC"
						oRec3.open strSQLid, oConn, 3
						if not oRec3.EOF then
						useAktid = oRec3("id")
						end if
						oRec3.close
						
						if len(useAktid) <> 0 then
						useAktid = useAktid
						else
						useAktid = 0
						end if
						
						if jobid <> 0 then
						    
						    if jobid = jobids(j) then
						    rdir_useAktid = useAktid
						    end if
						
						else
						
						if jobids(j) = 0 then
						rdir_useAktid = useAktid
						end if
						
						end if
						
						
						
				else
				useAktid = id
				rdir_useAktid = useAktid 
				end if
				    
				    
				    
				    '*********************************************************'
				    '********* Opdaterer timepriser på aktivitet *************'
				    '*********************************************************'
					
					
					'if func = "dbred" OR jobids(j) = 0 then 
					if request("FM_overfor_tp") = "j" OR func = "dbred" OR jobids(j) = 0 then
					
					medarb_tpris = request("FM_use_medarb_tpris")
					
					Dim intMedArbID 
					Dim b
					
						intMedArbID = Split(medarb_tpris, ", ")
						For b = 0 to Ubound(intMedArbID)
							oConn.execute("DELETE FROM timepriser WHERE jobid = "& jobids(j) &" AND aktid = "& useAktid &" AND medarbid = "& intMedArbID(b) &"")
							
							
						    if len(request("FM_6timepris_"&intMedArbID(b)&"")) <> 0 OR jobids(j) = 0 then
							   
							   
							       if len(request("FM_6timepris_"&intMedArbID(b)&"")) <> 0 then
							       this6timepris = SQLBlessP(replace(request("FM_6timepris_"&intMedArbID(b)&""), ".", ""))
							       else
							       this6timepris = 0
							       end if
							       
							        valuta6 = request("FM_valuta_600"& intMedArbID(b))
								    tprisalt = 6
								    
								    '** Skal altid være = 6 med mindre der nedarves fra job **'
								    
								    'tprisalt = request("FM_timepris_"&intMedArbID(b)&"")
								    'if tprisalt <> 0 then
								    'tprisalt = tprisalt
								    'else
								    'tprisalt = 6
								    'end if
								    
								
								
								    strSQLtp = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) "_
							        &" VALUES ("& jobids(j) &", "& useAktid &", "& intMedArbID(b) &", "& tprisalt &", "_
							        &" "& this6timepris &", "& valuta6 &")"
							        
							        'Response.Write strSQLtp
							        'Response.flush
							        
							        oConn.execute (strSQLtp)
        							
							        nyTimePris = this6timepris
        					        nyValuta = valuta6
        					
        					else
        					       '*** Nedarver fra job ***'    
        					        strSQLtimp = "SELECT 6timepris, 6valuta FROM timepriser WHERE jobid = "& jobids(j) &" AND medarbid = "& intMedArbID(b) &" AND aktid = 0"
									
										oRec2.open strSQLtimp, oConn, 3
										if not oRec2.EOF then
										nyTimePris = oRec2("6timepris")
										nyValuta = oRec2("6valuta")
										end if
										oRec2.close 
        					        
        					end if
							        
							        
							            fraDato = request("FM_opdatertpfra_individuel") 
                                        if isDate(fraDato) = false then
                                        fraDato = year(now) &"/"& month(now) & "/"& day(now)
                                        else
                                        fraDato = year(fraDato) &"/"& month(fraDato) & "/"& day(fraDato) 
                                        end if

                                        if request("FM_opdater_timepriser_individuel") = "1" then
							        
							                    if len(trim(nyTimePris)) <> 0 AND func = "dbred" then
							                    '************************************************************'
							                    '** Opdaterer timereg (Uanset om der foreligger fakturaer) **'
						                        '************************************************************'
						                        nyTimePris = replace(nyTimePris, ",", ".")
						                        nyValuta = nyValuta
						                    
						                               '**** Finder aktuel kurs ***'
						                               strSQL = "SELECT kurs FROM valutaer WHERE id = " & nyValuta
						                               oRec.open strSQL, oConn, 3
        						            
						                               if not oRec.EOF then
						                               nyKurs = replace(oRec("kurs"), ",", ".")
						                               end if 
						                               oRec.close
						            
						                        strSQL = "UPDATE timer SET timepris = "& nyTimePris &", valuta = "& nyValuta &", kurs = "& nyKurs &""_
						                        &" WHERE tmnr = "& intMedArbID(b) &" AND taktivitetid = "& useAktid   
        					       
					                            'Response.Write strSQL & "<br>"
					                            'Response.flush
					                           
					                            oConn.execute(strSQL)
							
							                    end if
							            
                                        end if
							
						                 
						    next
					
					         'Response.end
					end if 'overfor_tp
					 'end if  '** func dbred & jobids(0) **'
			        '**************************************
				   
                   
                   '********************************'
                    '***** Forretningsområder ******'
                    '********************************'
                    
                    

                    '*** nulstiller akt ****'
                    strSQLfor = "DELETE FROM fomr_rel WHERE for_aktid = "& useAktid
                    oConn.execute(strSQLfor)
                   

                    'Response.Write "her"
                    'Response.Flush

                    for afor = 0 to UBOUND(fomrArr)

                            'Response.Write "her2" & afor & "<br>"
                            'Response.Flush

                            if fomrArr(afor) <> 0 then

                            strSQLfomri = "INSERT INTO fomr_rel "_
                            &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                            &" VALUES ("& fomrArr(afor) &", "& jobid &", "& useAktid &", "& for_faktor &")"

                            'Response.Write strSQLfomri
                            oConn.execute(strSQLfomri)

                           end if


                    next

                    '********************************' 
				    

                                        '*************************************************************'
                                        '*** Opdaterer alle timeregistreringer vedbrug af fastpris ***'
                                        '*************************************************************'

                                        if func = "dbred" then
                                        '** Fra dato ved opdater timepriser ***'
                                        fraDato = request("FM_opdatertpfra") 
                                        if isDate(fraDato) = false then
                                        fraDato = year(now) &"/"& month(now) & "/"& day(now)
                                        else
                                        fraDato = year(fraDato) &"/"& month(fraDato) & "/"& day(fraDato) 
                                        end if

                                        if request("FM_opdater_timepriser") = "1" then

										   '**** Finder aktuel kurs ***'
						                   strSQL = "SELECT kurs FROM valutaer WHERE id = " & fasttp_val
						                   oRec5.open strSQL, oConn, 3
        						            
						                   if not oRec5.EOF then
						                   nyKurs = replace(oRec5("kurs"), ",", ".")
						                   end if 
						                   oRec5.close

                                     

			                            strSQLtp = "UPDATE timer SET timepris = "& fasttp &", valuta = "& fasttp_val &", kurs = "& nyKurs &""_
			                            &" WHERE tdato >= '"& fraDato &"' AND taktivitetid = "& id
					                  
					                    oConn.execute(strSQLtp)
                                        'Response.flush
                                        'Response.write strSQLtp
                                        'Response.end

                                        end if

                                        '**Fra dato opdater kostpriser **'
                                        if request("FM_opdater_kostpriser") = "1" then

								        strSQLkp = "UPDATE timer SET kostpris = "& fastkp &" WHERE tdato >= '"& fraDato &"' AND taktivitetid = "& id
					                    oConn.execute(strSQLkp)
                                        
                                        'Response.write strSQLkp


                                        end if
                                        end if 'dbred
				   
				    
					
					end if 'jobidswrt
					next 'jobid(j)
					
					
					'Response.end
						
						
						if jobid <> 0 then 'Bliver der oprettet en stam akt eller en alm akt.
							select case rdir
							case "jp" '* fra jobplanner
							Response.redirect "jobplanner.asp?menu=job&id="&jobid&"&jaid="&id
							case "webbl21"
							Response.Write("<script language=""JavaScript"">window.opener.location.href('webblik_joblisten21.asp');</script>")
                            Response.Write("<script language=""JavaScript"">window.close();</script>")
							case "treg2"
							Response.Write("<script language=""JavaScript"">window.opener.location.href('timereg_akt_2006.asp?showakt=1');</script>")
                            Response.Write("<script language=""JavaScript"">window.close();</script>")
							case "job2"
							Response.Write("<script language=""JavaScript"">window.opener.location.href('jobs.asp?menu=job&func=red&id="& jobid &"&int=1&rdir=job');</script>")
                            Response.Write("<script language=""JavaScript"">window.close();</script>")
                            case "job3"
                            'Response.Write "hej"
                            'Response.end 
							Response.Write("<script language=""JavaScript"">window.opener.location.href('jobs.asp?menu=job&func=red&id="& jobid &"&int=1&rdir=job&showdiv=forkalk');</script>")
                            Response.Write("<script language=""JavaScript"">window.close();</script>")
                           
							
							'case "treg"
							'Response.Redirect "timereg_2006_fs.asp"
							case "job"
							Response.Redirect "jobs.asp?menu=job&func=red&id="& jobid &"&int=1&rdir=job"
							case else  '** også "treg"
							Response.redirect "aktiv.asp?menu=job&id="&rdir_useAktid&"&jobid="&jobid&"&jobnavn="&request("jobnavn")&"&rdir="&rdir&"&nomenu="&nomenu
							end select 
						else 'hvis akt bliver oprettet som stamakt.
						Response.redirect "aktiv.asp?menu=job&stamaktid="&rdir_useAktid&"&func=favorit&id="&rdir_int_aktFav&"&stamakgrpnavn="&request("jobnavn")&"&nomenu="&nomenu
						end if 
						
					 'end if '*validering
					   end if '*validering
					 end if '*validering
				   end if '*validering
                   end if '*validering
                   end if '*validering
				end if
			end if
	
	
	'***********************************************************************************************	
	case "opret", "red", "opretstam", "redstam"
	
	
	
	'*** Her indlæses form til rediger/oprettelse af Aktivitet ***'
	if func = "opret" OR func = "opretstam" then
	intFomr = 0
	strNavn = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny aktivitet"
	dbfunc = "dbopr"
	
	if len(request("fb")) <> 0 then
	strFakturerbart = request("fb") 
	else
	strFakturerbart = 1
	end if
	
	intaktstatus = 1
	intAktfavorit = 0
	dblFaktor = 1
	tidslaas = 0
	
	
	    tidslaas_st_Dis = "disabled"
		tidslaas_sl_Dis = "disabled"
		tidslaas_man_Dis = "disabled"
		tidslaas_tir_Dis = "disabled"
		tidslaas_ons_Dis = "disabled"
		tidslaas_tor_Dis = "disabled"
		tidslaas_fre_Dis = "disabled"
		tidslaas_lor_Dis = "disabled"
		tidslaas_son_Dis = "disabled"
		
	tidslaas_st = "07:30:00"
	tidslaas_sl = "23:30:00"
	
	    tidslaas_man = 1
		tidslaas_tir = 1
		tidslaas_ons = 1
		tidslaas_tor = 1
		tidslaas_fre = 1
		tidslaas_lor = 0
		tidslaas_son = 0
	
		if func <> "opretstam" then
			call budgettimerTildelt(strjobnr)
			
			call skaljobSync(strjobnr)
			
		else
		totaltimertildelt = 0
		usejoborakt_tp = 0
	    fastpris = 0
		end if
		
    antalstk = 0 '1
    
    strFase = ""
    
    select case lto
    case "dencker" '* Antal stk.
    intBgr = 2
    case "xx" '* Intet valgt
    intBgr = 0
    case else '** Timer
    intBgr = 1
    end select
    
    select case lto
    case "dencker"
    easyreg = 0 '1
    case else
    easyreg = 0
    end select


        strProj_1 = 0
		strProj_2 = 1
		strProj_3 = 1
		strProj_4 = 1
		strProj_5 = 1
		strProj_6 = 1
		strProj_7 = 1
		strProj_8 = 1
		strProj_9 = 1
		strProj_10 = 1

    sortorder = 1000
    

    brug_fasttp = 0
    brug_fastkp = 0
    fasttp = 0
    fastkp = 0
    fasttp_val = 0
    fastkp_val = 0 
    
    strAvarenr = ""
    kostpristarif = "0"

	else
	
	
	skaljobSync(strjobnr)
	
	
	strSQL = "SELECT id, navn, beskrivelse, job, fakturerbar, projektgruppe1, "_
	&" projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, "_
	&" projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, "_
	&" aktstartdato, aktslutdato, editor, dato, budgettimer, aktfavorit, aktstatus, fomr, "_
	&" faktor, aktbudget, tidslaas, tidslaas_st, tidslaas_sl, antalstk, "_
	&" tidslaas_man, tidslaas_tir, tidslaas_ons, "_
	&" tidslaas_tor, tidslaas_fre, tidslaas_lor, tidslaas_son, fase, bgr, easyreg, sortorder, fravalgt, "_
    &" brug_fasttp, fasttp, fasttp_val, brug_fastkp, fastkp, fastkp_val, avarenr, kostpristarif, aktkonto "_
	&" FROM aktiviteter WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
		strNavn = oRec("navn")
		strBeskrivelse = oRec("beskrivelse")
		strjobnr = oRec("job")
		strTdato = oRec("aktstartdato")
		strUdato = oRec("aktslutdato")
		strProj_1 = oRec("projektgruppe1")
		strProj_2 = oRec("projektgruppe2")
		strProj_3 = oRec("projektgruppe3")
		strProj_4 = oRec("projektgruppe4")
		strProj_5 = oRec("projektgruppe5")
		strProj_6 = oRec("projektgruppe6")
		strProj_7 = oRec("projektgruppe7")
		strProj_8 = oRec("projektgruppe8")
		strProj_9 = oRec("projektgruppe9")
		strProj_10 = oRec("projektgruppe10")
		strDato = oRec("dato")
		strLastUptDato = oRec("dato") 
		strEditor = oRec("editor")
		strFakturerbart = oRec("fakturerbar")
		intAktfavorit = oRec("aktfavorit")
		strBudgettimer = oRec("budgettimer")
		intaktstatus = oRec("aktstatus")
		intFomr = oRec("fomr")
		dblFaktor = oRec("faktor")
		intBudget = oRec("aktbudget")
		tidslaas = oRec("tidslaas")
		
		if tidslaas = 1 then
		tidslaas_st_Dis = ""
		tidslaas_sl_Dis = ""
		tidslaas_man_Dis = ""
		tidslaas_tir_Dis = ""
		tidslaas_ons_Dis = ""
		tidslaas_tor_Dis = ""
		tidslaas_fre_Dis = ""
		tidslaas_lor_Dis = ""
		tidslaas_son_Dis = ""
		else
		tidslaas_st_Dis = "disabled"
		tidslaas_sl_Dis = "disabled"
		tidslaas_man_Dis = "disabled"
		tidslaas_tir_Dis = "disabled"
		tidslaas_ons_Dis = "disabled"
		tidslaas_tor_Dis = "disabled"
		tidslaas_fre_Dis = "disabled"
		tidslaas_lor_Dis = "disabled"
		tidslaas_son_Dis = "disabled"
		end if
		
		tidslaas_st = oRec("tidslaas_st")
		tidslaas_sl = oRec("tidslaas_sl")
		
		tidslaas_man = oRec("tidslaas_man")
		tidslaas_tir = oRec("tidslaas_tir")
		tidslaas_ons = oRec("tidslaas_ons")
		tidslaas_tor = oRec("tidslaas_tor")
		tidslaas_fre = oRec("tidslaas_fre")
		tidslaas_lor = oRec("tidslaas_lor")
		tidslaas_son = oRec("tidslaas_son")
		
    	antalstk = oRec("antalstk")
    	
    	strFase = oRec("fase")
    	intBgr = oRec("bgr")
		easyreg = oRec("easyreg")
		
        sortorder = oRec("sortorder")
		
        fravalgt = oRec("fravalgt")


        brug_fasttp = oRec("brug_fasttp") 
        fasttp = oRec("fasttp")
        fasttp_val = oRec("fasttp_val")

        brug_fastkp = oRec("brug_fastkp")
        fastkp = oRec("fastkp")
        fastkp_val = oRec("fastkp_val")

        strAvarenr = oRec("avarenr")
        kostpristarif = oRec("kostpristarif")
        aktkonto = oRec("aktkonto")

		end if
		
		oRec.close
		
		function SQLBless(s)
			dim tmp
			tmp = s
			tmp = replace(tmp, ",", ".")
			SQLBless = tmp
		end function
		
		if func <> "redstam" then
			call budgettimerTildelt(strjobnr)
			
			'** Forkalk på tilhørende aktive aktiviteter ***'
			sumakttimer = 0
			sumaktbudget = 0
			
			
	        strSQLaktSum = "SELECT SUM(budgettimer) sumakttimer, fakturerbar, SUM(aktbudget) AS sumaktbudget FROM aktiviteter "_
			&" WHERE job =  "& strjobnr & " GROUP BY job"
			oRec2.open strSQLaktSum, oConn, 3
			if not oRec2.EOF then
			
			sumakttimer = oRec2("sumakttimer")
			sumaktbudget = oRec2("sumaktbudget")
			
			end if
			oRec2.close
			
		end if	
		
		
		
		dbfunc = "dbred"
		varbroedkrumme = "Rediger aktivitet"
		varSubVal = "opdaterpil" 
	end if
	
	
	'if strFakturerbart = 1 then
	'varFakEks = "checked"
	'else
	'varFakEks = ""
	'end if
		
	%>
	
	
	
	<%if nomenu <> 1 then 
	dtop = 102
	%>
	
	<%if rdir <> "jp" then%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->


            <%if nomenu <> 1 then
            
                call menu_2014()    
                
            end if
    end if

	
	else
	dtop = 20
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<%
	end if
	
	'** Ændrer ID så det passer til Dato2 fil.
	if func = "opret" then
	id = 1
	else
	id = id
	end if 
	%>
	<!--#include file="inc/dato2.asp"-->
	<%
	'** Ændrer ID tilbage til opr ID.
	if func = "opret" then
	id = 0
	nedarvdato = "j"
	else
	id = id
	end if
	%>
	
	
	<script>


	    $(document).ready(function() {

          $("#timepristable").table_checkall();

            if ($("#FM_brug_fasttp").is(':checked') == true) {
                $("#div_akt_med_tp").hide(1000)
            }


            $("#FM_brug_fasttp").click(function () {

                if ($("#FM_brug_fasttp").is(':checked') == true) {
                    $("#div_akt_med_tp").hide(1000)
                } else {
                    $("#div_akt_med_tp").css("display", "")
                    $("#div_akt_med_tp").css("visibility", "visible")
                }

            });
    


	        function beregnakttotpris(id, val) {

           

	            var thisid = id
	            var thisval = val
	            var idlngt = id.length
	            var idtrim = id.slice(7, idlngt)

	            var varstk = $("#FM_stk_" + idtrim).val();
	            var vartimer = $("#FM_tim_" + idtrim).val();
	            var varpris = $("#FM_pri_" + idtrim).val();

	            varpris = varpris.replace(".", "")
	            varstk = varstk.replace(".", "")
	            vartimer = vartimer.replace(".", "")

	            varpris = varpris.replace(",", ".")
	            varstk = varstk.replace(",", ".")
	            vartimer = vartimer.replace(",", ".")

	            //alert(vartimer/1 + " " + varstk/1 + " "+varpris/1)

	            if (thisval == 0) {
	                akttotpris = (varpris / 1)
	            } else {
	                if (thisval == 1) {
	                    akttotpris = (vartimer / 1 * varpris / 1)
	                } else {
	                    akttotpris = (varstk / 1 * varpris / 1)
	                }
	            }

                 akttotpris = String(Math.round(akttotpris * 100) / 100).replace(".", ",")
                 $("#FM_akt_totpris_" + idtrim).val(akttotpris)        
                 $("#akt_totpris_" + idtrim).html("<b>" + akttotpris + "</b>")
	           


	        }


	        $(".bgr").change(function() {
	            var thisid = this.id
	            var thisval = this.value
	            beregnakttotpris(thisid, thisval)
	        });


	        $(".timstkpris").keyup(function() {
	            //alert(this.id)
	            var thisid = this.id
	            var idlngt = thisid.length

	            //alert(idlngt)
	            var idtrim = thisid.slice(7, idlngt)
	            var thisval = $("#FM_bgr_" + idtrim).val()
	            beregnakttotpris(thisid, thisval)
	        });







	    });



	    function vistidslaas() {
	        //alert(document.getElementById("FM_useduedate").checked)
	        if (document.getElementById("FM_tidslaas").checked == true) {
	            document.getElementById("FM_tidslaas_start").disabled = false
	            document.getElementById("FM_tidslaas_slut").disabled = false
	            document.getElementById("FM_tidslaas_man").disabled = false
	            document.getElementById("FM_tidslaas_tir").disabled = false
	            document.getElementById("FM_tidslaas_ons").disabled = false
	            document.getElementById("FM_tidslaas_tor").disabled = false
	            document.getElementById("FM_tidslaas_fre").disabled = false
	            document.getElementById("FM_tidslaas_lor").disabled = false
	            document.getElementById("FM_tidslaas_son").disabled = false

	            document.getElementById("FM_tidslaas_man").checked = true
	            document.getElementById("FM_tidslaas_tir").checked = true
	            document.getElementById("FM_tidslaas_ons").checked = true
	            document.getElementById("FM_tidslaas_tor").checked = true
	            document.getElementById("FM_tidslaas_fre").checked = true

	            document.getElementById("FM_tidslaas_lor").checked = false
	            document.getElementById("FM_tidslaas_son").checked = false


	        } else {
	            document.getElementById("FM_tidslaas_start").disabled = true
	            document.getElementById("FM_tidslaas_slut").disabled = true
	            document.getElementById("FM_tidslaas_man").disabled = true
	            document.getElementById("FM_tidslaas_tir").disabled = true
	            document.getElementById("FM_tidslaas_ons").disabled = true
	            document.getElementById("FM_tidslaas_tor").disabled = true
	            document.getElementById("FM_tidslaas_fre").disabled = true
	            document.getElementById("FM_tidslaas_lor").disabled = true
	            document.getElementById("FM_tidslaas_son").disabled = true
	        }
	    }



	    function expand() {
	        if (document.getElementById("progrp").style.display == "none") {
	            document.getElementById("progrp").style.display = "";
	            document.getElementById("progrp").style.visibility = "visible";
	        } else {
	            document.getElementById("progrp").style.display = "none";
	            document.getElementById("progrp").style.visibility = "hidden";
	        }
	    }

	    function showalttimer() {
	        if (document.getElementById("alttimer").style.visibility == "visible") {
	            document.getElementById("alttimer").style.visibility = "hidden"
	            document.getElementById("alttimer").style.display = "none"
	        } else {
	            document.getElementById("alttimer").style.visibility = "visible"
	            document.getElementById("alttimer").style.display = ""
	        }
	    }



	    function overfortp(mid, t) {
	        document.getElementById("FM_6timepris_" + mid + "").value = document.getElementById("FM_hd_timepris_" + mid + "_" + t).value
	        valuta = document.getElementById("FM_hd_valuta_" + mid + "_" + t).value


	        //alert(document.getElementById("FM_valuta_600"+mid+"").length)

	        for (i = 0; i < document.getElementById("FM_valuta_600" + mid + "").length; i++) {
	            //alert(i +" -- Val: "+ valuta)
	            if (document.getElementById("FM_valuta_600" + mid + "").options[i].value == valuta) {
	                document.getElementById("FM_valuta_600" + mid + "").options[i].selected = true
	            }
	        }


	    }

	    function addfase(fase) {
	        document.getElementById("FM_fase").value = fase
	    }
	
	
      
 </script>


	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:<%=dtop%>px;">

	
	<%if func = "redstam" then
			totaltimertildelt = 1
			totalikkefaktimertildelt = 1
	end if%>
	
	
	<%

    call showEasyreg_fn()

    if nomenu <> 1 then 
	 tTop = 0
	 tLeft = 0
	 tWdth = 800
            	
            	
	 call tableDiv(tTop,tLeft,tWdth)
     
	 else%>
	 <div style="width:800px; border:0px #5582d2 solid; background-color:#FFFFFF;">
     <%end if %>

	<table cellspacing="0" cellpadding="0" border="0" width="100%" bgcolor="#FFFFFF">
	<form action="aktiv.asp?menu=job&func=<%=dbfunc%>" method="post">
	
	<input type="hidden" name="nomenu" value="<%=nomenu%>">
	<input type="hidden" name="rdir" value="<%=rdir%>">
	<input type="hidden" name="id" value="<%=id%>">
	<input type="hidden" name="jobid" value="<%=jobid%>">
	<input type="hidden" name="jobnavn" value="<%=request("jobnavn")%>">
	<input type="hidden" name="FM_OLDaktId" value="<%=id%>">
	<input type="hidden" name="FM_timertildelt" value="<%=totaltimertildelt%>">
	<input type="hidden" name="FM_timertildeltikkeFakbar" value="<%=totalikkefaktimertildelt%>">
	
	<%if dbfunc = "dbred" then%>
	<input type="hidden" name="FM_OLDakttimerTildelt" value="<%=strBudgettimer%>">
	<input type="hidden" name="FM_OLDfakbar" value="<%=strFakturerbart%>">
	<tr bgcolor="#5582D2">
		<td height=30 class="alt" style="padding-left:10px;"><%=varbroedkrumme%></td>
		<td class=alt align=right style="padding-right:20px;">Sidst opdateret den <b><%=strLastUptDato%></b> af <b><%=strEditor%></b></td>
	</tr>
	<%else%>
	<input type="hidden" name="FM_OLDakttimerTildelt" value="0">
	<tr bgcolor="#5582D2">
		<td colspan="2" height=30 class="alt" style="padding-left:10px;"><%=varbroedkrumme%></td>
	</tr>
	<%end if%>
	<tr><td colspan=2 align="right" style="padding:20px 60px 10px 10px;"><input type="submit" value="Opdater >>" /></td></tr>
	<tr>
		<td style="padding:10px 0px 3px 40px;" colspan=2><img src="../ill/akt_24.png" alt="" border="0">&nbsp;<font color=red size=2>*</font> <b>Aktivitets navn:</b>
            <%if func = "red" then %>
            &nbsp;<span style="color:#999999;">(id: <%=id %>)</span>
            <%end if %>
            
            <br />
		<input type="text" name="FM_navn" value="<%=strNavn%>" style="width:450px;" maxlength=100>
        &nbsp;&nbsp;&nbsp;Akt. nr / Varenr:&nbsp;<input type="text" name="FM_avarenr" value="<%=strAvarenr%>" style="width:150px;" maxlength=50><br />
		<span style="color:#999999; font-size:9px;">(maks 100 karak. må ikke indeholde 'aprostrof eller "situations-tegn)</span> </td>
	</tr>
      
        <%if func = "opretstam" OR func = "redstam" then %>
        <tr>
         <%if fravalgt <> 0 then
				  chkFravalgt = "CHECKED"
				  else 
				  chkFravalgt = ""
				  end if%>

		<td colspan=2 style="padding:0px 40px 10px 40px;"> <input id="Checkbox1" name="FM_fravalgt" type="checkbox" value="1" <%=chkFravalgt%> />&nbsp;Fravalgt on/off (ved job-oprettelse, tilføj fra skabelon)
		    
				  
                   </td>
	</tr>
        <%end if %>
         

	<tr>
		<td style="padding:10px 0px 10px 40px;" colspan="2"><b>Status:</b><br /><select name="FM_aktstatus" style="width:200px;">
		<%if dbfunc = "dbred" then 
		select case intaktstatus
		case 1
		strStatusNavn = "Aktiv"
		case 2
		strStatusNavn = "Passiv"
		case 0
		strStatusNavn = "Lukket"
		end select
		%>
		<option value="<%=intaktstatus%>" SELECTED><%=strStatusNavn%></option>
		<%end if%>
	<option value="1">Aktiv</option>
	<option value="2">Passiv</option>
	<option value="0">Lukket</option>
		</select></td>
	</tr>
	<tr>
		<td style="padding:10px 0px 3px 40px;" colspan=2><b>Fase:</b><br />
		
		<%
		strSQLfas = "SELECT fase FROM aktiviteter WHERE job = "& jobid &" AND fase <> '' GROUP BY fase" 
		oRec3.open strSQLfas, oConn, 3
		while not oRec3.EOF 
		%>
		<a href="#" onclick="addfase('<%=replace(oRec3("fase"), "_", " ") %>')" class=rmenu><%=replace(oRec3("fase"), "_", " ") %></a>&nbsp;|&nbsp;
		<%
		oRec3.movenext
		wend
		oRec3.close
		
		if len(trim(strFase)) <> 0 then
		strFase = replace(strFase, "_", " ")
		else
		strFase = strFase
		end if
		 
		
		%>
		<br />
		<input type="text" id="FM_fase" name="FM_fase" value="<%=strFase%>" style="width:700px;" maxlength=100><br />
		<span style="color:#999999; font-size:9px;">Maks 100 karakterer. Må ikke indeholde .punktum, 'aprostrof, &-tegn eller mellemrum.</span></td>
	</tr>
	<%
	    level = session("rettigheder")
	    if level = 1 then 'AND (func = "opret" OR func = "opretstam") then
		whSQL = "jobstatus = 1 OR j.id = "& jobid 
		
            if (func = "opret" OR func = "opretstam") then
            sz = "6"
		    mul = "multiple"
		    hgt = 120
		    jtxt = " (multi-tildel samtidig på andre job)"
            else
            sz = "1"
		    mul = ""
		    hgt = 60
            jtxt = " (flyt aktivitet til job)"
            end if
		else
		whSQL = "j.id = "& jobid 
		sz = "1"
		mul = ""
		hgt = 60
		jtxt = ""
		end if
	
	%>
	<tr>
		<td style="padding:10px 0px 3px 40px;" valign="top" colspan=2><br><b>Kontakt | Job:</b> <%=jtxt %><br />
		<input type="hidden" name="FM_jnr" value="<%=jobid %>">
		<%
		
		
		strSQLjob = "SELECT jobnavn, jobnr, j.id, kkundenavn, kkundenr, count(a.id) AS antalA FROM job j "_
		&" LEFT JOIN kunder ON (kid = jobknr) "_
		&" LEFT JOIN aktiviteter a ON (a.job = j.id) "_
		&" WHERE "& whSQL &" GROUP BY j.id ORDER BY kkundenavn, jobnr"
		
		'Response.Write strSQLjob
		'Response.flush
		
		%>
		
            <select name="FM_jnr" id="FM_jnr" <%=mul%> size="<%=sz%>" style="width:450px;">
            <%if func = "opretstam" OR func = "redstam" then 
            
            'if cint(strjobnr) = 0 then
		    sel0j = "SELECTED"
		    'else
		    'sel0j = ""
		    'end if
            
            %>
            <option value="0" <%=sel0j%>>Stam-aktivitetsgruppe: <%=request("jobnavn") %></option>
            <%end if %>
              
		<%
		
		oRec2.open strSQLjob, oConn, 3
		t = 2
		while not oRec2.EOF
		
		if cint(jobid) = cint(oRec2("id")) AND sel0j <> "SELECTED" then
		selj = "SELECTED"
		else
		selj = ""
		end if%>
		  <option value="<%=oRec2("id") %>" <%=selj %>><%=oRec2("kkundenavn") %> (<%=oRec2("kkundenr") %>) | (<%=oRec2("jobnr") %>) <%=oRec2("jobnavn") %> | Antal atk.: <%=oRec2("antalA") %> </option>
           
		<%
		oRec2.movenext
		Wend
		oRec2.close
		
		%>
		 </select>
		 
		</td>
	</tr>
	<tr>
		<td style="padding:10px 0px 6px 40px;" valign="top" colspan=2><br><b>Beskrivelse:</b><br>
		
		  <%
	                    
	                    dim content
	                    content = strBeskrivelse
            			
			            
			            Set editorK = New CuteEditor
            					
			            editorK.ID = "FM_beskrivelse"
			            editorK.Text = content
			            editorK.FilesPath = "CuteEditor_Files"
			            editorK.AutoConfigure = "Minimal"
            			
			            editorK.Width = 700
			            editorK.Height = 280
			            editorK.Draw()
		                %>
		                <br />
            &nbsp;
		
		</td>
	</tr>
	<tr>
		<td valign="top" style="padding:5px 60px 2px 40px;" colspan=2>
		
		<%
		'**** Der kan ikke mere oprettes stamaktiviter via aktiviteter der bliver oprettet på et job ****
		'**** Derfor er der åbnet op for timer på opret stam akt.
		'if func <> "opretstam" AND func <> "redstam" then
		'end if
		'***************************************************
		%>
		
		<table width=100% border=0 cellpadding=4 cellspacing=2 style="background-color:#d6dff5;">
		<tr><td><b>Forkalkuleret timer:</b><br />
		<input type="text" name="FM_budgettimer" class="timstkpris" id="FM_tim_0" value="<%=formatnumber(strBudgettimer, 2)%>" size="5"> timer 
		</td>
		<td><b>Stk. / Antal:</b><br />
		<input type="text" name="FM_antalstk" class="timstkpris" id="FM_stk_0" value="<%=formatnumber(antalstk, 2)%>" size="5">
		</td>
		<td><b>Beregnings grundlag:</b><br />
		<%
		bgrSEL0 = "SELECTED"
	    bgrSEL1 = ""
	    bgrSEL2 = ""
	
	
	select case intBgr
	case 0
	bgrSEL0 = "SELECTED"
	aktsumtot = intBudget
	case 1
	bgrSEL1 = "SELECTED"
	aktsumtot = formatnumber(intBudget * strBudgettimer) 
	case 2
	bgrSEL2 = "SELECTED"
	aktsumtot = formatnumber(intBudget * antalstk)
	end select %>
	
    
        <select class="bgr" id="FM_bgr_0" name="FM_bgr" style="width:65px;">
            <option value=0 <%=bgrSEL0 %>>Ingen</option>
            <option value=1 <%=bgrSEL1 %>>Timer</option>
            <option value=2 <%=bgrSEL2 %>>Stk.</option>
        </select>
        </td>
		<td><b>Pris pr. stk. / time</b><br />
		<input type="text" class="timstkpris" id="FM_pri_0" name="FM_budget" value="<%=formatnumber(intBudget, 2)%>" size="10"> <%=basisValISO %>
		</td>
		<td style="width:100px;"">Pris ialt:<br />
		= <label id="akt_totpris_0" style="width:80px; top:5px;"><b><%=formatnumber(aktsumtot, 2) %></b></label> <%=basisValISO %>
          <input id="FM_akt_totpris_0" name="FM_akt_totpris" type="hidden" value="<%=formatnumber(aktsumtot, 2) %>" />
            &nbsp;
		</td></tr>
		<tr>
		<td colspan=5>
		<%if cint(SY_fastpris) = 1 AND _
        cint(SY_usejoborakt_tp) = 1 then
        opdjobvCHK = "CHECKED"
        else
        opdjobvCHK = ""
        end if
        %>
		     <br /><input id="opdjobv" name="opdjobv" value="1" type="checkbox" <%=opdjobvCHK %> /> <b>Sync. job</b>  (timer og omsætning, hvis job er fastpris og aktiviteterne er bestemmende for timepris) 
		</td>
            
		</tr></table>
    
    
          
		</td>
	</tr>
	<tr>
	    <td style="padding:10px 0px 5px 40px;" colspan=2><b>Faktor:</b>&nbsp;<br />
             
		<input type="text" name="FM_faktor" value="<%=formatnumber(dblfaktor, 2)%>" size="5"> (timer * faktor = enheder) <span style="color:#999999; font-size:10px;">Opdaterer ikke eksisterende timer på denne aktivitet.</span>
    </td>
	</tr>
	
	
	
    <tr>
		<td style="padding:40px 0px 10px 40px;"><b>Sortering:</b></td>
		<td style="padding:40px 0px 10px 0px;"><input id="Text2" name="SortOrder" value="<%=sortorder%>" style="width:60px;" type="text" /> (heltal)</td>
	</tr>
               
	

			<tr>
				<td valign=top style="padding:40px 0px 0px 40px; width:150px;"><b>Type:</b> (egenskab)
				
				
                    
              </td>
				<td style="padding:40px 0px 0px 0px;">
				        <select id="FM_fakturerbart" name="FM_fakturerbart" style="width:400px;">
                       <%call akttyper2009(1)
                       Response.Write aty_options
                       %>
                  
                    </select>
                    <%if level = 1 then %>
                    <br />
                    <font class=megetlillesort>
                    M) Medregnes i det daglige tidsforbrug<br />
                    E) Med på faktura<br />
                    Z) Fakturerbar<br />
                    F) Fradrag i løntimer<br />
                    S) Skjult på timereg. siden
                   </font>
                   <%end if %>
                   <br />
                    &nbsp;
                    
                </td>
			  </tr>
			

            
                  
    
    
    <%if cint(showEasyreg_val) = 1 then  %>
    <tr>
      <%if easyreg <> 0 then
				  chkEasyreg = "CHECKED"
				  else 
				  chkEasyreg = ""
				  end if%>
				  

		<td colspan=2 style="padding:5px 0px 10px 40px;">
		
                    <input id="FM_easyreg" name="FM_easyreg" type="checkbox" value="1" <%=chkEasyreg%> />&nbsp;Tilføj til Easyreg. (Indtastning af timer fordelt på flere job samtidig)</td>
	</tr>
    <%end if %>

   

    <%
    
                if func = "red" OR func = "redstam" then
                                    
                    strFomr_rel = ""
                    strSQLfrel = "SELECT for_fomr, fomr.navn FROM fomr_rel "_
                    &" LEFT JOIN fomr ON (fomr.id = for_fomr) WHERE for_aktid = "& id

                    'Response.Write strSQLfrel
                    'Response.flush
                    f = 0
                    oRec.open strSQLfrel, oConn, 3
                    while not oRec.EOF

                    strFomr_rel = strFomr_rel & "#"& oRec("for_fomr") &"#"
                    
                    f = f + 1
                    oRec.movenext
                    wend
                    oRec.close

                   

                else

                strFomr_rel = ""

                end if
                

    %>
	
	<tr>
		<td style="padding-left:40px;" valign=top><br /><b>Forretningsområde(r):</b><br />
        <span style="color:#999999;">Hvis forretningsområde ændres, sker dette med tilbagevirkende kraft. Hvis der vælges flere områder, tæller timer med i begge områder 100%</span></td>
		<td valign=top><br /><select name="FM_fomr" size="8" multiple="multiple" style="width:400px;">
		<option value="0">Ingen valgt</option>
		<%
		strSQL = "SELECT id, navn FROM fomr WHERE aktok = 1 ORDER BY navn"
		oRec.open strSQL, oConn, 3 
		while not oRec.EOF 
		
        if instr(strFomr_rel, "#"&oRec("id")&"#") <> 0 then
        fSel = "SELECTED"
        else
        fSel = ""
        end if
		%>
		<option value="<%=oRec("id")%>" <%=fSel%>><%=oRec("navn")%></option>
		<%
		oRec.movenext
		wend
		oRec.close 
		%>
		</select><br>
		<br>&nbsp;</td>
	</tr>
	<%

        call fomr_account_fn()

       if cint(fomr_account) = 1 then 'kontonr på aktivitet %>

        <tr>
		<td style="padding-left:40px;" valign=top><b>Konto:</b></td><td>


       <%if func = "opretstam" OR func = "redstam" then 
           
           aktkontoDis = ""
           aktkontoName = "FM_aktkonto" 
           
         else

           aktkontoDis = "Disabled"
           aktkontoName = "FM_aktkontoDis"

           %>
             <input type="hidden" value="<%=aktkonto%>" name="FM_aktkonto" />
            <%

        end if
           
           %>

            <select name="<%=aktkontoName %>" <%=aktkontoDis%> style="width:400px;">
              <option value="0">Ingen</option>

        <%
            licensindehaverKid = 0
            call licKid()

         strSQLkonto = "SELECT navn, kontonr, id FROM kontoplan WHERE kid = "& licensindehaverKid
         oRec2.open strSQLkonto, oConn, 3
         while not oRec2.EOF 
            
            
            
            if cint(aktkonto) = oRec2("id") then
            kontoSEL = "SELECTED"
            else
            kontoSEL = ""
            end if

           
            kontonrVal = " ("& oRec2("kontonr") &")"
            
         
            %><option value="<%=oRec2("id") %>" <%=kontoSEL%>><%=oRec2("navn") & kontonrVal%></option><%

         oRec2.movenext
         wend
         oRec2.close  %>

          

    </select>

        </td>
        </tr>
      
	
	
	<%end if 'aktkonto = 1


	'**** Der kan ikke mere oprettes stamaktiviter via aktiviteter der bliver oprettet på et job ****
		'**** Derfor er der åbnet op for timer på opret stam akt.
		
		if func <> "opretstam" AND func <> "redstam" then
		%>
		<input type="hidden" name="FM_favorit" id="FM_favorit" value="0">
		<%
		end if%>
		
			

    <%
    select case lto
    case "xintranet - local", "lw"
    showT = 1
    case else
    showT = 0
    end select
    
    
    if showT = 1 then %>

			<%if len(tidslaas) <> 0 then
			tidslaas = tidslaas
			else
			tidslaas = 0
			end if
			
			if tidslaas = 1 then
			tidslaas_chk = "CHECKED"
			else
			tidslaas_chk = ""
			end if %>
	
	<tr>
			<td style="padding:40px 0px 0px 40px;" valign=top>
				<input type="checkbox" name="FM_tidslaas" id="FM_tidslaas" value="1" <%=tidslaas_chk%> onclick="vistidslaas();"><b>Tidslås:</b> <br />
				Der skal <b>kun</b> kunne registreres timer<br /> på denne aktivitet i det anførte tidsinterval.
				</td>
			<td style="padding:40px 0px 0px 0px;" valign=top>
			<%
			
			
			if len(trim(tidslaas_st)) <> 0 AND len(trim(tidslaas_sl)) <> 0 then
			tidslaas_st = formatdatetime(tidslaas_st, 3)
			tidslaas_sl = formatdatetime(tidslaas_sl, 3)
			else
			tidslaas_st = "07:00:00"
			tidslaas_sl = "23:30:00"
			end if
			
			if tidslaas_man <> 0 then
			tidslaas_manChk = "CHECKED"
			else
			tidslaas_manChk = ""
			end if
			
			if tidslaas_tir <> 0 then
			tidslaas_tirChk = "CHECKED"
			else
			tidslaas_tirChk = ""
			end if
			
			if tidslaas_ons <> 0 then
			tidslaas_onsChk = "CHECKED"
			else
			tidslaas_onsChk = ""
			end if
			
			if tidslaas_tor <> 0 then
			tidslaas_torChk = "CHECKED"
			else
			tidslaas_torChk = ""
			end if
			
			if tidslaas_fre <> 0 then
			tidslaas_freChk = "CHECKED"
			else
			tidslaas_freChk = ""
			end if
			
			if tidslaas_lor <> 0 then
			tidslaas_lorChk = "CHECKED"
			else
			tidslaas_lorChk = ""
			end if
			
			if tidslaas_son <> 0 then
			tidslaas_sonChk = "CHECKED"
			else
			tidslaas_sonChk = ""
			end if
			
			
			%>
			
			Fra kl. <input type="text" name="FM_tidslaas_start" id="FM_tidslaas_start" size="5" value="<%=tidslaas_st%>" <%=tidslaas_st_Dis %>> til kl.
			<input type="text" name="FM_tidslaas_slut" id="FM_tidslaas_slut" size="5" value="<%=tidslaas_sl%>" <%=tidslaas_sl_Dis %>> (tt:mm:ss, 24:00:00 angives som 23:59:59)
			
			 <br />På følgende dage:&nbsp;
                         Man 
                        <input id="FM_tidslaas_man" name="FM_tidslaas_man" value="1" <%=tidslaas_manChk %> type="checkbox" <%=tidslaas_man_Dis %>/>&nbsp;&nbsp;
                         Tir
                        <input id="FM_tidslaas_tir" name="FM_tidslaas_tir" value="1" <%=tidslaas_tirChk %> type="checkbox" <%=tidslaas_tir_Dis %>/>&nbsp;&nbsp;
                         Ons 
                        <input id="FM_tidslaas_ons" name="FM_tidslaas_ons" value="1" <%=tidslaas_onsChk %> type="checkbox"  <%=tidslaas_ons_Dis %>/>&nbsp;&nbsp;
                         Tor 
                        <input id="FM_tidslaas_tor" name="FM_tidslaas_tor" value="1" <%=tidslaas_torChk %> type="checkbox" <%=tidslaas_tor_Dis %>/>&nbsp;&nbsp;
                         Fre 
                        <input id="FM_tidslaas_fre" name="FM_tidslaas_fre" value="1" <%=tidslaas_freChk %> type="checkbox" <%=tidslaas_fre_Dis %>/>&nbsp;&nbsp;
                         Lør 
                        <input id="FM_tidslaas_lor" name="FM_tidslaas_lor" value="1" <%=tidslaas_lorChk %> type="checkbox" <%=tidslaas_lor_Dis %>/>&nbsp;&nbsp;
                         Søn 
                        <input id="FM_tidslaas_son" name="FM_tidslaas_son" value="1" <%=tidslaas_sonChk %> type="checkbox" <%=tidslaas_son_Dis %>/>&nbsp;&nbsp;
                       <br />&nbsp;
			</td>
			</tr>

            <%end if %>
	
	<%if func <> "opretstam" AND func <> "redstam" then%>
			<tr>
				<td colspan="2" style="padding:40px 0px 0px 40px;"><b>Start og slut dato:</b>&nbsp;&nbsp;
					<%if func = "opret" then
					Response.write "(nedarves fra job)"
					end if%>
				</td>
			</tr>
			<tr>
				<td align=right style="padding:0px 0px 10px 40px;">Start dato:&nbsp;</td><td style="padding:0px 0px 10px 0px;"><select name="FM_start_dag">
				<option value="<%=strDag%>"><%=strDag%></option> 
				<option value="1">1</option>
			   	<option value="2">2</option>
			   	<option value="3">3</option>
			   	<option value="4">4</option>
			   	<option value="5">5</option>
			   	<option value="6">6</option>
			   	<option value="7">7</option>
			   	<option value="8">8</option>
			   	<option value="9">9</option>
			   	<option value="10">10</option>
			   	<option value="11">11</option>
			   	<option value="12">12</option>
			   	<option value="13">13</option>
			   	<option value="14">14</option>
			   	<option value="15">15</option>
			   	<option value="16">16</option>
			   	<option value="17">17</option>
			   	<option value="18">18</option>
			   	<option value="19">19</option>
			   	<option value="20">20</option>
			   	<option value="21">21</option>
			   	<option value="22">22</option>
			   	<option value="23">23</option>
			   	<option value="24">24</option>
			   	<option value="25">25</option>
			   	<option value="26">26</option>
			   	<option value="27">27</option>
			   	<option value="28">28</option>
			   	<option value="29">29</option>
			   	<option value="30">30</option>
				<option value="31">31</option></select>
				
				<select name="FM_start_mrd">
				<option value="<%=strMrd%>"><%=strMrdNavn%></option>
				<option value="1">jan</option>
			   	<option value="2">feb</option>
			   	<option value="3">mar</option>
			   	<option value="4">apr</option>
			   	<option value="5">maj</option>
			   	<option value="6">jun</option>
			   	<option value="7">jul</option>
			   	<option value="8">aug</option>
			   	<option value="9">sep</option>
			   	<option value="10">okt</option>
			   	<option value="11">nov</option>
			   	<option value="12">dec</option></select>
				
				
				<select name="FM_start_aar">
				<option value="<%=strAar%>">
				<%if id <> 0 OR nedarvdato = "j" then%>
				20<%=strAar%>
				<%else%>
				<%=strAar%>
				<%end if%></option>
				
			   	<%for x = -5 to 10
                useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
                  <option value="<%=right(useY, 2)%>"><%=useY%></option>
                <%next %>
				</select>
				&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=6')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a></td>
										
				
				</tr>
				<tr>
				<td valign=top align=right style="padding:5px 0px 10px 40px;">Slut dato:&nbsp;</td><td valign=top style="padding:0px 0px 10px 0px;"><select name="FM_slut_dag">
				<option value="<%=strDag_slut%>"><%=strDag_slut%></option> 
			   	<option value="1">1</option>
			   	<option value="2">2</option>
			   	<option value="3">3</option>
			   	<option value="4">4</option>
			   	<option value="5">5</option>
			   	<option value="6">6</option>
			   	<option value="7">7</option>
			   	<option value="8">8</option>
			   	<option value="9">9</option>
			   	<option value="10">10</option>
			   	<option value="11">11</option>
			   	<option value="12">12</option>
			   	<option value="13">13</option>
			   	<option value="14">14</option>
			   	<option value="15">15</option>
			   	<option value="16">16</option>
			   	<option value="17">17</option>
			   	<option value="18">18</option>
			   	<option value="19">19</option>
			   	<option value="20">20</option>
			   	<option value="21">21</option>
			   	<option value="22">22</option>
			   	<option value="23">23</option>
			   	<option value="24">24</option>
			   	<option value="25">25</option>
			   	<option value="26">26</option>
			   	<option value="27">27</option>
			   	<option value="28">28</option>
			   	<option value="29">29</option>
			   	<option value="30">30</option>
				<option value="31">31</option></select>
				
				<select name="FM_slut_mrd">
				<option value="<%=strMrd_slut%>"><%=strMrdNavn_slut%></option>
				<option value="1">jan</option>
			   	<option value="2">feb</option>
			   	<option value="3">mar</option>
			   	<option value="4">apr</option>
			   	<option value="5">maj</option>
			   	<option value="6">jun</option>
			   	<option value="7">jul</option>
			   	<option value="8">aug</option>
			   	<option value="9">sep</option>
			   	<option value="10">okt</option>
			   	<option value="11">nov</option>
			   	<option value="12">dec</option></select>
				
				
				<select name="FM_slut_aar">
				<option value="<%=strAar_slut%>">
				<%if id <> 0  OR nedarvdato = "j"  then%>
				20<%=strAar_slut%>
				<%else%>
				<%=strAar_slut%>
				<%end if%></option>
				
				
				
			   <%for x = -5 to 10
                useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
                  <option value="<%=right(useY, 2)%>"><%=useY%></option>
                <%next %>
				</select>
				
				&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=5')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a><br />&nbsp;&nbsp;
												
				
				</td>
			</tr>
		
		 <%end if %>
			
		
         
        		
			
				<tr>
					<td colspan="2" style="padding:5px 0px 5px 35px;">

					<br><br><b>Projektgrupper:</b><br>
                    <% if func <> "opretstam" AND func <> "redstam" then 
                    
                    
                        if func = "opret" then
                        neCHK = "CHECKED"
                        onclkfn = "expand()"
                        else
                        neCHK = ""
                        onclkfn = ""
                        end if%>
					<input type="checkbox" name="FM_pgrp_arvefode" value="0" <%=neCHK %> onclick="<%=onclkfn %>"> Nedarv fra job. (de projektgrupper der har adgang til jobbet skal også have adgang til denne aktivitet.)<br>
                    <%end if %>
					</td>
			    </tr>
              

			
				
				<%if func = "opretstam" OR func = "redstam" OR func = "red" then
				disp = ""
                vzb = "visible"
                else
                disp = "none"
                vzb = "hidden"
                end if
				%>
				
		
			
			<tr>
			<td colspan="2" style="padding-left:40px;">
                  

					<div id="progrp" style="position: relative; display:<%=disp%>; visibility:<%=vzb%>;">

                      

					<table cellpadding="2" cellspacing="1" border="0">
                    <tr>
						<td valign="top" style="padding-left:35;" colspan="2"><b>Angiv de projektgrupper der skal have adgang til denne aktivitet.</b><br />
                    <%if func = "opret" OR func = "red" then %>
                    Vælg mellem de projektgrupper der er tilføjet til jobbet.<br />&nbsp;
                    <%end if %></td>
                    </tr>
					
					
					
					<%
					p = 0
					for p = 1 to 10
					varSelected = ""
					select case p
					case 1
					strProj = strProj_1
					case 2
					strProj = strProj_2
					case 3
					strProj = strProj_3
					case 4
					strProj = strProj_4
					case 5
					strProj = strProj_5
					case 6
					strProj = strProj_6
					case 7
					strProj = strProj_7
					case 8
					strProj = strProj_8
					case 9
					strProj = strProj_9
					case 10
					strProj = strProj_10
					end select
					%>
					<tr>
						<td valign="top" style="padding-left:35;">Projektgruppe <%=p%>:</td>
						<td>
						<%
							
                            select case func
                            'case "redstam"
                            'strSQL = "SELECT projektgrupper.id AS id, projektgrupper.navn AS navn, a.projektgruppe1, a.projektgruppe2, a.projektgruppe3, a.projektgruppe4, a.projektgruppe5, "_
                            '&" a.projektgruppe6, a.projektgruppe7, a.projektgruppe8, a.projektgruppe9, a.projektgruppe10 "_
                            '&" FROM aktiviteter a LEFT JOIN projektgrupper ON (projektgrupper.id = a.projektgruppe1 OR projektgrupper.id = a.projektgruppe2 "_
                            '&" OR projektgrupper.id = a.projektgruppe3 OR projektgrupper.id = a.projektgruppe4 OR projektgrupper.id = a.projektgruppe5 "_
                            '&" OR projektgrupper.id = a.projektgruppe6 OR projektgrupper.id = a.projektgruppe7 OR projektgrupper.id = a.projektgruppe8 OR "_
                            '&" projektgrupper.id = a.projektgruppe9 OR projektgrupper.id = a.projektgruppe10) WHERE a.id = "& id  
                            case "red", "opret" 
                            strSQL = "SELECT projektgrupper.id AS id, projektgrupper.navn AS navn, job.projektgruppe1, job.projektgruppe2, job.projektgruppe3, job.projektgruppe4, job.projektgruppe5, "_
                            &" job.projektgruppe6, job.projektgruppe7, job.projektgruppe8, job.projektgruppe9, job.projektgruppe10 "_
                            &" FROM job LEFT JOIN projektgrupper ON (projektgrupper.id = job.projektgruppe1 OR projektgrupper.id = job.projektgruppe2 "_
                            &" OR projektgrupper.id = job.projektgruppe3 OR projektgrupper.id = job.projektgruppe4 OR projektgrupper.id = job.projektgruppe5 "_
                            &" OR projektgrupper.id = job.projektgruppe6 OR projektgrupper.id = job.projektgruppe7 OR projektgrupper.id = job.projektgruppe8 OR "_
                            &" projektgrupper.id = job.projektgruppe9 OR projektgrupper.id = job.projektgruppe10) WHERE job.id = "& strjobnr 
                            ' ORDER BY projektgrupper.id"
                            case else
                            strSQL = "SELECT navn, id FROM projektgrupper WHERE id <> 0 ORDER BY navn"

                            end select
							
                            'Response.write strSQL
                            'Response.flush

                            %>
                            <select name="FM_projektgruppe_<%=p%>">
                            <%

							
							oRec.open strSQL, oConn, 3
							While not oRec.EOF 
							projgId = oRec("id")
							projgNavn = oRec("navn")
							
							if cint(strProj) = cint(projgId) then
							varSelected = "SELECTED"
							gp = strProj
							else
							varSelected = ""
							gp = gp
							end if
							%>
							<option value="<%=projgId%>" <%=varSelected%>><%=projgNavn%> </option>
							<%
							oRec.movenext
							wend
							oRec.close
							%>
				</select>
				</td>
					</tr>
					<%
					select case p
					case 1
					gp1 = gp
					case 2
					gp2 = gp
					case 3
					gp3 = gp
					case 4
					gp4 = gp
					case 5
					gp5 = gp
					case 6
					gp6 = gp
					case 7
					gp7 = gp
					case 8
					gp8 = gp
					case 9
					gp9 = gp
					case 10
					gp10 = gp
					end select
					
					next%>
					
			</table>
			</div>


            
            <%if func = "opretstam" then %>
            <%


            uWdt = 350
            uTop = 1465
            uLeft = 420
            uTxt = "<b>Ved multi-tildel Stam-aktivitet på eksisterende job:</b> (gælder kun nu mens stam-aktiviteten oprettes)<br />"_
            &"<input name=""FM_overfor_tp"" id=""FM_overfor_tp"" value=""j"" type=""checkbox"" /><b>Brug valgte timepriser</b> på denne stam-aktivitet. (ellers nedarves timepriser fra de job aktiviteten overføres til)<br />"_
            &"<br><input type=""radio"" name=""FM_pgrp_arvefode"" value=""1""><b>Brug valgte projektgrupper.</b> Hvis projektgruppen ikke findes på jobbet, oprettes den (hvis ikke alle 10 projektgrupper på jobbet allerede er udfyldt). <br />"_
            &"<br><input type=""radio"" name=""FM_pgrp_arvefode"" value=""0"" CHECKED><b>Nedarv fra job</b> (projektgrupper nedarves automatisk fra de job aktiviteten tilføjes til)"
            
            call infoUnisportAB(uWdt, uTxt, uTop, uLeft) %>
            
		    <%end if %>
           

		</td>
		</tr>
	
    
    <%
    '** ved oprettese af stamakt. ****
    if func = "opretstam" OR func = "redstam" then
    %>
    
   
    <input type="hidden" name="FM_start_dag" value="1">
	<input type="hidden" name="FM_start_mrd" value="1">
	<input type="hidden" name="FM_start_aar" value="2001">
	<input type="hidden" name="FM_slut_dag" value="1">
	<input type="hidden" name="FM_slut_mrd" value="1">
	<input type="hidden" name="FM_slut_aar" value="2044">
	<!--<input type="hidden" name="FM_jnr" value="0">-->
	
    <input type="hidden" name="FM_favorit" value="<%=request("aktfavgp")%>">
	<!--
    <input type="hidden" name="FM_projektgruppe_1" value="10">
	<input type="hidden" name="FM_projektgruppe_2" value="1">
	<input type="hidden" name="FM_projektgruppe_3" value="1">
	<input type="hidden" name="FM_projektgruppe_4" value="1">
	<input type="hidden" name="FM_projektgruppe_5" value="1">
	<input type="hidden" name="FM_projektgruppe_6" value="1">
	<input type="hidden" name="FM_projektgruppe_7" value="1">
	<input type="hidden" name="FM_projektgruppe_8" value="1">
	<input type="hidden" name="FM_projektgruppe_9" value="1">
	<input type="hidden" name="FM_projektgruppe_10" value="1">
    -->
	
    <%end if '*opretstam%>

	
	<%if func = "red" OR func = "opretstam" OR func = "redstam" then%>
	<tr>
		<td style="padding-left:40px; padding-top:20px; padding-bottom:20px; padding-right:20px;" colspan=2>

        <%if showtp = 1 then
        showtpVal = 0
        else
        showtpVal = 1
        end if %>

     <%
    'Response.flush
	

	'dsp = ""
	'vz = "visible"
	


    if cint(brug_fasttp) = 1 then
    brug_fasttpCHK = "CHECKED"
    else
    brug_fasttpCHK = ""
    end if


     if cint(brug_fastkp) = 1 then
    brug_fastkpCHK = "CHECKED"
    else
    brug_fastkpCHK = ""
    end if
    


	%>
	<div style="position:relative; visibility:visible; width:720px; display:;" id="alttimer">

	<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr>
		<td colspan=4><br>
		<h3><img src="../ill/ac0001-24.gif" width="24" height="24" alt="" border="0">&nbsp;Timepriser på aktivitet</h3>
        <br />

        <b>Angiv <u>fast (ens)</u> time- og/eller kost -pris på denne aktivitet for alle medarbejdere:</b><br /><br />
        <div style="border:1px #CCCCCC solid; padding:20px; background-color:#F7f7f7;">

         <input type="checkbox" value="1" name="FM_brug_fastkp" <%=brug_fastkpCHK %> /> Angiv <b>fast kostpris</b> gældende for alle medarbejdere. 
         Kostpris: <input type="text" name="FM_fastkp" value="<%=fastkp %>" style="width:100px;" />  <%=basisValISO %>
         
         <%if func <> "redstam" AND func <> "opretstam" then %><br />
         <input type="checkbox" value="1" name="FM_opdater_kostpriser" /> Opdater kostpriser på eksisterende timer
         <%end if %>


            <%

                selTarif0 = ""
                selTarifA = ""
                selTarifB = ""
                selTarifC = ""
                selTarifD = ""


                select case kostpristarif
                case "0"
                selTarif0 = "SELECTED"
                case "A"
                selTarifA = "SELECTED"
                case "B"
                selTarifB = "SELECTED"
                case "C"
                selTarifC = "SELECTED"
                case "D"
                selTarifD = "SELECTED"
                end select%>
                
                
                 

        <br /><br />
            Benyt kostpris-tarif på denne aktivitet: (fra medarb.type)<br /> <select name="FM_kostpristarif">
                <option value="0" <%=selTarif0 %>>Nej (benyt alm. kostpris)</option>
                <option value="A" <%=selTarifA %>>Tarif A</option>
                <option value="B" <%=selTarifB %>>Tarif B</option>
                <option value="C" <%=selTarifC %>>Tarif C</option>
                <option value="D" <%=selTarifD %>>Tarif D</option>
           </select> <br /><span style="color:#999999;">(beruges f.eks ved natarbejde. Kan ikke benyttes, hvis der er angivet fast-kostpris ovenfor)</span>


         <br /><br>
        
            <%if cint(SY_laasmedtpbudget) = 1 then 'fra job
                
                lstTxt = "<br><span style=""color:#999999;"">Medarbejder timepris er låst til budget på aktiviteter (se job)</span>"
                valIdSQLkri = " id = "& fasttp_val

                if brug_fasttpCHK = "CHECKED" then
                brug_fasttpVal = 1
                else
                brug_fasttpVal = 0
                end if

                %>

             <input type="checkbox" value="1" Disabled name="" id="FM_brug_fasttp" <%=brug_fasttpCHK %> /> Angiv <b>fast timepris</b> gældende for alle medarbejdere. 
            <input type="hidden" name="FM_brug_fasttp" value="<%=brug_fasttpVal%>" />
            Timepris:  <input type="text" DISABLED name="" value="<%=fasttp %>" style="width:100px;" /> <input type="hidden" name="FM_fasttp" value="<%=fasttp %>" style="width:100px;" /> 
            <%
              else

                lstTxt = ""
                valIdSQLkri = " id <> 0 "
                %>
                 <input type="checkbox" value="1" <%=brug_fasttpDis %> name="FM_brug_fasttp" id="FM_brug_fasttp" <%=brug_fasttpCHK %> /> Angiv <b>fast timepris</b> gældende for alle medarbejdere. 
                Timepris:  <input type="text" name="FM_fasttp" value="<%=fasttp %>" style="width:100px;" /> 
                <%
           
               end if %>
       


       
        
        <%strSQLval = "SELECT valutakode, id, grundvaluta FROM valutaer WHERE "& valIdSQLkri &" ORDER BY grundvaluta DESC" 
        

            'response.write strSQLval
        %>
        <select name="FM_fasttp_val" style="width:60px;">
        <%

        oRec5.open strSQLval, oConn, 3
        while not oRec5.EOF 
        
        if cint(fasttp_val) = oRec5("id") OR ((func = "opret" OR func = "opretstam") AND oRec5("grundvaluta") = 1) then
        seltpVal = "SELECTED"
        else 
        seltpVal = ""
        end if 

        %>
        <option value="<%=oRec5("id") %>" <%=seltpVal %>><%=oRec5("valutakode") %></option>
        <%
        oRec5.movenext
        wend
        oRec5.close%>

        </select>&nbsp;<%=lstTxt %> <br />

                    <%if func <> "redstam" AND func <> "opretstam" then 
            
                        select case lto
                        case "dencker", "intranet - local"
                        useDate = dateAdd("yyyy", -3, now)
                        useDate = formatdatetime(useDate,2)
                        case else
                        useDate = formatdatetime(now,2)
                        end select

                        %>
                    <input type="checkbox" value="1" name="FM_opdater_timepriser" /> Opdater timepriser på eksisterende timer<br /><br />
                    Opdater time- og kost -priser (hvis valgt) fra d. <input type="text" name="FM_opdatertpfra" value="<%=useDate %>" style="font-size:9px; width:60px;" /> til dd.<br />
                    (også lukkede uger og hvis der foreligger faktura)
                    <%end if %>

        </div>


        <br /><br />
            
                <%if func <> "opretstam" AND showtp = 0 then %>
                <b> Eller angiv <a id="aktshowtp" href="aktiv.asp?menu=job&func=<%=func%>&id=<%=id%>&jobid=<%=jobid%>&jobnavn=<%=request("jobnavn")%>&rdir=<%=rdir %>&nomenu=<%=nomenu%>&showtp=<%=showtpVal%>&aktfavgp=<%=aktfavgp%>" class=vmenu> <u>individuel</u> timepris for hverenkelt medarbejder..</a></b>
                <%else %>
                 <b> Eller angiv <u>individuel</u> timepris for hverenkelt medarbejder herunder..</b> (klik fast timepris fra i feltet herover)
                <%end if %>
           
        
		            <%if func = "opretstam" OR func = "redstam" then%>
		            <br />Når et job oprettes kan det vælges om aktiviteter skal nedarve timeprisen fra jobbet eller om aktiviteten skal bruge den timepris der er angivet her.
        
            
		
		            <%else%>
	    
                    <!--
                    <br>Tildel timepriser på de medarbejderne der er tilknyttet dette job via deres projektgrupper.<br>
		
                    <font class=roed>Hvis der skiftes timepris, ændres alle hidtidige timeregistreringer til den nye timepris, også i perioder, der er faktureret, eller afsluttet.
		            <br /><b>En aktivitet kan kun antage en timepris pr. medarbejder.</b></font><br />
                    -->
		            <%end if%>
		<br><br>


          





        
	<% 

    if showtp = 1 OR func = "opretstam" then
        
		
        %>

            <div id="div_akt_med_tp"><%

              

            mtypeSQLKri = " medarbejdertype <> 0"
            tp_jobid = jobid
            tp_jobaktid = id

            call timepriser_akt %>
				
				<br>
				<%if func = "opretstam" OR func = "redstam" then%>
				
				<%else%>
				*) Aktivitetens timepris nedarves fra den valgte timepris på jobbet. Så længe "Valgt Timepris" er blank (tom) nedarves timepris på denne aktivitet fra jobbet.
				Hvis der ændres projektgrupper, skal du opdatere aktiviteten, før den aktuelle liste af medarbejdere vises.<br />
		        <br /><br />
                <input type="checkbox" value="1" name="FM_opdater_timepriser_individuel" /> Opdater timepriser på eksisterende timer for de viste medarbejdere<br /><br />
        Opdater time- og kost -priser (hvis valgt) fra d. <input type="text" name="FM_opdatertpfra_individuel" value="<%=formatdatetime(now,2) %>" style="font-size:9px; width:60px;" /> til dd.<br />
        (også lukkede uger og hvis der foreligger faktura)
        <br /><br />

            &nbsp;
				<%end if%>


              
            </div>
				</td>
		</tr>
		</table>
		</div>
        <%end if %>

	</td>
	</tr>
    <%else %>

    <tr>
		<td style="padding-left:40px; padding-top:20px; padding-bottom:20px; padding-right:20px;" colspan=2><img src="../ill/ac0001-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Medarb. timepriser på denne aktivitet:</b><br />
        Timepriser kan først ændres efter aktiviteten er oprettet, i første omgang nedarves de autimatisk fra jobbet. (Opret aktiviteten og klik derefter på rediger)

        </td>
	</tr>

	<%end if%>
	
	
	<tr><td colspan=2 align="right" style="padding:20px 40px 30px 10px;"><input type="submit" value="Opdater >>" /></td></tr>
	
	</table>
	
	</div><!-- table div -->
	
	
	

	</form>
	
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a><br>
	<br>
	<br>
	</div>
	
	
	
	
	
	
	<%case "favorit" 'viser liste med aktviteter i stamgruppe
	
	if len(request("stamaktid")) <> 0 then
	stamaktid = request("stamaktid")
	else
	stamaktid = 0
	end if

    call fomr_account_fn()
	%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	
	<%call menu_2014() %>

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; visibility:visible;">
	
	<%
   'oimg = "aktstam_48.png"
	oleft = 0
	otop = 0
	owdt = 800
	oskrift = "Stamaktivitets-gruppe: " & request("stamakgrpnavn")
	
	'call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
    %>
	
	<!-- job faneblad -->
	<table cellspacing="2" cellpadding="2" border="0" width=200>
	<tr><td colspan=2>
	    <a href="akt_gruppe.asp?menu=job&func=favorit" class=vmenu> << Stam-aktivitetsgrupper</a><br /><br />
    
	</td></tr>
		
	</table>
	
	<!-- slut -->
	
	<%
		tTop = 0
	tLeft = 0
	tWdth = 1004
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	%><h4>&nbsp;<%=oskrift %></h4>
       
	

          <div style="position:absolute; left:800px; top:40px;"><% 
                nWdt = 160
                nTxt = "Opret ny stam-aktivitet"
                nLnk = "aktiv.asp?menu=job&func=opretstam&jobid=0&id=0&jobnavn="&request("stamakgrpnavn")&"&aktfavgp="&id
                nTgt = ""
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) %> 
	        </div>

        <br /><br />

	<table cellspacing="0" cellpadding="0" border="0" width="100%" id="incidentlist">
 	
	<tr bgcolor="#5582D2">
	<td>&nbsp;</td>
		<td class='alt' style="width:200px;"><b>Stamaktiviteter</b></td>
		<td class=alt><b>Fase</b></td>
		<td class=alt><b>Forretningsområde</b></td>
		
		<td class='alt'><b>Type</b><br />(egenskab)</td>
          <td class="alt" style="padding-right:20px; width:120px;"><b>Budget antal, pris</b><br />(Stk. ell. timer)</td>
        <td class="alt" style="padding-right:20px; width:120px; white-space:nowrap;"><b>Real. time- og kost -pris</b><br />
            (fast eller individuel)</td>

        <%if cint(fomr_account) = 1 then %>
        <td class="alt" style="width:200px;"><b>Konto</b></td>
        <%end if %>
        <td class='alt' align="right"><b>Fjern fra gruppe</b><br />
		(slet)</td>
	<td>&nbsp;</td>
	</tr>
	<%

    if cint(fomr_account) = 1 then 
        cspa = 10
    else
        cspa = 9
    end if


	if id <> 0 then
	sqlKriakt = "aktfavorit = "& id &""
	else
	sqlKriakt = "aktfavorit <> 0"
	end if
	
	strSQL = "select a.job, a.navn, a.id, aktfavorit, b.navn AS gnavn, "_
	&" b.id AS gid, a.fakturerbar, a.fase, a.sortorder, a.easyreg, a.fravalgt, fasttp, brug_fasttp, fastkp, brug_fastkp, aktkonto, aktpris, antalstk, aktbudgetsum, aktbudget, bgr, budgettimer"_
	&" FROM aktiviteter a, akt_gruppe b "_
	&" WHERE "& sqlKriakt &" AND b.id = a.aktfavorit ORDER BY a.fase, a.sortorder, a.navn"
	a = 0
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	visning = 1
	call akttyper(oRec("fakturerbar"), visning)
	
	
	if cdbl(stamaktid) = oRec("id") then
	bgthis = "#ffff99"
	else
	    select case right(a,1)
        case 0,2,4,6,8
        bgthis = "#FFFFFF"
        case else
    	bgthis = "#EFF3FF"
        end select
	end if

        

        %>

	<tr>
		<td bgcolor="#cccccc" colspan="<%=cspa %>"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bgthis%>">
		<td valign="top">
			<input type="hidden" name="SortOrder" class="SortOrder" value="<%=oRec("sortorder")%>" />
	        <input type="hidden" name="rowId" value="<%=oRec("id")%>" />
		<img src="../ill/blank.gif" width="1" height="30" alt="" border="0"></td>
		<td style="white-space:nowrap;" ><img src="../ill/pile_drag.gif" alt='Klik, træk og sorter aktivitet' class="drag" border="0" />&nbsp;&nbsp;
		<%if oRec("gid") <> 2 then%>
		<a href="aktiv.asp?menu=job&func=redstam&jobid=0&id=<%=oRec("id")%>&jobnavn=<%=request("stamakgrpnavn")%>&aktfavgp=<%=id%>"><%=oRec("navn")%></a>
		<%else%>
		<%=oRec("navn")%>
		<%end if%>
		
		<%if oRec("easyreg") <> 0 then %>
		&nbsp;<span style="font-size:9px; font-family:arial;">(Easyreg.)</span>
		<%end if %>

        <%if oRec("fravalgt") <> 0 then %>
		&nbsp;<span style="font-size:9px; font-family:arial;">(Fravalgt)</span>
		<%end if %>
        
		</td>
		
		<td>
		<%if len(trim(oRec("fase"))) <> 0 then %>
	    <%=oRec("fase") %>
	    <%end if %>&nbsp;
		</td>
		
		<td style="width:200px; padding:4px 4px 4px 0px;" valign=top>

         <%
                    '*** Forretningsområder **' 
	                strFomr_navn = ""
                    strFomr_id = ""
                    visJobFomr = 0

                    strSQLfrel = "SELECT for_fomr, fomr.navn FROM fomr_rel "_
                    &" LEFT JOIN fomr ON (fomr.id = for_fomr) WHERE for_aktid = "& oRec("id") 

                    'Response.Write strSQLfrel
                    'Response.flush
                    f = 0
                    oRec3.open strSQLfrel, oConn, 3
                    while not oRec3.EOF

                    if f = 0 then
                    strFomr_navn = " ("
                    end if

                    strFomr_navn = strFomr_navn & oRec3("navn") & ", " 
                    strFomr_id = strFomr_id &",#"& oRec3("for_fomr") & "#"

                    if instr(strFomr_rel, "#"&oRec3("for_fomr")&"#") <> 0 then
                    visJobFomr = 1
                    else
                    visJobFomr = visJobFomr
                    end if

                    f = f + 1
                    oRec3.movenext
                    wend
                    oRec3.close

                    if f <> 0 then
                    len_strFomr_navn = len(strFomr_navn)
                    left_strFomr_navn = left(strFomr_navn, len_strFomr_navn - 2)
                    strFomr_navn = left_strFomr_navn & ")"

                        if len(strFomr_navn) > 50 then
                        strFomr_navn = left(strFomr_navn, 50) & "..)"
                        end if

                    end if		    
                    

                     
                     if f <> 0 then%>

                     <span style="color:#999999;">
                     <%=strFomr_navn %></span>

                     <%end if %>
		

		</td>
		
		
		<td style="padding-right:10px;"><%=akttypenavn %></td>
        <td style="padding-right:10px;">

            <%select case oRec("bgr")
                case "1"
                grundlagTxt = "<b>" & oRec("budgettimer") & "</b> t. á "& formatnumber(oRec("aktbudget"), 2)
                case "2"
                grundlagTxt = "<b>" & oRec("antalstk") & "</b> stk. á "& formatnumber(oRec("aktbudget"), 2)
                case else
                grundlagTxt = ""
                end select %>


          <%=grundlagTxt %>
        </td>
        
        <td style="white-space:nowrap;">
        

         


        <%if cint(oRec("brug_fasttp")) = 1 then %>
          Timepris: <%=formatnumber(oRec("fasttp"), 2) %>
            <%else %>
            <!--
        <span style="color:#999999;">Timepris:</span> Individuel
            -->
        <%end if %>

            <br />
            <%if cint(oRec("brug_fastkp")) = 1 then %>
           Kostpris: <%=formatnumber(oRec("fastkp"), 2) %>
            <%else %>
            <!--
        <span style="color:#999999;">Kostpris:</span> Individuel 
            -->
        <%end if %>
       
        </td>


         <%if cint(fomr_account) = 1 then %>
        <td>
            <% strSQLkonto = "SELECT navn, kontonr, id FROM kontoplan WHERE id = "& oRec("aktkonto")
         oRec2.open strSQLkonto, oConn, 3
         if not oRec2.EOF then
            
           %><%=oRec2("navn") & " ("& oRec2("kontonr") &")" %>


        </td>
        <%
         
         end if 
         oRec2.close
        end if%>

        <%if oRec("gid") <> 2 then%>
		<td align="right"><a href="aktiv.asp?func=favorit_fjern&id=<%=oRec("id")%>&favpgid=<%=id%>&favgpnavn=<%=request("stamakgrpnavn")%>"><img src='../ill/slet_16.gif' alt="Fjern fra gruppe" border='0'></a></td>
		<%else%>
		<td>&nbsp;</td>
		<%end if%>

		<td valign="top" align="right"><img src="../ill/blank.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	lastfavgp = oRec("aktfavorit")
	a = a + 1
	oRec.movenext
	wend
	oRec.close
	
	if a = 0 then%>
	<tr>
		<td bgcolor="#5582D2" colspan="<%=cspa %>"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_blank.gif" width="1" height="76" alt="" border="0"></td>
		<td height="76" valign="top" colspan="<%=cspa-2 %>">&nbsp;<br>Der er ingen stam-aktiviteter i denne gruppe!</td>
		<td valign="top" align="right"><img src="../ill/blank.gif" width="1" height="76" alt="" border="0"></td>
	</tr>
	<%end if%>
	
	</table>
	<!-- table div -->
	</div>
	
	
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a><br><br>
    <br />
	<br>
	<!--pagehelp-->

<%

itop = 20
ileft = 850
iwdt = 120
ihgt = 0
ibtop = 40 
ibleft = 150
ibwdt = 600
ibhgt = 250
iId = "pagehelp"
ibId = "pagehelp_bread"
call sideinfoId(itop,ileft,iwdt,ihgt,iId,phDsp,phVzb,ibtop,ibleft,ibwdt,ibhgt,ibId)
%>
	
	<b>Hvordan du bruger stam-aktivitetsgrupper:</b> (job-skabeloner)<br><br />
	Aktiviter i en stam-aktivitetsgruppe kan tilføjes til et job, ved at tilknytte stam-aktivitetsgruppen til jobbet når det oprettes eller redigeres.
	<br /><br />Der kan ved joboprettelse tilføjes optil 5 stam-aktivitetsgrupper.
	
</td>
	</tr></table></div>
	
	
	<br>
	<br>
	</div>
	<%case "favorit_fjern"
	
	'***********************************************************************
	' Fjerner favorrit akt. fra fav gruppe. Den bliver ikke slette da de
	' kan være den er i brug på et job. *
	' Hvis den er oprettet "on the fly" i opret akt. under job.
	'***********************************************************************
	 
	strSQL = "DELETE FROM aktiviteter WHERE id =" & id
	oRec.open strSQL, oConn, 3
	
	Response.redirect "aktiv.asp?menu=job&func=favorit&id="&request("favpgid")&"&stamakgrpnavn="&request("favgpnavn")&""
	
	
	
	case else%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<script src="inc/aktiv_jav.js" language="javascript" type="text/javascript"></script>


    <%if nomenu <> 1 then %>
    
  
    <%
    call menu_2014()   
        
        lftpx = 90
        toppx = 102 
        
    else

        lftpx = -30
        toppx = 20 

    end if %>



	<!-------------------------------Sideindhold------------------------------------->
	
	<%
	if len(request("jobid")) <> 0 then
	jobid = request("jobid")
	else
	jobid = 0
	end if
	

      
	
	%>
	<!-- job faneblad -->
	<!--
	<div id="jobnavn" style="position:relative; left:30px; top:52px; width:300px; visibility:visible; border:0px #000000 solid; z-index:1200;">
	
	<table cellspacing="0" cellpadding="5" border="0" width=100%>
		
		<tr>
			<td style="border:1px orange solid; border-bottom:0px orange solid;" bgcolor="#ffff99">
			-->
			<%
			
			
			strSQL = "SELECT jobnavn, jobnr, jobknr FROM job WHERE id = "& jobid
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
				jobnavnThis = oRec("jobnavn")
				jobnrThis = oRec("jobnr")
				kid = oRec("jobknr")
				
			end if
			oRec.close
			%>
			<!--<a href="jobs.asp?menu=job&func=red&id=<%=jobid%>&int=1&rdir=<%=rdir%>">&nbsp;<%=jobnavnThis %>&nbsp;(<%=jobnrThis%>)</a></td>
		</tr>
	</table>
	
	</div>-->
	<!-- slut -->
	<%
	
	
	oimg = "ikon_akt_48.png"
	oleft = 20
	otop = 20
	owdt = 600
	oskrift = "Aktiviteter & Faser"
	
	'call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	'*** Faneblade ***'
	'side = "aktiv"
	'div = ""
	'call faneblade(side, div)
	
	
	
	
	
	'*** timer tildelt på job ***'
	call budgettimerTildelt(strjobnr)
    call showEasyreg_fn()
	%>
	
	
	
	
	<div id="sindhold" style="position:relative; left:<%=lftpx%>px; top:<%=toppx%>px; border:0px #000000 solid; visibility:visible;">
	
	
	
	<%
	
        
        
        'public faTot, lastFaseTimer, lastFaseSum
	function faSumlinje(lastFase, lastFaseSum, lastFaseTimer)
	
	%>
	<tr bgcolor="#FFFFFF">
	    
	    <td colspan=6 align=right style="border-bottom:0px #cccccc solid; padding:5px 5px 15px 5px;">&nbsp;<b>Ialt på fase:</b> (kun viste aktiviteter)</td>
	    <td colspan=2 align=right style="padding:5px 10px 15px 5px; height:20px; border-bottom:0px #cccccc solid;">
            <div id="sltimer_<%=lcase(lastFase)%>" style="width:100px;"><b><%=formatnumber(lastFaseTimer, 2) & "</b> timer" %></div></td>
		<td colspan=2 align=right style="padding:5px 10px 15px 5px; height:20px; border-bottom:0px #cccccc solid;">
            <div id="slsum_<%=lcase(lastFase)%>" style="width:100px;"><b><%=formatnumber(lastFaseSum, 2) & "</b> "& basisValISO &" " %></div></td>
	    <td colspan=4 style="border-bottom:0px #cccccc solid; padding:5px 10px 15px 5px;">&nbsp; <input id="fa_<%=lcase(lastFase)%>" value="<%=fa %>" type="hidden" />
	    <input id="fatot_<%=lcase(lastFase)%>" value="<%=faTot %>" type="hidden" />
	    <input id="fatotbn_<%=faTot%>" value="<%=lcase(trim(lastFase))%>" type="hidden" />
		
	    <input id="fatot_val_<%=faTot %>" value="<%=formatnumber(lastFaseSum, 2) %>" type="hidden" />
	    <input id="fatottimer_val_<%=faTot %>" value="<%=formatnumber(lastFaseTimer, 2) %>" type="hidden" /></td>
	</tr>
	<%
	faTot = faTot + 1
	lastFaseSum = 0
	lastFaseTimer = 0
	end function
	
	'** skal job sync ***'
	skaljobSync(jobid)
	
	
	
	
	tTop = 20
	tLeft = 0
	tWdth = 1520
	


	call tableDiv(tTop,tLeft,tWdth)
	
	%>
    	     
   <form action="aktiv.asp?menu=job&func=opdaktliste&rdir=<%=rdir%>&jobid=<%=jobid%>&jobnavn=<%=request("jobnavn")%>&nomenu=<%=nomenu%>" method="post" id="form_aktlist">
       
	<table cellspacing="0" cellpadding="0" border="0" width=100% bgcolor="#FFFFFF" id="incidentlist">
       
             <input type="hidden" id="hd_jobid" value="<%=jobid%>"  />
	<tr>
	    <td colspan=5 style="padding:5px 70px 20px 10px;"><h4>Aktivitetsliste<br />
	    <span style="font-size:11px;">Job: <b><%=jobnavnThis%></b> (<%=jobnrThis %>)</span> </h4>


            <%
                  LnkUrl = "aktiv.asp?menu=job&func=opret&jobid="&jobid&"&id=0&jobnavn="&request("jobnavn")&"&fb=1&rdir=list&nomenu="&nomenu 
                  lnkTxt = "Opret ny aktivitet"
                  lnkAlt = lnkTxt 
                  lnkWdt = 120
                  lnkTgt = ""
                  call opretLink_2013(lnkUrl, lnkTxt, lnkAlt, lnkTgt, lnkWdt) 
                %>
            
        
            <br />

	       <input id="vispasluk" name="vispasluk" type="checkbox" value="1" <%=vispaslukCHK %> /> <b>Vis passive og lukkede aktiviteter</b><br />
	        <!--<input id="visbesk" name="visbesk" type="checkbox" value="1" <%=visbeskCHK %> /> <b>Vis beskrivelses felter</b> <span style="color:#999999;">(opdater liste for at gemme indstillinger)</span>-->
    
     
        
           
	    </td>
         
	    <td colspan=6><!--<a href="job_print.asp?id=<%=jobid%>&kid=<%=kid %>" class=vmenu>Print / PDF</a>-->&nbsp;</td>
	    <td align=right colspan=4 style="padding:20px 70px 20px 10px;"><input type="submit" name="statusliste" value="Opdater liste >> "></td>
	</tr>
	<tr bgcolor="#5582D2">
	<td class='alt' style="padding-left:10px; width:300px;"><b>Aktivitet</b><br />
   Beskrivelse
	</td>
    <td class='alt' style="padding-right:10px;"><b>St. og slut dato</b><br />
    dd-mm-åååå<br />
     Projektgrupper</td>
	<td class='alt' style="padding-right:10px;"><b>Type</b><br />
    Forretningsomr.</td>
	<td class='alt' style="padding-right:10px;">Fase</td>
    <%if cint(showEasyreg_val) = 1 then %>
    <td class='alt' style="padding-right:10px;">Easyreg.</td>
    <%else %>
    <td class='alt' style="padding-right:10px;">&nbsp;</td>
    <%end if %>
	<td style="padding-left:10px;" class='alt'>Realiseret %</td>
	<td class='alt' align=right style="padding-right:10px; width:70px;">Antal stk.</td>
	<td class='alt' align=right style="padding-right:10px; width:100px;"><b>Timer forkalk.</b>
	<br />Realiseret<br />
	Balance</td>
	<td class='alt'>Beregnings grundlag</td>
	<td class=alt align=right style="padding-right:10px; width:110px;"><b>Pris pr. stk / time</b><br />
	Pris i alt</td>
	<td class='alt'>Faktor<br />
    timer*faktor=enh.<br /><span style="color:#CCCCCC; font-size:10px;">Opdaterer ikke eksisterende timer</span></td>
    <td align="center" class='alt'>Status</td>

    <td class='alt'>Real. timepris og kostpris på medarbejdere<br /><br /><span style="color:#CCCCCC; font-size:10px;">Individuel eller fast</span></td>
	<td align="right" style="padding-right:10px;" class='alt'>Slet?</td>
	</tr>
	<%
	
	samletverdi = 0
	sort = Request("sort")
	select case sort
	case "slutdato"
	orderBy = "aktslutdato, a.navn"
	case else
	orderBy = "a.fase, a.sortorder, a.navn"
	end select
	
	lastFase = ""
	lastFaseSum = 0
	
	if vispasluk = 1 then
	vispaslukKri = " AND aktstatus <> - 1"
	else
	vispaslukKri = " AND aktstatus = 1"
	end if
	
	strSQL = "SELECT a.id, a.navn, job, beskrivelse, fakturerbar, aktstartdato, "_
	&" aktslutdato, budgettimer, aktstatus, aktbudget, fomr.navn AS fomr, antalstk, tidslaas, a.fase, a.sortorder, a.bgr, easyreg, faktor, "_
    &" projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, fasttp, fasttp_val, fastkp, fastkp_val "_
	&" FROM aktiviteter a LEFT JOIN fomr ON (fomr.id = a.fomr) "_
    &" WHERE job = "& jobid &" "& vispaslukKri &" ORDER BY "& orderBy 
	
	'Response.Write strSQl
	'Response.flush
	c = 0
	fa = 0
	faTot = 0
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
	        strSQL = "SELECT sum(timer) AS proafslut FROM timer WHERE TaktivitetId=" & oRec("id")
	        oRec3.open strSQL, oConn, 3
	
	        if len(oRec3("proafslut")) <> 0 then
	        proaf = oRec3("proafslut")
	        else
	        proaf = 0
	        end if
	
	        oRec3.close
	
	visning = 1
	call akttyper(oRec("fakturerbar"), visning)
	

	
	if len(trim(oRec("budgettimer"))) = "0" OR oRec("budgettimer") = "0" then 
	budgettimer = 0
	else
	budgettimer = oRec("budgettimer")
	end if
	
	if budgettimer > 0 then
	projektcomplt = ((proaf/budgettimer)*100)
	else
	projektcomplt = 100
	end if
	
	timerbal = (budgettimer - proaf)
	
	'if projektcomplt > 100 then
	'projektcomplt = 100
	'end if
	
	if projektcomplt > 100 then
	showprojektcomplt = projektcomplt
	projektcomplt = 100
	bgdiv2 = "Crimson"
	else
	showprojektcomplt = projektcomplt
	projektcomplt = projektcomplt
	bgdiv2 = "DarkSeaGreen"
	end if
	
	if cdbl(id) = oRec("id") then
	bgthis = "#ffff99"
	else
	    
	select case right(c,1)
	case 0,2,4,6,8
	bgthis = "#FFFFFF"
	case else
	bgthis = "#F3F3F3" '"#EFF3FF"
	end select
	    

	end if
	
	
	if lcase(lastFase) <> lcase(oRec("fase")) AND len(trim(oRec("fase"))) <> 0 then 
	
	if c <> 0 then
	    call faSumlinje(lastFase, lastFaseSum, lastFaseTimer)
	end if
	
	    'if cint(oRec("aktstatus")) <> 1 then 'første akt. aktiv = fase åben
	    'fswsb = "hidden"
	    'fsdsp = "none"
	    'else
	    fswsb = "visible"
	    fsdsp = ""
	    'end if
	
	%>
	
        <!--
	<tr>
		<td bgcolor="#FFFFFF" colspan="14" style="padding:2px 0px 0px 0px; border-top:1px #cccccc solid;"><img src="ill/blank.gif" width="1" border="0" height="1" /></td>
	</tr>
        -->
	<tr>
		<td bgcolor="#8caae6" colspan="14" style="padding:10px 10px 2px 10px; border-top:1px #cccccc solid;"><a href="#" id="showhidefase_<%=lcase(trim(oRec("fase")))%>" class="showhidefase" style="color:#FFFFFF;">+ fase: <%=replace(oRec("fase"), "_", " ")%></a>&nbsp;</td>
	</tr>
	
        <tr class="trfase_<%=lcase(trim(oRec("fase")))%>" bgcolor="#8caae6" style="visibility:<%=fswsb%>; display:<%=fsdsp%>;">
        <td>&nbsp;</td>
	

         <td colspan=1 valign=top style="padding:3px 10px 2px 3px; white-space:nowrap;">
        <input id="dt_st_<%=lcase(trim(oRec("fase")))%>" type="text" style="width:65px; font-size:9px;" value=""  class="fs_dt_st" />
        - <input id="dt_sl_<%=lcase(trim(oRec("fase")))%>" type="text" style="width:65px; font-size:9px;" value="" class="fs_dt_sl" />  
        </td>
        

        <td>&nbsp;</td>
    
        <td colspan="4" valign=top style="padding:5px 10px 2px 5px;">
            <input id="fs_navn_<%=lcase(trim(oRec("fase")))%>" name="" class="faseoskrift_navn" value="<%=replace(oRec("fase"), "_", " ")%>" type="text" style="width:210px;" />
        </td>

		<td align=right colspan="6" valign=top style="padding:5px 5px 2px 12px;">
		Status på akt. i fase:&nbsp;&nbsp;
		<select name="faseoskrift" class="faseoskrift" id="<%=lcase(trim(oRec("fase")))%>" style="font-size:9px;">
	<option value="1">Vælg..</option>
	<option value="1">Aktiv</option>
	<option value="0">Lukket</option>
	<option value="2">Passiv</option>
	</select>
		</td>
		<td valign=top align=right style="padding:5px 12px 2px 12px;">
            <input id="sl_<%=lcase(trim(oRec("fase")))%>" class="faseoskrift_slet" type="checkbox" value="1" />
           </td>
	</tr>
	
	
	<%
	
	lastFase = lcase(oRec("fase")) 
	fa = 0
	else
	%>
	<!--<tr>
		<td colspan="14"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>-->
	<%
	end if %>
	
	
	
	<tr class="trfase_<%=lcase(trim(oRec("fase")))%>" bgcolor="<%=bgthis%>" style="visibility:<%=fswsb%>; display:<%=fsdsp%>;">
	
	<td valign=top style="padding:3px 10px 2px 10px; white-space:nowrap;">
	    <img src="../ill/pile_drag.gif" alt='Klik, træk og sorter aktivitet' class="drag" border="0" />&nbsp;
        <a href="aktiv.asp?menu=job&func=red&id=<%=oRec("id")%>&jobid=<%=jobid%>&jobnavn=<%=request("jobnavn")%>&nomenu=1" class=vmenu>
         <img src="../ill/blyant.gif" border="0" /></a>
        <input id="FM_aktnavn" name="FM_aktnavn" type="text" value="<%=oRec("navn")%>" style="width:350px;" /> <br />
        <input id="Text1" name="FM_aktnavn" type="hidden" value="#" />
	   
       



	<input type="hidden" class="SortOrder" name="SortOrder" value="<%=oRec("sortorder")%>" />
	<input type="hidden" name="rowId" value="<%=oRec("id")%>" />
	
	
	
	<%if oRec("tidslaas") = "1" then %>
	(tidslås aktiv)
	<%end if %>
   

      
    <div class="dv_besk" style="position:relative; visibility:visible; display:; font-size:10px; padding:3px 0px 2px 0px; color:#999999; border:0px #999999 solid;"> 
        Beskrivelse:<br />
            <textarea name="<%="FM_beskrivelse_" & oRec("id") %>" style="font-size:10px; width:377px; height:50px; border:1px #CCCCCC solid;"><%=oRec("beskrivelse")%></textarea>
           <%
    
    
        
	                    
	                    'content = oRec("beskrivelse")
            			
			            
			            'Set editorK = New CuteEditor
            					
			            'editorK.ID = "FM_beskrivelse_" & oRec("id")
			            'editorK.Text = content
			            'editorK.FilesPath = "CuteEditor_Files"
			            'editorK.AutoConfigure = "Minimal" 'none
                        
                        'if c > 0 then
                        'editorK.SubSequent = true
                        'end if

            			
			            'editorK.Width = 380
			            'editorK.Height = 280
			            'editorK.Draw() %>
                        <!--<br />&nbsp;-->
	
        </div>   

 
  
   
    </td>

   

    <% 
    
	                  

    %>
	
	
    <td valign=top style="padding:3px 10px 2px 3px; white-space:nowrap;">
        <input id="af_st_dato_<%=lcase(trim(oRec("fase")))%>_<%=fa%>" type="text" style="width:65px; font-size:9px;" value="<%=formatdatetime(oRec("aktstartdato"), 2)%>" name="FM_st_dato" />
        - <input id="af_sl_dato_<%=lcase(trim(oRec("fase")))%>_<%=fa%>" type="text" style="width:65px; font-size:9px;" value="<%=formatdatetime(oRec("aktslutdato"), 2)%>" name="FM_sl_dato" />  
        <br />
         <!-- Projektgrupper -->
        
      
        <div style="font-size:10px; color:#5C75AA; height:65px; width:150px; padding:6px 0px 3px 0px; overflow:hidden;">
        Projektgrupper:<br />
        <%for pgid = 1 to 10 %>
            <%call prgNavnShort(oRec("projektgruppe"& pgid &""), pgid) %>

            <%select case pgid
               case 3,6,9
               %><br />
                <% 
               end select %>

        <%next %>

        </div><br />
        </td>

	<td valign=top style="padding:3px 10px 2px 3px; font-size:10px;">
        
            <%strFakturerbart = oRec("fakturerbar") %>

              <select id="Select1" name="FM_fakturerbart" style="background-color:#FFFFe1; width:100px; font-size:9px; font-family:arial;">
                       <%
                       call akttyper2009(1)
                       Response.Write aty_options
                       %>
                       
                      
                    </select>

                    

                    <%
                    '*** Forretningsområder **' 
	                strFomr_navn = ""
                    strFomr_id = ""
                    visJobFomr = 0

                    strSQLfrel = "SELECT for_fomr, fomr.navn FROM fomr_rel "_
                    &" LEFT JOIN fomr ON (fomr.id = for_fomr) WHERE for_aktid = "& oRec("id") 

                    'Response.Write strSQLfrel
                    'Response.flush
                    f = 0
                    oRec3.open strSQLfrel, oConn, 3
                    while not oRec3.EOF

                    if f = 0 then
                    strFomr_navn = ""
                    end if

                    strFomr_navn = strFomr_navn & oRec3("navn") & ", " 
                    strFomr_id = strFomr_id &",#"& oRec3("for_fomr") & "#"

                    if instr(strFomr_rel, "#"&oRec3("for_fomr")&"#") <> 0 then
                    visJobFomr = 1
                    else
                    visJobFomr = visJobFomr
                    end if

                    f = f + 1
                    oRec3.movenext
                    wend
                    oRec3.close

                    if f <> 0 then
                    len_strFomr_navn = len(strFomr_navn)
                    left_strFomr_navn = left(strFomr_navn, len_strFomr_navn - 2)
                    strFomr_navn = left_strFomr_navn & ""

                        if len(strFomr_navn) > 50 then
                        strFomr_navn = left(strFomr_navn, 50) & ".."
                        end if

                    end if		    
                    

                     
                     if f <> 0 then%>
                     <br />
                     <span style="color:#999999;"><b>Forretningsomr.:</b><br />
                     <%=strFomr_navn %></span>

                     <%end if %>
    
    </td>
	
	<td valign=top style="padding:3px 10px 2px 3px;">
	<%if len(oRec("fase")) <> 0 then %>
	<%strFase = replace(oRec("fase"), "_", " ")%>
    <%else
     strFase = ""  
     end if %>
       
        <input id="af_n_<%=lcase(trim(oRec("fase")))%>_<%=fa%>" name="FM_fase" value="<%=strFase%>" type="text" style="width:110px; font-size:9px;" />
        <input id="Hidden1" name="FM_fase" type="hidden" value="#" />
        <!--<input id="Hidden1" name="FM_fase" type="hidden" value="#" />-->
	
	&nbsp;
	</td>
	


    <%
    if cint(showEasyreg_val) = 1 then 
    
    if oRec("easyreg") = 1 then
    easCHK = "CHECKED"
    else
    easCHK = ""
    end if %>

    <td valign=top style="padding:3px 3px 2px 3px;">
	<input name="FM_easyreg_aid_<%=c %>" type="checkbox" value="1" <%=easCHK %> />
    <input name="FM_easyreg_aid" type="hidden" value="<%=c %>" />

	</td>
    <%else %>
    <td valign=top style="padding:3px 3px 2px 3px;">&nbsp;</td>

    <input name="FM_easyreg_aid_<%=c %>" type="hidden" value="0" />
    <input name="FM_easyreg_aid" type="hidden" value="<%=c %>" />
    <%end if %>    
	
	
		<td valign=top align="left" style="padding:5px 10px 5px 10px; width:110px; border:0px #000000 solid;">
		<%if proaf <> 0 then %>
		<div style="position:relative; border:1px #000000 solid; background-color:<%=bgdiv2%>; height:15px; width:<%=cint(left(projektcomplt, 3))%>px;">
		<%if projektcomplt > 0 then%>
		&nbsp;&nbsp;<%=formatpercent(showprojektcomplt/100, 0)%>
		<%end if%></div>
		<%end if %>
		</td>
		
	
	<td align="right" valign=top style="padding:3px 10px 2px 2px;">
        <input class="timstkpris" id="FM_stk_<%=oRec("id") %>" name="FM_antalstk" value="<%=oRec("antalstk")%>" type="text" style="width:60px; font-size:9px;" /><br /><b>Stk.</b>
       </td>
	
	<%antalstkialt = antalstkialt + oRec("antalstk") %>
	
		
	
	<td align=right valign=top style="padding:3px 10px 2px 2px;">
        <input class="timstkpris" name="FM_timer" id="FM_tim_<%=oRec("id") %>" type="text" value="<%=formatnumber(budgettimer, 2)%>" style="width:75px; font-size:9px;" />
        Real.: <%=formatnumber(proaf, 2)%> t. 
	
	
	<%
	'*** Kun typer der er med i timeregnskab ***'
	call akttyper2009prop(oRec("fakturerbar"))
	if aty_real <> 0 then
	timerRealIalt = (timerRealIalt + proaf)  
	end if
	 %>
	
	<!--
	<%=formatnumber(timerbal, 2) %>
	
	--></td>
	
	
	<%
	bgrSEL0 = "SELECTED"
	bgrSEL1 = ""
	bgrSEL2 = ""
		
	
	select case oRec("bgr")
	case 0 '** Intent angivet
	bgrSEL0 = "SELECTED"
	aktsumtot = oRec("aktbudget")
	case 1 '** timer
	bgrSEL1 = "SELECTED"
	aktsumtot = formatnumber(oRec("aktbudget") * budgettimer) 
	case 2 '** Stk
	bgrSEL2 = "SELECTED"
	aktsumtot = formatnumber(oRec("aktbudget") * oRec("antalstk"))
	end select %>
	
    <td valign=top style="padding:3px 3px 2px 3px;">
        <select class="bgr" id="FM_bgr_<%=oRec("id") %>" name="FM_bgr" style="width:65px; font-size:9px;">
            <option value=0 <%=bgrSEL0 %>>Ingen</option>
            <option value=1 <%=bgrSEL1 %>>Timer</option>
            <option value=2 <%=bgrSEL2 %>>Stk.</option>
        </select></td>
	
	<td align="right" valign=top style="padding:3px 10px 2px 2px;"><input class="timstkpris" name="FM_pris" id="FM_pri_<%=oRec("id") %>" value="<%=formatnumber(oRec("aktbudget"), 2)%>" type="text" style="width:75px; font-size:9px;"/>
	<br />
	<div id="akt_totpris_<%=oRec("id") %>" style="width:80px;"><b><%=formatnumber(aktsumtot, 2) %></b> <%=basisValISO %></div>
	<input id="FM_akt_totpris_<%=oRec("id") %>" name="FM_akt_totpris" type="hidden" value="#<%=formatnumber(aktsumtot, 2) %>" />
	
	    
	<input name="FM_sum_aid_<%=oRec("id")%>" id="af_sum_<%=lcase(trim(oRec("fase")))%>_<%=fa%>" type="hidden" value="<%=formatnumber(aktsumtot, 2) %>" />
	<input name="FM_sum_aid_fs_<%=oRec("id") %>" id="FM_sum_aid_fs_<%=oRec("id") %>" type="hidden" value="<%=lcase(trim(oRec("fase")))%>" />
	<input name="FM_sum_aid_fa_<%=oRec("id") %>" id="FM_sum_aid_fa_<%=oRec("id") %>" type="hidden" value="<%=fa%>" />
	
	<input name="FM_timer_aid_<%=oRec("id")%>" id="af_timer_<%=lcase(trim(oRec("fase")))%>_<%=fa%>" type="hidden" value="<%=formatnumber(budgettimer, 2)%>" />
	
	</td>

    <%
    if len(oRec("faktor")) <> 0 then
    dblFaktor = formatnumber(oRec("faktor"), 2)
    else
    dblFaktor = formatnumber(1,0)
    end if %>

       <td style="padding:3px 2px 2px 2px;" valign=top>
           <input id="FM_faktor" name="FM_faktor" value="<%=dblFaktor %>" style="width:40px; font-size:10px;" type="text" />
           <input id="Text3" name="FM_faktor" value="#" type="hidden" />

       </td>
        
	<td align="center" valign=top style="padding:3px 2px 2px 2px;">
	<%select case oRec("aktstatus")
	case 1
	stCHK0 = ""
	stCHK1 = "SELECTED"
	stCHK2 = ""
	selbgcol = "DarkSeaGreen"
	case 2
	stCHK0 = ""
	stCHK1 = ""
	stCHK2 = "SELECTED"
	selbgcol = "#cccccc"
	case 0
	stCHK0 = "SELECTED"
	stCHK1 = ""
	stCHK2 = ""
	selbgcol = "Crimson"
	end select
	%>
	<select name="FM_listestatus" id="af_<%=lcase(oRec("fase"))%>_<%=fa%>" style="background-color:<%=selbgcol%>; font-size:9px;">
	<option value="1" <%=stCHK1%>>Aktiv</option>
	<option value="0" <%=stCHK0%>>Lukket</option>
	<option value="2" <%=stCHK2%>>Passiv</option>
	</select><br />
        &nbsp;
	<input type="hidden" name="FM_listeid" id="FM_listeid" value="<%=oRec("id")%>">
	</td>

     

          <td valign="top" style="padding-left:10px;">
               <span style="font-size:9px; color:#5C75AA;">
       
        <%if cint(oRec("fasttp_val")) = 1 then %>
        <b>Timepris:</b> Fast/Ens (<%=formatnumber(oRec("fasttp"), 2) %>)
            <%else %>
        <b>Timepris:</b> Individuel
        <%end if %>

            <br />
            <%if cint(oRec("fastkp_val")) = 1 then %>
        <b>Kostpris:</b> Fast/Ens (<%=formatnumber(oRec("fastkp"), 2) %>)
            <%else %>
        <b>Kostpris:</b> Individuel 
        <%end if %>
                   </span>
            
        </td>

	<td valign=top align="center" style="padding:3px 10px 10px 10px;">
        <input name="FM_slet_aid_<%=c %>" id="af_sl_<%=lcase(trim(oRec("fase")))%>_<%=fa%>" type="checkbox" value="1" />
        <input name="FM_slet_aid" type="hidden" value="<%=c %>" />
	<!--
	<a href="aktiv.asp?menu=job&func=slet&jobid=<%=jobid%>&jobnavn=<%=request("jobnavn")%>&id=<%=oRec("id")%>&rdir=<%=rdir%>"><img src="../ill/slet_16.gif" alt="Slet aktivitet" border="0"></a>
	-->
	</td>
	</tr>
	<%
	
	lastFaseSum = lastFaseSum + aktsumtot
	lastFaseTimer = lastFaseTimer + budgettimer
	c = c + 1
	x = 0
	fa = fa + 1
	oRec.movenext
	wend
	oRec.close
	
	call faSumlinje(lastFase, lastFaseSum, lastFaseTimer)
	
	
	'if antalstkialt <> 0 then
	'antalstkialt = antalstkialt
	'else
	'antalstkialt = 0
	'end if
	
	
	'*** Samlet værdi på aktiviteter ****' 
	strSQL3 = "SELECT sum(aktbudgetsum) AS aktbudgetsum, sum(budgettimer) AS akttimer FROM aktiviteter WHERE job = " & jobid &" ORDER BY budgettimer"
	oRec3.open strSQL3, oConn, 3
	if not oRec3.EOF then
		aktbudgetsamlet = oRec3("aktbudgetsum")
		akttfaktimtildelt = oRec3("akttimer")
	end if
	oRec3.close
	
	if len(aktbudgetsamlet) <> 0 then
	aktbudgetsamlet = aktbudgetsamlet
	else
	aktbudgetsamlet = 0
	end if
	
	
	if len(akttfaktimtildelt) <> 0 then
	akttfaktimtildelt = akttfaktimtildelt
	else
	akttfaktimtildelt = 0
	end if
	
	

	%>
	
	<tr bgcolor="lightpink">
		<td colspan="8" valign=top align=right style="padding:10px 10px 10px 10px;"><b>Total:</b>&nbsp;</td>
           
            
		 <td colspan=2 align=right style="padding:10px 10px 10px 10px;">
		 <table cellpadding=0 cellspacing=0 border=0 width=100%>
		 <tr>
		    <td>Forkalkuleret:</td>
		    <td align=right><div id="fasertimertot"><b><%=formatnumber(akttfaktimtildelt, 2) &"</b> timer"%></div></td>
		 </tr>
		 <tr>
		    <td>Realiseret:</td>
		    <td align=right><b><%=formatnumber(timerRealIalt, 2) %></b> timer</td>
		 </tr>
		 
		 </table>
		
		 </td>
		 <td>&nbsp;</td>
		<td align=right style="padding:10px 10px 10px 10px;">
		<div id="fasersumtot"><b><%=formatnumber(aktbudgetsamlet, 2) &"</b> "&basisValISO%></div>
            <input id="fatot_ialt" value="<%=faTot-1 %>" type="hidden" /> 
		</td>
		 
		<td colspan="2">
            &nbsp;</td>
	</tr>	
	<%if cint(SY_fastpris) = 1 AND _
    cint(SY_usejoborakt_tp) = 1 then
    opdjobvCHK = "CHECKED"
    else
    opdjobvCHK = ""
    end if
    
    if c <> 0 then
    c = c
    else
    c = 0
    end if
     %>
	<tr>
	    <td colspan=8 style="padding:20px 70px 20px 10px;">
           <input id="synctimertiljob" name="opdjobv" type="checkbox" <%=opdjobvCHK %> /><b>Synkroniser forkalkulerede timer på aktiviteter med forkalkuledere timer på job</b> (forkalk. timer fra aktiviterer overføres til job)<br />
          
            <br /><br />Aktiviteter på listen: <%=formatnumber(c,0) %></td>
		<td align=right colspan=6 style="padding:20px 70px 20px 10px;"><input type="submit" name="statusliste" value="Opdater liste >>"></td>
	</tr>
	
			

	</table>
            </form>
	
	<!-- table div -->
	</div>
	
	<br><br>
	<b>Funktioner:</b><br>
	<a href="aktiv.asp?menu=job&func=opdprogp&jobid=<%=jobid%>" class=vmenu target="_blank">klik her</a> for at opdatere projektgrupper på alle ovenstående aktiviteter,<br> således at de matcher de projektgrupper der er tildelt på jobbet.
	    <!--
        <br><br><a href="timereg_2006_fs.asp" class=vmenu><< Til timeregistrering</a>
	    <br><a href="jobs.asp?menu=job&shokselector=1" class=vmenu><< Til joboversigt</a>
        -->
	<br>
	<br>
	&nbsp
	</div>
	<%end select%>


<%
end if
%>
<!--#include file="../inc/regular/footer_inc.asp"-->
