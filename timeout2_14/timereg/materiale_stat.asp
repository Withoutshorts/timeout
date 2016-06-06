<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="../inc/regular/mat_func.asp"-->
<!--#include file="inc/isint_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->


<%


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	

    call smileyAfslutSettings()
	
	if len(request("hidemenu")) <> 0 then
	hidemenu = request("hidemenu")
	else
	hidemenu = 0
	end if
	
	thisfile = "materiale_stat.asp"
	func = request("func")
	id = request("id") '** Jobid
	print = request("print")
	
	if len(id) <>  0 then
	id = id
		else
			if len(request.cookies("matstat")("jobidmat")) <> 0 then
			id = request.cookies("matstat")("jobidmat")
			else
			id = 0
			end if
	end if
	Response.cookies("matstat")("jobidmat") = id
	
	'Response.Write " id:" & id
	
	session.lcid = 1030 'DK
	editok = 0
	
	if len(request("medid")) <> 0 then
	medid = request("medid")
	else
		if len(request.cookies("matstat")("medid")) <> 0 then
		medid = request.cookies("matstat")("medid")
		else
		medid = 0
		end if
	end if
	
	level = session("rettigheder")
	
	Response.cookies("matstat")("medid") = medid
	Response.cookies("matstat").expires = date + 1
	
	'*** Altid = 1 (Brug datointerval)
	fmudato = 1
	
	select case func
	case "afregnalle"
	
	if len(trim(request("afregnalle_id"))) <> 0 then
	afregnalle_id = request("afregnalle_id")
	else
	afregnalle_id = 0
	end if
	
	strSQL = "UPDATE materiale_forbrug SET afregnet = 1 WHERE bilagsnr = '"& afregnalle_id & "'"
	'Response.Write strSQL & "<br>"
	oConn.execute(strSQL)
	
	Response.redirect "materiale_stat.asp?hidemenu="&hidemenu
	
	
	case "opdafregnet"
	
	
	'Response.Write request("erafregnet") & "<br>"
	
	fids = split(trim(request("erafregnet")), ", ")
	for x = 0 to UBOUND(fids)
	
	thisval = trim(request("mfid_"&fids(x)&""))
	thisvalGK = trim(request("mfidgk_"&fids(x)&""))
	gkDate = year(now) & "/"& month(now) &"/"& day(now)
	
	'** kun hvis den ikke har været afregnet / udbewtalt før skal gkaf og gkdato ændres ***'
	godkendt = 0
	findes = 0
	strSQLfindes = "SELECT godkendt FROM materiale_forbrug WHERE id = "& fids(x)
	oRec.open strSQLfindes, oConn , 3
	if not oRec.EOF then
	
	godkendt = oRec("godkendt")
	findes = 1
	
	end if
	oRec.close
	
	if cint(godkendt) <> 1 then
	
	strSQL = "UPDATE materiale_forbrug SET afregnet = "& thisval &", godkendt = "& thisvalGK 
	    if thisvalGK = 1 then
	    strSQL = strSQL &", gkaf = '"& session("user") &"', gkdato = '"& gkDate &"'"
	    end if
	strSQL = strSQL &" WHERE id = "& fids(x)
	'Response.Write strSQL & "<br>"
	oConn.execute(strSQL)
	
	else
	'*** Opdater ikke timestamp for gokenels, samt godkendt af ***'
	strSQL = "UPDATE materiale_forbrug SET afregnet = "& thisval &", godkendt = "& thisvalGK &" WHERE id = "& fids(x)
	'Response.Write strSQL & "<br>"
	oConn.execute(strSQL)
	
	
	end if
	
	next
	
	'Response.end
	
	Response.redirect "materiale_stat.asp?hidemenu="&hidemenu
   
	
	
	
	case "dbred"
		
		
		antal = request("FM_antal")

        call erDetInt(antal)
        if isInt <> 0 then
        %>
	    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	    <%
	    errortype = 154
	    call showError(errortype)
        Response.end
        end if
        isInt = 0


        matkobspris = replace(request("FM_matkobspris"), ".", "") 
        matkobspris = replace(matkobspris, ",", ".")

        call erDetInt(matkobspris)
        if isInt <> 0 then
        %>
	    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	    <%
	    errortype = 154
	    call showError(errortype)
        Response.end
        end if
        isInt = 0


        matsalgspris = replace(request("FM_matsalgspris"), ".", "")
        matsalgspris = replace(matsalgspris, ",", ".")

        call erDetInt(matkobspris)
        if isInt <> 0 then
        %>
	    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	    <%
	    errortype = 154
	    call showError(errortype)
        Response.end
        end if
        isInt = 0

         if len(trim(request("FM_matnavn"))) <> 0 then
        matnavn = trim(request("FM_matnavn"))
        matnavn = replace(matnavn, "'", "")
        else
        matnavn = ""
        end if

        if matnavn = "" then
        %>
	    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	    <%
	    errortype = 154
	    call showError(errortype)
        Response.end
        end if
        


        if len(trim(request("FM_jobnr"))) <> 0 then
         jobnr = trim(request("FM_jobnr"))
        else
         jobnr = 0
        end if
                
         if jobnr <> 0 then
         jobnr = jobnr
         else
         jobnr = "XX999xcvsdrf3"
         end if

         jidFundet = 0
         jobidThis = 0
         strSQLjobnr = "SELECT id FROM job WHERE jobnr = '"& jobnr &"'"
       
         oRec3.open strSQLjobnr, oConn, 3
         if not oRec3.EOF then
         jidFundet = 1
            
            jobidThis = oRec3("id")

         end if
         oRec3.close         



      
        if jidFundet = 0 then
        %>
	    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	    <%
	    errortype = 166
	    call showError(errortype)
        Response.end
        end if
      
        if len(trim(request("matrecid"))) <> 0 then
		matfbrecordid = request("matrecid")
        else
        matfbrecordid = 0
        end if


		
            strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		    matid = 0
            matantallager = 0
            matvarenr = 0
            matantal_org = 0
            
            '*** Henter materiale id og antal ***'
            strSQLsel = "SELECT matantal, matid, matvarenr FROM materiale_forbrug WHERE id = " & matfbrecordid
            oRec.open strSQLsel, oConn, 3
            if not oRec.EOF then
            
            matvarenr = oRec("matvarenr")
            matid = oRec("matid")
            matantal_org = oRec("matantal")
            'matnavn = oRec("matnavn")

            end if
            oRec.close


            if matvarenr <> 0 then

             '*** Henter materiale lagerstatus ***'
            strSQLsel = "SELECT antal AS matantallager FROM materialer WHERE id = " & matid
            oRec.open strSQLsel, oConn, 3
            if not oRec.EOF then
            
            matantallager = oRec("matantallager")
            'matnavn = oRec("matnavn")

            end if
            oRec.close
		

            antal_org = matantal_org 'request("FM_antal_org")
            antalDiff = (antal - (antal_org))

            beregnet = (matantallager-(antalDiff))
            beregnet = replace(beregnet, ".", "")
            beregnet = replace(beregnet, ",", ".")

            'Response.write "matantallager: "& matantallager & "<br>"
            'Response.write "antal_org: "& antal_org & "<br>" 
            'Response.write "antal: "& antal & "<br>"
            'Response.write "Nyt lager: beregnet: "& beregnet & "<br>"

            'Response.end
            end if


            antal = replace(antal, ".","") 
            antal = replace(antal, ",",".")

        '***** SLET *****'
		if antal = 0 then
		    
		    
		    '** Opdaterer antal på lager ***
	        if matvarenr <> 0 then
	        strSQL2 = "UPDATE materialer SET antal = "& beregnet &" WHERE id = "&matid
	        oCOnn.execute(strSQL2)
	        end if
		    
		
		 strSQL = "DELETE FROM materiale_forbrug WHERE id = "& matfbrecordid 

         '*** Indsætter i delete historik ****'
         matnavn = replace(matnavn, "'", "")
	     call insertDelhist("mat", matfbrecordid, 0, matnavn, session("mid"), session("user"))

		else
		        
		    '** Opdaterer antal på lager ***
	        if matvarenr <> 0 then

            strSQL2 = "UPDATE materialer SET antal = '"& beregnet &"' WHERE id = "& matid
            'response.write strSQL2
            'response.flush
           
	        oCOnn.execute(strSQL2)
	        end if


            '*** ER der flyttet job ***'
		    
		
		strSQL = "UPDATE materiale_forbrug SET matantal = "& antal &", editor = '"& session("user")&"', "_
        &" dato = '"& strDato &"', matkobspris = "& matkobspris &", matsalgspris = "& matsalgspris &", "_
        &" jobid = "& jobidThis &", matnavn = '"& matnavn &"' WHERE id = "& matfbrecordid 
		end if
		'Response.write strSQL
		'Response.end
        
        oConn.execute(strSQL)	
		
		
		progrp_medarb = request("FM_radio_projgrp_medarb")
		progrp = request("FM_progrupper")
		thisMiduse = request("FM_medarb")
		kundeid = request("FM_kunde")
		kundeans = request("FM_kundeans")
		jobans = request("FM_jobans")
		visKundejobans = request("FM_kundejobans_ell_alle")
		jobid = request("FM_job")
        jobsogVal = request("FM_jobsog")

		
		response.redirect "materiale_stat.asp?id="&id&""_
		&"&FM_radio_projgrp_medarb="&progrp_medarb&""_
		&"&FM_progrupper="&progrp&""_
		&"&FM_medarb="&thisMiduse&""_
		&"&FM_medarb_hidden="&thisMiduse&""_
		&"&FM_kunde="&kundeid&""_
		&"&FM_kundeans="&kundeans&""_
		&"&FM_jobans="&jobans&""_
		&"&FM_kundejobans_ell_alle="&visKundejobans&""_
		&"&FM_job="&jobid&"&hidemenu="&hidemenu&"&FM_jobsog="&jobsogVal
		
		
		
	case else
	
	thisJobnr = 0
	%>
	
	
	
	
	
	
	
	
	<%if print <> "j"  then%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<script src="inc/matstat_jav.js"></script>


 
        <%

        call menu_2014()
	
	tdclass = ""
	tblwdt = 900
	tdwtd = 120
            
          siTop = 102
           siLeft = 90%>
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sideindhold" style="position:absolute; left:<%=siLeft%>px; top:<%=siTop%>px; visibility:visible;">
	
	<%else
        
   %>
       <!--#include file="../inc/regular/header_hvd_inc.asp"-->
     
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sideindhold" style="position:absolute; left:20px; top:20px; visibility:visible;">
	<%
	tdclass = "lille"
	tblwdt = 600
	tdwtd = 80
	
	
	end if '*Print%>

	
	
	
	
	<%
	
	if print <> "j" then
	
	
	
	Knap1_bg = "#3B5998"
	Knap4_bg = "#5C75AA"
	Knap3_bg = "#3B5998"
	Knap2_bg = "#3B5998"
	
	addTopPx = 30
	
	    call matregmenu(-20,0, Knap1_bg,Knap2_bg,Knap3_bg,Knap4_bg)
	else
	
	addTopPx = 0
	
	end if
	
	oimg = "ikon_matforbrug.png"
	oleft = 0
	otop = addTopPx
	owdt = 400
	oskrift = "Materialeforbrug / Udlæg's statistik"
	
    if media = "print" OR print = "j" then
    media = "print"
	'call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	end if
	
	
	
	
	
	
	
	'**************************************************
	'***** Faste Filter kriterier *********************
	'**************************************************
		
	
	'*** Medarbejdere / projektgrupper
	selmedarb = session("mid")
	
	'*** Job og Kundeans ***
	call kundeogjobans()
	
    'call medarbogprogrp("matstat")
	medarbSQlKri = ""
	kundeAnsSQLKri = ""
	jobAnsSQLkri = ""
	jobAns2SQLkri = ""
	
	if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
	progrp = 0
	end if
	
	'Response.Write "medid first: "& left(request("FM_medarb"), 1)
	'Response.end
	
	'*** Rettigheder på den der er logget ind **'
	medarbid = session("mid")
	
	if len(request("FM_medarb")) <> 0 OR func = "export" then
	
	    if left(request("FM_medarb"), 1) = "0" then
	        
            if media <> "print" then
	        thisMiduse = request("FM_medarb_hidden")
    	    else
    	    thisMiduse = request("FM_medarb")
    	    end if
    	
    	intMids = split(thisMiduse, ", ")
	    else
	    thisMiduse = request("FM_medarb")
	    intMids = split(thisMiduse, ", ")
	    end if
	
	else
	    
	    thisMiduse = session("mid") 
	    intMids = split(thisMiduse, ", ")
	   
	end if
	
	'Response.Write "thisMiduse:"& request("FM_medarb") & thisMiduse & "<br>"

	media = request("media")
	
	'**** Kundekri ***
	if len(request("FM_kunde")) <> 0 then
	kundeid = request("FM_kunde")
	else
	kundeid = 0
	end if
	
	'*** Kundeans ***
	'strKnrSQLkri = ""
	strKnrSQLkri = " OR jobknr = 0 "
	
	'*** finder udfra valgte projektgrupper og medarbejdere
	'medarbSQlKri 
	'kundeAnsSQLKri
	
			    for m = 0 to UBOUND(intMids)
			    
			     if m = 0 then
			     medarbSQlKri = "(m.mid = " & intMids(m)
			     kundeAnsSQLKri = "kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
			     jobAnsSQLkri = "jobans1 = "& intMids(m)  
			     jobAns2SQLkri = "jobans2 = "& intMids(m)
                 jobAns3SQLkri = "jobans3 = "& intMids(m) & " OR jobans4 = "& intMids(m) & " OR jobans5 = "& intMids(m)
			     else
			     medarbSQlKri = medarbSQlKri & " OR m.mid = " & intMids(m)
			     kundeAnsSQLKri = kundeAnsSQLKri & " OR kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
			     jobAnsSQLkri = jobAnsSQLkri & " OR jobans1 = "& intMids(m)  
			     jobAns2SQLkri = jobAns2SQLkri & " OR jobans2 = "& intMids(m)
                 jobAns3SQLkri = jobAns3SQLkri & " OR jobans3 = "& intMids(m) & " OR jobans4 = "& intMids(m) & " OR jobans5 = "& intMids(m)
			     end if
			     
			    next
			    
			    medarbSQlKri = medarbSQlKri & ")"
			    
			jobAnsSQLkri =  " ("& jobAnsSQLkri &")"
			jobAns2SQLkri =  "xx (" & jobAns2SQLkri &")"
			jobAns3SQLkri =  "xx (" & jobAns3SQLkri &")"
			
	
	'** Er key acc og kundeansvarlig valgt **?
	if cint(kundeans) = 1 then
	kundeAnsSQLKri = kundeAnsSQLKri
	else
	kundeAnsSQLKri = " Kundeans1 <> -1 AND kundeans2 <> -1 "
	end if
	
	if len(request("FM_ignorerperiode")) <> 0 then
	ignper = request("FM_ignorerperiode")
	else
	ignper = 0
	end if
	
	'***** Valgt job eller søgt på Job ****
	'** hvis Sog = Ja
	call jobsog()

	
	'*** Aftale ****
	if len(request("FM_aftaler")) <> 0 then ' AND jobid <= 0 AND len(jobSogVal) = 0 then
	    aftaleid = request("FM_aftaler")
	else
		aftaleid = 0
	end if
	
	
	
	'**** Alle SQL kri starter med NUL records ****
	'jobidFakSQLkri = " OR jobid = -1 "
	jobnrSQLkri = " OR tjobnr = '-1' "
	jidSQLkri = " OR id = -1 "
	seridFakSQLkri = " OR aftaleid = -1 "
	
	
	
	if len(request("viskunabnejob")) <> 0 then
	viskunabnejob = request("viskunabnejob")
	    
	    if viskunabnejob = 0 then
	    jost0CHK = "CHECKED"
	    jost1CHK = ""
	    else
	    jost1CHK = "CHECKED"
	    jost0CHK = ""
	    end if
	    
	else
	    if len(trim(request.cookies("stat")("viskunabnejob"))) <> 0 then
	    viskunabnejob = request.cookies("stat")("viskunabnejob")
	    
	            
	            if viskunabnejob = 0 then
                jost0CHK = "CHECKED"
                jost1CHK = ""
                else
                jost1CHK = "CHECKED"
                jost0CHK = ""
                end if
	            
	            
	    else
	    jost0CHK = "CHECKED"
	    jost1CHK = ""
	    viskunabnejob = 0
	    end if
	end if
	
	
	
	
	response.cookies("stat")("viskunabnejob") = viskunabnejob
	'************ slut faste filter var **************				
	
	
		
	
	
			
			
			'**** Fordeling pr job og medarb eller sum ***
			if len(request("FM_visprjob_ell_sum")) <> 0 then
			visprjob_ell_sum = request("FM_visprjob_ell_sum")
				
				select case cint(visprjob_ell_sum) 
				case 0 
				visprjob_ell_sum_chk0 = "CHECKED"
				visprjob_ell_sum_chk1 = ""
				visprjob_ell_sum_chk2 = ""
				case 2
				visprjob_ell_sum_chk0 = ""
				visprjob_ell_sum_chk1 = ""
				visprjob_ell_sum_chk2 = "CHECKED"
				case else
				visprjob_ell_sum_chk0 = ""
				visprjob_ell_sum_chk1 = "CHECKED"
				visprjob_ell_sum_chk2 = ""
				end select
				
				Response.Cookies("matstat")("vtype") = visprjob_ell_sum
				
			else
			    
			    if request.Cookies("matstat")("vtype") <> "" then
			    
			    visprjob_ell_sum = request.Cookies("matstat")("vtype")
			    
			        select case cint(visprjob_ell_sum) 
				    case 0 
				    visprjob_ell_sum_chk0 = "CHECKED"
				    visprjob_ell_sum_chk1 = ""
				    visprjob_ell_sum_chk2 = ""
				    case 2
				    visprjob_ell_sum_chk0 = ""
				    visprjob_ell_sum_chk1 = ""
				    visprjob_ell_sum_chk2 = "CHECKED"
				    case else
				    visprjob_ell_sum_chk0 = ""
				    visprjob_ell_sum_chk1 = "CHECKED"
				    visprjob_ell_sum_chk2 = ""
				    end select
			    
			    else
			        
                    visprjob_ell_sum = 2 'Som udgiftsbilag
			    
                
                visprjob_ell_sum_chk0 = "CHECKED"
			    visprjob_ell_sum_chk1 = ""
			    visprjob_ell_sum_chk2 = ""
			    end if
			    
			end if
			
			        select case cint(visprjob_ell_sum) 
				    case 0 
				    visningsTypVal = "Materialeforbr. / Udlæg fordelt på job"
				    case 2
				    visningsTypVal = "Udgiftsbilag"
				    case else
				    visningsTypVal = "Totalforbrug"
				    end select
	
	
	
	
	call filterheader_2013(addTopPx-5,0,890,oskrift)	

         
	
	%>
	
	<table cellspacing=0 cellpadding=0 border=0 width=100% bgcolor="#FFFFFF">
	
	<%if print <> "j" then%>
	<form action="materiale_stat.asp?hidemenu=<%=hidemenu %>" method="post">
	
	<%end if %>
	
	 <%
         
         call medkunderjob %>
	

	
	 </td>
	    </tr>
	
	
	
	
	
	<%
    'response.end    
        
    if print = "j" then
	dontshowDD = "1"
	%>
	<!--#include file="inc/weekselector_s.asp"--> <!-- b -->
	
	<tr><td><b>Periode:</b></td><td>
	<%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%>
	</td></tr>
	
	<%
	else
	%>
	<tr>
			<!-- Brug altid datointerval, FM_usedatokri = 1 -->
			<input type="hidden" name="FM_usedatokri" id="FM_usedatokri" value="1">
			<td valign=top><b>Periode:</b><br>
			<!--#include file="inc/weekselector_s.asp"--> <!-- b -->
			</td>
		
	<%end if%>
	
	
	<%
	'*** visningstype kriteriuer, bilagsnr, personlige mm.
	
	if len(trim(request("FM_bilagsnr"))) <> 0 then 
	        sogliste = trim(request("FM_bilagsnr"))
	        useSog = 1
	        else
	        sogliste = ""
	        useSog = 0
	        end if
	        %> 
    
    
    <%if len(request("FM_intkode")) <> 0 then
			intKode = request("FM_intkode")
			
			
			'Response.Write "intKode  "& intKode 
			    
			    select case intKode
			    case 1
			    ikSEL0 = ""
			    ikSEL1 = "SELECTED"
			    ikSEL2 = ""
			    ikSEL3 = ""
			    intKodeSQLkri = " AND intkode = 1 "
			    intKodeVal = "Intern" 
			    case 2
			    ikSEL0 = ""
			    ikSEL1 = ""
			    ikSEL2 = "SELECTED"
			    ikSEL3 = ""
			    intKodeSQLkri = " AND intkode = 2 "
			    intKodeVal = "Ekstern"
			    'case 3
			    'ikSEL0 = ""
			    'ikSEL1 = ""
			    'ikSEL2 = ""
			    'ikSEL3 = "SELECTED"
			    'intKodeSQLkri = " AND intkode = 3 "
			    'intKodeVal = "Personlig"
			    case else
			    ikSEL0 = "SELECTED"
			    ikSEL1 = ""
			    ikSEL2 = ""
			    ikSEL3 = ""
			    intKodeSQLkri = " AND intkode <> -1 "
			    intKodeVal = "Alle"
			    end select
			
			
			else
			intKode = 0
			ikSEL0 = "SELECTED"
			ikSEL1 = ""
			ikSEL2 = ""
			ikSEL3 = ""
			intKodeSQLkri = " AND intkode <> -1 "
			intKodeVal = "Alle"
			end if %>
            
            
            
              <%
           if len(trim(request("FM_intkode"))) <> 0 AND print <> "j" then
                if len(trim(request("showonlypers"))) <> 0 then
                showonlypers = 1
                showonlypersCHK = "CHECKED"
                else
                showonlypers = 0
                showonlypersCHK = ""    
                end if
           else
               
               if Request.Cookies("mat")("cshowonlypers") <> "" then
                    showonlypers = request.Cookies("mat")("cshowonlypers")
                    
                    if showonlypers = 1 then
                    showonlypersCHK = "CHECKED"
                    else
                    showonlypersCHK = ""
                    end if
                    
                else
                    showonlypers = 0
                    showonlypersCHK = ""
               end if
         
           end if 
           
           Response.Cookies("mat")("cshowonlypers") = showonlypers
           
           
           
            if len(trim(request("FM_intkode"))) <> 0 AND print <> "j" then
                if len(trim(request("showonlyikkeafreg"))) <> 0 then
                showonlyikkeafreg = 1
                showonlyikkeafregCHK = "CHECKED"
                else
                showonlyikkeafreg = 0
                showonlyikkeafregCHK = ""    
                end if
           else
               
               if Request.Cookies("mat")("cshowonlyikkeafreg") <> "" then
                    showonlyikkeafreg = request.Cookies("mat")("cshowonlyikkeafreg")
                    
                    if showonlyikkeafreg = 1 then
                    showonlyikkeafregCHK = "CHECKED"
                    else
                    showonlyikkeafregCHK = ""
                    end if
                    
                else
                    showonlyikkeafreg = 0
                    showonlyikkeafregCHK = ""
               end if
         
           end if 
           
           
           Response.Cookies("mat")("cshowonlyikkeafreg") = showonlyikkeafreg
           
           if len(trim(request("FM_intkode"))) <> 0 AND print <> "j" then
                if len(trim(request("showonlyikkegk"))) <> 0 then
                showonlyikkegk = 1
                showonlyikkegkCHK = "CHECKED"
                else
                showonlyikkegk = 0
                showonlyikkegkCHK = ""    
                end if
           else
               
               if Request.Cookies("mat")("cshowonlyikkegk") <> "" then
                    showonlyikkegk = request.Cookies("mat")("cshowonlyikkegk")
                    
                    if showonlyikkegk = 1 then
                    showonlyikkegkCHK = "CHECKED"
                    else
                    showonlyikkegkCHK = ""
                    end if
                    
                else
                    showonlyikkegk = 0
                    showonlyikkegkCHK = ""
               end if
         
           end if 
           
           
           Response.Cookies("mat")("cshowonlyikkegk") = showonlyikkegk
           

           
           
           '**** slut visning type kriterier ***'%>     
	
	
	<%if print <> "j" then %>
	<td valign=top>
	<table cellspacing=0 cellpadding=5 border=0 width=100%>
		<tr>
			<td valign=top>
			<b>Visningsttype:</b><br />
			<input type="radio" name="FM_visprjob_ell_sum" id="FM_visprjob_ell_sum0" value="0" <%=visprjob_ell_sum_chk0%> onclick="submit();"> Vis Materialeforbr. / Udlæg fordelt på job<br />
			<input type="radio" name="FM_visprjob_ell_sum" id="FM_visprjob_ell_sum1" value="1" <%=visprjob_ell_sum_chk1%> onclick="submit();"> Vis totalforbrug<br />
			<input type="radio" name="FM_visprjob_ell_sum" id="FM_visprjob_ell_sum2" value="2" <%=visprjob_ell_sum_chk2%> onclick="submit();"> Vis som udgiftsbilag
			</td>
			</tr>
        <tr>
			<td valign=top>
			
			<div id="ysiu" style="postion:relative; display:none; visibility:hidden; padding:10px; background-color:#F7F7F7;">
			<b>Yderligere søgekriterier ved udgiftsbilag's visning.</b><br />
		    
		    <br />Søg på bilags nummer:<br /><input id="FM_bilagsnr" name="FM_bilagsnr" size="20" type="text" value="<%=sogliste %>" />&nbsp;% = wildcard, <%=tsa_txt_325 %><br />
			<%=tsa_txt_231 %>: <br />
			
			
			
			<select name="FM_intkode" id="FM_intkode" style="width:55px; font-size:9px;">
		    <option value="0" <%=ikSEL0 %>><%=tsa_txt_076 %></option>
		    <option value="1" <%=ikSEL1 %>><%=tsa_txt_232 %></option>
		    <option value="2" <%=ikSEL2 %>><%=tsa_txt_233 %></option>
		    <!--
		    <option value="3" <%=ikSEL3 %>><=tsa_txt_234 %></option>
		    -->
		    
    		</select>
    		
    		<br />
                
                
           
            <input id="Checkbox2" name="showonlypers" id="showonlypers" type="checkbox" <%=showonlypersCHK %> /> <%=tsa_txt_320 %>
             <br /><input id="Checkbox1" name="showonlyikkeafreg" id="showonlyikkeafreg" type="checkbox" <%=showonlyikkeafregCHK %> /> <%=tsa_txt_322 %>
             <br /><input id="Checkbox3" name="showonlyikkegk" id="showonlyikkegk" type="checkbox" <%=showonlyikkegkCHK %> /> <%=tsa_txt_324 %>
             
             
                </div>  
		    
		    </td>
		   
		
		</tr>
	    </table>
        
        </td>
    
    
		
		<tr><td align=right colspan=3>
		   <input type="submit" value=" Kør statistik >> ">
			&nbsp;
		</td></tr>
		</form>
		</table>
		
	
	
	<%else %>
	
	<!-- for at undgå javafejl onload  -->
	<div id="ysiu" style="postion:relative; display:none; visibility:hidden; width:1px; height:1px;">
        <input id="FM_visprjob_ell_sum2" type="checkbox" />
	</div>		
	<!-- -->
	
	
	  <tr><td><b>Visningstype:</b></td><td><%=visningsTypVal %></td></tr>
	  
	   <tr><td valign=top><b>Medarbejder(e):</b></td><td>
	   <% for m = 0 to UBOUND(intMids)
	        
	        call meStamdata(intMids(m))
	        Response.Write meTxt & "<br>"
	   next
	   %></td></tr>
	  
	<%if visningsTypVal = "Udgiftsbilag" then %>
	<tr><td><b>Søgning på bilags nr:</b></td><td><%=sogliste %>&nbsp;</td></tr>
	<tr><td><b>Intern kode:</b></td><td><%=intKodeVal%></td></tr>
	<tr><td><b>Personlig:</b></td><td>
	<%if cint(showonlypers) <> 0 then %>
	<%=tsa_txt_320 %>
	<%else %>
	Alle
	<%end if %></td></tr>
	
	<tr><td><b>Afregnet:</b></td><td>
	<%if cint(showonlyikkeafreg) <> 0 then %>
	<%=tsa_txt_322 %>
	<%else %>
	Alle
	<%end if %></td></tr>
	
	
	<tr><td><b><%=tsa_txt_323 %>:</b></td><td>
	<%if cint(showonlyikkegk) <> 0 then %>
	<%=tsa_txt_324 %>
	<%else %>
	Alle
	<%end if %></td></tr>
	
	<%end if '**udgiftbilag %>
	
	<%end if %>
	
	</table>
	
	<!-- filter div -->
	</td></tr></table>
	</div><br /><br />
	
	
	

	
	<%
	
	
	 '*** valgte job ***
		call valgtejob()
		
		
		'**** Valgte aftaler *****
		call valgteaftaler()
				
	
	
	    '*** For at spare (trimme) på SQL hvis der vælges alle job alle kunder og vis kun for jobanssvarlige ikke er slået til ****
		'*** Og der ikke er søgt på jobnavn ***
		'if cint(kundeid) = 0 AND cint(jobid) = 0 AND cint(jobans) = 0 AND cint(kundeans) = 0 AND len(trim(jobSogVal)) = 0  then 
		if cint(kundeid) = 0 AND cint(jobid) = 0 AND cint(jobans) = 0 AND cint(jobans2) = 0 AND cint(jobans3) = 0 _
				 AND cint(kundeans) = 0 AND len(trim(jobSogVal)) = 0 AND cint(aftaleid) = 0 AND cint(segment) = 0 then 
				
				
			'jidSQLkri =  " OR id <> 0 "
			jobkri = " mf.jobid <> 0 "
			'seridFakSQLkri = " OR aftaleid <> 0 
			
		end if
	
	
		'**************** Trimmer SQL states ************************
		
		len_jobSQLkri = len(jobkri)
		right_jobSQLkri = right(jobkri, len_jobSQLkri - 3)
		jobSQLkri =  right_jobSQLkri
		
		'len_jidSQLkri = len(jidSQLkri)
		'right_jidSQLkri = right(jidSQLkri, len_jidSQLkri - 3)
		'jidSQLkri =  right_jidSQLkri
		
		'jidSQLKri = replace(jidSQLkri, "id", "aktiviteter.job")
		
		
		'*****************************************************************************************************
	
	
	
	'*** SQL datoer ***
	strStartDato = strAar&"/"&strMrd&"/"&strDag
	strSlutDato = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	
	
	if cint(fmudato) = 1 then
	strDatoKri = " AND mf.forbrugsdato BETWEEN '"& strStartDato &"' AND '"& strSlutDato &"'"
	else
	strDatoKri = ""
	end if
	
	orderby = request("orderby") 
	if len(request("orderby")) = 0 OR orderby = "dato" then
	orderbyKri = "mf.forbrugsdato DESC, mf.matnavn"
	else
	    select case orderby 
        case "matnr" 
	    orderbyKri = "mf.matvarenr"
        case "jobnr" 
	    orderbyKri = "k.kkundenavn, jobnr, mf.matvarenr"
	    case else
	    orderbyKri = "mf.matnavn"
	    end select
	end if
	
	if cint(jobans) = 1 OR cint(kundeans) = 1 then
	medarbSQlKri = " usrid <> 0 "
	else
	medarbSQlKri = medarbSQlKri
	end if
	
	if cint(jobid) <> 0 then
	jobKri = "mf.jobid = "& jobid &""
	else
	jobKri = jobKri
	end if
	
	'jobnrSQLkri = jobKri
	
	
	
	
	                if print <> "j" then
                    mainTableWth = "1104"
                	else
                	mainTableWth = "950"
                	end if
	
	
	select case visprjob_ell_sum 
	case 0 '** Matforbrug fordelt pr. job ***'
	
	
	                '**** Lukket for redigering *'
	                call ersmileyaktiv()
	                call licensStartDato()
                	
	                stDato = startDatoAar &"/"& startDatoMd &"/"& startDatoDag
	                slDato = year(now) &"/"& month(now) &"/"& day(now)
                	
                	
	                strSQL = "SELECT  mid FROM medarbejdere WHERE mansat <> '2' ORDER BY mid"
	                oRec.open strSQL, oConn, 3
	                while not oRec.EOF 
                	
	                call afsluger(oRec("mid"), stdato, sldato)
                	
	                oRec.movenext
	                wend
	                oRec.close
                	

       
                	
	               
		        tTop = 20 + addTopPx
	            tLeft = 0
	            tWdth = mainTableWth
            	
            	
	            call tableDiv(tTop,tLeft,tWdth)
	            %>
	                
	                <table cellspacing="0" cellpadding="2" border="0" bgcolor="#ffffff" width=100%>
	                
	                <tr bgcolor="#5582D2">
		                <td style="height:30px;">&nbsp;</td>
		                <td class=alt>Kunde<br /><a href="materiale_stat.asp?menu=stat&orderby=jobnr&id=<%=id%><%=link%>&hidemenu=<%=hidemenu%>&FM_jobsog=<%=jobSogVal%>" class=alt><u><b>Job</b></u></a><br />
                            Aktivitet</td>
		                <td class=alt><a href="materiale_stat.asp?menu=stat&orderby=dato&id=<%=id%><%=link%>&hidemenu=<%=hidemenu%>&FM_jobsog=<%=jobSogVal%>" class=alt><u><b>Dato</b></u></a> (Medarb.)</td>
		                <td class=alt style="padding-left:5px;"><b>Gruppe / lager</b> (Ava. %)</td>
		                <td class=alt style="padding-left:5px;"><a href="materiale_stat.asp?menu=stat&orderby=navn&id=<%=id%><%=link%>&hidemenu=<%=hidemenu%>&FM_jobsog=<%=jobSogVal%>" class=alt><u><b>Navn</b></u></a>&nbsp;&nbsp;
		                <a href="materiale_stat.asp?menu=stat&orderby=matnr&id=<%=id%><%=link%>&hidemenu=<%=hidemenu%>&FM_jobsog=<%=jobSogVal%>" class=alt><u><b>(Mat. nr.)</b></u></a></td>
		                <td class=alt><b>Antal</b></td>
		                
		                <td class=alt align=right><b>På lager</b><br />(heltal)</td>
		                <td class=alt align=right><b>Min. lager</b></td>
		                <td class=alt align=right><b>Købspris pr. stk</b></td>
		                <td class=alt align=right><b>Salgspris pr. stk.</b></td>
                        <td class=alt>&nbsp;</td>

		                <td class=alt><b>Valuta</b></td>
		                <td class=alt align=right><b>Kurs</b></td>
		                <td class=alt align=right><b>Købspris ialt</b></td>
		                <td class=alt align=right><b>Salgspris ialt</b></td>
		                <td class=alt><b>Valuta</b></td>
		                <td class=alt>&nbsp;</td>
		                <td>&nbsp;</td>
	                </tr>
	                <%
                	
	                strExport = "Jobnavn;Job Nr;Aktivitet;Kundennavn;Kunde Nr;Forbrugs dato;Medarbejder;Medarb. Nr;Initialer;Gruppe;Navn;Varenr;Antal;Enhed;Købspris pr. stk.;Salgspris pr. stk;Valuta; Kurs; Købspris ialt; Salgspris ialt; Valuta;"
	                
	                strSQL = "SELECT m.mnavn AS medarbejdernavn, mf.id, mf.matvarenr AS varenr, mg.navn AS gnavn, mf.matenhed AS enhed, "_
	                &" mf.matnavn AS navn, mf.matantal AS antal, mf.dato, mf.editor, mf.matid, "_
	                & "mf.matkobspris, mf.matsalgspris, mf.jobid, mf.matgrp, mf.usrid, mf.forbrugsdato, mg.av, "_
	                &" j.jobans1, j.jobans2, j.jobnavn, j.jobnr, k.kkundenavn, k.kkundenr, m.init, f.fakdato, "_
	                &" m.mnr, ma.antal AS palager, ma.minlager, "_
	                &" mf.valuta, mf.intkode, v.valutakode, mf.kurs AS mfkurs, j.risiko, a.navn AS aktnavn "_
	                &" FROM materiale_forbrug mf"_
	                &" LEFT JOIN materiale_grp mg ON (mg.id = mf.matgrp) "_
	                &" LEFT JOIN medarbejdere m ON (mid = mf.usrid) "_
	                &" LEFT JOIN job j ON (j.id = mf.jobid) "_
                    &" LEFT JOIN aktiviteter a ON (a.id = mf.aktid) "_
	                &" LEFT JOIN kunder k ON (kid = jobknr)"_
	                &" LEFT JOIN materialer ma ON (ma.id = mf.matid)"_
	                &" LEFT JOIN fakturaer f ON (f.jobid = mf.jobid AND f.faktype = 0)"_
	                &" LEFT JOIN valutaer v ON (v.id = mf.valuta) "_
	                &" WHERE ("&jobKri&") AND ("& medarbSQlKri &") "& strDatoKri &" GROUP BY mf.id ORDER BY "& orderbyKri	
                	
	                'response.write strSQL
	                'Response.flush
                	
	                x = 0
                	
	                strMatids = "0"
	                'strMatAntal = "0"
	                strAntalFM = ""
	                oRec.open strSQL, oConn, 3
	                while not oRec.EOF 
                	
	                strMatids = strMatids &","& oRec("matid")
	                'strMatAntal = strMatAntal &","& oRec("antal")
                					
                					
                					
					                editok = 0
					                '*** rediger rettigheder ***
					                if level = 1 then
					                editok = 1
					                else
							                if cint(session("mid")) = jobans1 OR cint(session("mid")) = jobans2 OR (cint(jobans1) = 0 AND cint(jobans2) = 0) then
							                editok = 1
							                end if
					                end if
                					
                					
                			    
			                    '*** Er uge alfsuttet af medarb, er smiley og autogk slået til
                                erugeafsluttet = instr(afslUgerMedab(oRec("usrid")), "#"&datepart("ww", oRec("forbrugsdato"),2,2)&"_"& datepart("yyyy", oRec("forbrugsdato")) &"#")
                                
                           
                              
                              
                                strMrd_sm = datepart("m", oRec("forbrugsdato"), 2, 2)
                                strAar_sm = datepart("yyyy", oRec("forbrugsdato"), 2, 2)
                                strWeek = datepart("ww", oRec("forbrugsdato"), 2, 2)
                                strAar = datepart("yyyy", oRec("forbrugsdato"), 2, 2)

                                if cint(SmiWeekOrMonth) = 0 then
                                usePeriod = strWeek
                                useYear = strAar
                                else
                                usePeriod = strMrd_sm
                                useYear = strAar_sm
                                end if

                
                                call erugeAfslutte(useYear, usePeriod, oRec("usrid"))
		        
		                        'Response.Write "smilaktiv: "& smilaktiv & "<br>"
		                        'Response.Write "SmiWeekOrMonth: "& SmiWeekOrMonth &" ugeNrAfsluttet: "& ugeNrAfsluttet & " tjkDag: "& tjkDag &"<br>"
		                        'Response.Write "autolukvdatodato: "& autolukvdatodato & "<br>"
		                        'Response.Write "tjkDag: "& tjkDag & "<br>"
		                        'Response.Write "autolukvdato: "& autolukvdato & "<br>"
		                        'Response.Write "erugeafsluttet:" & erugeafsluttet & "<br>"
		        
		                          call lonKorsel_lukketPer(oRec("forbrugsdato"), oRec("risiko"))
		         
                                'if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                                 if ( (( datepart("ww", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 0) OR (datepart("m", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 1 )) AND cint(ugegodkendt) = 1 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("forbrugsdato")) = year(now) AND DatePart("m", oRec("forbrugsdato")) < month(now)) OR _
                                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("forbrugsdato")) < year(now) AND DatePart("m", oRec("forbrugsdato")) = 12)) OR _
                                (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", oRec("forbrugsdato")) < year(now) AND DatePart("m", oRec("forbrugsdato")) <> 12) OR _
                                (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", oRec("forbrugsdato")) > 1))) OR cint(lonKorsel_lukketIO) = 1 then
                              
                                ugeerAfsl_og_autogk_smil = 1
                                else
                                ugeerAfsl_og_autogk_smil = 0
                                end if 
                				
				                if len(oRec("fakdato")) <> 0 then
	                            fakdato = oRec("fakdato")
	                            else
	                            fakdato = "01/01/2002"
	                            end if
                				
                				
				                if (ugeerAfsl_og_autogk_smil = 0 _
				                OR ugeerAfsl_og_autogk_smil = 1 AND level = 1) _
				                AND cdate(fakdato) < cdate(oRec("forbrugsdato")) then
				                editok = editok
				                else
				                editok = 0
				                end if
                					
                	
	                bgthis = "#ffffff"
	                %>
	                <form name="<%=oRec("id")%>" id="<%=oRec("id")%>" action="materiale_stat.asp?func=dbred&id=<%=id%>&matrecid=<%=oRec("id")%>&hidemenu=<%=hidemenu%>" method="post">
	                <input type="hidden" name="FM_radio_projgrp_medarb" id="FM_radio_projgrp_medarb" value="<%=progrp_medarb%>">
		                <input type="hidden" name="FM_progrupper" id="FM_progrupper" value="<%=progrp%>">
		                <input type="hidden" name="FM_medarb" id="FM_medarb" value="<%=thisMiduse%>">
		                <input type="hidden" name="FM_medarb_hidden" id="FM_medarb_hidden" value="<%=thisMiduse%>">
		                <input type="hidden" name="FM_kunde" id="FM_kunde" value="<%=kundeid%>">
		                <input type="hidden" name="FM_kundeans" id="FM_kundeans" value="<%=kundeans%>">
		                <input type="hidden" name="FM_jobans" id="FM_jobans" value="<%=jobans%>">
		                <input type="hidden" name="FM_kundejobans_ell_alle" value="<%=visKundejobans%>">
		                <input type="hidden" name="FM_job" id="FM_job" value="<%=jobid%>">
		                <input type="hidden" name="FM_visprjob_ell_sum" id="FM_visprjob_ell_sum" value="<%=visprjob_ell_sum%>">
                        <input type="hidden" name="FM_jobsog" id="Hidden1" value="<%=jobSogVal%>">
                        
		               
		                
		                
	                <%if x <> 0 then%>
	                <tr>
		                <td bgcolor="#ffffff" colspan="18"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
                    </tr>
	                <%end if%>
	                <tr bgcolor="<%=bgthis%>">
		                <td height=20>&nbsp;</td>
		                <td style="width:200px;">
		                <%=left(oRec("kkundenavn"), 15)%> (<%=oRec("kkundenr")%>)<br>
                            <b><%=oRec("jobnavn")%> </b>
                            <%if editok = 1 AND print <> "j" then%>
                            <input type="text" name="FM_jobnr" value="<%=oRec("jobnr")%>" style="width:60px; font-size:9px;" />
                            <%else %>
                            &nbsp;(<%=oRec("jobnr")%>)
                            <%end if %>
                            <br /><span style="color:#999999; font-size:9px;"><%=oRec("aktnavn") %></span>
		                </td>
                		
		                <%
		                strExport = strExport & "xx99123sy#z" & Chr(34) & oRec("jobnavn") & Chr(34) &";"& Chr(34) & oRec("jobnr") & Chr(34) &";"& Chr(34) & oRec("aktnavn") & Chr(34) &";"& Chr(34) & oRec("kkundenavn") & Chr(34) &";" & Chr(34) &oRec("kkundenr") & Chr(34) &";"
		               
                        
                 
                        %>
                		
		                <td style="white-space:nowrap; width:150px;">
		                <%if len(oRec("forbrugsdato")) <> 0 then%>
		                <%=formatdatetime(oRec("forbrugsdato"), 1)%>
		                <%
		                strExport = strExport & Chr(34) & formatdatetime(oRec("forbrugsdato"), 2) & Chr(34) &";" 
		                else
		                strExport = strExport &";"  
		                end if
		                %>
		                <br><span style="font-size:10px;"><%=oRec("medarbejdernavn")%> (<%=oRec("mnr")%>) - <%=oRec("init")%></span><br />
                        <span style="font-size:8px; color:#999999;"><%=oRec("editor")%></span></td>
                		
		                <%
		                strExport = strExport & Chr(34) & oRec("medarbejdernavn")& Chr(34) &";"& Chr(34) & oRec("mnr") & Chr(34) &";"& Chr(34) & oRec("init") & Chr(34) &";"
		                %>
                		
		                <td style="padding-left:5px;">
		                  <span style="color:#999999;"><%if len(oRec("gnavn")) <> 0 then %>
		                <%=oRec("gnavn")%> (<%=oRec("av") %> %)
		                <%end if %></span>
		                &nbsp;
		                </td>
		                <td style="padding-right:10px; padding-left:5px; width:210px; white-space:nowrap;">
                            
                              <%if editok = 1 AND print <> "j" AND oRec("varenr") = "0" then%>
                            <input type="text" name="FM_matnavn" value="<%=oRec("navn")%>" style="width:180px; font-size:9px;" />
                            <%else %>
                             <input type="hidden" name="FM_matnavn" value="<%=oRec("navn")%>" />
                            <%=oRec("navn")%>
                            <%end if %>

                            <%if oRec("varenr") <> "0" AND len(trim(oRec("varenr"))) <> 0 then%>
                            &nbsp;(<%=oRec("varenr")%>)
                            <%end if %>

                           </td>
                		
		                <%
		                strExport = strExport & Chr(34) & oRec("gnavn")&" ("& oRec("av") &"%)" & Chr(34) &";" & Chr(34) &oRec("navn") & Chr(34) & ";" & Chr(34) & oRec("varenr") & Chr(34) &";"
		                %>
                		
		                <%if editok = 1 AND print <> "j" then%>
		                <td><input type="text" name="FM_antal" id="FM_antal" size="2" style="font-size:9px;" value="<%=oRec("antal")%>">&nbsp;<%=oRec("enhed")%>&nbsp;

		                </td>
		                
		                <%else%>
		                <td><%=oRec("antal")%>&nbsp;<%=oRec("enhed")%></td>
		                
		                <%end if%>
                		
		                <td align=right><%=oRec("palager")%>&nbsp;</td>
		                <td align=right><%=oRec("minlager")%>&nbsp;</td>
                		
		                <%
		                strExport = strExport & Chr(34) & oRec("antal") & Chr(34) &";" & Chr(34) & oRec("enhed") & Chr(34) &";"
		                
                        
                        '** priser 2/3 decimaler ***
                        
                        matkobspris = oRec("matkobspris")
                        matsalgspris = oRec("matsalgspris")
                        %>
                		

                        <%if editok = 1 AND print <> "j" then%>
                        <td align=right style="padding-right:3px;"><input type="text" name="FM_matkobspris" id="Text1" size="2" style="font-size:9px; width:80px;" value="<%=matkobspris%>"></td>
                        <td align=right style="padding-right:3px;"><input type="text" name="FM_matsalgspris" id="Text2" size="2" style="font-size:9px; width:80px;" value="<%=matsalgspris%>"></td>
                        <td><input type="image" src="../ill/soeg-knap.gif"></td>
                        <%else %>
		                <td align=right style="padding-right:3px;"><b><%=matkobspris%></b></td>
		                <td align=right style="padding-right:3px;"><%=matsalgspris%></td>
                        <td>&nbsp;</td>
                        <%end if %>


		                <td align=center><%=oRec("valutakode") %></td>
		                
		            <td align=right style="padding-right:5px;"><%=oRec("mfkurs") %></td>
            		<td align=right style="padding-right:3px;">
		                <%kobsprisialt = formatnumber(matkobspris/1 * oRec("antal")/1 * oRec("mfkurs")/100 , 2) %>
		                <b><%=kobsprisialt %></b></td>
		                <%
		                strExport = strExport & matkobspris &";"
		                kprisTot = kprisTot + (kobsprisialt/1)
		                %>
		              
		              <td align=right style="padding-right:3px;">
		                <%salgsprisialt = formatnumber(matsalgspris/1 * oRec("antal")/1 * oRec("mfkurs")/100 , 2) %>
		                <%=salgsprisialt %>
		                <%
		                strExport = strExport & matsalgspris &";"& oRec("valutakode") & ";" & oRec("mfkurs") & ";"& kobsprisialt &";"&salgsprisialt&";"& basisValISO &";"
		                sprisTot = sprisTot + (salgsprisialt/1)
		                %>
		                </td>
		                <td align=center><%=basisValISO %></td>
            		
                		
		                <td>&nbsp;</td>
		                <td>&nbsp;</td>
	                </tr>
	                 <tr>
		                <td style="border-bottom:1px #CCCCCC solid;" colspan="18"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	                </tr>
	                
	                </form>
	                <% 
                	
	                strAntalFM = strAntalFM &"<input id='FM_antal_b_"&oRec("matid")&"' name='FM_antal_b_"&oRec("matid")&"' type=hidden value='"&oRec("antal")&"' />" 
                	
                    'if session("mid") = 1 then
                    'Response.write strExport & "<br>"
                    'end if 

                    Response.flush
	                x = x + 1
	                oRec.movenext
	                wend
                	
	                if x = 0 then 
	                %>
	              
	                <tr bgcolor="#ffffff">
		                <td height=50>&nbsp;</td>
		                <td colspan=16><font color="red"><b>!</b></font>&nbsp;Der er ikke registreret materialer der matcher de valgte søgekriterier.</td>
		                <td>&nbsp;</td>
	                </tr>



	                <%
                        
                        
                        
                        
                        else%>
	                <tr bgcolor="#ffdfdf">
		                <td height="50">&nbsp;</td>
		                <td colspan=12 align=right><b>Total:</b></td>
		                <td align=right>
		                <%if len(kprisTot) <> 0 then%>
		                <b><%=formatnumber(kprisTot, 2)%></b>
		                <%end if%></td>
                		
                		
		                <td align=right>
		                <%if len(sprisTot) <> 0 then%>
		                <b><%=formatnumber(sprisTot, 2)%></b>
		                <%end if%></td>
                		
                		
		                <td><b><%=basisValISO %></b></td>
		                <td>
                            &nbsp;</td>
		                <td>&nbsp;</td>
	                </tr>
	                <%end if%>	
	               
	                </table>
	                </div>
	
	<%
	case 1  '** Total forbrug ***'
	%>
		
		        <% 
		        tTop = 20 + addTopPx
	            tLeft = 0
	            tWdth = 880
            	
            	
	            call tableDiv(tTop,tLeft,tWdth)
	            %>
	            
		        <table cellspacing=0 cellpadding=2 width=100% border=0>
		        <tr bgcolor="#5582d2">
			        <td style="padding:5px;" class=alt><b>Lager / Gruppe</b> (Ava. %)</td>
			        <td style="padding:5px;" class=alt><b>Navn</b></td>
			        <td style="padding:5px;" class=alt><b>Varenr</b></td>
			        <td style="padding:5px;" class=alt align=right><b>Antal brugt i per.</b></td>
			        <td style="padding:5px;" class=alt align=right><b>På lager</b></td>
			        <td style="padding:5px;" class=alt align=right><b>Min. lager</b></td>
		        </tr>
        		
        		
		        <%
        		
		        strExport = "Gruppe;Navn;Varenr;Antal;Enhed;På lager;Min. lager;"
        	    
        	    medarbSQlKri = replace(medarbSQlKri, "m.mid", "mf.usrid")
        		
		        strSQL = "SELECT mf.id, mf.matvarenr AS varenr, mg.navn AS gnavn, mf.matenhed AS enhed, "_
		        &" mf.matnavn AS navn, sum(mf.matantal) AS antal, mf.dato, mf.editor, "_
		        &" mf.matkobspris, mf.matsalgspris, mf.jobid, mf.matgrp, mf.usrid, mf.forbrugsdato, "_
		        &" m.antal AS palager, m.minlager, m.id, mf.matid, mg.av "_
		        &" FROM materiale_forbrug mf"_
		        &" LEFT JOIN materialer m ON (m.id = mf.matid) "_
		        &" LEFT JOIN materiale_grp mg ON (mg.id = mf.matgrp) "_
		        &" WHERE ("&jobKri&") AND ("& medarbSQlKri &") "& strDatoKri &" GROUP BY mf.navn, mf.matgrp, mf.matgrp ORDER BY mg.gnavn, mf.matnavn"
        		
		        'Response.write strSQL
		        'Response.flush
		        x = 0
		        oRec.open strSQL, oConn, 3
		        while not oRec.EOF
        		
        		
		        'if oRec("minlager") > oRec("palager") then
		        'bg = "#ffdfdf"
		        'else
		        'bg = "#ffffff"
		        'end if		
		        
		        select case right(x,1)
		        case 0,2,4,6,8
		        bgcol = "#ffffff"
		        case else
		        bgcol = "#EFf3FF"
		        end select
		        			
        								
		        %>
		        <tr bgcolor="<%=bgcol %>">
			        <td height=20 style="border-top:1px #cccccc solid; padding-left:5px;">	
	            <%if len(oRec("gnavn")) <> 0 then %>
		        <%=oRec("gnavn")%> (<%=oRec("av") %> %)
		        <%end if %>
		        &nbsp;</td>
			        <td style="border-top:1px #cccccc solid; padding-left:5px;" class="<%=tdclass%>"><%=oRec("navn")%></td>
			        <td style="border-top:1px #cccccc solid; padding-left:5px;" class="<%=tdclass%>"><%=oRec("varenr")%></td>
			        <td style="border-top:1px #cccccc solid; padding-right:5px;" align=right class="<%=tdclass%>"><%=oRec("antal")%>&nbsp;
			        <%if len(oRec("enhed")) <> 0 then%>
			        <%enh = oRec("enhed")%>
			        <%else 
			        enh = "Stk."
			        end if %>
        			
			        <%=enh%></td>
        			
			        <%if oRec("varenr") <> "0" then %>
			        <td style="border-top:1px #cccccc solid; padding-right:5px;" align=right class="<%=tdclass%>"><%=oRec("palager")%>&nbsp;<%=oRec("enhed")%>
        			
			        <%if oRec("minlager") > oRec("palager") then
			        Response.write "&nbsp;<font class=roed><b>!</b>"
			        end if
			        %>
			        </td>
			        <td style="border-top:1px #cccccc solid; padding-right:5px;" align=right><%=oRec("minlager")%>&nbsp;<%=oRec("enhed")%></td>
		            <%else %>
		            <td style="border-top:1px #cccccc solid; padding-right:5px;" align=right class="<%=tdclass%>">-</td>
			        <td style="border-top:1px #cccccc solid; padding-right:5px;" align=right>-</td>
		            <%end if %>
		        </tr>
		        <%
		        strExport = strExport & "xx99123sy#z" & oRec("gnavn") &";"&oRec("navn")&";"&oRec("varenr")&";"&oRec("antal")&";"&enh&";"&oRec("palager")&";"&oRec("minlager")&";"
        		
        		Response.flush
		        x = x + 1
		        oRec.movenext
		        wend
        		
		        oRec.close
        		
        		
		        if x = 0 then%>
        		
		        <tr bgcolor="#ffffff">
			        <td colspan=8 height=50 style="padding:20px;">
			        <font color="red"><b>!</b></font>&nbsp;Der er ikke registreret materialer der matcher de valgte søgekriterier.</td>
		        </tr>
		        <%end if%>
		        <tr bgcolor="#5582d2">
			        <td colspan=8>&nbsp;</td>
		        </tr>
		        </table>
		        </div>
        	
	<% 
	case 2 '*** Som udgiftsbilag ***'
	
	
	            tTop = 20 + addTopPx
	            tLeft = 0
	            tWdth = mainTableWth
            	
            	
	            call tableDiv(tTop,tLeft,tWdth)
            	
	            %>
	            <table width=100% cellspacing=0 cellpadding=2 border=0>
            		<form method="post" action="materiale_stat.asp?func=opdafregnet&hidemenu=<%=hidemenu %>">
		            <tr bgcolor="#5582d2">
            		       
		                <td class=alt><b><%=tsa_txt_230 %></b></td>
		                
		                
		                <td class=alt><b><%=tsa_txt_231%></b><br />
		                <%=tsa_txt_241 %></td>
		                <td class=alt><b><%=tsa_txt_183 %></b></td>
            			
            			<td class=alt><b><%=tsa_txt_244 %></b><br />
            			<%=tsa_txt_243 %><br />
                            Aktivitet
            			</td>
            			
            			<td class=alt><%=tsa_txt_077%></td>
            			
            			
            			
			            <td align=right class=alt><b><%=tsa_txt_202 %></b></td>
			            <td class=alt><b><%=tsa_txt_245%></b></td>
			            <td style="padding:2px 5px 2px 5px;" class=alt><b><%=tsa_txt_209 %></b></td>
			            <td align=center style="padding:2px 5px 2px 5px;" class=alt><b><%=tsa_txt_217 %></b><br />
			            <%=right(tsa_txt_263, 6) %></td>
            			
            			
            			
            			
			            <td align=right style="padding:2px 5px 2px 5px;" class=alt><b><%=tsa_txt_219 %></b></td>
			            <td align=right style="padding:2px 5px 2px 5px;" class=alt><b><%=tsa_txt_246%></b></td>
			            <td align=right style="padding:2px 5px 2px 5px;" class=alt><b><%=tsa_txt_247%></b></td>
            			<td align=right style="padding:2px 5px 2px 5px;" class=alt><b><%=tsa_txt_248%></b></td>
            			<td align=right style="padding:2px 5px 2px 5px;" class=alt><b><%=tsa_txt_246%></b></td>
            			
            			
			            <td class=alt><%=tsa_txt_221 %></td>
			            <td class=alt><%=tsa_txt_323 %>?<br />
			            <a href="#gkja" name"gkja" class=alt onclick="gkAlly(1)"><u>Ja</u></a> 
			            | <a href="#gkja" name"gkja" class=alt onclick="gkAlly(0)"><u>Nej</u></a></td>
			            <td class=alt><%=tsa_txt_253 %>
			            <%if level = 1 then %>
			            <br />
			            <a href="#gkja" name"gkja" class=alt onclick="afAlly(1)"><u>Ja</u></a> 
			            | <a href="#gkja" name"gkja" class=alt onclick="afAlly(0)"><u>Nej</u></a>
			            <%end if %>
			            </td>
            			
            			
		            </tr>
	            <%
	            
	            strExport = strExport & tsa_txt_230 &";"& tsa_txt_231 &";"& tsa_txt_241 &";"& tsa_txt_183 &";"_
	            &""&tsa_txt_243&";"&tsa_txt_249&";"&tsa_txt_244&";"&tsa_txt_250&";Aktivitet;"& tsa_txt_077 &";"&tsa_txt_321&";"& tsa_txt_202 &";"&tsa_txt_245&";"& tsa_txt_209 &";"&tsa_txt_217&";"&tsa_txt_219&";"_
	            &""&tsa_txt_246&";"&tsa_txt_247&";"&tsa_txt_248&";"&tsa_txt_246&";"&tsa_txt_323&";"&tsa_txt_253&";"  
	            ' "xx99123sy#z" 
	            
	            
	            if useSog = 1 then
	            sqlWh = "mf.bilagsnr LIKE '"& sogliste &"'"
	            strDatoKri = ""
	            else
	            sqlWh = medarbSQlKri '"usrid = "& usemrn 
	            end if
	            
	            if showonlypers = 1 then
	            strSQLper = " AND personlig = 1"
	            else
	            strSQLper = ""
	            end if
            	
            	
            	if showonlyikkeafreg = 1 then
            	SQLafr = " AND mf.afregnet = 0"
            	else
            	SQLafr = ""
            	end if
            	
            	
            	if showonlyikkegk = 1 then
            	SQLgk = " AND godkendt = 0"
            	else
            	SQLgk = ""
            	end if
            	
            	totsum = 0
            	totsumNone = 0
		        totsumInt = 0
		        totsumEks = 0
		        totsumPers = 0 
            	
	            strSQLmat = "SELECT m.mnavn AS medarbejdernavn, mf.id AS mfid, mf.matvarenr AS varenr, mg.navn AS gnavn, mf.matenhed AS enhed, "_
	            &" mf.matnavn AS navn, mf.matantal AS antal, mf.dato, mf.editor, "_
	            & "mf.matkobspris, mf.matsalgspris, mf.jobid, mf.matgrp, "_
	            &" mf.usrid, mf.forbrugsdato, j.id, j.jobnr, j.jobnavn, j.risiko, "_
	            &" mg.av, f.fakdato, k.kkundenavn, mf.afregnet, "_
	            &" k.kkundenr, mf.valuta, mf.intkode, mf.bilagsnr, v.valutakode, mf.kurs AS mfkurs, "_
	            &" mf.personlig, m.init, m.mnr, mf.godkendt, mf.gkaf, m.mid, a.navn AS aktnavn "_
	            &" FROM materiale_forbrug mf"_
	            &" LEFT JOIN materiale_grp mg ON (mg.id = matgrp) "_
	            &" LEFT JOIN medarbejdere m ON (mid = usrid) "_
	            &" LEFT JOIN job j ON (j.id = mf.jobid) "_
                &" LEFT JOIN aktiviteter a ON (a.id = mf.aktid) "_
	            &" LEFT JOIN fakturaer f ON (f.jobid = mf.jobid AND f.faktype = 0) "_
	            &" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_
	            &" LEFT JOIN valutaer v ON (v.id = mf.valuta) "_
	            &" WHERE ("& sqlWh &") "& strDatoKri &" AND "&jobKri&" "& intKodeSQLkri &" "& strSQLper &" "& SQLafr &" "& SQLgk &" GROUP BY mf.id ORDER BY mf.forbrugsdato DESC, f.fakdato DESC, mf.id DESC" 	
            	
	            'mf.jobid = "& id &"
	            'response.write strSQLmat
	            'Response.flush
            	
            	
            	
	            x = 0
	            oRec.open strSQLmat, oConn, 3
	            while not oRec.EOF
            	 
            	 
	             if len(oRec("fakdato")) <> 0 then
	             fakdato = oRec("fakdato")
	             else
	             fakdato = "01/01/2002"
	             end if
            	 
	             if oRec("mfid") = cint(lastId) then
	             bgthis = "#FFFF99"
	             else
	                 select case right(x, 1) 
	                 case 0,2,4,6,8
	                 bgthis = "#EFf3FF"
	                 case else
	                 bgthis = "#FFFFff"
	                 end select
	             end if
	             
	             strExport = strExport & "xx99123sy#z"
	             
	             %>
	            <tr bgcolor="<%=bgthis %>">
            	    
	                <td style="border-bottom:1px #cccccc solid; white-space:nowrap;">
                    <%if len(oRec("bilagsnr")) <> 0 then%>
		            <%=oRec("bilagsnr") %> 
		            <%end if %>
            	        &nbsp;
                        </td>
                        
                        <%
                          strExport = strExport & oRec("bilagsnr") &";"
	           
                        %>
                        
                        <td style="border-bottom:1px #cccccc solid; width:80px;">
                        <%select case oRec("intkode")
		            case 0
	                intKodeTxt = "-"
		            case 1
		            intKodeTxt = tsa_txt_239
		            case 2
		            intKodeTxt = tsa_txt_240
		            'case 3
		            'intKode = tsa_txt_241
		            end select %>
            		
		            
                        
                        <%if intKodeTxt <> "-" then %>
		                <%=intKodeTxt%><br />
		                <%end if %>
                            
            	    
            	            <%
                          strExport = strExport & intKodeTxt &";"
	           
                        %>
                        
                        <%if oRec("personlig") <> 0 then %>
		                    <font class=megetlillesilver><%=tsa_txt_241 %></font>
		                    <%
		                    
		                    end if %>
		                    
		                  
		                    
		<%
                          strExport = strExport & oRec("personlig") &";"
	           
                        %>
		
		</td>
            	    
            	    
		            <td style="border-bottom:1px #cccccc solid; width:80px;">
		            <%if len(oRec("forbrugsdato")) <> 0 then%>
		            <%=formatdatetime(oRec("forbrugsdato"), 2)%>
		            <%end if%>
		           
	                </td>
	                 <%
                          strExport = strExport & formatdatetime(oRec("forbrugsdato"), 2) &";"
	           
                        %>
	                
	                
	                <td style="border-bottom:1px #cccccc solid; width:180px;"><%=oRec("kkundenavn") %> (<%=oRec("kkundenr") %>)<br />
                        <b><%=oRec("jobnavn")%></b>&nbsp;(<%=oRec("jobnr")%>)<br />
                        <span style="color:#999999; font-size:9px;"><%=oRec("aktnavn") %></span>
	                </td>
            		
	                
	                <%
                          strExport = strExport & oRec("kkundenavn") & ";"& oRec("kkundenr") &";" & oRec("jobnavn") & ";" & oRec("jobnr") & ";"& oRec("aktnavn") &";"
	           
                        %>
	                
	                <td style="border-bottom:1px #cccccc solid; width:120px;"><%=left(oRec("medarbejdernavn"), 15)%> 
	                <%if len(trim(oRec("init"))) <> 0 then %>
	                &nbsp;[<%=oRec("init") %>]
	                <%end if%>
	                </td>
            	     
            	     <%
                          strExport = strExport & oRec("medarbejdernavn") &";"& oRec("init") &";"& oRec("antal") &";"
	           
                        %>
            	     
		            <td align=right style="padding-right:5px; border-bottom:1px #cccccc solid;" class=lille><b><%=oRec("antal")%></b>&nbsp;</td>
		            
		            <td style="border-bottom:1px #cccccc solid;"> <%if len(oRec("enhed")) <> 0 then
		            enh = oRec("enhed")
		            else
		            enh = tsa_txt_222 '"Stk."
		            end if %>
            		
		            <%=enh%></td>
            	    
            	   
            	    
	                <td style="border-bottom:1px #cccccc solid;"><b><%=oRec("navn")%></b>&nbsp;</td>
		            <td align=right style="border-bottom:1px #cccccc solid; padding-right:5px;"><%=oRec("varenr")%>
		            <br /><%=oRec("gnavn") %></td>
            	    
            	     <%
            	    strNavn = replace(oRec("navn"), "'", "")
            	    strNavn = replace(strNavn, chr(34), "")
                
            	    
                          strExport = strExport & enh &";"& strNavn &";" & oRec("varenr") & ";" & formatnumber(oRec("matkobspris"), 2) & ";"& oRec("valutakode") &";"& oRec("mfkurs") &";"
	           
                        %>
		            
		            <td align=right style="border-bottom:1px #cccccc solid; padding-right:5px;">
	                <b><%=formatnumber(oRec("matkobspris"), 2) %></b>
		            </td>
		            
		            <td align=right style="border-bottom:1px #cccccc solid; padding-right:5px;"><%=oRec("valutakode") %></td>
            		
            		<td align=right style="border-bottom:1px #cccccc solid; padding-right:5px;"><%=oRec("mfkurs") %></td>
            		<td align=right style="border-bottom:1px #cccccc solid; padding-right:5px;">
            		<%udlgSum = formatnumber(oRec("matkobspris")/1 * oRec("antal")/1 * oRec("mfkurs")/100 , 2) %>
            		<b><%=udlgSum %></b>
            		</td>
            		
            		<%
                          strExport = strExport & udlgSum &";"& basisValISO &";"
	           
                        %>
            	     
            		
            		<%
            		
            		
            		'*** Totaler ***'
            		
            		
            		select case oRec("intkode")
		            case 0
	                '** Ingen
            		totsumNone = totsumNone + udlgSum 
		            case 1
		            '** Intern
		            totsumInt = totsumInt + udlgSum 
		            case 2
		            '** Ekster
		            totsumEks = totsumEks + udlgSum
		            end select 
            		
            		
            		
            		'** Perso
            		if oRec("personlig") = 1 then
            		totsumPers = totsumPers + udlgSum 
            		end if
            		
            		
            		'** Ialt 
            		totsum = totsum + udlgSum %>
            		
            		<td align=right style="border-bottom:1px #cccccc solid; padding-right:5px;"><%=basisValISO %></td>
            		
	                <td style="border-bottom:1px #cccccc solid;">
            	    
	                <%
	                   '*** Er uge alfsuttet af medarb, er smiley og autogk slået til
                            erugeafsluttet = instr(afslUgerMedab(oRec("mid")), "#"&datepart("ww", oRec("forbrugsdato"),2,2)&"_"& datepart("yyyy", oRec("forbrugsdato")) &"#")
                            
                            'Response.Write "erugeafsluttet --" & erugeafsluttet  &"<br>"
                            'Response.flush
                          


                                strMrd_sm = datepart("m", oRec("forbrugsdato"), 2, 2)
                                strAar_sm = datepart("yyyy", oRec("forbrugsdato"), 2, 2)
                                strWeek = datepart("ww", oRec("forbrugsdato"), 2, 2)
                                strAar = datepart("yyyy", oRec("forbrugsdato"), 2, 2)

                                if cint(SmiWeekOrMonth) = 0 then
                                usePeriod = strWeek
                                useYear = strAar
                                else
                                usePeriod = strMrd_sm
                                useYear = strAar_sm
                                end if

                
                                call erugeAfslutte(useYear, usePeriod, oRec("mid"))
		        
		                        'Response.Write "smilaktiv: "& smilaktiv & "<br>"
		                        'Response.Write "SmiWeekOrMonth: "& SmiWeekOrMonth &" ugeNrAfsluttet: "& ugeNrAfsluttet & " tjkDag: "& tjkDag &"<br>"
		                        'Response.Write "autolukvdatodato: "& autolukvdatodato & "<br>"
		                        'Response.Write "tjkDag: "& tjkDag & "<br>"
		                        'Response.Write "autolukvdato: "& autolukvdato & "<br>"
		                        'Response.Write "erugeafsluttet:" & erugeafsluttet & "<br>"
		        
		                           call lonKorsel_lukketPer(regdato, oRec("risiko"))
		         
                             'if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                            if ( (( datepart("ww", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 0) OR (datepart("m", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 1 )) AND cint(ugegodkendt) = 1 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                            (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("forbrugsdato")) = year(now) AND DatePart("m", oRec("forbrugsdato")) < month(now)) OR _
                            (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("forbrugsdato")) < year(now) AND DatePart("m", oRec("forbrugsdato")) = 12)) OR _
                            (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", oRec("forbrugsdato")) < year(now) AND DatePart("m", oRec("forbrugsdato")) <> 12) OR _
                            (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", oRec("forbrugsdato")) > 1))) OR cint(lonKorsel_lukketIO) = 1 then
                          
                            ugeerAfsl_og_autogk_smil = 1
                            else
                            ugeerAfsl_og_autogk_smil = 0
                            end if 
            				
            				
            				
				            if (ugeerAfsl_og_autogk_smil = 0 _
				            OR ugeerAfsl_og_autogk_smil = 1 AND level = 1) _
				            AND cdate(fakdato) < cdate(oRec("forbrugsdato")) AND print <> "j" then
            				
			            %>
	                &nbsp;<a href="materialer_indtast.asp?id=<%=id%>&func=slet&matregid=<%=oRec("mfid")%>&fromsdsk=<%=fromsdsk%>&aftid=<%=aftid%>"><img src="../ill/slet_16.gif" alt="<%=tsa_txt_221 %>" border=0 /></a>
	                <%else%>
	                &nbsp;
	                <%end if %>
            	    
	                </td>
	                
	                <td style="border-bottom:1px #cccccc solid; padding-top:15px; width:60px;" align=center>
	                
	                 <%if cint(oRec("godkendt")) = 1 then
                        gkendt1CHK = "SELECTED"
                        gkendt0CHK = ""
                        gkendt = "Ja"
                        gkAf = oRec("gkaf")
                        else
                        gkendt1CHK = ""
                        gkendt0CHK = "SELECTED"
                        gkendt = "Nej" 
                        gkAf = "&nbsp;"
                        end if 
            
             strExport = strExport & gkendt &";"
            
            if print <> "j" then%>
           
          
            <select id="mfidgk_id_<%=x %>" name="mfidgk_<%=oRec("mfid")%>" style="width:50px; font-size:9px;">
                <option value="0" <%=gkendt0CHK %>>Nej</option>
                <option value="1" <%=gkendt1CHK %>>Ja</option>
            </select><br />
            <font class="megetlillesilver"><%=left(gkAf, 10)%>.</font>
           
            <%else 
                if cint(oRec("godkendt")) = 1 then
                Response.Write "<b>Ja</b><br><font class=megetlillesilver>"& gkAf
                else
                Response.Write "Nej"
                end if
            end if
            %>
	                
	                </td>
	                
	                
	                <td style="border-bottom:1px #cccccc solid; padding-top:15px;" align=center>
	                
	                 <%if cint(oRec("afregnet")) = 1 then
            efbCHK = "SELECTED"
            ikbCHK = ""
            afreg = "Ja;"
            else
            ikbCHK = "SELECTED"
            efbCHK = ""
            afreg = "Nej;" 
            end if 
            
             strExport = strExport & afreg &";"
            
            if print <> "j" AND level = 1 then%>
           
           <!--
           <input id="mfid_<%=oRec("mfid")%>" name="mfid_<%=oRec("mfid")%>" value="1" type="radio" <%=efbCHK %> /> <b>Ja</b><br />
           <input id="mfid_<%=oRec("mfid")%>" name="mfid_<%=oRec("mfid")%>" value="0" type="radio" <%=ikbCHK %> /> Nej
            -->
            
            <select id="mfid_id_<%=x%>" name="mfid_<%=oRec("mfid")%>" style="width:50px; font-size:9px;">
                <option value="0" <%=ikbCHK %>>Nej</option>
                <option value="1" <%=efbCHK %>>Ja</option>
            </select><br />&nbsp;
            
            
            <%else 
            
            %>
                        <input type="hidden" id="mfid_id_<%=x%>" name="mfid_<%=oRec("mfid")%>" value="<%=oRec("afregnet")%>" />
            <%
                if oRec("afregnet") = 1 then
                Response.Write "<b>Ja</b>"
                else
                Response.Write "Nej"
                end if
            end if
            %>
            
            <input id="erafregnet" name="erafregnet" type="hidden" value="<%=oRec("mfid")%>" />
           
	                
	                </td>
            		
	            </tr>
	           
	            <%
                Response.flush
	            x = x + 1
	            oRec.movenext
	            wend
            	
	            oRec.close
	            %>
	            <tr>
	                <td colspan=12 align=right style="padding-top:5px;">
                        <u><b>Total:&nbsp;</b></u></td>
	                <td align=right style="padding-right:5px; padding-top:5px;"><u><b><%=formatnumber(totsum,2) %></b></u></td>
	                <td align=right style="padding-right:5px; padding-top:5px;"><%=basisValISO %></td>
	                <td colspan=3>&nbsp;</td>
	            </tr>
	            <tr>
	                <td colspan=12 align=right class=lille>
                        Ikke angivet:&nbsp;</td>
	                <td align=right style="padding-right:5px;" class=lille><%=formatnumber(totsumNone,2) %></td>
	                <td align=right style="padding-right:5px;" class=lille><%=basisValISO %></td>
	                <td colspan=3>&nbsp;</td>
	           </tr>
	             <tr>
	                <td colspan=12 align=right class=lille>
                        Interne:&nbsp;</td>
	                <td align=right style="padding-right:5px;" class=lille><%=formatnumber(totsumInt,2) %></td>
	                <td align=right style="padding-right:5px;" class=lille><%=basisValISO %></td>
	                <td colspan=3>&nbsp;</td>
	            </tr>
	            <tr>
	                <td colspan=12 align=right class=lille>
                        Eksterne:&nbsp;</td>
	                <td align=right style="padding-right:5px;" class=lille><%=formatnumber(totsumEks,2) %></td>
	                <td align=right style="padding-right:5px;" class=lille><%=basisValISO %></td>
	                <td colspan=3>&nbsp;</td>
	            </tr>
	            <tr>
	                <td colspan=12 align=right class=lille>
                        Personlige:&nbsp;</td>
	                <td align=right style="padding-right:5px;" class=lille><b><%=formatnumber(totsumPers,2) %></b></td>
	                <td align=right style="padding-right:5px;" class=lille><%=basisValISO %></td>
	                <td colspan=3>&nbsp;</td>
	            </tr>
	            <tr>
	                <td colspan=14>
                        &nbsp;</td>
	                <td colspan=3 align=right>
                        &nbsp;<%if print <> "j" then %>
                        <input id="Submit1" type="submit" value="<%=tsa_txt_144 %> >> " />
                        <%end if %>
                        </td>
	            </tr>
                    <input id="antal_x" name="antal_x" value="<%=x%>" type="hidden" />
	             </form>
	            </table>
	            
	           </div>
	            <br /><br />&nbsp;
	           
	           <%if print <> "j" And level = 1 then %>
	          <br /><br />
	            <br /><br /><br />
	            <form method="post" action="materiale_stat.asp?func=afregnalle&hidemenu=<%=hidemenu %>" >
	            Marker alle bilags poster på følgende bilag som afregnet: <br />
                <input id="afregnalle_id" name="afregnalle_id" type="text" size=20 value="<%=sogliste %>" />
                &nbsp;<input id="Submit2" type="submit" value="Afregn --> " />
                </form>
                <%end if %>
                
                
                
                
                 
                <%'*** Fordeling på materiale grupper ****'
                
                 call tableDiv(tTop,tLeft,tWdth)
	            %>
	             
	            <table cellspacing="0" cellpadding="2" border="0" bgcolor="#ffffff" width=50%>
	            <tr>
	            <td colspan=2 style="padding:5px;"><h4>Fordeling på grupper</h4>
	            <b>Kun personlige</b> (udbetaling til medarbejdere)</td>
	            </tr>
	                
                <%
                
                sqlWh = replace(sqlWh, "m.mid", "mf.usrid")
                
                strSQLmg = "SELECT sum(mf.matkobspris*mf.matantal*mf.kurs/100) AS sumkobspris, mf.valuta, mf.matgrp, mg.navn AS gnavn FROM materiale_forbrug mf"_
            	&" LEFT JOIN materiale_grp mg ON (mg.id = mf.matgrp) "_
                &" WHERE ("& sqlWh &") "& strDatoKri &" AND "&jobKri&" "& intKodeSQLkri &""_
            	&" "& strSQLper &" "& SQLafr &" "& SQLgk &" AND personlig = 1 GROUP BY mf.matgrp ORDER BY mg.navn" 	
            	
	            'Response.Write strSQLmg
	            'response.flush
	            
	            g = 0 
	            grpIalt = 0
	            oRec4.open strSQLmg, oConn, 3
	            while not oRec4.EOF 
	             select case right(g, 1)
	             case 0,2,4,6,8
	             bgthis = "#FFFFFF"
	             case else
	             bgthis = "#EFF3FF"
	             end select%>
	            <tr bgcolor="<%=bgthis %>">
	                <td style="border-bottom:1px #999999 dashed;">
	                <%if len(trim(oRec4("gnavn"))) <> 0 then %>
	                <%=oRec4("gnavn") %>
	                <%else %>
	                (Ingen gruppe)
	                <%end if %></td>
	                <td style="border-bottom:1px #999999 dashed;" align=right><%=formatnumber(oRec4("sumkobspris"), 2 ) &" "& basisValISO%> </td>
	            </tr>
	            <%
	            grpIalt = grpIalt + oRec4("sumkobspris")
	            
	            g = g + 1
	            oRec4.movenext
	            wend
	            oRec4.close
	            %>
	            <tr bgcolor="#FFC0CB">
	                <td><b>Total:</b></td>
	                <td align=right><b><%=formatnumber(grpIalt, 2) &" "& basisValISO %> </b></td>
	            </tr>
	            </table>
	            </div>
	            
	
	<%end select '* pr job eller sum eller som udgiftsbilag%>
	
	
	
	
	
	
	
	
	
	
	
	<%'if x <> 0 then%>
	
	
	
	<%if print <> "j" then%>
	
	<br><br>
	
	<%if visprjob_ell_sum = 0 then %>
	<form action="materialer_ordrer.asp?func=opr" method="post" target="_blank"> 
			<input type="hidden" name="matids" id="matids" value="<%=strMatids%>">
			<%=strAntalFM%>
		    <input type="submit" value="Opret mat. ordre >>">
			
			</form>
	<%end if %>
	        
	        
	        
	         <%
	            
	           
	'*** Generalt link med alle variabler ***
	link = "&FM_visprjob_ell_sum="&visprjob_ell_sum&"&FM_job="&jobid&"&FM_jobans="&jobans&""_
	&"&FM_jobans2="&jobans2&"&FM_kundeans="&kundeans&"&FM_kunde="&kundeid&"&FM_medarb="&thisMiduse&""_
	&"&FM_medarb_hidden="&thisMiduse&"&FM_intkode="&intKode&"&FM_progrupper="&progrp&""_
	&"&FM_kundejobans_ell_alle="&visKundejobans&"&FM_jobsog="&jobSogValPrint&"&viskunabnejob0="&viskunabnejob0&"&viskunabnejob1="&viskunabnejob1&"&viskunabnejob2="&viskunabnejob2&"&FM_segment="&segment

                'if hidemenu = "1" then
                'ptop = 25 '72 + addTopPx
                'else
                ptop = 26 '72 + addTopPx
                'end if

                pleft = 930
                pwdt = 165

                call eksportogprint(ptop,pleft, pwdt)
                
               %>

                <form action="materiale_stat_eksport.asp" method="post" name=theForm2 onsubmit="BreakItUp2()" target="_blank"> <!--  -->
			    <input type="hidden" name="datointerval" id="datointerval" value="<%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%>">
			    <input type="hidden" name="txt1" id="txt1" value="">
                <textarea name="BigTextArea" id="BigTextArea" style="visibility:hidden; width:1px; height:1px;"><%=strExport%></textarea>
			    
			    <input type="hidden" name="txt20" id="txt20" value="">

                <tr>

                    <td align=center><input type="image" src="../ill/export1.png">
                    
                    </td>
                    </td><td>.csv fil eksport</td>
                    </tr>
                    <tr>
                    <td align=center>
                   <a href="materiale_stat.asp?menu=stat&print=j&id=<%=id%><%=link%>&hidemenu=<%=hidemenu %>" target="_blank"  class='vmenu'>
                   &nbsp;<img src="../ill/printer3.png" border=0 alt="" /></a>
                </td><td>Print version</td>
                   </tr>
                   
	                </form>
                   </table>
                </div>
                
                <%if cint(visprjob_ell_sum) = 1 OR cint(visprjob_ell_sum)  = 0 then %>
                <div style="position:relative; visibility:visible; top:100px; padding:20px; width:400px;"><h4>Side note(r):</h4>
	            Hvis du er Jobansvarlig eller Administrator har du mulighed for at ændre i antal. For at slette en materiale-registrering tast <b>0 (nul) i antal</b> og opdater. Antal på lager vil samtidig blive opdateret.
	            </div>
	            <%end if %>
	            <br><br><br><br><br><br><br><br><br><br>&nbsp;
                
                <%else%>

                <% 
                Response.Write("<script language=""JavaScript"">window.print();</script>")
                %>
                <%end if%>
	
	
	<%'end if%><!-- x <> 0 -->
	
	<br><br>&nbsp;
	</div>
	<%end select%>
	
	


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
