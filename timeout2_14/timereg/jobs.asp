<%response.buffer = true%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/job_func.asp"-->
<!--#include file = "CuteEditor_Files/include_CuteEditor.asp" -->
<!--#include file="../inc/regular/topmenu_inc.asp"--> 
<!--#include file="inc/convertDate.asp"-->
<!--#include file="inc/timbudgetsim_inc.asp"-->




<%


 'response.write "request(FM_fomr): "  & request("FM_fomr") & "<br>"
    

'** Jquery section 
if Request.Form("AjaxUpdateField") = "true" then

%>
<!--  Added by Lei 17-06-2013-->
    <!-- <link rel="stylesheet" href="../inc/jquery/jquery-ui-1.10.3.custom/css/ui-lightness/jquery-ui-1.10.3.custom.min.css" /> -->     
    <script src="../inc/jquery/jquery-1.10.1.min.js" type="text/javascript"></script>
    <script src="../inc/jquery/jquery-ui-1.10.3.custom/js/jquery-ui-1.10.3.custom.min.js" type="text/javascript"></script>

    <style>
        #sortable { list-style-type: none; margin: 0; padding: 0; width: 60%; }
        #sortable li { margin: 0 3px 3px 3px; padding: 0.4em; padding-left: 1.5em; font-size: 1.4em; height: 18px; }
        #sortable li span { position: absolute; margin-left: -1.3em; }
    </style>

    <script type="text/javascript">

        /*
        Restore globally scoped jQuery variables to the first version loaded
        (the newer version)
        */
        jq1101 = jQuery.noConflict(true);

    </script>
    <script src="../inc/jquery/dragSort.jquery.js" type="text/javascript"></script>
<!--End of added by Lei 17-06-2013-->


<%


Select Case Request.Form("control")
case "FM_sortOrder"
              Call AjaxUpdate("aktiviteter","sortorder","")

case "FN_getAktlisten"

               sfunc = request("func")

               if sfunc = "dbred" then          
              %>
              <script src="inc/job_jav_aktliste.js"></script> 
              <%
              
              jq_func = "red"
              jq_sort = ""
              jq_id = jq_visalle
             
              if len(trim(request.Form("cust"))) <> 0 then
              jq_id = request.Form("cust")
              else
              jq_id = 0
              end if
              
              jq_vispasluk = request("jq_visalle")

              'slto = request("slto") 'bruges kun ved test
              
               
              
              '*** Henter eksisterende ***'
	          call hentaktiviterListe(jq_id, jq_func, jq_vispasluk, sort)
	          'Response.write "<table><tr><td>Hej der1</td></tr></table>"


	          '*** ÆØÅ **'
              call jq_format(strAktListe)
              strAktListeTxt = jq_formatTxt
	          
	          Response.Write strAktListeTxt

             end if




case "FN_getKundelisten_2013"
              
              dim kundeArr_2013
              redim kundeArr_2013(25)
              
              if len(trim(request.Form("cust"))) <> 0 then
              jq_sog_val = request.Form("cust")
              else
              jq_sog_val = ""
              end if


              if jq_sog_val <> "" then
              jq_sog_valSQL = " AND (Kkundenavn LIKE '"& jq_sog_val &"%' OR Kkundenr LIKE '"& jq_sog_val &"%')"
              else
              jq_sog_valSQL = ""
              end if

              'jq_sog_valSQL = " AND (Kkundenavn LIKE 'p%')"

                strSQL = "SELECT Kkundenavn, Kkundenr, Kid, kundeans1, kundeans2 FROM kunder WHERE ketype <> 'e' "& jq_sog_valSQL &" ORDER BY Kkundenavn"
			    
                'if session("mid") = 34 then
                'Response.write "<option>"& jq_sog_valSQL &" :: "& strSQL& "</option>"
                'Response.end
                'end if
                
                z = 0
                oRec.open strSQL, oConn, 3
			    kans1 = ""
			    kans2 = ""
			    while not oRec.EOF
				
			    
                

            	strSQL2 = "SELECT mnavn, mnr FROM medarbejdere WHERE mid = "&oRec("kundeans1")
				oRec2.open strSQL2, oConn, 3 
				if not oRec2.EOF then
				kans1 = oRec2("mnavn")
				end if
				oRec2.close
				
				strSQL2 = "SELECT mnavn, mnr FROM medarbejdere WHERE mid = "&oRec("kundeans2")
				oRec2.open strSQL2, oConn, 3 
				if not oRec2.EOF then
				kans2 = " - &nbsp;&nbsp;" & oRec2("mnavn") 
				end if
				oRec2.close
				
				if len(kans1) <> 0 OR len(kans2) <> 0 then
				anstxt = "....................kontaktansv 1: "
				else
				anstxt = ""
				end if
				
			'strKlistTxt = strKlistTxt & "<option value='"&oRec("Kid")&"'>"&oRec("Kkundenavn")&"&nbsp;&nbsp;("&oRec("Kkundenr")&") " & anstxt &""& kans1 &"&nbsp;&nbsp;"& kans2 &"</option>"
            kundeArr_2013(z) = "<option value='"&oRec("Kid")&"'>"&oRec("Kkundenavn")&"&nbsp;&nbsp;("&oRec("Kkundenr")&") " & anstxt &""& kans1 &"&nbsp;&nbsp;"& kans2 &"</option>"
            

            'Response.flush
		    
			z = z + 1
			kans1 = ""
			kans2 = ""
			oRec.movenext
			wend
			oRec.close

              
              for z = 0 to z - 1
              '*** ÆØÅ **'
              call jq_format(kundeArr_2013(z))
              kundeArr_2013(z) = jq_formatTxt
	          
	          Response.Write kundeArr_2013(z)
              next


case "FN_getKundeKperslisten"
              
              dim kundeKpersArr
              redim kundeKpersArr(500)
              
              if len(trim(request.Form("cust2"))) <> 0 AND request.Form("cust2") <> 0 then
              kidThis = request.Form("cust2")
              else
              kidThis = -1
              end if

              if len(trim(request("jq_kundekpers"))) <> 0 then
              jq_kundekpers = request("jq_kundekpers")
              else
              jq_kundekpers = 0
              end if


                strSQL = "SELECT navn, id, titel, email FROM kontaktpers WHERE kundeid = "& kidThis
			    
                'Response.write "<option>"& strSQL& "</option>"
                'Response.end
                

                kundeKpersArr(0) = "<option value='0'>Alle</option>"
                z = 1

                oRec.open strSQL, oConn, 3
			   
			    while not oRec.EOF
				
			    
                if cint(jq_kundekpers) = oRec("id") then
                kpsSEL = "SELECTED"
                else
                kpsSEL = ""
                end if

                if len(trim(oRec("email"))) <> 0 then
                eTxt = ", " & oRec("email")
                else
                eTxt = ""
                end if


                if len(trim(oRec("titel"))) <> 0 then
                tTxt = oRec("titel")
                else
                tTxt = ""
                end if
                

            	
			kundeKpersArr(z) = "<option value='"&oRec("id")&"' "&kpsSEL&">"&oRec("navn")&" "& tTxt &" "& eTxt &"</option>"
            

            'Response.flush
		    
			z = z + 1
			
			oRec.movenext
			wend
			oRec.close

            kundeKpersArr(z) = "<option value='0'>Ingen kontaktpersoner fundet</option>"
              
               for z = 0 to UBOUND(kundeKpersArr)
              '*** ÆØÅ **'
              call jq_format(kundeKpersArr(z))
              kundeKpersArr(z) = jq_formatTxt
	          
	          Response.Write kundeKpersArr(z)
              next
		

end select
Response.end
end if



if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
    '** NOT READY ***'
    if len(trim(request("nomenu"))) <> 0 AND request("nomenu") <> "0" then                       
	nomenu = 1
    else
    nomenu = 0
    end if

    nomenu = 0
    '*********************


	function SQLBless3(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ".", ",")
		SQLBless3 = tmp
	end function
	
	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function
	
	function SQLBless2(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless2 = tmp
	end function
	

	
	'*** Altid = 1 for en sikkedrheds skyld, men den bruges ikke mere pr. 1/10-2005 ***
	'*** Kun eksterne job herefter *** 
	strInternt = 1 ' = request("int") 
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if

    tp_jobid = id
	
    '**** Hvilken div skal vises? ****
	if len(request("showdiv")) <> 0 then
	showdiv = request("showdiv")
	else
	showdiv = "stamdata"
	end if

    if len(trim(request("FM_tp_jobaktid"))) <> 0 then
    tp_jobaktid = request("FM_tp_jobaktid")
    else
    tp_jobaktid = 0
    end if


    if len(trim(request("FM_hd_tp_jobaktid"))) <> 0 then
    hd_tp_jobaktid = request("FM_hd_tp_jobaktid")
    else
    hd_tp_jobaktid = 0
    end if


    if len(trim(request("FM_mtype"))) <> 0 then 'AND request("FM_mtype") <> 0 
    mtype = request("FM_mtype")
    else
        select case lto
        case "epi", "epi_no", "epi_sta", "epi_ab", "epi_cati", "epi_uk"
        mtype = 10
        case else
        mtype = 0
        end select
    
    end if
    
    if mtype <> 0 then
    mtypeSQLKri = " medarbejdertype = " & mtype
    else
    mtypeSQLKri = " medarbejdertype <> 0"
    end if
    

    if len(trim(request("FM_sorter_tp"))) <> 0 then
    sorttp = request("FM_sorter_tp")
        if sorttp = "1" then
        sorttpCHK0 = ""
        sorttpCHK1 = "CHECKED"
        else
        sorttpCHK0 = "CHECKED"
        sorttpCHK1 = ""
        end if
    else
        
        if request.cookies("tsa")("sorttp") <> "" then
            

        if request.cookies("tsa")("sorttp") = "1" then
        sorttp = 1
        sorttpCHK0 = ""
        sorttpCHK1 = "CHECKED"
        else
        sorttp = 0
        sorttpCHK0 = "CHECKED"
        sorttpCHK1 = ""
        end if


        else
        sorttp = 0
        sorttpCHK0 = "CHECKED"
        sorttpCHK1 = ""
        end if
    end if

    response.cookies("tsa")("sorttp") = sorttp


    if len(trim(request("FM_fomr"))) <> 0 then
    strFomr_rel = request("FM_fomr")
    strFomr_rel = replace(strFomr_rel, "X234", "#")
    else
    strFomr_rel = "#0#"
    end if

   
    

                    '*** Forretningsområder **' 
	                strFomr_Gblnavn = ""
                    strFomr_relA = replace(strFomr_rel, "#", "")
                    strFomr_relA = split(strFomr_relA, ",")

                    fo = 0
                    for f = 0 to UBOUND(strFomr_relA)

                    strSQLfrel = "SELECT fomr.navn FROM fomr WHERE fomr.id = "& strFomr_relA(f) 

                    'Response.Write strSQLfrel
                    'Response.flush
                    
                    oRec3.open strSQLfrel, oConn, 3
                    if not oRec3.EOF then

                    if fo = 0 then
                    strFomr_Gblnavn = " ("
                    end if

                    strFomr_Gblnavn = strFomr_Gblnavn & oRec3("navn") & ", " 
                   
                    fo = fo + 1
                    end if
                    oRec3.close

                    next

                    if fo <> 0 then
                    len_strFomr_Gblnavn = len(strFomr_Gblnavn)
                    left_strFomr_Gblnavn = left(strFomr_Gblnavn, len_strFomr_Gblnavn - 2)
                    strFomr_Gblnavn = left_strFomr_Gblnavn & ")"

                    if len(strFomr_Gblnavn) > 50 then
                    strFomr_Gblnavn = left(strFomr_Gblnavn, 50) & "..)"
                    end if
                    end if		    

	
	thisfile = "jobs"
	rdir = request("rdir")
	
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<!-------------------------------Sideindhold------------------------------------->
	<%
	
	slttxt = "<b>A) Jeg ønsker at slette jobbet helt.</b><br />"_
	&"Du vil samtidig slette <b>alle aktiviteter, timeregistreringer, materialeforbrug, udlæg og<br> ressourcetimer</b> på dette job."
	slturl = "jobs.asp?menu=job&func=sletok&id="&id&"&kt=0&fm_kunde_sog="&request("fm_kunde_sog")&"&jobnr_sog="&request("jobnr_sog")&"&filt="&request("filt")&"&rdir="&rdir
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,110,90)
	
	slttxtb = "<b>B) Jeg vil nulstille jobbet.</b><br>Jeg ønsker ikke at slette job eller aktiviteter, men nulstille/slette<br> de timeregistreringer og ressourcetimer der findes."
    slturlb = "jobs.asp?menu=job&func=sletok&id="&id&"&kt=1&fm_kunde_sog="&request("fm_kunde_sog")&"&jobnr_sog="&request("jobnr_sog")&"&filt="&request("filt")&"&rdir="&rdir
	
	
	call sltque(slturlb,slttxtb,slturlalt,slttxtalt,540,90)
	
	slttxtc = "<b>C) Slet de tilføjede aktiviteter.</b><br>Jeg ønsker ikke at slette jobbet, men slette<br> de tilhørende aktivteter, timeregistreringer og ressourcetimer."
    slturlc = "jobs.asp?menu=job&func=sletok&id="&id&"&kt=2&fm_kunde_sog="&request("fm_kunde_sog")&"&jobnr_sog="&request("jobnr_sog")&"&filt="&request("filt")&"&rdir="&rdir
	
	
	call sltque(slturlc,slttxtc,slturlalt,slttxtalt,970,90)
	
	
	
	case "sletok"
	'*** Her slettes et job, dets aktiviteter og de indtastede timer på jobbet ***
	strSQL = "SELECT id, navn FROM aktiviteter WHERE job = "& id &"" 
	kt = request("kt") 'slet kun timer
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
		
		'*** Indsætter i delete historik ****'
	    call insertDelhist("akt", oRec("id"), 0, oRec("navn"), session("mid"), session("user"))
		
		'** Slet eller nulstil? **'
		if kt = "0" OR kt = "2" then
		oConn.execute("DELETE FROM aktiviteter WHERE id = "& oRec("id") &"")
		end if
		oConn.execute("DELETE FROM timer WHERE TAktivitetId = "& oRec("id") &"")
		
	oRec.movenext
	wend
	oRec.close
	
	'*** Sletter ressource timer på job ****
	oConn.execute("DELETE FROM ressourcer WHERE jobid = "& id &"")
	
	'*** Sletter ressource_md timer på job ****
	oConn.execute("DELETE FROM ressourcer_md WHERE jobid = "& id &"")
	
	'*** Sletter fra Guiden aktive job timer på job ****
	oConn.execute("DELETE FROM timereg_usejob WHERE jobid = "& id &"")
	
	'*** Sletter materialeforbrug ****
	oConn.execute("DELETE FROM materiale_forbrug WHERE jobid = "& id &"")

	
	'*** Sletter fakturaer på job ****'
	'*** Kan ikke slette job der findes fakturaer på pr. 16/11-2008 **'
	
	'strSQLfak = "SELECT fid FROM fakturaer WHERE jobid = "& id &""
	'oRec.open strSQLfak, oConn, 3
	
	'while not oRec.EOF 
	        
	'        oConn.execute("DELETE FROM faktura_det WHERE fakid = "& oRec("fid") &"")
	'        oConn.execute("DELETE FROM fak_med_spec WHERE fakid = "& oRec("fid") &"")
	'        oConn.execute("DELETE FROM fak_mat_spec WHERE matfakid = "& oRec("fid") &"")
	        
	'oRec.movenext
	'wend
	'oRec.close
	
	
	'oConn.execute("DELETE FROM fakturaer WHERE jobid = "& id &"")
	
	
	strSQL = "SELECT id, jobnavn, jobnr FROM job WHERE id = "& id &"" 
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
		
		'*** Indsætter i delete historik ****'
	    call insertDelhist("job", id, oRec("jobnr"), oRec("jobnavn"), session("mid"), session("user"))
		
	
	end if
	oRec.close
	
	
	
	
	'Response.flush
	
	if kt = "0" then
	oConn.execute("DELETE FROM job WHERE id = "& id &"")
	end if
	
	if rdir <> "webblik" then
	Response.redirect "jobs.asp?menu=job&shokselector=1&fm_kunde="&request("fm_kunde_sog")&"&jobnr_sog="&request("jobnr_sog")&"&filt="&request("filt")
	else
	Response.redirect "webblik_joblisten.asp"
	end if
	
	case "sletfil"
	
	'*** Her spørges om det er ok at der slettes en fil ***
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190px; top:150px; visibility:visible;">
	<br><br><br>
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td><img src="../ill/alert.gif" width="44" height="45" alt="" border="0">Du er ved at <b>slette</b> en fil. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	   <a href="jobs.asp?menu=jobs&func=sletfilok&id=<%=id%>&filnavn=<%=request("filnavn")%>&jobid=<%=request("jobid")%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	case "sletfilok"
	'*** Her slettes en fil ***
	'ktv
	'strPath =  "E:\www\timeout_xp\wwwroot\ver2_1\upload\"&lto&"\" & Request("filnavn")
	'Qwert
	strPath =  "d:\dkdomains\outzource\wwwroot\timeout_xp\wwwroot\ver2_1\inc\upload\"&lto&"\" & Request("filnavn")
	'Response.write strPath
	
	on Error resume Next 

	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	Set fsoFile = FSO.GetFile(strPath)
	fsoFile.Delete
	
	
	oConn.execute("DELETE FROM filer WHERE id = "& id &"")
	Response.redirect "jobs.asp?menu=job&shokselector=1&func=red&id="& request("jobid")&"&int=1"
	
	

    '************************************************************'
    '************** Opdater jobliste ****************************'
	'************************************************************'
    case "opdjobliste"
	
	'Response.write "Opdaterer liste"
	
    if len(trim(request("FM_formr_opdater"))) <> 0 then
    fomrJa = request("FM_formr_opdater")
    else
    fomrJa = 0
    end if

    if len(trim(request("FM_formr_opdater_akt"))) <> 0 then
    fomrJaAkt = request("FM_formr_opdater_akt")
    else
    fomrJaAkt = 0
    end if

    

    slutdato = request("FM_opdaterslutdato_dato")

    if len(trim(request("FM_opdaterslutdato"))) <> 0 then
    opdaterslutdato = 1
    else
    opdaterslutdato = 0
    end if

    slutdato = replace(slutdato, ".", "-")

    'Response.Write slutdato & IsDate('slutdato)
    if IsDate(slutdato) = True then

    'Response.Write "<br>" & slutdato
    slutdatoSQL = year(slutdato) &"/"& month(slutdato) & "/"& day(slutdato) 
    else

    opdaterslutdato = 0

    end if


   

    fomrArr = split(request("FM_fomr"), ",")

                    'for_faktor = 0
                    'for afor = 0 to UBOUND(fomrArr)
                    'for_faktor = for_faktor + 1 
                    'next

                    'if for_faktor <> 0 then
                    'for_faktor = for_faktor
                    'else
                    for_faktor = 1
                    'end if

                    'for_faktor = formatnumber(100/for_faktor, 2)
                    'for_faktor = replace(for_faktor, ",", ".")

	jobids = split(request("FM_listeid"), ",")
	'jobstatus = split(request("FM_listestatus"), ",")
	for t = 0 to UBOUND(jobids)
		

        jobStatus = request("FM_listestatus_"& trim(jobids(t)) &"")
        'Response.Write "FM_listestatus_"& trim(jobids(t)) &"" & "<br>"

         '*** Sæt lukkedato (skal være før det skifter status) ***'
        call lukkedato(jobids(t), jobStatus)

		strSQL = "UPDATE job SET jobstatus = " & jobStatus 
		
        
        if opdaterslutdato = 1 then
        strSQL = strSQL & ", jobslutdato = '"& slutdatoSQL & "'"
        end if


        strSQL = strSQL & " WHERE id = "& jobids(t)

        oConn.execute(strSQL)

        '**** SyncDatoer ****'
        jobnrThis = 0
        syncslutdato = 0
        strSQLj = "SELECT jobnr, syncslutdato FROM job WHERE id =  "& jobids(t)
        oRec5.open strSQLj, oConn, 3 
        if not oRec5.EOF then

        jobnrThis = oRec5("jobnr")
        syncslutdato = oRec5("syncslutdato")

        end if
        oRec5.close

        if jobStatus = 0 AND syncslutdato = 1 then
            call syncJobSlutDato(jobids(t), jobnrThis, syncslutdato)
        end if


       

        


		'Response.write strSQL & "<br>"

                            
                            
                            
                            '**** Sætter EASYreg aktiv på aktiviteter for alle medarbejdere (hvis der findes EASY regaktiviter) ****'
                            oEasyReg = 0
                            strSQLea = "SELECT easyreg, id FROM aktiviteter WHERE job = "& jobids(t) & " AND easyreg = 1"
				            oRec5.open strSQLea, oConn, 3
				            while not oRec5.EOF 
				                
                              oEasyReg = 1
                            
                            oRec5.movenext
				            wend 
				            oRec5.close
				                

                                if cint(oEasyReg) = 1 then
				                strSQLtreguse = "UPDATE timereg_usejob SET easyreg = " & jobids(t) & " WHERE jobid = " & jobids(t)
	   	                        'Response.Write strSQLtreguse & "<br>"
	   	                        'Response.flush
	   	                        oConn.execute(strSQLtreguse)
				                end if
				               
				            'Response.end




        '*** Forretningsområder ****'
	    if cint(fomrJa) = 1 then
                    

                    '*** nulstiller job (IKKE akt) ****'
                    strSQLfor = "DELETE FROM fomr_rel WHERE for_jobid = "& jobids(t)  & " AND for_aktid = 0"
                    oConn.execute(strSQLfor)

                    'Response.Write "her"
                    'Response.Flush
                    'fomrArrLink = "#0#"
                    for afor = 0 to UBOUND(fomrArr)

                            'Response.Write "her2" & afor & "<br>"
                            'Response.Flush

                            fomrArr(afor) = replace(fomrArr(afor), "#", "")

                            if fomrArr(afor) <> 0 then

                            strSQLfomri = "INSERT INTO fomr_rel "_
                            &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                            &" VALUES ("& fomrArr(afor) &", "& jobids(t) &", 0, "& for_faktor &")"

                            oConn.execute(strSQLfomri)

                            'fomrArrLink = fomrArrLink & ",#"& fomrArr(afor) & "#" 
            

                                '**** IKKE MERE 11.3.2015 *****'
                                '*** Altid Sync aktiviteter ved multiopdater ****'
                                if cint(fomrJaAkt) = 1 then
                                strSQLa = "SELECT id FROM aktiviteter WHERE job = "& jobids(t)
                                oRec3.open strSQLa, oConn, 3
                                while not oRec3.EOF 


                                     '*** nulstiller form på akt () ****'
                                    strSQLfor = "DELETE FROM fomr_rel WHERE for_jobid = "& jobids(t)  & " AND for_aktid = " & oRec3("id")
                                    oConn.execute(strSQLfor)

                                    strSQLfomrai = "INSERT INTO fomr_rel "_
                                    &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                                    &" VALUES ("& fomrArr(afor) &", "& jobids(t) &","& oRec3("id") &", "& for_faktor &")"

                                   oConn.execute(strSQLfomrai)

                                oRec3.movenext
                                wend
                                oRec3.close
                                end if

                                

                            end if


                    next

                    '********************************'



        end if

	next
	'response.write "jobs.asp?menu=job&jobnr_sog="&request("jobnr_sog")&"&filt="&request("filt")&"&fm_kunde_sog="&request("fmkunde")&"&FM_kunde="&request("FM_kunde")&"&FM_fomr="&request("FM_fomr")
    'Response.end
	
    formTemp = replace(request("FM_fomr"), "#", "X234")

    Response.redirect "jobs.asp?menu=job&jobnr_sog="&request("jobnr_sog")&"&filt="&request("filt")&"&fm_kunde_sog="&request("fmkunde")&"&FM_kunde="&request("FM_kunde")&"&FM_fomr="&formTemp
	
	
	
	
	case "dbopr", "dbred"
	'*** Her indsættes et nyt job i db ****
	
	
    showaspopup = request("showaspopup")
	
	public useleftdiv
	sub visErrorFormat
	
		if showaspopup <> "y" then
		%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		
		<%
		else
		%>
		<!--#include file="../inc/regular/header_hvd_inc.asp"-->
		<%
		useleftdiv = "j"
		end if
	
	end sub
	
	
				 '*tjekker om dato eksisterer og tildeler ny dato hvis den ikke gør ** 
				call dato_30(Request("FM_start_dag"), Request("FM_start_mrd"), Request("FM_start_aar"))
				strStartDay = strDay_30
				
				call dato_30(Request("FM_slut_dag"), Request("FM_slut_mrd"), Request("FM_slut_aar"))
				strSlutDay = strDay_30
	
		
		slutaar = right(Request("FM_slut_aar"), 2)
		if slutaar = "44" then
		slutaar = "2044"
		else
		slutaar = slutaar
		end if
		
		startDatoNum = cdate(strStartDay &"/"& Request("FM_start_mrd") &"/"& Request("FM_start_aar")) 
		slutDatoNum = cdate(strSlutDay &"/"& Request("FM_slut_mrd")  &"/"& slutaar)
		
		'Response.write startDatoNum & "==" & slutDatoNum
		
		'Response.Write "request(FM_kunde)" & request("FM_kunde")
		
        
    'Response.Write request("FM_kunde") & request("FM_navn") & " "& request("FM_jnr")
    'Response.end

		'** Tjekker om alle felter er udfyldt korrekt **
		if len(request("FM_navn")) = 0 OR len(request("FM_navn")) > 50 OR len(trim(request("FM_jnr"))) = 0 OR len(request("FM_kunde")) = 0 OR _
		len(request("FM_jnr")) > 20 OR startDatoNum > slutDatoNum OR _
		cint(instr(request("FM_navn"), "'")) > 0 then
		
		call visErrorFormat
		errortype = 14
		call showError(errortype)
		Response.end
		
		else
		
		
		    %>
			<!--#include file="inc/isint_func.asp"-->
			<%
			
			isInt = 0
			call erDetInt(request("prio")) 
			if isInt > 0 then
			    
			    call visErrorFormat
				
				errortype = 132
				call showError(errortype)
			
			response.End
			end if
			
			
			
			isInt = 0    
		    call erDetInt(request("FM_sandsynlighed"))
            if isInt > 0 then
            call visErrorFormat
    		
		    errortype = 31
		    call showError(errortype)
		    response.End
			end if


            '** forretningsområde mandatory 
            call fomr_mandatory_fn()
            if cint(fomr_mandatoryOn) = 1 AND len(trim(request("FM_fomr"))) = 0 then

            call visErrorFormat
			errortype = 177
			call showError(errortype)
			Response.end


            end if
			
			
			'Response.Write  request("FM_jnr")
			'Response.flush
			        
                    '*** Jobnr og tilbudsnummer må gerne indeholde nr. **'
                    isInt = 0
			        'call erDetInt(request("FM_jnr")) 
			        'if cint(isInt) > 0 OR instr(request("FM_jnr"), ".") <> 0 OR (request("FM_jnr") = "0" AND func = "dbred") then
				
				     '   call visErrorFormat
				
				     '   errortype = 15
				     '   call showError(errortype)
			
			        'isInt = 0
			
			        'else
			                'call erDetInt(request("FM_tnr"))
			                'if isInt > 0 then
				
				            'call visErrorFormat
				
				            'errortype = 15
				            'call showError(errortype)
			
			                'isInt = 0
			
			                'else
			    
			    if request("simpeludv") = "1" then
			    simpeludvEXT = "_s"
			    response.cookies("tsa")("be_simpel_udv") = "1"
			    else
			    simpeludvEXT = ""
			    response.cookies("tsa")("be_simpel_udv") = "2"
			    end if
			    
			    
				call erDetInt(request("FM_budget"&simpeludvEXT&""))
				if isInt > 0 OR instr(request("FM_budget"&simpeludvEXT&""), "-") <> 0 OR trim(lcase(request("FM_budget"&simpeludvEXT&""))) = "nan" then
				
				call visErrorFormat
				
				errortype = 16
				call showError(errortype)
				isInt = 0
						
						else
						
						
						call erDetInt(request("FM_budgettimer"&simpeludvEXT&""))
						if isInt > 0 then
						
						call visErrorFormat
						
						errortype = 20
						call showError(errortype)
						isInt = 0
								
								else
								
								call erDetInt(request("FM_ikkebudgettimer"&simpeludvEXT&""))
								if isInt > 0 then
								
								call visErrorFormat
								
								errortype = 20
								call showError(errortype)
								isInt = 0
								
								else
								
								
								if request("FM_fastpris") = "1" AND len(trim(request("FM_budgettimer"))) = 0 then 'OR request("FM_fastpris") = "1" AND request("FM_budgettimer") = "0" then
									
									call visErrorFormat
									
									errortype = 30
									call showError(errortype)
									
								
								
								
				                
				  else
				
				
				'**** Validering OK ***
				'** kundeid og navn *** 
				
				
				strKundeId = request("FM_kunde")
				varJobId = 0
				
				'Response.Write strKundeid
				'Response.end
				
				'*** Multible opret ***
				if instr(request("FM_kunde"), ",") <> 0 then
				    
				    kids = split(request("FM_kunde"), ",")
				    for x = 0 to UBOUND(kids)
				        
				        if kids(x) <> 0 then
				        kidSQLKri = kidSQLKri & " kid = "& kids(x) & " OR"
				        end if
				        
				        if x = 1 then
				        strKidUseOne = kids(x)
				        end if 
				    next
				    
				    len_kidSQLKri = len(kidSQLKri)
				    left_kidSQLKri = left(kidSQLKri, len_kidSQLKri - 3)
				    kidSQLKri = " ("& left_kidSQLKri &") "
				    
				else
			    kidSQLKri = " kid = "& strKundeId 
			    strKidUseOne = strKundeId
				end if
				
				
				
				

				
				
			    '**********************************
			    '**** Jobdata ****
			    '**********************************
			    
				
				strNavn = replace(request("FM_navn"),chr(34), "")
				
				'** Jobbeskrivelse **'
			    strBesk = SQLBless2(request("FM_beskrivelse"))
			    
			    '*** HTML Replace **'
                call htmlreplace(strBesk)
			    strBesk = htmlparseTxt
			    
			    '* Intern note **'
			    strInternBesk = SQLBless2(request("FM_internbesk"))
			    
			    '*** HTML Replace **'
			    call htmlreplace(strInternBesk)
			    strInternBesk = htmlparseTxt
			    
			    
			    
			    if len(request("FM_ski")) <> 0 then
			    ski = 1
			    else
			    ski = 0
			    end if
			    
			    if len(request("FM_abo")) <> 0 then
			    abo = 1
			    else
			    abo = 0
			    end if
			    
			    if len(request("FM_ubv")) <> 0 then
			    ubv = 1
			    else
			    ubv = 0
			    end if
			    
			    
			    if len(trim(request("FM_sandsynlighed"))) <> 0 then
			    sandsynlighed = formatnumber(request("FM_sandsynlighed"), 0)
			    else
			    sandsynlighed = 0
			    end if
                			    
			    
			    
				strOLDjobnr = request("FM_OLDjobnr")
				

                '***** Jobansvarlige ***'
				intJobans1 = request("FM_jobans_1")
				intJobans2 = request("FM_jobans_2")
				intJobans3 = request("FM_jobans_3")
				intJobans4 = request("FM_jobans_4")
				intJobans5 = request("FM_jobans_5")

                if len(trim(request("FM_jobans_proc_1"))) <> 0 then
				jobans_proc_1 = request("FM_jobans_proc_1")
                else
                jobans_proc_1 = 0
                end if

                jobans_proc_1 = replace(jobans_proc_1, ".", "")
                jobans_proc_1 = replace(jobans_proc_1, ",", ".")

				if len(trim(request("FM_jobans_proc_2"))) <> 0 then
				jobans_proc_2 = request("FM_jobans_proc_2")
                else
                jobans_proc_2 = 0
                end if

                jobans_proc_2 = replace(jobans_proc_2, ".", "")
                jobans_proc_2 = replace(jobans_proc_2, ",", ".")

				if len(trim(request("FM_jobans_proc_3"))) <> 0 then
				jobans_proc_3 = request("FM_jobans_proc_3")
                else
                jobans_proc_3 = 0
                end if

                jobans_proc_3 = replace(jobans_proc_3, ".", "")
                jobans_proc_3 = replace(jobans_proc_3, ",", ".")

				if len(trim(request("FM_jobans_proc_4"))) <> 0 then
				jobans_proc_4 = request("FM_jobans_proc_4")
                else
                jobans_proc_4 = 0
                end if

                jobans_proc_4 = replace(jobans_proc_4, ".", "")
                jobans_proc_4 = replace(jobans_proc_4, ",", ".")

				if len(trim(request("FM_jobans_proc_5"))) <> 0 then
				jobans_proc_5 = request("FM_jobans_proc_5")
                else
                jobans_proc_5 = 0
                end if

                jobans_proc_5 = replace(jobans_proc_5, ".", "")
                jobans_proc_5 = replace(jobans_proc_5, ",", ".")

                errProcVal = 0

                for i = 1 to 5
                      select case i 
                      case 1
                      procVal = jobans_proc_1
                      case 2
                      procVal = jobans_proc_2
                      case 3
                      procVal = jobans_proc_3
                      case 4
                      procVal = jobans_proc_4
                      case 5
                      procVal = jobans_proc_5
                      end select
                    
                   
                    
                    call erDetInt(procVal)
				    
                    if isInt > 0 then
                    errProcVal = 1
                    isInt = 0
                    end if
                
                next




                '***** Salg ansvarlige ******
                if len(trim(request("FM_salgsans_1"))) <> 0 then ' er slags ansvarlige slået til / vist
                salgsans1 = request("FM_salgsans_1")
				salgsans2 = request("FM_salgsans_2")
				salgsans3 = request("FM_salgsans_3")
				salgsans4 = request("FM_salgsans_4")
				salgsans5 = request("FM_salgsans_5")
                else
                salgsans1 = 0
				salgsans2 = 0
				salgsans3 = 0
				salgsans4 = 0
				salgsans5 = 0
                end if

                if len(trim(request("FM_salgsans_proc_1"))) <> 0 then
				salgsans_proc_1 = request("FM_salgsans_proc_1")
                else
                salgsans_proc_1 = 0
                end if

                salgsans_proc_1 = replace(salgsans_proc_1, ".", "")
                salgsans_proc_1 = replace(salgsans_proc_1, ",", ".")

				if len(trim(request("FM_salgsans_proc_2"))) <> 0 then
				salgsans_proc_2 = request("FM_salgsans_proc_2")
                else
                salgsans_proc_2 = 0
                end if

                salgsans_proc_2 = replace(salgsans_proc_2, ".", "")
                salgsans_proc_2 = replace(salgsans_proc_2, ",", ".")

				if len(trim(request("FM_salgsans_proc_3"))) <> 0 then
				salgsans_proc_3 = request("FM_salgsans_proc_3")
                else
                salgsans_proc_3 = 0
                end if

                salgsans_proc_3 = replace(salgsans_proc_3, ".", "")
                salgsans_proc_3 = replace(salgsans_proc_3, ",", ".")

				if len(trim(request("FM_salgsans_proc_4"))) <> 0 then
				salgsans_proc_4 = request("FM_salgsans_proc_4")
                else
                salgsans_proc_4 = 0
                end if

                salgsans_proc_4 = replace(salgsans_proc_4, ".", "")
                salgsans_proc_4 = replace(salgsans_proc_4, ",", ".")

				if len(trim(request("FM_salgsans_proc_5"))) <> 0 then
				salgsans_proc_5 = request("FM_salgsans_proc_5")
                else
                salgsans_proc_5 = 0
                end if

                salgsans_proc_5 = replace(salgsans_proc_5, ".", "")
                salgsans_proc_5 = replace(salgsans_proc_5, ",", ".")

           

                for i = 1 to 5
                      select case i 
                      case 1
                      procVal = salgsans_proc_1
                      case 2
                      procVal = salgsans_proc_2
                      case 3
                      procVal = salgsans_proc_3
                      case 4
                      procVal = salgsans_proc_4
                      case 5
                      procVal = salgsans_proc_5
                      end select
                    
                   
                    
                    call erDetInt(procVal)
				    
                    if isInt > 0 then
                    errProcVal = 1
                    isInt = 0
                    end if
                
                next
                '***** Salg ansvarlige SLUT
			 
               
				
                if len(request("FM_virksomheds_proc")) <> 0 then
                virksomheds_proc = request("FM_virksomheds_proc")
                else
                virksomheds_proc = 0
                end if 

                call erDetInt(virksomheds_proc)
                if isInt > 0 then
                errProcVal = 1
                isInt = 0
                end if




                
                if errProcVal <> 0 then
                call visErrorFormat
				errortype = 151
                call showError(errortype)

                Response.end
                end if

                'if len(trim(lcase(request("FM_budget"&simpeludvEXT&"")))) = "nan" then
                ''call visErrorFormat
				'errortype = 159
                'call showError(errortype)

                'Response.end
                'end if


				if len(trim(request("FM_budget"&simpeludvEXT&""))) = "0" then
				strBudget = 0
				else
				strBudget = request("FM_budget"&simpeludvEXT&"")
				strBudget = replace(strBudget, ".", "")
				strBudget = replace(strBudget, ",", ".")
				end if
				
				if len(trim(request("FM_fastpris"))) = "0" then 
				strFastpris = 0 '0: lbn. timer / 1: fastpris 
				else
				strFastpris = request("FM_fastpris")
				end if
				
				if len(trim(request("FM_jobfaktype"))) <> 0 then
				jobfaktype = request("FM_jobfaktype")
				else
				jobfaktype = 0
				end if
				
				
				if len(trim(request("prio"))) <> 0 then
				intprio = request("prio")
				else
 				intprio = -1
 				end if
 				
				response.Cookies("tsa")("faktype") = jobfaktype
				response.cookies("tsa").expires = date + 45
				
				
				usejoborakt_tp = request("FM_usejoborakt_tp")
				
				
				if len(trim(request("FM_budgettimer"&simpeludvEXT&""))) = 0 then
				strBudgettimer = 0
				else
				strBudgettimer = request("FM_budgettimer"&simpeludvEXT&"")
				strBudgettimer = replace(strBudgettimer, ".", "")
				strBudgettimer = replace(strBudgettimer, ",", ".")
				end if
				
				if len(trim(request("FM_ikkebudgettimer"&simpeludvEXT&""))) = "0" then
				ikkeBudgettimer = 0
				else
				ikkeBudgettimer = request("FM_ikkebudgettimer"&simpeludvEXT&"")
				end if
				
				strKunde = request("FM_kunde")
				
				
				'Response.Write "request(FM_kundese): " & request("FM_kundese")
				'Response.end
				'** Tilgængelig for kunde ***
				if len(request("FM_kundese")) <> 0 then
					if request("FM_kundese_hv") = 1 then
					intkundese = 2 '(når job er lukket)
					else
					intkundese = 1 '(når timer er indtastet)
					end if
				else
				intkundese = 0
				end if
				
				if len(request("FM_opr_kpers")) <> 0 then
				intKundekpers = request("FM_opr_kpers")
				else
				intKundekpers = 0
				end if
				
                
                if cint(intKundekpers) = 0 AND (lto = "dencker") AND rdir <> "redjobcontionue" then
		        call visErrorFormat
				errortype = 155
                call showError(errortype)
                Response.end
		        end if

				if len(request("FM_serviceaft")) <> 0 then
				intServiceaft = request("FM_serviceaft") '1
				else
				intServiceaft = 0
				end if
				
				
				
				strEditor = session("user")
				strDato = session("dato")
				intAktfavgp = "0, "& request("FM_favorit")
				strAktFase = replace("0, "& request("FM_favorit_fase"), "'", "")
				strAktFase = replace(strAktFase, "''", "")
				
				intKorsel = 0 'request("FM_favorit_korsel")
				
				if len(request("FM_lukafmjob")) <> 0 then
				lukafmjob = request("FM_lukafmjob")
				else
				lukafmjob = 0
				end if
				
				if len(request("FM_valuta")) <> 0 then
				valuta = request("FM_valuta")
				else
				valuta = 1 'main valuta
				end if
				
				rekvnr = replace(request("FM_rekvnr"), "'", "''")
				
				if len(trim(rekvnr)) = 0 AND lto = "ets-track" then
		        call visErrorFormat
				errortype = 149
                call showError(errortype)
                Response.end
		        end if

                if len(trim(request("FM_forvalgt"))) <> 0 then
                forvalgt = 1
                else
                forvalgt = 0
                end if
				
		        if len(trim(request("FM_bruttofortj"))) <> 0 AND lCase(request("FM_bruttofortj")) <> "nan" then
		        jo_bruttofortj = replace(request("FM_bruttofortj"), ".", "")
		        jo_bruttofortj = replace(jo_bruttofortj, ",", ".")
		        else
		        jo_bruttofortj = 0
		        end if
		        
		        call erdetint(jo_bruttofortj)
		        if isInt <> 0 then
		        call visErrorFormat
				errortype = 126
                call showError(errortype)
                Response.end
		        end if
		        isInt = 0

                
		        
		        if len(trim(request("FM_db"))) <> 0 AND lCase(request("FM_db")) <> "nan" then
		        
		        jo_dbproc = request("FM_db")
		         
		            call erdetint(jo_dbproc)
		            if isInt <> 0 then
		            call visErrorFormat
				    errortype = 127
                    call showError(errortype)
                    Response.end
                    else
                    jo_dbproc = formatnumber(request("FM_db"),0)
		            end if
		            isInt = 0
		            
				
				else
				jo_dbproc = 0
				end if
				
				
				
				if len(trim(request("FM_gnsinttpris"))) <> 0 AND lCase(request("FM_gnsinttpris")) <> "nan" then
				jo_gnstpris = replace(request("FM_gnsinttpris"), ".", "")
		        jo_gnstpris = replace(jo_gnstpris, ",", ".")
				else
				jo_gnstpris = 0
				end if
				
				call erdetint(jo_gnstpris)
		        if isInt <> 0 then
		        call visErrorFormat
				errortype = 124
                call showError(errortype)
                Response.end
		        end if
				isInt = 0
				
				if len(trim(request("FM_intfaktor"))) <> 0 then
				jo_gnsfaktor = replace(request("FM_intfaktor"), ",", ".")
				else
				jo_gnsfaktor = 0
				end if
				
				call erdetint(jo_gnsfaktor)
		        if isInt <> 0 then
		        call visErrorFormat
				errortype = 128
                call showError(errortype)
                Response.end
		        end if
		        isInt = 0
				
                '** Diff på timer og Sum job vs aktiviteter		
                diff_timer = request("FM_diff_timer")
                diff_sum = request("FM_diff_sum")

                diff_timer = replace(diff_timer, ".", "")
		        diff_timer = replace(diff_timer, ",", ".")

                diff_sum = replace(diff_sum, ".", "")
		        diff_sum = replace(diff_sum, ",", ".")

                if len(trim(diff_timer)) <> 0 then
                diff_timer = diff_timer
                else
                diff_timer = 0
                end if

                 if len(trim(diff_sum)) <> 0 then
                diff_sum = diff_sum
                else
                diff_sum = 0
                end if
 
                        
                if len(trim(request("FM_interntbelob"))) <> 0 AND lCase(request("FM_interntbelob")) <> "nan" then
                jo_gnsbelob = replace(request("FM_interntbelob"), ".", "")
		        jo_gnsbelob = replace(jo_gnsbelob, ",", ".")
                else
                jo_gnsbelob = 0
                end if
                
                call erdetint(jo_gnsbelob)
		        if isInt <> 0 then
		        call visErrorFormat
				errortype = 125
                call showError(errortype)
                Response.end
		        end if
		        isInt = 0

                if len(trim(request("FM_udgifter_ulev"))) <> 0 AND lCase(request("FM_udgifter_ulev")) <> "nan" then
                jo_udgifter_ulev = replace(request("FM_udgifter_ulev"), ".", "")
		        jo_udgifter_ulev = replace(jo_udgifter_ulev, ",", ".")
                else
                jo_udgifter_ulev = 0
                end if

                call erdetint(jo_udgifter_ulev)
		        if isInt <> 0 then
		        call visErrorFormat
				errortype = 1252
                call showError(errortype)
                Response.end
		        end if
		        isInt = 0


                if len(trim(request("FM_udgifter_intern"))) <> 0 AND lCase(request("FM_udgifter_intern")) <> "nan" then
                jo_udgifter_intern = replace(request("FM_udgifter_intern"), ".", "")
		        jo_udgifter_intern = replace(jo_udgifter_intern, ",", ".")
                else
                jo_udgifter_intern = 0
                end if

                call erdetint(jo_udgifter_intern)
		        if isInt <> 0 then
		        call visErrorFormat
				errortype = 1251
                call showError(errortype)
                Response.end
		        end if
		        isInt = 0


                
               
		        
		        if len(trim(request("FM_udgifter"))) <> 0 AND lCase(request("FM_udgifter")) <> "nan" then
		        udgifter = replace(request("FM_udgifter"), ".", "")
                udgifter = replace(udgifter, ",", ".")
                else
                udgifter = 0
                end if
                
                call erdetint(udgifter)
		        if isInt <> 0 then
		        call visErrorFormat
				errortype = 129
                call showError(errortype)
                Response.end
		        end if
		        isInt = 0
				
				
                if len(trim(request("FM_restestimat"))) <> 0 then
                restestimat = request("FM_restestimat")

                        call erdetint(restestimat)
		                if isInt <> 0 then
		                call visErrorFormat
				        errortype = 156
                        call showError(errortype)
                        Response.end
		                end if
		                isInt = 0

                        restestimat = abs(restestimat)

                else
                restestimat = 0
                end if


                '''Tjekekr om brutto beløb er mindre end netto beløb
                select case lto 
                case "epi", "epi_no", "epi_sta", "epi_ab", "epi_uk"
                strBudgetTjk = replace(strBudget, ".", ",")
                jo_gnsbelobTjk = replace(jo_gnsbelob, ".", ",")
                if cdbl(strBudgetTjk) < cdbl(jo_gnsbelobTjk) then
                
                        call visErrorFormat
				        errortype = 165
                        call showError(errortype)
                        Response.end
		               
                
                
                end if 
                end select


               

                stade_tim_proc = request("FM_stade_tim_proc")
       
				
                if len(trim(request("FM_syncaktdatoer"))) <> 0 then
                syncaktdatoer = 1
                else
                syncaktdatoer = 0
                end if

                
                if len(trim(request("FM_syncslutdato"))) <> 0 then
                syncslutdato = 1
                else
                syncslutdato = 0
                end if

                if len(trim(request("FM_brugaltfakadr"))) <> 0 then
                altfakadr = 1
                else
                altfakadr = 0
                end if
				
                if len(trim(request("FM_preconditions_met"))) <> 0 then
                preconditions_met = request("FM_preconditions_met")
                else
                preconditions_met = 0
                end if

                if len(trim(request("FM_fomr_konto"))) <> 0 then
                fomr_konto = request("FM_fomr_konto")
                else
                fomr_konto = 0
                end if


                jfak_moms = request("FM_jfak_moms")
                jfak_sprog = request("FM_jfak_sprog")



                if len(trim(request("FM_opdmedarbtimepriser"))) <> 0 then
                laasmedtpbudget = 1
                else
                        
                        if func = "dbopr" then
                        select case lto 
                        case "jttek", "intranet - local"
                        laasmedtpbudget = 1
                        case else
                        laasmedtpbudget = 0
                        end select    
                        
                        else
                        laasmedtpbudget = 0

                        end if
                end if



                '******* Sti til dokumenter på egen filserver *****'
                if len(trim(request("FM_filepath1"))) <> 0 then 
                filepath1 = request("FM_filepath1")
                filepath1 = replace(filepath1, "'", "")
                filepath1 = replace(filepath1, "\", "#")
                else
                filepath1 = ""
                end if



				
				'**********************************'
				'**** Henter valgte kunde(r) ****'
				'**********************************'
			    strSQLkun = "SELECT kid, kkundenavn FROM kunder WHERE ketype <> 'e' AND "& kidSQLKri &" ORDER BY Kkundenavn"
			    
			    'Response.Write strSQLkun 
			    'Response.flush
			    'Response.end
			    
			    oRec.open strSQLkun, oConn, 3
			    while not oRec.EOF
				
				
				if func = "dbopr" then
				
				'**********************************'
				'*** nyt jobnr / tilbudsnr ***'
				'**********************************'
				strSQL = "SELECT jobnr, tilbudsnr FROM licens WHERE id = 1"
				oRec5.open strSQL, oConn, 3
				if not oRec5.EOF then 
				    
				    strjnr = oRec5("jobnr") + 1
				    
				    if request("FM_usetilbudsnr") = "j" then
				    tlbnr = oRec5("tilbudsnr") + 1
				    else
				    tlbnr = 0
				    end if
				    
				end if
				oRec5.close

                    
				
				else
				  
				  strjnr = request("FM_jnr")
                
                 call alfanumerisk(strjnr)
                 strjnr = alfanumeriskTxt
                 strjnr = left(strjnr,20)

				  tlbnr = request("FM_tnr")  
				    
				end if
				
				
                 '************************************'
                 '***** Forrretningsområder **********'
                 '************************************'

                   


                    fomrArr = split(request("FM_fomr"), ",")

                    for_faktor = 0
                    'for afor = 0 to UBOUND(fomrArr)
                    'for_faktor = for_faktor + 1 
                    'next

                    'if for_faktor <> 0 then
                    'for_faktor = for_faktor
                    'else
                    'for_faktor = 1
                    'end if

                    'for_faktor = formatnumber(100/for_faktor, 2)
                    'for_faktor = replace(for_faktor, ",", ".")

                    'if len(trim(request("FM_fomr_syncakt"))) <> 0 then
                    'syncForAkt = 1
                    'else
                    syncForAkt = 0
                    'end if
              

                 '****************************************
				
				
				'*********************************************************************'
				'************************** Internt / eksternt job *******************'
				'*********************************************************************'
				'*** Altid Eksternt job ***'
						
						if request("FM_usetilbudsnr") = "j" then
						strStatus = 3 'tilbud
						else
							if request("FM_status") <> 3 then
							strStatus = request("FM_status")
							else
							strStatus = 1 'aktivt, hvis der ved en fejl at valgt tilbud
							end if
						end if
				


                
                '*******************************************'		
                '** Projektgrupper ****
                '*******************************************'		
				
                
                '*** Nedarv / Fød ***'
                '** 0 = Nedarv ******'
                '** 1 = Fød *********'

                pgrp_arvefode = request("pgrp_arvefode") ''Tilføjes længere ned hvor Stam-aktiviteter tilføjes
                
                '*** Progrp på job ***'
                
                strProjektgr1 = request("FM_projektgruppe_1")
				strProjektgr2 = request("FM_projektgruppe_2")
				strProjektgr3 = request("FM_projektgruppe_3")
				strProjektgr4 = request("FM_projektgruppe_4")
				strProjektgr5 = request("FM_projektgruppe_5")
				strProjektgr6 = request("FM_projektgruppe_6")
				strProjektgr7 = request("FM_projektgruppe_7")
				strProjektgr8 = request("FM_projektgruppe_8")
				strProjektgr9 = request("FM_projektgruppe_9")
				strProjektgr10 = request("FM_projektgruppe_10")

                

                'Response.write "strProjektgr1:" & strProjektgr1 &" strProjektgr2: "& strProjektgr2 &" strProjektgr3: "& strProjektgr3&" strProjektgr4: "& strProjektgr4&" strProjektgr5: "& strProjektgr5&" strProjektgr6: "& strProjektgr6&" strProjektgr7: "& strProjektgr7&" strProjektgr8: "& strProjektgr8&" strProjektgr9: "& strProjektgr9&" strProjektgr10: "& strProjektgr10
                'Response.end

                if request("FM_gemsomdefault") = "1" then
				response.cookies("job")("defaultprojgrp") = strProjektgr1
				response.cookies("job").expires = date + 65
				end if

              
				
				strFakturerbart = request("FM_fakturerbart")
				
					startDato = Request("FM_start_aar") &"/" & Request("FM_start_mrd") & "/" & strStartDay 
					if request("FM_datouendelig") = "j" then
					slutDato =  "2044/1/1" 
					else
					slutDato =  Request("FM_slut_aar") &"/" & Request("FM_slut_mrd") & "/" & strSlutDay 
					end if
				
				'************************************************************************************
				
				    
				    
				    '*** Jobnr findes ***
				    jobnrFindes = 0
				    
				    if func = "dbred" then
				    strSQL = "SELECT jobnr, id FROM job WHERE id <> "& id &" AND jobnr = '" & strjnr & "'"
					else
					strSQL = "SELECT jobnr, id FROM job WHERE jobnr = '" & strjnr &"'"
					end if 
					
					'Response.Write strSQL
					'Response.Flush
					
					
					oRec5.open strSQL, oConn, 3
		            if not oRec5.EOF then	
		            
		            jobnrFindesNR = oRec5("jobnr")
		            jobnrFindesID = oRec5("id")
		            jobnrFindes = 1 
		            
		            end if
					oRec5.close
					
					'Response.Write "<br> jobnrFindes" &  jobnrFindes & " "& jobnrFindesNR &" "& jobnrFindesID &"<br>"
					'Response.flush
					
					if cint(jobnrfindes) = 1 then
					%>
					<!--#include file="../inc/regular/header_inc.asp"-->
				    <%	
					errortype = 93
					call showError(errortype)
					Response.end
				    end if 
					
				    
					'*** tilbudsnr findes ***
					
					if request("FM_usetilbudsnr") = "j" then
					tilbudsnrFindes = 0
					
				    if func = "dbred" then
				    strSQL = "SELECT tilbudsnr, id FROM job WHERE id <> "& id &" AND tilbudsnr = '" &  tlbnr & "' AND tilbudsnr <> 0"
					else
					strSQL = "SELECT tilbudsnr, id FROM job WHERE tilbudsnr = '" & tlbnr & "'"
					end if 
					
					'Response.Write strSQL
					'Response.Flush
					'Response.end
					
					oRec5.open strSQL, oConn, 3
		            if not oRec5.EOF then	
		            
		            tilbudsnrFindesNR = oRec5("tilbudsnr")
		            tilbudsnrFindes = 1 
		            
		            end if
					oRec5.close
					
					end if 
					
					
					if cint(tilbudsnrFindes) = 1 then
					%>
					<!--#include file="../inc/regular/header_inc.asp"-->
				    <% 
					errortype = 94
					call showError(errortype)
					Response.end
				    end if 
					
					
					
					 if cint(jobnrFindes) <> 1 AND cint(tilbudsnrFindes) <> 1 then		
					 
					 'Response.Write "her oprettes"
					 		
					 	'*** Opretter Job ***
					 	if func = "dbopr" then  
					 	
					 	
							    strSQLjob = ("INSERT INTO job (jobnavn, jobnr, jobknr, jobTpris, jobstatus, jobstartdato,"_
							    &" jobslutdato, editor, dato, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, "_
							    &" projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, "_
							    &" fakturerbart, budgettimer, fastpris, kundeok, beskrivelse "_
							    &", ikkeBudgettimer, tilbudsnr, jobans1, jobans2, "_
							    &" serviceaft, kundekpers, lukafmjob, valuta, jobfaktype, rekvnr, "_
							    &" jo_gnstpris, jo_gnsfaktor, jo_gnsbelob, jo_bruttofortj, jo_dbproc, udgifter, "_
							    &" risiko, usejoborakt_tp, ski, job_internbesk, abo, ubv, sandsynlighed, jobans3, jobans4, jobans5, "_
                                &" diff_timer, diff_sum, jo_udgifter_intern, jo_udgifter_ulev, jo_bruttooms, "_
                                &" jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, restestimat, stade_tim_proc, "_
                                &" virksomheds_proc, syncslutdato, altfakadr, preconditions_met, laasmedtpbudget,"_
                                &" salgsans1, salgsans2, salgsans3, salgsans4, salgsans5, "_
                                &" salgsans1_proc, salgsans2_proc, salgsans3_proc, salgsans4_proc, salgsans5_proc, filepath1, fomr_konto, jfak_sprog, jfak_moms "_
                                &") VALUES "_
							    &"('"& strNavn &"', "_
							    &"'"& strjnr &"', "_ 
							    &""& oRec("kid") &", "_  
							    &""& jo_gnsbelob &", "_ 
							    &""& strStatus &", "_ 
							    &"'"& startDato &"', "_ 
							    &"'"& slutDato &"', "_
							    &"'"& strEditor &"', "_
							    &"'"& strDato &"', "_ 
							    &""& strProjektgr1 &", "_ 
							    &""& strProjektgr2 &", "_ 
							    &""& strProjektgr3 &", "_ 
							    &""& strProjektgr4 &", "_ 
							    &""& strProjektgr5 &", "_
							    &""& strProjektgr6 &", "_ 
							    &""& strProjektgr7 &", "_ 
							    &""& strProjektgr8 &", "_ 
							    &""& strProjektgr9 &", "_ 
							    &""& strProjektgr10 &", "_  
							    &""& strFakturerbart &", "_ 
							    &""& strBudgettimer  &", "_ 
							    &"'"& strFastpris & "',"_
							    &""& intkundese &", "_
							    &"'"& strBesk &"', "& SQLBless(ikkeBudgettimer) &", "_
							    &"'"& tlbnr &"', "& intJobans1 &", "& intJobans2 &", "_
							    &" "& intServiceaft &", "& intKundekpers &", "& lukafmjob &", "_
							    &" "& valuta &", "& jobfaktype &", '"& rekvnr &"', "_
							    &" "& jo_gnstpris &", "& jo_gnsfaktor &", "_
							    &" "& jo_gnsbelob &","& jo_bruttofortj &","& jo_dbproc &", "& udgifter &", "_
							    &" "& intprio &", "& usejoborakt_tp &", "& ski &", '"& strInternBesk &"',"_
							    &" "& abo &", "& ubv &", "& sandsynlighed &", "& intJobans3  &", "& intJobans4  &", "& intJobans5  &", "_
                                &" "& diff_timer &", "& diff_sum &", "& jo_udgifter_intern &", "& jo_udgifter_ulev &", "& strBudget &", "_
                                &" "& jobans_proc_1 & ", "& jobans_proc_2 & ", "& jobans_proc_3 & ", "& jobans_proc_4 & ", "& jobans_proc_5 & ", "& restestimat &", "& stade_tim_proc &","_
                                &" "& virksomheds_proc &", "& syncslutdato &", "& altfakadr &", "& preconditions_met &", "& laasmedtpbudget &", "_
                                &" "& salgsans1 &","& salgsans2 &","& salgsans3 &","& salgsans4 &","& salgsans5 &", "_
                                &" "& salgsans_proc_1 &","& salgsans_proc_2 &","& salgsans_proc_3 &","& salgsans_proc_4 &","& salgsans_proc_5 &", '"& filepath1 &"', "& fomr_konto &", "& jfak_sprog &", "& jfak_moms &""_
                                &")")
    							
							    'Response.write strFakturerbart & "<br><br>"
							    'Response.write strSQLjob
							    'Response.end 
    							
							    oConn.execute(strSQLjob)	
								
								
								'******************************************'
								'*** finder jobid på det netop opr. job ***'
								'******************************************'
                                't = 10
								'if t = 0 then
								
								    strSQL2 = "SELECT id FROM job WHERE id <> 0 ORDER BY id DESC " 'jobnr = " & strjnr
									oRec5.open strSQL2, oConn, 3
									if not oRec5.EOF then
									varJobIdThis = oRec5("id")
									varJobId = varJobIdThis
									end if
									oRec5.close
									
									
								'***************************************'
								'*** Opretter timepriser på job.       *'
								'***************************************'
                                
						        'Response.write "HER"
                                'response.end
                        
                                'timeE = now
                                'loadtime = datediff("s",timeA, timeE)
							    'Response.write "<br>Før timepriser: "& loadtime & "<br><br>"		


								strSQLpg = "SELECT id, navn FROM projektgrupper WHERE "_
								&" id = "& strProjektgr1 &" OR "_
								&" id = "& strProjektgr2 &" OR id = "& strProjektgr3 &" OR "_
								&" id = "& strProjektgr4 &" OR id = "& strProjektgr5 &" OR id = "& strProjektgr6 &" OR "_
								&" id = "& strProjektgr7 &" OR id = "& strProjektgr8 &" OR id = "& strProjektgr9 &" OR "_
								&" id = "& strProjektgr10 &" ORDER BY navn"
								oRec5.open strSQLpg, oConn, 3
								
								'Response.write strSQL
								while not oRec5.EOF
									
									strSQLmtp = "SELECT medarbejderid, projektgruppeid, mid, mnavn, timepris, timepris_a1, "_
									&" timepris_a2, timepris_a3, timepris_a4, timepris_a5, "_
									&" tp0_valuta, tp1_valuta, tp2_valuta, tp3_valuta, tp4_valuta, tp5_valuta "_
									&" FROM progrupperelationer "_
									&" LEFT JOIN medarbejdere ON (mid = progrupperelationer.medarbejderid) "_
									&" LEFT JOIN medarbejdertyper ON (medarbejdertyper.id = medarbejdertype) "_
									&" WHERE projektgruppeid = "& oRec5("id") &" AND mnavn <> '' AND mansat <> 2 ORDER BY mnavn"
									oRec3.open strSQLmtp, oConn, 3
									this6timepris = 0
									while not oRec3.EOF
								        
								        
								        '*** Finder valuta på job og finder timepris ***'
								        '*** på medarbejder der matcher valuta ***'
								        '*** Hvis findes ellers vælges 1 = DKK, Grundvaluta ****'
								            
								           valutafundet = 0
								             
								           for i = 0 to 5
								           
								                   'Response.Write "her" & i & "<br>"
								           
								                   select case i
								                   case 0
								                   tpris = oRec3("timepris")
								                   tprisValuta = oRec3("tp0_valuta")
								                   tprisAlt = 0
								                   case 1
								                   tpris = oRec3("timepris_a1")
								                   tprisValuta = oRec3("tp1_valuta")
								                   tprisAlt = 1
								                   case 2
								                   tpris = oRec3("timepris_a2")
								                   tprisValuta = oRec3("tp2_valuta")
								                   tprisAlt = 2
								                   case 3
								                   tpris = oRec3("timepris_a3")
								                   tprisValuta = oRec3("tp3_valuta")
								                   tprisAlt = 3
								                   case 4
								                   tpris = oRec3("timepris_a4")
								                   tprisValuta = oRec3("tp4_valuta")
								                   tprisAlt = 4
								                   case 5
								                   tpris = oRec3("timepris_a5")
								                   tprisValuta = oRec3("tp5_valuta")
								                   tprisAlt = 5
								                   end select
								           
                                                    'if lto = "cisu" then
                                
                                                    '    response.write "<br>Medid: "& oRec3("medarbejderid") &":"& oRec3("mid") &" projektgruppeID: "& oRec5("id") &" i: "& i &" valuta: " & valuta &" tprisValuta = "& tprisValuta &" valutafundet: "& valutafundet
                                                    '    response.flush 

                                                    'end if

								           
								                    if cint(valuta) = cint(tprisValuta) AND cint(valutafundet) = 0 then
								                    tprisUse = tpris
								                    tprisValutaUse = tprisValuta
								                    tprisAltUse = tprisAlt
								                    valutafundet = 1
								            
								                    end if
								            
							                next
								           
								           if cint(valutafundet) = 0 then
								           tprisUse = oRec3("timepris")
								           tprisValutaUse = oRec3("tp0_valuta")
								           tprisAltUse = 0
								           end if
								            
                                           tprisUse = replace(tprisUse, ".", "")
                                           tprisUse = replace(tprisUse, ",", ".")
								        
								        '*** Opretter timepriser på job **'
								        if instr(usedmids, "#"&oRec3("mid")&"#") = 0 then
								        strSQLtp = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) "_
								        &" VALUES ("& varJobId &", 0, "& oRec3("mid") &","& tprisAltUse &", "& tprisUse &", "& tprisValutaUse &")"
								        
                                        'if lto = "qwert" then
								        'Response.Write strSQLtp & "<br>"
								        'Response.flush
                                        'end if
								        
								        oconn.execute(strSQLtp)
								        end if
								   
								   usedmids = usedmids &oRec3("mid")&"#"
								   oRec3.movenext     
								   wend
								   oRec3.close
							    
							    
						     oRec5.movenext     
							 wend
							 oRec5.close
							   
                             'end if 't = 0
                            
                              'if lto = "qwert" then
							  'Response.end 
                              'end if 
								
                            'timeC = now
                            'loadtime = datediff("s",timeA, timeC)
							'Response.write "<br>Før stam: "& loadtime & "<br><br>"


					            '******************************************************************'
								'Tilknytter stam aktiviteter til det job der bliver oprettet her.**'
								'******************************************************************'
								
								intAktfavgp_use = split(intAktfavgp, ",")
								strAktFase_use = split(strAktFase, ", ")
                                firstLoop = 0
								for a = 0 to UBOUND(intAktfavgp_use)
								'for a = 1 to 1
								
                                    if len(trim(intAktfavgp_use(a))) <> 0 then
								    call tilknytstamakt(a, intAktfavgp_use(a), trim(strAktFase_use(1)), 0)
								
								        'if a = 0 then
									    'intAktfavgp_1 = intAktfavgp_use(a)
									    'Response.Write "intAktfavgp_1 "& intAktfavgp_1
									    'Response.end
									    'end if
								
                                    end if

								next
								
                                'Response.write "Aktiviterer tilknyttet"
                                ' 
                                'Response.end


                                'timeD = now
                                'loadtime = datediff("s",timeA, timeD)
							    'Response.write "<br>Efter stam: "& loadtime & "<br><br>"


                                '**** Indlæser Underlev grp / Salgsomkostninger **'
                                call addUlev
                               

							
								
				                '*************************************************'
			                    '**** Opdaterer jobnr og tilbuds nr i Licens *****'
				                '*************************************************'
				                nytjobnr = strjnr
				                tilbudsnrKri = ""
				                
				                if request("FM_usetilbudsnr") = "j" then
				                nyttlbnr = tlbnr
				                tilbudsnrKri = ", tilbudsnr = "& nyttlbnr &""
				                end if
				                
				                strSQL = "UPDATE licens SET jobnr = "&  nytjobnr &" "& tilbudsnrKri &" WHERE id = 1"
				                oConn.execute(strSQL)
				               
				                
				                
								
							else '** REDIGER
							


                             '*** Opdater LUKKE dato (før det skifter stastus) ***'
                             call lukkedato(id, strStatus)



							strSQL = "UPDATE job SET "_
							&" jobnavn = '"& strNavn &"', "_
							&" jobnr = '"& strjnr &"', "_ 
							&" jobknr = "& oRec("kid") &", "_  
							&" jobTpris = "& jo_gnsbelob &", "_ 
							&" jobstatus = "& strStatus &", "_ 
						    &" jobstartdato = '"& startDato &"', "_ 
							&" jobslutdato = '"& slutDato &"', "_
							&" editor = '"& strEditor &"', "_
							&" dato = '"& strDato &"', "_
							&" projektgruppe1 = "& strProjektgr1 &", "_ 
							&" projektgruppe2 = "& strProjektgr2 &", "_ 
							&" projektgruppe3 = "& strProjektgr3 &", "_ 
							&" projektgruppe4 = "& strProjektgr4 &", "_
							&" projektgruppe5 = "& strProjektgr5 &", "_
							&" projektgruppe6 = "& strProjektgr6 &", "_ 
							&" projektgruppe7 = "& strProjektgr7 &", "_ 
							&" projektgruppe8 = "& strProjektgr8 &", "_ 
							&" projektgruppe9 = "& strProjektgr9 &", "_
							&" projektgruppe10 = "& strProjektgr10 &", "_
							&" fakturerbart= "& strFakturerbart &", "_
							&" Budgettimer  = "& strBudgettimer &", "_
							&" fastpris = '"& strFastpris & "', kundeok = "& intkundese &", "_
							&" beskrivelse = '"& strBesk &"', ikkeBudgettimer = "& SQLBless(ikkeBudgettimer) &", "_
							&" jobans1 =  "& intJobans1 &", "_
							&" jobans2 = "& intJobans2 &", serviceaft = "& intServiceaft &", "_
							&" kundekpers = "& intKundekpers &", lukafmjob = "& lukafmjob &", "_
							&" valuta = "& valuta &", jobfaktype = "& jobfaktype &", rekvnr = '"& rekvnr &"', "_
							&" jo_gnstpris = "& jo_gnstpris &", jo_gnsfaktor = "& jo_gnsfaktor &", "_
							&" jo_gnsbelob = "& jo_gnsbelob &", jo_bruttofortj = "& jo_bruttofortj &", "_
							&" jo_dbproc = "& jo_dbproc &", udgifter = "& udgifter &", "_
							&" risiko = "& intprio &", usejoborakt_tp = "& usejoborakt_tp &", ski = "& ski &", "_
							&" job_internbesk = '"& strInternBesk &"', abo = "& abo &", ubv = "& ubv &", "_
							&" sandsynlighed = "& sandsynlighed &", tilbudsnr = '"& tlbnr &"', "_
                            &" jobans3 = "& intJobans3 &", jobans4 = "& intJobans4 &", jobans5 = "& intJobans5 &", "_
                            &" diff_timer = "& diff_timer &", diff_sum = "& diff_sum &", "_
                            &" jo_udgifter_intern = "& jo_udgifter_intern &", jo_udgifter_ulev = "& jo_udgifter_ulev &", jo_bruttooms = "& strBudget &", "_
                            &" jobans_proc_1 = "& jobans_proc_1 & ", jobans_proc_2 = "& jobans_proc_2 & ", jobans_proc_3 = "& jobans_proc_3 & ", jobans_proc_4 = "& jobans_proc_4 & ", jobans_proc_5 = "& jobans_proc_5 & ", "_
                            &" restestimat = "& restestimat &", stade_tim_proc = "& stade_tim_proc &", virksomheds_proc = "& virksomheds_proc &", "_
                            &" syncslutdato = "& syncslutdato &", altfakadr = "& altfakadr &", preconditions_met = "& preconditions_met &", laasmedtpbudget = "& laasmedtpbudget &", "_
                            &" salgsans1 = "& salgsans1 &", salgsans2 = "& salgsans2 &", salgsans3 = "& salgsans3 &", salgsans4 = "& salgsans4 &", salgsans5 = "& salgsans5 &", "_
                            &" salgsans1_proc = "& salgsans_proc_1 &", salgsans2_proc = "& salgsans_proc_2 &", salgsans3_proc = "& salgsans_proc_3 &", salgsans4_proc = "& salgsans_proc_4 &", "_
                            &" salgsans5_proc = "& salgsans_proc_5 &", filepath1 = '"& filepath1 &"', fomr_konto = "& fomr_konto &", jfak_sprog = "& jfak_sprog &", jfak_moms = "& jfak_moms &""_
							&" WHERE id = "& id 
							
							'Response.Write strSQL
							'Response.end
                            'Response.flush							
							oConn.execute(strSQL)
							                    




                                                '****** opdaterer tilbudsnr ved rediger ****'
                                                if request("FM_usetilbudsnr") = "j" then
				                                nyttlbnr = tlbnr
                                                     
                                                     senesteTilbnr = 0
                                                     strSQL = "SELECT tilbudsnr FROM licens WHERE id = 1"
                                                     oRec5.open strSQL, oConn, 3
                                                     if not oRec5.EOF then
                                                     senesteTilbnr = oRec5("tilbudsnr")
                                                     end if
                                                     oRec5.close

				                                end if
				                                
                                                if cdbl(nyttlbnr) <> 0 AND cdbl(nyttlbnr) >= cdbl(senesteTilbnr) then
				                                strSQL = "UPDATE licens SET tilbudsnr = "& nyttlbnr &" WHERE id = 1"
				                                oConn.execute(strSQL)
				                                end if

                                
                                                '***********************************'
                                                '** Opdaterer kundeoplysninger *****'
                                                '***********************************'

								
								                '*** Overfører gamle timeregistreringer til ny aftale (hvis der skiftes aftale) **'
								                if request("FM_overforGamleTimereg") = "1" then
                								
								                '*** Opdaterer timereg tabellen ****
								                strSQLtimer = "UPDATE timer SET "_
								                & " Tknavn = '"& replace(oRec("kkundenavn"), "'", "''") &"', Tknr = "& oRec("kid")&", "_
								                & " Tjobnavn = '"& strNavn &"', "_
								                & " Tjobnr = '"& strjnr &"', "_
								                & " fastpris = '"& strFastpris &"', "_
								                & " seraft = "& intServiceaft &" "_
								                & " WHERE Tjobnr = '"& strOLDjobnr & "'"
                								
                								'** Husk materiale forbrug ***
								                strSQLmat_forbrug = "UPDATE materiale_forbrug SET serviceaft = " & intServiceaft &""_
								                & " WHERE jobid = " & id
                								
								                oConn.execute(strSQLmat_forbrug)
                								
                								
								                else
                								
								                '*** Opdaterer timereg tabellen ****
								                strSQLtimer = "UPDATE timer SET "_
								                & " Tknavn = '"& replace(oRec("kkundenavn"), "'", "''") &"', Tknr = "& oRec("kid") &", "_
								                & " Tjobnavn = '"& strNavn &"', "_
								                & " Tjobnr = '"& strjnr &"', "_
								                & " fastpris = '"& strFastpris &"' "_
								                & " WHERE Tjobnr = '"& strOLDjobnr & "'"
                								
								                end  if
                								
                								
								                'Response.write strSQLtimer
								                'Response.flush
                								
								                oConn.execute(strSQLtimer)
								                
								                
								                '*** Opdaterer faktura tabel så faktura kunde id passer hvis der er skiftet kunde  ved rediger job.
								                '*** Adr. i adresse felt på faktura behodes til revisor spor. **'
								                strSQLFakadr = "UPDATE fakturaer SET fakadr = "& oRec("kid") &" WHERE jobid = " & id
								                oConn.execute(strSQLFakadr)
								
								
								varJobId = id
								

                               


                                '***** Sync Job slutDato ***********'
                                if strStatus = 0 then 'Kun ved luk job
                                syncslutdato = syncslutdato
                                else
                                syncslutdato = 0
                                end if

                                call syncJobSlutDato(varJobId, strjnr, syncslutdato)

                               
                                '****** Sync. datoer på aktiviteter *******'
                                if syncaktdatoer = 1 then
                                    
                                    '** Hvis sync jobslutdato er valgt ***
                                    if syncslutdato <> 1 then
                                    useSlutDato = slutDato
                                    else
                                    useSlutDato = useSyncDato 
                                    end if

                                strSQLaDatoer = "UPDATE aktiviteter SET aktstartdato = '"& startDato &"', aktslutdato = '"& useSlutDato &"' WHERE job = "& varJobId
                                oConn.execute(strSQLaDatoer)

                                end if

                               






								
				                '************************************************'
				                '**** Diffentierede timepriser ******************'
			                    '*** Opdater eksisterende time-registreringer ***'
			                    '************************************************'
                				
                                '** sync alle aktiviteter ***'
                		        if request("FM_sync_tp") = "1" then
                                syncAkt = 1
                                else
                                syncAkt = 0
                                end if

                                 


				                medarb_tpris = request("FM_use_medarb_tpris")
				                Dim intMedArbID 
				                Dim b
                				strMedabTimePriserSlet = " AND medarbid <> 0"

					                intMedArbID = Split(medarb_tpris, ", ")
					                For b = 0 to Ubound(intMedArbID)
                                        
                                         
                                         strMedabTimePriserSlet = strMedabTimePriserSlet & " AND medarbid <> "& intMedArbID(b)


                                        if cint(syncAkt) = 1 then
                                            
                                            '** ikke KM **
                                            fakbaraktid = ""
                                            strSQLatyp = "SELECT id, fakturerbar FROM aktiviteter WHERE job = " & id & " AND fakturerbar = 5"
                                            oRec5.open strSQLatyp, oConn, 3
                                            fb = 0
                                            while not oRec5.EOF 
                                                
                                                if fb <> 0 then
                                                fakbaraktid = fakbaraktid & " OR "
                                                end if

                                                if oRec5("fakturerbar") = 5 then
                                                fakbaraktid = fakbaraktid & "aktid <> "& oRec5("id") 
                                                else
                                                fakbaraktid = fakbaraktid 
                                                end if

                                                fb = fb + 1
                                            oRec5.movenext
                                            wend
                                            oRec5.close

                                            if len(trim(fakbaraktid)) <> 0 then
                                            fakbaraktid = " AND ("& fakbaraktid & ")"
                                            else
                                            fakbaraktid = ""
                                            end if



						                oConn.execute("DELETE FROM timepriser WHERE jobid = "& id &" AND medarbid = "& intMedArbID(b) &" "& fakbaraktid &"")
                                        
                                        else '** ellers kun job // Eller valgte akt **'
                                        oConn.execute("DELETE FROM timepriser WHERE jobid = "& id &" AND medarbid = "& intMedArbID(b) &" AND aktid = "& hd_tp_jobaktid &"")
                                        end if

                                         
						               
                                        'if len(request("FM_6timepris_"&intMedArbID(b)&"")) <> 0 OR jobids(j) = 0 then
                						
						                if len(request("FM_6timepris_"&intMedArbID(b)&"")) <> 0 then
						                call erDetInt(request("FM_6timepris_"&intMedArbID(b)&""))
						                    if isInt > 0 then
						                    this6timepris = 0
						                    else
						                    this6timepris = SQLBless(replace(request("FM_6timepris_"&intMedArbID(b)&""),".",""))
						                    end if
						                
						                valuta6 = request("FM_valuta_600"& intMedArbID(b))
						                
						                else
						                this6timepris = 0
						                valuta6 = 1
						                end if
						               
                						
                                       ' if hd_tp_jobaktid = 0 then '** Job / Akt, Hviklen radio bt er valgt ***'
                                       '*** Ved sync nedarver alle og der skal derfor ikke indsættes en record heller ikke for den aktivitet man står på ***'
                                       '**** Kun hvis man står på selve jobbet ***'

                                        if len(request("FM_6timepris_"&intMedArbID(b)&"")) <> 0 AND (cint(syncAkt) <> 1 OR (cint(syncAkt) = 1 AND hd_tp_jobaktid = 0)) then

                                            radiobtVal = SQLBless(replace(request("FM_hd_timepris_"&intMedArbID(b)&"_"&request("FM_timepris_"&intMedArbID(b)&"")),".",""))

                						    if this6timepris = radiobtVal then
                						    tprisalt = request("FM_timepris_"&intMedArbID(b)&"")
                						    else
                						    tprisalt = 6
                                            end if
                						
                                            
                                            
						                    strSQLupdTimePriser = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) "_
						                    &" VALUES ("& id &","& hd_tp_jobaktid &", "& intMedArbID(b) &", "& tprisalt &", "& this6timepris &", "& valuta6 &")"
                						
                                       
                                            'Response.Write strSQLupdTimePriser & "<br>"
                                            'Response.flush

                                             oConn.execute(strSQLupdTimePriser)

                                       
                                       

                                         end if
                                        






                                        
                                        '*********************************************************'
						                '** Opdaterer timereg (kun på aktiviteter der nedarver) **'
                                        '** Dvs dem der ikke findes i timepriser tabellen       **'

                                        '*** Else den specifikke aktivitet ***
						               
						                '*** Kræver at der minimun findes 1 akt. *****************
						              	'*** Ved sync er denne tom (lige slettet ovenfor) og alle **'
                                        '**' timeprsier på alle aktiviter forall medeb. opdateres ***'
                                        '*********************************************************'


                                        if cint(syncAkt) = 1 OR hd_tp_jobaktid = 0 then
										strSQLtimp = "SELECT aktid FROM timepriser WHERE jobid = "& id &" AND medarbid = "& intMedArbID(b) &" AND aktid <> 0"
										
										'Response.Write strSQLtimp & "<br>"
										'Response.flush
										
										notAkt = ""
										oRec5.open strSQLtimp, oConn, 3
										while not oRec5.EOF
										notAkt = notAkt & " AND taktivitetId <> "& oRec5("aktid") 
										oRec5.movenext
										wend
										oRec5.close 

                                        specifikakt = 0

                                        else

                                        specifikakt = 1
                                        notAkt = " AND taktivitetId = "& hd_tp_jobaktid 
                                        
                                        aktnavn = ""
                                        strSQLalleaktmnavn = "SELECT navn FROM aktiviteter WHERE id = "& hd_tp_jobaktid 
                                        oRec6.open strSQLalleaktmnavn, oConn, 3
                                        if not oRec6.EOF then
                                        aktnavn = oRec6("navn")
                                        end if
                                        oRec6.close

                                        end if
										
										

                                        '*** Opdaterer alle timeregistreringer på valgte akt. eller på alle aktiviteter ved sync ***'

										   '**** Finder aktuel kurs ***'
						                   strSQL = "SELECT kurs FROM valutaer WHERE id = " & valuta6
						                   oRec5.open strSQL, oConn, 3
        						            
						                   if not oRec5.EOF then
						                   nyKurs = replace(oRec5("kurs"), ",", ".")
						                   end if 
						                   oRec5.close

                                         '** Fra dato ved opdater timepriser ***'
                                        fraDato = request("FM_opdatertpfra") 
                                        if isDate(fraDato) = false then
                                        fraDato = year(now) &"/"& month(now) & "/"& day(now)
                                        else
                                        fraDato = year(fraDato) &"/"& month(fraDato) & "/"& day(fraDato) 
                                        end if

										
					                    
					                    
                                        if request("FM_opdateralleakt") = "1" AND specifikakt = 1 then 
                                        '** Specifik akt + alle med samme navn **' 
			                            strSQL = "UPDATE timer SET timepris = "& this6timepris  &", valuta = "& valuta6 &", kurs = "& nyKurs &""_
			                            &" WHERE (tmnr = "& intMedArbID(b) &" AND tdato >= '"& fraDato &"' AND (tfaktim <> 5)) AND taktivitetnavn = '" & aktnavn &"'"
                                        else
                                        '** Nedarv (IKKE kørsel) / specifik akt 
			                            strSQL = "UPDATE timer SET timepris = "& this6timepris  &", valuta = "& valuta6 &", kurs = "& nyKurs &""_
			                            &" WHERE (tmnr = "& intMedArbID(b) &" AND tjobnr = '"& strjnr &"' AND tdato >= '"& fraDato &"' AND (tfaktim <> 5)) " & notAkt
					                    end if

                                        'Response.Write strSQL & "<br>"
					                    'Response.end
					                    
					                    oConn.execute(strSQL)
                                                

                             next '** intMedArbID(b)




                            '*************************************************************************************
                            '*** Renser ud i timepriser på medarbejdere der ikke længere er tilknyttet jobbbet ***
                            '*************************************************************************************
                                       
                            '*** Pilot WWF 26-09-2011 *** 
                            'if lto = "wwf" then
                            strSQLRensTp = "DELETE FROM timepriser WHERE jobid = "& id &" "& strMedabTimePriserSlet 
                            'Response.Write strSQLRensTp
                            oConn.execute(strSQLRensTp)
                            'end if

                				
                				'Response.end
                                '**** Timepriser END ***'

				                
				                '**********************************************************'
								'*** Opdaterer allerede tilknyttet aktiviteter   **********'
								
								
								    if len(request("FM_aktnavn")) > 3 then
	                                len_aktnavn = len(request("FM_aktnavn"))
	                                left_aktnavn = left(request("FM_aktnavn"), len_aktnavn - 3)
	                                strAktnavn = left_aktnavn
	                                else
	                                strAktnavn = ""
	                                end if
                                	
	                                aktnavn = strAktnavn
	                                akttimer = request("FM_akttimer")

                                    if len(trim(request("FM_aktkonto"))) <> 0 then
                                    aktkonto = request("FM_aktkonto")
                                    else
                                    aktkonto = ""
                                    end if

	                                aktantalstk = request("FM_aktantalstk")
	                                aktpris = request("FM_aktpris")
	                                
	                                'Response.Write request("FM_aktpris")
	                                'Response.end
	                                aktids = request("FM_aktid")

                                    'Response.Write "aktids" & aktids
	                                'Response.end

	                                aktstatus = request("FM_aktstatus")
                                	
                                    'Response.Write request("FM_aktfase")
                                    'Response.end

	                                if len(request("FM_aktfase")) > 3 then
	                                len_Fasenavn = len(request("FM_aktfase"))
	                                left_Fasenavn = left(request("FM_aktfase"), len_Fasenavn - 3)
	                                strFasenavn = left_Fasenavn
	                                else
	                                strFasenavn = ", #"
	                                end if
                                	
	                                aktfaser = strFasenavn
	                                aktbgr = request("FM_aktbgr")
                                	
	                                'Response.Write request("FM_akt_totpris")
	                                'Response.end
                                	
	                                akttotpris = request("FM_akttotpris")
                                	
	                                'Response.Write request("FM_fase")
	                                'Response.end
                                	
	                                aktslet = request("FM_slet_aid")
                                	
	                                'Response.Write aktSlet
	                                'Response.end
								    
								    aktslet_aids = 0
								
								    call opdateraktliste(varJobId, aktids, aktnavn, akttimer, aktantalstk, aktfaser, aktbgr, aktpris, aktstatus, akttotpris, aktslet, aktslet_aids, aktkonto)

								
								'Response.end
								
								'**********************************************************'
								'**** Tilknytter flere Stam-aktivitets grupper til det  ***'
								'**** job der bliver redigeret                          ***'
								'**********************************************************'
								
                                'Response.write "intAktfavgp " & intAktfavgp & "<br>"
                                'Response.write "strAktFase " & strAktFase & "<br>"
                                'Response.end

								intAktfavgp_use = split(intAktfavgp, ",")
								strAktFase_use = split(strAktFase, ", ")
                                firstLoop = 0
								for a = 0 to UBOUND(intAktfavgp_use)
								'for a = 1 to 1
                                    'Response.Write "a" & a & " val: "& intAktfavgp_use(a) &"<br>"
                                    'Response.flush
                                    if len(trim(intAktfavgp_use(a))) <> 0 then
								    call tilknytstamakt(a, intAktfavgp_use(a), trim(strAktFase_use(1)), 0)
                                    end if
							
                                next
                                
                                'Response.write "her"
                                'Response.end

                             
								
								                

                                                 '*** Opdaterer projektgrupper på Eksisterende aktiviteter ved redigering af job ***'
                                                 '*** Så de følger jobbet. = Sync (nedarv på akt.) ***'
                    								
					                                if request("FM_opdaterprojektgrupper") = 1 then
                                                    
                                                    'Response.write "her"
                                                    'Response.end 
                                                    	
					                                strSQL = "SELECT projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM job WHERE id = "& varJobId 
                    	
					                                oRec5.open strSQL, oConn, 3
					                                if not oRec5.EOF then
                    									
					                                    oConn.execute("UPDATE aktiviteter SET"_
					                                    &" projektgruppe1 = "& oRec5("projektgruppe1") &" , projektgruppe2 = "& oRec5("projektgruppe2") &", "_
					                                    &" projektgruppe3 = "& oRec5("projektgruppe3") &", "_
					                                    &" projektgruppe4 = "& oRec5("projektgruppe4") &", "_
					                                    &" projektgruppe5 = "& oRec5("projektgruppe5") &", "_
					                                    &" projektgruppe6 = "& oRec5("projektgruppe6") &", "_
					                                    &" projektgruppe7 = "& oRec5("projektgruppe7") &", "_
					                                    &" projektgruppe8 = "& oRec5("projektgruppe8") &", "_
					                                    &" projektgruppe9 = "& oRec5("projektgruppe9") &", "_
					                                    &" projektgruppe10 = "& oRec5("projektgruppe10") &" WHERE job = "& varJobId &"")
                		
					                                end if
					                                oRec5.close
                    									
					                                end if
								                   
								
								
							                    '**** Indlæser Underlev grp / Salgsomkostninger **'
                                                call addUlev
							
							
							
							                   '***** Medarbejertyper timebudget ****'
                                               
                                               strSQLmtyp_tpDEL = "DELETE FROM medarbejdertyper_timebudget WHERE jobid = "& varJobId
                                               oConn.execute(strSQLmtyp_tpDEL)

                                               'Response.write request("FM_mtype_id") & "<hr>"
                                               
                                               mtype_id_arr = split(request("FM_mtype_id"), ",")
                                               


                                               for t = 0 to UBOUND(mtype_id_arr) 
                                                        
                                                        
                                                        tpy_id = trim(mtype_id_arr(t))

                                                        if len(trim(request("FM_mtype_timer_"& tpy_id &""))) <> 0 then
                                                        tpy_timer = request("FM_mtype_timer_"& tpy_id &"")
                                                        else
                                                        tpy_timer = 0
                                                        end if
                                                        
                                                        tpy_timer = replace(tpy_timer, ".", "")
                                                        tpy_timer = replace(tpy_timer, ",", ".") 

                                                        call erDetInt(tpy_timer)
                                                        if isint = 0 then
                                                        tpy_timer = tpy_timer
                                                        else
                                                        tpy_timer = 0
                                                        end if

                                                        '*** Nulstiller ikke ISINT da alle felter skal være iorden før record indlæses **'

                                                        if len(trim(request("FM_mtype_timepris_"& tpy_id &""))) <> 0 then
                                                        tpy_timepris = request("FM_mtype_timepris_"& tpy_id &"")
                                                        else
                                                        tpy_timepris = 0
                                                        end if
                                                        
                                                        tpy_timepris = replace(tpy_timepris, ".", "")
                                                        tpy_timepris = replace(tpy_timepris, ",", ".") 


                                                        call erDetInt(tpy_timepris)
                                                        if isint = 0 then
                                                        tpy_timepris = tpy_timepris
                                                        else
                                                        tpy_timepris = 0
                                                        end if


                                                        if len(trim(request("FM_mtype_kostpris_"& tpy_id &""))) <> 0 then
                                                        tpy_kostpris = request("FM_mtype_kostpris_"& tpy_id &"")
                                                        else
                                                        tpy_kostpris = 0
                                                        end if
                                                        
                                                        tpy_kostpris = replace(tpy_kostpris, ".", "")
                                                        tpy_kostpris = replace(tpy_kostpris, ",", ".") 


                                                        call erDetInt(tpy_timepris)
                                                        if isint = 0 then
                                                        tpy_timepris = tpy_timepris
                                                        else
                                                        tpy_timepris = 0
                                                        end if


                                                        if len(trim(request("FM_mtype_faktor_"& tpy_id &""))) <> 0 then
                                                        tpy_faktor = request("FM_mtype_faktor_"& tpy_id &"")
                                                        else
                                                        tpy_faktor = 0
                                                        end if
                                                        
                                                        tpy_faktor = replace(tpy_faktor, ".", "")
                                                        tpy_faktor = replace(tpy_faktor, ",", ".") 


                                                         call erDetInt(tpy_faktor)
                                                        if isint = 0 then
                                                        tpy_faktor = tpy_faktor
                                                        else
                                                        tpy_faktor = 0
                                                        end if


                                                        if len(trim(request("FM_mtype_belob_"& tpy_id &""))) <> 0 then
                                                        tpy_belob = request("FM_mtype_belob_"& tpy_id &"")
                                                        else
                                                        tpy_belob = 0
                                                        end if
                                                        
                                                        tpy_belob = replace(tpy_belob, ".", "")
                                                        tpy_belob = replace(tpy_belob, ",", ".") 

                                                        call erDetInt(tpy_belob)
                                                        if isint = 0 then
                                                        tpy_belob = tpy_belob
                                                        else
                                                        tpy_belob = 0
                                                        end if

                                                        if len(trim(request("FM_mtype_belob_ff_"& tpy_id &""))) <> 0 then
                                                        tpy_belob_ff = request("FM_mtype_belob_ff_"& tpy_id &"")
                                                        else
                                                        tpy_belob_ff = 0
                                                        end if
                                                        
                                                        tpy_belob_ff = replace(tpy_belob_ff, ".", "")
                                                        tpy_belob_ff = replace(tpy_belob_ff, ",", ".") 

                                                        call erDetInt(tpy_belob_ff)
                                                        if isint = 0 then
                                                        tpy_belob_ff = tpy_belob_ff
                                                        else
                                                        tpy_belob_ff = 0
                                                        end if

                                                        if isInt = 0 then

                                                        strSQLmtyp_tp = "INSERT INTO medarbejdertyper_timebudget "_
                                                        &" (jobid, typeid, timer, timepris, faktor, belob, belob_ff, kostpris) "_
                                                        &" VALUES "_
                                                        &"("& varJobId &", "& tpy_id &", "& tpy_timer &", "& tpy_timepris &", "& tpy_faktor &", "& tpy_belob &", "& tpy_belob_ff &", "& tpy_kostpris &")"


                                                        'Response.write strSQLmtyp_tp & "<br>"
                                                        'Response.flush

                                                        oConn.execute(strSQLmtyp_tp)

                                                        end if

                                               next


                                              
							
							
							end if 
                            'Response.end
							'**** Opret / Rediger Job ***'
				
                '******************************************************************************
                '****** WIP historik **********************************************************

                 ddDato = year(now) &"/"& month(now) &"/"& day(now) 
             
                 strSQLUpdjWiphist = "INSERT INTO wip_historik (dato, restestimat, stade_tim_proc, medid, jobid) VALUES ('"& ddDato &"', "& restestimat &", "& stade_tim_proc &", "& session("mid") &", "& id &")"
                 oConn.Execute(strSQLUpdjWiphist)

                        		



                '*************************************************************************'
                '***** Adviser jobansvarlige ****************************
                '*************************************************************************'
	            if len(trim(request.form("FM_adviser_jobans"))) <> 0 then
            	
		                    
                            jobans1 = 0
                            jobans2 = 0
                            jobans3 = 0    
                            jobans4 = 0
                            jobans5 = 0

				            '**** Finder jobansvarlige *****
				            strSQL = "SELECT job.id AS jid, jobnavn, jobnr, jobans1, jobans2, jobans3, jobans4, jobans5, job.beskrivelse, job_internbesk, "_
                            &" m1.mnavn AS m1mnavn, m1.email AS m1email, m1.mansat AS m1mansat, "_
				            &" m2.mnavn AS m2mnavn, m2.email AS m2email, m2.mansat AS m2mansat, "_
                            &" m3.mnavn AS m3mnavn, m3.email AS m3email, m3.mansat AS m3mansat, "_
                            &" m2.mnavn AS m4mnavn, m4.email AS m4email, m4.mansat AS m4mansat, "_    
                            &" m2.mnavn AS m5mnavn, m5.email AS m5email, m5.mansat AS m5mansat, "_    
                            &" kkundenavn, kkundenr, m2.init AS m2init, m1.init AS m1init, m3.init AS m3init, m4.init AS m4init, m5.init AS m5init FROM job "_
				            &" LEFT JOIN medarbejdere AS m1 ON (m1.mid = jobans1)"_
				            &" LEFT JOIN medarbejdere AS m2 ON (m2.mid = jobans2)"_
                            &" LEFT JOIN medarbejdere AS m3 ON (m3.mid = jobans3)"_
                            &" LEFT JOIN medarbejdere AS m4 ON (m4.mid = jobans4)"_
                            &" LEFT JOIN medarbejdere AS m5 ON (m5.mid = jobans5)"_
                            &" LEFT JOIN kunder AS k ON (k.kid = job.jobknr)"_
				            &" WHERE job.id = "& varJobId

                            'Response.Write strSQL
                            'Response.end

				            oRec5.open strSQL, oConn, 3
				            x = 0
				            if not oRec5.EOF then
            				
				            jobid = oRec5("jid")
				            
                            jobans1 = oRec5("m1mnavn")
                            jobans1Init = oRec5("m1init")
                            jobans1email = oRec5("m1email")
                            jobans1Mansat = oRec5("m1mansat")
                           
				            jobans2 = oRec5("m2mnavn")
				            jobans2Init = oRec5("m2init")
                            jobans2email = oRec5("m2email")
                            jobans2Mansat = oRec5("m2mansat")
				            
                             jobans3 = oRec5("m3mnavn")
				            jobans3Init = oRec5("m3init")
                            jobans3email = oRec5("m3email")
                            jobans3Mansat = oRec5("m3mansat")

                             jobans4 = oRec5("m4mnavn")
				            jobans4Init = oRec5("m4init")
                            jobans4email = oRec5("m4email")
                            jobans4Mansat = oRec5("m4mansat")

                             jobans5 = oRec5("m5mnavn")
				            jobans5Init = oRec5("m5init")
                            jobans5email = oRec5("m5email")
                            jobans5Mansat = oRec5("m5mansat")

                            jobnavnThis = oRec5("jobnavn")
                            intJobnr = oRec5("jobnr")
                            strkkundenavn = oRec5("kkundenavn")

                            strBesk = oRec5("beskrivelse")
                            job_internbesk = oRec5("job_internbesk")

                            'kkundenr = oRec("kkundenr")
            				
				            
				            end if
				            oRec5.close
            				
            				
				            '*** Henter afsender **
				            strSQL = "SELECT mnavn, email FROM medarbejdere"_
				            &" WHERE mid = "& session("mid")
				            oRec5.open strSQL, oConn, 3
            				
				            if not oRec5.EOF then
            				
				            afsNavn = oRec5("mnavn")
				            afsEmail = oRec5("email")
            				
				            end if
				            oRec5.close
            				
            				
            					
            						
            					
				            if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\jobs.asp" then
					  	            'Response.Write "isNULL(jobans1) & isNULL(jobans2)" &  isNULL(jobans1) &" "& isNULL(jobans2)
                                    'Response.end


                                    for m = 1 to 5

                                    select case m
                                    case 1
                                    jobAnsThis = jobans1
                                    jobAnsThisEmail = jobans1email
                                    jobAnsAnsat = jobans1Mansat
                                    case 2
                                    jobAnsThis = jobans2
                                    jobAnsThisEmail = jobans2email
                                    jobAnsAnsat = jobans2Mansat
                                    case 3
                                    jobAnsThis = jobans3
                                    jobAnsThisEmail = jobans3email
                                    jobAnsAnsat = jobans3Mansat
                                    case 4
                                    jobAnsThis = jobans4
                                    jobAnsThisEmail = jobans4email
                                    jobAnsAnsat = jobans4Mansat
                                    case 5  
                                    jobAnsThis = jobans5
                                    jobAnsThisEmail = jobans5email
                                    jobAnsAnsat = jobans5Mansat
                                    end select


                                    if (jobAnsThis <> "0" AND isNULL(jobAnsThis) <> true) then
                                    
                                    'Sender notifikations mail
		                            'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                            ' Sætter Charsettet til ISO-8859-1
		                            'Mailer.CharSet = 2
                                    'Mailer.FromName = "TimeOut Email Service | " & afsNavn &" "& afsEmail

                                    Set myMail=CreateObject("CDO.Message")
                                    myMail.From="timeout_no_reply@outzource.dk" 'TimeOut Email Service 

                                

		                            
                                
                                    '** Hvis problemer med mail pga spam filtre, kan der ændrs her så der bliver sendet fra en TO eller anden adresse der er godkendt af spamfilter
		                            'select case lto
                                    'case "hestia"
                                    'Mailer.FromAddress = "timeout_no_reply@outzource.dk"
                                    'case else
                                    'Mailer.FromAddress = afsEmail
		                            'end select            


                                    'Mailer.RemoteHost = "webout.smtp.nu" '"195.242.131.254" ' = smtp3.atznet.dk '"webout.smtp.nu" ' '"webmail.abusiness.dk" '"pasmtp.tele.dk"
                                    'Mailer.ContentType = "text/html"


                                     
						            'Mailer.AddRecipient "" & jobAnsThis & "", "" & jobAnsThisEmail & ""
                                    myMail.To= ""& jobAnsThis &"<"& jobAnsThisEmail &">"
                                   

						            'Mailer.Subject = "Til de jobansvarlige på: "& jobnavnThis &" ("& intJobnr &") | " & strkkundenavn  
                                    myMail.Subject= "Til de jobansvarlige på: "& jobnavnThis &" ("& intJobnr &") | " & strkkundenavn  


		                            strBody = "<br>"
                                    strBody = strBody &"<b>Kunde:</b> "& strkkundenavn & "<br>" 
						            strBody = strBody &"<b>Job:</b> "& jobnavnThis &" ("& intJobnr &") <br><br>"

                                    if jobans1 <> "0" AND isNULL(jobans1) <> true then
                                    strBody = strBody &"<b>Jobansvarlig:</b> "& jobans1 &" "& jobans1init &"<br><br>"
                                    end if

                                    if jobans2 <> "0" AND isNULL(jobans2) <> true then
                                    strBody = strBody &"<b>Jobejer:</b> "& jobans2 &" ("& jobans2init &") <br><br>"
		                            end if

                                    if len(trim(strBesk)) <> 0 then
                                    strBody = strBody &"<hr><b>Jobbeskrivelse:</b><br>"
                                    strBody = strBody & strBesk &"<br><br><br><br>"
                                    end if

                                    if len(trim(job_internbesk)) <> 0 then
                                    strBody = strBody &"<hr><b>Intern note:</b><br>"
                                    strBody = strBody & job_internbesk &"<br><br>"
                                    end if

                                    
                                    'strBody = strBody &"<br><br><br>https://outzource.dk/timeout_xp/wwwroot/ver2_10/login.asp?lto="&lto&"&tomobjid="&jobid&"&key="&strLicenskey

                                   'strBody = strBody &"<br><br><br>Gå til TimeOut ved at <a href='https://outzource.dk/timeout_xp/wwwroot/ver2_10/login.asp?lto="&lto&"&tomobjid="&jobid&"'>klikke her..</a>"
                                   '&key="&strLicenskey&"
                                    if jobAnsAnsat = "1" then
                                           if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                           strBody = strBody &"<br><br><br>Gå til TimeOut her:<br>https://outzource.dk/"&lto&"/default.asp?tomobjid="&jobid
                                           else
                                           strBody = strBody &"<br><br><br>Gå til TimeOut her:<br>https://timeout.cloud/"&lto&"/default.asp?tomobjid="&jobid
                                           end if
                                    end if
                                    
                                  


                                    strBody = strBody &"<br><br><br><br><br><br>Med venlig hilsen<br><i>" 
		                            strBody = strBody & session("user") & "</i><br><br>&nbsp;"



                                    select case lto 
                                    case "hestia"
            		                strBody = strBody &"<br><br><br><br>_______________________________________________________________________________________________<br>"
                                    strBody = strBody &"HESTIA Ejendomme, Rosenørnsgade 6, st., 8900 Randers C, Tlf. 70269010 - www.hestia.as<br><br>&nbsp;"
                                    end select
            		


            		
		                            'Mailer.BodyText = strBody
            		
		                            'Mailer.sendmail()
		                            'Set Mailer = Nothing


                                    'Mailer.BodyText = strBody
                                    myMail.HTMLBody= "<html><head></head><body>" & strBody & "</body>"

                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                                    'Name or IP of remote SMTP server
                                    
                                                if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                                   smtpServer = "webout.smtp.nu"
                                                else
                                                   smtpServer = "formrelay.rackhosting.com" 
                                                end if
                    
                                                myMail.Configuration.Fields.Item _
                                                ("http://schemas.microsoft.com/cdo/configuration/smtpserver")= smtpServer

                                    'Server port
                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
                                    myMail.Configuration.Fields.Update

                                    if len(trim(jobAnsThisEmail)) <> 0 then
                                    myMail.Send
                                    end if

                                    set myMail=nothing

            		
		                            'Mailer.sendmail()
		                            'Set Mailer = Nothing



                                    'Response.end

                                    end if
                              
                                    next
				            
                            end if 'c drev





                    end if' adviser
                            '*******************************************************


                       



                            '********************************************************
                            '*** Opret folder ****'
                            '********************************************************
                                if request("FM_opret_folder") <> "" then

                                                    
                                            '**** Opretter ikke fysisk en folder endnu ***'
                                        
                                             

                                           
                                            select case lto 
                                            case "xdencker", "xintranet - local"
                                            'Set objFSO = server.createobject("Scripting.FileSystemObject")
                                            'folderPath = "\\195.189.130.210\" 'timeout_xp\wwwroot\"
                                            'Set objnewFolder = objFSO.CreateFolder(folderPath & "DER1")
                                            'Set objnewFolder = nothing






                                            'ServerShare = "\\outzource.dk\dencker_job"
                                            'ServerShare = "\\remote.dencker.net\JOB_Timeout"
                                            ServerShare = "\\remote.dencker.net\to_job4"
                                            
                                            
                                            
                                            'UserName = "Administrator"
                                            UserName = "ad"
                                            'UserName = "TimeOut"
                                            'Password = "Sok!2637"
                                            Password = "ad1996"
                                            'Password = "betina00"
                                            'Password = "timeout"
                                            'Password = ""

                                            Set NetworkObject = CreateObject("WScript.Network")
                                            Set FSO = CreateObject("Scripting.FileSystemObject")

                                            NetworkObject.MapNetworkDrive "", ServerShare, False ', UserName, Password

                                            Set objnewFolder = FSO.CreateFolder(ServerShare &"/"& strNavn &"_"& strjnr)

                                            Set Directory = FSO.GetFolder(ServerShare)
                                            For Each FileName In Directory.Files
                                                response.write FileName.Name & "<br>"
                                                response.write fso.GetParentFolderName(FileName.Name) & "<br>"
                                            Next

                                            Set FileName = Nothing
                                            Set Directory = Nothing
                                            Set FSO = Nothing
                                            Set objnewFolder = nothing

                                            response.write "Folder opRettet!"
                                            response.end

                                            end select
                         
            	                            oprFysisk = 1
                                            if oprFysisk = 1000 then

	                                                if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\jobs.asp" then
            							
		                                                Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
		                                                Set objNewFile = nothing
		                                                Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
            	
	                                                else
            		
		                                                Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, false)
		                                                Set objNewFile = nothing
		                                                Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8, -1)
            		
	                                                end if
            	
	                                        file = "ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
            	
            	
	                                        objF.WriteLine(ekspTxt)


                                            objF.close

                                            end if

                                        '*** Insert into db folder ***'

                                        sqldatoDt = year(now) &"/"& month(now) &"/"& day(now)
                                        strSQL = "INSERT INTO foldere (navn, kundeid, kundese, jobid, editor, dato) VALUES "_
		                                &" ('"& strNavn &"_"& strjnr &"', "& oRec("kid") &", 0, "& varJobid &", '"& session("user") &"', '"& sqldatoDt &"')"
		                                
		
		
		
		                                oConn.execute(strSQL)
		


                                 end if

							
							
							'**** Tilknytter / Sletter / Opdaterer Uleverandører ***'
							
							'strSQLDelUlev = "DELETE FROM job_ulev_ju WHERE ju_jobid = "& varJobId
							'oConn.execute(strSQLDelUlev)
							
							For u = 1 to 50
							    
                                if len(trim(request("ulevid_"&u&""))) <> 0 then
                                ulevid = request("ulevid_"&u&"")
                                else
                                ulevid = 0
                                end if

                                ulevfase = "" 'replace(request("ulevfase_"&u&""), "'", "''")
							    ulevnavn = replace(request("ulevnavn_"&u&""), "'", "''")
							    
                                ulevstk = replace(request("ulevstk_"&u&""), ".", "")
							    ulevstk = replace(ulevstk, ",", ".")

                                ulevstkpris = replace(request("ulevstkpris_"&u&""), ".", "")
							    ulevstkpris = replace(ulevstkpris, ",", ".")

							    ulevpris = replace(request("ulevpris_"&u&""), ".", "")
							    ulevpris = replace(ulevpris, ",", ".")
							    
							    ulevfaktor = replace(request("ulevfaktor_"&u&""), ".", "")
							    ulevfaktor = replace(ulevfaktor, ",", ".")
							    
							    ulevbelob = replace(request("ulevbelob_"&u&""), ".", "")
							    ulevbelob = replace(ulevbelob, ",", ".")
    							
    							
    							
		                     if len(trim(ulevnavn)) <> 0 then


                                        if len(trim(ulevstk)) <> 0 then
    							        ulevstk = ulevstk
    							    
    							            call erDetInt(ulevstk)
    							            if isInt = 0 then
    							            ulevstk = ulevstk
    							            else
    							            ulevstk = 0
    							            end if
    							            isInt = 0
    							    
    							        else
    							        ulevstk = 0
    							        end if
                                        
                                        if len(trim(ulevstkpris)) <> 0 then
    							        ulevstkpris = ulevstkpris
    							    
    							            call erDetInt(ulevstkpris)
    							            if isInt = 0 then
    							            ulevstkpris = ulevstkpris
    							            else
    							            ulevstkpris = 0
    							            end if
    							            isInt = 0
    							    
    							        else
    							        ulevstkpris = 0
    							        end if
                                            
					            
					                    if len(trim(ulevpris)) <> 0 then
    							        ulevpris = ulevpris
    							    
    							            call erDetInt(ulevpris)
    							            if isInt = 0 then
    							            ulevpris = ulevpris
    							            else
    							            ulevpris = 0
    							            end if
    							            isInt = 0
    							    
    							        else
    							        ulevpris = 0
    							        end if
    							
    							
    							        if len(trim(ulevfaktor)) <> 0 then
    							        ulevfaktor = ulevfaktor
    							    
    							            call erDetInt(ulevfaktor)
    							            if isInt = 0 then
    							            ulevfaktor = ulevfaktor
    							            else
    							            ulevfaktor = 0
    							            end if
    							            isInt = 0
    							    
    							        else
    							        ulevfaktor = 0
    							        end if
    							
    							        if len(trim(ulevbelob)) <> 0 then
    							        ulevbelob = ulevbelob
    							        
    							            call erDetInt(ulevbelob)
    							            if isInt = 0 then
    							            ulevbelob = ulevbelob
    							            else
    							            ulevbelob = 0
    							            end if
    							            isInt = 0
    							        
    							        else
    							        ulevbelob = 0
    							        end if
    							

                                            if cdbl(ulevid) <> 0 then

                                            strSQLUpdUlev = "UPDATE job_ulev_ju SET "_
							                &" ju_fase = '"& ulevfase &"', ju_navn = '"& ulevnavn &"', ju_ipris = "& ulevpris &", "_
							                &" ju_faktor = "& ulevfaktor &", ju_belob = "& ulevbelob &",  ju_jobid = "& varJobId & ", ju_stk = "& ulevstk &", ju_stkpris = "& ulevstkpris &" WHERE ju_id = " & ulevid
							                
                                            'Response.Write strSQLInsUlev & "<br>"
                                            oConn.execute(strSQLUpdUlev)


                                            else
    							
						                    strSQLInsUlev = "INSERT INTO job_ulev_ju SET "_
							                &" ju_fase = '"& ulevfase &"', ju_navn = '"& ulevnavn &"', ju_ipris = "& ulevpris &", "_
							                &" ju_faktor = "& ulevfaktor &", ju_belob = "& ulevbelob &",  ju_jobid = "& varJobId& ", ju_stk = "& ulevstk &", ju_stkpris = "& ulevstkpris & ""_
							                
                                            'Response.Write "ulevid:" & ulevid & " -- "&  strSQLInsUlev & "<br>"
                                            oConn.execute(strSQLInsUlev)

                                            end if
    							
							   
                               
                               else
                                        
                                         if cint(ulevid) <> 0 then

                                         ulevDel = "DELETE FROM job_ulev_ju WHERE ju_id = " & ulevid
                                         oConn.execute(ulevDel)

                                         end if
                               
                               
                               end if
    							
							  
							
							next
							
						    'Response.end
							
							
						    '*****************************************************************************'
							'*** tilføjer job i timereg_usejob (Vis guide), både ved opret og rediger ****'
							'*****************************************************************************'
						
                        'repsonse.write "her"
                        'response.end

                         
							medarUseJobWrt = ""
							dtNow = year(now) & "-"& month(now) & "-"& day(now)
							
								strSQL = "SELECT DISTINCT(MedarbejderId) FROM progrupperelationer WHERE ("_
								&" ProjektgruppeId = "& strProjektgr1 &""_
								&" OR ProjektgruppeId =" & strProjektgr2 &""_
								&" OR ProjektgruppeId =" & strProjektgr3 &""_
								&" OR ProjektgruppeId =" & strProjektgr4 &""_
								&" OR ProjektgruppeId =" & strProjektgr5 &""_
								&" OR ProjektgruppeId =" & strProjektgr6 &""_
								&" OR ProjektgruppeId =" & strProjektgr7 &""_
								&" OR ProjektgruppeId =" & strProjektgr8 &""_
								&" OR ProjektgruppeId =" & strProjektgr9 &""_
								&" OR ProjektgruppeId =" & strProjektgr10 &""_
								&") GROUP BY MedarbejderId"
								
                                'Response.Write "strSQL "& strSQL & "<br><hr>"

								
								oRec5.open strSQL, oConn, 3
								while not oRec5.EOF

                                    
                                    medarbfundet = 0
                                    strSQLex = "SELECT id FROM timereg_usejob WHERE jobid = "& varJobId &" AND medarb = " & oRec5("MedarbejderId")
                                    
                                    'Response.Write strSQLex &"<br><br>"
                                    
                                    oRec4.open strSQLex, oConn, 3
                                    if not oRec4.EOF then
                                    medarbfundet = 1
                                            
                                            if cint(forvalgt) = 1 then 'sæt aktiv for alle tilmeldte medarbejdere i valgte projektgrupper

                                            strSQLfvlgt = "UPDATE timereg_usejob SET forvalgt = 1, forvalgt_sortorder = 0, forvalgt_af = "& session("mid") &", forvalgt_dt = '"& dtNow &"' WHERE id = " & oRec4("id") 

                                            oConn.execute(strSQLfvlgt)
                                            'Response.Write strSQLfvlgt & "<br>"

                                            end if

                                    End if
                                    oRec4.close

                                    if cint(medarbfundet) = 0 then 
                                        
                                        if cint(forvalgt) = 1 then 'sæt aktiv
                                        strSQL3 = "INSERT INTO timereg_usejob (medarb, jobid, forvalgt, forvalgt_sortorder, forvalgt_af, forvalgt_dt) VALUES "_
                                        &" ("& oRec5("MedarbejderId") &", "& varJobId &", 1, 0, "& session("mid") &", '"& dtNow &"')"

                                         oConn.execute(strSQL3)
                                        else
                                        
                                                
                                                call positiv_aktivering_akt_fn() 'wwf
                                                if cint(positiv_aktivering_akt_val) <> 1 then
                                                strSQL3 = "INSERT INTO timereg_usejob (medarb, jobid) VALUES ("& oRec5("MedarbejderId") &", "& varJobId &")"
                                                oConn.execute(strSQL3)
                                                end if
                                        end if

                                   
                                   
									'Response.Write strSQL3 & "<br>"
									'Response.Flush
                                    end if
									
                                    
                                    'Response.End 
									
								     medarUseJobWrt = medarUseJobWrt & " AND medarb <> "& oRec5("MedarbejderId")
									
									
								oRec5.movenext
								wend
								oRec5.close

                                'Response.End

							
							if func = "dbred" then
							'*** Sletter ikke længere aktuelle medarbejdere fra timereg usejob **'
                                'if len(trim(medarUseJobWrt)) <> 0 then
							    strSQLdelUseJob = "DELETE FROM timereg_usejob WHERE jobid = "& varJobId & "" & medarUseJobWrt

                                'Response.Write "<br><br>"& strSQLdelUseJob

                                oConn.execute(strSQLdelUseJob)
							    'end if
                            end if



                            '**** Sætter EASYreg aktiv på aktiviteter for alle medarbejdere ****'
                            
                            oEasyReg = 0
                            strSQLea = "SELECT easyreg, id FROM aktiviteter WHERE job = "& id & " AND easyreg = 1"
				            oRec5.open strSQLea, oConn, 3
				            while not oRec5.EOF 
				                
                              oEasyReg = 1
                            
                            oRec5.movenext
				            wend 
				            oRec5.close
				                

                                if cint(oEasyReg) = 1 then
				                strSQLtreguse = "UPDATE timereg_usejob SET easyreg = " & id & " WHERE jobid = " & id
	   	                        'Response.Write strSQLtreguse & "<br>"
	   	                        'Response.flush
	   	                        oConn.execute(strSQLtreguse)
				                end if
				               
				            'Response.end
				         
					        


					        '************************************************'
							' Opdaterer prjgp (fjerner) på aktiviteter      *'
							' så der ikke kan være progp på aktiviteter     *'
							' der ikke også findes på jobbet.               *'
							'************************************************'
							if func = "xxx" then	
                            '** Slået fra da det giver problemer iforhold til nedarv. BÅDE ved opret og rediger job, f.eks ved FØD job med nye stamaktivitetyer ved rediger job	
							'** Funktionen kan dog bruges til oprydning ***'		
							
							
								    strSQL = "SELECT id, job, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter WHERE job = "& varJobId
								    'Response.write strSQL
								    oRec5.open strSQL, oConn, 3
								    
									while not oRec5.EOF 
									
									call pjgcase(oRec5("projektgruppe1"), 1)
									call pjgcase(oRec5("projektgruppe2"), 2)
									call pjgcase(oRec5("projektgruppe3"), 3)
									call pjgcase(oRec5("projektgruppe4"), 4)
									call pjgcase(oRec5("projektgruppe5"), 5)
									call pjgcase(oRec5("projektgruppe6"), 6)
									call pjgcase(oRec5("projektgruppe7"), 7)
									call pjgcase(oRec5("projektgruppe8"), 8)
									call pjgcase(oRec5("projektgruppe9"), 9)
									call pjgcase(oRec5("projektgruppe10"), 10)
									
									oRec5.movenext
									wend 
								    oRec5.close
							
					
					        end if
					
					

                    '******************************************************************************************'
                    '**** Opdaterer medarbejder timeprsier til at følge budgetprisen på aktivitets linjerne ***'
                    '******************************************************************************************'
                    if cint(laasmedtpbudget) = 1 then

                   
                        call opd_aktfasttp(id,opdekstp,valuta)




                    end if



                    if func = "dbred" then
                    '***************************************
                    '**** BUDGET FY ************************
                    '***************************************
                    f = 0
                    FY0 = request("FM_fy_aar_0")
                    FY1 = request("FM_fy_aar_1")
                    FY2 = request("FM_fy_aar_2")
                    jobBudgetFY0 = request("FM_fy_hours_0")
                    fctimeprisFY0 = 0
                    fctimeprish2FY0 = 0
                    jobBudgetFY1 = request("FM_fy_hours_1")
                    jobBudgetFY2 = request("FM_fy_hours_2")
                    jobids = id
                    aktid = 0

                    for f = 0 to 2
                        call opdaterRessouceRamme(f, FY0, FY1, FY2, jobBudgetFY0, fctimeprisFY0, fctimeprish2FY0, jobBudgetFY1, jobBudgetFY2, jobids, aktid)
                    next

                    end if
					

                    '********************************'
                    '***** Forretningsområder ******'
                    '********************************'
                    
                    '*** nulstiller job ****'
                    strSQLfor = "DELETE FROM fomr_rel WHERE for_jobid = "& varJobId & " AND for_aktid = 0"
                    oConn.execute(strSQLfor)


                    '*** IKKE MERE 11.3.2015 (aktiviteter tælles altid med på job) ***'
                    '*** nulstiller akt ved sync ****'
                    'if cint(syncForAkt) = 1 then
                    '    strSQLfor = "DELETE FROM fomr_rel WHERE for_jobid = "& varJobId & " AND for_aktid <> 0"
                    '    oConn.execute(strSQLfor)
                    'end if

                    'Response.Write "her"
                    'Response.Flush

                    for afor = 0 to UBOUND(fomrArr)

                            'Response.Write "her2" & afor & "<br>"
                            'Response.Flush

                            if fomrArr(afor) <> 0 then

                            strSQLfomri = "INSERT INTO fomr_rel "_
                            &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                            &" VALUES ("& fomrArr(afor) &", "& varJobId &", 0, "& for_faktor &")"

                            oConn.execute(strSQLfomri)

                                '*** Sync aktiviteter ****'
                                'if cint(syncForAkt) = 1 then
                                'strSQLa = "SELECT id FROM aktiviteter WHERE job = "& varJobid
                                'oRec3.open strSQLa, oConn, 3
                                'while not oRec3.EOF 

                                '    strSQLfomrai = "INSERT INTO fomr_rel "_
                                '    &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                                '    &" VALUES ("& fomrArr(afor) &", "& varJobId &","& oRec3("id") &", "& for_faktor &")"

                                '    oConn.execute(strSQLfomrai)

                                'oRec3.movenext
                                'wend
                                'oRec3.close

                                'end if

                            end if


                    next

                    '********************************'

					
                    end if 'jobnummer findes		
							
					oRec.movenext
					wend
					oRec.close
                    '*** End for next Multible opret på kunder ***'
					
                    'Response.Write "func:" & func
                    'Response.end	
							
							 
							 
							 if len(request("fm_kunde_sog")) <> 0 then
									vmenukundefilt = request("fm_kunde_sog")
									else
									vmenukundefilt = 0
									end if
									
									if len(trim(request("jobnr_sog"))) <> 0 then
									strJobsog = request("jobnr_sog")
									else
									strJobsog = left(strNavn, 10)
									end if
							
							
							if request("FM_usetilbudsnr") = "j" then
							filt = "tilbud"
							else
							filt = request("filt") 
							end if
							

                             if len(trim(request("visrealtimerdetal"))) <> 0 AND request("visrealtimerdetal") <> 0 then
                             visrealtimerdetal = 1
                             else
                             visrealtimerdetal = 0
                             end if
							
							
                          
							
                            
							'timeB = now
                            'loadtime = datediff("s",timeA, timeB)
							'Response.write loadtime
                            'response.flush
							'Response.Write "<br>" & rdir
							'Response.end
						
                            
                        	


							select case request("rdir")
							case "redjobcontionue"
							rdirLink = "jobs.asp?menu=job&func=red&id="&varJobId&"&jobnr_sog="&strJobsog&"&filt="&filt&"&fm_kunde="&strKundeId&"&fm_kunde_sog="&vmenukundefilt&"&showdiv="&showdiv&"&FM_tp_jobaktid="&tp_jobaktid&"&FM_mtype="&mtype&"&visrealtimerdetal="&visrealtimerdetal
                            'Response.Write rdirLink
                            'Response.end
                            case "jbpla_w"
							strKlokkeslet = request("sttid")
							dato = request("dato")
							jobstKri = request("datostkri")
							jobslKri = request("datoslkri")
							rdirLink = "jbpla_w_opr.asp?func=step1&sttid="&strKlokkeslet&"&dato="&dato&"&datostkri="&jobstKri&"&datoslkri="&jobslKri&"&id=0&step=2&jobid="&varJobId&"&stepialt=3"
							case "sdsk"
							rdirLink = "<script language=""JavaScript"">window.opener.location.reload();</script>"
							rdirLink = rdirLink & "<script language=""JavaScript"">window.close();</script>"
							case "sdsk2"
							rdirLink = "sdsk.asp?func=red&id=0"
							case "webblik"
							rdirLink = "webblik_joblisten.asp"
							case "webblik_tilfakturering"
							rdirLink = "webblik_tilfakturering.asp"
							case "webblik_milepale"
							'rdirLink = "webblik_milepale.asp"
                            rdirLink = "<script language=""JavaScript"">window.opener.location.href(""webblik_milepale.asp"");</script>"
							rdirLink = rdirLink & "<script language=""JavaScript"">window.close();</script>"
							case "pg"
							rdirLink = "pg_allokering.asp?menu=job"
							case "kon"
							rdirLink = "kunder.asp?menu=kund"
							case "pipe"
							'rdirLink = "pipeline.asp?menu=job"
                            rdirLink = "<script language=""JavaScript"">window.opener.location.reload();</script>"
							rdirLink = rdirLink & "<script language=""JavaScript"">window.close();</script>"
							case "treg"	
						    rdirLink = "timereg_akt_2006.asp?showakt=1"
												
							case "seraft"
							rdirLink = "kunder.asp?menu=kund&func=red&id="&strKundeId&"&showdiv=seraft"
							
							case else
							
							'**til joblisten
							rdirLink = "jobs.asp?menu=job&shokselector=1&id="&varJobId&"&jobnr_sog="&strJobsog&"&filt="&filt&"&fm_kunde="&vmenukundefilt
							
                            
                          end select
							
						  
						    '*****************************************'
				            '******* Opret opgave Inceidnt i SDSK ****'
				            '*****************************************'
            				
				             if len(request("FM_opr_incident")) <> 0 then
            							    
				                if func = "dbred" then
				                kidforSDSK = strKundeId
				                else
				                kidforSDSK = strKidUseOne 'trim(kids(0))
				                end if
				            
				            Response.Write("<script language=""JavaScript"">window.open('sdsk.asp?func=opr1&id=0&FM_kontakt="&kidforSDSK&"&FM_emne="&strNavn&" ("& strjnr &")"&"','sdskwin');</script>")
                            %>
                            <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
                            
                            <body>
                            <%
                            itop = 100
                            ileft = 150
                            iwdt = 400
                            call sideinfo(itop,ileft,iwdt)
                            
                             %>
                            
                            
			                       <tr><td style="padding:10px;">
                                   Du har valgt at oprette en tilhørende incident (opgave), den er nu åbnet i et nyt vindue. <br /><a href="<%=rdirLink %>" class=vmenu>klik her for at fortsætte..</a>
                            
	                        </td></tr></table>
	                        </div>
	                        <br><br><br><br><br>&nbsp;
                            </body>
                            
                            
						    <%
						    Response.end
						    else
						    
						     if rdir = "sdsk" OR rdir = "pipe" or rdir = "webblik_milepale" then
						     response.Write rdirLink
						     else
						     response.Redirect rdirLink
						     end if
						     
						    end if
				
                '** Validering		  
						
					'end if
				'end if
			  end if
			 end if
			end if
		end if
	end if
		
	
	
	
	
	case "opret", "red"
	'*** Tjekker om det er fra ressourcebooking/popup at jobbet bliver oprettet ***
	showaspopup = request("showaspopup")
	headergif = "../ill/job_oprred_header.gif"
	
	
	
	
	if showaspopup <> "y" then%>
	
	
	
	
	<%
	if len(trim(request("step"))) <> 0 then
	step = request("step")
	else
		if request("func") = "red" then
		step = 2
		else
		step = 1
		end if
	end if
	
	
	if request("FM_kunde") = "0" AND step = 2 then
	%>
					 <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					 
					 
					 
			          <%
					errortype = 92
					call showError(errortype)
					isInt = 0
			        Response.End
			        
			        
	end if
	
    
    %>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    
    <%if step = 2 then %>
    <script src="inc/job_jav.js"></script>
    <%else %>
    <script src="inc/job_listen_jav.js"></script>
    <%end if 
    
    
    if nomenu <> "1" then%>
    
    

    <!--
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(3)%>
	</div>

	<div id="sekmenu" style="position:absolute; left:15px; top:82px; visibility:visible;">
	<%
	call jobtopmenu()
	%>
    </div>
    <%end if%>
        -->
	
	<%call menu_2014() %>



	<%else%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
    
	<%
	'level = 1
	'editok = 1 %>
	
	
	<%end if 'showaspopup%>
	<!--#include file="inc/job_inc.asp"-->
	<%
	
	if showaspopup <> "y" then
    call top
	varbroedkrumme = "Rediger job"
	
	else
	varbroedkrumme = "Opret job"
	%>
		<br>&nbsp;&nbsp;&nbsp;&nbsp;<font class="stor-blaa">Opret job ekspress</font>
	<%end if

	
	

	
	
	
	'*** Placering af side div ***
	if showaspopup <> "y" then
	sidedivLeft = 90

    if step = 1 then
    sidedivTop = 132
    else
	sidedivTop = 152
	end if  
      
	    if func <> "red" then
	    
	    if step = 1 then
	    mainTableWidth = 675
	    else
	    mainTableWidth = 475
	    end if
	    
	    
	    select case step
	    case 1
	    varbroedkrumme = "<h3 class=""hv"">Vælg kunde <span style='font-size:10px;'> - Opret nyt job step 1/2</span></h3>"
	    case 2  
	    varbroedkrumme = "<h3 class=""hv"">Jobstamdata <span style='font-size:10px;'> - Opret nyt job step 2/2</span></h3>"
	    end select
	    
	    else
	    
	    varbroedkrumme = "<h3 class=""hv"">Jobstamdata<span style='font-size:10px;'> - Rediger</span></h3>"
	    mainTableWidth = 475
	    
	    end if
	    
	else
	step = 2
	stdata_vzb = "visible"
	stdata_dsp = ""
	sidedivLeft = 90
	sidedivTop = 132
	mainTableWidth = 475
	'varbroedkrumme = varbroedkrumme & " (Ekspres oprettelse)"
	end if
	
     '************ Faneblade ******************
    if func = "red" then 'AND strInternt = "1" 
    'strJobnr & strKnr = request("FM_kunde")
    strSQLfb = "SELECT id, jobnavn, jobnr, jobknr FROM job, kunder WHERE id = " & id 
	
	'Response.Write strSQLfb
	'Response.flush
	
	oRec.open strSQLfb, oConn, 3
	
	if not oRec.EOF then
	fbNavn = oRec("jobnavn")
	fbjobnr = oRec("jobnr")
	'fbKnavn = oRec("kkundenavn")
	fbKnr = oRec("jobknr")
	end if
    oRec.close
    
    
    call faneblade("job", showdiv)
    end if
	

    if func = "opret" AND step = 2 then
	%>
	<!-- manual --->
    <div id="help" style="position:absolute; background-color:#ffffff; padding:1px 1px 0px 1px; width:170px; border:1px silver solid; border-bottom:0px; left:1020px; top:135px; visibility:visible; display:; z-index:9000000; overflow:hidden;">
    <table cellpadding=0 cellspacing=0 border=0 width=100%>
    <tr bgcolor="#FF6666"><td align=center>
    <%
    select case lto
    case "execon" %>
    <a href="../help_and_faq/TimeOut_job_oprettelse_rev_100924_execon.pdf" class=alt target="_blank">Manual til joboprettelse..</a>
    <%case "wwf"
    %>
    <a href="../help_and_faq/TimeOut_job_oprettelse_rev_110523_wwf.pdf" class=alt target="_blank">Manual til joboprettelse..</a>
    <%
    case else %>
    <a href="../help_and_faq/TimeOut_job_oprettelse_rev_110523.pdf" class=alt target="_blank">Manual til joboprettelse..</a>
    <%end select%>

    
    
        </td>
     </tr>
    </table>
	</div>
    <%end if %>

    <%select case lto
    case "epi", "epi_cati", "epi_no", "epi_sta", "epi_se", "intranet - local", "epi_uk"
    %>
      <!-- Popup MSG -->
	<div id="mtyp_msg1" style="position:absolute; top:200px; left:1250px; width:275px; border:1px #999999 solid; z-index:900000000; background-color:#FFFFe1; padding:10px 10px 10px 10px; visibility:hidden; display:none;">
    <b>Angiv</b>, hvor stor en del af bruttoomsætningen der skal tildeles de forskellige medarbejdertyper. <br /><br />
     Du kan enten<br /><br />
        <b>1) Tildele totalbeløb først</b> og defter angive timer på hver enkelt linje (lås beløb)<br /><br />
        <b>2) eller tildele timer på hver linje</b> og tilsidst overføre summen til Bruttoomsætningen.<br /><br />
        <span id="lukms1" style="color:#ea0957; font-size:9px;">[luk]</span>
	</div>
	
     <!-- Popup MSG -->
	<div id="mtyp_msg2" style="position:absolute; top:200px; left:1250px; width:275px; border:1px #999999 solid; z-index:900000000; background-color:#FFFFe1; padding:10px; visibility:hidden; display:none;">
    <b>Husk at</b> tildel timer på medarbejder-typerne. Faktor sættes ens på alle linier indenfor hver gruppe.<br /><br />
          <span id="lukms2" style="color:#ea0957; font-size:9px;">[luk]</span>
	</div>
	
    <%
    case else
    end select %>

  



    <!-------------------------------Sideindhold------------------------------------->

    <form action="jobs.asp?menu=job" method="post" name="jobdata" id="jobdata">
	<div id="sindhold" style="position:absolute; left:<%=sidedivLeft%>px; top:<%=sidedivTop%>px; visibility:<%=stdata_vzb%>; display:<%=stdata_dsp%>; z-index:50;">
	
	
	<%
    
   
	
	
	

    tTop = 0
	tLeft = 0
	tWdth = mainTableWidth
	
	
	call tableDiv_j(tTop,tLeft,tWdth)
	%>

    <table cellspacing="0" cellpadding="0" border="0" bgcolor="#FFFFFF" width="100%">
	
    <input id="showdiv" name="showdiv" value="<%=showdiv%>" type="hidden" />
	<!--<input type="hidden" name="func" id="func" value="<=dbfunc%>">-->
	<input type="hidden" name="fm_kunde_sog" value="<%=request("fm_kunde_sog")%>">
	<input type="hidden" id="rdir" name="rdir" value="<%=request("rdir")%>">
	<input type="hidden" id="jq_jobid" name="id" value="<%=id%>">
	<input type="hidden" name="int" id="int" value="<%=strInternt%>">
	<input type="hidden" name="jobnr_sog" value="<%=request("jobnr_sog")%>">
	<input type="hidden" name="filt" value="<%=request("filt")%>">
	<input type="hidden" name="showaspopup" value="<%=showaspopup%>">
    <input type="hidden" id="jq_func" name="jq_func" value="<%=func%>">
    <input type="hidden" id="lto" name="lto" value="<%=lto%>">
	
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan=2 valign="top"><img src="../ill/blank.gif" width="<%=(mainTableWidth-16)%>" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan=2 class="alt" valign="top" style="padding-top:4px;"><%=varbroedkrumme %></td>
	</tr>
	<%
	select case step
	case "1" 
	editok = 1
	
	%>
	<input type="hidden" name="step" value="<%=step+1%>">
	<input type="hidden" name="func" id="func" value="<%=func%>">
	
	<input type="hidden" name="FM_fastpris1" id="FM_fastpris1" value="0">
	
	
	<%if showaspopup = "y" then%>
	<input type="hidden" name="sttid" value="<%=request("sttid")%>">
	<input type="hidden" name="dato" value="<%=request("dato")%>">
	<input type="hidden" name="datostkri" value="<%=request("datostkri")%>">
	<input type="hidden" name="datoslkri" value="<%=request("datoslkri")%>">
	<%end if%>
	<%
	
	call kundeopl


	
	
	case "2" '"3"
	

    



	
	if func = "opret" then
	
	%>
	<input type="hidden" name="step" value="<%=step+1%>">
	<!-- Fra Step 1 og 2 -->
	<input type="hidden" name="FM_kunde" value="<%=request("FM_kunde")%>">
	<!--
	<input type="hidden" name="FM_kundese" value="<%=request("FM_kundese")%>">
	<input type="hidden" name="FM_kundese_hv" value="<%=request("FM_kundese_hv")%>">
	-->
	
	
	<%
	
	end if
	
	if func = "opret" AND step = 2 OR func = "red" then
	sendmailtokunde = "onClick='mailtokunde()'"
	else
	sendmailtokunde = ""
	end if
	
	thisMid = session("mid")

	
	                '*** Er der på licebns niveau mulighe for at lukke og afmelde job? ***'
	                lukafm = 0
					strSQLlicens = "SELECT lukafm FROM licens WHERE id = 1"
					oRec5.open strSQLlicens, oConn, 3
					if not oRec5.EOF then
					
					lukafm = oRec5("lukafm")
					
					end if
					oRec5.close 
	
	
   


	'*** Her indlæses form til rediger/oprettelse af job ***
	if func = "opret" then
	
	strjobnr = 0
	strtilbudsnr = 0
    strNexttilbudsnr = 0
	strNavn = "Jobnavn.."
	
	dbfunc = "dbopr"
	strFakturerbart = 1
	rowspan = "28"

	    select case lto
	    case "intranet - local", "synergi1", "cisu", "wilke"
	    strFastpris = 1 'default fastpris
	    case else
	    strFastpris = 2 'default løbende timer
	    end select
	
    intkundeok = 0
	varSubVal = "opretpil"
	
	if len(request.cookies("job")("defaultprojgrp")) <> 0 AND lto <> "execon" then 
	strProj_1 = request.cookies("job")("defaultprojgrp")
	else
	    select case lto 
        case "execon" 
        strProj_1 = 1
        case else
        strProj_1 = 10
        end select
	end if
	
	strProj_2 = 1
	strProj_3 = 1
	strProj_4 = 1
	strProj_5 = 1
	strProj_6 = 1
	strProj_7 = 1
	strProj_8 = 1
	strProj_9 = 1
	strProj_10 = 1
	
	intkundekpers = 0
	
	editok = 1
	
	jobans1 = 0
	jobans2 = 0
	
	ikkeBudgettimer = 0
	strBudgettimer = 0
	
	rekvnr = ""
	
	ski = 0
	abo = 0
	ubv = 0
	
	'*** Mulitible opret *** (bruges til kontakt personer) **'
	

	                
				    strKundeId = 0
				    kids = split(request("FM_kunde"), ",")
				    for x = 0 to UBOUND(kids)
				        
				        if x > 1 then
				        strKundeId = 0
				        else
				        strKundeId = kids(x)
				        end if
				        
				        if x > 1 then
				        multibletrue = 1
				        else
				        multibletrue = 0
				        end if
				    
				    next
				    
				   
	
	
	
	laasmedtpbudget = 0
	intServiceaft = 0
	
	if cint(lukafm) <> 0 then
	lukafmjob = 1
	else 
	lukafmjob = 0
	end if
	
	

	jobfaktype = 0
	jo_gnstpris = 0

    select case lto 
    case "epi", "epi_no", "epi_sta", "outz", "epi_ab", "epi_cati", "epi_uk"
    jo_gnsfaktor = 2
	case else
    jo_gnsfaktor = 1
    end select

    jo_gnsbelob = 0
    jo_bruttofortj = 0
    jo_dbproc = 0
    
    udgifter = 0
    usejoborakt_tp = 0
    

    if len(trim(strKundeId)) <> 0 then
    strKundeId = strKundeId
    else
    strKundeId = 0
    end if

    select case lto 
    case "fe"
    intprio = 1
    case "dencker"
        if cint(strKundeId) = 1 then
        intprio = -3
        else
        intprio = 100
        end if        
    case else
    intprio = 100
    end select
    
  
    
    sandsynlighed = 0
    jfak_moms = 1
    jfak_sprog = 1
    sdskpriogrp = 0
    valuta = basisValId
   
    if multibletrue = 0 then
    
    '** er der en SDSK priogruppe, valuta,moms og fak sprog på kunde ****
    strSQL = "SELECT sdskpriogrp, kfak_moms, kfak_sprog, kfak_valuta FROM kunder WHERE kid = "& strKundeId
    oRec.open strSQL, oConn, 3
    if not oRec.EOF then
    sdskpriogrp = oRec("sdskpriogrp")
    jfak_moms = oRec("kfak_moms")
    jfak_sprog = oRec("kfak_sprog")
    valuta = oRec("kfak_valuta")
    end if
    oRec.close
    
    end if 

   



    restestimat = 0
    stade_tim_proc = 0

    select case lto
    case "epi", "epi_no", "epi_sta", "intranet - local", "epi_ab", "epi_cati", "epi_uk"
	virksomheds_proc = 50
	syncslutdato = 1
    case else
    virksomheds_proc = 0
	syncslutdato = 0
    end select
    
    altfakadrCHK = ""
    altfakadr = 0

    preconditions_met = 0 '1 = ja, 0= ikke angivet, 2 = nej vent

    filepath1 = ""
    fomr_konto = 0

    useasfak = 0

	else '*** REDIGER JOB *****'
    
    vlgtmtypgrp = 0
    call mtyperIGrp_fn(vlgtmtypgrp,0) 
    call fn_medarbtyper()
    

	strSQL = "SELECT id, jobnavn, jobnr, kkundenavn, jobknr, "_
	&" jobTpris, jobstatus, jobstartdato, jobslutdato, projektgruppe1, projektgruppe2, "_
	&" projektgruppe3, projektgruppe4, projektgruppe5, job.dato, job.editor, "_
	&" fakturerbart, budgettimer, fastpris, kundeok, job.beskrivelse, "_
	&" ikkeBudgettimer, tilbudsnr, projektgruppe6, projektgruppe7, "_
	&" projektgruppe8, projektgruppe9, projektgruppe10, job.serviceaft, "_
	&" kundekpers, jobans1, jobans2, lukafmjob, valuta, jobfaktype, rekvnr, "_
	&" jo_gnstpris, jo_gnsfaktor, jo_gnsbelob, jo_bruttofortj, jo_dbproc, "_
	&" udgifter, risiko, sdskpriogrp, usejoborakt_tp, ski, job_internbesk, abo, ubv, sandsynlighed, "_
    &" diff_timer, diff_sum, jo_udgifter_ulev, jo_udgifter_intern, jo_bruttooms, restestimat, stade_tim_proc, virksomheds_proc, "_
    &" syncslutdato, lukkedato, altfakadr, preconditions_met, laasmedtpbudget, filepath1, fomr_konto, jfak_moms, jfak_sprog, useasfak "_
    &" FROM job, kunder WHERE id = " & id &" AND kunder.Kid = jobknr"
	
	'Response.Write strSQL
	'Response.flush
	
	oRec.open strSQL, oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("jobnavn")
	strjobnr = oRec("jobnr")
	strKnavn = oRec("kkundenavn")
	strKnr = oRec("jobknr")
	strBudget = oRec("jo_bruttooms") 'oRec("jobTpris")
	strStatus = oRec("jobstatus")
	strTdato = oRec("jobstartdato")
	strUdato = oRec("jobslutdato")
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
	strFakturerbart = oRec("fakturerbart")
	
	
	strFastpris = oRec("fastpris")
	intkundeok = oRec("kundeok")
	strBesk = oRec("beskrivelse")


    useasfak = oRec("useasfak")
    
        if cdbl(oRec("tilbudsnr")) = 0 then
	    
        strtilbudsnr = oRec("tilbudsnr")

	    strSQLtb = "SELECT tilbudsnr FROM licens WHERE id = 1"
	    oRec3.open strSQLtb, oConn, 3
        if not oRec3.EOF then

        strNexttilbudsnr = oRec3("tilbudsnr") + 1
        tbStyle = "color:#999999;"

        end if
        oRec3.close

        else

        strtilbudsnr = oRec("tilbudsnr")
        strNexttilbudsnr = strtilbudsnr
	    tbStyle = "color:#000000;"
        end if
    
    intServiceaft = oRec("serviceaft")
	intkundekpers = oRec("kundekpers")
	
	jobans1 = oRec("jobans1")
	jobans2 = oRec("jobans2")

   
	
	'*** Ikke fakturerbare timer bruges ikke mere, men gemmes pga. gamle job ***'
	strBudgettimer = oRec("budgettimer") + oRec("ikkeBudgettimer")
	
	
		if oRec("ikkeBudgettimer") > 0 then
		ikkeBudgettimer = oRec("ikkeBudgettimer")
		else
		ikkeBudgettimer = 0
		end if
	end if
	
	lukafmjob = oRec("lukafmjob") 
	
    valuta = oRec("valuta")
    jobfaktype = oRec("jobfaktype")
	
	rekvnr = oRec("rekvnr")
	
	jo_gnstpris = oRec("jo_gnstpris")
	jo_gnsfaktor = oRec("jo_gnsfaktor")
	jo_gnsbelob = oRec("jo_gnsbelob")
	jo_bruttofortj = oRec("jo_bruttofortj")
	jo_dbproc = oRec("jo_dbproc")
	
    jo_udgifter_ulev = oRec("jo_udgifter_ulev")
    jo_udgifter_intern = oRec("jo_udgifter_intern")

    jo_bruttooms = oRec("jo_bruttooms")

	udgifter = oRec("udgifter")
	
	intprio = oRec("risiko")
	sdskpriogrp = oRec("sdskpriogrp")
	
	usejoborakt_tp = oRec("usejoborakt_tp")
	
	ski = oRec("ski")
	abo = oRec("abo")
	ubv = oRec("ubv")
	job_internbesk = oRec("job_internbesk")
	
	intSandsynlighed = oRec("sandsynlighed")

    diff_timer = oRec("diff_timer")
    diff_sum = oRec("diff_sum")
	
    restestimat = oRec("restestimat")
    stade_tim_proc = oRec("stade_tim_proc")

    virksomheds_proc = oRec("virksomheds_proc")

    syncslutdato = oRec("syncslutdato")
    lkdato = oRec("lukkedato")
    altfakadr = oRec("altfakadr")

    preconditions_met = oRec("preconditions_met")
    laasmedtpbudget = oRec("laasmedtpbudget")

  
    if len(trim(oRec("filepath1"))) <> 0 then
    filepath1 = oRec("filepath1")
    filepath1 = replace(filepath1, "#", "\")
    else
    filepath1 = "" 
    end if

    fomr_konto = oRec("fomr_konto")

    jfak_moms = oRec("jfak_moms")
    jfak_sprog = oRec("jfak_sprog")

	
	oRec.close
	
	
	headergif = "../ill/job_oprred_header.gif"
	dbfunc = "dbred"
	varSubVal = "opdaterpil" 
	rowspan = "28"
	
	strKundeId = strKnr
	
	editok = 0
	
	
	end if
	
	if func = "red" then



            if altfakadr = 0 then
            altfakadrCHK = ""
            else
            altfakadrCHK = "CHECKED"
            end if

	
	        if jobfaktype = 0 then
	        jobfaktypeCHK0 = "CHECKED"
	        jobfaktypeCHK1 = ""
	        else
	        jobfaktypeCHK0 = ""
	        jobfaktypeCHK1 = "CHECKED"
	        end if 
	
	else
	    
           


	    if lto = "dencker" then
	        
	        jobfaktypeCHK0 = ""
	        jobfaktypeCHK1 = "CHECKED"
	    
	    else
	    
	        if request.Cookies("tsa")("faktype") <> "" then
    	        
	            if request.Cookies("tsa")("faktype") = "0" then
	            jobfaktypeCHK0 = "CHECKED"
	            jobfaktypeCHK1 = ""
	            else
	            jobfaktypeCHK0 = ""
	            jobfaktypeCHK1 = "CHECKED"
	            end if 
    	        
	        else
    	        
	            jobfaktypeCHK0 = "CHECKED"
	            jobfaktypeCHK1 = ""
    	    
	        end if
	    
	    end if
	
	end if
	
	
	%>
	<!--#include file="inc/dato2.asp"-->
	<%
	'****************************************************************
	'*** Tjekker admin rettigeheder eller om man er jobanssvarlig ***
	'****************************************************************
	
	
	if level <= 2 OR level = 6 then
	editok = 1
	else
			if cint(session("mid")) = jobans1 OR cint(session("mid")) = jobans2 OR (cint(jobans1) = 0 AND cint(jobans2) = 0) then
			editok = 1
			end if
	end if
	
	
	if editok = 1 then
	
	
	
		if cint(intkundeok) = 1 OR cint(intkundeok) = 2 then
		kundechk = "checked"
		else
		kundechk = ""
		end if
	
	
		if strFakturerbart = 1 then
		varFakEks = "checked"
		varFakInt = ""
		else
		varFakEks = ""
		varFakInt = "checked"
		end if
		
		if strFastpris = "1" then
		varFastpris1 = "checked"
		varFastpris2 = ""
		else
		varFastpris1 = ""
		varFastpris2 = "checked"
		end if
	
	
	%>
    <input type="hidden" name="slto" id="slto" value="<%=lto %>" />
           
            
	<input type="hidden" name="func" id="func" value="<%=dbfunc%>">
	<input type="hidden" name="FM_OLDjobnr" value="<%=strJobnr%>">
	
	<%if showaspopup = "y" then%>
	<input type="hidden" name="sttid" value="<%=request("sttid")%>">
	<input type="hidden" name="dato" value="<%=request("dato")%>">
	<input type="hidden" name="datostkri" value="<%=request("datostkri")%>">
	<input type="hidden" name="datoslkri" value="<%=request("datoslkri")%>">
	<%end if%>
	
	<!--<input type="hidden" name="seraft" value="<=intServiceaft%>">-->
	    
    <%
	    if func = "red" then
	        if len(request("fm_kunde_sog")) <> 0 then
		    fm_kunde_sog = request("fm_kunde_sog")
		    else
		    fm_kunde_sog = 0
		    end if
		end if
	
		rowspan = rowspan
		if func = "opret" then
		tp_height = "1529"
		else
		tp_height = "1529"
		end if
	

              
	
	if request("showoprtxt") = "j" then 'AND request("int") <> 0 then%>
	<div id=visopretmed name=visopretmed style="position:absolute; left:140; top:20; width:300px; visibility:visible; z-index:100; background-color:#fffff1; border:1px darkred solid; padding:3px;">
	<font color="ForestGreen"><b>Jobbet er oprettet!</b></font><br> Du kan nu tilføje ressourcer, se gannt chart, tilføje flere aktiviteter eller klikke dig videre til <a href="jobs.asp?menu=job&shokselector=1" class=vmenu>joboversigten</a> med det samme.<br>
	<br><br>
	<a href="#" onClick="hideoprmed()" class=red>[X] Luk denne meddelelse.</a> 
	</div>
	<%end if%>
	
    <%if func = "red" then%>
	<tr>
		<td colspan=4 bgcolor="#ffffff" height="30" style="  padding-top:5px; padding-left:20px;">
		Sidst opdateret den <b><%=strLastUptDato%></b> af <b><%=strEditor%></b>
		</td>
	</tr>
    <%end if%>

 
        <%
        '**** kundeoplysninger, eksternt job *****'
	    call kundeopl
	
	    '******** Kontaktpersoner ***********'
	    call kundeopl_aft_kpers
        %>
       
	
	
	
	<tr bgcolor="#FFFFFF">
		
		<td style="padding:10px 5px 10px 10px; white-space:nowrap;" colspan="4">
            

         
        
            <div style="border:1px #ea0957 solid; width:430px; padding:10px;">
        <img src="../ill/ikon_job_24.png" width="24" height="24" alt="" border="0">&nbsp;<font color=red size=2>*</font> <b>Jobnavn:</b> 
        
                <!--
                <br /><a href="file:///C:\SK\skole\Fra Historie Online.docx" target="_blank">C:\SK\skole\Fra Historie Online.docx</a>

                <div class="fileupload">
                <input type="file" class="file" name="f1" style="width:400px;" />
                </div>
                <input  type="text" id="myInputText" value="C:\SK\skole\Fra Historie Online.docx" />
                -->

        <%if func <> "red" then %>
        (Jobnr. / tilbudsnr. tildeles automatisk)
        <%end if %>

        <br />
        <%if func <> "red" then 
        focol = "#999999"
        twd = 420
        else
        focol = "#000000"
        twd = 290
        end if%>
		<input type="text" name="FM_navn" id="FM_navn" value="<%=strNavn%>" style="width:<%=twd %>px; color:<%=focol%>;">
        <%if func = "red" then %>
        &nbsp;<font color=red size=2>*</font> Jobnr.: <input type="text" name="FM_jnr" id="FM_jnr" value="<%=strJobnr%>" style="width:80px; color:<%=focol%>; font-style:italic;">
        <%else %>
        <input type="hidden" name="FM_jnr" id="Text2" value="0">
        <%end if %>
        
        <br /><span style="font-size:10px; font-family:arial; color:#999999;">Jobnavn maks 100 karakterer. " (situationstegn) er ikke tilladt i jobnavn / jobnr. alfanumerisk</span>

        <br /><br />
      
                        <%
                        
                        select case lto
                        case "synergi1", "intranet - local", "outz", "dencker", "essens", "jttek"
                        
                        if func <> "red" then
                        opFolderCHK = "CHECKED"
                        else
                        opFolderCHK = ""
                        end if

                        case else

                        opFolderCHK = ""

                        end select

                        if len(trim(id)) <> 0 then
                        joidThis = id
                        else
                        joidThis = 0
                        end if

                        strSQLdok = "SELECT fo.navn FROM foldere fo  "_
                        &" WHERE fo.jobid =" & joidThis & " AND fo.jobid <> 0" 

                        folderfindes = 0
                        oRec2.open strSQLdok, oConn, 3
                        if not oRec2.EOF then
                        folderfindes = 1
                        end if
                        oRec2.close
                        
                        
                        if cint(folderfindes) = 0 then %>
                        
                        <input type="checkbox" value="1" name="FM_opret_folder" <%=opFolderCHK %> /> Opret folder med samme navn som jobbet.
                        

                        <%end if %>


                       


		</div>
        </td>
	</tr>
	<tr>
		<td style="">&nbsp;</td>
		<td valign=top style="padding:5px 5px 10px 0px;" colspan=2>
		<%if func = "red" then%>
		<b>Tilbudsnr.:</b> <br />
		<%else %>
		<b>Tilbud?</b><br />
		<%end if %>
		
		<%if cint(strStatus) = 3 then
					chkusetb = "CHECKED"
					else
					chkusetb = ""
					end if
					%>
					
					<input type="checkbox" id="FM_usetilbudsnr" name="FM_usetilbudsnr" value="j" <%=chkusetb%>>Dette job har status som tilbud.
		
	
			
					<%if func = "red" then
                        
                     if strNexttilbudsnr <> strtilbudsnr then 
                     tlbplcholderTxt = strNexttilbudsnr
                     else
                    tlbplcholderTxt = ""
                    end if %>

                    &nbsp;<span style="color:#999999">(Næste ledige nummer: <%=tlbplcholderTxt %>)</span>

					<br /><input type="text" id="FM_tnr" name="FM_tnr" value="<%=strtilbudsnr%>" style="<%=tbStyle%>" size="20"> 
                  
					<%else%>
					<input type="hidden" name="FM_tnr" value="0">
					<%end if%>
				
					<input type="hidden" id="FM_nexttnr" value="<%=strNexttilbudsnr %>">

					<br /><br /><input id="Text1" name="FM_sandsynlighed" value="<%=intSandsynlighed %>" type="text" style="width:30px;"/> % sandsynlighed for at vinde tilbud. &nbsp;
            <br /><span style="font-size:10px; font-family:arial; color:#999999;">(Pipelineværdi = Brutto oms. - Udgifter lev. * sandsynlighed)</span>
            
	       
	</td>
    <td style="" width=8>&nbsp;</td>
	</tr>	

  
                        <%select case lto 
                        case "essens", "xintranet - local"
                           %> 
                        <tr><td colspan="4" style="padding:0px 10px 10px 10px;">
                        Sti på filserver:<br /> <input type="text" name="FM_filepath1" placeholder="c:\..." style="width:450px;" value="<%=filepath1 %>" /> <br /><br />
                        </td></tr>
               <%   

                        end select
                        %>
	
    	
	<tr bgcolor="#FFFFFF">
		<td colspan="4" style="padding:10px 10px 2px 10px;">
		<b>Tilknyt job til aftale?</b> <a href="#" onClick="serviceaft('0', '<%=strKundeId%>', '', '0')" class=vmenu>+ Opret ny aftale (reload)</a> <br>
		
		<%
		strSQL2 = "SELECT id, enheder, stdato, sldato, status, navn, pris, perafg, "_
		&" advitype, advihvor, erfornyet, varenr, aftalenr FROM serviceaft "_
		&" WHERE kundeid = "& strKundeId &" OR id = "& intServiceaft &" ORDER BY id DESC" 'AND status = 1
		
		'Response.write strSQL2
		%>
		<select name="FM_serviceaft" id="FM_serviceaft" style="width:450px;">
		<option value="0">Nej (fjern)</option>
		
		<%
		
		oRec2.open strSQL2, oConn, 3 
		while not oRec2.EOF 
		
		if oRec2("advitype") <> 0 then
		udldato = "&nbsp;&nbsp;&nbsp; startdato: " & formatdatetime(oRec2("stdato"), 2) 
		else
		udldato = ""
		end if%>
		
		<%
		if cint(intServiceaft) = cint(oRec2("id")) then
		serChk = "SELECTED"
		else
		serChk = ""
		end if
		
		if oRec2("status") <> 0 then
		stThis = "Aktiv"
		else
		stThis = "Lukket"
		end if%>
		<option value="<%=oRec2("id")%>" <%=serChk%>> <%=oRec2("navn")%>&nbsp;(<%=oRec2("aftalenr") %>)  <%=udldato%> (<%=stThis%>) </option>
		<%
		oRec2.movenext
		wend
		oRec2.close
		%>
		</select>
		
		<%if func = "red" then %><br>
		<input type="checkbox" name="FM_overforGamleTimereg" value="1"> Overfør (fjern) eksisterende 
		<b>timeregistreringer</b> og <b>materiale-forbrug</b> til den valgte aftale. 
      
        
        <!--<br />
		<br />Hvis der <b>ændres</b> aftale tilknytning undervejs i jobbets levetid, vil evt. oprettede
		fakturaer på jobbet ikke længere kunne ses under den <b>gamle aftale.</b> (faktura historik, aftale afstemning)<br />
		Fakturaer oprettet direkte på aftalen vil ikke blive berørt. -->
		<%end if%>
		  <br>&nbsp;</td>
		
	</tr>
	
	
	<%
	'**** Beskrivelse *****
	if showaspopup <> "y" then
	bcols = 62
	else
	bcols = 52
	end if%>
	<tr>
		<td style="">&nbsp;</td>
		<td style="padding:5px 0px 0px 2px;" colspan=2><b>Job beskrivelse:</b>
		<%if showaspopup <> "y" then %>
		&nbsp;<a href="#" id="a_internnote" class="vmenu"> + Intern note</a><br>
		<%end if%>
		
		                <%
		                dim content

	                   select case lto 
                            case "synergi1", "xintranet - local"
                                if func = "red" then 
                            content = strBesk
                            else
                            content = "Deadline:<br><br>Tidsestimat:"
                            end if
                            case else
	                    content = strBesk
                            end select
            			
            			
			            
			            Set editorK = New CuteEditor
            					
			            editorK.ID = "FM_beskrivelse"
			            editorK.Text = content
			            editorK.FilesPath = "CuteEditor_Files"
			            
                        select case lto
                        case "dencker", "jttek", "intranet - local"
                        editorK.AutoConfigure = "Compact"
                        case else
			            editorK.AutoConfigure = "Minimal"
                        end select
            			
			            editorK.Width = 450
			            editorK.Height = 280
			            editorK.Draw()
		                %>
		                
		    <br />
            &nbsp;
            
                           
                            
		
		</td>
		<td>
		
		<%if showaspopup <> "y" then %>
		<!-- intern note -->
        <div id="div_internnote" style="position:absolute; top:420px; left:20px; visibility:hidden; padding:20px; border:10px #CCCCCC solid; background-color:#FFFFFF; width:650px; z-index:2000000;">
	    
                           <table cellpadding="0" cellspacing="0" border="0" width=100%>
							  <tr>
							  <td style="padding:15px 20px 2px 10px;"><b>Intern note:</b><br />
							  Brug denne information til interne beskeder og arbejdsbeskrivelser. Vises på timereg. siden og ved faktura oprettelse.</td>
	                            <td align=right style="padding:0px 20px 2px 0px;"><a href="#" id="a2_internnote" class=red>[x]</a></td>
	                        </tr>
                            
                            	<tr>
		
		<td style="padding-left:5px; padding-top:6px; padding-bottom:4px;" colspan=2>
		
		                <%
		                dim content2
	                    content2 = job_internbesk
            			
			            
			            Set editorK = New CuteEditor
            					
			            editorK.ID = "FM_internbesk"
			            editorK.Text = content2
			            editorK.FilesPath = "CuteEditor_Files"
                        select case lto
                        case "dencker"
                        editorK.AutoConfigure = "Compact"
                        case else
			            editorK.AutoConfigure = "Minimal"
                        end select
            			
			            editorK.Width = 420
			            editorK.Height = 220
			            editorK.Draw()
		                %>
		                
		    <br />
            &nbsp;
		
		    </td>
		   
	    </tr>
        </table>
        </div>
		<%end if%>
		
		&nbsp;</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td style="padding:0px 20px 5px 0px; width:140px;"><b>Rekvisitions Nr.:</b><br />
        (indkøbsordrenr)</td>
		<td style="padding:0px 0px 5px 5px;"><input type="text" name="FM_rekvnr" id="FM_rekvnr" value="<%=rekvnr%>" size="40"></td>
		<td>&nbsp;</td>
	</tr>
	
    <%if (lto = "execon" OR lto = "immenso") then %>

	<%if cint(ski) = 1 then
	skiCHK = "CHECKED"
	else
	skiCHK = ""
	end if
	 %>
	<tr>
		<td style="" width=8>&nbsp;</td>
		<td colspan=2 style="padding:0px 20px 10px 0px;"><input id="Checkbox1" name="FM_ski" type="checkbox" value="1" <%=skiCHK %> /> <b>SKI aftale</b>&nbsp;
		(Staten og kommunernes indkøbs service)</td>
	    <td style="" width=8>&nbsp;</td>
	</tr>
	
	
	
	<%if cint(abo) = 1 then
	aboCHK = "CHECKED"
	else
	aboCHK = ""
	end if
	 %>
	<tr>
		<td style="" width=8>&nbsp;</td>
		<td colspan=2 style="padding:0px 20px 10px 0px;"><input id="FM_abo" name="FM_abo" type="checkbox" value="1" <%=aboCHK %> /> <b>Abonnement</b>&nbsp;(Lightpakke)</td>
	    <td style="" width=8>&nbsp;</td>
	</tr>
	
	
	<%if cint(ubv) = 1 then
	ubvCHK = "CHECKED"
	else
	ubvCHK = ""
	end if
	 %>
	<tr>
		<td style="" width=8>&nbsp;</td>
		<td colspan=2 style="padding:0px 20px 10px 0px;"><input id="Checkbox2" name="FM_ubv" type="checkbox" value="1" <%=ubvCHK %> /> <b>Udbudsvagten</b>&nbsp;
		(Job omfattet af Udbudsvagten)</td>
	    <td style="" width=8>&nbsp;</td>
	</tr>
	
    <%end if%>
   

	  <%if lto = "dencker" then
                    '*** Dencker opret SDSK incident ****' 
                 
                 call erSDSKaktiv() 
	
	
	
	                '**** Opret Incident på job ****'
	                if dsksOnOff <> 0 AND showaspopup <> "y" ANd multibletrue <> 1 AND sdskpriogrp <> 0 then
	
	                if lto = "xdencker" AND func <> "red" then
	                opIncCHK = "CHECKED"
	                else
	                opIncCHK = ""
	                end if%>
	                <tr bgcolor="#ffffff">
		                <td colspan=4 style="padding:10px 0px 5px 10px;">
                            <input id="FM_opr_incident" type="checkbox" name="FM_opr_incident" value="1" <%=opIncCHK %>/> <b>Opret tilhørende Incident</b> (opgave under ServiceDesk)<br />&nbsp;</td>
	                </tr>
	                <%end if 
                 
                 
        end if         %>				
					
					
					
					
                   
					
					
				
					
					
					<tr bgcolor="#FFFFFF">
						<td>&nbsp;</td>
						<td colspan=2 valign=top style="padding:10px 5px 0px 0px; white-space:nowrap;">
						<h3>Status og periode</h3></td>
						<td>&nbsp;</td>	
				   </tr>
				   <tr bgcolor="#FFFFFF">
				        <td>&nbsp;</td>
                        <td valign=top style="padding-top:3px; padding-bottom:4px;">
                        <b>Status:</b></td>
						<td valign=top style="padding-top:0px; padding-bottom:4px;">
						
						             <select name="FM_status" id="FM_status" style="width:160px;">
									<%
                                    lkDatoThis = ""
                                    if dbfunc = "dbred" then 
									select case strStatus
									case 1
									strStatusNavn = "Aktiv"
									case 2
									strStatusNavn = "Til Fakturering" 'passiv
									case 0
									strStatusNavn = "Lukket"
                                        if cdate(lkdato) <> "01-01-2002" then
                                        lkDatoThis = " ("& formatdatetime(lkdato, 2) & ")"
                                        end if
									case 3
									strStatusNavn = "Tilbud"
                                    case 4
									strStatusNavn = "Gennemsyn"
									end select
									%>
									<option value="<%=strStatus%>" SELECTED><%=strStatusNavn%> <%=lkDatoThis %></option>
									<%end if
                                    
                                  
                                    %>
									<option value="1">Aktiv</option>
									<option value="2">Til Fakturering</option> <!-- Passiv -->
									<option value="0">Lukket</option>
                                    <option value="4">Gennemsyn</option>
									
									<option value="3">Tilbud</option>
									
									</select> 

                                    <%if strStatus = 3 then
                                    %>
                                    <br /><span style="color:#999999;">Skifter automatisk til aktiv, hvis tilbuds status fravælges i toppen.</span>
                                    <%
                                    end if %>
									
						<td>&nbsp;</td>	
				   </tr>		
			       
                   
                  
					
					
		
	
								<%if showaspopup <> "y" then%>
									<tr bgcolor="#ffffff">
										<td style="">&nbsp;</td>
										<td style="padding-top:4px;"><b>Start dato:</b> <br /><span style="color:#999999; font-size:9px;">job er tilgængelig på timereg. siden fra denne dato</span></td>
										<td style="padding-top:4px;"><select name="FM_start_dag" id="FM_start_dag">
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
										<option value="31">31</option></select>&nbsp;
										
										<select name="FM_start_mrd" id="FM_start_mrd">
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
										
										
										<select name="FM_start_aar" id="FM_start_aar">
										<option value="<%=strAar%>">
										<%if id <> 0 then%>
										20<%=strAar%>
										<%else%>
										<%=strAar%>
										<%end if%></option>
										
										<%for x = -10 to 15
		                                useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
		                                  <option value="<%=right(useY, 2)%>"><%=useY%></option>
		                                <%next %>
										
									  
										</select>&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=6')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a></td>
										
										
										<td style="">&nbsp;</td>
										</tr>
										<tr bgcolor="#ffffff">
											<td style="padding-top:0px; padding-left:5px;" rowspan=2 align=right>&nbsp;</td>
											<td valign=top style="padding-top:5px; padding-bottom:4px;"><b>Slut dato:</b>&nbsp;
											</td>
											<td style="padding-bottom:4px;"><select name="FM_slut_dag" id="FM_slut_dag">
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
										<option value="31">31</option></select>&nbsp;
										
										<select name="FM_slut_mrd" id="FM_slut_mrd">
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
										
										
										<select name="FM_slut_aar" id="FM_slut_aar">
										<option value="<%=strAar_slut%>">
										<%if id <> 0 then%>
										20<%=strAar_slut%>
										<%else%>
										<%=strAar_slut%>
										<%end if%></option>
										
									   
									   <%for x = -10 to 15 
		                                useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
		                                <option value="<%=right(useY, 2)%>"><%=useY%></option>
		                                <%next %>
									   
									   </select>&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=5')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>&nbsp;&nbsp;
												<br />
												<%if strAar_slut = 44 then
												chald = "checked"
												else
												chald = ""
												end if%>
												<input type="checkbox" name="FM_datouendelig" value="j" <%=chald%>> Aldrig: (1. jan 2044)
                                                
                                               </td>
										<td style="">&nbsp;</td>
									</tr>
                                     
									<tr bgcolor="#ffffff">
										
										<td colspan=2 style="padding:10px 5px 10px 0px;"><b>Beregn slutdato.</b>
										 (angiv antal uger jobbet skal løbe over)
										 <input type="text" name="FM_antaluger" id="FM_antaluger" size="4" style="font-size:9px;">
										 <input type="button" value="beregn" onClick="beregnuger();" style="font-size:9px;">&nbsp;&nbsp;
										</td>
										<td style="">&nbsp;</td>
									</tr>

                                    <tr bgcolor="#ffffff"><td colspan=4 style="padding:10px 0px 0px 8px;">
                                                <%if syncslutdato = 1 then
                                                syncslutdatoCHK = "CHECKED"
                                                else
                                                syncslutdatoCHK = ""
                                                end if %>
                                                <b>Hold datoer aktuelle:</b>
                                                 <br />	<input type="checkbox" name="FM_syncslutdato" value="j" <%=syncslutdatoCHK %>> Sync. slutdato, så den følger sidste timereg. / sidste faktura / dd. (ved luk job) 
                                                  <br />	<input type="checkbox" name="FM_syncaktdatoer" value="j"> Sync. start- og slut- datoer på aktiviteter, så de nedarver fra job.<br />&nbsp;
                                       </td></tr>
									
	<tr bgcolor="#FFFFFF">
        <td colspan=4 style="padding:10px 0px 10px 8px;">
        <b>Restestimat:</b>&nbsp;&nbsp;<input id="FM_restestimat" name="FM_restestimat" type="text" size="6" value="<%=restestimat %>" />&nbsp;
        <select id="FM_stade_tim_proc" name="FM_stade_tim_proc">
           
           <%
           sele0 = ""
           sele1 = ""
           
           select case stade_tim_proc
           case 1
           sele1 = "SELECTED"
           case else
           sele0 = "SELECTED"
           end select
            %>
           
            <option <%=sele0 %> value="0">timer til rest.</option>
            <option <%=sele1 %> value="1">% afsluttet</option>
        </select><br />&nbsp;
       </td>
    </tr>    
									
									
								<%else
								'*** Start og slut dato sættes = med hinanden når der oprettes fra ressource booking 
								'*** Ændres senere til den dato der vælges i jbpla_w_opr.asp ***
								strDag = day(request("dato"))
								strMrd = month(request("dato"))
								strAar = year(request("dato"))%>
								<input type="hidden" name="FM_start_dag" value="<%=strDag%>">
								<input type="hidden" name="FM_start_mrd" value="<%=strMrd%>">
								<input type="hidden" name="FM_start_aar" value="<%=strAar%>">
								<input type="hidden" name="FM_slut_dag" value="<%=strDag%>">
								<input type="hidden" name="FM_slut_mrd" value="<%=strMrd%>">
								<input type="hidden" name="FM_slut_aar" value="<%=strAar%>">
								<%end if%>
	
	
	
	<!-- fordel timepriser -->
	
	
	
	
	<!-- Til stam akt. grupper -->
	
	<!-- stam akt tilde SLUT --->
	
	
	                 <!-- Forkalk. tmer Nettoomsætning og brutto omsætning ---->
	                <input type="hidden" name="FM_fakturerbart" value="1">
	
					<%if showaspopup <> "y" then%>
				   <input id="jobfaktype" name="FM_jobfaktype" value="0" type="hidden" />   
	
	
	
	
						
			
                                <%  
                            call salgsans_fn()
                             
                            if cint(showSalgsAnv) = 1 then 
                            jobansvHeaderTxt = "Job- og Salgs -ansvarlige"
                            selWdt = "80px"
                            else
                            jobansvHeaderTxt = "Jobansvarlig og jobejer"
                            selWdt = "240px"       
                            end if
                                    
                            select case lto 
                            case "epi", "epi_ab", "epi_no", "epi_sta", "intranet - local", "epi_uk"
                            mTypeExceptSQL = " AND (medarbejdertype <> 14 AND medarbejdertype <> 24)"
                            case else
                            mTypeExceptSQL = ""         
                            end select      
                            %>
			
	               
					<tr>
						<td style="">&nbsp;</td>
						<td colspan="2" style="padding-left:5px; padding-top:15px;" valign=top><h3><%=jobansvHeaderTxt %>
						
						<span style="font-size:9px; font-style:normal; font-weight:normal; line-height:12px;">- Hvis der angives jobansvarlig og / eller jobejer er det kun disse (eller administratorer) der kan 
						redigere jobbet, og oprette fakturaer på jobbet. Jobmedansvarlig  1-3 bruges til salg og bonus beregninger mm.</span></h3>
						
                            <table style="width:100%;"><tr>  <td>
                           


						<table style="width:100%;">
						
						
						<%for ja = 1 to 5 
					
						select case ja
						case 1
						'jbansImg = "<img src='../ill/ac0019-24.gif' width='24' height='24' alt='Jobansvarlig' border='0'>"
                        if cint(showSalgsAnv) = 1 then 
						jbansTxt = "Jobansv."
                        else
                        jbansTxt = "Jobansvarlig"
                        end if
						jobansField = "jobans1, jobans_proc_1"
						case 2
						'jbansImg = "<img src='../ill/ac0020-24.gif' width='24' height='24' alt='Jobejer' border='0'>"
						jbansTxt = "Jobejer"
						jobansField = "jobans2, jobans_proc_2"
						case 3
						'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
						jbansTxt = "Medansv. 1"
						jobansField = "jobans3, jobans_proc_3"
						case 4
						'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
						jbansTxt = "Medansv. 2"
						jobansField = "jobans4, jobans_proc_4"
						case 5
						'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
						jbansTxt = "Medansv. 3"
						jobansField = "jobans5, jobans_proc_5"
						end select
						
						%>
						<tr>
						<!--<td><%=jbansImg  %></td>-->
						<td>
						<b><%=jbansTxt %>:</b></td><td>
						&nbsp;&nbsp;<select name="FM_jobans_<%=ja %>" id="FM_jobans_<%=ja %>" style="width:<%=selWdt%>;">
						<option value="0">Ingen</option>
							<%

                            select case lto
                            case "hestia", "intranet - local" 
							mPassivSQL = " OR mansat = 3"
                            case else
                            mPassivSQL = ""
                            end select

							if func <> "red" then
							strSQL = "SELECT mnavn, mnr, mid, init, mansat FROM medarbejdere WHERE ((mansat = 1 "& mPassivSQL &")" & mTypeExceptSQL & ")"
							else
							strSQL = "SELECT mnavn, mnr, mid, mansat, "& jobansField &", init FROM medarbejdere "_
							&" LEFT JOIN job ON (job.id = "& id &") WHERE ((mansat = 1 "& mPassivSQL &") "& mTypeExceptSQL &") OR (mid = jobans"& ja &")"
							end if

                            strSQL = strSQL & " ORDER BY mnavn"

							oRec.open strSQL, oConn, 3 
							while not oRec.EOF 
							
							if func <> "red" then
							  if ja = 1 then
							  usemed = session("mid")
                              jobans_proc = 100
							  else
							  usemed = 0
                              jobans_proc = 0
							  end if

                              


							else
							    select case ja
						        case 1
						        usemed = oRec("jobans1")
                                jobans_proc = oRec("jobans_proc_1")
						        case 2
						        usemed = oRec("jobans2")
                                jobans_proc = oRec("jobans_proc_2")
						        case 3
						        usemed = oRec("jobans3")
                                jobans_proc = oRec("jobans_proc_3")
						        case 4
						        usemed = oRec("jobans4")
                                jobans_proc = oRec("jobans_proc_4")
						        case 5
						        usemed = oRec("jobans5")
                                jobans_proc = oRec("jobans_proc_5")
						        end select
							end if
							
								if cint(usemed) = oRec("mid") then
								medsel = "SELECTED"
								else
								medsel = ""
								end if

                                if len(trim(oRec("init"))) <> 0 then
                                opTxt = oRec("init") &" - "& oRec("mnavn")
                                else
                                opTxt = oRec("mnavn") '&" (" & oRec("mnr") &")"
                                end if

                                if oRec("mansat") <> 1 then
                                select case oRec("mansat")
                                case 2
                                opTxt = opTxt & " - Deaktiveret"
                                case 3 
                                opTxt = opTxt & " - Passiv"
                                end select
                                end if

                            %>
							<option value="<%=oRec("mid")%>" <%=medsel%>><%=opTxt %></option>
							<%

							oRec.movenext
							wend
							oRec.close 
							%>
						</select>
						</td><td style="white-space:nowrap;"><input id="FM_jobans_proc_<%=ja %>" name="FM_jobans_proc_<%=ja %>" value="<%=formatnumber(jobans_proc, 1) %>" type="text" style="width:40px;" /> %</td></tr>
                            
						<%next %>
						
						
						</table>
						

                             <%
                                
                           
                            if cint(showSalgsAnv) = 1 then %>
                            </td><td>
                           <table style="width:100%;">
						
						
						<%for sa = 1 to 5 
					
						select case sa
						case 1
						'jbansImg = "<img src='../ill/ac0019-24.gif' width='24' height='24' alt='Jobansvarlig' border='0'>"
						saansTxt = "Salgsansv. 1"
						salgsansField = "salgsans1, salgsans1_proc"
						case 2
						'jbansImg = "<img src='../ill/ac0020-24.gif' width='24' height='24' alt='Jobejer' border='0'>"
						saansTxt = "Salgsansv. 2"
						salgsansField = "salgsans2, salgsans2_proc"
						case 3
						'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
						saansTxt = "Salgsansv. 3"
						salgsansField = "salgsans3, salgsans3_proc"
						case 4
						'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
						saansTxt = "Salgsansv. 4"
						salgsansField = "salgsans4, salgsans4_proc"
						case 5
						'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
						saansTxt = "Salgsansv. 5"
						salgsansField = "salgsans5, salgsans5_proc"
						end select
						
						%>
						<tr>
						<!--<td><%=jbansImg  %></td>-->
						<td>
						<b><%=saansTxt %>:</b></td><td>
						&nbsp;&nbsp;<select name="FM_salgsans_<%=sa %>" id="Select4" style="width:<%=selWdt%>;">
						<option value="0">Ingen</option>
							<%
							
							if func <> "red" then
							strSQL = "SELECT mnavn, mnr, mid, init, mansat FROM medarbejdere WHERE (mansat = 1 " & mTypeExceptSQL &")"
							else
							strSQL = "SELECT mnavn, mnr, mid, mansat, "& salgsansField &", init FROM medarbejdere "_
							&" LEFT JOIN job ON (job.id = "& id &") WHERE (mansat = 1 "& mTypeExceptSQL &") OR (mid = salgsans"& sa &")"
							end if

                            strSQL = strSQL & " ORDER BY mnavn"

							
                                
                            oRec.open strSQL, oConn, 3 
							while not oRec.EOF 
							
							if func <> "red" then
							  if sa = 1 then
							  usemed = session("mid")
                              salgsans_proc = 0
							  else
							  usemed = 0
                              salgsans_proc = 0
							  end if

                              

							else
							    select case sa
						        case 1
						        usemed = oRec("salgsans1")
                                salgsans_proc = oRec("salgsans1_proc")
						        case 2
						        usemed = oRec("salgsans2")
                                salgsans_proc = oRec("salgsans2_proc")
						        case 3
						        usemed = oRec("salgsans3")
                                salgsans_proc = oRec("salgsans3_proc")
						        case 4
						        usemed = oRec("salgsans4")
                                salgsans_proc = oRec("salgsans4_proc")
						        case 5
						        usemed = oRec("salgsans5")
                                salgsans_proc = oRec("salgsans5_proc")
						        end select
							end if
							
								if cint(usemed) = oRec("mid") then
								medsel = "SELECTED"
								else
								medsel = ""
								end if


						        if len(trim(oRec("init"))) <> 0 then
                                opTxt = oRec("init") &" - "& oRec("mnavn")
                                else
                                opTxt = oRec("mnavn") '&" (" & oRec("mnr") &")"
                                end if

                                if oRec("mansat") <> 1 then
                                select case oRec("mansat")
                                case 2
                                opTxt = opTxt & " - Deaktiveret"
                                case 3 
                                opTxt = opTxt & " - Passiv"
                                end select
                                end if

                            %>
							<option value="<%=oRec("mid")%>" <%=medsel%>><%=opTxt %></option>
							<%
							oRec.movenext
							wend
							oRec.close 
							%>
						</select>
						</td><td style="white-space:nowrap;"><input id="FM_salgsans_proc_<%=sa %>" name="FM_salgsans_proc_<%=sa %>" value="<%=formatnumber(salgsans_proc, 1) %>" type="text" style="width:40px;" /> %</td></tr>
                            
						<%next %>
						
						
						</table>
                            <%end if %>

                            </td></tr></table>







						<br><br>
						<%
                        
                        if level = 1 then 
                             if lto <> "epi" OR (lto = "epi" AND (thisMid = 6 OR thisMid = 11 OR thisMid = 59 OR thisMid = 1 OR thisMid = 1720)) OR (lto = "epi_no" AND thisMid = 2) OR (lto = "epi_ab" AND thisMid = 2) OR (lto = "epi_sta" AND thisMid = 2) OR (lto = "epi_cati" AND thisMid = 2) OR (lto = "epi_uk" AND thisMid = 2) then%>
                            <input type="text" name="FM_virksomheds_proc" value="<%=formatnumber(virksomheds_proc, 0) %>" style="width:40px;" /> Virksomhedsandel af salg i %
                            <%else %>
                              <input type="hidden" name="FM_virksomheds_proc" value="<%=formatnumber(virksomheds_proc, 0) %>"/>
                            <%end if %>
                        <%else %>
                        <input type="hidden" name="FM_virksomheds_proc" value="<%=formatnumber(virksomheds_proc, 0) %>"/>
                        <%end if %>
					    
                        <br /><br />
                        <%if func <> "red" then

                            select case lto
                            case "synergi1", "qwert", "hestia"
                            advJobansCHK = "CHECKED"
                            case "dencker"

                                if session("mid") = 34 OR session("mid") = 49 then 'Aministrationen 
                                advJobansCHK = "CHECKED"
                                else
                                advJobansCHK = ""
                                end if
                            
                            case else
                            advJobansCHK = ""
                            end select
                        
                        else
                            
                            select case lto
                            case "hestia"
                            advJobansCHK = "CHECKED"
                            case else
                            advJobansCHK = ""
                            end select

                        end if %>

                        <input type="checkbox" value="1" name="FM_adviser_jobans" <%=advJobansCHK %> />Adviser valgte medarbejdere om at de er tilføjet som jobansvarlig / jobejer.

                        <br><br>&nbsp;
						</td>
						<td style="">&nbsp;</td>
					</tr>
					
		       
		
		        <!-- Avanceret indstillinger -->
				<tr bgcolor="#FFFFFF">
								<td colspan=4 style="padding:20px 0px 0px 10px;">
                              

                            
						 <%'**** Projektgrupper *****%>
                           
                            <table cellpadding=0 cellspacing=0 border=0 width=100%>
							<tr>
								<td style="">&nbsp;</td>
								<td colspan=2 style="padding-left:5px;">
                                 <h4>Projektgrupper <span style="font-size:9px; font-style:normal; font-weight:normal; line-height:11px;">- Hvem skal kunne registrere tid på dette job?</span></h4>
                                
								<% 
                                uTxt = "Hvis der fjernes en projektgruppe vil realiserede timer"_
                                &" på medarbejdere der kun er med i denne projektgruppe ikke længere kunne ses ved fakturaoprettelse."
								uWdt = 300
								
								call infoUnisport(uWdt, uTxt) 

                              %>

								<br>&nbsp;</td>
								<td style="">&nbsp;</td>
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
								<td style="">&nbsp;</td>
								<td valign="top" style="padding-left:5px; width:140px;">Projektgruppe <%=p%>:</td>
								<td><select name="FM_projektgruppe_<%=p%>" style="width:200px;">
								<%
									strSQL = "SELECT id, navn FROM projektgrupper ORDER BY navn"
									oRec.open strSQL, oConn, 3
									
									While not oRec.EOF 
									projgId = oRec("id")
									projgNavn = oRec("navn")
									
                                    'if projgNavn = "Alle" then
                                    'projgNavn = "Alle-gruppen (alle medarbejdere)"
                                    'end if

									if cint(strProj) = cint(projgId) then
									varSelected = "SELECTED"
									gp = strProj
									else
									varSelected = ""
									gp = gp
									end if
									%>
									<option value="<%=projgId%>" <%=varSelected%>><%=projgNavn%></option>
									<%
									oRec.movenext
									wend
									oRec.close%>
						</select>
                        
                       
                        
                        
                        </td>
							<td style="">&nbsp;</td>
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
							
							next
				
				


                %>
               <tr><td colspan=4 style="padding:20px 60px 40px 10px;">

                            
                     <%if func = "red" then
                                    
                                    if lto = "execon" OR lto = "immenso" OR lto = "synergi1" then
                                    syncAktProjGrpCHK = "CHECKED"
                                    else
                                    syncAktProjGrpCHK = ""
                                    end if
                                
                                %>
                                
                                
                               
								<input type="checkbox" name="FM_opdaterprojektgrupper" id="FM_opdaterprojektgrupper" value="1" <%=syncAktProjGrpCHK %>> <b>Overfør</b> (synkroniser) valgte projektgrupper, 
                                til <b>aktiviteterne</b> på dette job. 
								<br />
                                <%end if%>
                
                  <%
                           if lto <> "execon" then%>
                                <br>
								<input type="checkbox" name="FM_gemsomdefault" id="FM_gemsomdefault" value="1"><b>Skift standard forvalgt projektgruppe</b> til den gruppe der her vælges som projektgruppe 1.
								<span style="color:#999999;">Gemmes som cookie i 30 dage.</span>
                                <%end if %>
								
                               
                              

                                <%if func <> "red" then
                                 
                                    if lto = "jm" OR lto = "synergi1" OR lto = "micmatic" OR lto = "krj" then 'OR lto = "lyng" OR lto = "glad" then
                                    forvalgCHK = "CHECKED"
                                    else
                                    forvalgCHK = ""
                                    end if
                                 
                                 else
                                 
                                 forvalgCHK = ""

                                 end if %>

                                 <br /><br />
                              
								<input type="checkbox" name="FM_forvalgt" id="FM_forvalgt" <%=forvalgCHK %> value="1"><b>Tilføj "Push" job til aktiv-jobliste.</b> 
                                Gør dette job aktivt på aktiv-joblisten for alle de medarbejdere der er med i de ovenstående valgte projektgrupper. 
                              
                              
                                                               
                                 

                                 </td></tr>
                       </table>
                    



                </td></tr>
                
                <!-- forretningsområder --->
                <%
                
                
                if func = "red" then
                                    
                    strFomr_rel = ""
                    strFomr_navn = ""
                    
                    strSQLfrel = "SELECT for_fomr, fomr.navn FROM fomr_rel "_
                    &" LEFT JOIN fomr ON (fomr.id = for_fomr) WHERE for_jobid = "& id & " AND for_aktid = 0 GROUP BY for_fomr"

                    'Response.Write strSQLfrel
                    'Response.flush
                    f = 0
                    oRec.open strSQLfrel, oConn, 3
                    while not oRec.EOF

                    if f = 0 then
                    strFomr_navn = ""
                    end if

                    strFomr_rel = strFomr_rel & "#"& oRec("for_fomr") &"#"
                    strFomr_navn = strFomr_navn & oRec("navn") & ", " 

                    f = f + 1
                    oRec.movenext
                    wend
                    oRec.close

                    if f <> 0 then
                    len_strFomr_navn = len(strFomr_navn)
                    left_strFomr_navn = left(strFomr_navn, len_strFomr_navn - 2)
                    strFomr_navn = left_strFomr_navn & ")"

                        if len(strFomr_navn) > 50 then
                        strFomr_navn = left(strFomr_navn, 50) & "..)"
                        end if
                    end if

                else

                '**** forvalgt forretningsområder ****' 
                select case lto 
                case "hestia", "xintranet - local"
                    
                        
                            strSQLfrel = "SELECT id, navn FROM fomr WHERE id = 2"    
                           
                        
                            f = 0
                            oRec.open strSQLfrel, oConn, 3
                            while not oRec.EOF

                            'if f = 0 then
                            'strFomr_navn = " ("
                            'end if

                            strFomr_rel = strFomr_rel & "#"& oRec("id") &"#"
                            strFomr_navn = strFomr_navn & oRec("navn") & ", " 

                            f = f + 1
                            oRec.movenext
                            wend
                            oRec.close     

                case else
                     strFomr_navn = ""
                     strFomr_rel = ""
                end select


                end if
                
                if lto <> "execon" AND lto <> "xintranet - local" then %>

                <tr bgcolor="#FFFFFF">
								<td colspan=4 style="padding:10px 0px 0px 10px;">
                                <a href="#" class=vmenu id="a_tild_forr">+ Forretningsområder </a><br />
                                Forretningsområder bruges bl.a. til at se tidsforbrug på tværs af kunder og job, og til at se hvilke slags opgaver man bruger sin tid på. 
                                  <br /><br />
                                  Alle aktiviteter på dette job tæller altid med i de forretningsområder der er valgt på jobbet. Specifikke forretningsområder kan vælges på den enkelte aktivitet.  
                                  <br /><br />
                              
                                                      
                                <%
                                 ' uTxt = "Forretningsområder bruges bl.a. til at se tidsforbrug på tværs af kunder og job, og til at se hvilke slags opgaver man bruger sin tid på."
                                 ' uWdt = 300
								
								'call infoUnisport(uWdt, uTxt) 
                                
                                    
                                    call fomr_mandatory_fn()

                                    div_tild_forr_Pos = "relative"
                                    div_tild_forr_Lft = "0px"
                                    div_tild_forr_Top = "0px"
                                    div_tild_forr_z = "0"

                                    if cint(fomr_mandatoryOn) = 1 then
                                    div_tild_forr_VZB = "visible"
                                    div_tild_forr_DSP = ""

                                        if func = "opret" AND step = "2" then
                                        div_tild_forr_bdr = "10"
                                        div_tild_forr_pd = "20"
                                        else
                                        div_tild_forr_bdr = "0"
                                        div_tild_forr_pd = "20"
                                        end if

                                        
                                        if func = "opret" AND step = "2" then
                                        div_tild_forr_Pos = "absolute"
                                        div_tild_forr_Lft = "600px"
                                        div_tild_forr_Top = "500px"
                                        div_tild_forr_z = "400000"
                                        end if

                                    else
                                    div_tild_forr_VZB = "hidden"
                                    div_tild_forr_DSP = "none"

                                        div_tild_forr_bdr = "0"
                                        div_tild_forr_pd = "0"

                                    end if
                                    %>

                                

						
                           
                            <div id="div_tild_forr" style="position:<%=div_tild_forr_Pos%>; left:<%=div_tild_forr_Lft%>; top:<%=div_tild_forr_Top%>; z-index:<%=div_tild_forr_z%>; width:390px; visibility:<%=div_tild_forr_VZB%>; display:<%=div_tild_forr_DSP%>; border:<%=div_tild_forr_bdr%>px #CCCCCC solid;  padding:<%=div_tild_forr_pd%>px; background-color:#FFFFFF;">
                              <!--Forretningsområder er baseret på de forretningsområder der er tildelt på aktiviteteterne: <br /><b><%=strFomr_navn%></b>-->
                            
                            <table cellpadding=0 cellspacing=0 border=0 width=100%>
                                <% if cint(fomr_mandatoryOn) = 1 AND func = "opret" AND step = "2" then%>
                                <tr><td align="right"> <a href="#" class=vmenu id="a_tild_forr2" style="color:red;">[X]</a></td></tr>
                                <%end if %>
                            <tr><td valign=top>
                            <h4>Forretningsområder: <span style="font-size:11px; font-weight:lighter;">(konto)</span></h4>  
                                
                               
                                                            
                                <%

                                '** Finder kundetype, til forvalgte forrretningsområder '***
                                thisKtype_segment = 0
                                if func = "opret" AND step = 2 AND strKundeId <> "" then 


                                    strSQLktyp = "SELECT ktype FROM kunder WHERE kid = " & strKundeId
                                    oRec5.open strSQLktyp, oConn, 3
                                    if not oRec5.EOF then

                                    thisKtype_segment = oRec5("ktype")

                                    end if
                                    oRec5.close


                                end if

                                
                                
                                'strSQLf = "SELECT id, navn FROM fomr WHERE id <> 0 ORDER BY navn"
                                    strSQLf = "SELECT f.navn AS fnavn, f.id, f.konto, kp.kontonr AS kkontonr, kp.navn AS kontonavn, fomr_segment FROM fomr AS f "_
                                    &" LEFT JOIN kontoplan AS kp ON (kp.id = f.konto) WHERE f.id <> 0 AND f.jobok = 1 ORDER BY f.navn"


                                    'if session("mid") = 1 then
                                    'response.write strSQLf
                                    'response.flush
                                    'end if

                                    %>
                                <select name="FM_fomr" id="FM_fomr" multiple="multiple" size="12" style="width:380px;">
                                <option value="0">Ingen valgt</option>
                                    
                                    <%
                                    fa = 0
                                    strchkbox = ""
                                    oRec.open strSQLf, oConn, 3
                                    while not oRec.EOF

                                        if func = "opret" AND step = 2 then '*** Opret (forvalgt)

                                        if instr(oRec("fomr_segment"), "#"& thisKtype_segment &"#") <> 0 then
                                        fSel = "SELECTED"
                                        else
                                        fSel = ""
                                        end if


                                        else '** Rediger Forretningsområder
                                    
                                        if instr(strFomr_rel, "#"&oRec("id")&"#") <> 0 then
                                        fSel = "SELECTED"
                                        else
                                        fSel = ""
                                        end if

                                        end if



                                    if oRec("konto") <> 0 then
                                    kontonrVal = " ("& left(oRec("kontonavn"), 10) &" "& oRec("kkontonr") &")"

                                    if cint(fomr_konto) = cint(oRec("id")) then
                                    fokontoCHK = "CHECKED"
                                    else
                                    fokontoCHK = ""
                                    end if

                                    strchkbox = strchkbox & "<input type='radio' class='FM_fomr_konto' id='FM_fomr_konto_"& oRec("id") &"' name='FM_fomr_konto' value="& oRec("id") &" "& fokontoCHK &"> " & left(oRec("fnavn"), 20) &" "& kontonrVal &"<br>"
                                    else
                                    kontonrVal = ""
                                    end if 
                                    
                                    %>
                                    <option value="<%=oRec("id")%>" <%=fSel %>><%=oRec("fnavn") &" "& kontonrVal %></option>
                                    <%
                                        fa = fa + 1

                                    
                                    oRec.movenext
                                    wend
                                    oRec.close
                                    %>
                                </select>
                                <br />
                                <br />
                                <b>Forvalgt konto på faktura / ERP system</b><br />
                                Vælg herunder blandt de forretningsområder der har tilknyttet en omsætningskonto, og hvor fakturaer på dette job skal posteres på denne konto:<br />  
                                <%=strchkbox %>

                                <br />
                                <br />
                                
                                <%if func <> "red" then

                                    select case lto
                                    case "hestia", "intranet - local"
                                    fomr_sync_CHK = "CHECKED"
                                    case else
                                    fomr_sync_CHK = ""
                                    end select

                                else
                                fomr_sync_CHK = ""
                                end if

                                %>

                                <!--

                                <input id="FM_fomr_syncakt" type="checkbox" value="1" name="FM_fomr_syncakt" <%=fomr_sync_CHK%> /> Tildel / Synkroniser disse forretningsområder på underliggende aktiviteter.<br />
                                Alle valgte forretningsområder bliver fordelt ligeligt på hver aktivitet.
                                -->

                                <!--
                                <span style="font-size:11px; color:#999999;"><br />
                                Forretningsområder <b>skal</b> fordeles ned på aktiviteter for at blive talt med i statitikken over timeforbrug på hvert forretningsområde. Der kan manuelt tildeles forretningsområder nede på hver enkelt aktivitet.
                                </span>
                                --><br /><br />&nbsp;
                                </td>
                                </tr></table>
                            </div>
                                    
                </td>
                </tr>
                <%else 
                
                fomrIDs = replace("0" & strFomr_rel, "#", ",")
                fomrIDs = replace(fomrIDs, ",,", ",")
                fomrIDs = left(fomrIDs, len(fomrIDs)-1)
                %>
                <input id="Hidden2" type="hidden" name="FM_fomr" value="<%=fomrIDs%>" />



                <%end if%>

                 <!-- Avanceret indstillinger 2 -->
				<tr bgcolor="#FFFFFF">
								<td colspan=4 style="padding:10px 0px 0px 10px;">
                                <a href="#" class=vmenu id="a_tild_ava">+ Avanceret indstillinger </a><br />
                                Tildel bla. prioitet, faktura-indstillinger, pre-konditioner, kundeadgang mm.

								<br /><br />

                    
                      <div id="div_tild_ava" style="visibility:hidden; display:none; background-color:#FFFFFF;">
                     <table cellpadding=0 cellspacing=0 border=0 width=100%>
                    


                     <tr bgcolor="#FFFFFF">
				        <td>&nbsp;</td>
                        <td valign=top style="padding-top:30px;" colspan=3>
                        <h3>Prioitet:</h3></td>
                        <td style="">&nbsp;</td>
                    </tr>
                    <tr>
                        <td style="">&nbsp;</td>
						<td valign=top colspan=3>		
									
									
								
									
											<%strSQLr = "SELECT risiko FROM job WHERE risiko > -1 AND jobstatus = 1 ORDER BY risiko DESC"
									 
									 highestRiskval = 1
									 oRec4.open strSQLr, oConn, 3
									 if not oRec4.EOF then
									 highestRiskval = oRec4("risiko")
									 end if
									 oRec4.close
									 
									 %>
									 
									 <%strSQLr = "SELECT risiko FROM job WHERE risiko > -1 AND jobstatus  = 1 ORDER BY risiko"
									 
									 lowestRiskval = 0
									 oRec4.open strSQLr, oConn, 3
									 if not oRec4.EOF then
									 lowestRiskval = oRec4("risiko")
									 end if
									 oRec4.close
									 
									 %>

                                     <input id="prio" name="prio" type="text" value="<%=intprio %>" style="width:40px;" />
                                     Prioiteter på nuværende aktive job ligger mellem: <b><%=lowestRiskval%> - <%=highestRiskval%></b><br /><br />
									 <b>-1 = Internt job</b> vises ikke under fakturering og igangværende job.<br />
                                     <b>-2 = HR job</b> vises i HR mode på timereg. siden<br />
                                     <b>-3 = Internt job</b> men der skal kunne laves ressouceforecast på dette job. 
									 <br /><br />
                                     -1 / -2 / -3 medfører enkel visning af aktivitetslinjer på timereg. siden, dvs. der bliver ikke vist tidsforbrug, start- og slut -datoer mv. på aktiviteterne. <br /><br />&nbsp;
									
								
							</td>
							<td style="">&nbsp;</td>
						</tr>


                        
                        <%
                        preconditions_met_SEL0 = "SELECTED"
                        preconditions_met_SEL1 = ""
                        preconditions_met_SEL2 = ""

                        select case cint(preconditions_met)
                        case 0
                        preconditions_met_SEL0 = "SELECTED"
                        case 1
                        preconditions_met_SEL1 = "SELECTED"
                        case 2
                        preconditions_met_SEL2 = "SELECTED"
                        case else
                        preconditions_met_SEL0 = "SELECTED"
                        end select
                        %>


                           <tr bgcolor="#FFFFFF">
				        <td>&nbsp;</td>
                        <td valign=top style="padding-top:30px;" colspan=3>
                        <h3>Pre-konditioner opfyldt:<br /><span style="font-size:11px; font-weight:normal;">
                        Underleverandør klar, materialer indkøbt mm.</span></h3></td>
                        <td style="">&nbsp;</td>
                    </tr>
                    <tr>
                        <td style="">&nbsp;</td>
						<td valign=top colspan=3>	
                        
                        <select name="FM_preconditions_met" id="Select1" size="1" style="width:200px;">
                        <option value="0" <%=preconditions_met_SEL0 %>>Ikke angivet</option>
                        <option value="1" <%=preconditions_met_SEL1 %>>Ja</option>
                        <option value="2" <%=preconditions_met_SEL2 %>>Nej - afvent</option>
                        </select>
									
					</td>
							<td style="">&nbsp;</td>
						</tr>

                        <%
					
					
					
					'** Altid åben, indstilling vælges på timereg. siden **'
					'cint(lukafm) = 1
                    if cint(lukafm) = 1000 then%>
					<tr bgcolor="#FFFFFF">
						<td style="">&nbsp;</td>
						<td style="padding-left:5px; padding-top:30px; padding-bottom:4px;" colspan=3><h3>Mulighed for at lukke job fra timereg. siden?</h3>	
						<select name="FM_lukafmjob" style="font-size:9px; width:300px;">
									<%

                                    lukafmjob0CHK = "" 
									lukafmjob2CHK = ""
									lukafmjob3CHK = ""
                                    lukafmjob1CHK = ""

									select case lukafmjob
									case 1
									lukafmjob1CHK = "SELECTED"
									case 2
									lukafmjob2CHK = "SELECTED"
									case 3
									lukafmjob3CHK = "SELECTED"
									case 0
									lukafmjob0CHK = "SELECTED" 
									end select
									%>
									
									<option value="1" <%=lukafmjob1CHK %>>A - Luk job og send mail til jobansvarlig</option>
									<option value="2" <%=lukafmjob2CHK %>>B - Luk job, opret nyt (kopi) og send mail til jobansvarlig</option>
                                    <option value="3" <%=lukafmjob1CHK %>>C - Send mail til jobansvarlig, når en medarbejder er færdig med sin del</option>
									<option value="0" <%=lukafmjob0CHK %>>Ikke mulighed for at afmelde og lukke job fra timreg. siden.</option>
								
									</select>
									
							</td>
							<td style="">&nbsp;</td>
						</tr>
					    <%else %>
					    <input type="hidden" name="FM_lukafmjob" value="<%=lukafmjob %>">
					    <%end if %>
					    
					
					<%else%>
					<input type="hidden" name="FM_fastpris" value="0">
					<input type="hidden" name="FM_status" value="1">
					<%end if%>	
                
                  
                 


                   <tr>
						<td>&nbsp;</td>
						<td colspan=3 style="padding:30px 5px 10px 0px;"><h4>Fakturaindstillinger:<br /><span style="font-size:11px; font-weight:lighter;">(Nedarves fra kunde ved joboprettelse)</span></h4>
					 
					  
							<b>Valuta:</b><br /><select name="FM_valuta" style="width:380px;">
							<%strSQL = "SELECT id, valuta, valutakode, kurs FROM valutaer WHERE id <> 0 ORDER BY id " 
							oRec.open strSQL, oConn, 3
							while not oRec.EOF 
							 if oRec("id") = cint(valuta) then
							 vSEL = "SELECTED"
							 else
							 vSEL = ""
							 end if%>
							<option value="<%=oRec("id") %>" <%=vSEL %>><%=oRec("valuta") %> | <%=oRec("valutakode") %> | kurs: <%=oRec("kurs")%></option>
							<%
							oRec.movenext
							wend
							oRec.close%>
							
							</select>

                            <br /><br />
                             <b>Moms:</b> <br />

                           

                        <%strSQLmoms = "SELECT id, moms FROM fak_moms WHERE id <> 0 ORDER BY id " %>
                        <select name="FM_jfak_moms">

                            <%oRec6.open strSQLmoms, oConn, 3
                            while not oRec6.EOF 

                                if cint(jfak_moms) = cint(oRec6("id")) then
                                fakmomsSeL = "SELECTED"
                                else
                                fakmomsSeL = ""
                                end if

                            %><option value="<%=oRec6("id") %>" <%=fakmomsSeL %>><%=oRec6("moms") %>%</option><%
                  
                            oRec6.movenext
                            wend 
                            oRec6.close%>

                        </select>
        
                <br /><br />
        
                <b>Sprog:</b> <br />
                      <%strSQLsprog = "SELECT id, navn FROM fak_sprog WHERE id <> 0 ORDER BY id " %>
                    <select name="FM_jfak_sprog">

                        <%oRec6.open strSQLsprog, oConn, 3
                        while not oRec6.EOF 

                            if cint(jfak_sprog) = cint(oRec6("id")) then
                            faksprogSeL = "SELECTED"
                            else
                            faksprogSeL = ""
                            end if

                        %><option value="<%=oRec6("id") %>" <%=faksprogSeL %>><%=oRec6("navn") %></option><%
                  
                        oRec6.movenext
                        wend 
                        oRec6.close%>

                    </select>

							
						</td><td style="">&nbsp;</td></tr>
                
                
                
                <%if func = "opret" AND step = 2 OR func = "red" then %>
  
	          
        
                    
	              
                    <tr>
		            <td style="padding:30px 10px 10px 0px;" colspan=4>
                    <h3>Skal job være åben for kunde?</h3>
                    <input type="checkbox" name="FM_kundese" id="FM_kundese" value="1" <%=kundechk%>>&nbsp;<b>Gør job tilgængeligt for kontakt.</b><br>
		            Hvis tilgængelig for kontakt tilvælges, udsendes der en mail til kontakt-stamdata emailadressen, med link til kontakt loginside.
		
		
		            <%if func = "opret" then
		            hvchk1 = "checked"
		            hvchk2 = ""
		            else
			            if cint(intkundeok) = 2 then
			            hvchk1 = ""
			            hvchk2 = "checked"
			            else
			            hvchk1 = "checked"
			            hvchk2 = ""
			            end if
		            end if%>
		            <br><br>
		            <b>Hvis job åbnes for kontakt, hvornår skal registrerede timer så være tilgængelige?</b><br>
		            <input type="radio" name="FM_kundese_hv" value="0" <%=hvchk1%>>Offentliggør timer, så snart de er indtastet.<br>
		            <input type="radio" name="FM_kundese_hv" value="1" <%=hvchk2%>>Offentliggør først timer når jobbet er lukket. (afsluttet/godkendt)
		            <br>&nbsp;
                    </td></tr>
                  <%end if%>
                
                
                
                
                
                   </table>
                
                </div>
                <br />&nbsp;

                </td>
                </tr>

                <%




				
				

end if '** Rettigheder	
end select '*** Step %>


	
	
<%if editok = 1 then '*** rediger rettigheder ok%>	
	
	
	
	
	<%if func = "red" then
	submtxt = "Opdater & Afslut"
	else
        if step = "2" then
	    submtxt = "Opret & Afslut"
        else
        submtxt = "Opret"
        end if
	end if %>
	
	<tr bgcolor="#ffffff">
		<td colspan=4 align="right" style="padding:40px 40px 10px 10px;">
		<input id="Submit2" type="submit" value=" <%=submtxt %>  >> " <%=sendmailtokunde%> /></td>
	</tr>
	
	
	</table>
   
    </td>
	</tr>
	</table>
	</div><!-- tablediv -->
	
	

	<br><br><br>&nbsp;
	
	                     
	                    
	            <%
                '******************************************************************************
                '*** div projekt beregner ****'
                '******************************************************************************** 
                
                'call browsertype()

                'if browstype_client <> "ch" then

                
                call projektberegner

               

                'else
                
                'uWdt = 300
                'uLeft = 700
                'uTop = 100
                'uTxt = "Projektberegneren er pt. ikke understøttet af Chrome. Vi arbejder på en løsning, men indtil da henvises du til at benytte Internet Explorer, Firefox eller Safari."
                'call infoUnisportAB(uWdt, uTxt, uTop, uLeft)

                'end if    
                   
                        
						
	
	
	
	
	
	'*** Funktioner og Mini joboverblik ***'
	'if request("int") <> 0 AND showaspopup <> "y" then
	mini = 1
	if func = "red" AND showaspopup <> "y" AND mini = 1 then


     if len(trim(request("visrealtimerdetal"))) <> 0 AND request("visrealtimerdetal") <> 0 then
     visrealtimerdetal = 1
     visrealtimerdetalCHK = "CHECKED"
    else
     visrealtimerdetal = 0
     visrealtimerdetalCHK = ""
    end if


    call minioverblik

       
        
    
        '********************************************************'
        '**************** Milepæle ******************************'
        '********************************************************'

    oTop = 0
	oLeft = 487
	oWdth = 700
	oHgt = "" 
	oId = "milep_div"
	oZindex = 1000
    oVzb = "hidden"
    oDsp = "none"

	call tableDivAbs(oTop,oLeft,oWdth,oHgt,oId, oVzb, oDsp, oZindex)
    %>

    <table cellspacing="0" cellpadding="3" border="0" bgcolor="#EFF3FF" width="100%">
	<tr>
	<tr bgcolor="#5582D2">
		<td width="8" valign=top><img src="../ill/blank.gif" width="8" height="30" alt="" border="0"></td>
		<td colspan=4 style="padding-top:3px;"><h3 class="hv">Milepæle & Betalings terminer</h3></td>
		<td align=right valign=top><img src="../ill/blank.gif" width="8" height="30" alt="" border="0"></td>
	</tr>




	
	    

        <%for m = 0 to 1 %>
        

        <%if m = 0 then %>
        <tr bgcolor="#FFFFFF">
		<td valign="top" height=10>&nbsp;</td>
		<td colspan=3><br /><h4>Milepæle</h4></td>
        <td align=right>
        	<a href="javascript:popUp('milepale.asp?menu=job&func=opr&jid=<%=id%>&rdir=redjobcontionue&type=0','650','500','250','120');" target="_self"><span style="color:green; font-size:16px;"><b>+</b></span> Opret ny milepæl</a>
        </td>
		<td valign="top" align=right>&nbsp;</td>
	    </tr>
		<tr>
			<td valign="top" height=25 style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;">&nbsp;</td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;"">
			<b><%=formatdatetime(strTdato, 1)%></b></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;"">Job startdato</td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;"" colspan=2>&nbsp;</td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;"" align=right>&nbsp;</td>
		</tr>
        <%else %>
         <tr bgcolor="#FFFFFF">
		<td valign="top" height=10>&nbsp;</td>
		<td colspan=3><br /><h4>Terminer</h4></td>
        <td align=right>
        	<a href="javascript:popUp('milepale.asp?menu=job&func=opr&jid=<%=id%>&rdir=redjobcontionue&type=1','650','500','250','120');" target="_self"><span style="color:green; font-size:16px;"><b>+</b></span> Opret ny termin</a>
        </td>
		<td valign="top" align=right>&nbsp;</td>
	    </tr>
        <%end if %>
		
		<%
        if m = 0 then
        mtp_typ = " AND mpt_fak <> 1"
        else
        mtp_typ = " AND mpt_fak = 1"
        end if

		strSQL = "SELECT milepale.id AS id, milepale.navn AS navn, milepal_dato, "_
		&" milepale_typer.navn AS type, ikon, beskrivelse, milepale.editor, belob FROM milepale "_
		&" LEFT JOIN milepale_typer ON (milepale_typer.id = milepale.type) "_
		&" WHERE jid = "& id &" "& mtp_typ &" ORDER BY milepal_dato"
		x = 0
		oRec.open strSQL, oConn, 3
		while not oRec.EOF 
		
		select case right(x, 1)
		case 0,2,4,6,8
		bgthis = "#FFFFFF"
		case else
		bgthis = "#EFF3FF"
		end select
		%>
		<tr bgcolor="<%=bgthis %>">
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;">&nbsp;</td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;">
            <%if isNull(oRec("milepal_dato")) = false then %>
            <b><%=formatdatetime(oRec("milepal_dato"), 1)%></b>
            <%else %>
            - Ikke angivet
            <%end if %>
            <br />
			<span style="color:#999999;"><i><%=oRec("editor") %></i></span></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;">
			<span style="color:#999999;"><%=oRec("type")%></span>
            <%if m = 1 then %><br />
            <b><%=formatnumber(oRec("belob")) &" "& basisValISO%> </b>
            <%end if %></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;">
			<a href="javascript:popUp('milepale.asp?func=red&id=<%=oRec("id")%>&jid=<%=id%>&rdir=redjobcontionue','650','500','250','120');" target="_self" class=vmenu><img src="../ill/ac0038-16.gif" alt="Rediger" border="0"><%=oRec("navn")%></a>
			<%if len(oRec("beskrivelse")) <> 0 then%>
			<br><i><%=oRec("beskrivelse")%></i>
			<%end if%></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;"><a href="javascript:popUp('milepale.asp?milepale.asp?menu=job&id=<%=oRec("id")%>&jid=<%=id%>&func=slet&rdir=redjobcontionue','650','500','250','120');" target="_self"><img src="../ill/delete2.gif" alt="Slet" border="0"></a></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;" align="right">&nbsp;</td>
		</tr>
		
		<%
		x = x + 1
		oRec.movenext
		wend
		oRec.close

        select case right(x, 1)
		case 0,2,4,6,8
		bgthis = "#FFFFFF"
		case else
		bgthis = "#EFF3FF"
		end select
		%>
		

         <%if m = 0 then %>
		<tr bgcolor="<%=bgthis %>">
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;"><img src="../ill/blank.gif" width="1" height="25" alt="" border="0"></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;" height=25><b><%=formatdatetime(strUdato, 1)%></b></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;">Job slutdato</td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;" colspan=2>&nbsp;</td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;" align=right><img src="../ill/blank.gif" width="1" height="25" alt="" border="0"></td>
		</tr>
        <%end if %>
		

        <%next %>

	</table>
    </div> 
		
		
		<%end if
		'************************* Joboverblik / funktioner ***'
		%>
	
	
	
    <!--
	
	
	</div>side div

    -->
	
	<%
				
	
	
	
	
	
	
				'**********************************************************
				'**** Nedenstående (Timepriser) vises kun ved rediger *****
                '**** Og ved klik på timepriser ***'

                if antalmedarb < 25 then
                tpKriShow = 1
                else
                    if showdiv = "tpriser" then
                    tpKriShow = 1
                    else
                    tpKriShow = 0
                    end if
                end if

				if step = "2" AND func = "red" AND  tpKriShow = 1 then 'AND strInternt <> 0 '2

                call timepriser%>
                
                <%    
				end if%>	

             
	
	
	<%else%>
	<tr>
		<td colspan=4 bgcolor="#ffffff" height="100" style="border:1px #003399 solid; padding:20px;">
		<img src="../ill/alert_lille.gif" width="22" height="19" alt="" border="0">&nbsp;<font class=error><b>Fejl!</b></font><br>
		Du har ikke <b>rettigheder</b> til at redigere dette job. Du skal enten være jobansvarlig, 
		eller have administrator rettigheder for at redigere dette job.<br><br>
		Hvis alle medarbejdere skal kunne redigere et job <b>må der ikke</b> være valgt nogen jobansvarlige.
		<br>
		<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
		&nbsp;</td>
	</tr>
	</table>
	</div><!-- tablediv -->
	
	
	<%end if '** rettigheder%>
	
	
	</form>
	
	<br><br><br>&nbsp;
	</div><!--- side div -->

	
	
	<%if (lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "xintranet - local" OR lto = "epi_ab" OR lto = "epi_cati" OR lto = "epi_uk") AND func <> "red" AND step = 2 then 

    %>
    <div id="load" style="position:absolute; display:; visibility:visible; top:200px; left:135px; width:1000px; height:600px; background-color:#ffffff; border:10px #9ACD32 solid; padding:10px; z-index:100000000;">
    <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" /><br /><br />
	TimeOut forbereder dit nye projekt, vent veligst. <br />Der kan gå op til 20 sek...

        <br /><br /><br />

           <div id="auopr_5" style="color:#5582d2; font-size:14px; border:0px #999999 solid; width:980px; padding:1px;"></div><br />

    <div id="auopr_1">Indlæser aktiviteter..</div><br />
    <div id="auopr_2"></div><br />
    <div id="auopr_3"></div><br />
    <div id="auopr_4"></div><br />


      

	</td></tr></table>
	
	</div>
    <%
       
    'response.end
   
    Response.Write("<script>autoopret();</script>") 
    end if %>


	<%case else
        
        
        
    'response.write "|"& request("FM_fomr") &"|"
    'response.write "strFomr_rel: "& strFomr_rel
        %>
	
	
	
	
	
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <script src="inc/job_listen_jav.js"></script>
    
 <div id="loadbar" style="position:absolute; display:; visibility:visible; top:300px; left:280px; width:300px; background-color:#ffffff; border:10px #9ACD32 solid; padding:10px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	<br />
	Forventet loadtid: 5-45 sek. <br />(ved mere end 1000 job)
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	
	</td></tr></table>

	</div>


	

    <%call menu_2014() %>
	
	
	

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px;">
	
	
	<%


	
	
	
	
	Public left_intX
	Public function kommaFunc(x)
	if len(x) <> 0 then
	instr_komma = instr(x, ",")
		
		if instr_komma > 0 then
		left_intX = left(x, instr_komma + 2)
		else
		left_intX = x
		end if
	else
	left_intX = 0
	end if
	
	Response.write left_intX 
	end function
	
    if len(trim(request("filt"))) <> 0 then
	filt = request("filt")
    else
        if request.cookies("tsa")("statusfilt") <> "" then
        filt = request.cookies("tsa")("statusfilt")
        else
        filt = "1"
        end if
    end if
	
    response.cookies("tsa")("statusfilt") = filt
    
    aftfilt = request("aftfilt")
	sort = Request("sort")
	usedletter = request("l") 
	

   
	
	oimg = "ikon_job_48.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = "Joboversigt"
	
    if media = "print" OR request("print") = "j" then
	    call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
    end if
	

    'Response.write "filt" & filt
    'Response.end
    
    chk1 = ""
	chk2 = ""
	chk3 = ""
	chk4 = ""
	chk5 = ""

    varFilt = " AND (jobstatus = -2"
    for f = 0 to 4
        'Response.write instr(filt, f) & "<br>.."

        if instr(filt, f) <> 0 then
        varFilt = varFilt & " OR jobstatus = "& f
            
            select case f
            case 0
            chk0 = "CHECKED"
            case 1
            chk1 = "CHECKED"
            case 2
            chk2 = "CHECKED"
            case 3
            chk3 = "CHECKED"
            case 4
            chk4 = "CHECKED"
            end select
        
        
        end if



    next

    varFilt = varFilt & ")"

    'Response.write "varFilt" & varFilt


	

	
	select case aftfilt
	case "serviceaft"
	varFilt = varFilt & " AND serviceaft <> 0" 'Service aftaler.
	chk6 = "CHECKED"
	chk7 = ""
	chk8 = ""
	case "ikkeserviceaft"
	varFilt = varFilt & " AND serviceaft = 0" 'Ikke Service aftaler.
	chk6 = ""
	chk7 = "CHECKED"
	chk8 = ""
	case else
	varFilt = varFilt & "" 'Vis alle.
	chk6 = ""
	chk7 = ""
	chk8 = "CHECKED"
	end select
	
	
	'* er der valgt et bogstav ****
	if len(usedletter) <> 0 then
	varFilt = varFilt & " AND jobnavn LIKE '"&request("l")&"%'" 
	'showFilter = showFilter & " der starter med:<b> " & request("l") &"</b>"
	else
	varFilt = varFilt 
	'showFilter = showFilter
	end if
	
	
		if len(request("FM_kunde")) <> 0 then
		fmkunde = request("FM_kunde")
		else
		fmkunde = 0
		end if
	    
        if len(request("frompost")) <> 0 then
            if len(trim(request("FM_vis_timepriser"))) <> 0 then
            vis_timepriser = 1
            else
            vis_timepriser = 0
            end if

           
        else
            if request.Cookies("job")("vis_timepriser") <> "" then
            vis_timepriser = request.Cookies("job")("vis_timepriser")
            else
            vis_timepriser = 0
            end if
        end if

        Response.Cookies("job")("vis_timepriser") = vis_timepriser

        if vis_timepriser = 1 then
        vis_timepriserCHK = "CHECKED"
        else
        vis_timepriserCHK = ""
        end if

	
	'''** søg i akt ***'
	if len(trim(request("FM_sogakt"))) <> 0 then
			sogakt = 1
			sogaktCHK = "CHECKED"
	else
	        sogakt = 0
	        sogaktCHK = ""		
    end if	
	
	    
        
        
        '*************************************************
        '*** Er der søgt på en specifikt jobnr/jobnavn ***
		'*************************************************
        
        if len(request("FM_kunde")) <> 0 then 'er der søgt/submitted??
        jobnr_sog = request("jobnr_sog")
        response.cookies("tsa")("jobsog") = jobnr_sog
        visliste = 1
        else

            if  request.cookies("tsa")("jobsog") <> "" then
            jobnr_sog = request.cookies("tsa")("jobsog")
            else
            jobnr_sog = ""
            end if
        visliste = 0
        end if




		if len(jobnr_sog) <> 0 AND jobnr_sog <> "-1" OR (cint(visliste) = 1) then
			jobnr_sog = jobnr_sog
			
			if cint(sogakt) = 1 then
			sogeKri = sogeKri & " (a.navn LIKE '%"& jobnr_sog &"%' "
			else
			    

                if instr(jobnr_sog, ">") > 0 OR instr(jobnr_sog, "<") > 0 OR instr(jobnr_sog, "-") > 0 then
           
                if instr(jobnr_sog, ">") > 0 then
                sogeKri = sogeKri &" (j.jobnr > "& replace(trim(jobnr_sog), ">", "") &" "
                end if

                if instr(jobnr_sog, "<") > 0 then
                sogeKri = sogeKri &" (j.jobnr < '"& replace(trim(jobnr_sog), "<", "") &"' "
                end if

                if instr(jobnr_sog, "-") > 0 then
                jobnr_sogArr = split(jobnr_sog, "-")
               
                for t = 0 to 1
                
                if t = 0 then
                jobSogKriA = jobnr_sogArr(0)
                else
                jobSogKriB = jobnr_sogArr(1)
                end if
                
                next

                sogeKri = sogeKri &" (j.jobnr BETWEEN '"& trim(jobSogKriA) &"' AND '"& trim(jobSogKriB) &"'"
                end if

                else

                sogeKri = " (j.jobnr LIKE '"& jobnr_sog &"' OR j.jobnavn LIKE '"& jobnr_sog &"%' OR j.id LIKE '"& jobnr_sog &"' OR Kkundenavn LIKE '"& jobnr_sog &"%' OR Kkundenr LIKE '"& jobnr_sog &"'"
			

                end if
            
            end if
			
			sogeKri = sogeKri & ") AND "
			
			
			show_jobnr_sog = jobnr_sog
			
		else
			jobnr_sog = "-1"
			sogeKri = " (j.id = -1) AND "
			show_jobnr_sog = ""
		end if


		
    '**** Projektgrupper ****'
    if len(trim(request("FM_prjgrp"))) <> 0 then
    prjgrp = request("FM_prjgrp")
    response.Cookies("job")("prjgrp") = prjgrp
    else
        if request.Cookies("job")("prjgrp") <> "" then
        prjgrp = request.Cookies("job")("prjgrp")
        else
        prjgrp = 10 '** Alle
        end if
    end if
		
	
	call datocookie
	
	'**** Brug datokriterie filer ****
	if len(request("usedatokri")) <> 0 then
		if request("usedatokri") = "j" then
		usedatoKri = "j"
		datoKriJa = "CHECKED"
		datoKriNej = ""
		else
		usedatoKri = "n"
		datoKriJa = ""
		datoKriNej = "CHECKED"
		end if
	else
		usedatoKri = request.Cookies("job")("cusedatokri")
		if usedatoKri = "j" then
		datoKriJa = "CHECKED"
		datoKriNej = ""
		else
		usedatoKri = "n"
		datoKriJa = ""
		datoKriNej = "CHECKED"
		end if
	end if
	


    if len(trim(request("FM_hd_kpers"))) <> 0 then
    hd_kpers = request("FM_hd_kpers")
    else
    hd_kpers = -1
    end if

	'*** Indsætter cookie ***
	Response.Cookies("job")("cusedatokri") = usedatoKri
	Response.Cookies("job").Expires = date + 65
	
	if usedatoKri = "j" then
	datoKri = " AND jobslutdato BETWEEN '"& strStartDato &"' AND '"& strSlutDato &"'"
	else
	datoKri = ""
	end if	
	'***** 
	
	%>
	
	<%call filterheader_2013(0,0,900,oskrift)%>
	<table cellspacing=0 width=100% cellpadding=0 border=0>
    <tr><td valign=top>

    <table border=0 cellspacing=0 cellpadding=0>
    <form action="jobs.asp?menu=job&shokselector=1&frompost=1" method="post" name="jobnr" id="jobnr">
	<tr>
		<td colspan=2><b>Kunde:</b>&nbsp;&nbsp;
		<select name="FM_kunde" id="FM_kunde" style="width:380px;" onchange="submit();">
		<option value="0">Alle <%=writethis%></option>
		<%
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'e' ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(fmkunde) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%>&nbsp;(<%=oRec("Kkundenr")%>)</option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
		</select>
		</td>
	</tr>


    <tr>
    <input type="hidden" value="<%=hd_kpers %>" name="FM_hd_kpers" id="FM_hd_kpers" />
		<td colspan=2 style="padding:5px 2px 2px 0px;"><b>Kontaktperson:</b>&nbsp;&nbsp;
		<select name="FM_kundekpers" id="FM_kundekpers" style="width:342px;">
		<option value="0">Alle (vælg kontakt for at se kontaktpersoner)</option>
		</select>
        </td>
	</tr>




	<tr><td colspan=2>
	<%
	if len(request("FM_medarb_jobans")) <> 0 then
	medarb_jobans = request("FM_medarb_jobans")
	response.cookies("tsa")("jobans") = medarb_jobans
	else
	    if request.cookies("tsa")("jobans") <> "" then
	    medarb_jobans = request.cookies("tsa")("jobans")
	    else
	    medarb_jobans = 0
	    end if
	end if
	
	response.cookies("tsa").expires = date + 45
	%>
	
    <br /><br />
	<b>Jobansvarlig:</b> (1 ell. 2)&nbsp;&nbsp; 
	<select name="FM_medarb_jobans" id="FM_medarb_jobans" style="width:310px;">
            <option value="0">Alle (ignorer)</option>
            <%
            mNavn = "Alle"
            
             strSQL = "SELECT mnavn, mnr, mid, init FROM medarbejdere WHERE mansat <> '2' ORDER BY mnavn"
             oRec.open strSQL, oConn, 3
             while not oRec.EOF
                
                if cint(medarb_jobans) = oRec("mid") then
                selThis = "SELECTED"
                else
                selThis = ""
                end if
                
                
                %>
             <option value="<%=oRec("mid")%>" <%=selThis%>><%=oRec("mnavn") %> (<%=oRec("mnr") %>)
             <%if len(trim(oRec("init"))) <> 0 then %>
              - <%=oRec("init") %>
              <%end if %>
              </option>
             <%
             oRec.movenext
             wend
             oRec.close
             %>
             
              
        </select>
	</td></tr>
    <tr><td colspan=2><br />
    <b>Projektgrupper</b>: (vis kun job hvor valgte projektgruppe er tilknyttet..)<br />
    <select  name="FM_prjgrp" id="FM_prjgrp" style="width:433px;">
    <%call progrupperIdNavn(prjgrp) %>
    
    </select>
    
    </td></tr>
   

	<tr bgcolor="#FFFFFF">
		<td colspan=2><br /><b>Søg på jobnr, jobnavn ell. kunde:</b><br />
		<input type="text" name="jobnr_sog" id="jobnr_sog" value="<%=show_jobnr_sog%>" style="width:433px; border:2px #6CAE1C solid;">&nbsp;
		<br />
        (% = wildcard) Brug ">", "<" ell. "-" til at søge efter jobnr i et interval.<br />
            <input id="FM_sogakt" name="FM_sogakt" type="checkbox" value="1" <%=sogaktCHK%> /> Vis kun job hvor søgekriterie indgår i en aktivitet på jobbet. 

        <br /><br /><br />
        <div style="border:1px #CCCCCC solid; padding:5px;">
        <input type="checkbox" name="FM_vis_timepriser" <%=vis_timepriserCHK %> value="1" />
        <b>Vis som liste med medarbejder-timepriser</b> (maks. 10000 linier)
        <br />
        Sorter efter: <input type="radio" value="0" name="FM_sorter_tp" <%=sorttpCHK0 %> /> Aktiviteter, ell.
        &nbsp;<input type="radio" value="1" name="FM_sorter_tp" <%=sorttpCHK1 %> />  Medarbejdere
        </div>
        <br /><br />&nbsp;
		</td>
	</tr>
    </table>
    </td>
   
    <td valign=top>
    &nbsp;
     </td>
   
    <td valign=top>

    <table border=0 cellspacing=0 cellpadding=0>

     <tr><td colspan=2>
         <b>Forretningsområder:</b> <!--<%=strFomr_rel %> <%=left(strFomr_Gblnavn, 75) %>--><br />

                              
                                <%
                                
                                
                                strSQLf = "SELECT id, navn FROM fomr WHERE id <> 0 ORDER BY navn"%>
                                <select name="FM_fomr" id="Select2" multiple="multiple" size="6" style="width:410px;">
                                <%if instr(strFomr_rel, "#0#") <> 0 then 
                                f0sel = "SELECTED"
                                else
                                f0sel = ""
                                end if%>

                                <option value="#0#" <%=f0sel %>>Alle (ignorer)</option>
                                    
                                    <%oRec.open strSQLf, oConn, 3
                                    while not oRec.EOF
                                    
                                    if instr(strFomr_rel, "#"&oRec("id")&"#") <> 0 then
                                    fSel = "SELECTED"
                                    else
                                    fSel = ""
                                    end if
                                    
                                    %>
                                    <option value="#<%=oRec("id")%>#" <%=fSel %>><%=oRec("navn") %></option>
                                    <%
                                    oRec.movenext
                                    wend
                                    oRec.close
                                    %>
                                </select>

    </td></tr>

	<tr>
		<td colspan=2 valign=top style="padding-top:15px;"><b>Status filter:</b><br />
       
		
        <input type="CHECKBOX" name="filt" value="1" <%=chk1%>/> Aktive<br />
        <input type="CHECKBOX" name="filt" value="2" <%=chk2%>/> Passive / Til fakturering<br />
        <input type="CHECKBOX" name="filt" value="3" <%=chk3%>/> Tilbud<br />
        <input type="CHECKBOX" name="filt" value="4" <%=chk4%>/> Gennemsyn<br />
        <input type="CHECKBOX" name="filt" value="0" <%=chk0%>/> Lukkede<br />

		</td>
	</tr>
	<tr>
		<td colspan=2><br /><br /><b>Er job tilknyttet en aftale?</b><br />
		<input type="radio" name="aftfilt" value="visalle" <%=chk8%>>&nbsp;Vis alle&nbsp;&nbsp;
		<input type="radio" name="aftfilt" value="serviceaft" <%=chk6%>>&nbsp;Ja&nbsp;&nbsp;
		<input type="radio" name="aftfilt" value="ikkeserviceaft" <%=chk7%>>&nbsp;Nej&nbsp;&nbsp;
		</td>
	</tr>
   
	<tr>
		<td colspan=2><br /><br />

		<b>Vis kun job med slutdato i følgende periode:</b><br /> 
            Nej <input type="radio" name="usedatokri" value="n" <%=datoKriNej%>><br />
            Ja:
		<input type="radio" name="usedatokri" value="j" <%=datoKriJa%>> <span style="font-size:9px; color:#999999;">(Sorterer efter slutdato)</span> <br />
		<table width=100% cellpadding=0 cellspacing=0>
			<tr><!--#include file="inc/weekselector_b.asp"--></tr>
		</table>
		</td>
    </tr>
    <tr>
		<td colspan=2 align=right valign=bottom><br /><br /><br /><br /><input id="Submit1" type="submit" value="Vis job >>"/>
            &nbsp;</td>
	</tr>
	</form>
	</table>

    </td>
    </tr>
    </table>
	
	<!-- filter header -->
	</td></tr></table>
	</div>
	<br /><br />

	
	<%
	
	tTop = 0
	tLeft = 0
	tWdth = 1224
	
	
	call tableDiv(tTop,tLeft,tWdth)
	

    '** Vælger sql sætning efter sortering og filter ***
	
	if fmkunde <> 0 then
	varJobknrKri = "j.jobknr = "& Request("FM_Kunde") &" AND "
	else
		'if request("fromvemenu") = "j" then
		'varJobknrKri = "j.jobknr = 0 AND "
		'else
		varJobknrKri = ""
		'end if
	end if
	
	if usedatoKri = "j" then
	varOrder = "jobslutdato DESC"
	else
	varOrder = "kkundenavn, jobnavn"
	end if
	
    '** jobans **'
	if cint(medarb_jobans) = 0 then
	jobansKri = ""
	else
	jobansKri = " (jobans1 = " & medarb_jobans & " OR jobans2 = "& medarb_jobans &") AND "
	end if
	

    '** kontaktpersoner hos kunde **'
    if cint(hd_kpers) <> -1 AND cint(hd_kpers) <> 0  then
    kpersSQLkri = " AND kundekpers = " & hd_kpers
    else
    kpersSQLkri = ""
    end if


    if prjgrp <> 0 AND prjgrp <> 10 then
    prjgrpSQLkri = " AND (projektgruppe1 = "&prjgrp&" "_
    &" OR projektgruppe2 = "&prjgrp&" OR projektgruppe3 = "&prjgrp&" OR projektgruppe4 = "&prjgrp&" "_
    &" OR projektgruppe5 = "&prjgrp&" OR projektgruppe6 = "&prjgrp&" OR projektgruppe7 = "&prjgrp&" OR projektgruppe8 = "&prjgrp&" OR projektgruppe9 = "&prjgrp&" OR projektgruppe10 = "&prjgrp&")"
    else
    prjgrpSQLkri = ""
    end if 
	




    if cint(vis_timepriser) <> 1 then
	%>
	
	
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#ffffff" width=100%>
	<form action="jobs.asp?menu=job&func=opdjobliste&jobnr_sog=<%=jobnr_sog%>&filt=<%=filt%>&fm_kunde_sog=<%=fmkunde%>&FM_kunde=<%=fmkunde%>" method="post">
	<%sub tabletop%>
	
	<tr bgcolor="#5582D2">
		<td class='alt' valign=bottom style="width:200px;"><b>Job&nbsp;/&nbsp;Jobansv.<br />Kunde&nbsp;/&nbsp;Per.</b></td>
		<td class='alt' style="padding-right:10px;" valign=bottom><b>Aktiviteter</b></td>
        <td class='alt' style="padding-right:10px;" valign=bottom><b>Forretningsomr.</b></td>
        <td class='alt' valign=bottom style="width:150px; padding-left:20px;"><b>Realiseret %</b><br />(heraf fakturerbare)</td>
		<td class='alt' valign=bottom style="padding-right:5px; white-space:nowrap;" align=right><b>Brutto oms</b><br />(Budget)</td>
		<td align=right style="padding-right:5px;" valign=bottom class=alt style="width:100px;"><b>Timer forkalk.</b><br>
		Realiseret<br>
		<b>Balance</b></td>
		<td class='alt' valign=bottom style="padding-left:5px;"><b>Status</b>&nbsp;

   <table cellpadding=0 cellspacing=0 border=0 width="100">
       
   <tr><td style="font-size:9px; color:#CCCCCC;"><input type="radio" name="st_sel" value="1" id="st_cls_1" />Aktiv</td></tr>
   <tr><td style="font-size:9px; color:#CCCCCC;"><input type="radio" name="st_sel"  value="2" id="st_cls_2" />Passiv/Fak.</td></tr>  
   <tr><td style="font-size:9px; color:#CCCCCC;"><input type="radio" name="st_sel"  value="3" id="st_cls_3"/>Tilbud</td></tr>  
   <tr><td style="font-size:9px; color:#CCCCCC;"><input type="radio" name="st_sel" value="4" id="st_cls_4" />Gen.syn</td></tr> 
   <tr><td style="font-size:9px; color:#CCCCCC;"><input type="radio" name="st_sel" value="0" id="st_cls_0" />Lukket</td></tr> 
	
  </table>	
            
            </td>
		<td class='alt' valign=bottom align=center style="padding-left:5px; width:100px;"><b>Funktioner</b></td>
		<td class='alt' valign=bottom style="width:160px;"><b>Faktura hist.</b><br />
        Faktisk timepris<br />
        <span style="font-size:9px; line-height:10px; color:#CCCCCC;">Faktureret beløb <br />(ekskl. mat. og km.) <br />/ real. timer</span></td>
		<td class='alt' valign=bottom colspan=2><b>Projektgrupper</b>
		</td>
		
    </tr>
	<%
	end sub
	
	
	
	
	'***********************************************************
	'*** Main SQL **********************************************
	'***********************************************************
	strSQL = "SELECT j.id, jobnavn, jobnr, kkundenavn, kid, jobknr, jobTpris, jobstatus, jobstartdato, "_
	&" jobslutdato, j.budgettimer, fakturerbart, Kkundenr, ikkebudgettimer, jobans1, jobans2, jobans3, jobans4, jobans5, fastpris, jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, virksomheds_proc, "_
	&" s.navn AS aftnavn, rekvnr, tilbudsnr, sandsynlighed, jo_bruttooms, kundekpers, serviceaft, lukkedato, preconditions_met"
	
    strSQL = strSQL &", j.projektgruppe1, j.projektgruppe2, j.projektgruppe3, j.projektgruppe4, j.projektgruppe5, j.projektgruppe6, j.projektgruppe7, j.projektgruppe8, j.projektgruppe9, j.projektgruppe10 "
	   
	
	if cint(sogakt) = 1 then
	    strSQL = strSQL &" FROM aktiviteter AS a "
	    strSQL = strSQL &" LEFT JOIN job AS j ON (j.id = a.job)"
	    strSQL = strSQL &" LEFT JOIN kunder ON (kunder.kid = jobknr)"
	else
        strSQL = strSQL &" FROM job AS j "
        strSQL = strSQL &" LEFT JOIN kunder ON (kunder.kid = jobknr)"
    end if

	
	strSQL = strSQL &" LEFT JOIN serviceaft s ON (s.id = serviceaft)"_
	&" WHERE "& varJobknrKri &" "& sogeKri &" "& jobansKri &" kunder.Kid = jobknr " & varFilt &" "& datoKri &" "& prjgrpSQLkri &" "& kpersSQLkri &""_
	&" GROUP BY j.id ORDER BY "&varOrder&" LIMIT 5000"
	
	
	
	x = 0
	cnt = 0
	jids = 0
	totReal = 0
	gnsPrisTot = 0
	totRealialt = 0
    thisMid = session("mid")
	
	'Response.write strSQL
	'Response.Flush
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF

    call forretningsomrJobId(oRec("id"))


     if cint(visJobFomr) = 1 OR instr(strFomr_rel, "#0#") <> 0 then


	
	'** Til Export fil ***
	jids = jids & "," & oRec("id")
	
	if cnt = 0 then
	call tabletop
	end if
	

    if len(trim(oRec("budgettimer"))) = "0" OR oRec("budgettimer") = "0" then 
	budgettimer = 0
	else
	budgettimer = oRec("budgettimer")
	end if
	
	if oRec("ikkebudgettimer") > 0 then 
	ikkebudgettimer = oRec("ikkebudgettimer")
	else
	ikkebudgettimer = 0
	end if
	
	'** tildelt total ***
	budgettimertot = (ikkebudgettimer + budgettimer)


	call akttyper2009(2)

	'*** Real timer og Proafsluttet **************************
	strSQL = "SELECT sum(timer) AS proafslut FROM timer WHERE Tjobnr = '" & oRec("jobnr") &"' AND ("& aty_sql_realhours &") "
	oRec3.open strSQL, oConn, 3
	
	if len(oRec3("proafslut")) <> 0 then
	proaf = oRec3("proafslut")
	else
	proaf = 0
	end if

	oRec3.close

    totReal = proaf
    restt = (budgettimertot - proaf)
    
    realfakbare = 0
    '*** Real. fakturerbare timer **************************
	strSQL = "SELECT sum(timer) AS realfakbare FROM timer WHERE Tjobnr = '" & oRec("jobnr") &"' AND tfaktim = 1"
	oRec3.open strSQL, oConn, 3
	if not oRec3.EOF then

	if len(oRec3("realfakbare")) <> 0 then
	realfakbare = oRec3("realfakbare")
	else
	realfakbare = 0
	end if

	end if
	oRec3.close
	

    resttfakbare = (budgettimertot - realfakbare)

    'timerTotFakbarePajob = realfakbare 


    'if len(timerTotFakbarePajob) <> 0 then
	'timerTotFakbarePajob = timerTotFakbarePajob
	'else
	'timerTotFakbarePajob = 0
	'end if



		
		'*** bgcolor ***
		if cint(id) = oRec("id") then
		bc = "#FFFF99"
		else
			select case right(cnt, 1) 
			case 0,2,4,6,8
			bc = "#FFFFFF"
			case else
			bc = "#EFF3FF" '"#5582D2" '"#d2691e"
			end select
		end if
	
	
	
	sletkanp = "<img src='../ill/slet_16.gif' alt='Slet' border='0'>" 
	
	'if oRec("fakturerbart") = "1" then
		if oRec("jobstatus") = 3 then
		inteksimg = "<img src='../ill/filtertilbud.gif' width='10' height='21' alt='' border='0'>"
		else
		inteksimg = "<img src='../ill/blank.gif' width='15' height='15' alt='' border='0'>"
		'eksternt_job_trans
		end if
	'else
	'inteksimg = "<img src='../ill/internt_job_trans.gif' width='15' height='15' alt='' border='0'>"
	'end if
	
	


    '*** Oms. timer ****'
    jo_bruttooms = oRec("jo_bruttooms")
	jo_fastpris = oRec("fastpris")


    '** Prekonditioner ****'
    preconditions_met = oRec("preconditions_met")
    
	
	if budgettimer <> 0 then
	projektcomplt = ((proaf/budgettimertot)*100)
    projektcompltFakbare = ((realfakbare/budgettimertot)*100)
	else
	projektcompltFakbare = 100
    projektcomplt = 100
	end if
	
	if projektcomplt > 100 then
	showprojektcomplt = projektcomplt
	projektcomplt = 100
	bgdiv = "Crimson"
	else
	showprojektcomplt = projektcomplt
	projektcomplt = projektcomplt
	bgdiv = "yellowgreen"
	end if

    if projektcompltFakbare > 100 then
	showprojektcompltFakbare = projektcompltFakbare
	projektcompltFakbare = 100

	else
	showprojektcompltFakbare = projektcompltFakbare
	projektcompltFakbare = projektcompltFakbare
	
	end if

    
	
	'*** Antal aktiviteter på job ***
	strAktnavn = ""
    lastFase = ""
    strSQL2 = "SELECT id, navn, fase FROM aktiviteter WHERE job = "& oRec("id") & " ORDER BY fase, sortorder, navn"
	oRec5.open strSQL2, oConn, 3
	while not oRec5.EOF 
	x = x + 1

    if lastFase <> oRec5("fase") AND isNull(oRec5("fase")) <> true AND len(trim(oRec5("fase"))) <> 0 then

    strAktnavn = strAktnavn & "<br><b>"& replace(oRec5("fase"), "_", " ") & "</b><br>"
    lastFase = oRec5("fase")
    end if

    strAktnavn = strAktnavn & left(oRec5("navn"), 20) & "<br>"
	oRec5.movenext
	wend
	oRec5.close
	Antal = x
	
	''** Antal brugte fakturerbare timer **
	'strSQL3 = "SELECT sum(timer) AS timerTotFaktimerpajob FROM timer WHERE Tjobnr= " & oRec("jobnr") &" AND ("& aty_sql_realhours &")"
	'oRec3.open strSQL3, oConn, 3
	'if not oRec3.EOF then
	'timerTotFakbarePajob = oRec3("timerTotFaktimerpajob")
	'end if

	'oRec3.close
	
	'if len(timerTotFakbarePajob) <> 0 then
	'timerTotFakbarePajob = timerTotFakbarePajob
	'else
	'timerTotFakbarePajob = 0
	'end if
	
	'*** fakturerbare timer tildelt på aktiviteter **** 
	'strSQL3 = "SELECT sum(budgettimer) AS akttimer FROM aktiviteter WHERE job = " & oRec("id") &" AND fakturerbar = 1 ORDER BY budgettimer"
	'oRec3.open strSQL3, oConn, 3
	'if not oRec3.EOF then
	'	akttfaktimtildelt = oRec3("akttimer")
	'end if
	'oRec3.close
	
	'if len(akttfaktimtildelt) <> 0 then
	'akttfaktimtildelt = akttfaktimtildelt
	'else
	'akttfaktimtildelt = 0
	'end if
	
	
	'***Ikke fakturerbare timer tildelt på aktiviteter **** 
	'strSQL3 = "SELECT sum(budgettimer) AS aktnotftimer FROM aktiviteter WHERE job = " & oRec("id") &" AND fakturerbar = 0 ORDER BY budgettimer"
	'oRec3.open strSQL3, oConn, 3
	'if not oRec3.EOF then
	'	akttnotfaktimtildelt = oRec3("aktnotftimer")
	'end if
	'oRec3.close
	
	'if len(akttnotfaktimtildelt) <> 0 then
	'akttnotfaktimtildelt = akttnotfaktimtildelt
	'else
	'akttnotfaktimtildelt = 0
	'end if



	%>
	<tr>
		<td bgcolor="#d6dff5" colspan="11"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bc%>">
	
	<%
	'*** Tjekker rettigehder eller om man er jobanssvarlig ***
	editok = 0
	if level <= 2 OR level = 6 then
	editok = 1
	else
			if cint(session("mid")) = oRec("jobans1") OR cint(session("mid")) = oRec("jobans2") OR (cint(oRec("jobans1")) = 0 AND cint(oRec("jobans2")) = 0) then
			editok = 1
			else
			editok = 0   
			end if
	end if
	
	
	              
	
	
	'********* Jobnavn og nr ****
	%>
	<td valign=top style="padding:4px 10px 4px 2px; width:250px;">
      <b><%=oRec("kkundenavn")%></b>&nbsp;(<%=oRec("Kkundenr")%>)<br />
	<%
	'Response.write  cint(session("mid")) &"="& oRec("jobans1") &" OR " & cint(session("mid")) &" = " & oRec("jobans2") &" OR "& cint(oRec("jobans1")) &" = 0 OR" & cint(oRec("jobans2")) &" = 0 then"
			
	
	if editok = 1 then
		'if oRec("fakturerbart") = "1" then%>
		<a href="jobs.asp?menu=job&func=red&id=<%=oRec("id")%>&int=<%=oRec("fakturerbart")%>&jobnr_sog=<%=jobnr_sog%>&filt=<%=filt%>&fm_kunde_sog=<%=fmkunde%>"><%=oRec("jobnavn")%>&nbsp;&nbsp;(<%=oRec("jobnr")%>)</a>&nbsp;
		<%'else%>
		<!--<a href="jobs.asp?menu=job&func=red&id=<=oRec("id")%>&int=<=oRec("fakturerbart")%>&jobnr_sog=<=jobnr_sog%>&filt=<=filt%>&fm_kunde_sog=<=fmkunde%>" class=vmenu><=oRec("jobnavn")%>&nbsp;&nbsp;(<=oRec("jobnr")%>)</a>&nbsp;-->
		<%'end if
	else
		%>
		<b><%=oRec("jobnavn")%>&nbsp;&nbsp;(<%=oRec("jobnr")%>)</b>&nbsp;
		<%
	end if
	

    if oRec("fastpris") = 1 then
    strFasptpris = "Fastpris"
    else
    strFasptpris = "Lbn. timer"
    end if


    %>
    <span style="color:green; font-size:10px;">(<%=strFasptpris %>)</span>

    <%
	
	
	
	'******************************
	%>
    <span style="color:#999999; font-size:9px;"><i>
    <%
		
						'*** Jobansvarlige ***
                        if oRec("jobans1") <> 0 then
						call meStamData(oRec("jobans1"))
                        %>
						<%=meNavn%> 
                             <%if oRec("jobans_proc_1")  <> 0 then %>
                            (<%=oRec("jobans_proc_1") %>%)
                            <%end if %>

                        <%end if
                        

                        '*** Jobejer 2 ***
                        if oRec("jobans2") <> 0 then
						call meStamData(oRec("jobans2"))
                        %>
						, <%=meNavn%> 
                        
                        <%if oRec("jobans_proc_2")  <> 0 then %>
                        (<%=oRec("jobans_proc_2") %>%)
                        <%end if %>

                        <%end if

                        '*** Job medansvarlige ***
                        if oRec("jobans3") <> 0 then
						call meStamData(oRec("jobans3"))
                        %>
						, <%=meNavn%> 
                         <%if oRec("jobans_proc_3")  <> 0 then %>
                        (<%=oRec("jobans_proc_3") %>%)

                        <%end if
                        
                        end if


                        '*** Job medansvarlige ***
                        if oRec("jobans4") <> 0 then
						call meStamData(oRec("jobans4"))
                        %>
						, <%=meNavn%> 
                             <%if oRec("jobans_proc_4")  <> 0 then %>
                            (<%=oRec("jobans_proc_4") %>%)
                            <%end if %>

                        <%end if


                        '*** Job medansvarlige ***
                        if oRec("jobans5") <> 0 then
						call meStamData(oRec("jobans5"))
                        %>
						, <%=meNavn%> 
                             <%if oRec("jobans_proc_5")  <> 0 then %>
                            (<%=oRec("jobans_proc_5") %>%)
                            <%end if %>

                        <%end if

                        if level = 1 then
                             if lto <> "epi" OR (lto = "epi" AND (thisMid = 6 OR thisMid = 11 OR thisMid = 59 OR thisMid = 1720)) OR (lto = "epi_cati" AND (thisMid = 6 OR thisMid = 11 OR thisMid = 59)) OR (lto = "epi_no" AND thisMid = 2) OR (lto = "epi_ab" AND thisMid = 2) OR (lto = "epi_sta" AND thisMid = 2) OR (lto = "epi_uk" AND thisMid = 2) then 
                                if oRec("virksomheds_proc") <> 0 then%>
                                <br /><b><%=lto %>: </b> (<%=oRec("virksomheds_proc") %>%)
                                <%end if
                            end if
                        end if
                      
						
						'**********************
						%></i></span>


                        <%'*** kontakpersoner hos kunde ***'
                         
                         kpersNavn = ""
                         if cint(oRec("kundekpers")) <> 0 then

                         strSQLkpers = "SELECT navn FROM kontaktpers WHERE id = " & oRec("kundekpers")
                         oRec6.open strSQLkpers, oConn, 3
                         if not oRec6.EOF then

                         kpersNavn = oRec6("navn")

                         end if
                         oRec6.close

                         %>
                         <span style="color:#8caae6; font-size:9px;"><br />Kontaktpers.: <%=kpersNavn %></span>
                         <%

                         end if
                         
                         %>
							
						
	<br><span style="font-size:9px;"><%=formatdatetime(oRec("jobstartdato"), 1)%> - <%=formatdatetime(oRec("jobslutdato"), 1)%></span>
	
    
	<% 
	if len(trim(oRec("rekvnr"))) <> 0 then
    %>
    <span style="font-size:9px; color:#999999;">
    <%
	Response.Write "<br>Rekvnr.: "& oRec("rekvnr") 
    %>
     </span>
    <%
	end if

  
	if len(trim(oRec("aftnavn"))) <> 0 then
    %><span style="font-size:9px; background-color:#FFFFe1; color:#000000;"><%
	Response.Write "<br>Aftale: "& oRec("aftnavn") 
    %>
     </span>
    <%
	end if
	
	if oRec("jobstatus") = 3 then
    %><span style="font-size:9px; background-color:#ffdfdf; color:#000000;"><%
	Response.Write "<br>Tilbud: "& oRec("tilbudsnr") &" ("& oRec("sandsynlighed") &" %)"
    %>
     </span>
    <%
	end if


    select case preconditions_met
      case 0
    preconditions_met_Txt = ""
      case 1
    preconditions_met_Txt = "<br><span style='color:#6CAE1C; font-size:10px; background:#DCF5BD;'>Pre-konditioner opfyldt</span>"
    case 2
    preconditions_met_Txt = "<br><span style='color:red; font-size:10px; background:pink;'>Pre-konditioner ikke opfyldt!</span>"
    case else
    preconditions_met_Txt = ""
    end select


    %>
    <%=preconditions_met_Txt %>
	
	&nbsp;</td>
	
        <td valign=top style="padding:4px 10px 0px 5px;">
            Aktiviteter
		 <%
		'********* Aktiviteter ****
		if editok = 1 then%>
        <a href="#" onclick="Javascript:window.open('aktiv.asp?menu=job&jobid=<%=oRec("id")%>&jobnavn=<%=oRec("jobnavn")%>&rdir=job3&nomenu=1', '', 'width=1354,height=800,resizable=yes,scrollbars=yes')" class=vmenu>(<%=Antal%>)</a>
		<%else%>
		<b>(<%=Antal%>)</b>
		<%end if%>

        <%
        'if visAkt <> 1 then
        %>
        <div style="font-size:10px; height:80px; width:120px; overflow:auto; color:#999999; padding-right:5px;">
           
            
            <%=strAktnavn %></div>

        <%'end if
		'**************************
		%>
		</td>
		

        

        <td valign=top style="padding:4px 5px 5px 5px;" class=lille><%=strFomr_navn %>&nbsp;</td>

		<td align="left" valign=top style="padding:10px 5px 0px 5px;">
		<%if proaf <> 0 then %>
		
		
		<div style="font-size:9px; color:#999999,">Real.:
		<span style="width:<%=cint(left(projektcomplt, 3))%>px; background-color:<%=bgdiv%>; height:10px; padding:2px 0px 2px 2px; font-size:9px;">
		<%if showprojektcomplt > 0 then%>
		<%=formatpercent(showprojektcomplt/100, 0)%>
		<%end if%>
		</span>
        </div>

        <div style="font-size:9px; color:#999999;">Fakt.: 
        <span style="width:<%=cint(left(projektcompltFakbare, 3))%>px; background-color:#CCCCCC; color:#000000; height:10px; padding:2px 0px 2px 2px; font-size:9px;">
		<%if showprojektcompltFakbare > 0 then%>
		<%=formatpercent(showprojektcompltFakbare/100, 0)%>
		<%end if%>
		</span>
            </div>

      

		
		
		
		<%end if %>
		</td>
		<td align=right valign=top style="padding:4px 5px 0px 5px;" class=lille>
		<%=formatnumber(oRec("jo_bruttooms"), 2) &" "& basisValISO %> 
		</td>
		
		<td align=right valign=top style="padding:4px 15px 4px 4px;" class=lille>Forkalk: <%=formatnumber(budgettimertot)%><br>
		Realiseret:  <%=formatnumber(proaf)%><br>
        <span style="color:#999999;">Heraf fakturerbar.: <%=formatnumber(realfakbare, 2) %></span> <br />
        -----------------------<br />
        Bal. = <%=formatnumber(restt)%><br />
        <span style="color:#999999;"><%=formatnumber(resttfakbare) %></span>
        =====================
		</td>
		
		<td valign=top style="padding:4px 5px 0px 5px;"><%
		
		'if oRec("jobstatus") <> "3" then 'Tilbud
		
        stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
        stCHK3 = ""
        stCHK4 = ""
        stBgcol0 = "#999999"
		stBgcol1 = "#999999"
		stBgcol2 = "#999999"
        stBgcol3 = "#999999"
        stBgcol4 = "#999999"

        lkDato = ""

		select case oRec("jobstatus")
        case "0"
        lkDato = "("& formatdatetime(oRec("lukkedato"), 2) &")"
        stCHK0 = "CHECKED"
        stBgcol0 = "darkred"
		case "1"
		stCHK1 = "CHECKED"
        stBgcol1 = "green"
        case "2"
        stCHK2 = "CHECKED"
        stBgcol2 = "#999999"
		case "0"
        stCHK0 = "CHECKED"
        stBgcol0 = "Crimson"
        case "3"
        stCHK3 = "CHECKED"
        stBgcol3 = "#000000"
        case "4"
        stCHK4 = "CHECKED"
        stBgcol4 = "#5582d2"
		end select
		
		if editok = 1 then
		%>
        <!--
		<select name="FM_listestatus" id="FM_listestatus" style="background-color:<%=selbgcol%>; font-size:9px;">
		<option value="1" <%=stCHK1%>>Aktiv</option>
		<option value="0" <%=stCHK0%>>Lukket</option>
		<option value="2" <%=stCHK2%>>Passiv</option>
		</select>
        -->
        
        <input type="radio" class="FM_listestatus_1" name="FM_listestatus_<%=oRec("id")%>" value="1" id="FM_listestatus_1_<%=oRec("id")%>" <%=stCHK1%>/><span style="color:<%=stBgcol1%>; font-size:9px;">Aktiv</span><br />
        <input type="radio" class="FM_listestatus_2" name="FM_listestatus_<%=oRec("id")%>" value="2" id="FM_listestatus_2_<%=oRec("id")%>" <%=stCHK2%>/><span style="color:<%=stBgcol2%>; font-size:9px;">Passiv/Fak.</span><br />
        <input type="radio" class="FM_listestatus_3" name="FM_listestatus_<%=oRec("id")%>" value="3" id="FM_listestatus_3_<%=oRec("id")%>" <%=stCHK3%>/><span style="color:<%=stBgcol3%>; font-size:9px;">Tilbud<br />
        <input type="radio" class="FM_listestatus_4" name="FM_listestatus_<%=oRec("id")%>" value="4" id="FM_listestatus_4_<%=oRec("id")%>" <%=stCHK4%>/><span style="color:<%=stBgcol4%>; font-size:9px;">Gen.syn</span><br />
        <input type="radio" class="FM_listestatus_0" name="FM_listestatus_<%=oRec("id")%>" value="0" id="FM_listestatus_0_<%=oRec("id")%>" <%=stCHK0%>/><span style="color:<%=stBgcol0%>; font-size:9px;">Lukket <br /><%=lkDato %></span>

		<%else%>
		<%=stNavn%>
		<input type="hidden" name="FM_listestatus" id="Hidden1" value="<%=oRec("jobstatus")%>">
		
		<%end if %>
		
		<%'else%>
        <!--
		<b>Tilbud</b><br />
		<span style="font-size:9px; font-family:arial;"><%=oRec("tilbudsnr") %> <br /><%=oRec("sandsynlighed") %> %</span>
		<input type="hidden" name="FM_listestatus" id="FM_listestatus" value="<%=oRec("jobstatus")%>">
        -->
		<%'end if%>
		
		
		<input type="hidden" name="FM_listeid" id="FM_listeid" value="<%=oRec("id")%>">
		</td>
	
		<td valign=top style="padding:4px 5px 0px 5px; white-space:nowrap;">
		<%if editok = 1 then %>
		<a href="job_print.asp?id=<%=oRec("id")%>" class=rmenu>Print / PDF >></a><br />
		<a href="job_kopier.asp?func=kopier&id=<%=oRec("id")%>&fm_kunde=<%=oRec("kid")%>&filt=<%=request("filt")%>" class=rmenu>Kopier job >></a><br />
         <a href="materialer_indtast.asp?id=<%=oRec("id")%>&fromsdsk=0&aftid=<%=oRec("serviceaft")%>" target="_blank" class=rmenu>Indtast mat./udlæg >></a>

         <%
         
        viskunabnejob0 = "1"
        viskunabnejob1 = "" 
        viskunabnejob2 = ""
        select case oRec("jobstatus")
        case 0
        viskunabnejob0 = "1"
        case 1,2
        viskunabnejob2 = "1"
        case 3
        viskunabnejob1 = "1"
        end select

         select case lto
        case "synergi1", "xintranet - local"


        call licensStartDato()
        dtlink_stdag = startDatoDag
		dtlink_stmd = startDatoMd
		dtlink_staar = startDatoAar

        case else
        
        dtlink_stdag = datepart("d", oRec("jobstartdato"), 2, 2)
		dtlink_stmd = datepart("m", oRec("jobstartdato"), 2, 2)
		dtlink_staar = datepart("yyyy", oRec("jobstartdato"), 2, 2)

		
        end select

        

		dtlink_sldag = datepart("d", now, 2, 2)
		dtlink_slmd = datepart("m", now, 2, 2)
		dtlink_slaar = datepart("yyyy", now, 2, 2)
		
       
		
		dtlink = "FM_usedatokri=1&FM_start_dag="&dtlink_stdag&""_
	    &"&FM_start_mrd="&dtlink_stmd&"&FM_start_aar="&dtlink_staar&"&FM_slut_dag="&dtlink_sldag&""_
	    &"&FM_slut_mrd="&dtlink_slmd&"&FM_slut_aar="&dtlink_slaar
         
          %>

          <br />
          <a href="joblog.asp?nomenu=1&FM_job=<%=oRec("id") %>&FM_kunde=<%=oRec("kid")%>&FM_jobsog=<%=oRec("jobnr")%>&viskunabnejob0=<%=viskunabnejob0%>&viskunabnejob1=<%=viskunabnejob1%>&viskunabnejob2=<%=viskunabnejob2%>&<%=dtlink %>" class=rmenu target="_blank">Joblog >></a><br />
          <a href="jobprintoverblik.asp?menu=job&id=<%=oRec("id")%>&media=printjoboverblik" class=rmenu target="_blank">Print joboverblik >></a><br>
          <a href="timereg_akt_2006.asp?FM_kontakt=<%=oRec("kid")%>&FM_ignorer_projektgrupper=1&jobid=<%=oRec("id")%>&FM_sog_job_navn_nr=<%=oRec("jobnr")%>&usemrn=<%=session("mid")%>&showakt=1&fromsdsk=1" target="_blank" class=rmenu>Timeregistrering >> </a><br />
          
          <%if cint(useasfak) <= 2 then %>
          <a href="../timereg/erp_opr_faktura_fs.asp?FM_kunde=<%=oRec("kid")%>&FM_job=<%=oRec("id")%>&FM_aftale=0&<%=dtlink %>&reset=1&FM_usedatokri=1" target="_blank" class=rmenu>Opret faktura >> </a>
		  <%end if %>

		<%else %>
		&nbsp;
		<%end if %></td>
		<td class=lille valign=top style="padding-right:2px;">
		<div style="position:relative; top:0px, left:0px; width:160px; height:100px; overflow:auto; padding:5px;">
        <table cellspacing=0 border=0 cellpadding=0 width=100%>
		
		<%
		
		        '** Findes der fakturaer på job kan det ikke slettes **'
			    
			    deleteok = 0
			    faktotbel = 0
			    strSQLffak = "SELECT f.fid, f.faknr, f.aftaleid, f.faktype, f.jobid, f.fakdato, f.beloeb, "_
			    &" f.faktype, f.kurs, SUM(fd.aktpris) AS aktbel, brugfakdatolabel, labeldato, fakadr FROM fakturaer f "_
			    &" LEFT JOIN faktura_det AS fd ON (fd.fakid = f.fid AND fd.enhedsang <> 3)"_
			    &" WHERE jobid = " & oRec("id") & ""_
			    &" AND aftaleid = 0 AND shadowcopy = 0"_
			    &" GROUP BY f.fid ORDER BY f.fakdato DESC"
			    'Response.Write "strSQLffak<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>" & strSQLffak 
			    'Response.flush
			    f = 0
			    oRec3.open strSQLffak, oConn, 3
			    while not oRec3.EOF 
			        

                    if oRec3("brugfakdatolabel") = 1 then '** Labeldato
                    fakDato = "L: "& oRec3("labeldato") & "<br><span style=""color:#999999; size:9px,"">"& oRec3("fakdato") &"</span>"
                    else
                    fakDato = "F: "& oRec3("fakdato")
                    end if
			       
			          
                      %>
                      <tr><td class=lille valign=top>
                      <%
			        if cdate(oRec3("fakdato")) >= cdate("01-01-2006") AND editok = 1 then%>
			        <a href="erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&id=<%=oRec3("fid")%>&FM_jobonoff=<%=FM_jobonoffval%>&FM_kunde=<%=oRec3("fakadr")%>&FM_job=<%=oRec3("jobid")%>&FM_aftale=<%=oRec3("aftaleid")%>&fromfakhist=1" class="lgron" target="_blank"><b><%=oRec3("faknr") &"</b>"%></a></td>
                    <td align=right class=lille valign=top> <%=fakDato %> <br />
                    <%else%>
                    <b><%=oRec3("faknr") &"</b></td><td align=right class=lille valign=top> "& fakDato %><br />
                    <% end if
                    

                      %>

                      </td></tr>
                      <%
                    
                     
	                call beregnValuta(minus&(oRec3("beloeb")),oRec3("kurs"),100)
                    if oRec3("faktype") <> 1 then
                    belobGrundVal = valBelobBeregnet
                    else
                    belobGrundVal = -valBelobBeregnet
                    end if 
                    

                        if cDate(oRec3("fakdato")) < cDate("01-06-2010") AND (lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" OR lto = "epi_uk") then
                        belobKunTimerStk = belobKunTimerStk + oRec3("beloeb")
                        else
       
                    
	                        '** Kun aktiviteter timer, enh. stk. IKKE materialer og KM
	                        call beregnValuta(minus&(oRec3("aktbel")),oRec3("kurs"),100)
                            if oRec3("faktype") <> 1 then
                            belobKunTimerStk = valBelobBeregnet
                            else
                            belobKunTimerStk = -valBelobBeregnet
                            end if 


                     end if
            
                faktotbel = faktotbel + belobKunTimerStk
			    f = f + 1
			    oRec3.movenext
			    wend
			    oRec3.close
                
			    %>
                </table>
			    </div>
                

            <%if totReal <> 0 then 
	        gnstpris = faktotbel/totReal
	        else 
	        gnstpris = 0
	        end if%>
	        
            <%if gnstpris <> 0 then %>
            Faktisk timepris:<br />
	        <b><%=formatnumber(gnstpris) & " "& basisValISO %></b>
            
            <!-- <span class="qmarkhelp" id="qm0001" style="font-size:11px; color:#999999; font-weight:bolder;">?</span><span id="qmarkhelptxt_qm0001" style="visibility:hidden; color:#999999; display:none; padding:3px; z-index:4000;">(faktureret beløb - (materialeforbrug + km)) / timer realiseret</span>-->
	        
	        <%
            end if

	        if f <> 0 AND proaf <> 0 then
	        gnsPrisTot = gnsPrisTot + (faktotbel)
	        else
	        gnsPrisTot = gnsPrisTot
	        'totRealialt = totRealialt
	        end if
	        
	        totRealialt = totRealialt + (proaf)
	        %>

			    
		</td>
		<td align=right class=lille style="white-space:nowrap; padding:4px 4px 0px 4px;" valign=top>
        <% for p = 1 to 10
        
        pgid = oRec("projektgruppe"&p)

        if pgid <> 1 then
            call prgNavn(pgid, 200)
            Response.Write left(prgNavnTxt, 20) & "<br>"
        end if
        
        next %>
		   
		</td>

		<td valign=top style="padding:4px 2px 4px 10px;">
		<%
		    
		    if f = 0 AND editok = 1 then
		    deleteok = 1
		    else
		    deleteok = 0
		    end if
		    
		    
		if deleteok = 1 then%>
		<a href="jobs.asp?menu=job&func=slet&id=<%=oRec("id")%>&jobnr_sog=<%=jobnr_sog%>&filt=<%=filt%>&fm_kunde_sog=<%=fmkunde%>"><%=sletkanp%></a>&nbsp;
		<%else %>
		&nbsp;
		<%end if%>
		</td>
	</tr>
	<%
	cnt = cnt + 1
	x = 0
	lastKid = oRec("kid")

    select case right(cnt,1)
    case 0
    Response.flush
    end select

    end if' forretningsområde


    
	oRec.movenext
	wend
	
	if cnt > 0 then%>
    	<tr style="background-color:#FFFFFF;">
		<td colspan=20 style="padding-right:30px; padding-top:5px;" align=right><br /><input type="submit" name="statusliste" value="Opdater liste >>"><br />&nbsp;</td>
	</tr>
    <%end if %>
	</table>
	</div>
	<!-- end table div -->



    
	<%if cnt = 0 AND request("fromvemenu") <> "j" then%>
	
    <br /><br /><br /><br />

         <table cellspacing="0" cellpadding="2" border="0" bgcolor="#ffffff" width=100%>
	<tr>
		<td style="padding-left:20;"><br>
		<b>Der blev ikke fundet nogen job der macthedede de valgte søgekriterier.</b><br>
		Vær opmærksom på om du har valgt:<br>
			<ul>
			<li>Et <u>Status filter.</u>
			<li>En <u>Periode.</u> 
			<li>Eller en bestemt <u>Kunde.</u></ul>
			Der betyder at der ikke blev fundet nogen job.
		</td>
	</tr>
	
    </table>
    <br /><br />&nbsp;

	<%
    'response.end
    end if%>	
	
	
    <%
    if cnt > 0 then
    
    tTop = 40
	tLeft = 0
	tWdth = 600
	
	
	call tableDiv(tTop,tLeft,tWdth)
	 %>

    <br />
    <h3>Funktioner</h3>
	<table cellspacing=1 cellpadding=2 border=0 width=100%>
	

    <%if level = 1 then %>
     <tr style="background-color:#FFFFFF;"><td colspan=2>
 
         <b>Multiopdater forretningsområder:</b><br />

                              
                                <%
                                
                                strSQLf = "SELECT f.id, f.navn, kp.navn AS kontonavn, kp.kontonr FROM fomr AS f LEFT JOIN kontoplan AS kp ON (kp.id = f.konto) WHERE f.id <> 0 AND f.jobok = 1 ORDER BY f.navn"
                                
                                'response.write strSQLf
                                'response.flush
         
                                %>
                                <select name="FM_fomr" id="Select3" multiple="multiple" size="5" style="width:450px;">
                                <option value="#0#">Ingen (fjern)</option>
                               
                                    
                                    <%oRec3.open strSQLf, oConn, 3
                                    while not oRec3.EOF
                                    
                                    if oRec3("kontonr") <> 0 then
                                        kontnrTxt = " ("& oRec3("kontonavn") &" "& oRec3("kontonr") &")"
                                    else 
                                        kontnrTxt = ""
                                    end if
                                    
                                    %>
                                    <option value="#<%=oRec3("id") %>#"><%=oRec3("navn") & kontnrTxt %></option>
                                    <%
                                    oRec3.movenext
                                    wend
                                    oRec3.close
                                    %>
                                </select>

        <br />
         <input id="Checkbox3" type="checkbox" name="FM_formr_opdater" value="1" /> Opdater alle job på listen med ovenstående forretningsområder. (forvalgt konto på job skal ændres manuelt)<br />
          <input id="Checkbox3" type="checkbox" name="FM_formr_opdater_akt" value="1" /> Opdater også alle tilhørende aktiviteter 


    </td></tr>
    <%end if %>


    <%dtNow = day(now) & "-"& month(now) & "-"& year(now) %>

    <tr style="background-color:#FFFFFF;">
		<td style="padding-right:30px; padding-top:10px;">
        <b>Slutdato:</b><br />
        <input type="checkbox" name="FM_opdaterslutdato" value="1">Sæt slutdato på alle ovenstående job = <input type="text" name="FM_opdaterslutdato_dato" value="<%=dtNow%>" style="width:75px;" /> dd-mm-yyyy</td>
	</tr>
	<tr style="background-color:#FFFFFF;">
		<td style="padding-right:30px; padding-top:5px;" align=right><br /><input type="submit" name="statusliste" value="Opdater liste >>"></td>
	</tr></form>
    </table>

    	</table>
	</div>
	<!-- end table div -->

	<%end if%>
	

    <%
    oRec.Close
    
    
    else 'vis_timepriser%>
    <br />
    <b>Timepriser på job</b> <br />
    Hvis aktiviteten ikke er vist, følger den timeprisen på jobbet.

    <table cellspacing="0" cellpadding="2" border="0" bgcolor="#ffffff" width=100%>
	
            <tr bgcolor="#5582D2">
                <td class="alt">Kontakt</td>
                <td class="alt">Job</td>
                <td class="alt">Fase</td>
                <td class="alt">Aktivitet</td>
                <td class="alt">Medarbejder</td>
                <td class="alt">Initialer</td>
                <td class="alt">Timepris</td>
                <td class="alt">Valuta</td>
                
            </tr>
            <%
            
            jids = 0

            if sorttp = 0 then
            sortTpBy = "a.fase, a.navn, m.mnavn"
            else
            sortTpBy = "m.mnavn, a.fase, a.navn"
            end if
              

                        strSQL = "SELECT tp.jobid, tp.aktid, tp.medarbid, tp.timeprisalt, tp.6timepris, tp.6valuta, j.jobnavn, j.jobnr, a.navn AS aktnavn, a.fase, m.mnavn, m.init, m.mnr, "_
                        &" v.valutakode, j.jobans1, j.jobans2, j.id, "_
                        &" k.kkundenavn, k.kkundenr FROM job AS j, kunder AS k"_
                        &" LEFT JOIN timepriser AS tp ON (tp.jobid = j.id) "_
                        &" LEFT JOIN aktiviteter AS a ON (a.id = tp.aktid) "_
                        &" LEFT JOIN medarbejdere AS m ON (m.mid = tp.medarbid) "_
                        &" LEFT JOIN valutaer AS v ON (v.id = tp.6valuta) "_
                        &" WHERE "& varJobknrKri &" "& sogeKri &" "& jobansKri &" k.kid = j.jobknr " & varFilt &" "& datoKri &" "& replace(prjgrpSQLkri, "projektgruppe", "j.projektgruppe") &""_ 
                        &" AND j.jobnavn IS NOT NULL AND m.mnavn IS NOT NULL GROUP BY tp.jobid, tp.aktid, tp.medarbid ORDER BY k.kkundenavn, j.jobnavn, j.jobnr, "& sortTpBy &" LIMIT 0,10000"
						'Response.write strSQL & "<br>"
						'Response.flush
                        oRec2.open strSQL, oConn, 3 
						    
                            cnt = 0
                            while not oRec2.EOF 


                             call forretningsomrJobId(oRec2("id"))


                            if cint(visJobFomr) = 1 OR instr(strFomr_rel, "#0#") <> 0 then


                            select case right(cnt,1)
                            case 0,2,4,6,8
                            bgthis = "#FFFFFF"
                            case else
                            bgthis = "#EFF3FF"
                            end select




						    'thissel = oRec2("timeprisalt")
						    this6timepris = formatnumber(oRec2("6timepris"), 2)
						    this6valuta = oRec2("valutakode")

                            
                            if IsNull(oRec2("aktnavn")) = False OR oRec2("aktid") = 0 then

                            %>
                            <tr bgcolor="<%=bgthis%>">
                                <td style="padding:3px; border-bottom:1px #999999 solid; white-space:nowrap;"><b><%=left(oRec2("kkundenavn"), 65) & "</b> ("& oRec2("kkundenr") &")" %></td>

                                <td style="border-bottom:1px #999999 solid; white-space:nowrap;">

                                <%
                                editok = 0
	                            if level <= 2 OR level = 6 then
	                            editok = 1
	                            else
			                            if cint(session("mid")) = oRec2("jobans1") OR cint(session("mid")) = oRec2("jobans2") OR (cint(oRec2("jobans1")) = 0 AND cint(oRec2("jobans2")) = 0) then
			                            editok = 1
			                            else
			                            editok = 0   
			                            end if
	                            end if
	
	
	              
	
	
	                           
	                            if editok = 1 then
		                           %>
                                   <a href="jobs.asp?menu=job&func=red&id=<%=oRec2("id")%>&int=1&jobnr_sog=<%=jobnr_sog%>&filt=<%=filt%>&fm_kunde_sog=<%=fmkunde%>&showdiv=tpriser" class=vmenu><%=oRec2("jobnavn")%>&nbsp;&nbsp;(<%=oRec2("jobnr")%>)</a>&nbsp;
		                           <%
	                            else
		                            %>
		                            <b><%=oRec2("jobnavn")%>&nbsp;&nbsp;(<%=oRec2("jobnr")%>)</b>&nbsp;
		                            <%
	                            end if
                                %>


                                </td>
                                <td style="border-bottom:1px #999999 solid; white-space:nowrap;"><%=oRec2("fase") %>&nbsp;</td>
                                <td style="border-bottom:1px #999999 solid; white-space:nowrap;"><%=oRec2("aktnavn") & " <span style=""font-size:8px; color:#cccccc;"">"& oRec2("aktid") %></span>&nbsp;</td>
                                <td style="border-bottom:1px #999999 solid; white-space:nowrap;"><%=oRec2("mnavn") & " ("& oRec2("mnr") &")" %></td>
                                <td style="border-bottom:1px #999999 solid; white-space:nowrap;"><%=oRec2("init")  %>&nbsp;</td>
                                <td style="border-bottom:1px #999999 solid; white-space:nowrap;"><%=this6timepris %></td>
                                <td style="border-bottom:1px #999999 solid; white-space:nowrap;"><%=this6valuta %></td>
                            </tr>
                          
                            <%
                           
                            select case right(cnt,1)
                            case 0
                            Response.flush
                            case else
                            
                            end select
                           
                            if lastJob <> oRec2("jobid") then
                            jids = jids & "," & oRec2("jobid")
                            lastJob = oRec2("jobid")
                            end if

                            cnt = cnt + 1


                            end if 'Is null (akt slettet)

                            end if 'Forretningsomr

                            oRec2.movenext
						    wend 
						oRec2.close 

    
            
            
            %>
            </table>
	</div>
	<!-- end table div -->



    <%end if %>
	
	
	
	
	

    <%  

    'if cint(cnt) <> 0 then


ptop = 0
pleft = 946
pwdt = 210

call eksportogprint(ptop,pleft, pwdt)
%>

<form action="job_eksport.asp" method="post" target="_blank">
<tr> <input id="Hidden3" name="jids" value="<%=jids%>" type="hidden" />
    <td valign=top align=center>
   <input type=image src="../ill/export1.png" />
    </td>
    <td class=lille><input id="Submit5" type="submit" value="A) Eksportér jobdata >> " style="font-size:9px; width:130px;" />
    <input type="hidden" value="1" name="eksDataStd" /><br />
    <input type="checkbox" value="1" name="xeksDataStd" checked disabled /> Stamdata<br />
    <input type="checkbox" value="1" name="eksDataNrl" /> Nøgletal, Realiseret, Forr.omr., Projektgrupper mm.<br />
    <input type="checkbox" value="1" name="eksDataJsv" /> 
        <%
        call salgsans_fn()    
        if cint(showSalgsAnv) = 1 then  %>
        Job- og salgs -ansvarlige
        <%else %>
        Jobansvarlige
        <%end if %><br />
    <input type="checkbox" value="1" name="eksDataAkt" /> Aktiviteter<br />
    <input type="checkbox" value="1" name="eksDataMile" /> Milepæle
    </td>
</tr>
</form>


<form action="job_eksport.asp?optiprint=3" method="post" target="_blank">
<tr> <input id="Hidden5" name="jids" value="<%=jids%>" type="hidden" />
<input id="Hidden6" name="sorttp" value="<%=sorttp%>" type="hidden" />
    <td valign=top align=center>
   <input type=image src="../ill/export1.png" />
    </td>
    <td class=lille><input id="Submit3" type="submit" value="B) Eksportér timepriser >> " style="font-size:9px; width:130px;" /></td>
</tr>
</form>




<!--
<form action="job_eksport.asp?optiprint=4" method="post" target="_blank">
<tr> <input id="Hidden6" name="jids" value="<%=jids%>" type="hidden" />
    <td valign=top align=center>
   <input type=image src="../ill/export1.png" />
    </td>
    <td class=lille><input id="Submit4" type="submit" value="C) Eksportér jobansvarlige >> " style="font-size:9px; width:130px;" /></td>
</tr>
</form>

<form action="job_eksport.asp?optiprint=5" method="post" target="_blank">
<tr> <input id="Hidden7" name="jids" value="<%=jids%>" type="hidden" />
    <td valign=top align=center>
   <input type=image src="../ill/export1.png" />
    </td>
    <td class=lille><input id="Submit7" type="submit" value="D) Eksportér job og akt. >> " style="font-size:9px; width:130px;" /></td>
</tr>
</form>

-->

<form action="job_eksport.asp?optiprint=1" method="post" target="_blank">
<tr>
<input id="Hidden4" name="jids" value="<%=jids%>" type="hidden" />
    <td valign=top align=center>
   <input type=image src="../ill/printer3.png" />
    </td><td class=lille><input id="Submit6" type="submit" value="E) Print som arbejdskort >>" style="font-size:9px; width:130px;" /></td>
</tr>
</form>


<%if instr(lto, "epi") <> 0 OR lto = "intranet - local" then %>
<form action="job_eksport.asp?optiprint=6" method="post" target="_blank">
<tr> <input id="Hidden5" name="jids" value="<%=jids%>" type="hidden" />
    <td valign=top align=center>
   <input type=image src="../ill/export1.png" />
    </td>
    <td class=lille><input id="Submit7" type="submit" value="F) Eksportér som BU fil >> " style="font-size:9px; width:130px;" /></td>
</tr>
</form>

<%end if %>
	
	
	
    <!--
	<br /><br /><a href="job_eksport.asp?jids=<%=jids%>" class=vmenu target="_blank">&nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a>&nbsp;
	<br><a href="job_eksport.asp?jids=<%=jids%>&optiprint=1" class=vmenu target="_blank">Eksporter ovenstående liste af job. (optimeret til print)&nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a>&nbsp;
	<br><br>
    -->
	
	
	
	

	<%
	
	strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus <> 3"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	antalEksterneJob = oRec("antal")
	end if
	oRec.close 
	
	if len(antalEksterneJob) <> 0 then
	antalEksterneJob = antalEksterneJob
	else
	antalEksterneJob = 0
	end if
	%>
	
	<!--<b><=antalEksterneJob%></b> Eksterne job.
	<br>-->
	
	
	<%
	strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus = 1"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	antalEksterneAktiveJob = oRec("antal")
	end if
	oRec.close 
	
	if len(antalEksterneJob) <> 0 then
	antalEksterneAktiveJob = antalEksterneAktiveJob
	else
	antalEksterneAktiveJob = 0
	end if
	
	strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus = 0"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	antalEksterneLukkedeJob = oRec("antal")
	end if
	oRec.close 
	
	if len(antalEksterneJob) <> 0 then
	antalEksterneLukkedeJob = antalEksterneLukkedeJob
	else
	antalEksterneLukkedeJob = 0
	end if
	
	
	'*** Tilbud ***'
	strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus = 3"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	antalEksternePassiveJob = oRec("antal")
	end if
	oRec.close 
	
	if len(antalEksterneJob) <> 0 then
	antalEksternePassiveJob = antalEksternePassiveJob
	else
	antalEksternePassiveJob = 0
	end if
	
	antalTilbud = antalEksternePassiveJob
	
	
       

        %><tr><td colspan="2"><br /><br />
            <%
                nWdt = 120
                nTxt = "Opret nyt job"
                nLnk = "jobs.asp?menu=job&func=opret&id=0&int=1"
                nTgt = ""
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) %>
          
                <br />
            
            <%if lto = "outz" OR lto = "intranet - local" then %>

            <%
                nWdt = 120
                nTxt = "Opret nyt job fra fil"
                nLnk = "createJobFromFile.aspx?lto="&lto&"&editor="&session("user")
                nTgt = "_blank"
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) %>

                <br /><br />
                 <%end if
	
        
	
	if totRealialt < 1 then
	totRealialt = 1
	end if
	
	
	uTxt = "<b>Antal job oprettet i TimeOut:</b><br><b>"& antalEksterneAktiveJob & "</b> Aktive job."_
	&"<br><b>" & antalEksterneLukkedeJob &"</b> Lukkede job"_
	&"<br><b>" & antalTilbud & "</b> Tilbud"
	
    if cint(vis_timepriser) <> 1 then
    uTxt = uTxt  &"<br><b>" & cnt & "</b> Job i denne visning"
	else
    uTxt = uTxt  &"<br><b>" & cnt & "</b> timepris linier i denne visning."
    end if

    uTxt = uTxt  &"<br><br>Gns. faktisk timepris:<br><b> "& formatnumber(gnsPrisTot/totRealialt) &" "& basisValISO &" </b> "
	'&" Gns. faktisk timepris = (faktureret beløb ekskl. mat. og km. / real. timer)<br>"_
	'&" Gælder for de viste job, hvor der forefindes fakturaer og realiserede timer.<br>Gns. timepris er vægtet i forhold til timeforbrug på de enkelte job."
	
                Response.write uTxt


	%>

	</td></tr>
</table></div>	
<%'end if %>




	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>

	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>


<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br><br>
	<br>
&nbsp;





	<a href="Javascript:history.back()" class="vmenu"><< Tilbage </a>
	<br>
	<br>
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->

