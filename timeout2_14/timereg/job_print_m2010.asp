<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/job_func.asp"-->
<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" -->



<%
'** Bruges til PDF visning ***
if len(request("nosession")) <> 0 then
nosession = request("nosession")
else
nosession = 0
end if



                     

if len(trim(session("user"))) = 0 AND cint(nosession) = 0  then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	func = request("func")
	media = request("media")
	
	if len(trim(request("id"))) <> 0 then
	id = request("id")
	else
	id = 0
	end if
	
	if len(trim(request("kid"))) <> 0 then
	kid = request("kid")
	else
	kid = 0
	end if
	
	if len(request("tb_skabeloner")) <> 0 then
    skabelonid = request("tb_skabeloner")
    else
    skabelonid = 0
    end if


    if len(trim(request("FM_dokument"))) <> 0 then
    dokid = request("FM_dokument")
    else
    dokid = 0
    end if
    
    
    '** 1 = skjul, 0 = vis
    '*** data indstillings valg ****'
	if len(trim(request("kinfo"))) <> 0 then
	kinfo = request("kinfo")
	else
	    
	    if len(trim(request.cookies("job")("kinfo"))) <> 0 then
	    kinfo = request.cookies("job")("kinfo")
	    else
	    kinfo = 0
	    end if
	
	end if
	response.cookies("job")("kinfo") = kinfo
	
	if len(trim(request("jbinfo"))) <> 0 then
	jbinfo = request("jbinfo")
	else
	    
	    if len(trim(request.cookies("job")("jbinfo"))) <> 0 then
	    jbinfo = request.cookies("job")("jbinfo")
	    else
	    jbinfo = 0
	    end if
	
	end if
	response.cookies("job")("jbinfo") = jbinfo
	
	
	if len(trim(request("jbprisinfo"))) <> 0 then
	jbprisinfo = request("jbprisinfo")
	else
	    
	    select case lto
	    case "wowern"
	    jbprisinfo = 1
	    
	     case else
	    
	    if len(trim(request.cookies("job")("jbprisinfo"))) <> 0 then
	    jbprisinfo = request.cookies("job")("jbprisinfo")
	    else
	    jbprisinfo = 0
	    end if
	    
	    end select
	
	end if
	response.cookies("job")("jbprisinfo") = jbprisinfo
	
	
	if len(trim(request("jbperansinfo"))) <> 0 then
	jbperansinfo = request("jbperansinfo")
	else
	    
	    if len(trim(request.cookies("job")("jbperansinfo"))) <> 0 then
	    jbperansinfo = request.cookies("job")("jbperansinfo")
	    else
	    jbperansinfo = 0
	    end if
	
	end if
	response.cookies("job")("jbperansinfo") = jbperansinfo
	
	if len(trim(request("aktinfo"))) <> 0 then
	aktinfo = request("aktinfo")
	else
	    
	    if len(trim(request.cookies("job")("aktinfo"))) <> 0 then
	    aktinfo = request.cookies("job")("aktinfo")
	    else
	    aktinfo = 0
	    end if
	
	end if
	response.cookies("job")("aktinfo") = aktinfo
	
	
	if len(trim(request("aktinfo_besk"))) <> 0 then
	aktinfo_besk = request("aktinfo_besk")
	else
	    
	    if len(trim(request.cookies("job")("aktinfo_besk"))) <> 0 then
	    aktinfo_besk = request.cookies("job")("aktinfo_besk")
	    else
	    aktinfo_besk = 0
	    end if
	
	end if
	response.cookies("job")("aktinfo_besk") = aktinfo_besk
	
	
	if len(trim(request("vedrinfo"))) <> 0 then
	vedrinfo = request("vedrinfo")
	
	if vedrinfo = "-" then
	vedrinfo = ""
	end if
	
	else
	    
	    if len(trim(request.cookies("job")("vedrinfo"))) <> 0 then
	    vedrinfo = request.cookies("job")("vedrinfo")
	    else
	        select case lto
	        case "essens"
	        vedrinfo = "Rekvisition"
	        case "dencker"
	        vedrinfo = "Følgeseddel"
	        case "wowern"
	        vedrinfo = "Budget"
	        case else
	        vedrinfo = "Job"
	        end select
	    end if
	
	end if
	response.cookies("job")("vedrinfo") = vedrinfo
	
	
	
	
	if len(trim(request("aktinfo_priser"))) <> 0 then
	aktinfo_priser = request("aktinfo_priser")
	else
	    
	    if len(trim(request.cookies("job")("aktinfo_priser"))) <> 0 then
	    aktinfo_priser = request.cookies("job")("aktinfo_priser")
	    else
	    aktinfo_priser = 0
	    end if
	
	end if
	response.cookies("job")("aktinfo_priser") = aktinfo_priser
	
	
	if len(trim(request("aktinfo_timer"))) <> 0 then
	aktinfo_timer = request("aktinfo_timer")
	else
	    
	    if len(trim(request.cookies("job")("aktinfo_timer"))) <> 0 then
	    aktinfo_timer = request.cookies("job")("aktinfo_timer")
	    else
	    aktinfo_timer = 1
	    end if
	
	end if
	response.cookies("job")("aktinfo_timer") = aktinfo_timer
	
	
	
	'*** Valgte aktiviteter ***'
	if len(trim(request("FM_vis_akt"))) <> 0 then
	vlgAkt = request("FM_vis_akt")
	else
	    if len(trim(request("aktinfo"))) <> 0 then
	    vlgAkt = 0
	    else
	    vlgAkt = -1
	    end if
	end if
	
	
	'*** Valgte faser ***'
	if len(trim(request("FM_vis_fase"))) <> 0 then
	vlgFase = request("FM_vis_fase")
	else
	    if len(trim(request("aktinfo"))) <> 0 then
	    vlgFase = 0
	    else
	    vlgFase = -1
	    end if
	    
	end if
	
	
	'select case lto
	'case "wowern" 
	'mileinfo = 0
	
	'case else
	
	if len(trim(request("mileinfo"))) <> 0 then
	mileinfo = request("mileinfo")
	else
	    
	    if len(trim(request.cookies("job")("mileinfo"))) <> 0 then
	    mileinfo = request.cookies("job")("mileinfo")
	    else
	    mileinfo = 1
	    end if
	
	end if
	response.cookies("job")("mileinfo") = mileinfo
	
	'end select
	
	thisfile = "job_print.asp"
	
	
	                
	                
	                '**** afslutning ****'
	                if len(trim(request("FM_vis_afslutning"))) <> 0 then
		            afslutningVal = trim(request("FM_vis_afslutning"))
		            else
            		    
		                if request.Cookies("job")("afslutning") <> "" then
		                afslutningVal = request.Cookies("job")("afslutning")
		                else
		                afslutningVal = ""
		                end if
            		
            		
		            end if
            		
            		
		            if len(request("FM_gem_afs")) <> 0 then
		            gemafs = 1
		            else
		            gemafs = 0
		            end if
            		
		            if gemafs = 1 then
		            response.Cookies("job")("afslutning") = afslutningVal
		            end if
	
	
	select case lto
	case "wowern", "intranet - local"
	     
	     '** force footer indhold ***'
	     if media <> "pdf" AND func <> "print" then
	     afslutningVal = "<img src=""../ill/wowern/wowern_footer_091208.gif"" border=0>"
	     end if
    
    case "outz"
            
            '** force footer indhold ***'
	     if media <> "pdf" AND func <> "print" then
	     afslutningVal = "<img src=""../ill/outz/outzource_logo_rgb.gif"" border=0>"
	     end if
            
            
 	case else
	afslutningVal = afslutningVal
	end select
	
	'*** Print afslutnings TXT **'
	afslutningsTXT = afslutningVal 'request("FM_vis_afslutning")
	
	
	%>
	
	
	
	
	
	<%
	
	if func = "sletskabelon" then
	
	ts_id = request("sletskabelonid")
	strSQLdelts = "DELETE FROM tilbuds_skabeloner WHERE ts_id = " & ts_id
	oConn.execute(strSQLdelts)
	
	Response.Redirect "job_print.asp?id="&id
	
	
	end if
	
	
	if func = "print" AND request("printgemfil") = "1" then
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
    filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
	prNavn = "print_"& filnavnDato &" "& filnavnKlok &".txt"
	prTxt = request("FM_vis_indledning")
	prEditor = session("user")
	
	strSQLinsfil = "INSERT INTO filer (filnavn, filertxt, editor, dato, type, jobid, kundeid, folderid, "_
	&" adg_kunde, adg_alle, incidentid) VALUES ('"& prNavn &"', '"&prTxt&"', '"& prEditor &"', "_
	&" '"& filnavnDato &"', 5, "& id &", "& kid &", 1000, 0, 1, 0)"
	
	oConn.execute(strSQLinsfil)
	
	call TimeOutVersion()

	'** gem fysisk ***'
	file = prNavn
	
	Set objFSO = server.createobject("Scripting.FileSystemObject")
	
	if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\job_print.asp" then
							
		Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\upload\"&file&"", True, False)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\upload\"&file&"", 8)
	
	else
		
		Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"&file&"", True, false)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"&file&"", 8, -1)
		
	end if
	
	
    'call htmlparseCSV(skabelonTxt)
	'htmlparseCSVtxt
	
	objF.WriteLine(prTxt)
    objF.close
	
	
	end if
	
	if func = "gem" then
	
	'*** Opdaterer felter ***'
	'jobbeskupd = replace(request("FM_jobbesk"), "'", "''")
	'strSQLu = "UPDATE job SET beskrivelse = '"& jobbeskupd &"' WHERE id = " & id
	'oConn.execute(strSQLu) 
	
	if len(trim(request("gemsomskabelon"))) = 0 AND len(trim(request("gemsomfil"))) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 146
	call showError(errortype)
	
	Response.end
	
	end if
	
	'*** Gemmer skabelon ****'
	if request("gemsomskabelon") = "1" OR request("gemsomfil") = "1" then
	
	if len(trim(request("skabelonnavn"))) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 144
	call showError(errortype)
	
	Response.end
	end if
	
	if request("gemsomfil") = "1" AND (len(trim(request("foid"))) = 0 OR request("foid") = 0) then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 145
	call showError(errortype)
	
	Response.end
	
	end if
	
	skabelonNavn = replace(request("skabelonnavn"), "'", "''")
	skabelonTxt = replace(request("FM_gem_txt"), "'", "''")
	tsDato = year(now) & "/"& month(now) &"/"& day(now)
	tsEditor = session("user")
	ts_kundeid = request("gemsomskabelon_kid")
	
	ts_folderid = request("foid")
	
	
	if request("gemsomskabelon") = "1" then
	strSQLinsSkab = "INSERT INTO tilbuds_skabeloner (ts_navn, ts_txt, ts_dato, ts_editor, ts_kundeid) VALUES "_
	&" ('"& skabelonNavn& "', '"& skabelonTxt &"', '"& tsDato &"', '"& tsEditor &"', "& ts_kundeid &")"
	
	oConn.execute(strSQLinsSkab)
	
	end if
	
	if request("gemsomfil") = "1" then
	
	
	if cint(dokid) <> 0 then
	
	strSQLinsfil = "UPDATE filer SET filnavn = '"& skabelonNavn &".txt', "_
	&" filertxt = '"&skabelonTxt&"', editor = '"& tsEditor &"', dato = '"& tsdato &"', type = 5, "_
	&" jobid = "& id &", kundeid = "& ts_kundeid &", folderid = "& ts_folderid &""_
	&" WHERE id = " & dokid
	
	else
	strSQLinsfil = "INSERT INTO filer (filnavn, filertxt, editor, dato, type, jobid, kundeid, folderid, "_
	&" adg_kunde, adg_alle, incidentid) VALUES ('"& skabelonNavn &".txt', '"&skabelonTxt&"', '"& tsEditor &"', "_
	&" '"& tsdato &"', 5, "& id &", "& ts_kundeid &", "& ts_folderid &", 0, 1, 0)"
	end if
	
	oConn.execute(strSQLinsfil)
	'Response.Write strSQLinsfil
	
	'** gem fysisk ***'
	file = skabelonNavn&".txt"
	
	Set objFSO = server.createobject("Scripting.FileSystemObject")
	
	if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\job_print.asp" then
							
		Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\upload\"&file&"", True, False)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\upload\"&file&"", 8)
	
	else
		
		Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"&file&"", True, false)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"&file&"", 8, -1)
		
	end if
	
	
    'call htmlparseCSV(skabelonTxt)
	'htmlparseCSVtxt
	
	objF.WriteLine(skabelonTxt)
    objF.close
	
	
	
	
	end if
	
	'Response.flush
	'Response.end
    Response.Redirect "job_print.asp?id="&id&"&FM_dokument="&dokid
	
	
	
	end if
	
	end if
	
	
	
	
	if func <> "print" then%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(3)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call jobtopmenu()
	%>
	</div>
	<%else %>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%end if 
	
	
	%>
	<script>


	    function gemtxt() {
	        document.getElementById("FM_gem_txt").value = document.getElementById("FM_vis_indledning").value
	    }


	    function gemtxtPdf() {
	        document.getElementById("FM_vis_indledning_pdf").value = document.getElementById("FM_vis_indledning").value
	        document.getElementById("FM_vis_afslutning_pdf").value = document.getElementById("FM_vis_afslutning").value
	  
	    }





	    $(document).ready(function() {

	        $("#FM_vis_akt").click(function() {

	            var n = $("#FM_vis_akt:checked").length;

	            if (n == 1) {
	                $(".chkakt").attr("checked", "checked")
	                //$("#af_95").checked = false;
	                //document.getElementById("af_95").checked = false
	            } else {
	                $(".chkakt").removeAttr("checked");

	            }



	        });


	        $(".faseoskrift").click(function() {

	            var thisid = this.id
	            //alert(thisid)
	            var n = $("#" + thisid + ":checked").length;
	            //alert(n)

	            for (i = 0; i < 100; i++) {
	                if (n == 1) {
	                    $("#af_" + thisid + "_" + i).attr("checked", "checked")
	                } else {
	                    $("#af_" + thisid + "_" + i).removeAttr("checked");

	                }
	            }


	        });

	        $("#aktinfo_0").click(function() {

	            var thisid = this.id
	            //alert(thisid)
	            var n = $("#" + thisid + ":checked").length;
	            //alert(n)

	            //for (i = 0; i < 100; i++) {
	            if (n == 1) {
	                $(".aktoskrift").attr("checked", "checked");
	                $(".faseoskrift").attr("checked", "checked");
	            } else {
	                $(".aktoskrift").removeAttr("checked");
	                $(".faseoskrift").removeAttr("checked");

	            }
	            //}


	        });

	        $("#aktinfo_1").click(function() {

	            var thisid = this.id
	            //alert(thisid)
	            var n = $("#" + thisid + ":checked").length;
	            //alert(n)

	            //for (i = 0; i < 100; i++) {
	            if (n == 1) {
	                $(".aktoskrift").removeAttr("checked");
	                $(".faseoskrift").removeAttr("checked");
	            } else {
	                $(".faseoskrift").attr("checked", "checked");
	                $(".aktoskrift").attr("checked", "checked");
	            }
	            //}


	        });




	    });

	    function setskabid() {
	        document.getElementById("sletskabelonid").value = document.getElementById("tb_skabeloner").value
	    }
        
	 
	</script>
	
	<%
	
	        
            
            
	
	
	if func <> "print" then
	        lft = 20
	        tp = 142
        	
	        sdWdt = 1104
	        
	        
	          '*** Henter gemt dokument ***'
	            dokTxt = ""
	            doknavn = ""
	            folderid = 0
	            if cint(dokid) <> 0 then
	            strSQL = "SELECT filertxt, filnavn, folderid FROM filer WHERE id = " & dokid
                oRec.open strSQL, oConn, 3
                while not oRec.EOF
                
                pkt = instr(oRec("filnavn"), ".")
                
                if pkt <> 0 then
                left_filnavn = left(oRec("filnavn"), pkt-1)
                doknavn = left_filnavn
                else
                doknavn = oRec("filnavn")
                end if
                
                folderid = oRec("folderid")
                dokTxt = oRec("filertxt") 
                oRec.movenext
                wend
                oRec.close
                end if
        	
        	
	        '*** Henter gemt skabelon ***'
	        skabelonTxt = ""
	        if cint(skabelonid) <> 0 then
	        strSQL = "SELECT ts_txt FROM tilbuds_skabeloner WHERE ts_id = " & skabelonid
            oRec.open strSQL, oConn, 3
            while not oRec.EOF
            skabelonTxt = oRec("ts_txt") 
            oRec.movenext
            wend
            oRec.close
            end if
        	
        	
	       
            
            
        	
	        '*** Henter det job der skal printes ***
		        strSQL = "SELECT id, jobnavn, jobnr,"_
		        &" k.kkundenavn, k.kid, k.kkundenr, k.adresse, k.postnr, k.city, k.telefon, k.land, "_
		        &" jobknr, "_
		        &" jobTpris, jobstatus, jobstartdato, jobslutdato, "_
		        &" job.dato, job.editor, tilbudsnr, "_
		        &" fakturerbart, budgettimer, fastpris, kundeok, job.beskrivelse, "_
		        &" ikkeBudgettimer, tilbudsnr, serviceaft, kundekpers, jobans1, jobans2 "_
		        &" FROM job "_
		        &" LEFT JOIN kunder k ON (k.Kid = jobknr) "_
		        &" WHERE id = " & id 
		        oRec.open strSQL, oConn, 3
        				
        		
		        if not oRec.EOF then	
        				
			        kid = oRec("kid")	
			        strNavn = oRec("jobnavn")
			        strjobnr = oRec("jobnr")
			        tilbudsnr = oRec("tilbudsnr")
        			
			        strKnavn = oRec("kkundenavn")
			        strKundeId = oRec("jobknr")
			        strKundenr = oRec("kkundenr")
			        strAdresse = oRec("adresse")
			        strPostnr = oRec("postnr")
			        strBy = oRec("city")
			        strTLF = oRec("telefon")
			        strLand = oRec("land")
        			
			        strBudget = oRec("jobTpris")
			        if strBudget > 0 then
			        strBudget = strBudget
			        else
			        strBudget = 0
			        end if
        			
			        if oRec("ikkeBudgettimer") > 0 then
			        ikkeBudgettimer = oRec("ikkeBudgettimer")
			        else
			        ikkeBudgettimer = 0
			        end if
        			
			        if oRec("budgettimer") > 0 then
			        strBudgettimer = oRec("budgettimer")
			        else
			        strBudgettimer = 0
			        end if
        			
			        strTdato = oRec("jobstartdato")
			        strUdato = oRec("jobslutdato")
        			
			        strFastpris = oRec("fastpris")
        			
			        strBesk = oRec("beskrivelse")
			        'strtilbudsnr = oRec("tilbudsnr")
        			
			        'if request("FM_aftaler") = 1 then
			        'intServiceaft = oRec("serviceaft")
			        'else
			        'intServiceaft = 0
			        'end if
        			
			        intkundekpers = oRec("kundekpers")
        			
				        strSQL2 = "SELECT mnavn AS jobans_mnavn1, email, init, mnr FROM "_
				        &" medarbejdere WHERE mid = "& oRec("jobans1") 
				        oRec2.open strSQL2, oConn, 3 
				        while not oRec2.EOF 
        				
				        jobans1 = oRec2("jobans_mnavn1") '&" ("& oRec2("mnr") &") "
        				
				        if len(oRec2("init")) <> 0 then
				        jobans1 = jobans1 & " - " & oRec2("init")
				        end if
        				
				        if len(oRec2("email")) <> 0 then
				        jobans1 = jobans1 & ", " & oRec2("email")
				        end if
        				
			            oRec2.movenext
				        wend
				        oRec2.close
        				
				        strSQL3 = "SELECT mnavn AS jobans_mnavn2, email, init, mnr FROM "_
				        &" medarbejdere WHERE mid =  "& oRec("jobans2")
				        oRec2.open strSQL3, oConn, 3 
				        while not oRec2.EOF 
        				
				        jobans2 = oRec2("jobans_mnavn2") ' &" ("& oRec2("mnr") &") "
        				
				        if len(oRec2("init")) <> 0 then
				        jobans2 = jobans2 & " - " & oRec2("init")
				        end if
        				
				        if len(oRec2("email")) <> 0 then
				        jobans2 = jobans2 & ", " & oRec2("email")
				        end if
        				
				        oRec2.movenext
				        wend
				        oRec2.close
        				 
        			
		        end if
		        oRec.close
    
    else
		    lft = 20
	        tp = 50
	        

	        
	        
	        if media = "pdf" then
	        
	        if len(trim(request("pdfvalid"))) <> 0 then
	        pdfvalid = request("pdfvalid")
	        else
	        pdfvalid = 0
	        end if
	                
	                strSQLsel = "SELECT pdf_txt, pdf_footer FROM pdf_values WHERE id = "& pdfvalid
                    oRec.open strSQLsel, oConn, 3
                    if not oRec.EOF then
                    
                    indledningVal = oRec("pdf_txt")
                    afslutningsTXT = oRec("pdf_footer")
                    

                    end if
                    oRec.close
                    
                    
                    	                   
	        else
		    indledningVal = request("FM_vis_indledning")
    		end if
    		
    		if media <> "pdf" then
		    sdWdt = 800
		    else
		    sdWdt = 650
		    end if
		    
	end if '** Print
	
	
	if media <> "pdf" then
	tblWdt = "100%"
	else
	tblWdt = "100%"
	end if%>
	
	
	
	
	<div id="sindhold" style="position:absolute; left:<%=lft%>; top:<%=tp%>; visibility:visible; z-index:50; border:0px #999999 dashed; padding:0px; width:<%=sdWdt%>px;">
	
	    
	   
	    
	    
	    <%if func <> "print" then %>
	   
    <h4>Print job / tilbud (og gem som dok.)</h4>
    
    <a href="jobs.asp?menu=job&func=red&id=<%=id %>&int=1&rdir=<%=rdir%>"><< til rediger job</a>
      <table cellspacing=0 cellpadding=0 border=0 width="100%">
	      <form action="job_print.asp?kid=<%=kid %>" method="post">
     
	    <tr>
	    
	    <td valign=top>
    
        <%'** finder projegrupp rel **'
        if level <> 1 then
        call hentbgrppamedarb(session("mid"))
        strPgrpSQLkri = strPgrpSQLkri
        else
        strPgrpSQLkri = ""
        end if
    
        strSQLjob = "SELECT jobnavn, jobnr, j.id, kkundenavn, kkundenr FROM job j "_
	    &"LEFT JOIN kunder k ON (kid = j.jobknr) WHERE jobstatus = 1 OR jobstatus = 3 "& strPgrpSQLkri &" ORDER BY kkundenavn, jobnavn" 
	   
	    'Response.Write strSQLjob
	    'Response.flush
	   
	    %>
	    
	     <b>Vælg job ell. tilbud:</b>
            <select id="jobsel" name="id" style="width:500px;" onchange="submit();">
            <option value="0">Vælg..</option>
            <%
            oRec4.open strSQLjob, oConn,  3 
            while not oRec4.EOF 
            if cint(id) = oRec4("id") then
            jSel = "SELECTED"
            else
            jSel = ""
            end if
            
            %>
            <option value="<%=oRec4("id") %>" <%=jSel %>><%=oRec4("jobnavn") %> (<%=oRec4("jobnr") %>) ][ <%=oRec4("kkundenavn") %> (<%=oRec4("kkundenr") %>)</option>
            
            <%
            oRec4.movenext
            wend
            oRec4.close
            %>
            
                
            </select>
            &nbsp;
            <input id="Submit7" type="submit" value=" Vælg job >> " /><br />
            Kun job du har adgang til via dine projektgrupper, eller alle hvis du er administrator.
            
            </form>
            
            
            <!-- vælg job slut -->
            
            
            
            
            <%if id <> 0 then %>           
            <!-- Henter data --->
            
	    
	    
	  
	
	<%  
	    '** Afsender information ****'
		strSQLlogo = "SELECT f.filnavn, k.kid, k.logo, k.kkundenavn, k.kkundenr, k.adresse, "_
		&" k.postnr, k.city, k.telefon, k.land, k.fax, k.cvr, k.bank, k.regnr, k.kontonr FROM kunder k "_
		&" LEFT JOIN filer f ON (f.id = k.logo) WHERE k.useasfak = 1"
		
		'Response.Write strSQLlogo
		'Response.flush
		
		oRec.open strSQLlogo, oConn, 3
		if not oRec.EOF then
		
		afsKnavn = oRec("kkundenavn")
		afsAdr = oRec("adresse")
		afsPostnr = oRec("postnr")
		afsLand = oRec("land")
		afsBy = oRec("city")
		afsTlf = oRec("telefon")
		afsFax = oRec("fax")
		
		afsBank = oRec("bank")
		afsRegnr = oRec("regnr")
		afsKontonr = oRec("kontonr")
		afsCVR = oRec("cvr")
				
		end if
		oRec.close
		
		strkundeHTML = "<br><table cellspacing=0 cellpadding=0 border=0 width="""&tblWdt&"""><tr><td valign=top>"
		
		'** Kunde / Modtager **'
		
		strkundeHTML = strkundeHTML & strKnavn &"<br>"& strAdresse & "<br>"_
		& strPostnr &" "& strBy & "<br>" & strLand & "<br><br>"  
		
		if len(trim(intkundekpers)) <> 0 then
		intkundekpers = intkundekpers
		else
		intkundekpers = 0
		end if
 		
		strSQLkpers = "SELECT navn FROM kontaktpers WHERE id = "& intkundekpers 
	    oRec.open strSQLkpers, oConn, 3
	    while not oRec.EOF
	    
	    strkundeHTML = strkundeHTML & "Att:"& oRec("navn") & "<br><br>"
	    
	    oRec.movenext
	    wend
	    oRec.close 
	    
	    
	    strkundeHTML = strkundeHTML & "</td><td valign=bottom align=right>"
	    strkundeHTML = strkundeHTML & afsBy &", d. "& formatdatetime(date(),1) &"</td></tr></table>"
	    
	   
        'strjobHTML = "<br><br><table cellspacing=0 cellpadding=0 border=0 width="""&tblWdt&"""><tr><td valign=top>"
		
		
		strjobHTML = "<br><br>"
	    %>
	    
	    
	    
	    <%
	    'if tilbudsnr <> 0 then
		'strJobHTML = strjobHTML & "<h4>Vedr. "& vedrinfo &": "& tilbudsnr &"</h4>"
	    'else
	    '*** Jobnavn / vedr. text ***
		vedrtxtval = "<h4>"& vedrinfo &""
		strJobHTML = strJobHTML & vedrtxtval &" "& strNavn & " ("&strjobnr&")</h4>"
	    'end if 
	    
	    select case vedrinfo
	    case "Job"
	    jSel = "SELECTED"
	    case "Tilbud"
	    tSel = "SELECTED"
	    case "Rekvisition"
	    rSel = "SELECTED"
	    case "Budget"
	    bSel = "SELECTED"
	    case "Følgeseddel"
	    fSel = "SELECTED"
	    case ""
	    iSel = "SELECTED"
	    end select
		
		
		strChkjobHTML = "Overskrift:<SELECT name=""vedrinfo"">"_
		&"<option value=""Job"" "& jSel &">Job</option>"_
		&"<option value=""Tilbud"" "& tSel &">Tilbud</option>"_
		&"<option value=""Rekvisition"" "& rSel &">Rekvisition</option>"_
        &"<option value=""Følgeseddel"" "& fSel &">Følgeseddel</option>"_
		&"<option value=""Budget"" "& bSel &">Budget</option>"_
		&"<option value=""-"" "& iSel &">Ingen</option>"_
		&"</SELECT><h4>"& strNavn & " ("&strjobnr&")</h4>"
		
		
		if len(trim(strBesk)) <> 0 then
		strJobHTML = strJobHTML & "<b>Beskrivelse</b><br>"& strBesk & "<br><br>"
		end if
		
		
		strChkjobHTML = strChkjobHTML & "<b>Beskrivelse</b><br>"& left(strBesk, 200) & "<br>"
		
		
		'if lto <> "wowern" then
		strJobPrisHTML = strJobPrisHTML & "<h4>Pris: DKK "& formatnumber(strBudget, 2) &"</h4>"
		'end if
		
		'if prodtilbud = 2 then
		
		    if len(trim(jobans1)) <> 0 then
		    strJobInfoHTML = strJobInfoHTML & "Jobansvarlig/vor. ref.: " & jobans1 & "<br>"
            end if  
            
            if len(trim(jobans2)) <> 0 then
		    strJobInfoHTML = strJobInfoHTML & "Jobejer: " & jobans2 & "<br>"
            end if   
        
		strJobInfoHTML = strJobInfoHTML & "<br><b>Periode:<br></b>Arbejdet forventes udført i perioden "& formatdatetime(strTdato, 1) &" - "& formatdatetime(strUdato, 1) & "<br><br>&nbsp;"
		'end if
		
		      
        'strJobHTML = strJobHTML & strJobPrisHTML & strJobInfoHTML          
        
        
        if aktinfo_besk = "0" then
        aktinfo_beskCHK0 = "CHECKED"
        aktinfo_beskCHK1 = ""
        else
        aktinfo_beskCHK0 = ""
        aktinfo_beskCHK1 = "CHECKED"
        end if
        
        if aktinfo_priser = "0" then
        aktinfo_priserCHK0 = "CHECKED"
        aktinfo_priserCHK1 = ""
        else
        aktinfo_priserCHK0 = ""
        aktinfo_priserCHK1 = "CHECKED"
        end if
        
        if aktinfo_timer = "0" then
        aktinfo_timerCHK0 = "CHECKED"
        aktinfo_timerCHK1 = ""
        else
        aktinfo_timerCHK0 = ""
        aktinfo_timerCHK1 = "CHECKED"
        end if
        
        
        select case lto
        case "wowern"
        strAktHTML = ""
        case else
        strAktHTML = "<br><br><h4>Udspecificering:</h4>"
        end select
        
        strAktHTML = strAktHTML & "<table cellspacing=0 cellpadding=2 border=0 width="""&tblWdt&""">"
		strChkHTML = strAktHTML & "<tr><td colspan=3><b>Akt. indstillinger:</b><br>"_
		&"Vis akt. beskrivelser <input id=""Checkbox1"" name=""aktinfo_besk"" value=""0"" type=""radio"" "& aktinfo_beskCHK0 &" />ja <input id=""Checkbox1"" name=""aktinfo_besk"" value=""1"" type=""radio"" "& aktinfo_beskCHK1 &" />nej<br>"_
		&"Vis akt. priser <input id=""Checkbox1"" name=""aktinfo_priser"" value=""0"" type=""radio"" "& aktinfo_priserCHK0 &" />ja <input id=""Checkbox1"" name=""aktinfo_priser"" value=""1"" type=""radio"" "& aktinfo_priserCHK1 &" />nej<br>"_
		&"Vis akt. timer <input id=""Checkbox1"" name=""aktinfo_timer"" value=""0"" type=""radio"" "& aktinfo_timerCHK0 &" />ja <input id=""Checkbox1"" name=""aktinfo_timer"" value=""1"" type=""radio"" "& aktinfo_timerCHK1 &" />nej<br>&nbsp;"_
		&"</td></tr>"
		
	
			strSQLselAkt = "SELECT id AS aktid, navn, beskrivelse, job, fakturerbar,  "_
			&" aktstartdato, aktslutdato, editor, dato, budgettimer, aktfavorit, aktstatus, fomr, faktor, aktbudget, fase, bgr, antalstk, aktbudgetsum "_
			&" FROM aktiviteter WHERE job =" & id &" ORDER BY fase, sortorder, navn"
			
			oRec.open strSQLselAkt, oConn, 3
			
			         
			         
			         if aktinfo_timer = "0" then
				     cpsF = 2
				     totPDright = 0
				     else
				     cpsF = 1
				        if media <> "pdf" then
				        totPDright = 35
				        else
				        totPDright = 0
				        end if
				     end if
			
			af = 0
			a = 0
			lastFase = ""
			sumTotFase = 0
			aktbudgetTot = 0
			while not oRec.EOF 
				
				'select case lto 
				'case "wowern", "intranet - local"
				'strNavn = oRec("navn")
				
				'if len(strNavn) > 3 then
				'len_strNavn = len(strNavn)
				'right_strNavn = right(strNavn, len_strNavn - 3)
				'strNavn = right_strNavn 
				'end if
				
				'case else
				strNavn = oRec("navn")
				'end select
				
				strBeskrivelse = oRec("beskrivelse")
				strTdato = oRec("aktstartdato")
				strUdato = oRec("aktslutdato")
				strFakturerbart = oRec("fakturerbar")
				useBudgettimer = oRec("budgettimer")
				intaktstatus = oRec("aktstatus")
			    aktbudget = oRec("aktbudget")
			    aktbudgetsum = oRec("aktbudgetsum")
			    aktbgr = oRec("bgr")
			    aktAntalStk = oRec("antalstk")
			    strFase = oRec("fase")
				
				
				if lcase(lastFase) <> lcase(strFase) AND len(trim(strFase)) <> 0 then
				 
				 if instr(vlgFase, "#"&lcase(strFase)&"#") <> 0 OR (vlgFase = "-1" AND sumTotFase <> 0) then
				    
				    '*** Sum tot. på fase ***'
				    if a <> 0 AND aktinfo_priser = "0" then
				    strAktHTML = strAktHTML & "<tr><td colspan="&cpsF&"><span class=""jobpr16""><b>Ialt:</b></span></td><td align=right valign=top style=""width:40px;""><span class=""jobpr16""><b>DKK</b></span></td><td align=right valign=top style=""width:80px;""><span class=""jobpr16""><b>"&formatnumber(sumTotFase, 2)&"</b></span></td></tr>"
				    sumTotFase = 0
				    end if
				    
				    '*** Ny fase ***'
				    if vlgFase = "-1" then '(ved første indlæsning af side)
				    antal_aktifase = 0
				    strSQLaktifase = "SELECT COUNT(*) AS antal_aktifase FROM aktiviteter WHERE job = "& id &" AND fase = '"& strFase & "' AND aktbudgetsum <> 0 GROUP BY fase"
				    oRec3.open strSQLaktifase, oConn, 3
				    if not oRec3.EOF then
				    antal_aktifase = oRec3("antal_aktifase")
				    end if
				    oRec3.close
				    
				        if antal_aktifase <> 0 then
				          strAktHTML = strAktHTML & "<tr><td colspan=5><br><br><span class=""jobpr18""><b>"&replace(strFase, "_", " ")&"</b></span></td></tr>"
				          fsCHK = "CHECKED"
				        else
				          fsCHK = ""
				        end if
				    
				    else
				    
				    strAktHTML = strAktHTML & "<tr><td colspan=5><br><br><span class=""jobpr18""><b>"&replace(strFase, "_", " ")&"</b></span></td></tr>"
				    fsCHK = "CHECKED"
				    end if
				    
				   
				    
				 else
				 fsCHK = ""
				 end if
				 
				strChkHTML = strChkHTML & "<tr><td colspan=5><br><br><input id="&lcase(strFase)&" name=""FM_vis_fase"" "&fsCHK&" value=""#"&lcase(strFase)&"#"" type=""checkbox"" class=""faseoskrift"" /><b>"&replace(strFase, "_", " ")&"</b></td></tr>"    
				af = 0
				end if 
				
				
				if instr(vlgAkt, "#"& oRec("aktid") &"#") <> 0 OR (vlgAkt = "-1" AND aktbudgetsum <> 0) then 
				strAktHTML = strAktHTML & "<tr><td style=""padding-right:20px; width:400px;"">"
				end if
				
				select case right(af, 1)
				case 0,2,4,6,8
				bgcol = "#D6DFf5"
				case else
				bgcol = "#ffffff"
				end select 
				
				strChkHTML = strChkHTML & "<tr style=""background-color:"& bgcol &";""><td style=""padding-right:20px; width:400px;"" class=lille>" 
				
				'** CHKbokse
				if instr(vlgAkt, "#"& oRec("aktid") &"#") <> 0 OR (vlgAkt = "-1" AND aktbudgetsum <> 0) then
				cbCHK = "CHECKED"
				else
				cbCHK = ""
				end if
				
				strAktcbox = "<input id=""af_"&lcase(strFase)&"_"&af&""" name=""FM_vis_akt"" type=""checkbox"" "&cbCHK&" value='#"& oRec("aktid") &"#' class=""aktoskrift""/>&nbsp;"
				
				strChkHTML = strChkHTML & strAktcbox
				
				if instr(vlgAkt, "#"& oRec("aktid") &"#") <> 0 OR (vlgAkt = "-1" AND aktbudgetsum <> 0) then 
				strAktHTML = strAktHTML & "<span class=""jobpr16"">"& strNavn & "</span>"
				end if
				
				strChkHTML = strChkHTML & "<b>" & strNavn & "</b>"
				
				
				
				
				if len(trim(strBeskrivelse)) <> 0 then
				
				
				
				if aktinfo_besk = "0" then
				    
				    if instr(vlgAkt, "#"& oRec("aktid") &"#") <> 0 OR (vlgAkt = "-1" AND aktbudgetsum <> 0) then 
				    strAktHTML = strAktHTML & "<br /><span class=""jobpr14"">"& strBeskrivelse & "</span><br>&nbsp;" 'htmlparseCSVtxt 
				    end if
				
				end if
				     
                call htmlparseCSV(strBeskrivelse)
				  
				strChkHTML = strChkHTML & "&nbsp;<font class=lillesort>"& left(htmlparseCSVtxt, 80) & "</font>"
				end if
				
				if instr(vlgAkt, "#"& oRec("aktid") &"#") <> 0 OR (vlgAkt = "-1" AND aktbudgetsum <> 0) then 	
				strAktHTML = strAktHTML & "</td>"
				end if	
				strChkHTML = strChkHTML & "</td>"
				
				'** Udvideet visning ***'	
				'if prodtilbud = 2 then	
				'strAktHTML = strAktHTML & "<td class=lille>"& formatdatetime(strTdato, 2) & " til<br> "& formatdatetime(strUdato, 2) &"</td>"       
				'end if	
				
				
				call akttyper2009Prop(oRec("fakturerbar"))
				
				if aktbgr <> 2 then
				
				'** sætter r på time **'
				if useBudgettimer > 1 AND aty_enh = "time" then
				aty_enh = aty_enh & "r"
				useAntal = useBudgettimer
				end if
				
				else
				aty_enh = "stk."
				useAntal = aktAntalStk
				end if
				
				if aktinfo_timer = "0" then
				
				    if instr(vlgAkt, "#"& oRec("aktid") &"#") <> 0 OR (vlgAkt = "-1" AND aktbudgetsum <> 0) then 
    				
				            strAktHTML = strAktHTML & "<td align=right valign=top style=""padding-right:30px;""><span class=""jobpr16"">"& formatnumber(useAntal, 2) &" "& aty_enh &" á "& formatnumber(aktbudget, 2) &"</span></td>"        
				            
            	    end if
				
				end if
				
				if aktinfo_priser = "0" then
				
				if instr(vlgAkt, "#"& oRec("aktid") &"#") <> 0 OR (vlgAkt = "-1" AND aktbudgetsum <> 0) then 
				strAktHTML = strAktHTML & "<td align=right valign=top style=""width:40px;"">DKK</td><td align=right valign=top style=""width:80px;""><span class=""jobpr16"">"& formatnumber(aktbudgetsum, 2) &"</span></td></tr>"        
				sumTotFase = sumTotFase + aktbudgetsum
				end if
				
				end if
				
				
				
				strChkHTML = strChkHTML & "<td align=right valign=top style=""white-space:nowrap class=lille>"& formatnumber(useAntal, 2) &" "& aty_enh &"</td>"        
				
				strChkHTML = strChkHTML & "<td align=right valign=top class=lille>DKK</td><td align=right valign=top class=lille> "& formatnumber(aktbudgetsum, 2) &"</td></tr>"
				
				
				if instr(vlgAkt, "#"& oRec("aktid") &"#") <> 0 OR (vlgAkt = "-1" AND aktbudgetsum <> 0) then 
				aktbudgetTot = aktbudgetTot + aktbudgetsum
				end if
					     
			a = a + 1
			af = af + 1	
			lastFase = strFase	
			oRec.movenext
			wend
			oRec.close
			
			if aktinfo_priser = "0" then
			
			        '*** Sum tot. på fase ***'
				    if a <> 0 AND sumTotFase <> 0 then
				     strAktHTML = strAktHTML & "<tr><td colspan="&cpsF&"><span class=""jobpr16""><b>Ialt:</b></span></td><td align=right valign=top style=""width:40px;""><span class=""jobpr18""><b>DKK</b></span></td><td align=right valign=top style=""width:80px;""><span class=""jobpr18""><b>"&formatnumber(sumTotFase, 2)&"</b></span></td></tr>"
				    sumTotFase = 0
				    end if
			
			
			
			strAktHTML = strAktHTML & "<tr><td colspan=5 align=right style=""padding-top:40px; padding-right:"&totPDright&"px;""><span class=""jobpr18"">Ialt ekskl. moms, <b>DKK "& formatnumber(aktbudgetTot, 2) &"</b></span></td></tr>"        
			end if
			
			strAktHTML = strAktHTML & "</table><br><br>&nbsp;"
		    
		
		    strChkHTML = strChkHTML & "<tr><td colspan=5 align=right style=""padding-top:40px;"" class=lille>Ialt ekskl. moms, <b>DKK "& formatnumber(aktbudgetTot, 2) &"</b></td></tr>"        
			strChkHTML = strChkHTML & "</table><br><br>&nbsp;"
		
		    
		    
			
	    'if prodtilbud = 2 then
		        strMileHTML = "<br><br><h4>Milepæle:</h4><table cellspacing=0 cellpadding=0 border=0 width=""100%"">"
		
		        strSQL = "SELECT navn, milepal_dato, jid, beskrivelse FROM milepale WHERE jid = "& id & " ORDER BY milepal_dato"
		        oRec.open strSQL, oConn, 3 
		        x = 0
		        while not oRec.EOF 
		        strMileHTML = strMileHTML &"<tr><td valign=top>"& formatdatetime(oRec("milepal_dato"), 2) & ": "& oRec("navn") & "<br>"& oRec("beskrivelse") & "<br><br></td></tr>"
		        
		        x = x + 1
		        oRec.movenext
		        wend
		        oRec.close 
		      
		      strMileHTML = strMileHTML &"</table>"
		
		
		strHilsenHTML =  "<table cellspacing=0 cellpadding=0 border=0 width=""100%""><tr><td style=""padding-top:10px;""><span class=""jobpr14"">Med venlig hilsen <br><b>"& afsKnavn &"</b><br><br>"& session("user") &"</span></td></tr></table>"        
		    
		
		'end if
		
            
            
            
   
    '******************************************
    '***** Filter ****************************
    '*******************************************
          
    tTop = 25
	tLeft = 0
	tWdth = 1104
	
	
	call tableDiv(tTop,tLeft,tWdth)
	%>
	
                    
                    <table cellspacing=0 cellpadding=0 border=0 width=100%>
                    
	    
                    <tr><td valign=top> 
                    
                    <%
                          tdheight = 450
                          ptop = 0
                          pleft = 0
                          pwdt = 360
            
                         call filteros09(ptop, pleft, pwdt, "Hent skabelon ell. gemt dokument..", 3, tdheight)
                        
                         %>
                    
                    
                    <%if len(request("ignorerKid")) <> 0 then 
                    igKidSQL = " AND ts_kundeid <> 0"
                    chkigKid = "CHECKED"
                    ignkidVal = 1
                    else
                        
                        if request.Cookies("tsa")("ignkidval") <> "" then
                        ignkidVal = request.Cookies("tsa")("ignkidval")
                        
                        if ignkidVal = 0 then
                        chkigKid = ""
                        igKidSQL = " AND ts_kundeid = "& kid
                        else
                        igKidSQL = " AND ts_kundeid <> 0"
                        chkigKid = "CHECKED"
                        end if
                        
                        else
                        ignkidVal = 0 
                        chkigKid = ""
                        igKidSQL = " AND ts_kundeid = "& kid
                        end if
                    
                    end if
                    
                    response.Cookies("tsa")("ignkidval") = ignkidVal
                    
                    level = session("rettigheder")
                    
                    if level <= 2 OR level = 6 then%>
                    </td>
                    </tr>
                    <tr>
                    <form action="job_print.asp?func=sletskabelon&id=<%=id %>&kid=<%=kid %>" method="post">
     
                    <td align=right style="padding:5px 50px 0px 0px;">
                    
                    
                   
	            <input id="sletskabelonid" name="sletskabelonid" value="0" type="hidden" />
        <input id="Submit3" type="submit" value="Slet Skabelon >>" style="font-size:9px; background-color:#cccccc;" />
	    </form>
	        </td>
	                
	    </tr>
	     <tr>
	     
	      <form action="job_print.asp?id=<%=id%>&kid=<%=kid %>" method="post">
           
	        <td valign=top style="padding:0px 10px 10px 10px;">
                    
                    
                    
                    <input id="Checkbox1" name="ignorerKid" value="1" type="checkbox" <%=chkigKid %> /> Vis alle (ignorer tilhørsforhold til kunde)<br />
                    <%
                    end if
                    
                    strSQL_skab = "SELECT ts_navn, ts_id, ts_dato, ts_txt FROM tilbuds_skabeloner WHERE ts_id <> 0 "& igKidSQL &" ORDER BY ts_navn"
                    
                    %>
                    <br />
                    <b>Skabeloner:</b><br />
                    Skabelonnavn (dato)<br />
                    <select id="tb_skabeloner" name="tb_skabeloner" size=4 style="width:325px;" onchange="setskabid()">
                      
                    
                    <%
                    oRec4.open strSQL_skab, oConn, 3
                    while not oRec4.EOF 
                    
                    if cint(skabelonid) = oRec4("ts_id") then
                    tSel = "SELECTED"
                    else
                    tSel = ""
                    end if 
                    
                    
                    %>
                    <option value="<%=oRec4("ts_id") %>" <%=tSel %>><%=oRec4("ts_navn") %> (<%=oRec4("ts_dato") %>)</option>
                    
                    
                       
                    <%
                    oRec4.movenext
                    wend
                    oRec4.close
                       %>
                        </select>
                        
                        <br /><br />
                        <input id="Submit5" type="submit" value=" Hent skabelon >> " />
                        
                        </td>
                        </form>
                        </tr>
                        
                        
                         
                    
                        <tr>
                        <form action="job_print.asp?id=<%=id%>&kid=<%=kid %>" method="post">
                        <td valign=top style="padding:10px;"><br /><b>Dokumenter oprettet på dette job:</b> (kun .txt filer) <a href="filer.asp?nomenu=1&jobid=0&kundeid=<%=kid %>" class=vmenu target=_blank>Se filarkiv >></a> <br />
                        Folder ][ filnavn (dato)<br />
                        
                        <%strSQLdok = "SELECT f.filnavn, f.id AS fid, f.dato, fo.navn AS foldernavn FROM filer f "_
                        &" LEFT JOIN foldere fo ON (fo.id = f.folderid) WHERE f.jobid =" & id & " ORDER BY foldernavn, filnavn" 
                        
                        'Response.Write strSQLdok
                        'Response.flush
                        
                        %>
                            <select id="Select1" name="FM_dokument" size=4 style="width:325px; font-size:9px;">
                                
                            
                        <%
                        oRec.open strSQLdok, oConn, 3
                        while not oRec.EOF
                        
                        if cint(dokid) = oRec("fid") then
                        tSel = "SELECTED"
                        else
                        tSel = ""
                        end if 
                        %>
                        <option value="<%=oRec("fid") %>" <%=tSel %>><%=oRec("foldernavn") %> ][ <%=oRec("filnavn") %> (<%=oRec("dato") %>)</option>
                        <%
                        oRec.movenext
                        wend
                        oRec.close
                        %>
                        </select>
                        <br /><br />
                            <input id="Submit4" type="submit" value=" Hent dokument >> " />
                        </td>
                        
                        </tr>
                        
                        </form>
                        </table>
                        
                        
                        </td>
                        <form action="job_print.asp?id=<%=id%>&kid=<%=kid %>&func=gem" method="post">
                        <input id="FM_gem_txt" name="FM_gem_txt" type="hidden" />
                        <td valign=top>
                        
   
	
                    
                    
                 
                    
                    <%
                          
                          ptop = 0
                          pleft = 0
                          pwdt = 450
            
                         call filteros09(ptop, pleft, pwdt, "Gem dokument  i filarkiv?", 3, tdheight)
                        
                         %>
                    
                   
                    <input id="gemsomskabelon_kid" name="gemsomskabelon_kid" value="<%=kid %>" type="hidden" />
                    <input id="gemsomskabelon" name="gemsomskabelon" type="checkbox" value="1"/> Gem som skabelon<br />
                    
                    <%'if cint(dokid) <> 0 then
                    dkCHK = "CHECKED"
                    'else
                    'dkCHK = ""
                    'end if
                     %>
                    <input id="gemsomfil" name="gemsomfil" type="checkbox" value="1" <%=dkCHK %> /> Gem som dokument (overskriver ved rediger eksisterende dok.)
                    <br /><br />
                    <b>Hvis du ønsker at gemme som dokument i fil-arkivet skal du vælge en folder:</b><br />
                    Der kan oprettes standard foldere i kontrolpanelet.
                     
                    <br />
                    <a href="javascript:popUp('filer.asp?func=oprfo&kundeid=<%=kid%>&jobid=<%=id%>','600','500','250','120');" target="_self" class=vmenu><img src="../ill/folder_ikon.gif" width="17" height="15" alt="" border="0">&nbsp;Opret folder >></a> (re-loader side!)
                    <br />
                    
                    <br />
                    Folder(e) (standard foldere på kunde eller med tilknytning til dette job)
                   
                    <%
                    strSQLfol = "SELECT f.id as foid, f.navn AS fonavn, f.jobid, j.jobnavn, j.jobnr FROM foldere f "_
                    &" LEFT JOIN job j ON (j.id = f.jobid) WHERE f.jobid = " & id & " OR f.id = 1000 OR (f.kundeid = "& kid &" AND f.jobid = 0) ORDER BY f.navn"
                    
                     
                    %>
                    
                    <br /> <select id="foid" name="foid" size="10" style="width:400px; font-size:9px;">
                    <%
                    oRec3.open strSQLfol, oConn, 3
                    while not oRec3.EOF 
                    
                    if cint(folderid) = oRec3("foid") then
                    fSel = "SELECTED"
                    else
                    fSel = ""
                    end if
                    %>
                    <option value="<%=oRec3("foid") %>" <%=fSel %>><%=oRec3("fonavn") %> 
                   
                    </option>
                    <%
                    oRec3.movenext
                    wend
                    oRec3.close
                    %>
                    </select>
		           <br />
                   <br /><b>Navn:</b> <input name="skabelonnavn" id="skabelonnavn" value="<%=doknavn%>" style="width:200px;" type="text" />.txt
                   <input id="dokid2" name="FM_dokument" value="<%=dokid %>" type="hidden" />
                   &nbsp;&nbsp;<input onclick="gemtxt();" id="Submit2" type="submit" value=" Gem >> " />
                   
                   </td>
                   </tr>
                   
                   </form>
                   
                   </table>
                   </div>
                  
                        
                        
                        
                        </td>
                        
                        <td valign=top>
                        <%
                        
                          ptop = 0
                          pleft = 0
                          pwdt = 274
                         
            
                         call filteros09(ptop, pleft, pwdt, "Eksport og Print", 2, tdheight)
                        
                         %>
                        
        
        <!--                
		<a href="job_print.asp?&id=<=id%>&func=print&kid=<=kid%>" class=vmenu>Print</a>
		-->
		
		<table cellspacing=0 cellpadding=0 border=0>
		<!-- https://outzource.dk/timeout_xp/wwwroot/ver2_1/pdf/ -->
		
		<form action="job_make_pdf.asp?lto=<%=lto%>&id=<%=id%>&kid=<%=kid%>" target=_blank method="post">
		<tr>
		    <td><input onclick="gemtxtPdf();" id="Submit8" type="submit" style="font-size:9px;" value=" PDF " />
            <input id="FM_vis_indledning_pdf" name="FM_vis_indledning_pdf" type="hidden" /></td>
            <input id="FM_vis_afslutning_pdf" name="FM_vis_afslutning_pdf" type="hidden" /></td>
            </form>
            <form action="job_print.asp?id=<%=id%>&func=print&kid=<%=kid %>" method="post" target=_blank>
	    
		    <td>&nbsp;&nbsp;|&nbsp;&nbsp;</td>
		    <td><!-- Main Print Form -->
                <input id="Submit6" type="submit" value=" Print " style="font-size:9px;" />
               
                </td>
		</tr>
		<tr><td colspan=3><br />
                
                <%if lto = "dencker" OR lto = "intranet - local" then
                gfpCHK = "CHECKED"
                else
                gfpCHK = ""
                end if %>
                <input id="Checkbox2" name="printgemfil" value="1" type="checkbox" <%=gfpCHK %> /> Gem fil i filarkiv ved print.
                <br /><span style="background-color:pink;"><b>Angiv filnavn og folder til venstre</b></span>
                </td></tr>
		</table>
		
		
		
		
		
		
		
		
		<br /><br />
		
	
		<b>Gem PDF ovenfor, og email faktura til:</b><br />
		<%
		strSQL = "SELECT navn, email FROM kontaktpers WHERE kundeid = " & kid
		'Response.Write strSQL 
		'Response.Flush
		
		oRec.open strSQL, oConn, 3
        while not oRec.EOF
        
        Response.Write "<i>"& oRec("navn") & "</i>, <a href='mailto:"&oRec("email")&"&subject=Vedr.: "& strjobnr&"_"&tilbudsnr &"' class=vmenu>" & oRec("email") & "</a><br>"
        
        oRec.movenext
        wend
        oRec.close
        
        %>
        <br />Kontaktpersoner oprettes under fanebladet "kontakter" i hovemenuen.
        </td>
        </tr>
        </table>
        </div>
		                
                        </td>
                          </tr>
                          
                          
                          <!-- table div -->
                          </table>  
                           </div>
                          
                          
                          
                          <br /><br /> <br /><br />
                          <h4>Jobinfo og aktiviteter hentet fra det valgte job / skabelon:</h4>
                          <table width=1104 cellpadding=0 cellspacing=5 border=0>
                          <tr>
                            <td valign=top>
                                                             
                        
                    <%
		                dim content
		                
		                if kinfo <> "1" then
		                strkundeHTMLedit = strkundeHTML
		                else
		                strkundeHTMLedit = ""
		                end if
		                
		                if jbinfo <> "1" then
		                strJobHTMLedit = strJobHTML
		                else
		                strJobHTMLedit = ""
		                end if
		                
		                if cint(jbprisinfo) <> 1 then
		                strJobPrisHTMLedit = strJobPrisHTML
		                else
		                strJobPrisHTMLedit = ""
		                end if
		                
		                 if jbperansinfo <> "1" then
		                jbperansinfoHTMLedit = strJobInfoHTML
		                else
		                jbperansinfoHTMLedit = ""
		                end if
		                
		                '*** Styres pr. akt **'
		                'if aktinfo <> "1" then
		                aktinfoHTMLedit = strAktHTML
		                'else
		                'aktinfoHTMLedit = ""
		                'end if
		                
		                if cint(mileinfo) <> 1 then
		                mileinfoHTMLedit = strMileHTML
		                else
		                mileinfoHTMLedit = ""
		                end if
		                
		                
		                
		                
	                    content = strkundeHTMLedit & strJobHTMLedit & strJobPrisHTMLedit & jbperansinfoHTMLedit & aktinfoHTMLedit & mileinfoHTMLedit & strHilsenHTML
            			
			            
			            Set editorK = New CuteEditor
            					
			            editorK.ID = "FM_vis_indledning"
			            
			            
			            if cint(skabelonid) <> 0 then
			            editorK.Text = skabelonTxt
			            else
			                if cint(dokid) <> 0 then
			                editorK.Text = dokTxt
			                else
			                editorK.Text = content
			                end if
			            end if
			            
			            editorK.FilesPath = "CuteEditor_Files"
			            editorK.AutoConfigure = "default"
            			'editor.ConfigurationPath = "CuteEditor_Files/Configuration/myTools_to.config"
			            editorK.Width = 790
			            editorK.Height = 1080
			            editorK.Draw()
		                %>
		                
		                <br />
		                For at lave et sideskift indtastes et ##.
		                <br />
		                 <br />
                    <br /><br /><b>Fast footer på print:</b> (vises i bunden af hver printet side, kontakt OutZourCE for opsætning m. eget logo mv.)&nbsp;
                    <br />
                    <textarea id="FM_vis_afslutning" name="FM_vis_afslutning" cols="97" rows="4"><%=afslutningVal %></textarea>
                    <br /><input id="FM_gem_afs" name="FM_gem_afs" type="checkbox" value="1" />Gem denne footer på alle kunder. (overskriver eksisterende)<br />
                    <br /><input id="Submit1" type="submit" value=" Gem / print >> " />
                    
		                
		                </td>
		                </form>
		                <form action="job_print.asp?id=<%=id%>&kid=<%=kid %>" method="post">
		                <td valign=top style="padding:2px; border:1px #cccccc solid; background-color:#FFFFFF;">
		                
		                 <%
                          tdheight = 1200
                          ptop = 0
                          pleft = 0
                          pwdt = 300
            
                         call filteros09(ptop, pleft, pwdt, "Overblik / data valg", 4, tdheight)
                        
                         %>
		                
		                
		                <b>Jobinfo og aktiviteter på det aktuelle job / tilbud</b>
		                Hvis der vælges en gemt skabelon eller dok. kan original data på jobbet ses her.<br /><br />
		                Desuden kan du vælge hvilke data der skal vises når du henter et job/tilbud til print, eller du kan gen-indlæse det aktuelle job/tilbud med nedenstående valg:<br /><br />
		               
		               <%if kinfo <> "1" then
		               kinfoCHK0 = "CHECKED"
		               kinfoCHK1 = ""
		               else
		               kinfoCHK0 = ""
		               kinfoCHK1 = "CHECKED"
		               end if %>
		                
		                    <b style="background-color:#CCCCCC;"><u>Vis kunde info / adr.</u></b>
                            <input id="Radio1" name="kinfo" type="radio" value="0" <%=kinfoCHK0 %> /> ja <input id="Radio1" name="kinfo" type="radio" value="1" <%=kinfoCHK1 %> /> nej
		                <%=strkundeHTML%>
		                
		                
		                <%if jbinfo <> "1" then
		               jbinfoCHK0 = "CHECKED"
		               jbinfoCHK1 = ""
		               else
		               jbinfoCHK0 = ""
		               jbinfoCHK1 = "CHECKED"
		               end if %>
		                <br /><br />
		                 <b style="background-color:#CCCCCC;"><u>Vis job- navn og -beskrivelse</u></b>
		                  <input id="Radio3" type="radio" name="jbinfo" value="0" <%=jbinfoCHK0 %> /> ja <input id="Radio3" name="jbinfo" type="radio" value="1" <%=jbinfoCHK1 %> /> nej
		                <%=strChkjobHTML%>
		                
		                
		                <%if cint(jbprisinfo) <> 1 then
		               jbprisinfoCHK0 = "CHECKED"
		               jbprisinfoCHK1 = ""
		               else
		               jbprisinfoCHK0 = ""
		               jbprisinfoCHK1 = "CHECKED"
		               end if %>
		               <br />
		                <b style="background-color:#CCCCCC;"><u>Vis samletpris, stor.</u></b>
		                   <input id="Radio2" type="radio" name="jbprisinfo" value="0" <%=jbprisinfoCHK0 %> /> ja <input id="Radio4" name="jbprisinfo" type="radio" value="1" <%=jbprisinfoCHK1 %> /> nej
		               <%=strJobPrisHTML%>
		                
		                
		                <%if jbperansinfo <> "1" then
		               jbperansinfoCHK0 = "CHECKED"
		               jbperansinfoCHK1 = ""
		               else
		               jbperansinfoCHK0 = ""
		               jbperansinfoCHK1 = "CHECKED"
		               end if %>
		              
		                <b style="background-color:#CCCCCC;"><u>Vis periode og job ansvarlige</u></b>
		                     <input id="Radio5" type="radio" name="jbperansinfo" value="0" <%=jbperansinfoCHK0 %> /> ja <input id="Radio6" name="jbperansinfo" type="radio" value="1" <%=jbperansinfoCHK1 %> /> nej
		              <br />
		                <%=strJobinfoHTML%>
		                
		                
		                
		                  <%if aktinfo <> "1" then
		               aktinfoCHK0 = "CHECKED"
		               aktinfoCHK1 = ""
		               else
		               aktinfoCHK0 = ""
		               aktinfoCHK1 = "CHECKED"
		               end if %>
		                <br /><br />
		                <b style="background-color:#CCCCCC;"><u>Vis aktiviteter</u></b>
		                <input id="aktinfo_0" type="radio" name="aktinfo" value="0" <%=aktinfoCHK0 %> /> alle <input id="aktinfo_1" name="aktinfo" type="radio" value="1" <%=aktinfoCHK1 %> /> ingen
		              
		                <%=strChkHTML%>
		                
		                
		                
		                
		                 <%if cint(mileinfo) <> 1 then
		               mileinfoCHK0 = "CHECKED"
		               mileinfoCHK1 = ""
		               else
		               mileinfoCHK0 = ""
		               mileinfoCHK1 = "CHECKED"
		               end if %>
		                  <br /><br />
		                <b style="background-color:#CCCCCC;"><u>Vis milepæle</u></b>
		                 <input id="Radio9" type="radio" name="mileinfo" value="0" <%=mileinfoCHK0 %> /> ja <input id="Radio10" name="mileinfo" type="radio" value="1" <%=mileinfoCHK1 %> /> nej
		              
		                <%=strMileHTML %>
		                
		                
		                <%=strHilsenHTML %>
		                
		                <!-- filteros09 slut -->
		                </td>
		                </tr>
		                </table>
		                
                        </div>
		                
		                <table width=100% cellpadding=0 cellspacing=0 border=0>
		                 <tr><td align=center><br />
                     <input id="Submit9" type="submit" value=" Gen-indlæs >> " />
		                
                    </td></tr>
                    </form>
                    </table>
		                
		                </td>
                    
                    </tr>
                   
                    </table>
                   
                   
                    
                    <br /><br /><br />&nbsp;
                    
                <%end if 'id <> 0 %>   
                    
        <%else %>
        
        
        
        
            <%
            
            
            '*** sideombrydning ***
            if len(indledningVal) <> 0 then 
            
            
            if instr(indledningVal, "##") <> 0 then
            
            s1 = instr(indledningVal, "##") - 1
            p1_indledningVal = left(indledningVal, s1)
            len_p1val = len(p1_indledningVal)
            len_indledningVal = len(indledningVal)
            p2_indledningVal = ""
            p2_indledningVal = p2_indledningVal & "<img src=""../ill/blank.gif"" height=1 width=1 style='page-break-before:always'><br><br>"& right(indledningVal, len_indledningVal-(len_p1val+2))
            
            x = 2
            else
            
            p1_indledningVal = indledningVal
            p2_indledningVal = ""
            
            x = 1
            
            end if
            
            
            for p = 1 to x 
            
            
            if p = 1 then
            Response.Write p1_indledningVal
            else
            Response.Write p2_indledningVal
            end if 
            
            
            
            select case lto
            case "wowern", "intranet - local"
                if media <> "pdf" then
                
                    select case p
                    case 1
                    prdtop = "950"
                    prdlft = "75"
                    case 2
                    prdtop = "2075"
                    prdlft = "75"
                    end select
                    
                else
                
                    select case p
                    case 1
                    prdtop = "800"
                    prdlft = "0"
                    case 2
                    prdtop = "2075"
                    prdlft = "0"
                    end select
                
                end if
            
            case "outz"    
            
             if media <> "pdf" then
                
                    select case p
                    case 1
                    prdtop = "900"
                    prdlft = "100"
                    case 2
                    prdtop = "2025"
                    prdlft = "100"
                    end select
                    
                else
                
                    select case p
                    case 1
                    prdtop = "800"
                    prdlft = "0"
                    case 2
                    prdtop = "1975"
                    prdlft = "0"
                    end select
                
                end if
                
            case "essens"
            prdtop = "950"
            prdlft = "0"
            case else
            prdtop = "950"
            prdlft = "0"
            end select
            
            %>
            <div style="position:absolute; top:<%=prdtop%>px; left:<%=prdlft%>px; border:0px #999999 dashed; padding:0px;">
            <center>
            <%=afslutningsTXT %>
            </center>
            </div>
            <br /><br />&nbsp;
            
            <%next %>
            
           
            
            
            <%end if %>
            
            
            
        
        <%end if %> 
                  
                  
        
		
		
		<%
		
	    if showfiler = 99999 then
		
		if len(trim(request("FM_vis_filer"))) <> 0 AND request("FM_vis_filer") <> 0 then
		visfiler = 1
		else
		visfiler = 0
		end if
		
		if cint(visfiler) = 1 OR func <> "print" then %>
		                <br><br>
		                <h4><input id="FM_vis_filer" name="FM_vis_filer" type="checkbox" value="1" />Filer:</h4>
		                <%
                		
                		
	                        jobidSQL = "AND fo.jobid = " & id
	                        kundeseSQL = ""
	                        kundeseFilSQL = ""
	                        gamleFilerKri = ""
	                        kundeid = strKundeId
                	
                	
                	
		                strSQL = "SELECT fo.kundeid AS kundeid, fo.navn AS foldernavn, "_
		                &" fi.adg_kunde, fi.adg_admin, fi.adg_alle, fo.id AS foid, fo.kundese, "_
		                &" fo.jobid AS jobid, filnavn, fi.id AS fiid, fi.dato, jobnr, jobnavn, jobans1, jobans2, fo.dato AS fodato"_
		                &" FROM foldere fo "_
		                &" LEFT JOIN filer fi ON (fi.folderid = fo.id "& kundeseFilSQL &") "_
		                &" LEFT JOIN job j ON (j.id = fo.jobid) "_
		                &" WHERE "& gamleFilerKri &" fo.kundeid = "& kundeid &" "& jobidSQL &" "& kundeseSQL &" ORDER BY foid, filnavn"
                		
		                'Response.write strSQL
		                'Response.flush
                		
		                oRec.open strSQL, oConn, 3 
		                x = 0
		                lastfoid = 0
		                while not oRec.EOF 
			                if lastfoid <> oRec("foid") then%>
			                <br><b><%=oRec("foldernavn")%></b>
			                <%end if%>
		                &nbsp;<a href="../inc/upload/<%=lto%>/<%=oRec("filnavn")%>" class='vmenulille' target="_blank"><%=oRec("filnavn")%></a><br>
		                <%
		                lastfoid = oRec("foid")
		                x = x + 1
		                oRec.movenext
		                wend
		                oRec.close 
		               
		               
                        
        end if
        end if 'showfiler
   
		
   
	
	
	
	%>
	

	</div><!-- side div -->
	
	
		<%
	    
	    
	    
	    if func = "print" then
	    
	    Response.Write("<script language=""JavaScript"">window.print();</script>")
	    
	    end if
	
	        
	       
	
	
	response.Cookies("job").expires = date + 31
	
	
end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
