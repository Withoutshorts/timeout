<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->






<%

      '**** Søgekriterier AJAX **'
        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
        case "FN_jobstatus"
                
            if len(trim(request("jobstatus"))) <> 0 then
            jobstatus = request("jobstatus")
            else
            jobstatus = 1
            end if

        
            if len(trim(request("jobid"))) <> 0 then
            jobid = request("jobid")
            else
            jobid = 0
            end if
                    

            strSQL = "UPDATE job SET jobstatus = "& jobstatus &" WHERE id = "& jobid
            oConn.execute(strSQL)
    
        Response.Redirect "job_print.asp?id="&jobid
                
                  
        end select
        
        response.end
        end if


%>
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/job_func.asp"-->
<!--#include file = "CuteEditor_Files/include_CuteEditor.asp" -->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

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
    
    'fletCHK3 = ""
    fletCHK2 = ""
    fletCHK1 = ""

    if len(trim(request("FM_flet"))) <> 0 then
    flet = request("FM_flet")
        select case flet 
        case 1 
        fletCHK1 = "CHECKED"
        'case 3
        'fletCHK3 = "CHECKED"
        case else
        fletCHK2 = "CHECKED"
        end select
    else
    flet = 0
    'fletCHK0 = "CHECKED"
    end if
    
    if len(trim(request("FM_flet_brodtxt"))) <> 0 then
    flet_brodtxt = 1
    else
    flet_brodtxt = 0
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


    if request("FM_kundeid") <> 0 then
    kunde_id = request("FM_kundeid")
    else
    kunde_id = "0,0"
    end if

    dim kunde_ids
    redim kunde_ids(1)
    kunde_idArr = split(kunde_id, ",")
    for x = 0 to UBOUND(kunde_idArr)
    kunde_ids(x) = kunde_idArr(x)
    next



    if len(trim(request("levinfo"))) <> 0 then
	levinfo = request("levinfo")
	else
	    
	    if len(trim(request.cookies("job")("levinfo"))) <> 0 then
	    levinfo = request.cookies("job")("levinfo")
	    else
	    levinfo = 0
	    end if
	
	end if
	response.cookies("job")("levinfo") = levinfo

    if request("FM_levid") <> 0 then
    lev_id = request("FM_levid")
     'response.cookies("job")("levid") = lev_id
    else
    lev_id = "0,0"
    end if


    dim lev_ids
    redim lev_ids(1)
    lev_idArr = split(lev_id, ",")
    for x = 0 to UBOUND(kunde_idArr)
    lev_ids(x) = lev_idArr(x)
    next
	
	if len(trim(request("jbinfo"))) <> 0 then
	jbinfo = request("jbinfo")
	else
	    
	    
	    select case lto
	    case "wowern"
	    jbinfo = 0

        case "dencker"
	    jbinfo = 0
	    
	     case else
	    
	   if len(trim(request.cookies("job")("jbinfo"))) <> 0 then
	    jbinfo = request.cookies("job")("jbinfo")
	    else
	    jbinfo = 0
	    end if
	    
	    end select
	    
	
	end if
	response.cookies("job")("jbinfo") = jbinfo
	
	
	if len(trim(request("jbprisinfo"))) <> 0 then
	jbprisinfo = request("jbprisinfo")
	else
	    
	   
        select case lto
	    case "synergi1", "intranet - local"
	    jbprisinfo = 0

        case "dencker"
        jbprisinfo = 1

	    case else
	    
	    if len(trim(request.cookies("job")("jbprisinfo"))) <> 0 then
	    jbprisinfo = request.cookies("job")("jbprisinfo")
	    else
        jbprisinfo = 1 '** skjul
        end if
	    
	    end select

	
	end if
	response.cookies("job")("jbprisinfo") = jbprisinfo
	


	'*** Vis jobansvarlig og periode ***'
	if len(trim(request("jbperansinfo"))) <> 0 then
	jbperansinfo = request("jbperansinfo")
	else
	    
	   
	    
	    if request.cookies("job")("jbperansinfo") <> "" then
            
            select case lto
            case "synergi1", "jttek"
            jbperansinfo = 1
             case "dencker"
            jbperansinfo = 1

            case else
	        jbperansinfo = request.cookies("job")("jbperansinfo")
            end select

	    else
            
            
                select case lto
	            case "xx"
	            jbperansinfo = 0
	            case else
	            jbperansinfo = 1
                end select

	    jbperansinfo = 0
	    end if
	    
	   
	end if
	response.cookies("job")("jbperansinfo") = jbperansinfo
	


    '*** Vis udsepcificering af aktiviteter ***'
	if len(trim(request("aktinfo"))) <> 0 then
	aktinfo = request("aktinfo")
	else
	    
	    if len(trim(request.cookies("job")("aktinfo"))) <> 0 then
	    aktinfo = request.cookies("job")("aktinfo")
	    else

            select case lto
            case "jttek"
            aktinfo = 1
            case else
	        aktinfo = 0
            end select

	    
	    end if
	
	end if
	response.cookies("job")("aktinfo") = aktinfo
	
	

    '** Vis akt. beskrivelse ****'
	if len(trim(request("aktinfo_besk"))) <> 0 then
	aktinfo_besk = request("aktinfo_besk")
	else
	    
	    if len(trim(request.cookies("job")("aktinfo_besk"))) <> 0 then
	    aktinfo_besk = request.cookies("job")("aktinfo_besk")
	    else
            select case lto
            case "synergi1", "intranet - local"
	        aktinfo_besk = 1
            case else
            aktinfo_besk = 0
            end select

	    end if
	
	end if
	response.cookies("job")("aktinfo_besk") = aktinfo_besk


    '*** Valuta EXT ****'
    select case lto
    case "intranet - local", "synergi1"
    valExtCode = "Kr."
    case else
    valExtCode = basisValISO '"DKK"
    end select
	
	
	if len(trim(request("vedrinfo"))) <> 0 then
	vedrinfo = request("vedrinfo")
	
	if vedrinfo = "-" then
	vedrinfo = ""
	end if
	
	else
	    
	    
	    select case lto
	        case "essens"
	        'vedrinfo = "Rekvisition"
            vedrinfo = "Job"
	        case "dencker", "jttek", "intranet - local"
	        vedrinfo = "Følgeseddel<br>"
	        case "wowern"
	        vedrinfo = "Budget"
            case "syngergi1"
            vedrinfo = "Prisoverslag"
	       
	     case else
	    
	    if len(trim(request.cookies("job")("vedrinfo"))) <> 0 then
	    vedrinfo = request.cookies("job")("vedrinfo")
	    else
	    vedrinfo = "Job"
	    end if
	    
	    end select
	    
	    
	   
	
	end if
	response.cookies("job")("vedrinfo") = vedrinfo
	
	
	
	
	if len(trim(request("aktinfo_priser"))) <> 0 then
	aktinfo_priser = request("aktinfo_priser")
	else
	    
         select case lto
            case "synergi1", "intranet - local", "dencker"
	        aktinfo_priser = 1
            case else

                if len(trim(request.cookies("job")("aktinfo_priser"))) <> 0 then
	            aktinfo_priser = request.cookies("job")("aktinfo_priser")
	            else
	            aktinfo_priser = 0 
                end if

            
            end select

	  
	
	end if
	response.cookies("job")("aktinfo_priser") = aktinfo_priser
	
	
	if len(trim(request("aktinfo_timer"))) <> 0 then
	aktinfo_timer = request("aktinfo_timer")
	else
	    
         select case lto
            case "synergi1", "intranet - local"
            aktinfo_timer = 1
            case "dencker", "jttek"
	        aktinfo_timer = 0
            case else
                if len(trim(request.cookies("job")("aktinfo_timer"))) <> 0 then
	            aktinfo_timer = request.cookies("job")("aktinfo_timer")
	            else
	            aktinfo_timer = 1
                end if

            end select

	   
	
	end if
	response.cookies("job")("aktinfo_timer") = aktinfo_timer
	

    if len(trim(request("aktinfo_sum"))) <> 0 then
	aktinfo_sum = request("aktinfo_sum")
	else
	    
        
        select case lto
        case "synergi1", "intranet - local", "dencker"
	    aktinfo_sum = 1

	    case else
        
            if len(trim(request.cookies("job")("aktinfo_sum"))) <> 0 then
	        aktinfo_sum = request.cookies("job")("aktinfo_sum")
	        else
            aktinfo_sum = 0
            end if
        
        end select

	    
	
	end if
	response.cookies("job")("aktinfo_sum") = aktinfo_sum
	
	
	'*** Valgte aktiviteter ***'
	if len(trim(request("FM_vis_akt"))) <> 0 then
	vlgAkt = request("FM_vis_akt")
	else
	    if len(trim(request("aktinfo"))) <> 0 then
	    vlgAkt = "1" ' Skal være 1 dvs vis ingen, da der IKKE er valgt nogen aktiviteter
	    else
	    vlgAkt = "-1"
	    end if
	end if

        'response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
        'response.write "vlgAkt: "& vlgAkt & " ("& request("FM_vis_akt") &") - "& request("aktinfo")

    '*** Valgte omkostninger ***'
	if len(trim(request("FM_vis_ulev"))) <> 0 then
	vlgUlev = request("FM_vis_ulev")
	else
	    if len(trim(request("ulevinfo"))) <> 0 then
	    vlgUlev = 0
	    else
	    vlgUlev = -1
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

    if len(trim(request("ulevinfo"))) <> 0 then
	ulevinfo = request("ulevinfo")
	else

        select case lto
        case "synergi1", "intranet - local", "dencker"
	    ulevinfo = 1
	    case else
            if len(trim(request.cookies("job")("ulevinfo"))) <> 0 then
	        ulevinfo = request.cookies("job")("ulevinfo")
	        else
            ulevinfo = 1    
            end if
        end select
	
	end if
	response.cookies("job")("ulevinfo") = ulevinfo


    if len(trim(request("ulevinfo_timer"))) <> 0 then
	ulevinfo_timer = request("ulevinfo_timer")
	else
	    
	    if len(trim(request.cookies("job")("ulevinfo_timer"))) <> 0 then
	    ulevinfo_timer = request.cookies("job")("ulevinfo_timer")
	    else
	    ulevinfo_timer = 1
	    end if
	
	end if
	response.cookies("job")("ulevinfo_timer") = ulevinfo_timer
    
    if len(trim(request("ulevinfo_priser"))) <> 0 then
	ulevinfo_priser = request("ulevinfo_priser")
	else
	    
         select case lto
         case "synergi1", "intranet - local"
	     ulevinfo_priser = 0
         case else
                 if len(trim(request.cookies("job")("ulevinfo_priser"))) <> 0 then
	            ulevinfo_priser = request.cookies("job")("ulevinfo_priser")
	            else
	            ulevinfo_priser = 1
	            end if
         end select

	    
	
	end if
	response.cookies("job")("ulevinfo_priser") = ulevinfo_priser

	
     if len(trim(request("ulevinfo_sum"))) <> 0 then
	ulevinfo_sum = request("ulevinfo_sum")
	else
	    
	    if len(trim(request.cookies("job")("ulevinfo_sum"))) <> 0 then
	    ulevinfo_sum = request.cookies("job")("ulevinfo_sum")
	    else
	    

            select case lto
            case "synergi1", "intranet - local"
	        ulevinfo_sum = 0
            case else
            ulevinfo_sum = 0 
            end select

	    end if
	
	end if
	response.cookies("job")("ulevinfo_sum") = ulevinfo_sum


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
     case "jttek", "intranet - local"

         if media <> "pdf" AND func <> "print" then
	     afslutningVal = ""'<img src=""../ill/jttek/jt_teknik_invoice.jpg"" border=0>"
	     end if

    case "essens", "xintranet - local"
         
         if media <> "pdf" AND func <> "print" then
	     afslutningVal = "<img src=""../ill/blank.gif"" border=0>"
	     end if

	case "wowern"
	     
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
	
	
	
    '********* DOK TYPE ********************'
    if len(trim(request("FM_doktype"))) <> 0 then
    dokType = request("FM_doktype")
    else 
    dokType = "10" 
    end if

    'dokTyper
    '10 Dokument
    '11 Skabelon
    '13 Tilbud
    '20 Følgeseddel - Slut levering
    '21 Følgeseddel - Del levering
    '30 Ordre
    '31 Rekvisition



	'********* SLET SKABELON *******************
	if func = "sletskabelon" then
	
	ts_id = request("sletskabelonid")
	strSQLdelts = "DELETE FROM tilbuds_skabeloner WHERE ts_id = " & ts_id
	oConn.execute(strSQLdelts)
	
	Response.Redirect "job_print.asp?id="&id
	
	
	end if



	
	'********* PRINT + gem ved print *******************
	if func = "print" AND request("printgemfil") = "1" then
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
    filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
	
	if len(trim(request("print_skabelonnavn"))) <> 0 then
	

    call kundenavnPDF(prNavn)
    prNavn = strKundenavnPDFtxt

    prNavn = request("print_skabelonnavn") &".txt"

	else
    prNavn = "print_"& filnavnDato &" "& filnavnKlok &".txt"
	end if
	
	
    prDoktype = request("print_doktype") 
	prTxt = request("FM_vis_indledning")
    jobbeskTxt = prTxt 
    
    prTxt = replace(prTxt, "##job_besk_slut##", "")
    prTxt = replace(prTxt, "##job_besk##", "")
    '##job_besk## - ##job_besk_slut##



	prEditor = session("user")
	
	if len(trim(request("print_foid"))) <> 0 then
	prFolderid = request("print_foid")
	else
	prFolderid = 1000
	end if
	
     '** Opdater eksisterende / eller gem ny ****'
    filfindes = 0
    strSQLfindesfil = "SELECT filnavn, id FROM filer WHERE filnavn = '"& prNavn &"' AND jobid = "& id 
    oRec3.open strSQLfindesfil, oConn, 3
    if not oRec3.EOF then
    filfindes = oRec3("id")
    end if
    oRec3.close
	
	if filfindes <> 0 then
	

        call updFil(prNavn,jobbeskTxt,prEditor,filnavnDato,prDoktype,id,kid,prFolderid,filfindes)
   

	'strSQLinsfil = "UPDATE filer SET filnavn = '"& prNavn &"', "_
	'&" filertxt = '"&jobbeskTxt&"', editor = '"& prEditor &"', dato = '"& filnavnDato &"', type = "& dokType &", "_
	'&" jobid = "& id &", kundeid = "& kid &", folderid = "& prFolderid &""_
	'&" WHERE id = " & filfindes

    
	'Response.Write strSQLinsfil
	'Response.end
    'oConn.execute(strSQLinsfil)

    else

         call insFil(prNavn,jobbeskTxt,prEditor,filnavnDato,prDoktype,id,kid,prFolderid)

	'strSQLinsfil = "INSERT INTO filer (filnavn, filertxt, editor, dato, type, jobid, kundeid, folderid, "_
	'&" adg_kunde, adg_alle, incidentid) VALUES ('"& prNavn &"', '"&jobbeskTxt&"', '"& prEditor &"', "_
	'&" '"& filnavnDato &"', "& dokType &", "& id &", "& kid &", "& prFolderid &", 0, 1, 0)"
	
    'Response.Write strSQLinsfil
	'Response.end
	'oConn.execute(strSQLinsfil)
	
    end if

  

    


    call TimeOutVersion()
	
	'** gem fysisk ***'
	    'file = replace(prNavn, " ", "_")
        file = prNavn
    	
        'Response.write "lto " & lto &" "& file
        'Response.end

	    Set objFSO = server.createobject("Scripting.FileSystemObject")
    	
	    if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\job_print.asp" then
    							
		    Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\upload\"&lto&"\"&file&"", True, False)
		    Set objNewFile = nothing
		    Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\upload\"&lto&"\"&file&"", 8)
    	
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
	
            

	
	'if len(trim(request("gemsomskabelon"))) = 0 AND len(trim(request("gemsomfil"))) = 0 then

	'<!--include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	'errortype = 146
	'call showError(errortype)
	
	'Response.end
	
	'end if
	
	'*** Gemmer skabelon ****'
	'if request("gemsomskabelon") = "1" OR request("gemsomfil") = "1" then
	
	    if len(trim(request("skabelonnavn"))) = 0 then
	    %>
	    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	    <%
	    errortype = 144
	    call showError(errortype)
	
	    Response.end
	    end if
	
	    if dokType <> "11" AND (len(trim(request("foid"))) = 0 OR request("foid") = 0) then
	    %>
	    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	    <%
	    errortype = 145
	    call showError(errortype)
	
	    Response.end
	
	    end if
	
	skabelonNavn = request("skabelonnavn")

           

      
    call kundenavnPDF(skabelonNavn)
    skabelonNavn = strKundenavnPDFtxt
	
    skabelonTxt = replace(request("FM_gem_txt"), "'", "''")
    jobBeskTxt = skabelonTxt 

    
    'skabelonTxt = replace(skabelonTxt, "##job_besk_slut##", "")
    'skabelonTxt = replace(skabelonTxt, "##job_besk##", "")
    '##job_besk## - ##job_besk_slut##

   


	tsDato = year(now) & "/"& month(now) &"/"& day(now)
	tsEditor = session("user")
	ts_kundeid = request("gemsomskabelon_kid")
	
	ts_folderid = request("foid")
	
	    '******** Gem som skabelon ****'
	    if dokType = "11" then
	    strSQLinsSkab = "INSERT INTO tilbuds_skabeloner (ts_navn, ts_txt, ts_dato, ts_editor, ts_kundeid) VALUES "_
	    &" ('"& skabelonNavn & "', '"& skabelonTxt &"', '"& tsDato &"', '"& tsEditor &"', "& ts_kundeid &")"
	
	    oConn.execute(strSQLinsSkab)
	
	    end if


	'** Gem som fil ****'
	if dokType <> "11" then
	
    '** Opdater eksisterende / eller gem ny ****'
    filfindes = 0
    strSQLfindesfil = "SELECT filnavn, id FROM filer WHERE (filnavn = '"& skabelonNavn &".txt' AND jobid = "& id &") OR id = "& dokid  
    oRec3.open strSQLfindesfil, oConn, 3
    if not oRec3.EOF then
    filfindes = oRec3("id")
    end if
    oRec3.close

    'Response.write "dokid "& dokid & "<br>"


	
	if cint(dokid) <> 0 OR filfindes <> 0 then
	
    if filfindes <> 0 then
    dokidThis = filfindes
    else
    dokidThis = dokid
    end if

     call updFil(skabelonNavn &".txt",skabelonTxt,tsEditor,tsdato,dokType,id,ts_kundeid,ts_folderid,dokidThis)

	'strSQLinsfil = "UPDATE filer SET filnavn = '"& skabelonNavn &".txt', "_
	'&" filertxt = '"&skabelonTxt&"', editor = '"& tsEditor &"', dato = '"& tsdato &"', type = "& dokType &", "_
	'&" jobid = "& id &", kundeid = "& ts_kundeid &", folderid = "& ts_folderid &""_
	'&" WHERE id = " & dokidThis

    
	'Response.Write strSQLinsfil
	'Response.end
    'oConn.execute(strSQLinsfil)

	else

     
     call insFil(skabelonNavn &".txt",skabelonTxt,tsEditor,tsdato,dokType,id,ts_kundeid,ts_folderid)
	
    'strSQLinsfil = "INSERT INTO filer (filnavn, filertxt, editor, dato, type, jobid, kundeid, folderid, "_
	'&" adg_kunde, adg_alle, incidentid) VALUES ('"& skabelonNavn &".txt', '"&skabelonTxt&"', '"& tsEditor &"', "_
	'&" '"& tsdato &"', "& dokType &", "& id &", "& ts_kundeid &", "& ts_folderid &", 0, 1, 0)"


    'Response.Write strSQLinsfil &"<br><br><br>"
	'Response.end
    'oConn.execute(strSQLinsfil)
	'Response.Write strSQLinsfil

   
	

    strSQLsel = "SELECT id FROM filer WHERE id <> 0 ORDER BY id DESC"
    oRec.open strSQLsel, oConn, 3
    if not oRec.EOF then
    dokid = oRec("id")
    end if
    oRec.close

	end if 'dok findes


	
	
        	
        	call TimeOutVersion()
        	
	        '** Gem fysisk ***'
            '** Overskriver eksisterende **'
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
        	
        	
	'**** Opdater jobbeskrivelse ***'
    if len(trim(request("jobbeskoskriv"))) <> 0 AND request("jobbeskoskriv") <> 0 then
    jobbeskoskriv = 1
    else
    jobbeskoskriv = 0
    end if
     
     call gemJobBesk(jobBeskTxt, jobbeskoskriv)       
	
	
	
	end if
	
	'Response.flush
	'Response.end
    'Response.Write "job_print.asp?id="&id&"&FM_dokument="&dokid
    'Response.end 
    Response.Redirect "job_print.asp?id="&id&"&FM_dokument="&dokid
	
	
	
	'end if
	
	end if 'gem
	
	
	
	
	if func <> "print" then%>

	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--include file="../inc/regular/topmenu_inc.asp"-->

<!--
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(3)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call jobtopmenu()
	%>
	</div>
    -->

     <%call menu_2014() %>

	<%else %>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
    

	<%end if 
	
	
	%>
	<script>


	    function gemtxt() {
	        document.getElementById("FM_gem_txt").value = document.getElementById("FM_vis_indledning").value

	    }

	    function xjobbesk() {

	        var r = confirm("Gem brødtekst som jobbeskrivelse på job? \n(Tekst mellem: ##job_besk## og ##job_besk_slut##) ")
	        if (r == true) {
            document.getElementById("jobbeskoskriv").value = 1
            } else {
            document.getElementById("jobbeskoskriv").value = 0
            }

    }

    function xjobbeskPdf() {

        var r = confirm("Gem brødtekst som jobbeskrivelse på job? \n(Tekst mellem: ##job_besk## og ##job_besk_slut##) ")
        if (r == true) {
            document.getElementById("jobbeskoskrivPdf").value = 1
        } else {
            document.getElementById("jobbeskoskrivPdf").value = 0
        }

    }

    function xjobbeskPr() {

        var r = confirm("Gem brødtekst som jobbeskrivelse på job? \n(Tekst mellem: ##job_besk## og ##job_besk_slut##) ")
        if (r == true) {
            document.getElementById("jobbeskoskrivPr").value = 1
        } else {
            document.getElementById("jobbeskoskrivPr").value = 0
        }

    }


	    function xgemtxtPdf() {
	        document.getElementById("FM_vis_indledning_pdf").value = document.getElementById("FM_vis_indledning").value
	        document.getElementById("FM_vis_afslutning_pdf").value = document.getElementById("FM_vis_afslutning").value
	        document.getElementById("pdf_skabelonnavn").value = document.getElementById("skabelonnavn").value
            document.getElementById("pdf_doktype").value = document.getElementById("doktype").value
	        document.getElementById("pdf_foid").value = document.getElementById("foid").value

	        //if (confirm("Filnavn: " + document.getElementById("pdf_skabelonnavn").value)) {
	        //    window.open.location.href = "job_make_pdf.asp"
                
            //}

	    }

	    function xgemtxtPrint() {

            //alert(document.getElementById("skabelonnavn").value)
	        document.getElementById("print_skabelonnavn").value = document.getElementById("skabelonnavn").value
	        document.getElementById("print_foid").value = document.getElementById("foid").value
            document.getElementById("print_doktype").value = document.getElementById("doktype").value
	    }




	  


	    $(document).ready(function () {


        
         //$("#a_fltbrodtxt").mouseover(function () {

         //  $(this).css("cursor", "pointer");

         //}

	    
        $("#a_fltbrodtxt").click(function () {
            
        
            $("#div_fltbrodtxt").hide(1000);
	           
        });



         



	        pagehgt = $("#pghgt").val()
	        //alert("her" +pagehgt)
	        $("#bgimg").height(pagehgt+'px')

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
	
	        
    

    '*******************************************************************************'
    '**** Henter side til redigering eller print/PDF        
    '*******************************************************************************'        
	
	
	if func <> "print" then
	        
            
            lft = 90
	        tp = 82
        	sdWdt = 1104
	        
	        
	         rekvnr = ""
     
	            '*** Henter det job der skal printes ***
		        strSQL = "SELECT id, jobnavn, jobnr,"_
		        &" k.kkundenavn, k.kid, k.kkundenr, k.adresse, k.postnr, k.city, k.telefon, k.land, "_
		        &" jobknr, "_
		        &" jobTpris, jobstatus, jobstartdato, jobslutdato, "_
		        &" job.dato, job.editor, tilbudsnr, "_
		        &" fakturerbart, budgettimer, fastpris, kundeok, job.beskrivelse, "_
		        &" ikkeBudgettimer, tilbudsnr, serviceaft, kundekpers, jobans1, jobans2, jo_bruttooms, rekvnr "_
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
        			
			        strBudget = oRec("jo_bruttooms") 'oRec("jobTpris")
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
        			
                    strJobStatus = oRec("jobstatus")

                    '*** Job Sbek incl. flette koder ***'
			        strBesk = oRec("beskrivelse") 
			        'strtilbudsnr = oRec("tilbudsnr")
        			
			        'if request("FM_aftaler") = 1 then
			        'intServiceaft = oRec("serviceaft")
			        'else
			        'intServiceaft = 0
			        'end if
        			
                    rekvnr = oRec("rekvnr")

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


                 '*** Henter gemt dokument ***'
	            dokTxt = ""
	            doknavn = ""
	            folderid = 0
                dokType = "10"
	            if cint(dokid) <> 0 then
	                   
        
                        strSQL = "SELECT filertxt, filnavn, folderid, type FROM filer WHERE id = " & dokid
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

                        dokType = oRec("type")

                        oRec.movenext
                        wend
                        oRec.close

                        'doknavnstyle = "color:#000000;"

                else
                '*** genindlæst job, forslå filnavn = jobnavn ***'
                doknavn = replace(strNavn, " ", "_")
                doknavn = replace(doknavn, ">", "--")
                doknavn = replace(doknavn, "&", "-oo-")
                doknavn = year(now) & "_"& month(now) & "_" & day(now) & "_"& replace(formatdatetime(now, 3), ":", "") &"_"& doknavn 
                'doknavnstyle = "color:#999999;"
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

    
    else 
    
    '****************************************************************'
    '**** DETTE ER den side der bliver vist ved PRINT ell. PDF ******'
    '****************************************************************'
	
    '**** Sætter Margins, sideWdt mm. ******'
    	    
            select case lto 
            case "essens"
         
            %>
               
               <%if media = "pdf" then 
               
                lft = 89
	            tp = 100 '50
                sdWdt = 660 '668
                else 'print
                
                lft = 40 '85
	            tp = 100
                sdWdt = 660 '668

               
                
               end if %>
             
		
		<% 

            case "synergi1", "intranet - local"
             %>
               
               <%if media = "pdf" then 
               
                lft = 20 '60
	            tp = 130
                sdWdt = 495
                pageheight = 1169
                
                else 'print
                
                lft = 20
	            tp = 70
                sdWdt = 495 '525
                pageheight = 1169
                
               end if
             
            
            case "jttek"
             %>
               
               <%if media = "pdf" then 
               
                lft = 200 '100
	            tp = 200
                sdWdt = 600 '720
                pageheight = 1000 '1050 '1169
                
                else 'print
                
                lft = 20
	            tp = 200
                sdWdt = 700
                pageheight = 1169
                
               end if

        


             case "dencker"
             %>
               
               <%if media = "pdf" then 
               
                lft = 20 '60
	            tp = 150
                sdWdt = 720
                pageheight = 1169
                
                else 'print
                
                lft = 20
	            tp = 150
                sdWdt = 720
                pageheight = 1169
                
               end if

        
            
            
            case else
            
            lft = 20
	        tp = 50
            pageheight = 1169

            if media <> "pdf" then
		    sdWdt = 800
		    else
		    sdWdt = 650
		    end if

            end select
           




	        
            '********************************************************************************************************
	        '********************************************************************************************************
	        
            ' Henter PDF txt **********
            '*******************************************************************************************************

            side2 = 0

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
                    indledningVal = replace(indledningVal, "##job_besk_slut##", "")
                    indledningVal = replace(indledningVal, "##job_besk##", "") 

                    afslutningsTXT = oRec("pdf_footer")
                    

                    end if
                    oRec.close
                    
                    '**** Sætter topmargin på side 2 *****'
                    select case lto
                    case "synergi1", "intranet - local", "jttek"
                    topm = "120" '270
                    case else
                    topm = "100"
                    end select
                    
                    if instr(indledningVal, "##breakpage##") <> 0 then
                    pageheight = 2311
                    end if%>
                    <form><input type="hidden" id="pghgt" value="<%=pageheight %>" /></form>                      
                    <%

                    if instr(indledningVal, "##breakpage##") <> 0 then
                    indledningVal = replace(indledningVal, "##breakpage##", "<BR style=""page-break-before: always""><img src=""../ill/blank.gif"" width=""10"" height="&topm&" border=""0""><br>")
                    side2 = 1
                    end if
                    '&"<div style=""position:absolute; top:1169px; left:0px; background-color:#ffffff; background-image: url(../ill/synergi/synergi_logo_og_adresse_out_2.jpg); background-size:100%; border:0px #999999 solid; width:827px; height:1169px;"">&nbsp;</div>")
                    
                  
                    	                   
	        else


            '**** Sætter topmargin på side 2 *****'
            select case lto
            case "synergi1", "intranet - local", "jttek"
            topm = "120"
            case else
            topm = "100"
            end select

              


		    indledningVal = request("FM_vis_indledning")

            indledningVal = replace(indledningVal, "##job_besk_slut##", "")
            indledningVal = replace(indledningVal, "##job_besk##", "")




            if instr(indledningVal, "##breakpage##") <> 0 then
            pageheight = 2311
            end if%>

                      <form><input type="hidden" id="pghgt" value="<%=pageheight %>" /></form>                      
            <%
                  
                   

            indledningVal = replace(indledningVal, "##breakpage##", "<table border=0 cellspacing=0 width=100% cellpadding=0 style='page-break-before:always;'><tr><td><h4></h4><img src=""../ill/blank.gif"" width=""10"" height="&topm&" border=""0""></td></tr></table>")
            '&"<div style=""position:absolute; top:1169px; left:0px; background-color:#ffffff; background-image: url(../ill/synergi/synergi_logo_og_adresse_out_2.jpg); background-size:100%; border:0px #999999 solid; width:827px; height:1169px;"">&nbsp;</div>")
            
            
            
    		end if
    		




            '********************************************************************************************'
            '****** Baggrund IMG og Side skift / Side 2 BG imng ***'
            '********************************************************************************************'

            select case lto
              case "jttek", "intranet - local"

                 if media = "pdf" then %>

                <div id="Div4" style="position:absolute; top:0; width:821px; height:1049px; padding:0px; left:0; border:0px #000000 dashed; z-index:-1;">
             <img alt="" src="https://outzource.dk/timeout_xp/wwwroot/ver2_14/ill/jttek/jttek_brevpapir_20140703.gif" width="821" height="1049" border="0"  /><!-- 769 1087 width="758" height="969" -->
             </div>
               
               <br />
              <%if cint(side2) = 1 then %>

              <div id="Div4" style="position:absolute; top:0; width:821px; height:1049px; padding:0px; left:0; border:0px #000000 solid; z-index:-1;">
             <img alt="" src="https://outzource.dk/timeout_xp/wwwroot/ver2_14/ill/jttek/jttek_brevpapir_20140703.gif" width="821" height="1049" border="0"  /><!-- 769 1087 -->
             </div>
           
             <%end if %>
             
             <%end if 

            case "dencker"
                    
                     if media = "pdf" then%>
                    <div id="fakturaside_bg" style="position:absolute; left:0px; top:0px; visibility:visible; background-color:#ffffff; border:<%=bd%>px #8caae6 solid; z-index:1; padding:0px 0px 0px 0px;">
                <img src="../ill/dencker/brevpapir_bg_850_2013.gif" />
                </div>
                    <%end if

            case "essens"

            if media = "pdf" then 
            %>  
                
                <!--
                <div id="jobside_bg" style="position:absolute; left:0px; top:0px; visibility:visible; background-color:#ffffff; border:0px #8caae6 solid; z-index:1; padding:0px 0px 0px 0px;">
                <img src="../ill/essens/essens_v3_brevpapir_5.gif" />
                
                </div>
                -->

              <div id="Div5" style="position:absolute; top:0px; height:1072px; width:758px; padding:0px 0px 0px 0px; left:20px; border:0px #000000 solid; z-index:-1;">
             <!--<img alt="" src="https://outzource.dk/timeout_xp/wwwroot/ver2_10/ill/essens/essens_brevpapir_V4_758_1050_3.png" width="758" height="1050" border="0"  />--><!-- 769 1087 -->
            <img alt="" src="https://outzource.dk/timeout_xp/wwwroot/ver2_14/ill/essens/essens_brevpapir_v6_lowres.png" width="758" height="1072" border="0"  />
              </div>

                <%
                else
                 %>

                 <div id="Div2" style="position:absolute; top:0px; height:1072px; width:758px; padding:0px; left:-28px; border:0px #000000 solid; z-index:-1;">
             <!--<img alt="" src="https://outzource.dk/timeout_xp/wwwroot/ver2_10/ill/essens/essens_brevpapir_V4_758_1050_3.png" width="758" height="1050" border="0"  />--><!-- 769 1087 -->
                      <img alt="" src="https://outzource.dk/timeout_xp/wwwroot/ver2_14/ill/essens/essens_brevpapir_v6_lowres.png" width="758" height="1072" border="0"  />
             </div>

                 <!--
                <div id="Div2" style="position:absolute; left:0px; top:0px; visibility:visible; background-color:#ffffff; border:1px #8caae6 solid; z-index:1; padding:0px 0px 0px 0px;">
                <img src="../ill/essens/essens_v3_brevpapir_1_p.gif" />
                </div>
                -->
                <%
                end if

            case "synergi1", "xintranet - local"
           

              if media = "pdf" then %>

                <div id="Div4" style="position:absolute; top:85px; height:969px; width:758px; padding:0px; left:-28px; border:0px #000000 solid; z-index:-1;">
             <img alt="" src="https://outzource.dk/timeout_xp/wwwroot/ver2_10/ill/synergi/synergi_12.jpg" width="758" height="969" border="0"  /><!-- 769 1087 -->
             </div>
               
               <br />
              <%if cint(side2) = 1 then %>

               <div id="Div3" style="position:absolute; top:1230px; height:969px; width:758px; padding:0px; left:-28px; border:0px #000000 solid; z-index:-1;">
              <img alt="" src="https://outzource.dk/timeout_xp/wwwroot/ver2_10/ill/synergi/synergi_12.jpg" width="758" height="969" border="0"  />
             </div>
           
             <%end if %>
             
             <%end if %>

             
             <%end select
    		





		    
	end if '** Print
	
	


    '**** Sætter Global TableWidth ********'

        select case lto 
        case "synergi1"
        tblWdt = "605"
        case else
	    tblWdt = "100%"
       end select


    '**************************************
	%>



    <script src="inc/jobprint_jav.js"></script> 
	
	
	
	
	<div id="sindhold" style="position:absolute; left:<%=lft%>px; top:<%=tp%>px; visibility:visible; z-index:50; border:0px #999999 dashed; padding:0px; width:<%=sdWdt%>px;">
	


	    
	    <%if func <> "print" then %>


	   
    <h4>Print / PDF Center<br /><span style="font-size:12px;">Og gem som dokument i filarkiv  (.txt / .pdf)</span></h4>
    
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
	   lastknr = 0
	    %>
	    
	     <b>Vælg job ell. tilbud:</b>
            <select id="jobsel" name="id" style="width:500px;" onchange="submit();">
            <option value="0">Vælg..</option>
            <%
            oRec4.open strSQLjob, oConn,  3 
            while not oRec4.EOF 


            if lastknr <> oRec4("kkundenr") then
            %>
              <option value="<%=oRec4("id") %>"disabled></option>
                <option value="<%=oRec4("id") %>"disabled></option>
             <option value="<%=oRec4("id") %>" disabled> <%=oRec4("kkundenavn") %> (<%=oRec4("kkundenr") %>)</option>
            <%end if

            if cint(id) = oRec4("id") then
            jSel = "SELECTED"
            else
            jSel = ""
            end if
            
            %>
            <option value="<%=oRec4("id") %>" <%=jSel %>><%=oRec4("jobnavn") %> (<%=oRec4("jobnr") %>) .............<%=left(oRec4("kkundenavn"), 20) %></option>
            
            <%

            if lastknr <> oRec4("kkundenr") then
            lastknr = oRec4("kkundenr")
            end if

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
		
		'strkundeHTML = strkundeHTML & strKnavn &"<br>"& strAdresse & "<br>"_
		'& strPostnr &" "& strBy & "<br>" & strLand & "<br><br>"  
		
		'if len(trim(intkundekpers)) <> 0 then
		'intkundekpers = intkundekpers
		'else
		'intkundekpers = 0
		'end if

        if len(trim(request("FM_kundeid"))) = 0 then
        kunde_ids(0) = kid
        kunde_ids(1) = intkundekpers
        end if

 		
		'strSQLkpers = "SELECT navn FROM kontaktpers WHERE id = "& intkundekpers 
	    'oRec.open strSQLkpers, oConn, 3
	    'while not oRec.EOF
	    
	    'strkundeHTML = strkundeHTML & "Att: "& oRec("navn") & "<br><br>"
	    
	    'oRec.movenext
	    'wend
	    'oRec.close 
	    
        strkundeHTML = strkundeHTML & "##header_adr##"
	    
        select case lto
        case "synergi1"
        strkundeHTML = strkundeHTML & "<br><br><br><br><br>"
	    strkundeHTML = strkundeHTML & afsBy &", den ##header_dt## </td></tr></table>"
        case else
	    strkundeHTML = strkundeHTML & "</td><td valign=bottom align=right><br><br><br><br><br>"
	    strkundeHTML = strkundeHTML & afsBy &", d. ##header_dt## </td></tr></table>"
	    end select

        
	   
        'strjobHTML = "<br><br><table cellspacing=0 cellpadding=0 border=0 width="""&tblWdt&"""><tr><td valign=top>"
		
		
		strjobHTML = "<br><br>"
	    %>
	    
	    
	    
	    <%
	    'if tilbudsnr <> 0 then
		'strJobHTML = strjobHTML & "<h4>Vedr. "& vedrinfo &": "& tilbudsnr &"</h4>"
	    'else
	    '*** Jobnavn / vedr. text ***
        select case lto
        case "jttek"
        vedrtxtval = "<br><br><br><b><u>"& replace(vedrinfo, "<br>", "") &"</u></b><span style=""float:right;"">"
        vedrtxtval = vedrtxtval & "<img src=""ill/blank.gif"" width=450 height=1 border=0>"
        vedrtxtval = vedrtxtval & "<b><u>No. "& strjobnr &"</u></span><br><br><br>"
        case "dencker", "intranet - local"
        vedrtxtval = "<br><br><br><b><u>"& replace(vedrinfo, "<br>", "") &"</u></b>"
        vedrtxtval = vedrtxtval & "<br><br>"
            
            if len(trim(rekvnr)) <> 0 then
            vedrtxtval = vedrtxtval & "Rekvisitionsnr.: "& rekvnr
            end if

        vedrtxtval = vedrtxtval & "<b><br><br><br>"
        case else
		vedrtxtval = "<br><br><br><h4>"& vedrinfo &"</h4><b>"
        end select


        select case lto
        case "essens"
                
                  if vedrinfo = "Rekvisition" then
                  strJobHTML = strJobHTML & vedrtxtval &" på sagsnr. "& strjobnr &"</b><br><br>"
                  else
		          strJobHTML = strJobHTML & vedrtxtval &" "& strNavn & " ("&strjobnr&")</b><br><br>"
                  end if

        case "synergi1", "xintranet - local"
                  
                  
		          strJobHTML = strJobHTML & vedrtxtval &" "& strNavn & " </b><br><br>"
                  
                  'if vedrinfo = "Prisoverslag" then
                  'strJobHTML = strJobHTML &"<br><br><b>Data</b><br>"
                  'strJobHTML = strJobHTML &"<table width=100% border=0 cellspacing=1 cellpadding=2>"
                  'strJobHTML = strJobHTML &"<tr><td>Oplag:</td><td>1.000 stk.</td></tr>"
                  'strJobHTML = strJobHTML &"<tr><td>Format:</td><td>297 x 297 mm</td></tr>"
                  'strJobHTML = strJobHTML &"<tr><td>Omfang:</td><td>24 sider</td></tr>"
                  'strJobHTML = strJobHTML &"<tr><td>Farver:</td><td>4 + 4 + mat kacering</td></tr>"
                  'strJobHTML = strJobHTML &"<tr><td>Papir:</td><td>400 g matbestrøget</td></tr>"
                  'strJobHTML = strJobHTML &"<tr><td>Færdiggørelse:</td><td>Renskæres, optages i sæt, hulles og spiraleres wire-O</td></tr>"
                  'strJobHTML = strJobHTML &"</table><br><br>"
                  'end if

        case else
            strJobHTML = strJobHTML & vedrtxtval &" "& strNavn & " ("&strjobnr&")</b><br><br>"
        
        end select

      
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
	    case "Følgeseddel<br>"
	    fSel = "SELECTED"
        case "Følgeseddel dellevering<br>"
	    fdSel = "SELECTED"
        case "Prisoverslag"
	    pSel = "SELECTED"
	    case ""
	    iSel = "SELECTED"
	    end select
		
		
		strChkjobHTML = "<b>Overskrift:</b><br> <SELECT name=""vedrinfo"" id=""vedrinfo"" style=""width:250px;"">"_
		&"<option value=""Job"" "& jSel &">Job</option>"_
		&"<option value=""Tilbud"" "& tSel &">Tilbud</option>"_
		&"<option value=""Rekvisition"" "& rSel &">Rekvisition</option>"_
        &"<option value=""Følgeseddel<br>"" "& fSel &">Følgeseddel</option>"_
        &"<option value=""Følgeseddel dellevering<br>"" "& fdSel &">Følgeseddel dellevering</option>"_
		&"<option value=""Budget"" "& bSel &">Budget</option>"_
        &"<option value=""Prisoverslag"" "& pSel &">Prisoverslag</option>"_
		&"<option value=""-"" "& iSel &">Ingen</option>"_
		&"</SELECT><br><br><b>"& strNavn & " ("&strjobnr&")</b><br>"
		
		
		if len(trim(strBesk)) <> 0 then
            select case lto
            case "synergi1"
            strJobHTML = strJobHTML & "<br><br><br>"& strBesk & "<br><br>"
            case else
		    strJobHTML = strJobHTML & "<br><br><b>Beskrivelse:</b><br>"& strBesk & "<br><br>"
		    end select
        end if
		
		call htmlparseCSV(strBesk)
        strBesk = htmlparseCSVtxt
		strChkjobHTML = strChkjobHTML & "<br><span style=""color:#999999;""><i>"& left(strBesk, 100) & "...</i></span><br>"
		
		
		'if lto <> "wowern" then
        select case lto
        case "synergi1"
        strJobPrisHTML = strJobPrisHTML & "<table cellspacing=0 cellpadding=0 border=0 width="""&tblWdt&"""><tr><td><u>Prisoverslag</u><br>  "& valExtCode &" "& formatnumber(strBudget, 2) &"</td></tr></table><br>"
		case else
		strJobPrisHTML = strJobPrisHTML & "<table cellspacing=0 cellpadding=0 border=0 width="""&tblWdt&"""><tr><td align=right><u>Samlet pris ekskl. moms, <b> "& valExtCode &" "& formatnumber(strBudget, 2) &"</b></u></td></tr></table><br>"
		end select
		
		'if prodtilbud = 2 then
		
		    if len(trim(jobans1)) <> 0 then
		    strJobInfoHTML = strJobInfoHTML & "<b>Jobansvarlig / vor. ref.:</b> " & jobans1 & "<br>"
            end if  
            
            if len(trim(jobans2)) <> 0 then
		    strJobInfoHTML = strJobInfoHTML & "<b>Jobejer:</b> " & jobans2 & "<br>"
            end if   
        
		strJobInfoHTML = strJobInfoHTML & "<br><b>Periode:<br></b>Arbejdet forventes udført i perioden "& formatdatetime(strTdato, 1) &" - "& formatdatetime(strUdato, 1) & "<br><br>&nbsp;"
		'end if

        '*****************
        '** header_adr ***
        '*****************

        '*** Hvis
        '*** Rekvisition / Følgeseddel

        modtagerAdrfundet = 0
        
        if vedrinfo = "Følgeseddel<br>" then 'forvalgt leveringssted fra kpers/filial

        strSQLk = "SELECT navn, adresse, postnr, town, land FROM kontaktpers WHERE kundeid = "& kunde_ids(0) & " AND kptype = 2" 
         
        'Response.write strSQLk


        oRec3.open strSQLk, oConn, 3
        if not oRec3.EOF then

        strHeader_adrHTML = oRec3("navn") &"<br>" & oRec3("adresse") & "<br>" & oRec3("postnr") & " " & oRec3("town")

        if oRec3("land") <> "Danmark" then
        strHeader_adrHTML = strHeader_adrHTML & "<br>"& oRec3("land")
        end if
        
        modtagerAdrfundet = 1
        end if
        oRec3.close


        end if


        if cint(modtagerAdrfundet) = 0 then
        
        strSQLk = "SELECT kkundenavn, kkundenr, kid, adresse, postnr, city, land FROM kunder WHERE kid = "& kunde_ids(0) 

        oRec3.open strSQLk, oConn, 3
        if not oRec3.EOF then

        strHeader_adrHTML =  oRec3("kkundenavn") &"<br>" & oRec3("adresse") & "<br>" & oRec3("postnr") & " " & oRec3("city")

        if oRec3("land") <> "Danmark" then
        strHeader_adrHTML = strHeader_adrHTML & "<br>"& oRec3("land")
        end if
      
        end if
        oRec3.close

       

        end if



         '** att pers. **
        if kunde_ids(1) <> 0 then
         strSQLk = "SELECT navn FROM kontaktpers WHERE id = "& kunde_ids(1) 
         
        oRec3.open strSQLk, oConn, 3
        if not oRec3.EOF then

        strHeader_adrHTML = strHeader_adrHTML & "<br><br>Att: "& oRec3("navn") 
      
        end if
        oRec3.close
        end if



        '**************
        '** alt_adr ***
        '**************
        strSQLlev = "SELECT kkundenavn, kkundenr, kid, adresse, postnr, city, land FROM kunder WHERE kid = "& lev_ids(0) 

        oRec3.open strSQLlev, oConn, 3
        if not oRec3.EOF then

        strAlt_adrHTML =  oRec3("kkundenavn") &"<br>" & oRec3("adresse") & "<br>" & oRec3("postnr") & " " & oRec3("city")
      
        end if
        oRec3.close

        
        '** att pers. **
        if lev_ids(1) <> 0 then
         strSQLlev = "SELECT navn FROM kontaktpers WHERE id = "& lev_ids(1) 

        oRec3.open strSQLlev, oConn, 3
        if not oRec3.EOF then

        strAlt_adrHTML = strAlt_adrHTML & "<br><br>Att: "& oRec3("navn") 
      
        end if
        oRec3.close
        end if

		      
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

        if aktinfo_sum = "0" then
        aktinfo_sumCHK0 = "CHECKED"
        aktinfo_sumCHK1 = ""
        else
        aktinfo_sumCHK0 = ""
        aktinfo_sumCHK1 = "CHECKED"
        end if
        
        
        select case lto
        case "wowern"
        strAktHTML = ""
        case "synergi1", "intranet - local"
            
           
        strAktHTML = ""
        case else
            
            if vlgAkt <> "1" AND vlgAkt <> "-1" then '0:Vis alle  'AND vlgAkt <> "-1" then
            strAktHTML = "<br><br><b>Udspecificering:</b>"
            end if
        
        end select

        
        

           select case lto 
            case "dencker", "jttek"
            strAktHTML = strAktHTML & "<table cellspacing=0 cellpadding=2 border=0 width="""&tblWdt&""">"
            case else
            strAktHTML = strAktHTML & "##job_besk##<table cellspacing=0 cellpadding=2 border=0 width="""&tblWdt&""">"
            end select
        
		
        cspanFull =  1
        if aktinfo_timer = "0" then
        cspanFull = cspanFull + 1
        end if

        if aktinfo_priser = "0" then
        cspanFull = cspanFull + 2
        end if

                       if aktinfo <> "1" then
		               aktinfoCHK0 = "CHECKED"
		               aktinfoCHK1 = ""
		               else
		               aktinfoCHK0 = ""
		               aktinfoCHK1 = "CHECKED"
		               end if


                       
		               strChkHTML = "<b style=""background-color:#D6dff5;""><u>Vis aktiviteter</u></b><br>"_
		               &"<input id=""aktinfo_0"" type=""radio"" name=""aktinfo"" value=""0"" "&aktinfoCHK0&" /> Ja (vælg alle) <input id=""aktinfo_1"" name=""aktinfo"" type=""radio"" value=""1"" "&aktinfoCHK1&" /> Nej (vælg ingen)"
        

        strChkHTML = strChkHTML &"<hr><table cellspacing=0 cellpadding=0 border=0 width="""&tblWdt&""">"
		strChkHTML = strChkHTML &"<tr><td colspan=4 bgcolor=""#FFFFFF"">"
		strChkHTML = strChkHTML &"Vis akt. beskrivelser <input id=""Checkbox1"" name=""aktinfo_besk"" value=""0"" type=""radio"" "& aktinfo_beskCHK0 &" />ja <input id=""Checkbox1"" name=""aktinfo_besk"" value=""1"" type=""radio"" "& aktinfo_beskCHK1 &" />nej<br>"_
		&"Vis priser pr. linje <input id=""Checkbox1"" name=""aktinfo_priser"" value=""0"" type=""radio"" "& aktinfo_priserCHK0 &" />ja <input id=""Checkbox1"" name=""aktinfo_priser"" value=""1"" type=""radio"" "& aktinfo_priserCHK1 &" />nej<br>"_
		&"Vis timer/stk. pr. linje <input id=""Checkbox1"" name=""aktinfo_timer"" value=""0"" type=""radio"" "& aktinfo_timerCHK0 &" />ja <input id=""Checkbox1"" name=""aktinfo_timer"" value=""1"" type=""radio"" "& aktinfo_timerCHK1 &" />nej<br>"_
		&"Vis samlet beløb af aktiviteter<input id=""Checkbox1"" name=""aktinfo_sum"" value=""0"" type=""radio"" "& aktinfo_sumCHK0 &" />ja <input id=""Checkbox1"" name=""aktinfo_sum"" value=""1"" type=""radio"" "& aktinfo_sumCHK1 &" />nej<br>&nbsp;"_
		&"</td></tr>"
	    
        
                        
             'call akttyper2009(2)        
             'aty_sql_fakbar          	
            select case lto
            case "synergi1", "intranet - local"
            strSQLaktFakbar = " AND fakturerbar = 1"
            case "essens"
            strSQLaktFakbar = " AND (fakturerbar = 1 OR fakturerbar = 2 OR fakturerbar = 61) "
            case "dencker"
            strSQLaktFakbar = " AND (fakturerbar = 1 OR fakturerbar = 2 OR fakturerbar = 61) "
            case else
            strSQLaktFakbar = " AND fakturerbar = 1"
            end select
            
	
			strSQLselAkt = "SELECT id AS aktid, navn, beskrivelse, job, fakturerbar,  "_
			&" aktstartdato, aktslutdato, editor, dato, budgettimer, aktfavorit, aktstatus, "_
			&" fomr, faktor, aktbudget, fase, bgr, antalstk, aktbudgetsum "_
			&" FROM aktiviteter WHERE job =" & id &" "& strSQLaktFakbar &" ORDER BY fase, sortorder, navn"
			
			'Response.Write strSQLselAkt
			'Response.flush
			oRec.open strSQLselAkt, oConn, 3
			
			         
			         
			         if aktinfo_timer = "0" then
				     cpsF = 2
				     'totPDright = 0
				     else
				     cpsF = 1
				        'if media <> "pdf" then
				        'totPDright = 35
				        'else
				        'totPDright = 0
				        'end if
				     end if
			
			af = 0
			a = 0
			lastFase = "-1fdf78#9d##"
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
				                strAktHTML = strAktHTML & "<tr><td colspan="&cpsF&"><span class=""jobpr16""><b>I alt:</b></span></td><td align=right valign=top style=""width:40px;""><span class=""jobpr16""><b>"& valExtCode &"</b></span></td><td align=right valign=top style=""width:80px;""><span class=""jobpr16""><b>"&formatnumber(sumTotFase, 2)&"</b></span></td></tr>"
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
				                      strAktHTML = strAktHTML & "<tr><td colspan="&cspanFull&"><br><br><span class=""jobpr18""><b>"&replace(strFase, "_", " ")&"</b></span></td></tr>"
				                      fsCHK = "CHECKED"
				                    else
				                      fsCHK = ""
				                    end if
				    
				                else
				    
				                strAktHTML = strAktHTML & "<tr><td colspan="&cspanFull&"><br><br><span class=""jobpr18""><b>"&replace(strFase, "_", " ")&"</b></span></td></tr>"
				                fsCHK = "CHECKED"
				                end if
				    
				   
				    
				             else
				             fsCHK = ""
				             end if
				 
				strChkHTML = strChkHTML & "<tr><td colspan="&cspanFull&" bgcolor=""#FFFFFF""><br><br><input id="&lcase(strFase)&" name=""FM_vis_fase"" "&fsCHK&" value=""#"&lcase(strFase)&"#"" type=""checkbox"" class=""faseoskrift"" /><b>"&replace(strFase, "_", " ")&"</b></td></tr>"    
			    lastFase = strFase	
                af = 0
				end if 
				
				
				if instr(vlgAkt, "#"& oRec("aktid") &"#") <> 0 OR (vlgAkt = "-1" AND (aktbudgetsum <> 0 OR lto = "synergi1")) then  
				strAktHTML = strAktHTML & "<tr><td style=""padding-right:20px; width:"&sdWdt&"px;"">"
				end if
				
				select case right(af, 1)
				case 0,2,4,6,8
				bgcol = "#D6DFf5"
				case else
				bgcol = "#ffffff"
				end select 
				
				strChkHTML = strChkHTML & "<tr style=""background-color:"& bgcol &";""><td style=""padding-right:20px;"" class=lille>" 
				
				'** CHKbokse
				if instr(vlgAkt, "#"& oRec("aktid") &"#") <> 0 OR (vlgAkt = "-1" AND (aktbudgetsum <> 0 OR lto = "synergi1")) then
				cbCHK = "CHECKED"
				else
				cbCHK = ""
				end if
				
				strAktcbox = "<input id=""af_"&lcase(strFase)&"_"&af&""" name=""FM_vis_akt"" type=""checkbox"" "&cbCHK&" value='#"& oRec("aktid") &"#' class=""aktoskrift""/>&nbsp;"
				
				strChkHTML = strChkHTML & strAktcbox
				
				if instr(vlgAkt, "#"& oRec("aktid") &"#") <> 0 OR (vlgAkt = "-1" AND (aktbudgetsum <> 0 OR lto = "synergi1")) then 
			
            	    select case lto
                    case "synergi1", "intranet - local"
                     strAktHTML = strAktHTML & "<span class=""jobpr16""><u>"& strNavn & "</u></span>"
                    case else
                     strAktHTML = strAktHTML & "<span class=""jobpr16"">"& strNavn & "</span>"
                    end select
               
				
                end if
				
				strChkHTML = strChkHTML & "<b>" & strNavn & "</b>"' & lcase(lastFase) &" <> " & lcase(strFase) &" AND " & len(trim(strFase)) 
				
				
				
				
				if len(trim(strBeskrivelse)) <> 0 then
				
				
				
				        if aktinfo_besk = "0" then
				    
				            if instr(vlgAkt, "#"& oRec("aktid") &"#") <> 0 OR (vlgAkt = "-1" AND (aktbudgetsum <> 0 OR lto = "synergi1")) then 
				            strAktHTML = strAktHTML & "<br /><span class=""jobpr14"">"& strBeskrivelse & "</span><br>&nbsp;" 'htmlparseCSVtxt 
				            end if
				
				       
				     
                        

                        else

                            select case lto
                            case "synergi1"
                            strAktHTML = strAktHTML & "<br>&nbsp;"
                            case else
                            strAktHTML = strAktHTML & ""
                            end select

                        end if

                        call htmlparseCSV(strBeskrivelse)
				  
				        strChkHTML = strChkHTML & "&nbsp;<font class=lillesort>"& left(htmlparseCSVtxt, 80) & "</font>"

                else
                    
                    select case lto
                    case "synergi1"
                    strAktHTML = strAktHTML & "<br>&nbsp;"
                    case else
                    strAktHTML = strAktHTML & ""
                    end select

				end if
				
				if instr(vlgAkt, "#"& oRec("aktid") &"#") <> 0 OR (vlgAkt = "-1" AND (aktbudgetsum <> 0 OR lto = "synergi1")) then 	
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
    				
				            strAktHTML = strAktHTML & "<td align=right valign=top style=""padding-right:40px; white-space:nowrap;""><span class=""jobpr16"">"& formatnumber(useAntal, 2) &" "& aty_enh 
                    
                            select case lto 
                            case "dencker"
                            strAktHTML = strAktHTML & "</span></td>" 
                            case else
                            strAktHTML = strAktHTML & " á "& formatnumber(aktbudget, 2) &"</span></td>" 
                            end select

                                   
				            
            	    end if
				
				end if
				
				
                if aktinfo_priser = "0" then
				
				if instr(vlgAkt, "#"& oRec("aktid") &"#") <> 0 OR (vlgAkt = "-1" AND (aktbudgetsum <> 0 OR lto = "synergi1")) then 
				strAktHTML = strAktHTML & "<td align=right valign=top style=""width:40px;"">"& valExtCode &"</td><td align=right valign=top style=""width:80px;""><span class=""jobpr16"">"& formatnumber(aktbudgetsum, 2) &"</span></td>"        
				sumTotFase = sumTotFase + aktbudgetsum
				end if

               
			    end if
				
				
				
				strChkHTML = strChkHTML & "<td align=right valign=top style=""white-space:nowrap;"" class=lille>"& formatnumber(useAntal, 2) &" "& aty_enh &"</td>"        
				strChkHTML = strChkHTML & "<td align=right valign=top class=lille style=""padding-left:10px;"">"& valExtCode &"</td><td align=right valign=top class=lille> "& formatnumber(aktbudgetsum, 2) &"</td></tr>"
				
				
				if instr(vlgAkt, "#"& oRec("aktid") &"#") <> 0 OR (vlgAkt = "-1" AND aktbudgetsum <> 0) then 
				aktbudgetTot = aktbudgetTot + aktbudgetsum
				end if

                if instr(vlgAkt, "#"& oRec("aktid") &"#") <> 0 OR (vlgAkt = "-1" AND (aktbudgetsum <> 0 OR lto = "synergi1")) then 
                strAktHTML = strAktHTML & "</tr>"
                end if
					     
			a = a + 1
			af = af + 1	
		
			oRec.movenext
			wend
			oRec.close
			
			if aktinfo_priser = "0" then
			
			        '*** Sum tot. på fase ***'
				    if a <> 0 AND sumTotFase <> 0 then
				     strAktHTML = strAktHTML & "<tr><td colspan="&cpsF&"><span class=""jobpr16""><b>I alt:</b></span></td><td align=right valign=top style=""width:40px;""><span class=""jobpr18""><b>"& valExtCode &"</b></span></td><td align=right valign=top style=""width:80px;""><span class=""jobpr18""><b>"&formatnumber(sumTotFase, 2)&"</b></span></td></tr>"
				    sumTotFase = 0
				    end if
			
			end if

            
            if request("media") = "pdf" then 
			pdright = 0
			else
			pdright = 0
			end if
			        
                    if aktinfo_sum = "0" then
			        strAktHTML = strAktHTML & "<tr><td colspan="& cspanFull &" align=right style=""padding-top:40px; width:"&tblWdt&"px; padding-right:"&pdright&";""><span class=""jobpr18"">I alt ekskl. moms, <b>"& valExtCode &" "& formatnumber(aktbudgetTot, 2) &"</b></span></td></tr>"        
			        end if
			
			strAktHTML = strAktHTML & "</table>"
            
            select case lto 
            case "dencker", "jttek"
            case else
            strAktHTML = strAktHTML & "##job_besk_slut##"
            end select
            
            strAktHTML = strAktHTML & "<br><br>&nbsp;"

		    
            strChkHTML = strChkHTML & "<tr><td colspan=4 align=right style=""padding-top:40px;"" class=lille>I alt ekskl. moms, <b>"& valExtCode &" "& formatnumber(aktbudgetTot, 2) &"</b></td></tr>"        
			strChkHTML = strChkHTML & "</table><br><br>&nbsp;"
		
		    
            


            '*****************************************************'
            '*** Omkostninger ************************************'
            '*****************************************************'

             if ulevinfo_priser = "0" then
            ulevinfo_priserCHK0 = "CHECKED"
            ulevinfo_priserCHK1 = ""
            else
            ulevinfo_priserCHK0 = ""
            ulevinfo_priserCHK1 = "CHECKED"
            end if
        
            if ulevinfo_timer = "0" then
            ulevinfo_timerCHK0 = "CHECKED"
            ulevinfo_timerCHK1 = ""
            else
            ulevinfo_timerCHK0 = ""
            ulevinfo_timerCHK1 = "CHECKED"
            end if

            if ulevinfo_sum = "0" then
            ulevinfo_sumCHK0 = "CHECKED"
            ulevinfo_sumCHK1 = ""
            else
            ulevinfo_sumCHK0 = ""
            ulevinfo_sumCHK1 = "CHECKED"
            end if


             cspanFull =  1
            if ulevinfo_timer = "0" then
            cspanFull = cspanFull + 1
            end if

            if ulevinfo_priser = "0" then
            cspanFull = cspanFull + 2
            end if


            select case lto 
            case "synergi1", "intranet - local"
            strUlevHTML = "<table cellspacing=0 cellpadding=2 border=0 width="&tblWdt&">"
            case else
            strUlevHTML = "<b>Omkostninger</b><table cellspacing=0 cellpadding=2 border=0 width="&tblWdt&">"
            end select
		    
             
             strChkUlevHTML = "<hr><table cellspacing=0 cellpadding=0 border=0 width=""100%""><tr><td colspan=4>"_
		     &"Vis priser pr. linje <input id=""Checkbox1"" name=""ulevinfo_priser"" value=""0"" type=""radio"" "& ulevinfo_priserCHK0 &" />ja <input id=""Checkbox1"" name=""ulevinfo_priser"" value=""1"" type=""radio"" "& ulevinfo_priserCHK1 &" />nej<br>"_
		     &"Vis antal pr. linje <input id=""Checkbox1"" name=""ulevinfo_timer"" value=""0"" type=""radio"" "& ulevinfo_timerCHK0 &" />ja <input id=""Checkbox1"" name=""ulevinfo_timer"" value=""1"" type=""radio"" "& ulevinfo_timerCHK1 &" />nej<br>"_
		     &"Vis samlet beløb af omkostninger<input id=""Checkbox1"" name=""ulevinfo_sum"" value=""0"" type=""radio"" "& ulevinfo_sumCHK0 &" />ja <input id=""Checkbox1"" name=""ulevinfo_sum"" value=""1"" type=""radio"" "& ulevinfo_sumCHK1 &" />nej<br>&nbsp;"_
		     &"</td></tr>"
		
		        strSQLUlev = "SELECT ju_id, ju_navn, ju_ipris, ju_faktor, "_
		        &" ju_belob, ju_fase, ju_stk, ju_stkpris FROM job_ulev_ju WHERE ju_jobid = "& id  & " ORDER BY ju_navn "
		        oRec.open strSQLUlev, oConn, 3 
		        x = 0
		        while not oRec.EOF 
		        


                if oRec("ju_belob") <> 0 then
                ulevbudgetsum = oRec("ju_belob")
                else
                ulevbudgetsum = 0
                end if

                if instr(vlgUlev, "#"& oRec("ju_id") &"#") <> 0 OR (vlgUlev = "-1" AND (ulevbudgetsum <> 0 OR lto = "synergi1")) then 
                strUlevHTML = strUlevHTML &"<tr><td valign=top style=""width:"&sdWdt&"px;"">"& oRec("ju_navn") &"</td>"
                
                ulevbudgetTot = ulevbudgetTot + ulevbudgetsum

                end if

                select case right(x, 1)
				case 0,2,4,6,8
				bgcol = "#FFCC66"
				case else
				bgcol = "#ffffff"
				end select 
				
				strChkUlevHTML = strChkUlevHTML &"<tr style=""background-color:"& bgcol &";""><td style=""padding-right:20px;"" class=lille>"
                
                
                '** CHKbokse
				if instr(vlgUlev, "#"& oRec("ju_id") &"#") <> 0 OR (vlgUlev = "-1" AND (ulevbudgetsum <> 0 OR lto = "synergi1")) then
				ulCHK = "CHECKED"
				else
				ulCHK = ""
				end if
				
				strChkUlevHTML = strChkUlevHTML &"<input id=""ul_"&x&""" name=""FM_vis_ulev"" type=""checkbox"" "&ulCHK&" value='#"& oRec("ju_id") &"#' />&nbsp;"
				strChkUlevHTML = strChkUlevHTML & oRec("ju_navn") &"</td>"


                


                if ulevinfo_timer = "0" then
				      '*** vlgUlev = "-1" = alle ? 
				    if instr(vlgUlev, "#"& oRec("ju_id") &"#") <> 0 OR (vlgUlev = "-1" AND (ulevbudgetsum <> 0 OR lto = "synergi1")) then 
    				
				            strUlevHTML = strUlevHTML & "<td align=right valign=top style=""padding-right:40px;""><span class=""jobpr16"">"& formatnumber(oRec("ju_stk"), 2) 
            
                            select case lto 
                            case "dencker"
                            strUlevHTML = strUlevHTML & "</span></td>" 
                            case else
                            strUlevHTML = strUlevHTML & " á "& formatnumber(oRec("ju_stkpris"), 2) &"</span></td>" 
                            end select
                                
				            
            	    end if
				
				end if
				
				if ulevinfo_priser = "0" then
				
				    if instr(vlgUlev, "#"& oRec("ju_id") &"#") <> 0 OR (vlgUlev = "-1" AND (ulevbudgetsum <> 0 OR lto = "synergi1")) then 
				    strUlevHTML = strUlevHTML & "<td align=right valign=top style=""width:40px;"">"& valExtCode &"</td><td align=right valign=top style=""width:80px;""><span class=""jobpr16"">"& formatnumber(ulevbudgetsum, 2) &"</span></td></tr>"        
				    end if
				
				end if
		        
                strChkUlevHTML = strChkUlevHTML & "<td align=right valign=top style=""white-space:nowrap;"" class=lille>"& formatnumber(oRec("ju_stk"), 2) &"</td>"        
				strChkUlevHTML = strChkUlevHTML & "<td align=right valign=top class=lille>"& valExtCode &"</td><td align=right valign=top class=lille> "& formatnumber(ulevbudgetsum, 2) &"</td></tr>"

                

		        x = x + 1
		        oRec.movenext
		        wend
		        oRec.close 
		      
		        
                 if ulevinfo_sum = "0" then
			     strUlevHTML = strUlevHTML & "<tr><td colspan="& cspanFull &" align=right style=""padding-top:40px; width:"&tblWdt&"px; padding-right:"&pdright&";""><span class=""jobpr18"">I alt ekskl. moms, <b>"& valExtCode &" "& formatnumber(ulevbudgetTot, 2) &"</b></span></td></tr>"        
			     end if
                
                
                strUlevHTML = strUlevHTML &"</table><br><br><br>"

                strChkUlevHTML = strChkUlevHTML & "<tr><td colspan=4 align=right style=""padding-top:40px;"" class=lille>I alt ekskl. moms, <b>"& valExtCode &" "& formatnumber(ulevbudgetTot, 2) &"</b></td></tr>"        
			    strChkUlevHTML = strChkUlevHTML & "</table><br><br>&nbsp;"


			
	    'if prodtilbud = 2 then
		        strMileHTML = "<hr><b>Milepæle</b><hr><table cellspacing=0 cellpadding=0 border=0 width=""100%"">"
		
		        strSQL = "SELECT navn, milepal_dato, jid, beskrivelse FROM milepale WHERE jid = "& id & " ORDER BY milepal_dato"
		        oRec.open strSQL, oConn, 3 
		        x = 0
		        while not oRec.EOF 
		        strMileHTML = strMileHTML &"<tr><td valign=top>"& formatdatetime(oRec("milepal_dato"), 2) & ": "& oRec("navn") & "<br>"& oRec("beskrivelse") & "<br><br>&nbsp;</td></tr>"
		        
		        x = x + 1
		        oRec.movenext
		        wend
		        oRec.close 
		      
		      strMileHTML = strMileHTML &"</table>"
		
		
        strHilsenHTML = ""

        select case lto
        case "synergi1"
        strHilsenHTML = strHilsenHTML &"<table cellspacing=0 cellpadding=0 border=0 width=""100%""><tr><td style=""padding-top:10px;"">"_
        &"Priser må betragtes som budgetramme, da opgaven endnu ikke kendes i detaljer.<br><br>"_
        &"Alle priser er inkl. levering til én adresse i Danmark og ekskl. moms.<br><br>"_
        &"Vi henviser iøvrigt til vores salgs og leveringsbetingelser på<br>"_
        &"www.syngergi1.dk/salg_levering</td></tr></table><br><br>"        
		
        case "dencker", "intranet - local", "jttek"
        strHilsenHTMLst = "<table cellspacing=0 cellpadding=2 border=0 width=""100%"">"_
         &"<tr><td colspan=2><b>Note:</b></td>"_
         &"<tr><td valign=top id=""td_hilsSt_1"" styke=""width:40px;"">[ ]</td><td> Dellevering"_
         &"</td></tr>"_
         &"<tr><td valign=top id=""td_hilsSt_2"">[ ]</td><td> Slutlevering for denne del at ordren"_
         &"</td></tr>"_
         &"<tr><td valign=top id=""td_hilsSt_3"">[ ]</td><td> Slutlevering for hele ordren"_
         &"</td></tr>"_
         &"<tr><td valign=top id=""td_hilsSt_4"">[ ]</td><td id=""td_hilsSt_4_txt""> Andet:<br>"_
         &"</td></tr>"_
         &"</table><br><br>"   

          %> 
          <div id="dv_flgseddel_footer" style="display:none; visibility:hidden;"><%=strHilsenHTMLst %></div>
          <%

       

        end select

        select case lto
        case "synergi1"
        afsenderName = "Thomas Hansen"
        case else
        afsenderName = session("user")
        end select


		strHilsenHTML = strHilsenHTML &"<table cellspacing=0 cellpadding=0 border=0 width=""100%""><tr><td style=""padding-top:10px;"">Med venlig hilsen <br>"& afsKnavn &"<br><br>"& afsenderName &"</td></tr></table>"        
		
              
		
		'end if
		
            
            
            
   
    '******************************************
    '***** Filter ****************************
    '*******************************************



                    
                    strjobstatusSEL0 = ""
                    strjobstatusSEL1 = ""
                    strjobstatusSEL2 = ""
                    strjobstatusSEL3 = ""
                    strjobstatusSEL4 = ""
                    
                    nejtakval = strjobstatus
                  
                select case strjobstatus
                case 0
                    strjobstatusSEL0 = "SELECTED" 
                case 1
                    strjobstatusSEL1 = "SELECTED"
                case 2
                    strjobstatusSEL2 = "SELECTED"
                case 3
                    strjobstatusSEL3 = "SELECTED"
                case 4
                    strjobstatusSEL4 = "SELECTED"
                case else
                    strjobstatusSEL1 = "SELECTED"
                end select %>

                <div id="dv_jobst" style="position:absolute; left:835px; top:200px; width:240px; border:4px yellowgreen solid; visibility:hidden; display:none; background-color:#ffffff; padding:20px; z-index:200000;">
		        <form>

               

                      <input id="jobs_jobid" value="<%=id%>" type="hidden" />
                      <input id="jobs_pdf_pnt" value="0" type="hidden" />
                   <br /><b>Skift jobstatus</b> <br />
                    Det er muligt at skifte status på jobbet her.<br /> 
                    <select id="jobstatus">
                        <%select case lto
                       case "dencker", "intranet - local", "jttek"
                            
                       case else %>
                   <option value="<%=nejtakval %>">Nej Tak - videre..</option>
                        <%end select %>
                   <option value="1" <%=strjobstatusSEL1 %>>Aktiv</option>
                   <option value="0" <%=strjobstatusSEL0 %>>Lukket</option>
                   <option value="2" <%=strjobstatusSEL2 %>>Til fak.</option>
                   <option value="3" <%=strjobstatusSEL3 %>>Tilbud</option>
                   <option value="4" <%=strjobstatusSEL4 %>>Gennemsyn</option>

                              </select>
                    
                    
                    <%select case lto
                       case "dencker", "jttek", "intranet - local" %>
                      <br><br><br>
                    <b>Angiv leveringsstatus:</b> (på følgeseddel)<br />
                    <table cellspacing="0" cellpadding="0" border="0" width="100%">
                     <tr><td><input type="radio" name="statushilsenTxtSet" id="statushilsenTxtSet_0" value="0" /></td><td>Spring dette over</td></tr>
                     <tr><td><input type="radio" name="statushilsenTxtSet" id="statushilsenTxtSet_1" value="1" /></td><td> Dellevering</td></tr>
                     <tr><td><input type="radio" name="statushilsenTxtSet" id="statushilsenTxtSet_2" value="2" CHECKED /></td><td> Slutlevering for denne del at ordren</td></tr>
                     <tr><td><input type="radio" name="statushilsenTxtSet" id="statushilsenTxtSet_3" value="3" /></td><td> Slutlevering for hele ordren</td></tr>
                     <tr><td valign="top"><input type="radio" name="statushilsenTxtSet" id="statushilsenTxtSet_4" value="4" /></td><td> Andet: <textarea id="statushilsenTxtTxt" style="width:200px; height:50px;"></textarea></td></tr>
                     </table><br><br>  
                    <%end select %>
                    
                    &nbsp;<input type="button" value=" Videre >> " id="jobstatus_bt" /><br />
                    
                    <!--<label id="lbs_jobst" style="color:yellowgreen; font-size:12px; visibility:hidden; display:;"><i>V - Opdateret</i</label>-->
                  
                        

                   

                         
                    
                    </form>
                    <br /><br />&nbsp;
              </div>

    
    
    <%
                
    tTop = 25
	tLeft = 0
	tWdth = 1104
	
	
	call tableDiv(tTop,tLeft,tWdth)
	%>


	
                    
                    <table cellspacing=0 cellpadding=0 border=0 width=100%>
                    
	    
                    <tr><td valign=top> 
                    
                    <%
                          tdheight = 480
                          ptop = 0
                          pleft = 0
                          pwdt = 460
            
                         call filteros09(ptop, pleft, pwdt, "1) Hent fil <span style=""font-size:11px; color:#999999; font-weight:lighter;""><br>Skabelon ell. gemt dokument</span>", 3, tdheight)
                        
                         %>
                    
                    
                    <%if len(trim(request("formsbm"))) <> 0 then 
                        
                        if len(request("ignorerKid")) <> 0 then
                        igKidSQL = " AND ts_kundeid <> 0"
                        chkigKid = "CHECKED"
                        ignkidVal = 1
                        else
                        ignkidVal = 0 
                        chkigKid = ""
                        igKidSQL = " AND ts_kundeid = "& kid
                        end if

                    else
                        
                        select case lto
                        case "essens", "synergi1", "intranet - local"

                        igKidSQL = " AND ts_kundeid <> 0"
                        chkigKid = "CHECKED"
                        ignkidVal = 1
                        
                        case else

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
                        
                        end select

                    end if
                    
                    response.Cookies("tsa")("ignkidval") = ignkidVal
                    
                    level = session("rettigheder")
                    
                    if level <= 2 OR level = 6 then%>
                    </td>
                    </tr>
                        <form action="job_print.asp?func=sletskabelon&id=<%=id %>&kid=<%=kid %>&formsbm=1" method="post">
                    <tr>
                    
     
                    <td align=right style="padding:5px 50px 0px 0px;">
                    <input id="sletskabelonid" name="sletskabelonid" value="0" type="hidden" />
                    <input id="Submit3" type="submit" value="[X ] Slet Skabelon" style="font-size:9px; color:red; width:100px; background-color:#ffffff;" />
	                </td>
	                 
	    </tr>

                        </form>
                         <form action="job_print.asp?id=<%=id%>&kid=<%=kid %>&formsbm=1" method="post">
	     <tr>
	     
	     
           
	        <td valign=top style="padding:0px 10px 10px 10px;">
                    
                    
                    
                    <input id="Checkbox1" name="ignorerKid" value="1" type="checkbox" <%=chkigKid %> /> Vis alle (ignorer tilhørsforhold til kunde)
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
                        
                        <img src="../ill/blank.gif" width="1" height="2" /><br />

                        
                         <input id="Checkbox5" type="checkbox" name="FM_flet" value="1" <%=fletCHK %> />Flet skabelon med datavalg (se fletkoder) <br />
                         

                        <input id="Submit5" type="submit" value=" Hent skabelon >> " />
                        
                        </td>
                        
                        </tr>
                        </form>
                        
                         
                    <form action="job_print.asp?id=<%=id%>&kid=<%=kid %>&formsbm=1" method="post">
                        <tr>
                        
                        <td valign=top style="padding:10px;"><b>Dokumenter (kun .txt filer) oprettet på dette job:</b><br />
                        Folder \ filnavn (dato) &nbsp; <a href="filer.asp?nomenu=1&jobid=0&kundeid=<%=kid %>" class="vmenu" target="_blank">Se filarkiv >></a> <br />
                        
                        <%
                                
                        strSQLdok = "SELECT f.filnavn, f.id AS fid, f.dato, fo.navn AS foldernavn FROM filer f "_
                        &" LEFT JOIN foldere fo ON (fo.id = f.folderid) WHERE f.jobid =" & id & " ORDER BY foldernavn, filnavn" 
                        
                        'Response.Write strSQLdok
                        'Response.flush
                        
                        %>
                            <select id="Select1" name="FM_dokument" size=4 style="width:425px;">
                                
                            
                        <%
                        oRec.open strSQLdok, oConn, 3
                        while not oRec.EOF
                        
                        if cint(dokid) = oRec("fid") then
                        tSel = "SELECTED"
                        else
                        tSel = ""
                        end if 
                         
                        


                         if instr(oRec("filnavn"), "pdf") = 0 then%>
                        <option value="<%=oRec("fid") %>" <%=tSel %>><%=oRec("foldernavn") %> \ <%=oRec("filnavn") &"  (" & oRec("dato")%>)</option>
                        <% end if


                           

                        oRec.movenext
                        wend
                        oRec.close
                        %>
                        </select>

                          
                        
                        <input type ="checkbox" name="FM_flet_brodtxt" value="1" /> Flet (manuelt) brødtekst fra dok. med nye aktiviteter og priser <!--(henter kun brødtekst fra dok.)-->
                       <br />
                            <input id="Submit4" type="submit" value=" Hent dokument >> " />
                        </td>
                        
                        </tr>
                        
                        </form>
                        </table>
                        
                        
                        </td>
                        <form action="job_print.asp?id=<%=id%>&kid=<%=kid %>&func=gem&formsbm=1" method="post">
                        <input id="FM_gem_txt" name="FM_gem_txt" type="hidden" />
                        <td valign=top>
                        
   
	
                    
                    
                 
                    
                    <%
                          
                          ptop = 0
                          pleft = 0
                          pwdt = 350
            
                         call filteros09(ptop, pleft, pwdt, "3) Gem dokument <span style=""font-size:11px; color:#999999; font-weight:lighter;""><br>Gem i filarkiv eller som skabelon..?</span>", 3, tdheight)
                        
                         %>
                    
                   
                  
                    <b>Hvis du ønsker at gemme som dokument i fil-arkivet skal du vælge en folder:</b>
                    
                     
                    <br />
                    <a href="javascript:popUp('filer.asp?func=oprfo&kundeid=<%=kid%>&jobid=<%=id%>','600','500','250','120');" target="_self" class=vmenu><img src="../ill/folder_ikon.gif" width="17" height="15" alt="" border="0">&nbsp;Opret folder >></a>
                    <span style="color:#999999;">(re-loader side!) </span>
                    <br />
                    
                    <br />
                    <b>Vælg Folder:</b>
                   
                    <%
                    strSQLfol = "SELECT f.id as foid, f.navn AS fonavn, f.jobid, j.jobnavn, j.jobnr FROM foldere f "_
                    &" LEFT JOIN job j ON (j.id = f.jobid) WHERE f.jobid = " & id & " OR (f.kundeid = "& kid &" AND f.jobid = 0) OR (f.kundeid = 0) ORDER BY f.navn"
                    
                     
                    %>
                    
                    <br /> <select id="foid" name="foid" size="5" style="width:324px;">
                    <%
                    oRec3.open strSQLfol, oConn, 3
                    while not oRec3.EOF 
                    

                    if cdbl(folderid) <> 0 then

                        if cdbl(folderid) = oRec3("foid") then
                        fSel = "SELECTED"
                        else
                        fSel = ""
                        end if

                    else

                        'if lcase(left(trim(oRec3("fonavn")), 4)) = lcase(mid(trim(doknavn),11, 4)) then
                        '** Potentielt kan der være flere folderde på et job ***'
                        if cdbl(id) = oRec3("jobid") then
                        fSel = "SELECTED"
                        else
                        fSel = ""
                        end if
                    
                    end if


                    %>
                    <option value="<%=oRec3("foid") %>" <%=fSel %>><%=oRec3("fonavn") %> \ 
                   
                    </option>
                    <%
                    oRec3.movenext
                    wend
                    oRec3.close
                    %>
                    </select>
		           <br />
                   <br /><b>Filnavn:</b> <br /><input name="skabelonnavn" id="skabelonnavn" value="<%=doknavn%>" style="width:320px;" type="text" /> .txt / .pdf<br />
               

                          
                            <%
                                
                                dokType10CHK = ""
                                dokType11CHK = ""
                                dokType13CHK = ""
                                dokType20CHK = ""
                                dokType21CHK = ""
                                dokType30CHK = "" 
                                
                                select case dokType
                                case "10"
                                dokType10CHK = "SELECTED"
                                case "11"
                                dokType11CHK = "SELECTED"
                                case "13"
                                dokType13CHK = "SELECTED"
                                case "20"
                                dokType20CHK = "SELECTED"
                                case "21"
                                dokType21CHK = "SELECTED"
                                case "30"
                                dokType30CHK = "SELECTED"
                                case "31"
                                dokType31CHK = "SELECTED"
                                case else
                                dokType10CHK = "SELECTED"
                                end select %>

                    
                  
                    <br />Type:<br /><select name="FM_doktype" id="doktype">
                    <option value="10" <%=dokType10CHK %>>Dokument</option>
                    <option value="11" <%=dokType11CHK %>>Skabelon</option>
                        <!-- Der Bruges Overskrifter fra dataflet istedet
                    <option value="13" <%=dokType13CHK %>>Tilbud</option>
                    <option value="20" <%=dokType20CHK %>>Følgeseddel - Slut levering</option>
                    <option value="21" <%=dokType21CHK %>>Følgeseddel - Del levering</option>
                    <option value="30" <%=dokType30CHK %>>Ordre</option>
                    <option value="31" <%=dokType31CHK %>>Rekvisition</option>
                        -->
                         </select>
                          <br /><br />

                            <!--

                                    <%'if cint(dokid) <> 0 then
                    dkCHK = "CHECKED"
                    'else
                    'dkCHK = ""
                    'end if
                     %>
                    <input id="gemsomfil" name="gemsomfil" type="checkbox" value="1" <%=dkCHK %> /> Gem som dokument <span style="color:#999999; padding:1px;">(overskriver eksisterende fil med samme navn)</span>
             
                   

                                <br />
                   
                    <input id="gemsomskabelon" name="gemsomskabelon" type="checkbox" value="1" /> Gem som skabelon<br />
                        <br />

                                 -->


                     <input id="gemsomskabelon_kid" name="gemsomskabelon_kid" value="<%=kid %>" type="hidden" />
                   <input id="dokid2" name="FM_dokument" value="<%=dokid %>" type="hidden" />

                                
                    
                    <%
                    'select case lcase(lto)
                    'case "syngergi1", "intranet - local"
                    'jobbeskoskrivCHK = "CHECKED"
                    'case else
                    'jobbeskoskrivCHK = ""
                    'end select
                   
                     %>
                    <input id="jobbeskoskriv" name="jobbeskoskriv" type="CHECKBOX" value="1" /> Overskriv jobbeskrivelse på job <br /><span style="color:#999999; padding:1px;">(##job_besk## - ##job_besk_slut##)</span>
                    <br />

                    <input onclick="gemtxt();" id="Submit2" type="submit" value=" Gem >> " />
                    <!-- jobbesk(); -->
                 
                  
                   
                   </form>
                   

                   

                
                     </td>
                   </tr>
                
                     
                   </table>
                   </div>
                  
                        
                        
                        
                        </td>
                        
                        <td valign=top>
                        <%
                        
                          ptop = 0
                          pleft = 0
                          pwdt = 274
                         
            
                         call filteros09(ptop, pleft, pwdt, "4) Eksport og Print <span style=""font-size:11px; color:#999999; font-weight:lighter;""><br>Gem samtidig i filarkiv</span>", 2, tdheight)
                        
                         %>
                        
        
        <!--                
		<a href="job_print.asp?&id=<=id%>&func=print&kid=<=kid%>" class=vmenu>Print</a>
		-->
		
		<table cellspacing=0 cellpadding=2 border=0>
		<!-- https://outzource.dk/timeout_xp/wwwroot/ver2_1/pdf/ -->
		
		<form action="job_make_pdf.asp?lto=<%=lto%>&id=<%=id%>&kid=<%=kid%>" target="_blank" method="post" id="mk_pdf">

        <%if lto = "synergi1" OR lto = "intranet - local" OR lto = "dencker" then
                gfpdfCHK = "CHECKED"
                else
                gfpdfCHK = ""
                end if %>

		<tr>
		    <td><input id="gempdf" type="button" style="font-size:9px;" value=" PDF >> " /> <!-- jobbeskPdf(); --><!-- onclick="gemtxtPdf();" -->
		    <input id="Checkbox3" name="gem_pdf_somfil" value="1" type="checkbox" <%=gfpdfCHK %> /> Gem samtidig fil i filarkiv.

              

            <input id="xjobbeskoskrivPdf" name="xjobbeskoskriv" type="hidden" value="0" />    
            <input id="FM_vis_indledning_pdf" name="FM_vis_indledning_pdf" type="hidden" /></td>
            <input id="FM_vis_afslutning_pdf" name="FM_vis_afslutning_pdf" type="hidden" /></td>
            <input id="pdf_skabelonnavn" name="pdf_skabelonnavn" value="" type="hidden" />
            <input id="pdf_doktype" name="pdf_doktype" value="<%=doktype %>" type="hidden" />
            <input id="pdf_foid" name="pdf_foid" value="" type="hidden" />
            </form>



            <form action="job_print.asp?id=<%=id%>&func=print&kid=<%=kid %>" method="post" target="_blank" id="mk_pnt">
	    
		    </tr>
		    <tr>
		    <td><!-- Main Print Form -->
                <input id="gempnt"  type="button" value=" Print >> " style="font-size:9px;" /><!-- jobbeskPr(); --><!-- onclick="gemtxtPrint();" -->
                <%if lto = "dencker" OR lto = "intranet - local" OR lto = "synergi1" OR lto = "essens" OR lto = "jttek" then
                gfpCHK = "CHECKED"
                else
                gfpCHK = ""
                end if %>
                <input id="xjobbeskoskrivPr" name="xjobbeskoskriv" type="hidden" value="0" />
                <input id="Checkbox2" name="printgemfil" value="1" type="checkbox" <%=gfpCHK %> /> Gem samtidig fil i filarkiv.
                 <input id="print_skabelonnavn" name="print_skabelonnavn" value="" type="hidden" />
                <input id="print_foid" name="print_foid" value="" type="hidden" />
                <input id="print_doktype" name="print_doktype" value="<%=doktype %>" type="hidden" />
                
                <!--<br />
                <br />
                <input id="Radio7" type="radio" CHECKED /> Gem på web, <input id="Radio8" type="radio" /> Gem på filserver 
                <br />Sti til filserver: 
                <input id="Text1" type="text" value="c:\\kunder\kundenavn\job" style="width:200px; font-size:9px;" /><br />&nbsp;
                -->
                <br />
                <span style="color:#999999; padding:1px;">(overskriver eksisterende fil med samme navn)</span>
                </td>
		</tr>
		

              
		</table>
		
		
		
		
		<br /><br />
		
	
		<b>Gem PDF ovenfor, og email til:</b><br />
		<%

        

        

      

		strSQL = "SELECT navn, email FROM kontaktpers WHERE kundeid = " & kid
		'Response.Write strSQL 
		'Response.Flush
		
		oRec.open strSQL, oConn, 3
        while not oRec.EOF

                
        
        Response.Write "<i>"& oRec("navn") & "</i>, <a href='mailto:"&oRec("email")&"&subject=Vedr.: "& strjobnr&"_"& tilbudsnr &"' class=vmenu>" & oRec("email") & "</a><br>"
        
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
                          <h4>Jobinfo og aktiviteter hentet fra det valgte job / skabelon:<span style="font-size:11px; font-weight:lighter; line-height:13px; color:#000000;"><br />Herunder har Du mulighed for at redigere i teksten inden Du gemmer eller udskriver dit dokument.<br />
                              Du kan også flette dokumentet med datavalg yderst til højre.
                                                                                          </span></h4>
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
		                
                        if levinfo <> "1" then
                        strLevHTMLedit = strLevHTML
                        else
                        strLevHTMLedit = ""
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
		                

                         if cint(ulevinfo) <> 1 then
		                ulevinfoHTMLedit = strUlevHTML
		                else
		                ulevinfoHTMLedit = ""
		                end if
		                
		                '*************************************************'
                        '************* Editor Indhold ********************'
                        '*************************************************'
                        if media <> "print" AND media <> "pdf" then
                        %>
                         <!-- sideskift markering -->
                         <div id="Div1" style="position:absolute; top:2000px; left:-10px; height:2px; border-top:1px darkred dashed; padding-left:20px; z-index:20000000;">
                         <span style="font-size:10px;">Sideskift: indsæt: ##breakpage##</span>
                         <img alt="" src="../ill/blank.gif" width="810" height="1" border="0"  />
                         </div>

                        <%end if
                        
                        strkundeHTMLedit = replace(strkundeHTMLedit, "##header_adr##", strHeader_adrHTML)
                        strkundeHTMLedit = replace(strkundeHTMLedit, "##header_dt##", formatdatetime(date(), 1))
                        
		                
	                    content = strkundeHTMLedit & strJobHTMLedit  & jbperansinfoHTMLedit & aktinfoHTMLedit & ulevinfoHTMLedit & strJobPrisHTMLedit & mileinfoHTMLedit & strHilsenHTML
            			
                        if cint(flet) = 1 then
                               
                               if levinfo <> "1" then
                                    skabelonTxt = replace(skabelonTxt, "##alt_adr##", strAlt_adrHTML) 
                               end if

                               if kinfo <> "1" then
                                    skabelonTxt = replace(skabelonTxt, "##header_adr##", strHeader_adrHTML) 
                                    skabelonTxt = replace(skabelonTxt, "##header_dt##", formatdatetime(date(), 1))
                               end if

                               
                                skabelonTxt = replace(skabelonTxt, "##job_jobnr##", ""& strjobnr &"")

                        end if


                        'if flet <> 0 then
                        'skabelonTxt = replace(skabelonTxt, "&lt;&lt;Budget&gt;&gt;", aktinfoHTMLedit)
                        'end if
			            
			            Set editorK = New CuteEditor
            					
			            editorK.ID = "FM_vis_indledning"
			            
			            
			            if cint(skabelonid) <> 0 then
                            
                            if cint(flet) = 2 then
                            editorK.Text = content
                            else
			                editorK.Text = skabelonTxt
			                end if

			            else
			                if cint(dokid) <> 0 then
                                if cint(flet_brodtxt) = 1 then
                                   
                                 
                                   
                                  '**** Henter ny centent fra job ****'
                                   dokTxt_from = instr(dokTxt, "##job_besk##")
                                   dokTxt_to = instr(dokTxt, "##job_besk_slut##")
                                   dokTxt_len = (dokTxt_to-dokTxt_from) 

                                    if dokTxt_from <> 0 AND dokTxt_to <> 0 then 
                                    dokTxtRplc = mid(dokTxt, dokTxt_from+12, dokTxt_len-12)
                                    dokTxtRplc = replace(dokTxtRplc, "jobpr16", "")
                                    end if



                                   '**** Henter ny centent fra job ****'
                                   cntTxt_from = instr(content, "##job_besk##")
                                   cntTxt_to = instr(content, "##job_besk_slut##")
                                   cntTxt_len = (cntTxt_to-cntTxt_from) 

                                    if cntTxt_from <> 0 AND cntTxt_to <> 0 then 
                                    contentRplc = mid(content, cntTxt_from+12, cntTxt_len-12)
                                    contentRplc = replace(contentRplc, "jobpr16", "")
                                    end if

                                    
                                    contentUse = replace(dokTxt, dokTxtRplc, contentRplc)
                                   'call gemJobBesk(content, 2, jobbeskTxthold)
                                   
                                   ''htmlparseCSV(dokTxtRplc)
                                   htmlreplace(dokTxtRplc)
                                   dokTxtRplc = htmlparseTxt 'htmlparseCSVtxt

                                   %>
                                   <div id="div_fltbrodtxt" style="position:absolute; background-color:#FFFFFF; padding:20px; top:550px; left:400px; width:550px; height:300px; border:10px #6CAE1C solid; z-index:500;">
                                   Brødtekst fra det dokument du ønsker at flette fra.<br />
                                   <b> Kopier og sæt ind hvor det passer i teksten:</b>&nbsp;&nbsp;&nbsp;&nbsp;<a href="#" id="a_fltbrodtxt" class=red>[Luk]</a><br /><br />
                                   <span style="width:510px; height:200px; border:1px #CCCCCC solid; padding:5px; overflow:auto;"><%=dokTxtRplc %></span>
                                   
                                   </div>
                                   <%
                                   'fletcontent_left = left(content, cntTxt_from+12)
                                   'fletcontent_right = right(content, cntTxt_to-192)

                                   'fletcontent = fletcontent_left&fletcontent_right

                                editorK.Text = content 'content 'dokTxtRplc 'contentRplc 'dokTxt 'contentUse 'jobbeskTxthold 'jobBeskTxt 'content 'jobBeskTxt 'jobBeskTxt 'content 'dokTxt' "der flettes brødteask"
                               

                                else
			                    editorK.Text = dokTxt
			                    end if
                            else
			                editorK.Text = content
			                end if
			            end if
			            
			            editorK.FilesPath = "CuteEditor_Files"
                        
                        select case lto
                        case "dencker", "jttek", "intranet - local" 'kun dencker må indsætte billeder
                        editorK.AutoConfigure = "Compact" 'Default compact
                        case else
			            editorK.AutoConfigure = "Minimal"
                        end select    
                        
            			'editorK.ConfigurationPath = "ver2_14/timereg/CuteEditor_Files/Configuration/full.config"
			            editorK.Width = 790
			            editorK.Height = 1480
			            editorK.Draw()
                                      
		                %>
		                
		                <br />
		                
		                <%
		                uTxt = "<b>Sideskift:</b><br>For at lave et sideskift indtastes et ##breakpage##"
						uWdt = 300
								
								call infoUnisport(uWdt, uTxt) 
		                 %>
		                
                    <br /><br /><b>Fast footer på print/PDF:</b>  
                    <!-- <br />
                    Vis som: 
                                <select id="vis_footer_header_tb" name="vis_footer_header_tb">
                                    <option value=1>Footer</option>
                                    <option value=2>Header</option>
                                    <option value=0>Vis ikke</option>
                                </select>
                                -->
                                
                    <!--Vis på side:  <select id="vis_footer_header_side" name="vis_footer_header_side">
                                    <option value=0>Alle sider</option>
                                    <option value=1>Sidste side</option>
                                    <option value=2>Første side</option>
                                   </select>
                                   -->
                          
                         
                    
                    <br />
                    <textarea id="FM_vis_afslutning" name="FM_vis_afslutning" cols="97" rows="4"><%=afslutningVal %></textarea>
                    <br /><input id="FM_gem_afs" name="FM_gem_afs" type="checkbox" value="1" />Gem denne footer på alle kunder. (overskriver eksisterende)<br />
                    <!--<br /><input id="Submit1" type="submit" value=" Gem / print >> " />-->
                    
		                
		                </td>
		                </form>
		                <form action="job_print.asp?id=<%=id%>&kid=<%=kid %>&tb_skabeloner=<%=skabelonid %>&formsbm=1" method="post">




		                <td valign=top style="padding:2px; border:1px #cccccc solid; background-color:#FFFFFF;">
		                
		                 <%
                         '*******************************************************'
                         '** Datavalg
                         '*******************************************************'


                          tdheight = 1550
                          ptop = 0
                          pleft = 0
                          pwdt = 300
            
                         call filteros09(ptop, pleft, pwdt, "2) Datavalg / Flet <label style=""font-size:11px; color:#000000; font-weight:lighter; line-height:13px;""><br>Flet datavalg med valgte skabelon ell. genindlæs jobdata med nedenstående datavalg </label>", 4, tdheight)
                        
                         %>
		                
		                
		               
		               
		               <%if kinfo <> "1" then
		               kinfoCHK0 = "CHECKED"
		               kinfoCHK1 = ""
		               else
		               kinfoCHK0 = ""
		               kinfoCHK1 = "CHECKED"
		               end if %>
		                
		                    <b style="background-color:#D6dff5;"><u>Header</u></b> <span style="color:#999999;"> fletkoder, adr.: ##header_adr##,<br /> dato: ##header_dt##</span> <br />
                            <input id="Radio1" name="kinfo" type="radio" value="0" <%=kinfoCHK0 %> /> ja <input id="Radio1" name="kinfo" type="radio" value="1" <%=kinfoCHK1 %> /> nej
                               <br />Vælg: <select name="FM_kundeid" style="font-size:9px; width:200px;">
                              <option value="0">Ingen</option>
                            <%
                            
                            strSQLlev = "SELECT kkundenavn, kkundenr, kid, kp.navn AS kpersnavn, kp.id AS kpid FROM kunder AS k "_
                            &" LEFT JOIN kontaktpers AS kp ON (kundeid = kid) WHERE kid <> 0 ORDER BY kkundenavn, kp.navn "

                            lastKid = 0
                            oRec3.open strSQLlev, oConn, 3

                            k = 0
                            while not oRec3.EOF

                            if cint(kunde_ids(0)) = oRec3("kid") then
                            kIDSEL = "SELECTED"
                            else
                            kIDSEL = ""
                            end if
                            
                            if lastKid <> oRec3("kid") then
                                if k <> 0 then%>
                                <option value="<%=oRec3("kid")%>"></option>
                                <%end if %>

                            <option value="<%=oRec3("kid")%>,0" <%=kIDSEL%>><%=oRec3("kkundenavn") %> (<%=oRec3("kkundenr") %>)</option>
                            <%
                            end if

                             if cint(kunde_ids(1)) = oRec3("kpid") then
                            kpIDSEL = "SELECTED"
                            else
                            kpIDSEL = ""
                            end if

                            %>
                              <option value="<%=oRec3("kid")%>,<%=oRec3("kpid") %>" <%=kpIDSEL%>><%=oRec3("kpersnavn") %></option>

                            <%


                            lastKid = oRec3("kid")
                            k = k + 1
                            oRec3.movenext
                            wend
                            oRec3.close
                            %>

                          
                            </select>

		               
		                
		                

                         <%if levinfo <> "1" then
		               levinfoCHK0 = "CHECKED"
		               levinfoCHK1 = ""
		               else
		               levinfoCHK0 = ""
		               levinfoCHK1 = "CHECKED"
		               end if %>
		                 <br /><br /><br />
		                    <b style="background-color:#D6dff5;"><u>Alt. Adresse</u></b><span style="color:#999999;"> fletkode: ##alt_adr## </span><br />
                            <input id="Radio7" name="levinfo" type="radio" value="0" <%=levinfoCHK0 %> /> ja <input id="Radio8" name="levinfo" type="radio" value="1" <%=levinfoCHK1 %> /> nej
		                
		                    

                            
                            <br />Vælg: <select name="FM_levid" style="font-size:9px; width:200px;">
                              <option value="0">Ingen</option>
                            <%
                            
                            strSQLlev = "SELECT kkundenavn, kkundenr, kid, kp.navn AS kpersnavn, kp.id AS kpid FROM kunder AS k "_
                            &" LEFT JOIN kontaktpers AS kp ON (kundeid = kid) WHERE kid <> 0 ORDER BY kkundenavn, kp.navn "

                            lastKid = 0
                            oRec3.open strSQLlev, oConn, 3

                            k = 0
                            while not oRec3.EOF

                            if cint(lev_ids(0)) = oRec3("kid") then
                            levIDSEL = "SELECTED"
                            else
                            levIDSEL = ""
                            end if
                            
                            if lastKid <> oRec3("kid") then
                                
                                if k <> 0 then%>
                                <option value="<%=oRec3("kid")%>"></option>
                                <%end if %>

                            <option value="<%=oRec3("kid")%>,0" <%=levIDSEL%>><%=oRec3("kkundenavn") %> (<%=oRec3("kkundenr") %>)</option>
                            <%
                            end if

                            if cint(lev_ids(1)) = oRec3("kpid") then
                            levIDkpSEL = "SELECTED"
                            else
                            levIDkpSEL = ""
                            end if

                            %>
                              <option value="<%=oRec3("kid")%>,<%=oRec3("kpid") %>" <%=levIDkpSEL%>><%=oRec3("kpersnavn") %></option>

                            <%


                            lastKid = oRec3("kid")
                            k = k + 1
                            oRec3.movenext
                            wend
                            oRec3.close
                            %>

                          
                            </select>

                        <br />
                        <br />



		                <%if jbinfo <> "1" then
		               jbinfoCHK0 = "CHECKED"
		               jbinfoCHK1 = ""
		               else
		               jbinfoCHK0 = ""
		               jbinfoCHK1 = "CHECKED"
		               end if %>
		                <br /><br />
		                 <b style="background-color:#D6dff5;"><u>Vis jobnavn og beskrivelse</u></b> <span style="color:#999999;">fletkoder: ##job_jobnr##</span><br />
                         
		                  <input id="Radio3" type="radio" name="jbinfo" value="0" <%=jbinfoCHK0 %> /> ja <input id="Radio3" name="jbinfo" type="radio" value="1" <%=jbinfoCHK1 %> /> nej<br />
                         <br /> 
		                <%=strChkjobHTML%>
		                
		                
		                <%if cint(jbprisinfo) <> 1 then
		               jbprisinfoCHK0 = "CHECKED"
		               jbprisinfoCHK1 = ""
		               else
		               jbprisinfoCHK0 = ""
		               jbprisinfoCHK1 = "CHECKED"
		               end if %>
		               <br /><br /><br />
		                <b style="background-color:#D6dff5;"><u>Vis samletpris, stor</u></b>
		                   <input id="Radio2" type="radio" name="jbprisinfo" value="0" <%=jbprisinfoCHK0 %> /> ja <input id="Radio4" name="jbprisinfo" type="radio" value="1" <%=jbprisinfoCHK1 %> /> nej
                            <hr />
		               <%=strJobPrisHTML%>
		                
                       
                        <br />
                        <br /><br />

		                
		                <%if jbperansinfo <> "1" then
		               jbperansinfoCHK0 = "CHECKED"
		               jbperansinfoCHK1 = ""
		               else
		               jbperansinfoCHK0 = ""
		               jbperansinfoCHK1 = "CHECKED"
		               end if %>
		              
		                <b style="background-color:#D6dff5;"><u>Vis periode og job ansvarlige</u></b>
		                     <input id="Radio5" type="radio" name="jbperansinfo" value="0" <%=jbperansinfoCHK0 %> /> ja <input id="Radio6" name="jbperansinfo" type="radio" value="1" <%=jbperansinfoCHK1 %> /> nej
		              <hr />
		                <%=strJobinfoHTML%>
		                

                       <br />
                       <br />
                       <br />

		                
		                
		                <%'**** Aktiviteter ***' %> 
		                <%=strChkHTML%>
		                


                        <%'**** Salgs omkostninger *' %>
                         <%if cint(ulevinfo) <> 1 then
		               uLevinfoCHK0 = "CHECKED"
		               uLevinfoCHK1 = ""
		               else
		               uLevinfoCHK0 = ""
		               uLevinfoCHK1 = "CHECKED"
		               end if %>
		                  <br /><br />
		                <b style="background-color:#D6dff5;"><u>Vis Salgsomkostninger</u></b>
		                 <input id="Radio12" type="radio" name="ulevinfo" value="0" <%=uLevinfoCHK0 %> /> ja <input id="Radio13" name="ulevinfo" type="radio" value="1" <%=uLevinfoCHK1 %> /> nej
		                


                        
		                <%=strChkUlevHTML %>
		                
		                
		                
		                 <%if cint(mileinfo) <> 1 then
		               mileinfoCHK0 = "CHECKED"
		               mileinfoCHK1 = ""
		               else
		               mileinfoCHK0 = ""
		               mileinfoCHK1 = "CHECKED"
		               end if %>
		                  <br /><br />
		                <b style="background-color:#D6dff5;"><u>Vis milepæle</u></b>
		                 <input id="Radio9" type="radio" name="mileinfo" value="0" <%=mileinfoCHK0 %> /> ja <input id="Radio10" name="mileinfo" type="radio" value="1" <%=mileinfoCHK1 %> /> nej
		              
		                <%=strMileHTML %>
		                
		                
		                <%=strHilsenHTML %>
		                
		                <!-- filteros09 slut -->
		                </td>
		                </tr>
		                </table>
                        
                        </div>
		                
		                <table width=100% cellpadding=0 cellspacing=0 border=0>
		                 <tr><td> <br />
                         <input id="Checkbox4" type="radio" name="FM_flet" value="1" <%=fletCHK1 %> /> Flet skabelon med ovenstående datavalg.<br />
                         <input id="Radio11" type="radio" name="FM_flet" value="2" <%=fletCHK2 %> /> Indlæs jobdata igen med ovenstående datavalg.
                         <!--<input id="Radio14" type="radio" name="FM_flet" value="3" <%=fletCHK3 %> /> Flet dokument med ovenstående datavalg (beholder brødteskt fra dok.) --></td></tr>
                         <tr>
                         <td align=center><br />

                      
                     <input id="Submit9" type="submit" value=" Gen-indlæs / Flet >> " /><br />&nbsp;
		                
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
            
            select case lto
            case "wowern"
                
                    prdtop_1 = "865"
                    prdlft_1 = "0"
                    
                    prdtop_2 = "1365"
                    prdlft_2 = "0"

            case "jttek", "intranet - local"
                
                    prdtop_1 = "705"
                    prdlft_1 = "0"
                    
                    prdtop_2 = "1365"
                    prdlft_2 = "0"
                   
                
            case "essens"
                    
                    prdtop_1 = "865"
                    prdlft_1 = "0"
                    
                    prdtop_2 = "1465"
                    prdlft_2 = "0"
                    
            case else    
            
             
                    prdtop_1 = "865"
                    prdlft_1 = "0"
                    
                    prdtop_2 = "1465"
                    prdlft_2 = "0"
                   
            end select
            



            
            
            if instr(indledningVal, "##breakpage##") <> 0 then
            
            s1 = instr(indledningVal, "##breakpage##") - 1
            p1_indledningVal = left(indledningVal, s1)
            len_p1val = len(p1_indledningVal)
            len_indledningVal = len(indledningVal)

            
            
            '*** Footer side 1 ***
            p1_indledningVal = p1_indledningVal & "<div id=""printpdf_1"" style=""position:absolute; border:0px #999999 solid; top:"&prdtop_1&"px; left:"&prdlft_1&"px; z-index:20000000;"">"
            p1_indledningVal = p1_indledningVal & "<center>"& afslutningsTXT & "</center></div>"
           
            
            '*** Footer side 2 ***
            p2_indledningVal = ""
            p2_indledningVal = p2_indledningVal & "<div id=""printpdf_2"" style=""position:absolute; top:"&prdtop_2&"px; left:"&prdlft_2&"px; z-index:20000000;"">"
            p2_indledningVal = p2_indledningVal & "<center>"& afslutningsTXT & "</center></div>"
            p2_indledningVal = p2_indledningVal & "<img src=""../ill/blank.gif"" height=1 width=1 style='page-break-before:always'><br><br>"& right(indledningVal, len_indledningVal-(len_p1val+2))

            
            
            x = 2
            else
            
           

            
            '*** Footer side 1 ***
            p1_indledningVal = indledningVal
            p1_indledningVal = p1_indledningVal & "<div id=""printpdf_1"" style=""position:absolute; top:"&prdtop_1&"px; left:"&prdlft_1&"px; z-index:20000000;"">"
            p1_indledningVal = p1_indledningVal & "<center>"& afslutningsTXT & "</center></div>"
            
            p2_indledningVal = ""

           
            x = 1
            
            end if
            
            
            for p = 1 to x 
            
            
            if p = 1 then
            Response.Write p1_indledningVal 
            else
            Response.Write p2_indledningVal
            end if 
            
            
            
            %>
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
