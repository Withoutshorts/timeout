<%response.buffer = true%>


<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="NO-CACHE">
<META HTTP-EQUIV="EXPIRES" CONTENT="Mon, 22 Jul 2002 11:12:01 GMT">
    

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/smiley_inc.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->


<%
 
 public minThisDIFF
 function beregnSek(loginTid, logudTid, logudHH)


 'Response.Write loginTid & " - " & logudTid & "<br>"

 '**** Beregner Minutter mellem logind og logud ***********
if len(loginTid) <> 0 AND len(logudTid) <> 0 then
					                    
		'*** Er der angivet et logud ****'
		if logudHH <> "99" then
                     					
		'loginTidAfr = left(formatdatetime(loginTid, 3), 5)
		'logudTidAfr = left(formatdatetime(logudTid, 3), 5)
                    		
		minThisDIFF = datediff("s", loginTid, logudTid)/60
                    					
		else
                    					
		minThisDIFF = 0
		end if

       'Response.Write minThisDIFF & "<br>"
       'Response.flush
					                    
end if
'***

 end function

                            

 function nyePauser(p, thison, startval, hdn)


            
             if startval = 1 then ' 1: Standard pauser forvalgt / else 0 min pause forvalgt

             if p = 1 then
             caseVal = stPauseLic_1
             else
             caseVal = stPauseLic_2
             end if

             else

             caseVal = 0

             end if



            if cint(hdn) = 0 then

            
            if (thison = 1) then
             %>
             <br /><br /><b>Pause <%=p %>:</b> 

            <select name="p<%=p %>" id="p<%=p %>" style="font-family:arial; font-size:9px;">
            
             <% 
             
            
            
             
            selM0 = "SELECTED"
             selM5 = ""
             selM10 = ""
             selM15 = ""
             selM20 = ""
             selM25 = ""
             selM30 = ""
             selM35 = ""
             selM40 = ""
             selM45 = ""
             selM50 = ""
             selM55 = ""

             Select case caseVal
             case 0
             selM0 = "SELECTED"
             case 5
             selM5 = "SELECTED"
             case 10
             selM10 = "SELECTED"
             case 15
             selM15 = "SELECTED"
             case 20
             selM20 = "SELECTED"
             case 25
             selM25 = "SELECTED"
             case 30
             selM30 = "SELECTED"
             case 35
             selM35 = "SELECTED"
             case 40
             selM40 = "SELECTED"
             case 45
             selM45 = "SELECTED"
             case 50
             selM50 = "SELECTED"
             case 55
             selM55 = "SELECTED"
             end select
             
               
             
             %>
             <option value="0" <%=selM0%>>0 min</option>
             <option value="5" <%=selM5%>>5 min</option>
             <option value="10" <%=selM10%>>10 min</option>
             <option value="15" <%=selM15%>>15 min</option>
              <option value="20" <%=selM20%>>20 min</option>
              <option value="25" <%=selM25%>>25 min</option>
              <option value="30" <%=selM30%>>30 min</option>

               <option value="35" <%=selM35%>>35 min</option>
             <option value="40" <%=selM40%>>40 min</option>
              <option value="45" <%=selM45%>>45 min</option>
              <option value="50" <%=selM50%>>50 min</option>

               <option value="55" <%=selM55%>>55 min</option>
              
            </select> min.<br />
        
            <br />
            Kommentar pause <%=p %>:
	        <br /><textarea id="FM_komm_p<%=p %>" name="FM_komm_p<%=p %>" style="font-size:9px; font-family:arial; width:350px; height:46px;"></textarea>
	        <%
            end if


            else 'hdn = automatisk tilføj standard pause ved opret logind. Vises hidden, dvs ikke mulighed for at redigere

                %>
                <input type="hidden" name="p<%=p %>" value="<%=caseVal %>"/> 
                <input type="hidden" name="FM_komm_p<%=p %>"/> 
                <%

            end if
            
            end function
            


sub crmaktstatheader
%>
<table border=0 cellpadding=0 cellspacing=0 width="450">
	<tr>
	<td valign="top" width="163"><img src="../ill/logo_bg.gif" width="163" height="53" alt="" border="0"></td>
	<td valign="bottom"><b>Timeout Kontrolpanel - Stempelur indstillinger</b><br>
	Tilføj, fjern eller rediger Stempelur indstillinger.</td>
	</tr>
	</table><br>
<%
end sub


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	if len(request("lastid")) <> 0 then
	lastID = request("lastid")
	else
	lastID = 0
	end if 
	
	if len(request("showonlyone")) <> 0 then
	showonlyone = request("showonlyone")
	else
	showonlyone = 0
	end if
	
	if len(request("medarbSel")) <> 0 then
	medarbSel = request("medarbSel")
	else
	medarbSel = 0
	end if
	
	if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
	    if request.Cookies("stat")("progrp") <> "" then
	    progrp = request.Cookies("stat")("progrp")
	    else
	    progrp = 0
	    end if
	end if
	
	response.Cookies("stat")("progrp") = progrp
	
	'*** Rettigheder på den der er logget ind **'
	medarbid = session("mid")
	
	if len(request("FM_medarb")) <> 0 then
	
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
	    
	    if request.Cookies("stat")("intmids") <> "" then
	    thisMiduse = request.Cookies("stat")("intmids") 
	    intMids = split(thisMiduse, ", ")
	    else
	    thisMiduse = session("mid") 
	    intMids = split(thisMiduse, ", ")
	    end if
	   
	end if
	
	
	response.Cookies("stat")("intmids") = thisMiduse
	
	if showonlyone = 1 then
	showalle = 0
	medarbSQLkri = " AND m.mid = " & medarbSel
	else
	showalle = 1
	medarbSQLkri = " AND m.mid <> 0 "
	end if
	
	if len(request("hidemenu")) <> 0 then
	hidemenu = request("hidemenu")
	else
	hidemenu = 0
	end if
	
	rdir = request("rdir")

    'Response.write "rdir: "& rdir
    'Response.end

		
	thisfile = "stempelur"
	
	select case func
	case "sletlogind"
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<%
	slttxt = "<b>Slet dette logind?</b><br />"_
	&"Du er ved at slette et logind. Er dette korrekt?"
	slturl = "stempelur.asp?menu="&request("menu")&"&func=sletlogindok&id="&id&"&medarbSel="&medarbSel&"&showonlyone="&showonlyone&"&hidemenu="&hidemenu&"&rdir="&rdir
	
	call sltquePopup(slturl,slttxt,slturlalt,slttxtalt,20,10)
	
	case "sletlogindok"
	
    thisMid = 0
    thisLgNavn = ""
    thisLgNr = ""
    strSQL = "SELECT mid, login, logud, id FROM login_historik WHERE id = "& id &""
    oRec.open strSQL, oConn, 3
	while not oRec.EOF 
		
	        thisMid = oRec("mid")
            thisLgNavn = oRec("login") & " - "& oRec("logud")
		    thisLgNr = oRec("id")

	oRec.movenext
	wend
	oRec.close

    strSQL = "SELECT mid, mnavn, mnr FROM medarbejdere WHERE mid = "& thisMid &"" 
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
		
		'*** Indsætter i delete historik ****'
	    call insertDelhist("kom_ga", thisMid, thisLgNr, thisLgNavn, session("mid"), session("user"))
		
	oRec.movenext
	wend
	oRec.close


	
	strSQLdel = "DELETE FROM login_historik WHERE id = "& id
	oConn.execute(strSQLdel)
	
	'select case rdir
	'case "week"
	'Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
    'case "treg"
    'Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
	'''Response.Write("<script language=""JavaScript"">window.opener.location.href('timereg_akt_2006.asp?showakt=3');</script>")
    'case else
	'Response.Write("<script language=""JavaScript"">window.opener.location.href('stempelur.asp?menu=stat&func=stat&medarbSel="&medarbSel&"&showonlyone="&showonlyone&"&hidemenu="&hidemenu&"');</script>")
    'end select						
	
    Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
	Response.Write("<script language=""JavaScript"">window.close();</script>")
	
	
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:10; visibility:visible;">
	<%call crmaktstatheader%>
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> en Stempelur indstilling. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><a href="stempelur.asp?menu=tok&func=sletok&id=<%=id%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	case "sletok"


   

	'*** Her slettes en stempelurs indstilling ***
	oConn.execute("DELETE FROM stempelur WHERE id = "& id &"")

    'Response.write "her"
    'Response.end

	Response.redirect "stempelur.asp?menu=tok&shokselector=1"
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		<%
		errortype = 8
		call showError(errortype)
		
		else
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		if len(request("FM_faktor")) <> 0 then
		intFaktor = replace(request("FM_faktor"), ",", ".")
		else
		intFaktor = 0
		end if
		
		if len(request("FM_minimum")) <> 0 then
		intMinimum = replace(request("FM_minimum"), ",", "") 
		intMinimum = replace(intMinimum, ".", "")
		else
		intMinimum = 0
		end if
		
		if len(request("FM_forvalgt")) <> 0 then
		intForvalgt = request("FM_forvalgt")
		else
		intForvalgt = 0
		end if
		
			if intForvalgt = 1 then
			oConn.execute("UPDATE stempelur SET forvalgt = 0")
			end if
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO stempelur (navn, editor, dato, faktor, forvalgt, minimum) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& intFaktor &", "& intForvalgt &", "& intMinimum &")")
		else
		oConn.execute("UPDATE stempelur SET navn ='"& strNavn &"', "_
		&" forvalgt = "& intForvalgt &", editor = '" &strEditor &"', dato = '" & strDato &"', faktor = "& intFaktor &", minimum = "& intMinimum &" WHERE id = "&id&"")
		end if
		
        
		
		Response.redirect "stempelur.asp?menu=tok&shokselector=1"
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	
	else
	strSQL = "SELECT navn, editor, dato, faktor, forvalgt, minimum FROM stempelur WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	intFaktor = oRec("faktor")
	intForvalgt = oRec("forvalgt")
	intMinimum = oRec("minimum")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:10; visibility:visible;">
	<%call crmaktstatheader%>
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr><form action="stempelur.asp?menu=tok&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
    	<td valign="top" colspan="2"><font class="pageheader">Kontrolpanel | Stempelur indstillinger | <%=varbroedkrumme%></font></td>
	</tr>
	<%if dbfunc = "dbred" then%>
	<tr>
		<td colspan="2" valign="bottom" style="height:30;">Sidst opdateret den <b><%=formatdatetime(strDato, 0)%></b> af <b><%=strEditor%></b></td>
	</tr>
	<%end if%>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td><b>Navn:</b></td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td><b>Omregnings faktor:</b></td>
		<td><input type="text" name="FM_faktor" value="<%=intFaktor%>" size="4" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td valign=top style="padding-top:3px;"><b>Minimums indtastning:</b></td>
		<td><input type="text" name="FM_minimum" value="<%=intMinimum%>" size="4" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;<b>(Angivet i hele min.)</b><br>
		<font class=lillesort>Hvis minimum er sat til 60 min.<br> 
		vil alle login's på mindre end 60 min. blive takseret til 1 time.</font></td>
	</tr>
	<tr>
		<td colspan=2>
		<%
		if intForvalgt = 1 then
		chk = "CHECKED"
		else
		chk = ""
		end if
		%>
		<br><input type="checkbox" name="FM_forvalgt" value="1" <%=chk%>>&nbsp;<b>Denne indstilling skal stå som forvalgt ved login.</b></td>
	</tr>
	<tr>
		<td colspan="2"><br><br><img src="ill/blank.gif" width="100" height="1" alt="" border="0"><input type="image" src="../ill/<%=varSubVal%>.gif"></td>
	</tr>
	</form>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	
	<%
	case "dbloginhist"
	'******************************************************************************************************
	'******************************************************************************************************

    '*** Indlæser fra medarberjderes egen logind historik liste med åbne felter og fra rediger/opret ny indtastning

    '******************************************************************************************************
    '******************************************************************************************************

	'**** ignorer tidsinterval / fast st. tid *********
	
	call ersmileyaktiv()
	
	stpause = -1
    stpause2 = -1
	
	'Response.Write request("id") & "<br>"
    'Response.Write request("mid") & "<br>"
    'Response.Write request("logindato") & "<br>"

    'Response.end



	ids = split(request("id"), ",")
	useMid = split(request("mid"), ",")

   

    'Response.Write request("logindato")
    'Response.end

    if len(trim(request("logindato"))) = 0  then 
    %>
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <%
    useleftdiv = "c"
    errortype = 60
	call showError(errortype)
	Response.end

    end if

    loginDato = split(request("logindato"), ",")

    '**loginhist variable ***'
    usemrn = request("usemrn")
    varTjDatoUS_man = request("varTjDatoUS_man")
	varTjDatoUS_son = request("varTjDatoUS_son")
    nomenu = request("nomenu")

    lnk = "usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&varTjDatoUS_son="&varTjDatoUS_son&"&rdir="&rdir&"&nomenu="&nomenu
	
   
	

	
	'*** Logind og logud tider ****'
    'Response.write request("FM_login_hh") & "<hr>"
    login_hh = replace(request("FM_login_hh"), ",","")
	login_mm = replace(request("FM_login_mm"), ",","")
    logud_hh = replace(request("FM_logud_hh"), ",","")
	logud_mm = replace(request("FM_logud_mm"), ",","")

 
    login_hh = split(login_hh, " #")
	login_mm = split(login_mm, " #")
	logud_hh = split(logud_hh, " #")
	logud_mm = split(logud_mm, " #")
	

	'*** Logind og logud tider Oprindelig ****'
	oprlogin_hh = split(request("oprFM_login_hh"), ",")
	oprlogin_mm = split(request("oprFM_login_mm"), ", ")
	
	oprlogud_hh = split(request("oprFM_logud_hh"), ",")

    'Response.write "request(oprFM_logud_mm) : "& request("oprFM_logud_mm") & "<br>" 
    'Response.end


	oprlogud_mm = split(request("oprFM_logud_mm"), ", ")
	
    kommThis = replace(request("FM_kommentar"), "#,", "#")
	kommThis = split(kommThis, ", #")
    'Response.Write "K "& request("FM_kommentar")
	'Response.end
	
    'Response.write "stru:"&  request("FM_stur")
    'Response.end
	stur = split(request("FM_stur"), ",")
	
	ipn = request.servervariables("REMOTE_ADDR")
	

    '*** Manuelle pauser '******
	if len(request("p1")) <> 0 then
    p1 = request("p1")
    else
    p1 = -1
    end if
    
    if len(request("p2")) <> 0 then
    p2 = request("p2")
    else
    p2 = -1
    end if
    
    p1on = request("p1on")
    p2on = request("p2on")
   
        'Response.Write "request(p1on) " & request("p1on")
        'Response.Write "<br>request(p2on) " & request("p2on")
        'Response.Write "<br>p2on " & p2on
        'Response.Write "<br>request(p1) " & request("p1")
        'Response.Write "<br>request(p2) " & request("p1")
        'Response.Write "<br>request(p2) " & request("p2")
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="inc/isint_func.asp"-->
	<%
	
	                function SQLBless2(s)
				    dim tmp
				    tmp = s
				    tmp = replace(tmp, "'", "''")
				    SQLBless2 = tmp
				    end function
	
	
	
	'**** Pauser *****
	p1_komm = SQLBless2(trim(request("FM_komm_p1")))
    p2_komm = SQLBless2(trim(request("FM_komm_p2")))
    
    

	useleftdiv = "c"
	
	
	
	for a = 0 to UBOUND(ids) 
	
	
	        if len(trim(loginDato(a))) = 0 OR len(trim(login_hh(a))) = 0 OR len(trim(login_mm(a))) = 0 then '*** Spring over / Ignorer indtastning **'
		
            '** Logind = 0 er ok ***

                    '** Tjekker om der er tastet ind i logud felter eller logind min. **'
                    if len(trim(login_mm(a))) <> 0 OR len(trim(logud_mm(a))) <> 0 OR len(trim(logud_hh(a))) <> 0 then 


			        errortype = 60
			        call showError(errortype)
			        Response.end

                    end if
			
			else
			
			
            '*** Logud = 0 er ok ****
			
            '**Tjekker om logud er indstastet korrekt, hvis de er tastet ****'
            if len(trim(logud_hh(a))) <> 0 then 
			    call erDetInt(trim(logud_hh(a)))
			    if cint(isInt) > 0 OR instr(logud_hh(a), ".") <> 0 then
				
				    errortype = 60
				    call showError(errortype)
				    Response.end
			
			    isInt = 0
                end if
            end if

            if len(trim(logud_mm(a))) <> 0 then 
			    call erDetInt(trim(logud_mm(a)))
			    if cint(isInt) > 0 OR instr(logud_mm(a), ".") <> 0 then
				
				    errortype = 60
				    call showError(errortype)
				    Response.end
			
			    isInt = 0
                end if
            end if
			
			'else
					
					if isdate(trim(loginDato(a))) <> true then 
					'OR cint(len(trim(loginDato(a)))) <> 10 then
					
					'Response.Write "loginDato(a)" & loginDato(a) & " len:" & cint(len(trim(loginDato(a))))
					'Response.end
					
					
                    errortype = 66
					call showError(errortype)
					Response.end    
					        
					       
					
					else
			        
			        
					'call erDetInt(trim(logud_mm(a))) 
					'if cint(isInt) > 0 OR instr(trim(logud_mm(a)), ".") <> 0 OR logud_mm(a) > 59 then
						
					'	errortype = 58
					'	call showError(errortype)
					'	Response.end
						
                    'isInt = 0
					
					'else
					        
					        'Response.Write "login_hh(a) " & login_hh(a) & "<br>"
							'Response.flush
					        
							call erDetInt(trim(login_hh(a))) 

                         

                             if cint(isInt) > 0 OR instr(login_hh(a), ".") <> 0 then

                            	errortype = 59
								call showError(errortype)
								Response.end
							
							isInt = 0

                            end if


							if login_hh(a) > 23 then
			
								
								errortype = 59
								call showError(errortype)
								Response.end
							
							isInt = 0
							
							else
							
									call erDetInt(trim(login_mm(a))) 
									if isInt > 0 OR instr(login_mm(a), ".") <> 0 OR login_mm(a) > 59 then
										
										errortype = 59
										call showError(errortype)
										Response.end
									
									isInt = 0
									
									else
					
					
					loginHH = left(trim(login_hh(a)), 2)
					loginMM = left(trim(login_mm(a)), 2)	
	                
	                '** Tjekker om der er angivet et logud tidspunkt ***'
	                if len(trim(logud_hh(a))) <> 0 AND len(trim(logud_mm(a))) <> 0 then
					logudHH = left(trim(logud_hh(a)), 2)
					logudMM = left(trim(logud_mm(a)), 2)

                    'Response.Write "her 5<br>"
					else
					logudHH = "99" 'loginHH '"99"
					logudMM = "99" 'loginMM '"99"

                    'Response.Write "her 6<br>"
					end if
					
					oprloginHH = left(trim(oprlogin_hh(a)), 2)
					oprloginMM = left(trim(oprlogin_mm(a)), 2)	
	
					oprlogudHH = left(trim(oprlogud_hh(a)), 2)
					oprlogudMM = left(trim(oprlogud_mm(a)), 2)
					
					
					
					'****** Kommentar **************'
					komm = SQLBless2(trim(kommThis(a)))
                    komm = replace(komm, "#", "")

					'komm = SQLBless2(trim(request("FM_kommentar_"&a&"")))
					'komm = "xx"
                    'Response.Write "  a:" & a & request("FM_kommentar_"&a&"") & "<br>"
					
					
					
					'*** Tjekker logind om det ligger i ignorer periode ***'
					call stpauser_ignorePer(loginDato(a),loginHH,loginMM)
					
					loginTid = loginTid
					logudTid = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) & " " & logudHH &":"& logudMM & ":00"
					'Response.Write logudTid &"<br>"
					
					'******************************************
					
					
					
					
					
					'************ ?? ***********
					if a = 0 then
					loginTidp = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) & " 00:00:00"
					logudTidp = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) & " 00:00:00"
					loginDTp = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a))
					end if
					'************************
					
					
					
					            
					            
					        call beregnSek(loginTid, logudTid, logudHH)
					        
					        
					        '*** er logud tid mindre end logind tidspunkt **'
					        if minThisDIFF < 0 then
					        
					            'errortype = 139
								'call showError(errortype)
								'Response.end
					        
					            
                                'minThisDIFF = abs(minThisDIFF)

                                '** Hvis minThisDIFF < 0 må logud være fortaget efter midnat hvorfor der lægges et døgn til loguddatoen ***'
                                logudDatoAddOneD = dateadd("d", 1, loginTid)
                                logudTid = year(logudDatoAddOneD) &"/"& month(logudDatoAddOneD)&"/"& day(logudDatoAddOneD) & " " & logudHH &":"& logudMM & ":00"
					            

                            call beregnSek(loginTid, logudTid, logudHH)
                            
                            minThisDIFF = abs(minThisDIFF)   

					        end if
					        'else
					        
					        
					        
					        '** Er logud tid ændret (Ved rediger) ?? **************'
					        '*** Eller er det første logud på dette logind? *******'    
					        logudrettetJa = 0
					        manuelt_afsluttet = manuelt_afsluttet
					        tiderRettet = 0
					       
					       'Response.Write "manuelt_afsluttet: " & manuelt_afsluttet & "<br>"
					        
					        strSQLA = "SELECT logud, login, manuelt_afsluttet, login_first, logud_first FROM login_historik WHERE logud IS NOT NULL AND id = " & ids(a)
					        
                            oprLogud = ""
                            oprLogin = ""
                            logudfirstWasNull = 0

                            'Response.Write strSQLA
                            'Response.flush
                            
                            oRec.open strSQLA, oConn, 3
                            if not oRec.EOF then
                            'Response.Write "first: "& oRec("logud_first")
                            'Response.flush
                                
                                if isNull(oRec("logud_first")) <> true then
                                oprLogud = formatdatetime(oRec("logud_first"), 2) &" "& formatdatetime(oRec("logud_first"), 3)
                                else
                                oprLogud = logudTid
                                logudfirstWasNull = 1
                                end if

                                if isNull(oRec("login_first")) <> true then
                                oprLogin = formatdatetime(oRec("login_first"), 2) &" "& formatdatetime(oRec("login_first"), 3)
                                else
                                oprLogin = loginTid
                                end if

                                logudrettetJa = 1
                                
                                'Response.Write "<br>her 7 oprLogud: "& oprLogud & "<br>"

                               
                                'if oRec("manuelt_afsluttet") < 3 then
                                manuelt_afsluttet = oRec("manuelt_afsluttet")
                                'else
                                'manuelt_afsluttet = 3
                                'end if
                            end if
                            oRec.close
                            
                            'Response.Write "manuelt_afsluttet: " & manuelt_afsluttet & "<br>"
					       
                            
                                        if cint(logudrettetJa) = 1 then
                                        len_oprLogud = len(oprLogud)
                                
                                            if len_oprLogud <> 0 then
                                            oprLogud_left = left(oprLogud, len_oprLogud- 3)
                                            oprLogud = oprLogud_left &":00"
                                            
                                            'Response.Write "her 11<br>"
                                            else
                                             
                                            oprLogud = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) &" "& oprlogud_hh(a) &":"& oprlogud_mm(a) &":00"
                                            'Response.Write "her 2"
                                            end if
                            
                            
                                            len_oprLogin = len(oprLogin)

                                            if len_oprLogin <> 0 then
                                            oprLogin_left = left(oprLogin, len_oprLogin - 3)
                                            oprLogin = oprLogin_left &":00"
                                            else
                                            oprLogin = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) & " " & oprlogin_hh(a) &":"& oprlogin_mm(a) &":00"
                                            end if
                                        
                            
                                        else
                                        'Response.Write "her 4"

                                        oprLogud = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) &" "& oprlogud_hh(a) &":"& oprlogud_mm(a) &":00"
                                        oprLogin = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) &" "& oprlogin_hh(a) &":"& oprlogin_mm(a) &":00"
                                        
                                        end if
                            
                          
                            '*** use_ig_sltid = ignorer periode ***
                            
                            '** er der angivet logud ***'

                            'Response.Write "her 3 tiderrettet: "&  tiderRettet  & " ("& logudrettetJa &")<br>"
                            'Response.Write oprLogin &" ("& oprloginHH &") - " & loginTid & " OR oprlogud: "& oprLogud &" ("& logudHH &" / "& oprlogudHH &") <>  nyt logud:" & logudTid & "<br>logudHH: " & logudHH
                            

                            if logudHH <> "99" AND oprlogudHH <> "00" AND (isNull(logudHH) <> true AND isNull(oprlogudHH) <> true) AND (logudHH <> "" AND oprlogudHH <> "") then  'AND logudHH <> "00" then '99 '13 = 00, dvs over 24 ==> Ikke angivet før
                                    
                                    if instr(oprLogin, "#") <> 0 then
                                    oprLogin = loginTid
                                    end if
                
                                    if instr(oprLogud, "#") <> 0 then
                                    oprLogud = logudTid
                                    end if

                                    'oprLogin = replace(oprLogin, "#", "0")
                                    'oprLogud = replace(oprLogud, "#", "0")
                                    
                                    'Response.write "oprLogin:"  & oprLogin & " <> loginTid: " & loginTid & "oprLogud: "& oprLogud & "logudTi:"& logudTid
                                    'response.flush

                                    if (cDate(oprLogin) <> cDate(loginTid)) OR (cDate(oprLogud) <> cDate(logudTid)) then
                                    tiderRettet = 1
                                    end if

                                      'Response.Write "her 21 tiderrettet: "&  tiderRettet  & "<br>"
                                       
                                    

                            else
                                 '*** Ellers tjekkes kun logind tid **'
                                 
                                 if oprloginHH <> "00" then 'første logind

                                 'Response.write "oprLogin:"  & oprLogin & " <> loginTid: " & loginTid
                                 'Response.end
                                 'if oprLogin

                                 if instr(oprLogin, "#") = 0 then 'hvis minutter kun er angivet med 1 tal istedet for 2

                                 if (cDate(oprLogin) <> cDate(loginTid)) then
                                 tiderRettet = 1
                                 end if

                                 end if

                                 end if
                            
                                  'Response.Write "her 20 tiderrettet: "&  tiderRettet  & "<br>"

                            end if
					
					        
					        'Response.end
					       
					
					
					
					
					        '*** Validere at login og logud tider er blevet ændret ved rediger. **'
					        '*** I så fald skal der angives en kommentar ***

                            if logudHH <> "99" then '** skal ikke afprøves hvis logudtid ikke er angivet

				            if cint(tiderRettet) = 1 OR (lto = "dencker" AND (stur(a) = 5 OR stur(a) = 6)) OR (lto = "outz" AND stur(a) = 2) then
				    
				   
					        call erStempelurOn()
			                if cint(stempelur_ignokomkravOn) = 1 then		
                            useKommTvunget = 0         
                            else
                            useKommTvunget = 1
                            end if

                            'Response.end
					       ' select case lcase(lto)
                            'case "fk_bpm", "fk", "xintranet - local", "kejd_pb", "kejd_pb2", "tec"
                            'useKommTvunget = 0   
                            'case else
                            'useKommTvunget = 1
                            'end select


					            if (len(trim(komm)) = 0 AND useKommTvunget = 1) AND manuelt_afsluttet < 3 then
    					        useleftdiv = "m"
					            errortype = 130
					            'if lto = "outz" then
					   
					    
					            'call showError(errortype)
					            'else
					            call showError(errortype)
					            'end if
					    
					            Response.end
    					
					            end if
					
					        end if


                           '*** Kommentar ***'
					        '************************************************************************
					
					
					            '*****************************************************************'
				                '*** Skal der afsendes mail (Ved ændring af logind / logud tid)??? 
				                '*** Hvis logud er manuelt indtastet fra logud stat siden. *** 
				                if ((logudrettetJa = 1 AND tiderRettet = 1) AND manuelt_afsluttet < 3) OR (lto = "dencker" AND stur(a) = 5 OR stur(a) = 6) OR (lto = "outz" AND stur(a) = 2) then
				        
    					
				        	
					                                '***** Oprettter Mail object ***''
			                                        if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\stempelur.asp" then
                                			
                                			
			                                        'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
			                                        '' Sætter Charsettet til ISO-8859-1 
			                                        'Mailer.CharSet = 2
			                                        '' Afsenderens navn 
			                                        'Mailer.FromName = "TimeOut Stempelur Service"
			                                        '' Afsenderens e-mail 
			                                        'Mailer.FromAddress = "timeout_no_reply@outzource.dk"
			                                        'Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk"
			                                        'Mailer.ContentType = "text/html"
                                		    
                            
                                                    Set myMail=CreateObject("CDO.Message")
                                                    myMail.From="timeout_no_reply@outzource.dk" 
            
                                                

                        		                    '*** Finder medarbejder ***'
                        		                    mnavn = "-"
                                            
                                                  
                        		            
                                                    strSQLmedarb = "SELECT mnavn, mnr, email, init FROM medarbejdere WHERE mid = "& useMid(a) 'medarbSel
                                		    
                                                   
                                		    
                        		                    oRec.open strSQLmedarb, oConn, 3
                        		                    if not oRec.EOF then
                                		    
                        		                    mnavn = oRec("mnavn") & " ["& oRec("init") &"]"
                                                    moNavn = mnavn
                                                    moEmail = oRec("email")
                                		    
                        		                    end if
                        		                    oRec.close


                                                    '*** Hvem har rettet ****'
                                                    call meStamdata(session("mid"))
                                                    afsenderNavn = meNavn
                                                    afsenderEmail = meEmail 
                                                    afsenderInit = meInit
                                		    
                                		            '*** Type navn ***'
                                		            stempelur_typenavn = ""
                                		            strSQL_stemptypnavn = "SELECT navn FROM stempelur WHERE id = "& stur(a)
                                		            oRec.open strSQL_stemptypnavn, oConn, 3
                                		    
                                		            if not oRec.EOF then
                                		    
                                		            stempelur_typenavn = oRec("navn")
                                		    
                                		            end if
                                		            oRec.close
                                		    
        		                            
                                                    'Mailens emne
                                                    if (lto = "dencker" AND (stur(a) = 5 OR stur(a) = 6)) OR (lto = "outz" AND stur(a) = 2) then
                                                    'Mailer.Subject = "Medarbejder "& mnavn &" har haft et "& stempelur_typenavn
                                                    myMail.Subject= "Medarbejder "& mnavn &" har haft et "& stempelur_typenavn 
                                                    else
			                                        'Mailer.Subject = "Medarbejder "& mnavn &" har ændret Komme/Gå tid. (type: "& stempelur_typenavn &") - "& lto 
                                                    myMail.Subject= "Medarbejder "& mnavn &" har ændret Komme/Gå tid. (type: "& stempelur_typenavn &") - "& lto
        			                                end if
        			                        
			                                        'Modtagerens navn og e-mail
			                                        select case lto
			                                        case "dencker" 

                                                     'Mailer.AddRecipient ""& moNavn &"", ""& moEmail &""
		                                                myMail.To= ""& moNavn &"<"& moEmail &">"	     
			                                    
			                                            if stur(a) = 5 OR stur(a) = 6 then
			                                            'Mailer.AddRecipient "Karen", "krt@dencker.net"
                                                        'myMail.To= "Karen<krt@dencker.net>"
			                                            'Mailer.AddBCC "Anders Dencker", "ad@dencker.net"
                                                        'myMail.Bcc= "Anders Dencker<ad@dencker.net>"
                                                        'Mailer.AddBCC "Per Kristensen", "pk@dencker.net"
                                                        myMail.Bcc= "Karen<krt@dencker.net>;Anders Dencker<ad@dencker.net>;Per Kristensen<pk@dencker.net>"
                                                        
                                                        'Mailer.AddBCC "OutZourCE", "support@outzource.dk"
			                                            else
			                                            'Mailer.AddRecipient "Anders Dencker", "ad@dencker.net"
                                                        'myMail.To= "Anders Dencker<ad@dencker.net>"
                                                        'Mailer.AddBCC "Per Kristensen", "pk@dencker.net"
                                                        myMail.Bcc= "Anders Dencker<ad@dencker.net>;Per Kristensen<pk@dencker.net>"
                                                        'Mailer.AddBCC "OutZourCE", "support@outzource.dk"
			                                            end if
			                                    
			                                                                              

			                                        case "outz"
			                                        'Mailer.AddRecipient "OutZourCE", "sk@outzource.dk"
                                                    

                                                        if cint(useMid(a)) <> cint(session("mid")) then
                                                        myMail.To= ""& moNavn &"<"& moEmail &">"
		                                                end if	  
                                        
                                                        myMail.Bcc= "OutZourCE<support@outzource.dk>"

			                                        case "cst"
			                                        'Mailer.AddRecipient "C.S.T", "th@copst.com"
                                                    myMail.To= "C.S.T<th@copst.com>"
                                                    case "jttek"
			                                        'Mailer.AddRecipient "JT-Teknik", "jt@jtteknik.dk"
                                                    myMail.To= "Annette Østergaard<an@jtteknik.dk>"
                                                    myMail.Bcc= "JT-Teknik<jt@jtteknik.dk>" 
                                                    case "tec", "esn"
                                                    
                                                        if cint(useMid(a)) <> cint(session("mid")) then
                                                        'Mailer.AddRecipient ""& moNavn &"", ""& moEmail &""
                                                        myMail.To = ""& moNavn &"<"& moEmail &">"
		                                                end if	                                        

                                                    case "akelius"
                
                                                        if len(trim(moNavn)) <> 0 then
                                                        'Mailer.AddRecipient ""& moNavn &"", ""& moEmail &""
                                                        myMail.To = ""& moNavn &"<"& moEmail &">"
		                                                end if

                                                    case else
			                                        'Mailer.AddRecipient "OutZourCE", "timeout_no_reply@outzource.dk"
                                                    'myMail.To= "OutZourCE<support@outzource.dk>"

                                                        if len(trim(moNavn)) <> 0 then
                                                        'Mailer.AddRecipient ""& moNavn &"", ""& moEmail &""
                                                        myMail.To = ""& moNavn &"<"& moEmail &">"
		                                                end if

			                                        end select
                                			
        			                       
                                			
			                                                ' Selve teksten
					                                        strBody = "Der er blevet ændret i Komme/Gå tid, eller der er brugt en indstilling der skal begrundes."_
                                                            &"<br><br><b>Rettelsen er udført af:</b> "& afsenderNavn &" ["& afsenderInit &"], "& afsenderEmail &""_
					                                        &"<br><br><b>Type:</b> "& stempelur_typenavn &""_
					                                        &"<br><br><b>Medarbejder:</b><br>"& mnavn &"<br><br><b>Logind dato og klokkeslet:</b> "& formatdatetime(loginTid, 1) &" kl. "& formatdatetime(loginTid, 3) &" til "& formatdatetime(logudTid, 3) &"<br><br>"_ 
					                                        &"<b>Kommentar:</b><br> "_
					                                        &""& komm &"<br><br>"_
					                                        &"Med venlig hilsen<br><br>TimeOut Email Nofikationsservice<br><br>&nbsp;" 
        					                        



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

                                                            if len(trim(moEmail)) <> 0 then

                                                            select case lto
                                                            case "tec", "esn"
                                                                
                                                                    if cint(useMid(a)) <> cint(session("mid")) then
                                                                    myMail.Send
                                                                    end if

                                                            case else
                                                                
                                                                if len(trim(moNavn)) <> 0 then
                                                                myMail.Send
                                                                end if
        
                                                            end select
        
                                                            end if
                                                            set myMail=nothing




					                                        'If Mailer.SendMail Then
                                    				
					                                        'Else
    					                                    '    Response.Write "Fejl...<br>" & Mailer.Response
  					                                        'End if	
                                				
                                			
			                                        end if ''** Mail
        					                    
                    
                                               
        					                
        					
					       
				                end if 'tiderrRettet
					 
                       end if '*** Logudtid ikke angivet #99#
                       '****************************************************************************************************


					
					
					thisDato = year(loginDato(a)) &"/"& month(loginDato(a)) &"/"& day(loginDato(a))
					
                    'Response.Write "oprLogud: " & oprLogud & "logudHH"& logudHH
                    'Response.flush
                    if logudHH <> "99" then
                    oprLogud_first = thisDato &" "& formatdatetime(oprLogud, 3)
                    end if
					
                    'Response.write "logudTid A:" & logudTid
					        
					        '*** Findes der allerede en indtastning på denne dato i dette tidsrum ****'
                            'Response.Write logudTid
                            'Response.flush
                            if logudHH <> "99" then
					            call stempelur_tidskonflikt(ids(a), useMid(a), loginTid, logudTid, thisDato, 3)
                            end if
					   
                       
                       'Response.write "<br>logudTid C:" & logudTid
                               
					
					    '*** er logud angivet`eller tomt? ***'
					    if logudHH <> "99" then
					    logudTid = logudTid 
					    else
					    logudTid = ""
					    end if
					
					    
                        

					
					'*******************************************************************'
					'**** Opretter nyt / rediger / angiver slut tid på eksisterende ****'
					'*******************************************************************'
					if ids(a) <> 0 then 
					
					    
					     if logudrettetJa = 1 then
					       
					      
					       
					       if manuelt_afsluttet = 3 then
					       manuelt_afsluttet = 3
					       else
					       manuelt_afsluttet = tiderRettet '** 0/1
					       end if
					     
					     else  
					       
					      
					        
					       manuelt_afsluttet = manuelt_afsluttet 
					        
					     end if
					   
					   
					   
					    '***** Alt er ok, Opdaterer logind *****'
		                			     

					    strSQL = "UPDATE login_historik SET dato = '"& thisDato &"', login = '"& loginTid &"'"
					    
					    '** Er logud angivet ***'
					    if logudHH <> "99" then
                        strSQL = strSQL & ", logud = '"& logudTid &"', "
					    else
					    strSQL = strSQL & ", logud = NULL, "
					    end if
					    
					    strSQL = strSQL &" stempelurindstilling = "& stur(a) &", minutter = "& minThisDIFF &", "
					    
					    if manuelt_afsluttet <> 0 then
					    strSQL = strSQL &" manuelt_afsluttet = "& manuelt_afsluttet &","
					    end if
					    
					    strSQL = strSQL & "kommentar = '"& komm &"'"
    					
					    '*** Er det første logud ***?
					    '*** Ellers må logud_fisrt og login_fisrt ikke rettes ***'

					    if cint(logudfirstWasNull) = 1 AND logudHH <> "99" then '(logudrettetJa <> 1 AND logudHH <> "99" AND oprlogudHH = "00") OR 
					    strSQL = strSQL & ", logud_first = '"& logudTid &"'"
					    end if
                        'oprLogud_first
					
					
					    strSQL = strSQL & " WHERE id = " & ids(a)
					
					    
					    
					    'Response.Write "<br>"& strSQL & "<br>"
					    'Response.end

                                                        

                                                    '*********************************************************************************************'
                                                    '*** Opdaterer pauser : ALTID manuel indtastet ved rediger
                                                    '*********************************************************************************************'

                                                    psloginTidp = loginTid
                                                    pslogudTidp = logudTid

                                                    LoginDatoPaus = thisDato

                                                    strUsrId = useMid(a)

                                                 

                                                    
                                                    '**** Tømmer pauser så der er altid kun er indlæst 2 pauser pr. dag pr. medarb. ****
                                                    if p1 <> -1 OR p2 <> -1 then
                                                    strSQLpDel = "DELETE FROM login_historik WHERE stempelurindstilling = -1 AND dato = '"& year(LoginDatoPaus) &"/"& month(LoginDatoPaus) &"/"& day(LoginDatoPaus) &"' AND mid = "& strUsrId
	                                                'Response.write strSQLpDel
                                                    oConn.execute(strSQLpDel)
                                                    end if

                                                     
                                                     '** Manuel tilføjet '****
                                                    if cint(p1) > 0 then
                                                    '** p1 **
                                                    call tilfojPauser(0,strUsrId,LoginDatoPaus,psloginTidp,pslogudTidp,p1,p1_komm)
                                                    end if

                                                   
                                                    'response.write "TILFØJER PAUSE 1 p1: "& p1 

                                                    
                                                    
                                                    if cint(p2) > 0 then
                                                    '** p2 **
                                                    call tilfojPauser(0,strUsrId,LoginDatoPaus,psloginTidp,pslogudTidp,p2,p2_komm)
                                                    end if

                                                   
                                                    'response.write "<br>TILFØJER PAUSE 2 p2: "& p2


                                                    'Response.end


                                                    '**** End auto pauser *************'
                                                



					
					else
					
					        
					       
					        
	                         
					'****** Opretter nyt logind *****'
					
					'**** tjekker om periode er lukket ***'
					
					
					            '** er periode godkendt ***'
		                        tjkDag = thisDato
		                        erugeafsluttet = instr(afslUgerMedab(useMid(a)), "#"&datepart("ww", tjkDag,2,2)&"_"& datepart("yyyy", tjkDag) &"#")
                                
                		        
		                        'Response.Write "smilaktiv: "& smilaktiv & "<br>"
		                        'Response.Write "ugeNrAfsluttet: "& ugeNrAfsluttet & "<br>"
		                        'Response.Write "autolukvdatodato: "& autolukvdatodato & "<br>"
		                        'Response.Write "tjkDag: "& tjkDag & "<br>"
		                        'Response.Write "autolukvdato: "& autolukvdato & "<br>"
		                        'Response.Write "erugeafsluttet:" & erugeafsluttet & "<br>"
                		        
		                        call lonKorsel_lukketPer(tjkDag, -2)
                		         
                                if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", tjkDag) = year(now) AND DatePart("m", tjkDag) < month(now)) OR _
                                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", tjkDag) < year(now) AND DatePart("m", tjkDag) = 12)) OR _
                                (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", tjkDag) < year(now) AND DatePart("m", tjkDag) <> 12) OR _
                                (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", tjkDag) > 1))) OR cint(lonKorsel_lukketIO) = 1 then
                              
                                ugeerAfsl_og_autogk_smil = 1
                                else
                                ugeerAfsl_og_autogk_smil = 0
                                end if 
                					
				                if ugeerAfsl_og_autogk_smil = 1 AND level <> 100 then	
				                
				                errortype = 108
				                regDato = day(tjkDag) &"-"& month(tjkDag) &"-"& year(tjkDag)
					            call showError(errortype)
					            Response.end   
					            
					            
					            end if
					            '***
					
					
					strSQL = "INSERT INTO login_historik  "_
					&"(dato, login, mid, stempelurindstilling, minutter, "_
					&" login_first "
					
					if logudHH <> "99" then 
					strSQL = strSQL & ", logud, logud_first"
					end if
					
					strSQL = strSQL & ", manuelt_oprettet, manuelt_afsluttet, kommentar, ipn) "_
					&" VALUES ('"& thisDato &"', '"& loginTid &"', "_
					&" "& useMid(a) &", "& stur(a) &", "& minThisDIFF &", '"& loginTid &"'"
					
					if logudHH <> "99" then 
					strSQL = strSQL & ", '"& logudTid &"', '"& logudTid &"'"
					end if
					
					
					strSQL = strSQL & ", 1, "& manuelt_afsluttet &", '"& komm &"', '"& ipn &"')"
				    
				    
				    'Response.write strSQL & "<br>"
					'Response.flush
                                
                                
                                
                                                    

                                                    '*********************************************************************************************'
                                                    '*** Tilføjer pauser 
                                                    '*** opret nyt logind kan kun forekomme fra:
                                                    '*** 1: automatisk logind fra terminal (cls_stempelur) IKKE HER --> Altid standard pause 1 og 2
                                                    '*** 2: logind side (cls_stempelur) IKKE HER --> Altid standard pause 1 og 2
                                                    '*** 3: fra personlig redigerbar logind historik fra timereg. siden, hvis der indtastest på en dato hvor der ikke allerede findes et logind. HER --> Kan veksle mellem standrard pauser og manuelt rettet
                                                    '*** 4: Rediger / opret nyt logind manuelt via link / popup HER --> Kan veksle mellem standrard pauser og manuelt rettet

                                                    '*********************************************************************************************'
                                                                                                     
                                                    psloginTidp = loginTid
                                                    pslogudTidp = psloginTidp

                                                    LoginDatoPaus = thisDato

                                                    strUsrId = useMid(a)

                                                    'p1_grp, p2_grp findes her: stPauserFralicens()
                                                    call stPauserFralicens(LoginDatoPaus)

                                                    '** tjekker om der skal tilføjes pause til projektgruppe ***' 
                                                    call stPauserProgrp(strUsrId, p1_grp, p2_grp)

                                                 

                                                    
                                                    '**** Tømmer pauser så der er altid kun er indlæst 2 pauser pr. dag pr. medarb. ****
                                                    if cint(p1_prg_on) = 1 OR cint(p2_prg_on) = 1 then
                                                    strSQLpDel = "DELETE FROM login_historik WHERE stempelurindstilling = -1 AND dato = '"& year(LoginDatoPaus) &"/"& month(LoginDatoPaus) &"/"& day(LoginDatoPaus) &"' AND mid = "& strUsrId
	                                                'Response.write strSQLpDel
                                                    oConn.execute(strSQLpDel)
                                                    end if


                                                    

                                                    
                                                    '** Manuel tilføjet / eller standard pause
                                                    if p1 > -1 then
                                                    psKomm_1 = p1_komm
                                                    psMin_1 = p1
                                                    else
                                                    psKomm_1 = ""
                                                    psMin_1 = stPauseLic_1
                                                    end if
                                                    
                                                   
                                                    if cint(p1_prg_on) = 1 AND cint(p1on) = 1 OR p1 > 0 then
                                                    '** p1 **
                                                    call tilfojPauser(0,strUsrId,LoginDatoPaus,psloginTidp,pslogudTidp,psMin_1,psKomm_1)

                                                    end if


                                                     '** Manuel tilføjet / eller standard pause
                                                    if p2 > -1 then
                                                    psKomm_2 = p2_komm
                                                    psMin_2 = p2
                                                    else
                                                    psKomm_2 = ""
                                                    psMin_2 = stPauseLic_2
                                                    end if

                                                    if cint(p2_prg_on) = 1 AND cint(p2on) = 1 OR p2 > 0 then
                                                    '** p2 **
                                                    call tilfojPauser(0,strUsrId,LoginDatoPaus,psloginTidp,pslogudTidp,psMin_2,psKomm_2)
                                                    end if

                                                    


                                                    'Response.end


                                                    '**** End auto pauser *************'


					
					end if 'ids(a) <> 0
					
					'Response.write strSQL & "<br>"
					'Response.end
				    oConn.execute(strSQL)
				    idLast = ids(s)
		    
		    
					
					       
	                        
	                        
							    end if '** validering
						    end if
						'end if
					'end if
				'end if 
			end if
		end if
	    
	   next
	                        
	    
        
                   
	   'Response.end                     
	                        
	                        
                            
                           
	                        

	                        
	                        select case rdir
                            case "lgnhist"
                            
                            'Response.write lnk
                            'Response.end
                            Response.redirect "logindhist_2011.asp?"&lnk
                            'Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
                            'Response.Write("<script language=""JavaScript"">window.close();</script>")
	                        
                            
	                        case "treg", "popup" 'fra timereg.siden
                            Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
	                        Response.Write("<script language=""JavaScript"">window.close();</script>")
                            
                            case "sesaba"
                            Response.redirect "../sesaba.asp?logudDone=1"
                            Response.end
                            case else
	                        'Response.Write("<script language=""JavaScript"">window.opener.location.href('stempelur.asp?menu=stat&func=stat&medarbSel="&medarbSel&"&showonlyone="&showonlyone&"&hidemenu="&hidemenu&"&lastid="&idLast&"');</script>")
                            Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
                            Response.Write("<script language=""JavaScript"">window.close();</script>")
                            end select						
	
	                        
	                        
	    
	
	case "redloginhist", "oprloginhist"
	
	
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
     <script src="inc/stempelur_jav.js"></script>
	<%
	
	
	    if rdir = "sesaba" then
	    sideDivTop = 20
	    sideDivLeft = 160
	    opdAfsl = "Stempelur, angiv logind og logud tid"
	    bgcol = "#ffffff"
	    else
	    
	    sideDivTop = 0
	    sideDivLeft = 10
	    opdAfsl = "Logind historik (Stempelur) - Opdater"
	    bgcol = "#ffffff"
	    end if
	
	%>
	
	
	<div id="sindhold2" style="position:relative; width:600px; left:<%=sideDivLeft%>px; top:<%=sideDivTop%>px; visibility:visible;">
   
	
	<% 
	
	oimg = "ikon_stempelur_48.png"
	oleft = 0
	otop = 0
	owdt = 600
	oskrift = opdAfsl
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	
	dim loginDT, logudDT, datoDT, stur, kommentar, loginDato, idThis, firsttime
	useDato = year(now) & "/" & month(now) &"/" & day(now)
	
	if func = "redloginhist" then
	
	            redim loginDT(50), logudDT(50), datoDT(50), stur(50)
	            redim kommentar(50), loginDato(50), idThis(50), firsttime(50)
	

	
	
	            if id <> 0 then
	
	
	    
	                strSQL2 = "SELECT id, dato, mid FROM login_historik WHERE id = "& id 
	                oRec2.open strSQL2, oConn, 3
                    if not oRec2.EOF then
        
                    useDato = oRec2("dato")
                    useMid = oRec2("mid")
        
                    end if
                    oRec2.close
	
	
	            else
	
	                if len(trim(medarbSel)) <> 0 then
                    medarbSel = medarbSel
                    else
                        if len(trim(session("mid"))) <> 0 then
                        medarbSel = session("mid")
                        else
                        %>
                        <!--#include file="../inc/regular/header_inc.asp"-->
                        <%
	                    errortype = 5
	                    call showError(errortype)
                        Response.end
                        end if
                    end if
	
	                strSQL2 = "SELECT id, dato, mid FROM login_historik WHERE mid = "& medarbSel &" ORDER BY login_first DESC limit 0, 1" 'AND dato = '"& useDato &"'
	                oRec2.open strSQL2, oConn, 3
                    if not oRec2.EOF then
        
                    useDato = oRec2("dato")
                    useMid = oRec2("mid")
        
                    end if
                    oRec2.close
	
        
                end if
                            

                            '**** Seneste 3 dage ***'
                            useDato = year(useDato) &"-"& month(useDato) &"-"& day(useDato)
                            useISNULLDatoLimit = dateAdd("d", - 3, useDato)
                            useISNULLDatoLimit = year(useISNULLDatoLimit) &"-"& month(useISNULLDatoLimit) &"-" & day(useISNULLDatoLimit)
	
	
	                        strSQL = "SELECT l.login, l.logud, l.dato, l.stempelurindstilling, l.kommentar, l.id FROM login_historik l "
	                        strSQL = strSQL & " WHERE l.mid = "& useMid &" AND l.stempelurindstilling <> -1 AND (l.dato = '"& useDato &"' OR (logud IS NULL AND l.dato > '"& useISNULLDatoLimit &"' AND l.dato <= '"& useDato &"')) ORDER BY login_first DESC"
	
                            'l.dato = '"& useDato &"' AND
                            'if lto = "dencker" then
	                        'Response.Write strSQL &" ID:" & id & "medarbSel:" & medarbSel
	                        'Response.end
                            'end if
	
	                        a = 0
	                        oRec.open strSQL, oConn, 3 
	                        while not oRec.EOF 
	
	                        loginDT(a) = oRec("login")
	
	                        if len(trim(oRec("logud"))) <> 0 then
	                        'Response.Write "<br>"& a & " her:" & oRec("logud") & "<br>" 
	                        logudDT(a) = oRec("logud")
	                        firsttime(a) = 0
	                        else
                                '** skal logud tid fødes eller skal fetet være tomt.
                                '** fødes ved logud
                                '** tomt ved rediger logind historik for dagen
                                if id <> 0 then
	                            logudDT(a) = "" 'now '"00:00:00"
                                else
                                logudDT(a) = now 
                                end if
	                        'Response.Write "<br>"& a & " der:" & oRec("logud") & "<br>"
	                        firsttime(a) = 1 
	                        end if
	
	                        datoDT(a) = oRec("dato")
	                        stur(a) = oRec("stempelurindstilling")
	                        kommentar(a) = oRec("kommentar")
	                        idThis(a) = oRec("id")
	
	                        a = a + 1
	                        oRec.movenext
	                        wend
	                        oRec.close 
	

                            'Response.end
	
	
	else
	
	a = 1
	

    '*** Forvalgt stempelurr ****'
    call fv_stempelur()

	redim loginDT(a), logudDT(a), datoDT(a), stur(a), kommentar(a), loginDato(a), idThis(a)
	redim firsttime(a)
	
	i = 0
	
	datoDT(i) = day(now) & "-" & month(now) & "-"& year(now)
	useMid = medarbSel
	idThis(i) = 0
	loginDT(i) = "" 'now
	logudDT(i) = "" 'now
	stur(i) = forvalgt_stempelur '1
	kommentar(i) = ""
	loginDato(i) = useDato
	firsttime(i) = 2 
	
	
	
	end if
	
	
	
	tTop = 10
    tLeft = 0
    tWdth = 560


    call tableDiv(tTop,tLeft,tWdth)
	
    if len(trim(useMid)) <> 0 then
    useMid = useMid
    else
    useMid = medarbSel
    end if
	
	%>
	<form action="stempelur.asp?menu=stat&func=dbloginhist&FM_usedatokri=1&medarbSel=<%=useMid%>&showonlyone=<%=showonlyone%>&hidemenu=<%=hidemenu%>&rdir=<%=rdir %>" method="post">
	<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr>  
	    <td style="padding:20px 0px 0px 25px;">
	
	<h4><%= weekdayname(weekday(useDato, 1))%> d. <%=formatdatetime(useDato, 1)%> 
    <%if id <> 0 then %>
    <!--<span style="font-size:9px;"><br />Samt uafsluttede loginds foregående 3 dage<br />-->
    <%
    call meStamdata(useMid)
    Response.write meTxt
     %></span>
    <%end if %></h4>
        &nbsp;
 
   
	
	<%
	
	for a = 0 TO a - 1
	
	if len(trim(loginDT(a))) <> 0 then
	loginDT(a) = loginDT(a)
    loginDThh = left(formatdatetime(loginDT(a), 3), 2)
    loginDTmm = mid(formatdatetime(loginDT(a), 3), 4, 2)

    oprloginDThh = loginDThh
    oprloginDTmm = loginDTmm

    if len(trim(logudDT(a))) <> 0 then
    logudDThh = left(formatdatetime(logudDT(a), 3), 2)
    logudDTmm = mid(formatdatetime(logudDT(a), 3), 4, 2)
     oprlogudDThh = logudDThh
    oprlogudDTmm = logudDTmm
    else
    logudDThh = ""
    logudDTmm = ""
    oprlogudDThh = "00"
    oprlogudDTmm = "00"
    end if

    

   

	'loginDato(a) = formatdatetime(loginDT(a), 2)
	else
	loginDT(a) = "" '01-01-2001 12:33:00
    loginDThh = ""
    loginDTmm = ""

    logudDThh = ""
    logudDTmm = ""

    oprloginDThh = "00"
    oprloginDTmm = "00"

    oprlogudDThh = "00"
    oprlogudDTmm = "00"
	'loginDato(a) = "01-01-2001"
	end if
	
	'Response.write DatoDT(a)
	'Response.flush
	
  
	
	
	if cdbl(idThis(a)) = cdbl(id) then
	bgthis = "#FFFFFF"
	bdcol = "#FFFFFF" 
	else
	bgthis = "#FFFFFF"
	bdcol = "#FFFFFF" 
	end if
    
    bdpx = "2" %>
	
	    <%select case firsttime(a) 
	    case 1 
        'stempelur_ind
	    strTxt = "<img src=../ill/blank.gif /><span style=""background-color:#FFFFe1;""><b>Denne komme/gå tid er aktivt!</B></span><br />"_
	    &"Du vil få den logud tid der er angivet nedenfor."
	    bdcol = "yellowgreen"
	    bgthis = bgthis
        bdpx = "5"
	    case 0
        'stempelur_ok
        strTxt = "<img src=../ill/blank.gif /><span style=""background-color:#FFFFe1;"">Denne komme/gå tid er allerede gemt i databasen.</span>"
	    bgthis = bgthis
	    bdcol = bdcol
	    case 2
	    strTxt = "<img src=../ill/blank.gif /><span style=""background-color:#FFFFe1;"">Tilføj komme/gå</span>"
	    bgthis = bgthis
	    bdcol = bdcol
	   end select 
       
       'if id = 0 then
       'strTxt = "Angiv logind og logud tider nedenfor."
       'else
       'strTxt = "&nbsp;"
       'end if%>
	
	<br />
	 <table cellspacing="0" cellpadding="0" border="0" width="550">
	 <tr bgcolor="<%=bgthis %>">
	    <td colspan=4 style="padding:10px 0px 10px 10px;">
	    <%= strTxt %>
	   </td></tr>
	
	<tr bgcolor="<%=bgthis %>">
		<td valign=top style="padding:0px 0px 0px 10px;"><b>Dato:</b>
		<input type="hidden" name="id" id="id" value="<%=idThis(a)%>">
		<input type="hidden" name="mid" id="mid" value="<%=useMid%>">
		<br />
		<%
		't = 1000
		'if t = 1000 then
		if level = 1 OR id = 0 then%>
		<input type="text" size="8" name="logindato" id="logindato" value="<%=datoDT(a)%>"><span style="color:#999999;"><br />Eks: 03-04-2006</span>
		<%else%>
		
		<%=formatdatetime(datoDT(a),1) %> <!--left(formatdatetime(loginDT, 3), 5)--> 
		<input type="hidden" name="logindato" id="logindato" value="<%=datoDT(a)%>">
		<%end if%></td>
	
		<td valign=top style="padding:0px 0px 0px 10px;"><b>Logind:</b><br />
		
		<%'if level <= 2 OR level = 6 then%>
		<input type="text" name="FM_login_hh" id="FM_login_hh_<%=a%>" size=1 value="<%=loginDThh%>"> 
		<b>:</b> <input type="text" name="FM_login_mm" id="FM_login_mm_<%=a%>" size=1 value="<%=loginDTmm%>">
		&nbsp;&nbsp;tt:mm

        <!-- bruges til arrray split -->
	    <input type="hidden" name="FM_login_hh" id="Hidden3" value="#">
		<input type="hidden" name="FM_login_mm" id="Hidden4"  value="#">

		<%'else%>
		
        <!-- 
		%=loginDThh &":"&loginDTmm%>
		<input type="hidden" name="FM_login_hh" id="FM_login_hh" value="<=loginDThh%>">
		<input type="hidden" name="FM_login_mm" id="FM_login_mm"  value="<=loginDTmm%>">
        -->
	
		<%'end if%></td>
		
		<input type="hidden" name="oprFM_login_hh" id="hidden5" value="<%=oprloginDThh%>">
		<input type="hidden" name="oprFM_login_mm" id="Hidden2"  value="<%=oprloginDTmm%>">

         <input type="hidden" name="oprFM_login_hh" id="Hidden1" value="#">
		<input type="hidden" name="oprFM_login_mm" id="Hidden8"  value="#">
	    
	    
	    
	
	
		<td valign=top style="padding:0px 0px 0px 10px;"><b>Logud:</b> <!--(<a href="#" id="a_<=a%>" class="rmenu">ikke endnu</a>)--><br />
		
		
		<input type="text" name="FM_logud_hh" id="FM_logud_hh_<%=a%>" size=1 value="<%=logudDThh%>"> <b>:</b>
		<input type="text" name="FM_logud_mm" id="FM_logud_mm_<%=a%>" size=1 value="<%=logudDTmm%>">&nbsp;&nbsp;tt:mm

        <!-- bruges til arrray split -->
		 <input type="hidden" name="FM_logud_hh" id="Hidden6" value="#">
		<input type="hidden" name="FM_logud_mm" id="Hidden9"  value="#">
		<% 
		'** Hvis der ikke før er angivet en logud tid ***'
		'if len(trim(logudDT(a))) = 0 then
		'logudDT(a) = formatdatetime(now, 3)
		'else
		'logudDT(a) = logudDT(a) 
		'end if
		%>
		
		<input type="hidden" name="oprFM_logud_hh" id="oprFM_logud_hh" value="<%=oprlogudDThh%>">
		<input type="hidden" name="oprFM_logud_mm" id="oprFM_logud_mm" value="<%=oprlogudDTmm%>">
        <!-- bruges til arrray split -->
		 <input type="hidden" name="oprFM_logud_hh" id="Hidden10" value="#">
		<input type="hidden" name="oprFM_logud_mm" id="Hidden11"  value="#">
		
		
		
		
		</td>
		
		
		
	
		<td valign=top style="padding:0px 0px 0px 10px;"><b>Stempelurindst:</b><br />
		<select name="FM_stur" style="width:100px;">
         <%if lto <> "fk_bpm" AND lto <> "kejd_pb" AND lto <> "kejd_pb2" then %>
		<!--<option value="0">Ingen</option>-->
		<%end if

		strSQL3 = "SELECT id, navn FROM stempelur ORDER BY navn"
		oRec3.open strSQL3, oConn, 3 
		while not oRec3.EOF 
		
		if stur(a) = oRec3("id") then
		sel = "SELECTED"
		else
		sel = ""
		end if
		%>
		<option value="<%=oRec3("id")%>" <%=sel%>><%=oRec3("navn")%></option>
		
		<%
		oRec3.movenext
		wend
		oRec3.close 
		%></select>
		</td>
	</tr>
	<tr bgcolor="<%=bgthis %>">
	    <td colspan=4 style="padding-bottom:30px; padding-left:10px;"><br /><b>Kommentar:</b>
	    <br /><textarea id="FM_kommentar_<%=a %>" name="FM_kommentar" style="width:440px; height:60px;"><%=kommentar(a) %></textarea></td>
         <input type="hidden" name="FM_kommentar" id="Hidden13"  value="#">  
	</tr>
	</table>
	<%
    '** Standard pauser fra Licens ****
    psDt = datoDT(a) 'now
    next %>
	
	<br />
	
	
	    
	   
        
        <%
        
        '** Standard pauser fra Licens ****
        call stPauserFralicens(psDt) 


        '*** Adgang for specielle projektgrupper ****'
        call stPauserProgrp(useMid, p1_grp, p2_grp)

       
        

        '***** Opret nye pauser ****'

         
            'select case lto 
            'case "xx"
            '    hdn = 0
            'case else
            '    hdn = 1
            'end select


        %>
        <table cellspacing=0 cellpadding=0 border=0 width=500>
	    <tr bgcolor="#FFFFFF">
	    <td style="padding-left:10px;"><br />

        <%if func = "redloginhist" then %>

        <b>Tilføj/Rediger pauser på denne dato:</b><br />
        <br> Der kan kun tilføjes pauser på en dato med et gyldigt logind.
	    <br />Pauser bliver automatisk fratrukket login timer den pågældende dato.
        <br />Standard pause(r) er: <b><%=stPauseLic_1 %> min.</b>
            
            <%if stPauseLic_2 <> 0 then %>
             og <b><%=stPauseLic_2 %></b> min.
            <%end if %>
        <%
     
            
            
            '*** Henter allerede oprettede pauser på denne dato
            strSQL = "SELECT minutter, kommentar FROM  "_
            &" login_historik WHERE stempelurindstilling = -1 AND "_
            &" dato = '"& useDato &"' AND mid = "& useMid &" ORDER BY minutter DESC" 
            
            'Response.Write strSQL & " ID " & id
            'Response.Flush
            
            
            p = 1
            oRec.open strSQL, oConn, 3
            while not oRec.EOF
            %>
        
            <br /><br /><b>Pause <%=p %>:</b> 
            <select name="p<%=p %>" id="p<%=p %>" style="font-family:arial; font-size:9px;">
            
             <% 
           
             
             caseVal = oRec("minutter")
             
            
             selM0 = "SELECTED"
             selM5 = ""
             selM10 = ""
             selM15 = ""
             selM20 = ""
             selM25 = ""
             selM30 = ""
             selM35 = ""
             selM40 = ""
             selM45 = ""
             selM50 = ""
             selM55 = ""

             Select case caseVal
             case 0
             selM0 = "SELECTED"
             case 5
             selM5 = "SELECTED"
             case 10
             selM10 = "SELECTED"
             case 15
             selM15 = "SELECTED"
             case 20
             selM20 = "SELECTED"
             case 25
             selM25 = "SELECTED"
             case 30
             selM30 = "SELECTED"
             case 35
             selM35 = "SELECTED"
             case 40
             selM40 = "SELECTED"
             case 45
             selM45 = "SELECTED"
             case 50
             selM50 = "SELECTED"
             case 55
             selM55 = "SELECTED"
             end select
             
               
             
             %>
             <option value="0" <%=selM0%>>0 min</option>
             <option value="5" <%=selM5%>>5 min</option>
             <option value="10" <%=selM10%>>10 min</option>
             <option value="15" <%=selM15%>>15 min</option>
              <option value="20" <%=selM20%>>20 min</option>
              <option value="25" <%=selM25%>>25 min</option>
              <option value="30" <%=selM30%>>30 min</option>

               <option value="35" <%=selM35%>>35 min</option>
             <option value="40" <%=selM40%>>40 min</option>
              <option value="45" <%=selM45%>>45 min</option>
              <option value="50" <%=selM50%>>50 min</option>

               <option value="55" <%=selM55%>>55 min</option>
              
            </select> min.<br />
        
            
            <br />Kommentar pause <%=p %>:
	        <br /><textarea id="FM_komm_p<%=p %>" name="FM_komm_p<%=p %>" style="font-size:9px; font-family:arial; width:350px; height:46px;"><%=oRec("kommentar") %></textarea>

            <%
            if p = 1 then
            pOnVal = p1on
            else
            pOnVal = p2on
            end if
            %>

	        <input id="Hidden7" name="p<%=p%>on" value="<%=pOnVal %>" type="hidden" />
            <% 
            p = p + 1
            oRec.movenext
            wend
            oRec.close 

              

            end if 'func
            
            
            '*******************************************************
            '*** Der er ikke oprettet pauser på denne dato endnu ***
            '*** Eller der er tilføjet 1 pause *********************
            p1_fo = 0
            p2_fo = 0

            if (func = "redloginhist" AND id = 0) OR func = "oprloginhist" then
            startval = 1 'Standard Pause x antal min. er forvalgt
            hdn = 1
            else
            startval = 0 'Pause 0 min er forvalgt
            hdn = 0
            end if
        
          

            if cint(p) = 1 then

                if p1on <> 0 AND p1_prg_on = 1 then
                call nyePauser(1, p1on, startval, hdn)
                %>
                <input id="p1on" name="p1on" value="<%=p1on %>" type="hidden" />
                <%
                p1_fo = 1
                end if

                if p2on <> 0 AND p2_prg_on = 1 then
                call nyePauser(2, p2on, startval, hdn)
                 %>
                 <input id="p2on" name="p2on" value="<%=p2on %>" type="hidden" />
                <%
                p2_fo = 1
                end if

            end if
            
            
            if cint(p) = 2 then
                if p2on <> 0 AND p2_prg_on = 1 then
                call nyePauser(2, p2on, startval, hdn)
                %>
                 <input id="p2on" name="p2on" value="<%=p2on %>" type="hidden" />
                <%
                p2_fo = 1
                end if
            end if

            'end if
        

        if cint(p) = 1 AND p1_fo = 0 AND p2_fo = 0 then
        %>
        <br /><br /><br />
        <span style="color:#999999;">
        Du har ikke rettigheder til at tilføje pauser, <br />eller der skal ikke angives pauser på denne dag i ugen.
        </span>

        <%
        end if
        %>
            
            
	    </td></tr>
	    
	    
	<tr>
		<td align=right style="padding-right:30px;"><br><br />&nbsp;
		<%if id <> 0 then %>
            <input id="Submit1" type="submit" value="Opdater >>" />
		<%else %>
            <input id="Submit1" type="submit" value="Gem >>" />
        <%end if %>		
		</td>
	</tr>
	</table>
    </form>
    
    
    </td></tr></table>
	</div><!-- table div -->
	</div><!-- sidediv -->
	
	
     <br /><br /> <br /><br /> <br /><br /> <br /><br />
    &nbsp;
    
	
	<%case "stat"
	
	
	
	
	if hidemenu = 0 then%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <script src="inc/stempelur_jav.js"></script>
	
    <!--
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	
	<%call tsamainmenu(7)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	if showonejob <> 1 then
		call stattopmenu()
	end if
	%>
	</div>
    -->
	<%

    call menu_2014()

	sideDivTop = 102 '132
	sideDivLeft = 90
	
	else
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	
	sideDivTop = 20
	sideDivLeft = 20
	
	end if

    if len(trim(request("FM_type"))) <> 0 then
    typ = request("FM_type")
    response.Cookies("stat")("logindtyp") = typ
    else
        if request.Cookies("stat")("logindtyp") <> "" then
        typ = request.Cookies("stat")("logindtyp")
        else 
        typ = 0
        end if
    end if
	
	if len(request("FM_sumprmed")) <> 0 then
	showTot = 1
	sumChk = "CHECKED"
	else
	sumChk = ""
	showTot = 0
	end if
	%>


    
       <div id="loadbar" style="position:absolute; display:; visibility:visible; top:360px; left:200px; width:300px; background-color:#ffffff; border:10px #9ACD32 solid; padding:10px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	<br />
	Forventet loadtid:
	
	ca.: <b>3-45 sek.</b><br />
    <br />&nbsp;
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
    
  
    <br />
	<br />&nbsp,
    </td></tr>
    
 
        </table>

	</div>

    

	
	<div id="sindhold" style="position:absolute; left:<%=sideDivLeft%>px; top:<%=sideDivTop%>px; visibility:visible;">
	
	<% 
	oimg = "ikon_stempelur_48.png"
	oleft = 0
	otop = 0
	owdt = 600
	oskrift = "Komme/Gå tider (logind-historik)"
	
    if media = "print" then
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
    end if
	%>
	
	<%call filterheader_2013(0,0,800,oskrift)%>
	<table cellspacing="0" cellpadding="0" border="0" width=100%>
	<form action="stempelur.asp?menu=stat&func=stat&FM_usedatokri=1&lastid=<%=lastid%>&showonlyone=<%=showonlyone%>&hidemenu=<%=hidemenu%>" method="post">
	<tr>
			
			
			<%call progrpmedarb %>
			
            <%strSQLlogtyp = "SELECT id, navn FROM stempelur WHERE id <> 0 ORDER BY navn"%>

			</tr>
			<tr><td valign=top><b>Type:</b><br />
			<select name="FM_type">
            <option value=0>Vis alle</option>
            <%
            oRec3.open strSQLlogtyp, oConn, 3
            while not oRec3.EOF  
            
            if cint(typ) = oRec3("id") then
            typSEL = "SELECTED"
            else
            typSEL = ""
            end if
            %>
            <option value="<%=oRec3("id") %>" <%=typSEL %>><%=oRec3("navn") %></option>
            <%
            oRec3.movenext
            wend
            oRec3.close %>
            </select><br /><br />
			<input type="checkbox" name="FM_sumprmed" id="FM_sumprmed" value="1" <%=sumChk%>> Vis sum fordelt pr. medarbejder og logind type.&nbsp;&nbsp;</td>
			
			
			
			
			
			<td valign=top><b>Periode:</b><br>
			<!--#include file="inc/weekselector_s.asp"-->
			</td>
	</tr>
	<tr>
			<td colspan=2 align=right>
                <input id="Submit2" type="submit" value="Kør >> " /></td>
	</tr></form>
	</table>
	
	
	<!-- filter header -->
	</td></tr></table>
	</div>
	<br />&nbsp;
	<%
	
	Response.flush
	
	
	'*** Alle timer, uanset fomr. ***
	sqlDatoStart = strDag&"-"&strMrd&"-"&strAar
	sqlDatoSlut = strDag_slut&"-"&strMrd_slut&"-"&strAar_slut
	d_end = datediff("d", strDag&"/"&strMrd&"/"&strAar, strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 2, 2)
    layout = 0
	
	if datediff("d", strDag&"/"&strMrd&"/"&strAar, strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 2, 2) > 365 then
	%>
	<br /><br /><br />
	<div style="position:relative; background-color:#ffffff; border:10px #CCCCCC solid; padding:20px;">
	Datointerval er mere 365 dage. Vælg et midre interval.
	</div>
	<%
	
	Response.end
	
	end if
	
	call ersmileyaktiv()
	
	call stempelurlist(thisMiduse, showtot, layout, sqlDatoStart, sqlDatoSlut, typ, d_end, lnk)
	
	%>
	
	
	</div>
	<br /><br /><br /><br /><br /><br /><br />&nbsp;
	<%case else%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	</script>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20px; top:10px; visibility:visible;">
	<%call crmaktstatheader%>
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr>
    <td valign="top"><font class="pageheader">Kontrolpanel | Stempelur indstillinger</font>
	<br>
	<br>
	Sortér efter <a href="stempelur.asp?menu=tok&sort=navn">Navn</a> eller <a href="stempelur.asp?menu=tok&sort=nr">Id nr.</a>
	<img src="../ill/blank.gif" width="200" height="1" alt="" border="0">
	<a href="stempelur.asp?menu=tok&func=opret">Opret ny Stempelur indstilling <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a><br>
	<br>&nbsp;
	</td>
	</tr>
	</table>
	
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2><img src="../ill/tabel_top_left.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan=5 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top rowspan=2><img src="../ill/tabel_top_right.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td width="50" class=alt><b>Id</b></td>
		<td width="300"class=alt><b>Stempelur indstillinger</b></td>
		<td><b>Faktor</b></td>
		<td><b>Minimums indtast.</b></td>
		
		<td class=alt>&nbsp;</td>
	</tr>
	<%
	sort = Request("sort")
	if sort = "navn" then
	strSQL = "SELECT id, navn, faktor, forvalgt, minimum FROM stempelur ORDER BY navn"
	else
	strSQL = "SELECT id, navn, faktor, forvalgt, minimum FROM stempelur ORDER BY id"
	end if
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	%>
	<tr>
		<td bgcolor="#003399" colspan="7"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="#ffffff" onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'#ffffff');">
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td><%=oRec("id")%></td>
		<td><a href="stempelur.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a> 
		<%if oRec("forvalgt") = 1 then%>
		&nbsp;<b>(Forvalgt)</b>
		<%end if%>
		</td>
		<td><%=formatnumber(oRec("faktor"), 2)%></td>
		<td><%=formatnumber(oRec("minimum"), 0)%> min.</td>
		<td><a href="stempelur.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="" border="0"></a></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>	
	<tr bgcolor="#5582D2">
		<td width="8" valign=top height=20 style="border-bottom:1px #003399 solid; border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=5 valign="top" style="border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top style="border-bottom:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	</table>
	
	<br><br><br>
	<a href="Javascript:window.close()" class=red>[Luk vindue]</a><br><br>&nbsp;
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
