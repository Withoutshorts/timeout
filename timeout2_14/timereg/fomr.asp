<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%
sub crmaktstatheader
%>
<table border=0 cellpadding=0 cellspacing=0 width="450">
	<tr>
	<td valign="top" width="163"><img src="../ill/logo_bg.gif" width="163" height="53" alt="" border="0"></td>
	<td valign="bottom"><b>Timeout Kontrolpanel - Forretningsområder</b><br>
	Tilføj, fjern eller rediger Forretningsområder.</td>
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
	
	
	'** Faste filter kri ***'
	thisfile = "fomr"
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	
	call fomr_account_fn()
	
	
	
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190px; top:102px; visibility:visible;">
	<h4>Forretningsområder</h4>
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> et forretningsområde. Er dette korrekt?<br />
        Du vil samtidig slette alle relationer til dette forretningsområde.</td>
	</tr>
	<tr>
	   <td><a href="fomr.asp?menu=tok&func=sletok&id=<%=id%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
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
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM fomr WHERE id = "& id &"")
    oConn.execute("DELETE FROM fomr_rel WHERE for_fomr = "& id &"")

	Response.redirect "fomr.asp?menu=tok&shokselector=1"
	
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

        business_area_label = SQLBless(request("FM_business_area_label"))
        business_unit = SQLBless(request("FM_business_unit"))

		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)

        if len(trim(request("FM_aktok"))) <> 0 then
        aktok = 1
		else
        aktok = 0
        end if


        if len(trim(request("FM_jobok"))) <> 0 then
        jobok = 1
		else
        jobok = 0
        end if

        if len(trim(request("FM_konto"))) <> 0 then
        konto = request("FM_konto")
        else
        konto = 0
        end if

        fomr_segment = request("FM_fomr_segment")


		if func = "dbopr" then
		oConn.execute("INSERT INTO fomr (navn, editor, dato, jobok, aktok, konto, business_unit, business_area_label, fomr_segment) VALUES "_
        &" ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& jobok &", "& aktok &", "& konto &", '"& business_unit &"', '"& business_area_label &"', '"& fomr_segment &"')")
		else
		oConn.execute("UPDATE fomr SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', jobok = "& jobok &", aktok = "& aktok &", "_
        &" konto = "& konto &", business_unit = '"& business_unit &"', business_area_label = '"& business_area_label &"', fomr_segment = '"& fomr_segment &"' WHERE id = "& id &"")
		end if
		
		Response.redirect "fomr.asp?menu=tok&shokselector=1"
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "Opret" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
    jobokCHK = "CHECKED"
	aktokCHK = ""
    konto = 0
    fomr_segment = ""

	else
	strSQL = "SELECT navn, editor, dato, konto, aktok, jobok, business_unit, business_area_label, fomr_segment FROM fomr WHERE id= " & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = formatdatetime(oRec("dato"), 2)
	strEditor = oRec("editor")


    business_unit = oRec("business_unit")
    business_area_label = oRec("business_area_label")

    if oRec("jobok") = 1 then
    jobokCHK = "CHECKED"
    else
    jobokCHK = ""
    end if

	if oRec("aktok") = 1 then
    aktokCHK = "CHECKED"
    else
    aktokCHK = ""
    end if

    konto = oRec("konto")

    fomr_segment = oRec("fomr_segment")

	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdater" 
	end if
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->

    <%call menu_2014() %>

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px;  visibility:visible;">
	<h4>Forretningsområder <span style="font-size:9px; font-weight:lighter;"> - <%=varbroedkrumme %></span></h4>

        	<form action="fomr.asp?menu=tok&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
	<table cellspacing="0" cellpadding="2" border="0" width="100%">

    	
	<%if dbfunc = "dbred" then%>
	<tr>
		<td colspan="2" valign="bottom" style="height:30;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></td>
	</tr>
	<%end if%>
	
	<tr>
		<td style="width:350px; padding-top:20px;"><span style="color:red;">*</span> <b>Forretningsområde:</b></td>
		<td valign="top" style="padding-top:20px;"><input type="text" name="FM_navn" value="<%=strNavn%>" style="width:300px;">&nbsp;&nbsp;&nbsp;&nbsp;Label: <input type="text" name="FM_business_area_label" value="<%=business_area_label%>" style="width:80px;"></td>
	</tr>
    
	<tr>
		<td style="width:350px; padding-top:20px;"><b>Unit / afdeling:</b></td>
		<td valign="top" style="padding-top:20px;"><input type="text" name="FM_business_unit" value="<%=business_unit%>" style="width:430px;"></td>
	</tr>

    <tr>
		<td style="padding-top:20px;" colspan="2"><b>Indstillinger:</b></td>
		
	</tr>
        <tr>
		<td>Dette forretningsonmråde kan bruges på jobniveau?</td>
		<td valign="top"><input type="checkbox" name="FM_jobok" <%=jobokCHK %>></td>
	</tr>
    <tr>
		<td>Dette forretningsonmråde kan bruges på aktivitetsniveau?</td>
		<td valign="top"><input type="checkbox" name="FM_aktok" <%=aktokCHK %>></td>
	</tr>
    <%if cint(fomr_account) = 0 then %>
     <tr>
		<td style="padding-top:20px;"><b>Konto</b><br />
          <span style="color:#999999;">Der kan vælges mellem konti oprettet på licensindehaveren (jeres eget firma). Når dette forretningsområde tilknyttes et job, vil den
            valgte konto stå forvalgt ved fakturering.</span>
		</td>
		<td valign="top" style="padding-top:20px;"><select name="FM_konto" style="width:350px;">
              <option value="0">Ingen</option>

        <%
            licensindehaverKid = 0
            call licKid()

        strSQLkonto = "SELECT navn, kontonr, id FROM kontoplan WHERE kid = "& licensindehaverKid
         oRec2.open strSQLkonto, oConn, 3
         while not oRec2.EOF 
            
            if cint(konto) = oRec2("id") then
            kontoSEL = "SELECTED"
            else
            kontoSEL = ""
            end if

           
            kontonrVal = " ("& oRec2("kontonr") &")"
            
         
            %><option value="<%=oRec2("id") %>" <%=kontoSEL %>><%=oRec2("navn") & kontonrVal%></option><%

         oRec2.movenext
         wend
         oRec2.close  %>

          

    </select></td>
	</tr>
  
    <%end if %>


    <tr><td style="padding-top:20px;" colspan="2"><b>Dette forretningsområde skal være forvalgt på job oprettet på kunder i følgende segmenter:</b><br />

        <%strSQL = "SELECT id, navn FROM kundetyper WHERE id <> 0 ORDER BY navn"
            
            oRec.open strSQL, oConn, 3
            while not oRec.EOF 

            
            if instr(fomr_segment, "#"& oRec("id") &"#") <> 0 then
            fomr_segmentCHK = "CHECKED"
            else
            fomr_segmentCHK = ""
            end if


            %>
            <input type="checkbox" name="FM_fomr_segment" value="#<%=oRec("id") %>#" <%=fomr_segmentCHK %> /> <%=oRec("navn") %><br />
            <%


            oRec.movenext
            wend 
            oRec.close
              %>



        </td></tr>
	<tr>
		<td colspan="2" align="right"><br /><input type="submit" value="<%=varSubVal %> >>" /></td>
	</tr>
	
	</table>
                </form>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%case "stat"%>
	
	
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	



    <script src="inc/fomr_jav.js"></script>


	<%

        call menu_2014()

    
	
	
	'**************************************************
	'***** Faste Filter kriterier *********************
	'**************************************************
		
	
	'*** Medarbejdere / projektgrupper
	selmedarb = session("mid")
	
	'*** Job og Kundeans ***
	call kundeogjobans()
	
    'call medarbogprogrp("oms")
	medarbSQlKri = ""
	kundeAnsSQLKri = ""
	jobAnsSQLkri = ""
	jobAns2SQLkri = ""
	
	if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
	progrp = 0
	end if
	
    if len(trim(request("FM_fordel_jobakt"))) <> 0 then
    fordel_jobakt = request("FM_fordel_jobakt")
    else
    fordel_jobakt = 0
    end if

    if cint(fordel_jobakt) = 0 then
        fordel_jobakt0CHK = "CHECKED"
    else
        fordel_jobakt1CHK = "CHECKED"
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
			     else
			     medarbSQlKri = medarbSQlKri & " OR m.mid = " & intMids(m)
			     kundeAnsSQLKri = kundeAnsSQLKri & " OR kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
			     jobAnsSQLkri = jobAnsSQLkri & " OR jobans1 = "& intMids(m)  
			     jobAns2SQLkri = jobAns2SQLkri & " OR jobans2 = "& intMids(m)
			     end if
			     
			    next
			    
			    medarbSQlKri = medarbSQlKri & ")"
			    
			jobAnsSQLkri =  "AND ("& jobAnsSQLkri &")"
			jobAns2SQLkri =  "AND (" & jobAns2SQLkri &")"
			
	
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
	jobidFakSQLkri = " OR jobid = -1 "
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



    if media <> "export" then 
	ldTop = 400
	ldLft = 240
	else
	ldTop = 15
	ldLft = 10
	end if
	%>
	<div id="load" style="position:absolute; display:; visibility:visible; top:<%=ldTop%>px; left:<%=ldLft%>px; width:300px; background-color:#ffffff; border:10px #9ACD32 solid; padding:10px; z-index:100000000;">
    <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" /><br />
	&nbsp;&nbsp;&nbsp;&nbsp;Vent veligst. <br />&nbsp;&nbsp;&nbsp;&nbsp;Forventet loadtid: 3 - 20 sek...
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	</td></tr></table>
	
	</div>
	<%
	
	


	
	timermedarbSQLkri = replace(medarbSQLkri, "m.mid", "tmnr")	
	
	
	if request("print") <> "j" then
	pleft = 90
	ptop = 102
	'ptopgrafik = 348
	else
	pleft = 20
	ptop = 20
	'ptopgrafik = 90
	end if
	%>	
	<div id="Div1" style="position:absolute; left:<%=pleft%>px; top:<%=ptop%>px; visibility:visible;">
	
	
	<%
	
	oimg = "pie_chart_48_hot.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = "Forretningsområder"
	
    if request("print") = "j" then
	media = "print"
        call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
    end if
	
	if request("print") <> "j" then 
	
	call filterheader_2013(0,0,800,oskrift)%>
    	    
	  
	    <table cellspacing=0 cellpadding=0 border=0 width=100% bgcolor="#FFFFFF">
	    <form action="fomr.asp?rdir=<%=rdir %>&menu=stat&func=stat" method="post">
	    
	<%end if %>
	    <%call medkunderjob %>
	    
	    
	    </td>
	    </tr>
        <tr><td colspan="2"><br /><b>Vis forretingsområder fordelt på:</b> <br />
            <input type="radio" name="FM_fordel_jobakt" value="0" <%=fordel_jobakt0CHK %> />Job<br />
            <input type="radio" name="FM_fordel_jobakt" value="1" <%=fordel_jobakt1CHK %> />Aktiviteter
            </td></tr>
	<tr>
			<!-- Brug altid datointerval, FM_usedatokri = 1 -->
			<input type="hidden" name="FM_usedatokri" id="FM_usedatokri" value="1">
			<td><br /><br /><b>Periode:</b><br>
			<!--#include file="inc/weekselector_s.asp"--> <!-- b -->
			
			</td>
			<td align=right valign=bottom>
			<%if request("print") <> "j" then%>
			<img src="../ill/blank.gif" width="250" height="1" alt="" border="0"><input type="submit" value=" Kør >> ">
			<%end if%>
			
			</td>
		</tr>
		</form>
		</table><br>
		
		
		<!-- fiulter DIV-->
		</td></tr>
		</table>
		</div>

	
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
				
				
			jidSQLkri =  " OR id <> 0 "
			jobnrSQLkri = " OR tjobnr <> '0' "
			'seridFakSQLkri = " OR aftaleid <> 0 
			
		end if
	
	
		'**************** Trimmer SQL states ************************
		
		len_jobnrSQLkri = len(jobnrSQLkri)
		right_jobnrSQLkri = right(jobnrSQLkri, len_jobnrSQLkri - 3)
		jobnrSQLkri =  right_jobnrSQLkri
		
		len_jidSQLkri = len(jidSQLkri)
		right_jidSQLkri = right(jidSQLkri, len_jidSQLkri - 3)
		jidSQLkri =  right_jidSQLkri
		
		jidSQLKri = replace(jidSQLkri, "id", "aktiviteter.job")
		
		
		'*****************************************************************************************************
	
    Response.flush

	
	
	'*** Alle timer, uanset fomr. ***
	call akttyper2009(2)
	totaltimer = 0
	sqlDatoStart = strAar&"/"&strMrd&"/"&strDag
	sqlDatoSlut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	
	'if ignper <> 1 then
	tdatoSQLkri = " AND tdato BETWEEN '"& sqlDatoStart &"' AND '"& sqlDatoSlut &"'"
	'else
	'tdatoSQLkri = ""
	'end if
	
    strJobSQLKri = " AND (j.jobnr = 0 "
    strJobTxt = ""

	strSQL = "SELECT sum(timer.timer) AS timertot, tjobnr, tjobnavn, tknavn FROM job AS j "_
    &" LEFT JOIN timer ON (tjobnr = j.jobnr) WHERE ("& aty_sql_realhours &") "_
    &" AND ("& timermedarbSQLkri &") AND ("& jobnrSQLkri &") "& jobstKri &" " & tdatoSQLkri & " GROUP BY tjobnr ORDER BY tknavn, tjobnavn"
	
    'if lto = "jttek" then
    'Response.Write strSQL
    'Response.flush
    'end if

	'Response.write "timermedarbSQLkri: "& timermedarbSQLkri & "<br>"& strSQL
	'Response.flush
	totaltimer = 0
    lastKnavn = ""
	oRec.open strSQL, oConn, 3 
	While not oRec.EOF 
	totaltimer = totaltimer + oRec("timertot")
	
    '** Led kun efter forretningeområder på job der er registreret på i periode
    strJobSQLKri = strJobSQLKri & " OR j.jobnr = '"& oRec("tjobnr") & "'"

    if lastKnavn <> oRec("tknavn") then
    strJobTxt = strJobTxt & "<br><br><b>"& oRec("tknavn") & "</b><br>"
    end if

    strJobTxt = strJobTxt & oRec("tjobnavn") & " ("& oRec("tjobnr") &"): " & oRec("timertot") & " timer<br>" 

    lastKnavn = oRec("tknavn")
    oRec.movenext
    wend
	oRec.close
	
	if totaltimer <> 0 then
	totaltimer = totaltimer
	else
	totaltimer = 1
	end if

    strJobSQLKri = strJobSQLKri & ")"
	
	Response.write "Indtastet i alt i den valgte periode, på de valgte medarbejdere og job,<br> uanset om aktivitet er tilknyttet et forretningsområde: <b>"& formatnumber(totaltimer, 2)&"</b> timer.<br>"

	
	%>
	<br><br>
	<%
	tTop = 0
    tLeft = 0
    tWdth = 810


    call tableDiv(tTop,tLeft,tWdth)

	%>
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan=3 valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right width="8" valign=top rowspan=2><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class=alt><b>Forretningsområder</b></td>
		<td class=alt align=right><b>Timer</b></td>
		<td class=alt align=right><b>%-del af total</b> (<%=formatnumber(totaltimer, 2)%>)</td>
	</tr>
	<%
	
	call akttyper2009(2)

    dim fomrjobakt, ktypeid, fomridSum, ktypenavn, fomrnavn, fomridSumTot 
    redim fomrjobakt(50), fomrnavn(50), ktypeid(500), ktypenavn(500), fomridSum(500,50),  fomridSumTot(50) '** ktyper , Forretningsomr

    '*** Forretningsområder Navne **'
     '&" LEFT JOIN fomr AS f ON (f.id = fr.for_id) "_
     '  f.navn AS forromrnavn 
    strSQLf = "SELECT f.navn AS forromrnavn, f.id FROM fomr AS f WHERE f.id <> 0"
     oRec6.open strSQLf, oConn, 3
    while not oRec6.EOF 

    fomrnavn(oRec6("id")) = oRec6("forromrnavn")

    oRec6.movenext
    wend
    oRec6.close


    jobnrWrt = 0
    isjobnrWrt = "#0#" 

    ja = fordel_jobakt
    for ja = ja to ja 'fordeling mellem job og aktiviteter 


        if cint(ja) = 0 then 
        %><h4>Forretningsområder fordelt på job</h4><%
        else
        %><h4>Forretningsområder fordelt på aktiviteter</h4><%
        end if


    '*** Forretningsområder Relationer **'
    jobnrSQLkri = replace(jobnrSQLkri, "tjobnr", "j.jobnr")

    strSQLfor = "SELECT fr.for_fomr, fr.for_aktid, fr.for_jobid, for_faktor, j.jobnr, kt.id AS ktypeid, kt.navn AS ktypenavn FROM job AS j"_
    &" LEFT JOIN fomr_rel AS fr ON (fr.for_jobid = j.id) "_
    &" LEFT JOIN kunder k ON (k.kid = j.jobknr)"_
    &" LEFT JOIN kundetyper kt ON (kt.id = k.ktype)"_
    &" WHERE fr.for_jobid <> 0 "& strJobSQLKri &" GROUP BY kt.id, fr.for_fomr, fr.for_jobid, fr.for_aktid"

    'AND ("& jobnrSQLkri &")
    'strSQL = strSQL & ", kt.id AS ktypeid, kt.navn AS ktypenavn"
    '&" LEFT JOIN aktiviteter AS a ON (a.id = fr.for_aktid) "_
    'fr.for_aktid, fr.for_jobid'

    'if lto = "jttek" then
    'Response.Write "<br><br>Forretningsområder Relationer<br>"& strSQLfor & "<br><br>"
    'Response.flush
    'end if

    x = 0
    oRec6.open strSQLfor, oConn, 3
    while not oRec6.EOF 
                
                
                if cint(ja) = 1 then
                fomrjobakt(oRec6("for_fomr")) = " taktivitetid = " & oRec6("for_aktid")
                grpBYja = "tjobnr"

                else
                fomrjobakt(oRec6("for_fomr")) = " tjobnr = '"& oRec6("jobnr") & "'"
                grpBYja = "taktivitetid"

                if instr(isjobnrWrt, ",#"& oRec6("jobnr") &"#") = 0 then
                jobnrWrt = 0
                else
                jobnrWrt = 1
                end if

                end if

                        if (ja = 1 AND oRec6("for_aktid") <> 0) OR (ja = 0 AND oRec6("for_jobid") AND cint(jobnrWrt) = 0) then
                
                        if isNULL(oRec6("ktypeid")) = true then
                        ktypeIDthis = 0
                        else
                        ktypeIDthis = oRec6("ktypeid")
                        end if

                        'Response.write "ktypeIDthis: "& ktypeIDthis & "<br>"
                        'Response.flush

                        if isNull(oRec6("ktypenavn")) = true then
                        ktypenavn(0) = ""
                        else
                        ktypenavn(oRec6("ktypeid")) = oRec6("ktypenavn")
                        end if

                        'fomrjobakt(oRec6("for_fomr")) = oRec6("forromrnavn")
                        '"& replace(oRec6("for_faktor"), ",", ".") &"/100), 2)
                        strSQLt = "SELECT ROUND(SUM(t.timer), 2) AS timerfomr FROM timer AS t WHERE " & fomrjobakt(oRec6("for_fomr")) & " AND ("& aty_sql_realhours &") AND ("& timermedarbSQLkri &") " & tdatoSQLkri & " GROUP BY "& grpBYja &"" 

                        'Response.write "<br><br>"& strSQLt
                        'Response.flush
                        oRec5.open strSQLt, oConn, 3
                        if not oRec5.EOF then

               
                        fomridSum(ktypeIDthis, oRec6("for_fomr")) = fomridSum(ktypeIDthis, oRec6("for_fomr")) + (oRec5("timerfomr"))


                        'oRec5.movenext
                        end if
                        oRec5.close


                        isjobnrWrt = isjobnrWrt & ",#"& oRec6("jobnr") &"#" 

                        end if 'ja = 0/1

    x = x + 1
    oRec6.movenext
    wend
    oRec6.close


   


    l = 0
    for k = 0 to 500

    
     if ktypenavn(k) <> "" then
     
     

            'if x = 0 then
            'lastKtypeNavn = oRec("ktypenavn")
            'end if

        %>
        <tr>
		    <td bgcolor="#cccccc" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	    </tr>
        <tr>
            <td bgcolor="#8cAAe6" class=alt colspan=5 style="padding:6px 4px 2px 7px;">Segment: <b> <%= ktypenavn(k)%></b></td>
        </tr>

        <%
        end if
      


            for f = 0 to 50

            if fomridSum(k,f) > 0 then
            

            select case right(l, 1)
	        case 0, 2, 4, 6, 8
	        bgthis = "#eff3ff"
	        case else
	        bgthis = "#FFFFFF"
	        end select
            
            %>
	        <tr>
		        <td bgcolor="#cccccc" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	        </tr>
	        <tr bgcolor="<%=bgthis%>">
		        <td style="height:20px;">&nbsp;</td>
		        <td><%=fomrnavn(f)%>:</td>
		        <td align=right>
		
		        <b><%=formatnumber(fomridSum(k,f))%></b>&nbsp;timer</td>
		        <td align=right><%=formatpercent((fomridSum(k,f)/totaltimer), 2)%></td>
		        <td>&nbsp;</td>
	         </tr>
	        <%
            fomridSumTot(f) = fomridSumTot(f) + fomridSum(k,f) 
            timerfomr = timerfomr + fomridSum(k,f)
       
            l = l + 1     
            'Response.Write "fomr: "& fomrnavn(f) &":: "& fomridSum(k,f) & "<br>"
            end if

            next

    next


     next 'ja

    %>
      
       <tr>
		    <td bgcolor="#cccccc" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	    </tr>
       <tr>
		    <td colspan="5"><img src="ill/blank.gif" width="1" height="30" border="0" alt=""></td>
	    </tr>
    
        <tr>
            <td bgcolor="#FFFFFF" colspan=5 style="padding:16px 4px 2px 7px;">&nbsp;</td>
        </tr>
        <tr>
            <td bgcolor="#FFFFFF" colspan=5 style="padding:6px 4px 2px 7px;"><b>Totaler:</b></td>
        </tr>

     <%


    '************ totaler ********************'

    l = 0
   
            for f = 0 to 50

            if fomridSumTot(f) > 0 then
            

            select case right(l, 1)
	        case 0, 2, 4, 6, 8
	        bgthis = "#eff3ff"
	        case else
	        bgthis = "#FFFFFF"
	        end select
            
            %>
	        <tr>
		        <td bgcolor="#cccccc" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	        </tr>
	        <tr bgcolor="<%=bgthis%>">
		        <td style="height:20px;">&nbsp;</td>
		        <td><%=fomrnavn(f)%>:</td>
		        <td align=right>
		
		        <b><%=formatnumber(fomridSumTot(f))%></b>&nbsp;timer</td>
		        <td align=right><%=formatpercent((fomridSumTot(f)/totaltimer), 2)%></td>
		        <td>&nbsp;</td>
	         </tr>
	        <%

            fomridSumGTot = fomridSumGTot + fomridSumTot(f)
            l = l + 1
           'Response.Write "fomr: "& fomrnavn(f) &":: "& fomridSum(k,f) & "<br>"
            end if

            next




    	if l = 0 then%>
	
	<tr bgcolor="#eff3ff">
		<td height=20>&nbsp;</td>
		<td colspan=3><br><br>Der er <b>ikke</b> oprettet nogen forretningsområder, eller der er ikke indtastet timer på de oprettede forretningsområder. <br>
		Forretningsområder kan oprettes af admin brugere i kontrolpanelet.<br><br>&nbsp;</td>
		<td>&nbsp;</td>
	 </tr>
	<%end if%>
	
	
	<tr bgcolor="#FFFFFF">
		<td colspan=2 style="padding:6px 4px 2px 7px;"><b>Ialt:</b></td>
		
     
        <td align=right><b><%=fomridSumGTot %></b> timer</td>
        <td>&nbsp;</td>
	
		<td align=right valign=top><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	</table>
	</div>
	
    <br />
	Timer indtastet på aktiviter uden forretningsområde tilknyttet ~ <b><%=formatnumber(totaltimer - timerfomr, 2) %></b> timer.<br />
    Timeforbrug på forretningsområder er fordelt således at hvis et job eller en aktivitet dækker flere forretningsområder tæller timer 100% med i begge områder.
	
	<br><br>&nbsp;


    <br /><br />
    <b><span id="sp_grundlag" style="color:#5582d2;">+ Grundlag</span></b>
    <div style="width:780px; padding:20px; background-color:#ffffff; visibility:hidden; display:none;" id="div_grundlag" >
    
    <%=strJobTxt %>
	
	</div>

    <br><br><br><br><br><br>&nbsp;


     <!--<script>
         document.getElementById("load").style.visibility = "hidden";
         document.getElementById("load").style.display = "none";
			    </script>
         -->







    <%
    ' ****  siden slutter her PERFORMANCE ***' 
    Response.end

    for t = 0 to 1
	'for_faktor
	strSQL = "SELECT aktiviteter.id AS aid, t.timer/100 AS timerfomr, fomr.id, fomr.navn AS fomrnavn"

    if t = 0 then
    strSQL = strSQL & ", kt.id AS ktypeid, kt.navn AS ktypenavn"
    end if

	strSQL = strSQL &" FROM fomr "_
    &" LEFT JOIN fomr_rel ON (for_fomr = fomr.id) "_
    &" LEFT JOIN job AS j ON (j.id = for_jobid) "_
    &" LEFT JOIN aktiviteter ON (aktiviteter.id = for_aktid) "_
	&" LEFT JOIN timer AS t ON (taktivitetid = aktiviteter.id)" 

    if t = 0 then
    strSQL = strSQL &" LEFT JOIN kunder k ON (k.kid = tknr)"_
    &" LEFT JOIN kundetyper kt ON (kt.id = k.ktype)"
    grpBY = "k.ktype, kid, fomr.id"
    ordBY = "k.ktype, fomrnavn"
	else
    grpBY = "fomr.id"
    ordBY = "fomrnavn"
    end if
	
     strSQL = strSQL &" WHERE fomr.navn <> '' AND t.timer IS NOT NULL AND t.timer > 0 AND ("& aty_sql_realhours &") AND ("& timermedarbSQLkri &") AND ("& jobnrSQLkri &") " & tdatoSQLkri &" "& jobstKri &" GROUP BY "& grpBY &" ORDER BY "& ordBY


    'if lto = "jttek" then
    'response.write strSQL
	'Response.flush
    'end if
	
    lastKtype = -1
    lastKtypeNavn = ""
    x = 0
	timerfomr = 0
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
	select case right(x, 1)
	case 0, 2, 4, 6, 8
	bgthis = "#eff3ff"
	case else
	bgthis = "#FFFFFF"
	end select

    if t = 0 then

        if (lastKtype <> oRec("ktypeid") OR x = 0) AND oRec("ktypenavn") <> "" then

            'if x = 0 then
            'lastKtypeNavn = oRec("ktypenavn")
            'end if

        %>
        <tr>
		    <td bgcolor="#cccccc" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	    </tr>
        <tr>
            <td bgcolor="#8cAAe6" class=alt colspan=5 style="padding:6px 4px 2px 7px;">Segment: <b><%=oRec("ktypenavn")%></b></td>
        </tr>

        <%
        end if
    
    else

        if x = 0 then

         %>
        <tr>
		    <td bgcolor="#cccccc" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	    </tr>
        <tr>
            <td bgcolor="#FFFFFF" colspan=5 style="padding:16px 4px 2px 7px;">&nbsp;</td>
        </tr>
        <tr>
            <td bgcolor="lightpink" colspan=5 style="padding:6px 4px 2px 7px;"><b>Totaler:</b></td>
        </tr>

        <%

        end if

    end if
	
    if t = 1 OR (oRec("timerfomr") <> 0 AND t = 0) then
	%>
	<tr>
		<td bgcolor="#cccccc" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bgthis%>">
		<td style="height:20px;">&nbsp;</td>
		<td><%=oRec("fomrnavn")%>:</td>
		<td align=right>
		
		<b><%=formatnumber(oRec("timerfomr"))%></b>&nbsp;timer</td>
		<td align=right><%=formatpercent((oRec("timerfomr")/totaltimer), 2)%></td>
		<td>&nbsp;</td>
	 </tr>
	<%
	
	timerfomr = timerfomr + oRec("timerfomr")
    
    if t = 0 then

    if oRec("ktypeid") <> "" then
    lastKtype = oRec("ktypeid")
    else
	lastKtype = -1
    end if

    lastKtypeNavn = oRec("ktypenavn")

    end if

	x = x + 1

    end if
	
	
	oRec.movenext
	wend
	oRec.close 

    next
	
	
	

	case else%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	

    <%call menu_2014() %>

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; width:1000px; visibility:visible;">
	<h4>Forretningsområder</h4>

	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
    <td valign="top">
	Sortér efter <a href="fomr.asp?menu=tok&sort=navn">Navn</a> eller <a href="fomr.asp?menu=tok&sort=nr">Id nr.</a>
	<img src="../ill/blank.gif" width="500" height="1" alt="" border="0"></td><td align="right">

         <%
                nWdt = 200
                nTxt = "Opret nyt forretningsområde "
                nLnk = "fomr.asp?menu=tok&func=opret"
                nTgt = ""
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) %>

	
	<br>&nbsp;
	</td>
	</tr>
	</table>
	
    <div style="padding:20px; background-color:#FFFFFF;">
	<table cellspacing="0" cellpadding="2" border="0" width="100%">
	
	<tr bgcolor="#5582D2">
		<td class=alt style="height:20px;"><b>Id</b></td>
		<td class=alt><b>Forretningsområde</b></td>
        <td class="alt">Unit/ Afd.</td>
        <%if cint(fomr_account) <> 1 then %>
        <td class=alt><b>Konto</b></td>
        <%end if %>
        <td class=alt>Job (bruges af antal job)</td>
        <td class=alt>Akt. (bruges af antal aktiviteter)</td>
		<td class=alt>&nbsp;</td>
	</tr>
	<%
	sort = Request("sort")

    strSQL = "SELECT f.id, f.navn, COUNT(relakt.for_fomr) AS relakt_antal, kp.kontonr AS kkontonr, kp.navn AS kpnavn, jobok, aktok, business_unit, business_area_label FROM fomr AS f"
    strSQL = strSQL & " LEFT JOIN kontoplan AS kp ON (kp.id = f.konto) "
    strSQL = strSQL & " LEFT JOIN fomr_rel AS relakt ON (relakt.for_fomr = f.id AND relakt.for_aktid <> 0) GROUP BY f.id "

    if sort = "navn" then
	strSQL = strSQL & "ORDER BY f.navn"
	else
	strSQL = strSQL & "ORDER BY f.id"
	end if

    'Response.Write strSQL
    'Response.flush
    f = 0
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
        
        select case right(f,1)
        case 0,2,4,6,8
        bgt = "#EFF3FF"
        case else 
        bgt = "#FFFFFF"
        end select

         if cint(fomr_account) <> 1 then 
         fcspan = 7
         else
         fcspan = 6
         end if
	%>
	<tr>
		<td bgcolor="#999999" colspan="<%=fcspan %>" style="padding:0px 0px 0px 0px;"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bgt %>">
		
		<td><%=oRec("id")%></td>
		<td><a href="fomr.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a> 

            <%if oRec("business_area_label") <> "" then
                %>

            (<%=oRec("business_area_label")%>)
                
            <%
             end if %>

		</td>
        <td><%=oRec("business_unit")%></td>
         <%if cint(fomr_account) <> 1 then %>
        <td><%=oRec("kpnavn") &" "& oRec("kkontonr")%></td>
        <%end if %>
        <td>
            <%if cint(oRec("jobok")) = 1 then %>
            <i>V</i>&nbsp;&nbsp;
                <%
                    antalfomrJob = 0
                    strSQLantaljob = " SELECT COUNT(reljob.for_fomr) AS reljob_antal FROM fomr_rel AS reljob WHERE (for_fomr = "& oRec("id") &" AND for_jobid <> 0) "  
                    
                    oRec2.open strSQLantaljob, oConn, 3
	                if not oRec2.EOF then

                    antalfomrJob = oRec2("reljob_antal")

                    end if
                    oRec2.close
                    %>
                
            (<a href="jobs.asp?FM_fomr=X234<%=oRec("id") %>X234" class="vmenu" target="_blank"><%=antalfomrJob%></a>)
            
            <%else %>

            <%end if %>
            
        </td>
        <td>
            
             <%if cint(oRec("aktok")) = 1 then %>
            <i>V</i>&nbsp;&nbsp;(<%=oRec("relakt_antal")%>)
            <%else %>

            <%end if %>
            
            </td>
        
		<td><a href="fomr.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="" border="0"></a></td>
		
	</tr>
	<%
        f = f + 1
	x = 0
	oRec.movenext
	wend
	%>	
	
	</table>
        </div>
	
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
