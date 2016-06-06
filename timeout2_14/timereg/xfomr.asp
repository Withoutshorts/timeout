<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->

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
	
	
	
	
	
	
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:10; visibility:visible;">
	<%call crmaktstatheader%>
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> et forretningsområde. Er dette korrekt?<br />
        Du vil samtidig slette alle relationer til dette forretniingsområde.</td>
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
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO fomr (navn, editor, dato) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"')")
		else
		oConn.execute("UPDATE fomr SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"' WHERE id = "&id&"")
		end if
		
		Response.redirect "fomr.asp?menu=tok&shokselector=1"
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
	strSQL = "SELECT navn, editor, dato FROM fomr WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
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
	<tr><form action="fomr.asp?menu=tok&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
    	<td valign="top" colspan="2"><font class="pageheader">Kontrolpanel | Forretningsområder | <%=varbroedkrumme%></font></td>
	</tr>
	<%if dbfunc = "dbred" then%>
	<tr>
		<td colspan="2" valign="bottom" style="height:30;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></td>
	</tr>
	<%end if%>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td>Navn:</td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
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
	<%case "stat"%>
	
	<script>

	    function clearJobsog() {
	        document.getElementById("FM_jobsog").value = ""
	    }
	
	</script>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	<%call tsamainmenu(7)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	if showonejob <> 1 then
		call stattopmenu()
	end if
	
	%>
	</div>
	<%
	
	
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
	
	timermedarbSQLkri = replace(medarbSQLkri, "m.mid", "tmnr")	
	
	
	if request("print") <> "j" then
	pleft = 20
	ptop = 132
	'ptopgrafik = 348
	else
	pleft = 10
	ptop = 65
	'ptopgrafik = 90
	end if
	%>	
	<div id="Div1" style="position:absolute; left:<%=pleft%>; top:<%=ptop%>; visibility:visible;">
	
	
	<%
	
	oimg = "ac0023-24_header.gif"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = "Forretningsområder"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	if request("print") <> "j" then 
	
	call filterheader(0,0,800,pTxt)%>
    	    
	  
	    <table cellspacing=0 cellpadding=10 border=0 width=100% bgcolor="#FFFFFF">
	    <form action="fomr.asp?rdir=<%=rdir %>&menu=stat&func=stat" method="post">
	    
	<%end if %>
	    <%call medkunderjob %>
	    
	    
	    </td>
	    </tr>
	<tr>
			<!-- Brug altid datointerval, FM_usedatokri = 1 -->
			<input type="hidden" name="FM_usedatokri" id="FM_usedatokri" value="1">
			<td><b>Periode:</b><br>
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
		if cint(kundeid) = 0 AND cint(jobid) = 0 AND cint(jobans) = 0 _
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
	
	
	
	'*** Alle timer, uanset fomr. ***
	
	totaltimer = 0
	sqlDatoStart = strAar&"/"&strMrd&"/"&strDag
	sqlDatoSlut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	
	'if ignper <> 1 then
	tdatoSQLkri = " AND tdato BETWEEN '"& sqlDatoStart &"' AND '"& sqlDatoSlut &"'"
	'else
	'tdatoSQLkri = ""
	'end if
	
	strSQL = "SELECT sum(timer.timer) AS timertot FROM timer WHERE (tfaktim = 1 OR tfaktim = 2 OR tfaktim = 6 OR tfaktim = 13 OR tfaktim = 14 OR tfaktim = 20 OR tfaktim = 21) AND ("& timermedarbSQLkri &") AND ("& jobnrSQLkri &") " & tdatoSQLkri 
	
	'Response.write strSQL
	'Response.flush
	
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	totaltimer = oRec("timertot")
	end if
	oRec.close
	
	if totaltimer <> 0 then
	totaltimer = totaltimer
	else
	totaltimer = 1
	end if
	
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

    for t = 0 to 1
	
	strSQL = "SELECT aktiviteter.id AS aid, sum(timer.timer*for_faktor/100) AS timerfomr, fomr.id, fomr.navn AS fomrnavn"

    if t = 0 then
    strSQL = strSQL & ", kt.id AS ktypeid, kt.navn AS ktypenavn"
    end if

	strSQL = strSQL &" FROM fomr "_
    &" LEFT JOIN fomr_rel ON (for_fomr = fomr.id) "_
    &" LEFT JOIN aktiviteter ON (aktiviteter.id = for_aktid) "_
	&" LEFT JOIN timer ON (taktivitetid = aktiviteter.id AND ("& aty_sql_realhours &") AND ("& timermedarbSQLkri &") AND ("& jobnrSQLkri &") " & tdatoSQLkri &")"

    if t = 0 then
    strSQL = strSQL &" LEFT JOIN kunder k ON (k.kid = tknr)"_
    &" LEFT JOIN kundetyper kt ON (kt.id = k.ktype)"_
	&" WHERE fomr.navn <> '' GROUP BY k.ktype, fomr.id ORDER BY k.ktype, fomrnavn"
    else
    strSQL = strSQL &" WHERE fomr.navn <> '' GROUP BY fomr.id ORDER BY fomrnavn"
    end if
	
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
	
	
	
	if x = 0 then%>
	<tr>
		<td bgcolor="#003399" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="#eff3ff">
		<td height=20 style="border-left:1px #003399 solid;">&nbsp;</td>
		<td colspan=3><br><br>Der er <b>ikke</b> oprettet nogen forretningsområder, eller der er ikke indtastet timer på de oprettede forretningsområder. <br>
		Forretningsområder kan oprettes af admin brugere i kontrolpanelet.<br><br>&nbsp;</td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	 </tr>
	<%end if%>
	
	
	<tr bgcolor="#5582D2">
		<td width="8" valign=top height=20><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=3 valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	</table>
	</div>
	
    <br />
	Timer indtastet på aktiviter uden forretningsområde tilknyttet ~ <%=formatnumber(totaltimer - timerfomr, 2) %> timer.<br />
    Timeforbrug på forretningsområder er normal-fordelt således at hvis en aktivitet dækker over 3 forretningsområder tæller hver time 1/3.
	
	<br><br>&nbsp;
	
	</div>
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
	<div id="sindhold" style="position:absolute; left:20; top:10; visibility:visible;">
	<%call crmaktstatheader%>
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr>
    <td valign="top"><font class="pageheader">Kontrolpanel | Forretningsområder</font>
	<br>
	<br>
	Sortér efter <a href="fomr.asp?menu=tok&sort=navn">Navn</a> eller <a href="fomr.asp?menu=tok&sort=nr">Id nr.</a>
	<img src="../ill/blank.gif" width="200" height="1" alt="" border="0">
	<a href="fomr.asp?menu=tok&func=opret">Opret nyt forretningsområde <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a><br>
	<br>&nbsp;
	</td>
	</tr>
	</table>
	
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2><img src="../ill/tabel_top_left.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan=4 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top rowspan=2><img src="../ill/tabel_top_right.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td width="50" class=alt><b>Id</b></td>
		<td class=alt><b>Forretningsområde</b></td>
        <td class=alt>Bruges af antal aktiviteter</td>
		<td class=alt>&nbsp;</td>
	</tr>
	<%
	sort = Request("sort")

    strSQL = "SELECT id, navn, COUNT(for_fomr) AS fomr_antal FROM fomr"

	strSQL = strSQL & " LEFT JOIN fomr_rel ON (for_fomr = id) GROUP BY fomr.id "

    if sort = "navn" then
	strSQL = strSQL & "ORDER BY navn"
	else
	strSQL = strSQL & "ORDER BY id"
	end if

    'Response.Write strSQL
    'Response.flush

	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	%>
	<tr>
		<td bgcolor="#003399" colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="#ffffff" onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'#ffffff');">
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td><%=oRec("id")%></td>
		<td><a href="fomr.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
        <td><%=oRec("fomr_antal")%></td>
		<td><a href="fomr.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="" border="0"></a></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>	
	<tr bgcolor="#5582D2">
		<td width="8" valign=top height=20 style="border-bottom:1px #003399 solid; border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=4 valign="top" style="border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top style="border-bottom:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	</table>
	
	<br><br><br>
	<a href="Javascript:window.close()" class=red>[Luk vindue]</a>
	<br><br>&nbsp;</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
